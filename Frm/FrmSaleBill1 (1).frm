VERSION 5.00
Object = "{C115893A-A3BF-43AF-B28D-69DB846077F3}#1.0#0"; "vsflex8u.ocx"
Object = "{0AFE7BE0-11B7-4A3E-978D-D4501E9A57FE}#1.0#0"; "c1sizer.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{FE5DCFAD-BC1D-11D2-94CF-004005455FAA}#1.4#0"; "ImpulseButton.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form FrmSaleBill 
   Caption         =   "›« Ê—… «·»Ì⁄"
   ClientHeight    =   7470
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9885
   HelpContextID   =   160
   Icon            =   "FrmSaleBill.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   RightToLeft     =   -1  'True
   ScaleHeight     =   7470
   ScaleWidth      =   9885
   Begin C1SizerLibCtl.C1Elastic EleMain 
      Height          =   7470
      Left            =   0
      TabIndex        =   21
      TabStop         =   0   'False
      Top             =   0
      Width           =   9885
      _cx             =   17436
      _cy             =   13176
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
      GridCols        =   4
      Frame           =   3
      FrameStyle      =   0
      FrameWidth      =   1
      FrameColor      =   -2147483628
      FrameShadow     =   -2147483632
      FloodStyle      =   1
      _GridInfo       =   $"FrmSaleBill.frx":038A
      AccessibleName  =   ""
      AccessibleDescription=   ""
      AccessibleValue =   ""
      AccessibleRole  =   9
      Begin ImpulseButton.ISButton CmdInfo 
         Height          =   615
         Left            =   9210
         TabIndex        =   73
         Top             =   15
         Width           =   660
         _ExtentX        =   1164
         _ExtentY        =   1085
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
         ButtonImage     =   "FrmSaleBill.frx":0416
         ButtonImageHover=   "FrmSaleBill.frx":10F0
         ColorButton     =   16777215
         DrawFocusRectangle=   0   'False
      End
      Begin C1SizerLibCtl.C1Elastic Ele 
         Height          =   435
         Index           =   3
         Left            =   15
         TabIndex        =   50
         TabStop         =   0   'False
         Top             =   6465
         Width           =   9855
         _cx             =   17383
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
            Left            =   8850
            Locked          =   -1  'True
            RightToLeft     =   -1  'True
            TabIndex        =   51
            TabStop         =   0   'False
            Top             =   30
            Visible         =   0   'False
            Width           =   300
         End
         Begin MSDataListLib.DataCombo DCboUserName 
            Height          =   315
            Left            =   2700
            TabIndex        =   52
            Top             =   45
            Width           =   1410
            _ExtentX        =   2487
            _ExtentY        =   556
            _Version        =   393216
            Enabled         =   0   'False
            Text            =   ""
            RightToLeft     =   -1  'True
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
            Left            =   6510
            RightToLeft     =   -1  'True
            TabIndex        =   160
            Top             =   30
            Width           =   1035
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "Œ’Ê„« "
            Height          =   255
            Index           =   50
            Left            =   7560
            RightToLeft     =   -1  'True
            TabIndex        =   159
            Top             =   90
            Width           =   600
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
            Left            =   8190
            RightToLeft     =   -1  'True
            TabIndex        =   158
            Top             =   30
            Width           =   1035
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "«·’«›Ì"
            Height          =   255
            Index           =   49
            Left            =   5910
            RightToLeft     =   -1  'True
            TabIndex        =   157
            Top             =   90
            Width           =   570
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
            Left            =   4860
            RightToLeft     =   -1  'True
            TabIndex        =   74
            Top             =   30
            Width           =   1035
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "«·„” Œœ„"
            Height          =   315
            Index           =   1
            Left            =   4125
            RightToLeft     =   -1  'True
            TabIndex        =   58
            Top             =   75
            Width           =   720
         End
         Begin VB.Label XPTxtCount 
            Alignment       =   2  'Center
            BackColor       =   &H00E2E9E9&
            Height          =   240
            Left            =   45
            RightToLeft     =   -1  'True
            TabIndex        =   57
            Top             =   90
            Width           =   705
         End
         Begin VB.Label XPTxtCurrent 
            Alignment       =   2  'Center
            BackColor       =   &H00E2E9E9&
            Height          =   240
            Left            =   1020
            RightToLeft     =   -1  'True
            TabIndex        =   56
            Top             =   90
            Width           =   675
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "—ﬁ„ «·”Ã·"
            Height          =   240
            Index           =   2
            Left            =   1710
            RightToLeft     =   -1  'True
            TabIndex        =   55
            Top             =   90
            Width           =   975
         End
         Begin VB.Label lbl 
            Alignment       =   2  'Center
            BackColor       =   &H00E2E9E9&
            Caption         =   "/"
            Height          =   240
            Index           =   0
            Left            =   780
            RightToLeft     =   -1  'True
            TabIndex        =   54
            Top             =   90
            Width           =   210
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "«·≈Ã„«·Ï"
            Height          =   255
            Index           =   3
            Left            =   9240
            RightToLeft     =   -1  'True
            TabIndex        =   53
            Top             =   75
            Width           =   600
         End
      End
      Begin C1SizerLibCtl.C1Elastic Ele 
         Height          =   540
         Index           =   1
         Left            =   15
         TabIndex        =   40
         TabStop         =   0   'False
         Top             =   6915
         Width           =   9855
         _cx             =   17383
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
            Left            =   8820
            TabIndex        =   41
            Top             =   90
            Width           =   975
            _ExtentX        =   1720
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
            ColorToggledText=   -2147483631
            ColorToggledHoverText=   16711680
            ColorTextShadow =   4210752
         End
         Begin ImpulseButton.ISButton Cmd 
            Height          =   375
            Index           =   1
            Left            =   7680
            TabIndex        =   42
            Top             =   90
            Width           =   1035
            _ExtentX        =   1826
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
            Left            =   6600
            TabIndex        =   43
            Top             =   90
            Width           =   990
            _ExtentX        =   1746
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
            Left            =   5550
            TabIndex        =   44
            Top             =   90
            Width           =   960
            _ExtentX        =   1693
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
            Left            =   4350
            TabIndex        =   45
            Top             =   90
            Width           =   1080
            _ExtentX        =   1905
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
            Left            =   3285
            TabIndex        =   46
            Top             =   90
            Width           =   975
            _ExtentX        =   1720
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
            TabIndex        =   47
            TabStop         =   0   'False
            Top             =   90
            Width           =   960
            _ExtentX        =   1693
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
            Left            =   2190
            TabIndex        =   48
            Top             =   90
            Width           =   975
            _ExtentX        =   1720
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
            Left            =   1080
            TabIndex        =   49
            Top             =   90
            Width           =   1020
            _ExtentX        =   1799
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
         Height          =   1800
         Index           =   0
         Left            =   15
         TabIndex        =   25
         TabStop         =   0   'False
         Top             =   645
         Width           =   9855
         _cx             =   17383
         _cy             =   3175
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
         Begin VB.CommandButton cmdCommand1 
            Caption         =   "Command1"
            Height          =   495
            Left            =   3960
            RightToLeft     =   -1  'True
            TabIndex        =   165
            Top             =   1080
            Width           =   1215
         End
         Begin VB.TextBox TxtCashCustomerName 
            Alignment       =   1  'Right Justify
            Height          =   345
            Left            =   5535
            RightToLeft     =   -1  'True
            TabIndex        =   4
            Top             =   1080
            Width           =   2910
         End
         Begin VB.TextBox TxtEmployeeID 
            Alignment       =   1  'Right Justify
            Height          =   315
            Left            =   1965
            RightToLeft     =   -1  'True
            TabIndex        =   7
            Top             =   60
            Width           =   555
         End
         Begin VB.TextBox TxtStoreID 
            Alignment       =   1  'Right Justify
            Height          =   315
            Left            =   7905
            RightToLeft     =   -1  'True
            TabIndex        =   5
            Top             =   1440
            Width           =   540
         End
         Begin VB.TextBox TxtCusID 
            Alignment       =   1  'Right Justify
            Height          =   315
            Left            =   7905
            RightToLeft     =   -1  'True
            TabIndex        =   2
            Top             =   765
            Width           =   540
         End
         Begin VB.ComboBox CboSaleType 
            Height          =   315
            Left            =   120
            RightToLeft     =   -1  'True
            Style           =   2  'Dropdown List
            TabIndex        =   12
            Top             =   1050
            Width           =   2400
         End
         Begin VB.TextBox TxtTransSerial 
            Alignment       =   1  'Right Justify
            Height          =   345
            Left            =   6435
            RightToLeft     =   -1  'True
            TabIndex        =   0
            Top             =   60
            Width           =   2010
         End
         Begin C1SizerLibCtl.C1Elastic Ele 
            Height          =   735
            Index           =   8
            Left            =   3735
            TabIndex        =   66
            TabStop         =   0   'False
            Top             =   0
            Width           =   2655
            _cx             =   4683
            _cy             =   1296
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
            CaptionPos      =   6
            WordWrap        =   -1  'True
            MaxChildSize    =   0
            MinChildSize    =   0
            TagWidth        =   0
            TagPosition     =   0
            Style           =   1
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
            Begin ImpulseButton.ISButton CmdInvProfit 
               Height          =   390
               Left            =   60
               TabIndex        =   67
               Top             =   165
               Width           =   585
               _ExtentX        =   1032
               _ExtentY        =   688
               ButtonStyle     =   1
               ButtonPositionImage=   1
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
               ButtonImage     =   "FrmSaleBill.frx":1DCA
               ColorButton     =   14871017
               DrawFocusRectangle=   0   'False
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "‰”»… «·—»Õ"
               ForeColor       =   &H00C00000&
               Height          =   255
               Index           =   23
               Left            =   1680
               RightToLeft     =   -1  'True
               TabIndex        =   72
               Top             =   420
               Width           =   885
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "ﬁÌ„… «·—»Õ"
               ForeColor       =   &H00C00000&
               Height          =   255
               Index           =   22
               Left            =   1680
               RightToLeft     =   -1  'True
               TabIndex        =   71
               Top             =   150
               Width           =   885
            End
            Begin VB.Label lblInvPrecent 
               Alignment       =   2  'Center
               BackColor       =   &H00E2E9E9&
               Caption         =   "0"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   9.75
                  Charset         =   178
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000000&
               Height          =   240
               Left            =   735
               RightToLeft     =   -1  'True
               TabIndex        =   70
               Top             =   390
               Width           =   1095
            End
            Begin VB.Label LblInvProfit 
               Alignment       =   2  'Center
               BackColor       =   &H00E2E9E9&
               Caption         =   "0"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   9.75
                  Charset         =   178
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000000&
               Height          =   240
               Left            =   735
               RightToLeft     =   -1  'True
               TabIndex        =   69
               Top             =   135
               Width           =   1095
            End
         End
         Begin VB.ComboBox XPCboDiscountType 
            Height          =   315
            Left            =   1440
            RightToLeft     =   -1  'True
            Style           =   2  'Dropdown List
            TabIndex        =   10
            Top             =   720
            Width           =   1080
         End
         Begin VB.TextBox XPTxtDiscountVal 
            Alignment       =   1  'Right Justify
            Height          =   315
            Left            =   120
            RightToLeft     =   -1  'True
            TabIndex        =   11
            Top             =   720
            Width           =   900
         End
         Begin VB.ComboBox CboPayMentType 
            Height          =   315
            Left            =   120
            RightToLeft     =   -1  'True
            Style           =   2  'Dropdown List
            TabIndex        =   9
            Top             =   390
            Width           =   2400
         End
         Begin MSDataListLib.DataCombo DBCboClientName 
            Height          =   315
            Left            =   5535
            TabIndex        =   3
            Top             =   765
            Width           =   2370
            _ExtentX        =   4180
            _ExtentY        =   556
            _Version        =   393216
            ListField       =   "6"
            BoundColumn     =   ""
            Text            =   ""
            RightToLeft     =   -1  'True
         End
         Begin MSDataListLib.DataCombo DCboStoreName 
            Height          =   315
            Left            =   5535
            TabIndex        =   6
            Top             =   1440
            Width           =   2370
            _ExtentX        =   4180
            _ExtentY        =   556
            _Version        =   393216
            ListField       =   "7"
            Text            =   ""
            RightToLeft     =   -1  'True
         End
         Begin MSComCtl2.DTPicker XPDtbBill 
            Height          =   345
            Left            =   6450
            TabIndex        =   1
            Top             =   420
            Width           =   1980
            _ExtentX        =   3493
            _ExtentY        =   609
            _Version        =   393216
            Format          =   29556737
            CurrentDate     =   38784
         End
         Begin ImpulseButton.ISButton XPBtnNewClients 
            Height          =   360
            Left            =   4995
            TabIndex        =   30
            TabStop         =   0   'False
            Top             =   750
            Width           =   480
            _ExtentX        =   847
            _ExtentY        =   635
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
            ButtonImage     =   "FrmSaleBill.frx":2164
            ColorButton     =   14871017
            ColorHighlight  =   16777215
            ColorShadow     =   -2147483631
            ColorOutline    =   -2147483631
            DrawFocusRectangle=   0   'False
         End
         Begin MSDataListLib.DataCombo DcboEmp 
            Height          =   315
            Left            =   120
            TabIndex        =   8
            Top             =   45
            Width           =   1815
            _ExtentX        =   3201
            _ExtentY        =   556
            _Version        =   393216
            ListField       =   "7"
            Text            =   ""
            RightToLeft     =   -1  'True
         End
         Begin ImpulseButton.ISButton CmdCash 
            Height          =   270
            Index           =   0
            Left            =   5055
            TabIndex        =   85
            TabStop         =   0   'False
            Top             =   1140
            Width           =   330
            _ExtentX        =   582
            _ExtentY        =   476
            ButtonStyle     =   1
            ButtonPositionImage=   4
            Caption         =   ""
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
            BackStyle       =   0
            ButtonImage     =   "FrmSaleBill.frx":24FE
            ColorButton     =   14871017
            ColorHighlight  =   16777215
            ColorShadow     =   -2147483631
            ColorOutline    =   -2147483631
            DrawFocusRectangle=   0   'False
         End
         Begin ImpulseButton.ISButton CmdCash 
            Height          =   270
            Index           =   1
            Left            =   4785
            TabIndex        =   86
            TabStop         =   0   'False
            Top             =   1140
            Visible         =   0   'False
            Width           =   315
            _ExtentX        =   556
            _ExtentY        =   476
            ButtonStyle     =   1
            ButtonPositionImage=   4
            Caption         =   ""
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
            BackStyle       =   0
            ButtonImage     =   "FrmSaleBill.frx":2898
            ColorButton     =   14871017
            ColorHighlight  =   16777215
            ColorShadow     =   -2147483631
            ColorOutline    =   -2147483631
            DrawFocusRectangle=   0   'False
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "«”„ «·⁄„Ì· «·‰ﬁœÏ"
            Height          =   300
            Index           =   33
            Left            =   8490
            RightToLeft     =   -1  'True
            TabIndex        =   84
            Top             =   1140
            Width           =   1320
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "”Ì«”… «·»Ì⁄"
            Height          =   240
            Index           =   32
            Left            =   2535
            RightToLeft     =   -1  'True
            TabIndex        =   78
            Top             =   1050
            Width           =   1140
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "«”„ «·„ÊŸ›"
            Height          =   255
            Index           =   25
            Left            =   2535
            RightToLeft     =   -1  'True
            TabIndex        =   65
            Top             =   75
            Width           =   1140
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "‰Ê⁄ «·Œ’„"
            Height          =   315
            Index           =   10
            Left            =   2535
            RightToLeft     =   -1  'True
            TabIndex        =   33
            Top             =   750
            Width           =   1140
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "ÿ—Ìﬁ… «·œ›⁄"
            Height          =   315
            Index           =   9
            Left            =   2535
            RightToLeft     =   -1  'True
            TabIndex        =   32
            Top             =   390
            Width           =   1140
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "ﬁÌ„…"
            Height          =   330
            Index           =   8
            Left            =   1035
            RightToLeft     =   -1  'True
            TabIndex        =   31
            Top             =   720
            Width           =   360
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "«”„ «·„Œ“‰"
            Height          =   255
            Index           =   24
            Left            =   8865
            RightToLeft     =   -1  'True
            TabIndex        =   29
            Top             =   1470
            Width           =   960
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "«”„ «·⁄„Ì·"
            Height          =   300
            Index           =   7
            Left            =   8430
            RightToLeft     =   -1  'True
            TabIndex        =   28
            Top             =   780
            Width           =   1350
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   " «—ÌŒ «·›« Ê—…"
            Height          =   285
            Index           =   6
            Left            =   8385
            RightToLeft     =   -1  'True
            TabIndex        =   27
            Top             =   420
            Width           =   1380
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "—ﬁ„ «·›« Ê—…"
            Height          =   255
            Index           =   5
            Left            =   8745
            RightToLeft     =   -1  'True
            TabIndex        =   26
            Top             =   75
            Width           =   1020
         End
      End
      Begin C1SizerLibCtl.C1Tab XPTab301 
         Height          =   3990
         Left            =   15
         TabIndex        =   22
         Top             =   2460
         Width           =   9855
         _cx             =   17383
         _cy             =   7038
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
         BackColor       =   12648447
         ForeColor       =   0
         FrontTabColor   =   14871017
         BackTabColor    =   12648447
         TabOutlineColor =   -2147483632
         FrontTabForeColor=   16711680
         Caption         =   "«·√’‰«›|«·√Ê—«ﬁ «·„«·Ì…|≈” ﬁÿ«⁄«  ⁄·Ï «·›« Ê—…|ﬁÌÊœ «·ÌÊ„Ì…"
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
         Picture(0)      =   "FrmSaleBill.frx":2C32
         Picture(1)      =   "FrmSaleBill.frx":2FCC
         Picture(2)      =   "FrmSaleBill.frx":3366
         Begin C1SizerLibCtl.C1Elastic Ele 
            Height          =   3525
            Index           =   19
            Left            =   11100
            TabIndex        =   133
            TabStop         =   0   'False
            Top             =   45
            Width           =   9765
            _cx             =   17224
            _cy             =   6218
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
            Height          =   3525
            Index           =   15
            Left            =   10800
            TabIndex        =   126
            TabStop         =   0   'False
            Top             =   45
            Width           =   9765
            _cx             =   17224
            _cy             =   6218
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
            _GridInfo       =   $"FrmSaleBill.frx":3700
            AccessibleName  =   ""
            AccessibleDescription=   ""
            AccessibleValue =   ""
            AccessibleRole  =   9
            Begin C1SizerLibCtl.C1Elastic Ele 
               Height          =   495
               Index           =   18
               Left            =   15
               TabIndex        =   142
               TabStop         =   0   'False
               Top             =   1545
               Width           =   9735
               _cx             =   17171
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
               Begin VB.TextBox TxtTaxServiceValue 
                  Alignment       =   1  'Right Justify
                  Height          =   345
                  Left            =   6150
                  MaxLength       =   4
                  RightToLeft     =   -1  'True
                  TabIndex        =   148
                  Top             =   60
                  Width           =   1095
               End
               Begin VB.CheckBox ChkTaxSerivce 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "÷—»Ì… Œœ„…"
                  Height          =   225
                  Left            =   8580
                  RightToLeft     =   -1  'True
                  TabIndex        =   143
                  Top             =   120
                  Width           =   1095
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Height          =   255
                  Index           =   54
                  Left            =   4470
                  RightToLeft     =   -1  'True
                  TabIndex        =   164
                  Top             =   90
                  Width           =   1275
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
                  TabIndex        =   153
                  Top             =   90
                  Width           =   375
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "«·ﬁÌ„…"
                  Enabled         =   0   'False
                  Height          =   315
                  Index           =   43
                  Left            =   7245
                  RightToLeft     =   -1  'True
                  TabIndex        =   149
                  Top             =   90
                  Width           =   480
               End
            End
            Begin C1SizerLibCtl.C1Elastic Ele 
               Height          =   495
               Index           =   17
               Left            =   15
               TabIndex        =   140
               TabStop         =   0   'False
               Top             =   1035
               Width           =   9735
               _cx             =   17171
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
               Begin VB.TextBox TxtTaxStampValue 
                  Alignment       =   1  'Right Justify
                  Height          =   345
                  Left            =   6150
                  MaxLength       =   4
                  RightToLeft     =   -1  'True
                  TabIndex        =   147
                  Top             =   60
                  Width           =   1095
               End
               Begin VB.CheckBox ChkTaxStamp 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "œ„€…"
                  Height          =   255
                  Left            =   8520
                  RightToLeft     =   -1  'True
                  TabIndex        =   141
                  Top             =   90
                  Width           =   1095
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Height          =   255
                  Index           =   53
                  Left            =   4470
                  RightToLeft     =   -1  'True
                  TabIndex        =   163
                  Top             =   90
                  Width           =   1275
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
                  Left            =   5775
                  RightToLeft     =   -1  'True
                  TabIndex        =   154
                  Top             =   90
                  Width           =   315
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "«·ﬁÌ„…"
                  Enabled         =   0   'False
                  Height          =   285
                  Index           =   41
                  Left            =   7245
                  RightToLeft     =   -1  'True
                  TabIndex        =   145
                  Top             =   90
                  Width           =   540
               End
            End
            Begin C1SizerLibCtl.C1Elastic Ele 
               Height          =   495
               Index           =   16
               Left            =   15
               TabIndex        =   138
               TabStop         =   0   'False
               Top             =   525
               Width           =   9735
               _cx             =   17171
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
               Begin VB.TextBox TxtTaxAddValue 
                  Alignment       =   1  'Right Justify
                  Height          =   345
                  Left            =   6150
                  MaxLength       =   4
                  RightToLeft     =   -1  'True
                  TabIndex        =   146
                  Top             =   60
                  Width           =   1095
               End
               Begin VB.CheckBox ChkTaxAdd 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "÷—»Ì»… Œ’„ Ê≈÷«›… (√—»«Õ  Ã«—Ì…)"
                  Height          =   435
                  Left            =   7845
                  RightToLeft     =   -1  'True
                  TabIndex        =   139
                  Top             =   30
                  Width           =   1770
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Height          =   255
                  Index           =   52
                  Left            =   4470
                  RightToLeft     =   -1  'True
                  TabIndex        =   162
                  Top             =   90
                  Width           =   1275
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
                  Left            =   5775
                  RightToLeft     =   -1  'True
                  TabIndex        =   152
                  Top             =   90
                  Width           =   375
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "«·ﬁÌ„…"
                  Enabled         =   0   'False
                  Height          =   285
                  Index           =   39
                  Left            =   7305
                  RightToLeft     =   -1  'True
                  TabIndex        =   144
                  Top             =   90
                  Width           =   480
               End
            End
            Begin VB.TextBox TxtBillComment 
               Alignment       =   1  'Right Justify
               Height          =   1035
               Left            =   15
               MultiLine       =   -1  'True
               RightToLeft     =   -1  'True
               ScrollBars      =   2  'Vertical
               TabIndex        =   127
               Top             =   2475
               Width           =   9735
            End
            Begin C1SizerLibCtl.C1Elastic Ele 
               Height          =   495
               Index           =   4
               Left            =   15
               TabIndex        =   134
               TabStop         =   0   'False
               Top             =   15
               Width           =   9735
               _cx             =   17171
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
               Begin VB.CheckBox XPChkTAX 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "÷—»Ì»… «·„»Ì⁄« "
                  Height          =   225
                  Left            =   8160
                  RightToLeft     =   -1  'True
                  TabIndex        =   136
                  Top             =   120
                  Width           =   1455
               End
               Begin VB.TextBox XPTxtTaxValue 
                  Alignment       =   1  'Right Justify
                  Height          =   345
                  Left            =   6150
                  MaxLength       =   4
                  RightToLeft     =   -1  'True
                  TabIndex        =   135
                  Top             =   75
                  Width           =   1095
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Height          =   255
                  Index           =   51
                  Left            =   4470
                  RightToLeft     =   -1  'True
                  TabIndex        =   161
                  Top             =   90
                  Width           =   1275
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
                  Left            =   5775
                  RightToLeft     =   -1  'True
                  TabIndex        =   151
                  Top             =   90
                  Width           =   375
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "«·ﬁÌ„…"
                  Enabled         =   0   'False
                  Height          =   165
                  Index           =   4
                  Left            =   7305
                  RightToLeft     =   -1  'True
                  TabIndex        =   137
                  Top             =   135
                  Width           =   480
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
               TabIndex        =   150
               Top             =   2055
               Width           =   9735
            End
         End
         Begin C1SizerLibCtl.C1Elastic Ele 
            Height          =   3525
            Index           =   7
            Left            =   45
            TabIndex        =   23
            TabStop         =   0   'False
            Top             =   45
            Width           =   9765
            _cx             =   17224
            _cy             =   6218
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
            _GridInfo       =   $"FrmSaleBill.frx":3775
            AccessibleName  =   ""
            AccessibleDescription=   ""
            AccessibleValue =   ""
            AccessibleRole  =   9
            Begin C1SizerLibCtl.C1Elastic Ele 
               Height          =   690
               Index           =   2
               Left            =   30
               TabIndex        =   24
               TabStop         =   0   'False
               Top             =   30
               Width           =   9705
               _cx             =   17119
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
                  Left            =   3990
                  RightToLeft     =   -1  'True
                  Style           =   2  'Dropdown List
                  TabIndex        =   16
                  Top             =   315
                  Width           =   1290
               End
               Begin VB.TextBox TxtQuantity 
                  Alignment       =   1  'Right Justify
                  Enabled         =   0   'False
                  Height          =   360
                  Left            =   915
                  MaxLength       =   10
                  RightToLeft     =   -1  'True
                  TabIndex        =   18
                  Top             =   315
                  Width           =   870
               End
               Begin VB.TextBox TxtSerial 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00FFFFFF&
                  Enabled         =   0   'False
                  Height          =   360
                  Left            =   2145
                  MaxLength       =   20
                  RightToLeft     =   -1  'True
                  TabIndex        =   17
                  Top             =   315
                  Width           =   1785
               End
               Begin VB.TextBox TxtPrice 
                  Alignment       =   1  'Right Justify
                  Height          =   360
                  Left            =   300
                  MaxLength       =   10
                  RightToLeft     =   -1  'True
                  TabIndex        =   19
                  Top             =   315
                  Width           =   555
               End
               Begin MSDataListLib.DataCombo DCboItemsName 
                  Height          =   315
                  Left            =   5280
                  TabIndex        =   15
                  Top             =   315
                  Width           =   2760
                  _ExtentX        =   4868
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
               Begin MSDataListLib.DataCombo DCboItemsCode 
                  Height          =   315
                  Left            =   8115
                  TabIndex        =   14
                  Top             =   315
                  Width           =   1530
                  _ExtentX        =   2699
                  _ExtentY        =   556
                  _Version        =   393216
                  Text            =   ""
                  RightToLeft     =   -1  'True
               End
               Begin ImpulseButton.ISButton CmdAdd 
                  Height          =   360
                  Left            =   60
                  TabIndex        =   20
                  Top             =   315
                  Width           =   240
                  _ExtentX        =   423
                  _ExtentY        =   635
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
                  ButtonImage     =   "FrmSaleBill.frx":37E7
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
               Begin ImpulseButton.ISButton CmdSearch 
                  Height          =   285
                  Left            =   1785
                  TabIndex        =   77
                  Top             =   330
                  Width           =   300
                  _ExtentX        =   529
                  _ExtentY        =   503
                  ButtonStyle     =   1
                  ButtonPositionImage=   1
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
                  ButtonImage     =   "FrmSaleBill.frx":3B81
                  ColorButton     =   14871017
                  DrawFocusRectangle=   0   'False
               End
               Begin VB.Label lbl 
                  Alignment       =   2  'Center
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "ﬂÊœ «·’‰›"
                  Height          =   285
                  Index           =   31
                  Left            =   8115
                  RightToLeft     =   -1  'True
                  TabIndex        =   64
                  Top             =   45
                  Width           =   1530
               End
               Begin VB.Label lbl 
                  Alignment       =   2  'Center
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "≈”„ «·’‰›"
                  Height          =   285
                  Index           =   30
                  Left            =   5280
                  RightToLeft     =   -1  'True
                  TabIndex        =   63
                  Top             =   15
                  Width           =   2760
               End
               Begin VB.Label lbl 
                  Alignment       =   2  'Center
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "Õ«·… «·’‰›"
                  Height          =   285
                  Index           =   29
                  Left            =   3990
                  RightToLeft     =   -1  'True
                  TabIndex        =   62
                  Top             =   15
                  Width           =   1290
               End
               Begin VB.Label lbl 
                  Alignment       =   2  'Center
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "«·”Ì—Ì«·"
                  Height          =   285
                  Index           =   28
                  Left            =   2085
                  RightToLeft     =   -1  'True
                  TabIndex        =   61
                  Top             =   15
                  Width           =   1845
               End
               Begin VB.Label lbl 
                  Alignment       =   2  'Center
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "«·ﬂ„Ì…"
                  Height          =   285
                  Index           =   27
                  Left            =   990
                  RightToLeft     =   -1  'True
                  TabIndex        =   60
                  Top             =   45
                  Width           =   795
               End
               Begin VB.Label lbl 
                  Alignment       =   2  'Center
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "«·”⁄—"
                  Height          =   285
                  Index           =   26
                  Left            =   300
                  RightToLeft     =   -1  'True
                  TabIndex        =   59
                  Top             =   15
                  Width           =   555
               End
            End
            Begin VSFlex8UCtl.VSFlexGrid FG 
               Height          =   2385
               Left            =   30
               TabIndex        =   13
               Top             =   735
               Width           =   9705
               _cx             =   17119
               _cy             =   4207
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
               SelectionMode   =   0
               GridLines       =   1
               GridLinesFixed  =   2
               GridLineWidth   =   1
               Rows            =   2
               Cols            =   16
               FixedRows       =   1
               FixedCols       =   1
               RowHeightMin    =   300
               RowHeightMax    =   0
               ColWidthMin     =   0
               ColWidthMax     =   0
               ExtendLastCol   =   -1  'True
               FormatString    =   $"FrmSaleBill.frx":3F1B
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
               TabIndex        =   75
               Top             =   3135
               Width           =   8775
               _ExtentX        =   15478
               _ExtentY        =   1058
               ButtonWidth     =   609
               ButtonHeight    =   953
               Appearance      =   1
               _Version        =   393216
            End
            Begin VB.Label LblItemsCount 
               Alignment       =   2  'Center
               BackColor       =   &H00404040&
               ForeColor       =   &H0000FFFF&
               Height          =   360
               Left            =   30
               RightToLeft     =   -1  'True
               TabIndex        =   76
               Top             =   3135
               Width           =   450
            End
         End
         Begin C1SizerLibCtl.C1Elastic Ele 
            Height          =   3525
            Index           =   5
            Left            =   10500
            TabIndex        =   34
            TabStop         =   0   'False
            Top             =   45
            Width           =   9765
            _cx             =   17224
            _cy             =   6218
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
            _GridInfo       =   $"FrmSaleBill.frx":4185
            AccessibleName  =   ""
            AccessibleDescription=   ""
            AccessibleValue =   ""
            AccessibleRole  =   9
            Begin C1SizerLibCtl.C1Elastic Ele 
               Height          =   1440
               Index           =   10
               Left            =   0
               TabIndex        =   79
               TabStop         =   0   'False
               Top             =   2085
               Width           =   9765
               _cx             =   17224
               _cy             =   2540
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
               _GridInfo       =   $"FrmSaleBill.frx":41F2
               AccessibleName  =   ""
               AccessibleDescription=   ""
               AccessibleValue =   ""
               AccessibleRole  =   9
               Begin C1SizerLibCtl.C1Elastic Ele 
                  Height          =   375
                  Index           =   14
                  Left            =   15
                  TabIndex        =   119
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
                     TabIndex        =   120
                     Top             =   60
                     Width           =   870
                  End
                  Begin ImpulseButton.ISButton CmdCheque 
                     Height          =   285
                     Left            =   2610
                     TabIndex        =   130
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
                     Index           =   19
                     Left            =   6750
                     RightToLeft     =   -1  'True
                     TabIndex        =   132
                     Top             =   60
                     Width           =   870
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
                     TabIndex        =   131
                     Top             =   60
                     Width           =   1080
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
                     TabIndex        =   122
                     Top             =   60
                     Width           =   1650
                  End
                  Begin VB.Label lbl 
                     Alignment       =   1  'Right Justify
                     BackColor       =   &H00E2E9E9&
                     Height          =   285
                     Index           =   18
                     Left            =   4185
                     RightToLeft     =   -1  'True
                     TabIndex        =   121
                     Top             =   60
                     Width           =   870
                  End
               End
               Begin VSFlex8UCtl.VSFlexGrid FgCheques 
                  Height          =   1020
                  Left            =   1740
                  TabIndex        =   80
                  Top             =   405
                  Width           =   8010
                  _cx             =   14129
                  _cy             =   1799
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
                  FormatString    =   $"FrmSaleBill.frx":426C
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
               Height          =   1545
               Index           =   6
               Left            =   0
               TabIndex        =   87
               TabStop         =   0   'False
               Top             =   540
               Width           =   9765
               _cx             =   17224
               _cy             =   2725
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
               _GridInfo       =   $"FrmSaleBill.frx":43A0
               AccessibleName  =   ""
               AccessibleDescription=   ""
               AccessibleValue =   ""
               AccessibleRole  =   9
               Begin VSFlex8UCtl.VSFlexGrid FgInstallments 
                  Height          =   705
                  Left            =   1815
                  TabIndex        =   88
                  Top             =   465
                  Width           =   7935
                  _cx             =   13996
                  _cy             =   1244
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
                  FormatString    =   $"FrmSaleBill.frx":4411
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
                  TabIndex        =   89
                  TabStop         =   0   'False
                  Top             =   1185
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
                     TabIndex        =   129
                     Top             =   60
                     Width           =   1020
                  End
                  Begin VB.Label LblStartValue 
                     Alignment       =   1  'Right Justify
                     BackColor       =   &H00E2E9E9&
                     Height          =   225
                     Left            =   30
                     RightToLeft     =   -1  'True
                     TabIndex        =   128
                     Top             =   60
                     Width           =   225
                  End
                  Begin VB.Label LblInstallSeprator 
                     Alignment       =   1  'Right Justify
                     BackColor       =   &H00E2E9E9&
                     ForeColor       =   &H00FF0000&
                     Height          =   225
                     Left            =   1980
                     RightToLeft     =   -1  'True
                     TabIndex        =   125
                     Top             =   60
                     Width           =   240
                  End
                  Begin VB.Label LblPrecenValue 
                     Alignment       =   1  'Right Justify
                     BackColor       =   &H00E2E9E9&
                     Height          =   225
                     Left            =   7545
                     RightToLeft     =   -1  'True
                     TabIndex        =   124
                     Top             =   60
                     Width           =   270
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
                     TabIndex        =   123
                     Top             =   60
                     Width           =   435
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
                     TabIndex        =   99
                     Top             =   60
                     Width           =   750
                  End
                  Begin VB.Label LblPrecenType 
                     Alignment       =   1  'Right Justify
                     BackColor       =   &H00E2E9E9&
                     Height          =   225
                     Left            =   8280
                     RightToLeft     =   -1  'True
                     TabIndex        =   98
                     Top             =   60
                     Width           =   645
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
                     TabIndex        =   97
                     Top             =   60
                     Width           =   870
                  End
                  Begin VB.Label LblInstallTotal 
                     Alignment       =   1  'Right Justify
                     BackColor       =   &H00E2E9E9&
                     Height          =   225
                     Left            =   6030
                     RightToLeft     =   -1  'True
                     TabIndex        =   96
                     Top             =   60
                     Width           =   585
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
                     Left            =   5085
                     RightToLeft     =   -1  'True
                     TabIndex        =   95
                     Top             =   60
                     Width           =   930
                  End
                  Begin VB.Label LblInstallCount 
                     Alignment       =   1  'Right Justify
                     BackColor       =   &H00E2E9E9&
                     Height          =   225
                     Left            =   4740
                     RightToLeft     =   -1  'True
                     TabIndex        =   94
                     Top             =   60
                     Width           =   330
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
                     TabIndex        =   93
                     Top             =   60
                     Width           =   690
                  End
                  Begin VB.Label LblFirstInstallDate 
                     Alignment       =   1  'Right Justify
                     BackColor       =   &H00E2E9E9&
                     Height          =   225
                     Left            =   3240
                     RightToLeft     =   -1  'True
                     TabIndex        =   92
                     Top             =   60
                     Width           =   765
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
                  Begin VB.Label LblInstallmentType 
                     Alignment       =   1  'Right Justify
                     BackColor       =   &H00E2E9E9&
                     ForeColor       =   &H00FF0000&
                     Height          =   225
                     Left            =   1320
                     RightToLeft     =   -1  'True
                     TabIndex        =   90
                     Top             =   60
                     Width           =   645
                  End
               End
               Begin C1SizerLibCtl.C1Elastic Ele 
                  Height          =   435
                  Index           =   12
                  Left            =   15
                  TabIndex        =   100
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
                  Begin VB.CheckBox ChkInstall 
                     Alignment       =   1  'Right Justify
                     BackColor       =   &H00E2E9E9&
                     Caption         =   " ﬁ”Ìÿ"
                     Height          =   360
                     Left            =   1815
                     RightToLeft     =   -1  'True
                     TabIndex        =   104
                     Top             =   15
                     Width           =   855
                  End
                  Begin VB.TextBox XPTxtSerial 
                     Alignment       =   1  'Right Justify
                     Height          =   360
                     Index           =   1
                     Left            =   5280
                     Locked          =   -1  'True
                     RightToLeft     =   -1  'True
                     TabIndex        =   103
                     Top             =   30
                     Width           =   1260
                  End
                  Begin VB.TextBox XPTxtValue 
                     Alignment       =   1  'Right Justify
                     Height          =   360
                     Index           =   1
                     Left            =   7185
                     MaxLength       =   10
                     RightToLeft     =   -1  'True
                     TabIndex        =   102
                     Top             =   30
                     Width           =   1110
                  End
                  Begin VB.CheckBox XPChkPayType 
                     Alignment       =   1  'Right Justify
                     BackColor       =   &H00E2E9E9&
                     Caption         =   "¬Ã· "
                     Height          =   330
                     Index           =   1
                     Left            =   8820
                     RightToLeft     =   -1  'True
                     TabIndex        =   101
                     Top             =   30
                     Width           =   870
                  End
                  Begin ImpulseButton.ISButton CmdINSTALLMENT 
                     Height          =   435
                     Left            =   60
                     TabIndex        =   105
                     Top             =   -15
                     Width           =   1725
                     _ExtentX        =   3043
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
                     ButtonImage     =   "FrmSaleBill.frx":44E2
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
                     Left            =   2685
                     TabIndex        =   106
                     Top             =   30
                     Width           =   1320
                     _ExtentX        =   2328
                     _ExtentY        =   609
                     _Version        =   393216
                     Format          =   29556737
                     CurrentDate     =   38784
                  End
                  Begin VB.Label lbl 
                     Alignment       =   1  'Right Justify
                     BackColor       =   &H00E2E9E9&
                     Caption         =   " «—ÌŒ «·«” Õﬁ«ﬁ"
                     Height          =   300
                     Index           =   21
                     Left            =   4065
                     RightToLeft     =   -1  'True
                     TabIndex        =   109
                     Top             =   75
                     Width           =   1155
                  End
                  Begin VB.Label lbl 
                     Alignment       =   1  'Right Justify
                     BackColor       =   &H00E2E9E9&
                     Caption         =   "«·ﬁÌ„…"
                     Height          =   345
                     Index           =   15
                     Left            =   8310
                     RightToLeft     =   -1  'True
                     TabIndex        =   108
                     Top             =   75
                     Width           =   420
                  End
                  Begin VB.Label lbl 
                     Alignment       =   1  'Right Justify
                     BackColor       =   &H00E2E9E9&
                     Caption         =   "„”·”·"
                     Height          =   330
                     Index           =   14
                     Left            =   6600
                     RightToLeft     =   -1  'True
                     TabIndex        =   107
                     Top             =   75
                     Width           =   540
                  End
               End
            End
            Begin C1SizerLibCtl.C1Elastic Ele 
               Height          =   540
               Index           =   11
               Left            =   0
               TabIndex        =   110
               TabStop         =   0   'False
               Top             =   0
               Width           =   9765
               _cx             =   17224
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
               Begin VB.TextBox XPTxtValue 
                  Alignment       =   1  'Right Justify
                  Height          =   375
                  Index           =   0
                  Left            =   7200
                  MaxLength       =   10
                  RightToLeft     =   -1  'True
                  TabIndex        =   113
                  Top             =   60
                  Width           =   1095
               End
               Begin VB.TextBox XPTxtSerial 
                  Alignment       =   1  'Right Justify
                  Height          =   375
                  Index           =   0
                  Left            =   5280
                  Locked          =   -1  'True
                  RightToLeft     =   -1  'True
                  TabIndex        =   112
                  Top             =   60
                  Width           =   1275
               End
               Begin VB.CheckBox XPChkPayType 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "‰ﬁœ«"
                  Height          =   345
                  Index           =   0
                  Left            =   8940
                  RightToLeft     =   -1  'True
                  TabIndex        =   111
                  Top             =   90
                  Width           =   720
               End
               Begin MSDataListLib.DataCombo DcboBox 
                  Height          =   315
                  Left            =   2340
                  TabIndex        =   114
                  Top             =   105
                  Width           =   2055
                  _ExtentX        =   3625
                  _ExtentY        =   556
                  _Version        =   393216
                  Text            =   ""
                  RightToLeft     =   -1  'True
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "ÿ—Ìﬁ… «·œ›⁄"
                  Height          =   345
                  Index           =   20
                  Left            =   1110
                  RightToLeft     =   -1  'True
                  TabIndex        =   118
                  Top             =   120
                  Visible         =   0   'False
                  Width           =   1095
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "«·ﬁÌ„…"
                  Height          =   345
                  Index           =   13
                  Left            =   8295
                  RightToLeft     =   -1  'True
                  TabIndex        =   117
                  Top             =   90
                  Width           =   450
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "„”·”·"
                  Height          =   345
                  Index           =   12
                  Left            =   6555
                  RightToLeft     =   -1  'True
                  TabIndex        =   116
                  Top             =   90
                  Width           =   615
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "«”„ «·Œ“‰…"
                  Height          =   345
                  Index           =   11
                  Left            =   4350
                  RightToLeft     =   -1  'True
                  TabIndex        =   115
                  Top             =   90
                  Width           =   870
               End
            End
         End
      End
      Begin C1SizerLibCtl.C1Elastic Ele 
         Height          =   615
         Index           =   9
         Left            =   15
         TabIndex        =   35
         TabStop         =   0   'False
         Top             =   15
         Width           =   9180
         _cx             =   16193
         _cy             =   1085
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
         Caption         =   "›« Ê—… «·»Ì⁄ "
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
            Height          =   285
            Left            =   3915
            RightToLeft     =   -1  'True
            TabIndex        =   156
            Top             =   0
            Visible         =   0   'False
            Width           =   345
         End
         Begin VB.TextBox TxtFillData 
            Alignment       =   1  'Right Justify
            BackColor       =   &H000000FF&
            Height          =   285
            Left            =   3540
            RightToLeft     =   -1  'True
            TabIndex        =   155
            Top             =   0
            Visible         =   0   'False
            Width           =   345
         End
         Begin VB.TextBox XPTxtBillID 
            Alignment       =   1  'Right Justify
            BackColor       =   &H000000FF&
            Height          =   285
            Left            =   3150
            Locked          =   -1  'True
            RightToLeft     =   -1  'True
            TabIndex        =   68
            Top             =   0
            Visible         =   0   'False
            Width           =   360
         End
         Begin ImpulseButton.ISButton XPBtnMove 
            Height          =   345
            Index           =   0
            Left            =   1455
            TabIndex        =   36
            Top             =   30
            Width           =   600
            _ExtentX        =   1058
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
            ButtonImage     =   "FrmSaleBill.frx":487C
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
            Left            =   765
            TabIndex        =   37
            Top             =   30
            Width           =   660
            _ExtentX        =   1164
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
            ButtonImage     =   "FrmSaleBill.frx":4C16
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
            Left            =   2085
            TabIndex        =   38
            Top             =   30
            Width           =   630
            _ExtentX        =   1111
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
            ButtonImage     =   "FrmSaleBill.frx":4FB0
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
            Left            =   60
            TabIndex        =   39
            Top             =   30
            Width           =   645
            _ExtentX        =   1138
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
            ButtonImage     =   "FrmSaleBill.frx":534A
            ColorHighlight  =   4194304
            ColorHoverText  =   16777215
            ColorShadow     =   -2147483631
            ColorOutline    =   -2147483631
            DrawFocusRectangle=   0   'False
            DisabledImageStyle=   1
            ColorToggledHoverText=   16777215
            ColorTextShadow =   16777215
         End
         Begin ImpulseButton.ISButton CmdNotes 
            Height          =   345
            Left            =   5835
            TabIndex        =   81
            Top             =   120
            Width           =   630
            _ExtentX        =   1111
            _ExtentY        =   609
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
            ButtonImage     =   "FrmSaleBill.frx":56E4
            DrawFocusRectangle=   0   'False
         End
         Begin ImpulseButton.ISButton CmdRetruns 
            Height          =   345
            Left            =   6510
            TabIndex        =   82
            Top             =   120
            Width           =   645
            _ExtentX        =   1138
            _ExtentY        =   609
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
            ButtonImage     =   "FrmSaleBill.frx":5A7E
            DrawFocusRectangle=   0   'False
         End
         Begin VB.Label LblShortcutKeys 
            BackColor       =   &H00FFFFFF&
            BackStyle       =   0  'Transparent
            Caption         =   "ÃœÌœ F12 Or Enter ,  ⁄œÌ· F11 , Õ›Ÿ F10 ,  —«Ã⁄ F9 ,Õ–› F8 ,»ÕÀ F7 "
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
            Height          =   195
            Left            =   90
            RightToLeft     =   -1  'True
            TabIndex        =   83
            Top             =   390
            Width           =   5370
         End
      End
   End
End
Attribute VB_Name = "FrmSaleBill"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Rs As ADODB.Recordset
Dim TTP As clstooltip
Dim NewGrid As New ClsGrid
Dim SaleReport As ClsSaleReport
Dim cSearchDcbo(4)   As clsDCboSearch
Dim Dcombos As ClsDataCombos
Public BolPrint As Boolean
Public WithEvents m_Menu1 As Menu
Attribute m_Menu1.VB_VarHelpID = -1
Dim WithEvents m_MenuRefesh As Menu
Attribute m_MenuRefesh.VB_VarHelpID = -1
Dim WithEvents m_MenuCusBalance As Menu
Attribute m_MenuCusBalance.VB_VarHelpID = -1
Dim WithEvents m_MenuViewList As Menu
Attribute m_MenuViewList.VB_VarHelpID = -1
Dim WithEvents m_MenuViewNotes As Menu
Attribute m_MenuViewNotes.VB_VarHelpID = -1
Dim WithEvents m_MenuScreenPremission As Menu
Attribute m_MenuScreenPremission.VB_VarHelpID = -1
Dim WithEvents StrCashCustomerPhone As TextBox
Attribute StrCashCustomerPhone.VB_VarHelpID = -1
Dim WithEvents StrCashCustomerMobile As TextBox
Attribute StrCashCustomerMobile.VB_VarHelpID = -1
Dim WithEvents StrCashCustomerAddress As TextBox
Attribute StrCashCustomerAddress.VB_VarHelpID = -1
Dim WithEvents m_FrmSearch As Form
Attribute m_FrmSearch.VB_VarHelpID = -1
Private Sub CboPayMentType_Change()
On Error GoTo ErrTrap
If Me.TxtModFlg.text = "N" Or Me.TxtModFlg.text = "E" Then
    If CboPayMentType.ListIndex = 0 Then
        XPChkPayType(0).Enabled = False
        XPChkPayType(1).Enabled = False
        XPChkPayType(2).Enabled = False
        XPChkPayType(0).Value = Checked
        XPChkPayType(1).Value = Unchecked
        XPChkPayType(2).Value = Unchecked
        XPTxtValue(0).text = XPTxtSum.text
        XPTxtValue(1).text = ""
    Else
        XPChkPayType(0).Enabled = True
        XPChkPayType(1).Enabled = True
        XPChkPayType(2).Enabled = True
        XPChkPayType(0).Value = Unchecked
        XPChkPayType(1).Value = Unchecked
        XPChkPayType(2).Value = Unchecked
        XPTxtValue(0).text = ""
    End If
End If
Exit Sub
ErrTrap:
End Sub
Private Sub CboPayMentType_Click()
CboPayMentType_Change
End Sub

Private Sub ChkInstall_Click()
If ChkInstall.Value = vbChecked Then
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
If ChkTaxAdd.Value = Checked Then
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
If ChkTaxSerivce.Value = Checked Then
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
If ChkTaxStamp.Value = Checked Then
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
Dim AskOption As Boolean
Dim intDef As Integer
Dim Msg As String
Dim StrSQL As String
Dim RsTest As ADODB.Recordset
Dim RsOptions As ADODB.Recordset
BolPrint = True
On Error GoTo ErrTrap
Select Case Index
    Case 0
        If DoPremis(Do_New, Me.name, True) = False Then
            Exit Sub
        End If
        
        If SystemOptions.SysRegisterState = DemoRun Then
            Set RsTest = New ADODB.Recordset
            StrSQL = "Select Count(Transaction_ID) AS CountX From Transactions"
            RsTest.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText
            If Not (RsTest.BOF Or RsTest.EOF) Then
                If RsTest("CountX").Value >= 50 Then
                    Msg = "≈‰ Â  ‰”Œ… ⁄—÷ «·»—‰«„Ã ... »—Ã«¡ «·√ ’«· »«·œ⁄„ «·›‰Ï"
                    Msg = Msg & Chr(13) & "002-0123591024 - 0226210707"
                    MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
                    Exit Sub
                End If
            End If
        End If
        
        clear_all Me
        ClearNotes
        TxtModFlg.text = "N"
        XPTxtBillID.text = CStr(new_id("Transactions", "Transaction_ID", "", True))
        SetDefaults
        NewGrid.GridDefaultValue 1
        Me.DCboUserName.BoundText = User_ID
        intDef = GetSetting(StrAppRegPath, "DefaultOptions", "DefaultClient", 2)
        DBCboClientName.BoundText = intDef
        intDef = GetSetting(StrAppRegPath, "DefaultOptions", "DefaultSaleStore", 1)
        DCboStoreName.BoundText = intDef
        Set RsOptions = New ADODB.Recordset
        RsOptions.Open "tbloptions", Cn, adOpenStatic, adLockReadOnly, adCmdTable
        If Not (RsOptions.BOF Or RsOptions.EOF) Then
            Me.DcboBox.BoundText = IIf(IsNull(RsOptions("SalesBoxID").Value), "", RsOptions("SalesBoxID").Value)
        End If
        XPTab301.CurrTab = 0
        '------------------
        Me.XPDtbBill.SetFocus
        '--------------------
    Case 1
        If DoPremis(Do_Edit, Me.name, True) = False Then
            Exit Sub
        End If
        'If AvailableDeal = True Then
            '«·√ﬁ”«ÿ «·„”Ã·… ⁄·Ï «·›« Ê—…
            If XPTxtValue(1).Tag <> "" Then
                StrSQL = "select * From InstallMent where NoteID=" & XPTxtValue(1).Tag
                Set RsTest = New ADODB.Recordset
                RsTest.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText
                If Not (RsTest.EOF Or RsTest.BOF) Then
                    Msg = "·ﬁœ  „  ﬁ”Ìÿ «·ﬁÌ„ «·¬Ã·… ⁄·Ï Â–Â «·›« Ê—…" & Chr(13)
                    Msg = Msg + " ⁄œÌ· «·›« Ê—… ”ÌƒœÌ ≈·Ï Õ–› Â–Â «·√ﬁ”«ÿ" & Chr(13)
                    Msg = Msg + "Â·  —€» ›Ì  ⁄œÌ· Â–Â «·›« Ê—…ø"
                    If MsgBox(Msg, vbYesNo + vbQuestion + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title) = vbNo Then
                        Exit Sub
                    End If
                End If
            End If
            '«·√ﬁ”«ÿ «·„”œœ… ⁄·Ï «·›« Ê—…
            If XPTxtValue(1).Tag <> "" Then
                StrSQL = "select * From ReceiptQestForBill where NoteID=" & XPTxtValue(1).Tag
                Set RsTest = New ADODB.Recordset
                RsTest.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText
                If Not (RsTest.EOF Or RsTest.BOF) Then
                    Msg = "·ﬁœ  „  Õ’Ì· »⁄÷ «·√ﬁ”«ÿ «·„”Ã·… ⁄·Ï Â–Â «·›« Ê—…" & Chr(13)
                    Msg = Msg + "Ê·« Ì„ﬂ‰  ⁄œÌ· »Ì«‰« Â«" & Chr(13)
                    Msg = Msg + "≈–« ﬂ‰   —€» ›Ì  ⁄œÌ· »Ì«‰«  Â–Â «·›« Ê—…" & Chr(13)
                    Msg = Msg + "ÌÃ» Õ–› ⁄„·Ì«  «· Õ’Ì· «·Œ«’… »Â«"
                    MsgBox Msg, vbOKOnly + vbQuestion + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
                    Exit Sub
                End If
            End If
            '⁄„·Ì«  «·’Ì«‰… «·„— »ÿ… »«·›« Ê—…
            StrSQL = "select * From MaintenanceJuncTransaction where Transaction_ID=" & Trim(XPTxtBillID.text)
            Set RsTest = New ADODB.Recordset
            RsTest.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText
            If Not (RsTest.EOF Or RsTest.BOF) Then
                Msg = "·ﬁœ  „ ≈Ã—«¡ »⁄÷ ⁄„·Ì«  «·’Ì«‰… ⁄·Ï Â–Â «·›« Ê—… Ê·« Ì„ﬂ‰  ⁄œÌ·Â«"
                Msg = Msg + "≈–« ﬂ‰   —€» ›Ì  ⁄œÌ· »Ì«‰«  Â–Â «·›« Ê—…" & Chr(13)
                Msg = Msg + "ÌÃ» Õ–› ⁄„·Ì«  «·’Ì«‰… «·Œ«’… »Â«"
                MsgBox Msg, vbOKOnly + vbQuestion + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
                Exit Sub
            End If
            TxtModFlg.text = "E"
            Me.DCboUserName.BoundText = User_ID
        'End If
    Case 2
        SaveData
    Case 3
        Undo
    Case 4
        If DoPremis(Do_Delete, Me.name, True) = False Then
            Exit Sub
        End If
       Del_TransAction
    Case 5
        If DoPremis(Do_Search, Me.name, True) = False Then
            Exit Sub
        End If
        If m_FrmSearch Is Nothing Then
            Set m_FrmSearch = New FrmBuySearch
            m_FrmSearch.DealingForm = InvoiceTransaction
            m_FrmSearch.Caption = "«·»ÕÀ ⁄‰ ⁄„·Ì… »Ì⁄"
            Set m_FrmSearch.RetrunFrm = Me
            m_FrmSearch.Show vbModeless, MDIFrmMain
        Else
            Msg = "Â‰«ﬂ ‘«‘… »ÕÀ Œ«’… »‘«‘… ›« Ê—… «·»Ì⁄ «·Õ«·Ì…"
            Msg = Msg & Chr(13) & "Ÿ«Â—… «„«„ﬂ ›⁄·«...·«Ì„ﬂ‰ ⁄—÷ «ﬂÀ— „‰ ‘«‘… »ÕÀ ·ﬂ· ‘«‘… ›« Ê—…"
            MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
            m_FrmSearch.ZOrder 0
            'm_FrmSearch.SetFocus
        End If
    Case 7
        If DoPremis(Do_Print, Me.name, True) = False Then
            Exit Sub
        End If
        If Me.XPTxtBillID.text = "" Then
            Msg = "·« ÊÃœ ›Ê« Ì— ·Ì „ ÿ»«⁄ Â«"
            MsgBox Msg, vbOKOnly + vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
            Exit Sub
        End If
        AskOption = GetSetting(StrAppRegPath, "View_Type", "ShowMe", False)
        If AskOption = False Then
            FrmSallReportOptions.Show vbModal
            If FrmSallReportOptions.UserCanceled = True Then
                Unload FrmSallReportOptions
                Exit Sub
            End If
            Unload FrmSallReportOptions
        End If
        PrintReport
    Case 6
        Unload Me
End Select
Exit Sub
ErrTrap:
End Sub


Private Sub CmdCash_Click(Index As Integer)
Select Case Index
    Case 0
    Case 1
End Select
End Sub

Private Sub CmdCheque_Click()
Load FrmChecks
FrmChecks.TxtModFlg.text = Me.TxtModFlg.text
FrmChecks.XPTxtBillID.text = Me.XPTxtBillID.text
Set FrmChecks.PutFg = Me.FgCheques
FrmChecks.Show vbModal
SumChecks

End Sub

Private Sub cmdCommand1_Click()

MsgBox Val(FG.Cell(flexcpData, 1, FG.ColIndex("UnitID")))
MsgBox Val(FG.Cell(flexcpData, 1, FG.ColIndex("ColorID")))

End Sub

Private Sub CmdHelp_Click()
SystemOptions.SysHelp.HHTopicID = Me.HelpContextID
SystemOptions.SysHelp.HHDisplayTopicID Me.hwnd
End Sub

Private Sub CmdInfo_Click()
Dim xPoint As POINTAPI
MDIFrmMain.MnuInvSalesOptions.Visible = True
MDIFrmMain.MnuInvInsertTemp.Visible = True
MDIFrmMain.MnuInvSalesOptions.Checked = Me.Ele(8).Visible
MDIFrmMain.MnuInvSales_Mnu4.Enabled = Me.CmdNotes.Visible
Me.PopupMenu MDIFrmMain.MnuInvSales, vbPopupMenuRightAlign Or vbPopupMenuRightButton

'ClientToScreen Me.CmdInfo.hwnd, xPoint
'Me.PopupMenu MDIFrmMain.MnuInvSales, , (xPoint.X * Screen.TwipsPerPixelX), (xPoint.Y * Screen.TwipsPerPixelY)
'Me.PopupMenu MDIFrmMain.MnuInvSales, vbPopupMenuRightAlign + vbPopupMenuRightButton, (xPoint.X * Screen.TwipsPerPixelX), (xPoint.Y * Screen.TwipsPerPixelY)

End Sub

Private Sub CmdINSTALLMENT_Click()
On Error GoTo ErrTrap
Dim Msg As String
Dim I As Integer
If XPTxtValue(1).text = "" Then
    Msg = "ÌÃ»  ÕœÌœ «·ﬁÌ„… «·¬Ã·… ﬁ»·  ”ÃÌ· «·√ﬁ”«ÿ"
    MsgBox Msg, vbOKOnly + vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
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
        .Retrive Val(XPTxtValue(1).Tag)
    Else
        .Tag = "N"
        .Txt(1).text = XPTxtValue(1).text
        .LblNoteID.Caption = XPTxtSerial(1).text
        .CboPrecenType.ListIndex = Val(Me.LblPrecenType.Tag)
        .Txt(3).text = Val(LblPrecenValue.Caption)
        .Txt(5).text = Val(LblInstallCount.Caption)
        If IsDate(Me.LblFirstInstallDate.Caption) Then
            .Dtp_First.Value = Me.LblFirstInstallDate.Caption
        End If
        .Txt(7).text = Val(LblInstallSeprator.Caption)
        If Val(LblInstallmentType.Tag) = 0 Then
            .OptInt(0).Value = True
        ElseIf Val(LblInstallmentType.Tag) = 1 Then
            .OptInt(1).Value = True
        ElseIf Val(LblInstallmentType.Tag) = 2 Then
            .OptInt(2).Value = True
        End If
        With .FG
            .Rows = Me.FgInstallments.Rows
            For I = 1 To Me.FgInstallments.Rows - 1
                .TextMatrix(I, .ColIndex("Serial")) = I
                .TextMatrix(I, .ColIndex("Value")) = _
                    Me.FgInstallments.TextMatrix(I, Me.FgInstallments.ColIndex("Value"))
                .TextMatrix(I, .ColIndex("Due_Date")) = _
                    Me.FgInstallments.TextMatrix(I, Me.FgInstallments.ColIndex("Due_Date"))
            Next I
            .AutoSize 0, .Cols - 1, False
        End With
    End If
    .Show vbModal
End With
Exit Sub
ErrTrap:
End Sub

Private Sub CmdInvProfit_Click()
If SystemOptions.SysMainStockCostMethod = LastPurPriceType Or _
    SystemOptions.SysMainStockCostMethod = ModernWeightAverage Then
    NewGrid.ShowInvProfDialog
End If
'If Me.TxtModFlg.Text = "R" Then
'
'Else
'    NewGrid.ShowInvProfDialog
'End If
End Sub

Private Sub CmdNotes_Click()
ShowRelatedNotes Val(Me.XPTxtBillID.text), 1
End Sub

Private Sub CmdNotes_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim StrTemp As String

If Val(Me.CmdNotes.Tag) = 0 Then
    Me.CmdNotes.ToolTipText = ""
Else
    StrTemp = " ÊÃœ ⁄·Ï Â–Â «·Õ—ﬂ… ⁄„·Ì«  „«·Ì… „ﬁœ«—Â« : " & Val(Me.CmdNotes.Tag)
    Me.CmdNotes.ToolTipText = StrTemp
End If
End Sub


Private Sub CmdRetruns_Click()
ShowRelatedTransactions Val(Me.XPTxtBillID.text), 1
End Sub

Private Sub CmdRetruns_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim StrTemp As String

If Val(Me.CmdRetruns.Tag) = 0 Then
    Me.CmdRetruns.ToolTipText = ""
Else
    StrTemp = " ÊÃœ ⁄·Ï Â–Â «·Õ—ﬂ… Õ—ﬂ«   Ã«—Ì… √Œ—Ï ·Â« ⁄·«ﬁ… »Â« ≈Ã„«·ÌÂ«: " & Val(Me.CmdRetruns.Tag)
    Me.CmdRetruns.ToolTipText = StrTemp
End If
End Sub

Private Sub CmdSearch_Click()
'Dim LngItemID As Long
'Dim LngStoreID As Long
'LngItemID = Val(Me.DCboItemsName.BoundText)
'LngStoreID = Val(Me.DCboStoreName.BoundText)
'If LngItemID = 0 Or LngStoreID = 0 Then
'    Exit Sub
'End If
'Load FrmSerialList
'FrmSerialList.RetrunType = 1
'Set FrmSerialList.m_TextBox = Me.TxtSerial
'FrmSerialList.GetData LngItemID, LngStoreID
'FrmSerialList.Show vbModal
End Sub
Private Sub DBCboClientName_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = vbRightButton Then
    MDIFrmMain.MnuCusTools.Tag = Me.DBCboClientName.BoundText
    Me.PopupMenu MDIFrmMain.MnuCusTools
End If
End Sub

Private Sub Ele_DblClick(Index As Integer)
On Error GoTo ErrTrap
If Index = 9 Then
    If Me.WindowState = vbNormal Then
        Me.WindowState = vbMaximized
    Else
        Me.WindowState = vbNormal
    End If
End If
Exit Sub
ErrTrap:
End Sub

Private Sub Form_Activate()
Set m_Menu1 = MDIFrmMain.MnuInvInsertTemp
Set m_MenuRefesh = MDIFrmMain.MnuInvSales_Refresh
Set m_MenuCusBalance = MDIFrmMain.MnuInvSales_Mnu1
Set m_MenuViewList = MDIFrmMain.MnuInvViewList
Set m_MenuViewNotes = MDIFrmMain.MnuInvSales_Mnu4
Set m_MenuScreenPremission = MDIFrmMain.MnuInvSales_Mnu7
If TxtTransSerial.Enabled = True Then
    TxtTransSerial.SetFocus
End If
End Sub



Private Sub lbl_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
If Val(lbl(Index).Caption) <> 0 Then
    lbl(Index).ToolTipText = WriteNo(lbl(Index).Caption, 0, True)
End If
End Sub


Private Sub LblInstallCount_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
LblInstallCount.ToolTipText = WriteNo(LblInstallCount.Caption, 0, True)
End Sub
Private Sub LblInstallTotal_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
LblInstallTotal.ToolTipText = WriteNo(LblInstallTotal.Caption, 0, True)
End Sub
Private Sub LblInvProfit_Change()
CalculateInvPrecent
End Sub
Private Sub LblPrecenValue_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
LblPrecenValue.ToolTipText = WriteNo(LblPrecenValue.Caption, 0, True)
End Sub
Private Sub LblTotal_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
LblTotal.ToolTipText = WriteNo(LblTotal.Caption, 0, True)
End Sub
Private Sub m_FrmSearch_Unload(Cancel As Integer)
Set m_FrmSearch = Nothing
End Sub

Private Sub m_Menu1_Click()
On Error GoTo ErrTrap
With FrmBuySearch
    .DealingForm = InsertTemplateToInvoice
    .Caption = "«·⁄—Ê÷ «·Ã«Â“…"
    .FG.TextMatrix(0, .FG.ColIndex("Transaction_ID")) = "ﬂÊœ «·⁄—÷"
    .FG.TextMatrix(0, .FG.ColIndex("BillDate")) = "«”„ «·⁄—÷"
    .FG.TextMatrix(0, .FG.ColIndex("ClientNmae")) = " «—ÌŒ «·⁄—÷"
    .FG.TextMatrix(0, .FG.ColIndex("StorName")) = "ﬁÌ„… «·⁄—÷"
    .XPChkSearchType.Visible = False
    .TxtVal.Visible = True
    .XPLbl(2).Visible = True
    .XPLbl(1).Visible = False
    .XPLbl(0).Visible = False
    .XPLbl(3).Visible = True
    .XPLbl(4).Visible = True
    .Show vbModal
End With
Exit Sub
ErrTrap:
End Sub

Private Sub m_MenuCusBalance_Click()
Dim cReport As ClsCustemerReport
Dim LngCusID As Long

With Me.FG
    If Me.DBCboClientName.BoundText = "" Then Exit Sub
    LngCusID = Val(Me.DBCboClientName.BoundText)
    OpenScreen PopUpShowCustomerBalanceScreen, LngCusID, 0
End With
End Sub

Private Sub m_MenuRefesh_Click()
Dim Msg As String
On Error GoTo ErrTrap
If Me.TxtModFlg.text <> "R" Then
    Msg = " ÕœÌÀ «·»Ì«‰«  €Ì— „ «Õ ≈·« «‰  ﬂÊ‰ «·‘«‘… ›Ï Õ«·… «·⁄—÷ ›ﬁÿ..!"
    'MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
    'Exit Sub
End If
LoadCombosData
NewGrid.FillGrid
Rs.Requery
Exit Sub
ErrTrap:
End Sub

Private Sub m_MenuScreenPremission_Click()
ShowScreenPermission Me.name
End Sub

Private Sub m_MenuViewList_Click()
Dim FrmView As FrmViewList
Dim FG As VSFlex8UCtl.vsFlexGrid
Dim StrSQL As String
Dim Rs As ADODB.Recordset
Dim StrComboList As String
Dim GrdBack As ClsBackGroundPic
Dim cProgress As ClsProgress
Dim BolFrmLoaded As Boolean
Set FrmView = New FrmViewList
Set FG = FrmView.vsfGroup1.vsFlexGrid

With FG
    .Cols = 10
    .RowHeightMin = 320
    .TextMatrix(0, 0) = "—ﬁ„ «·»—‰«„Ã"
    .TextMatrix(0, 1) = "—ﬁ„ «·›« Ê—…"
    .TextMatrix(0, 2) = " «—ÌŒ «·›« Ê—…"
    .ColDataType(2) = flexDTDate
    .TextMatrix(0, 3) = "«”„ «·⁄„Ì·"
    .TextMatrix(0, 4) = "ÿ—Ìﬁ… «·œ›⁄"
    StrComboList = "#0;‰ﬁœÏ|#1;√Ã·"
    .ColComboList(4) = StrComboList
    
    .TextMatrix(0, 5) = "«”„ «·„Œ“‰"
    .TextMatrix(0, 6) = "«”„ «·„ÊŸ›"
    
    .TextMatrix(0, 7) = "‰Ê⁄ «·Œ’„"
    .TextMatrix(0, 8) = "ﬁÌ„… «·Œ’„"
    .TextMatrix(0, 9) = "≈Ã„«·Ï «·›« Ê—…"

    ',
    'QryTransactionsTotal.TransSum
    'QryTransactionsTotal.TransNet,
    If SystemOptions.SysDataBaseType = SQLServerDataBase Then
        StrSQL = "SELECT QryTransactionsTotal.Transaction_ID, QryTransactionsTotal.Transaction_Serial," & _
        "QryTransactionsTotal.Transaction_Date,dbo.TblCustemers.CusName, QryTransactionsTotal.PaymentType, " & _
        "dbo.TblStore.StoreName,dbo.TblEmployee.Emp_Name ,QryTransactionsTotal.Trans_DiscountType," & _
        "QryTransactionsTotal.Trans_Discount,QryTransactionsTotal.TotalAfterTax"
        StrSQL = StrSQL + " FROM dbo.QryTransactionsTotal() QryTransactionsTotal LEFT OUTER JOIN"
        StrSQL = StrSQL + " dbo.TblStore ON QryTransactionsTotal.StoreID = dbo.TblStore.StoreID LEFT OUTER JOIN"
        StrSQL = StrSQL + " dbo.TblEmployee ON QryTransactionsTotal.Emp_ID = dbo.TblEmployee.Emp_ID LEFT OUTER JOIN"
        StrSQL = StrSQL + " dbo.TblCustemers ON QryTransactionsTotal.CusID = dbo.TblCustemers.CusID"
        StrSQL = StrSQL + " WHERE QryTransactionsTotal.Transaction_Type=2 "
        StrSQL = StrSQL + " Order  By QryTransactionsTotal.Transaction_ID"
    ElseIf SystemOptions.SysDataBaseType = AccessDataBase Then
        StrSQL = "SELECT QryTransactionsTotal.Transaction_ID , QryTransactionsTotal.Transaction_Serial," & _
        "QryTransactionsTotal.Transaction_Date,TblCustemers.CusName, QryTransactionsTotal.PaymentType," & _
        "TblStore.StoreName,TblEmployee.Emp_Name ,QryTransactionsTotal.Trans_DiscountType," & _
        "QryTransactionsTotal.Trans_Discount,QryTransactionsTotal.TotalAfterTax "
        StrSQL = StrSQL + "FROM (TblEmployee RIGHT JOIN (TblCustemers RIGHT JOIN QryTransactionsTotal " & _
        "ON TblCustemers.CusID = QryTransactionsTotal.CusID) ON TblEmployee.Emp_ID = QryTransactionsTotal.Emp_ID) " & _
        "LEFT JOIN TblStore ON QryTransactionsTotal.StoreID = TblStore.StoreID "
        StrSQL = StrSQL + " WHERE QryTransactionsTotal.Transaction_Type=2 "
        StrSQL = StrSQL + " Order  By QryTransactionsTotal.Transaction_ID"
    End If
    Set Rs = New ADODB.Recordset
    Rs.Open StrSQL, Cn, adOpenKeyset, adLockReadOnly, adAsyncExecute + adAsyncFetch
    Set cProgress = New ClsProgress
    BolFrmLoaded = True
    cProgress.ProgressType = Waiting
    cProgress.StartProgress
    Do While Rs.State = adStateExecuting
        DoEvents
    Loop
    If BolFrmLoaded = True Then
        cProgress.StopProgess
        Set cProgress = Nothing
    End If
    Set .DataSource = Rs
    .TextMatrix(0, 0) = "—ﬁ„ «·»—‰«„Ã"
    .TextMatrix(0, 1) = "—ﬁ„ «·›« Ê—…"
    .TextMatrix(0, 2) = " «—ÌŒ «·›« Ê—…"
    .ColDataType(2) = flexDTDate
    .TextMatrix(0, 3) = "«”„ «·⁄„Ì·"
    .TextMatrix(0, 4) = "ÿ—Ìﬁ… «·œ›⁄"
    StrComboList = "#0;‰ﬁœÏ|#1;√Ã·"
    .ColComboList(4) = StrComboList
    .TextMatrix(0, 5) = "«”„ «·„Œ“‰"
    .TextMatrix(0, 6) = "«”„ «·„ÊŸ›"
    
    .TextMatrix(0, 7) = "‰Ê⁄ «·Œ’„"
    .TextMatrix(0, 8) = "ﬁÌ„… «·Œ’„"
    .TextMatrix(0, 9) = "≈Ã„«·Ï «·›« Ê—…"
    .ColKey(9) = "TotalAfterTax"
    'Rs.Close
    'Set Rs = Nothing
End With
Set GrdBack = New ClsBackGroundPic
FrmView.vsfGroup1.vsFlexGrid.WallPaper = GrdBack.Picture
FrmView.vsfGroup1.SetRTL = True
FrmView.vsfGroup1.TotalOnColKey = "TotalAfterTax"
FrmView.vsfGroup1.update
FrmView.Show

End Sub

Private Sub m_MenuViewNotes_Click()
CmdNotes_Click
End Sub

Private Sub TxtFillData_Change()
If TxtFillData.text = "F" Then
    NewGrid.Calculate 1, , , True
End If
End Sub

Private Sub TxtTransSerial_KeyDown(KeyCode As Integer, Shift As Integer)
Dim StrSearch As String
Dim VarBookMark As Variant
Dim Msg As String

If Me.TxtModFlg.text = "R" Then
    If KeyCode = vbKeyReturn Then
        If Trim$(TxtTransSerial.text) <> "" Then
            StrSearch = Trim$(TxtTransSerial.text)
            If Not (Rs.BOF Or Rs.EOF) Then
                If Rs.EditMode = adEditNone Then
                    VarBookMark = Rs.Bookmark
                    Rs.Find "Transaction_Serial='" & StrSearch & "'", , adSearchForward, adBookmarkFirst
                    If Not (Rs.BOF Or Rs.EOF) Then
                        Me.Retrive Rs("Transaction_ID").Value
                    Else
                        Rs.Bookmark = VarBookMark
                        Msg = "Â–Â «·›« Ê—… €Ì— „ÊÃÊœ…...!!!"
                        MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
                    End If
                End If
            End If
        End If
    End If
End If
End Sub


Private Sub TxtTransSerial_KeyPress(KeyAscii As Integer)
KeyAscii = KeyAscii_Num(KeyAscii, Me.TxtTransSerial.text, 1)
End Sub



Private Sub XPBtnMove_Click(Index As Integer)
On Error GoTo ErrTrap
Select Case Index
    Case 0
        If Not (Rs.EOF Or Rs.BOF) Then
            Rs.MovePrevious
            If Rs.BOF Then Rs.MoveFirst
        End If
    Case 1
        If Not (Rs.EOF Or Rs.BOF) Then
            Rs.MoveFirst
        End If
    Case 2
        If Not (Rs.EOF Or Rs.BOF) Then
            Rs.MoveLast
        End If
    Case 3
        If Not (Rs.EOF Or Rs.BOF) Then
            Rs.MoveNext
            If Rs.EOF Then Rs.MoveLast
        End If
End Select
Retrive
Exit Sub
ErrTrap:
End Sub
'
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
On Error GoTo ErrTrap
If KeyCode = vbKeyReturn Then
    If Me.TxtModFlg.text = "R" And Not (Me.ActiveControl Is TxtTransSerial) Then
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
If KeyCode = vbKeyF7 Then
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
If KeyCode = vbKeyDelete Then
    If Me.ActiveControl Is FG Then
        If TxtModFlg.text = "N" Or TxtModFlg.text = "E" Then
            'XPBtnRemove_Click
        End If
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
            FG.SetFocus
        End If
    End If
End If
If Shift = VBRUN.ShiftConstants.vbShiftMask Then
    'vbKeyX
    If KeyCode = vbKeyEscape Then
        Cmd_Click (6)
    End If
End If
Exit Sub
ErrTrap:
End Sub
Private Sub Form_Load()
Dim StrSQL As String
Dim Num As Integer
Dim StrList As String
Dim BGround As New ClsBackGroundPic
Dim ShowTax As Boolean

On Error GoTo ErrTrap
If SystemOptions.UserInterface = EnglishInterface Then
    SetInterface Me
    ChangeLang
End If

Set Cmd(0).ButtonImage = MDIFrmMain.ImgLstTree.ListImages("New").Picture
Set Cmd(1).ButtonImage = MDIFrmMain.ImgLstTree.ListImages("Edit").Picture
Set Cmd(2).ButtonImage = MDIFrmMain.ImgLstTree.ListImages("save").Picture
Set Cmd(3).ButtonImage = MDIFrmMain.ImgLstTree.ListImages("Undo").Picture
Set Cmd(4).ButtonImage = MDIFrmMain.ImgLstTree.ListImages("Del").Picture
Set Cmd(5).ButtonImage = MDIFrmMain.ImgLstTree.ListImages("Search").Picture
Set Cmd(6).ButtonImage = MDIFrmMain.ImgLstTree.ListImages("Exit").Picture
Set Cmd(7).ButtonImage = MDIFrmMain.ImgLstTree.ListImages("Print").Picture
Set CmdHelp.ButtonImage = MDIFrmMain.ImgLstTree.ListImages("Help").Picture
'Set m_menu1.ButtonImage = MDIFrmMain.ImgLstTree.ListImages("Excute").Picture
Set NewGrid.Grid = FG

ShowTax = GetSetting(StrAppRegPath, "SallBill", "HaveTaxOnSalles", False)
Ele(4).Visible = ShowTax
NewGrid.GridTrans = InvoiceTransaction
Set NewGrid.TxtInvID = Me.XPTxtBillID
Set NewGrid.TxtModFlag = TxtModFlg
Set NewGrid.TxtTotal = XPTxtSum
Set NewGrid.CboDiscount_Type = XPCboDiscountType
Set NewGrid.TxtDiscount_Val = XPTxtDiscountVal
Set NewGrid.TxtValueCash = XPTxtValue(0)
Set NewGrid.TxtValueDelay = XPTxtValue(1)
Set NewGrid.TxtValuechque = XPTxtValue(2)
'--------------------------------------
Set NewGrid.TxtTaxValue = Me.XPTxtTaxValue
Set NewGrid.TxtAddTax = Me.TxtTaxAddValue
Set NewGrid.TxtStampTax = Me.TxtTaxStampValue
Set NewGrid.TxtServiceTax = Me.TxtTaxServiceValue
'------------------------------------------------
Set NewGrid.TxtFillData = TxtFillData
Set NewGrid.StoreName = Me.DCboStoreName
Set NewGrid.DtpBillDate = Me.XPDtbBill
Set NewGrid.CmdAddSerialLIst = Me.CmdSearch
'Set NewGrid.CboDiscountType = CboDiscountType
' ⁄»∆… »Ì«‰«  «·√’‰«›
Set NewGrid.DcboItemName = DCboItemsName
Set NewGrid.DCboItemCode = DCboItemsCode
Set NewGrid.CboItemCase = CboItemCase
Set NewGrid.CmdAddData = CmdAdd
Set NewGrid.TxtSerial = TxtSerial
Set NewGrid.TxtQuantity = TxtQuantity
Set NewGrid.TxtPrice = TxtPrice
Set NewGrid.LblInvProfit = Me.LblInvProfit
Set NewGrid.LblItemsCount = Me.LblItemsCount
Set NewGrid.GrdTBar = Me.TBar
Set NewGrid.LblTotalAll = Me.LblTotalAll
Set NewGrid.LblDiscountsTotal = Me.LblDiscountsTotal
Set NewGrid.LblTaxSalesValue = Me.lbl(51)
Set NewGrid.LblTaxAddValue = Me.lbl(52)
Set NewGrid.LblTaxStampValue = Me.lbl(53)
Set NewGrid.LblTaxServiceValue = Me.lbl(54)

NewGrid.FillGrid
FG.WallPaper = BGround.Picture
AddTip
XPTab301.CurrTab = 0
XPDtbBill.Value = Date
If SystemOptions.UserInterface = ArabicInterface Then
    With XPCboDiscountType
        .Clear
        .AddItem "·«ÌÊÃœ Œ’„"
        .AddItem "Œ’„ »ﬁÌ„…"
        .AddItem "Œ’„ »‰”»…"
    End With
    With CboPayMentType
        .Clear
        .AddItem "‰ﬁœ«"
        .AddItem "¬Ã·"
    End With
    With Me.CboSaleType
        .Clear
        .AddItem "ﬁÿ«⁄Ì"
        .AddItem " Ã«—Ï"
    End With
ElseIf SystemOptions.UserInterface = EnglishInterface Then
    With XPCboDiscountType
        .Clear
        .AddItem "No Discount"
        .AddItem "Value Discount"
        .AddItem "Precetage Discount"
    End With
    With CboPayMentType
        .Clear
        .AddItem "Cash"
        .AddItem "Due"
    End With
    With Me.CboSaleType
        .Clear
        .AddItem "Retail"
        .AddItem "WholeSale"
    End With
End If
'--------------------------------
Set Dcombos = New ClsDataCombos
LoadCombosData
'--------------------------------
If SystemOptions.UserInvoiceShowProfit = 0 Then
    Me.Ele(8).Visible = False
Else
    Me.Ele(8).Visible = True
End If
SetDtpickerDate Me.XPDtbBill
'----------------------------
SetDtpickerDate Me.DtpDelayDate
'≈⁄œ«œ Ã—œ «·√ﬁ”«ÿ
ChkInstall.Value = Unchecked
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
Me.XPChkTAX.Value = vbUnchecked
XPChkTAX_Click
Me.ChkTaxAdd.Value = vbUnchecked
ChkTaxAdd_Click
Me.ChkTaxStamp.Value = vbUnchecked
ChkTaxStamp_Click
Me.ChkTaxSerivce.Value = vbUnchecked
ChkTaxSerivce_Click
'---------------------------
Resize_Form Me, TransactionSize
'----------------------------
DB_CreateField "Transactions", "TransactionComment", adVarWChar, adColNullable, 255, , " ”ÃÌ· „·«ÕŸ«  ⁄·Ï «·›« Ê—…", False, True
'----------------------------
StrSQL = "SELECT * FROM Transactions WHERE Transaction_Type=2 Order by Transaction_ID"
Set Rs = New ADODB.Recordset
Rs.Open StrSQL, Cn, adOpenKeyset, adLockOptimistic, adCmdText
If Not (Rs.BOF Or Rs.EOF) Then
    Rs.MoveLast
End If
Retrive
Me.TxtModFlg.text = "R"
Exit Sub
ErrTrap:
End Sub
Private Sub Form_Unload(Cancel As Integer)
On Error GoTo ErrTrap
Dim I As Integer
If Rs.State = adStateOpen Then
    If Not (Rs.EOF Or Rs.BOF) Then
        If Rs.EditMode <> adEditNone Then
            Rs.CancelUpdate
        End If
    End If
    Rs.Close
End If

Set Dcombos = Nothing
For I = LBound(cSearchDcbo) To UBound(cSearchDcbo)
    Set cSearchDcbo(I) = Nothing
Next I
Set Rs = Nothing
Set TTP = Nothing
NewGrid.Class_Terminate
Set NewGrid = Nothing
Set SaleReport = Nothing

Set m_Menu1 = Nothing
Set m_MenuRefesh = Nothing
If Not m_FrmSearch Is Nothing Then
    Unload m_FrmSearch
    Set m_FrmSearch = Nothing
End If
Exit Sub
ErrTrap:
End Sub
Private Sub TxtModFlg_Change()
On Error GoTo ErrTrap
Dim RsTest As ADODB.Recordset
Dim StrSQL As String
Select Case Me.TxtModFlg.text
    Case "R"
        If SystemOptions.UserInterface = ArabicInterface Then
            Me.Caption = "›« Ê—…«·»Ì⁄"
        ElseIf SystemOptions.UserInterface = EnglishInterface Then
            Me.Caption = "Bill Invoice"
        End If
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
        
        XPCboDiscountType.Locked = True
        Me.XPDtbBill.Enabled = False
        Me.DBCboClientName.Locked = True
        Me.DCboStoreName.Locked = True
        
        Me.XPTxtDiscountVal.Locked = True
        XPChkPayType(0).Enabled = False
        XPChkPayType(1).Enabled = False
        XPChkPayType(2).Enabled = False
        XPTxtValue(0).Enabled = False
        XPTxtSerial(0).Enabled = False
        XPTxtValue(1).Enabled = False
        XPTxtSerial(1).Enabled = False
        
        FG.Editable = flexEDNone
        XPChkTAX.Enabled = False
        If Rs.RecordCount < 1 Then
            Me.XPBtnMove(0).Enabled = False
            Me.XPBtnMove(1).Enabled = False
            Me.XPBtnMove(2).Enabled = False
            Me.XPBtnMove(3).Enabled = False
            Me.Cmd(1).Enabled = False
            Me.Cmd(4).Enabled = False
            Me.Cmd(5).Enabled = False
            Me.Cmd(7).Enabled = False
        End If
        
        CboPayMentType.Locked = True
        DtpDelayDate.Enabled = False
        If Not m_Menu1 Is Nothing Then
            m_Menu1.Enabled = False
        End If
        CmdINSTALLMENT.Enabled = False
        CmdCheque.Enabled = False
        '⁄—÷ «·√ﬁ”«ÿ «·„”Ã·… ⁄·Ï «·›« Ê—…
        If XPTxtValue(1).Tag <> "" Then
            StrSQL = "select * From InstallMent where NoteID=" & XPTxtValue(1).Tag
            Set RsTest = New ADODB.Recordset
            RsTest.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText
            If Not (RsTest.EOF Or RsTest.BOF) Then
                CmdINSTALLMENT.Enabled = True
                CmdINSTALLMENT.Caption = "⁄—÷ «·√ﬁ”«ÿ «·„”Ã·…"
            Else
                CmdINSTALLMENT.Enabled = False
                CmdINSTALLMENT.Caption = " ﬁ”Ìÿ «·ﬁÌ„… «·¬Ã·…"
            End If
        End If
        Ele(2).Enabled = False
        DcboEmp.Enabled = False
        XPChkTAX.Enabled = False
        ChkTaxAdd.Enabled = False
        ChkTaxSerivce.Enabled = False
        ChkTaxStamp.Enabled = False
    Case "N"
        If SystemOptions.UserInterface = ArabicInterface Then
            Me.Caption = "›« Ê—…«·»Ì⁄( ÃœÌœ )"
        ElseIf SystemOptions.UserInterface = EnglishInterface Then
            Me.Caption = "Bill Invoice(New)"
        End If
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
        FG.Enabled = True
        FG.Rows = FG.FixedRows
        FG.Rows = 2
        XPCboDiscountType.Locked = False
        Me.XPDtbBill.Enabled = True
        XPDtbBill.Value = Date
        Me.DBCboClientName.Locked = False
        CboPayMentType.Locked = False
        Me.DCboStoreName.Locked = False
        Me.XPTxtDiscountVal.Locked = False
        
        XPChkPayType(0).Enabled = True
        XPChkPayType(1).Enabled = True
        XPChkPayType(2).Enabled = True
        XPChkPayType(0).Value = Unchecked
        XPChkPayType(1).Value = Unchecked
        XPChkPayType(2).Value = Unchecked
        FG.Editable = flexEDKbdMouse
        XPChkTAX.Enabled = True
        XPTxtTaxValue.text = ""
        XPChkTAX.Value = Unchecked
        XPCboDiscountType.ListIndex = 0
        CboPayMentType.ListIndex = 0
'        XPFillData.Enabled = True
        DtpDelayDate.Enabled = True
        m_Menu1.Enabled = True
        DtpDelayDate.Value = Date
       
        CmdINSTALLMENT.Enabled = False
        CmdCheque.Enabled = False
        Ele(2).Enabled = True
        CboItemCase.ListIndex = 0
        
        Me.LblInvProfit.Caption = "0.0"
        Me.LblInvProfit.ForeColor = vbBlack
        
        DcboEmp.Enabled = True
        XPChkTAX.Enabled = True
        ChkTaxAdd.Enabled = True
        ChkTaxSerivce.Enabled = True
        ChkTaxStamp.Enabled = True
        
    Case "E"
        If SystemOptions.UserInterface = ArabicInterface Then
            Me.Caption = "›« Ê—…«·»Ì⁄(   ⁄œÌ· )"
        ElseIf SystemOptions.UserInterface = EnglishInterface Then
            Me.Caption = "Bill Invoice( Edit )"
        End If
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
        FG.Enabled = True
        XPCboDiscountType.Locked = False
        Me.XPDtbBill.Enabled = True
        Me.DBCboClientName.Locked = False
        Me.DCboStoreName.Locked = False
        Me.XPTxtDiscountVal.Locked = False
        CboPayMentType.Locked = False
        XPChkPayType(0).Enabled = True
        XPChkPayType(1).Enabled = True
        XPChkPayType(2).Enabled = True
        DtpDelayDate.Enabled = True
        If XPChkPayType(0).Value = Checked Then
            XPChkPayType_Click (0)
        End If
        If XPChkPayType(1).Value = Checked Then
            XPChkPayType_Click (1)
        End If
        If XPChkPayType(2).Value = Checked Then
            XPChkPayType_Click (2)
        End If
        If CboPayMentType.ListIndex = 0 Then
            CboPayMentType_Change
        End If
        FG.Editable = flexEDKbdMouse
        XPBtnNewClients.Enabled = True
        XPChkTAX.Enabled = True
        If Not m_Menu1 Is Nothing Then
            m_Menu1.Enabled = False
        End If
        If XPChkPayType(1).Value = vbChecked Then
            If XPTxtValue(1).text <> "" Then
                CmdINSTALLMENT.Enabled = True
                CmdINSTALLMENT.Caption = " ﬁ”Ìÿ «·ﬁÌ„… «·¬Ã·…"
            Else
                CmdINSTALLMENT.Enabled = False
            End If
        End If
        If Me.XPChkPayType(2).Value = vbChecked Then
            CmdCheque.Enabled = True
        Else
            CmdCheque.Enabled = False
        End If
        DBCboClientName_Change
        Ele(2).Enabled = True
        
        DcboEmp.Enabled = True
        XPChkTAX.Enabled = True
        ChkTaxAdd.Enabled = True
        ChkTaxSerivce.Enabled = True
        ChkTaxStamp.Enabled = True
        
End Select
Exit Sub
ErrTrap:
End Sub
Public Sub Retrive(Optional LngID As Long = 0)
Dim RsDetails As New ADODB.Recordset
Dim StrSQL As String
Dim RsNotes As New ADODB.Recordset
Dim RsTest  As ADODB.Recordset
Dim RsReplace As ADODB.Recordset
Dim LngPartID As Long
Dim RsPartDetails As ADODB.Recordset
Dim I As Long

On Error GoTo ErrTrap
'---------------------------------------------
'Here We Reset all Setting
Me.CmdNotes.Visible = False
Me.CmdNotes.Tag = ""
Me.CmdRetruns.Visible = False
Me.CmdRetruns.Tag = ""

ChkTaxAdd.Value = vbUnchecked
Me.TxtTaxAddValue.text = ""
ChkTaxStamp.Value = vbUnchecked
Me.TxtTaxStampValue.text = ""
ChkTaxStamp.Value = vbUnchecked
Me.TxtTaxStampValue.text = ""
ChkTaxSerivce.Value = vbUnchecked
Me.TxtTaxServiceValue.text = ""
'---------------------------------------------
If Rs.RecordCount < 1 Then
    XPTxtCurrent.Caption = 0
    XPTxtCount.Caption = 0
    Exit Sub
End If
If Rs.EOF Or Rs.BOF Then
    Exit Sub
End If
If LngID <> 0 Then
    Rs.Find "Transaction_ID=" & LngID, , adSearchForward, adBookmarkFirst
    If Rs.BOF Or Rs.EOF Then
        Exit Sub
    End If
End If
TxtFillData.text = "T"
Screen.MousePointer = vbArrowHourglass
XPTxtBillID.text = IIf(IsNull(Rs("Transaction_ID").Value), "", Val(Rs("Transaction_ID").Value))
TxtTransSerial.text = IIf(IsNull(Rs("Transaction_Serial").Value), "", Rs("Transaction_Serial").Value)
XPDtbBill.Value = IIf(IsNull(Rs("Transaction_Date").Value), "", (Rs("Transaction_Date").Value))
XPCboDiscountType.ListIndex = IIf(IsNull(Rs("Trans_DiscountType").Value), -1, Val(Rs("Trans_DiscountType").Value))
CboPayMentType.ListIndex = IIf(IsNull(Rs("PaymentType").Value), 0, Rs("PaymentType").Value)
XPTxtDiscountVal.text = IIf(IsNull(Rs("Trans_Discount").Value), "", (Rs("Trans_Discount").Value))
Me.DBCboClientName.BoundText = IIf(IsNull(Rs("CusID").Value), "", Rs("CusID").Value)
Me.DCboUserName.BoundText = IIf(IsNull(Rs("UserID").Value), "", Rs("UserID").Value)
FG.Clear flexClearScrollable, flexClearEverything
Me.DCboStoreName.BoundText = IIf(IsNull(Rs("StoreID").Value), "", Rs("StoreID").Value)
Me.DcboEmp.BoundText = IIf(IsNull(Rs("Emp_ID").Value), "", Rs("Emp_ID").Value)
XPTxtTaxValue.text = IIf(IsNull(Rs("TaxValue").Value), "", (Rs("TaxValue").Value))
XPChkTAX.Value = IIf(Rs("TaxFound") = True, Checked, Unchecked)
If IsNull(Rs("SaleType").Value) Then
    Me.CboSaleType.ListIndex = 0
Else
    Me.CboSaleType.ListIndex = IIf(Rs("SaleType").Value = 0, 0, 1)
End If
If Not (IsNull(Rs("CashCustomerName").Value)) Then
    Me.TxtCashCustomerName.text = Rs("CashCustomerName").Value
Else
    Me.TxtCashCustomerName.text = ""
End If
'÷—»Ì… «·Œ’„ Ê«·≈÷«›…
If Not IsNull(Rs("TaxAddValue").Value) Then
    If Rs("TaxAddValue").Value > 0 Then
        ChkTaxAdd.Value = vbChecked
        Me.TxtTaxAddValue.text = Rs("TaxAddValue").Value
    End If
End If
'÷—»Ì… «·œ„€…
If Not IsNull(Rs("TaxStampValue").Value) Then
    If Rs("TaxStampValue").Value > 0 Then
        ChkTaxStamp.Value = vbChecked
        Me.TxtTaxStampValue.text = Rs("TaxStampValue").Value
    End If
End If
'÷—»Ì… «·Œœ„…
If Not IsNull(Rs("TaxServiceValue").Value) Then
    If Rs("TaxServiceValue").Value > 0 Then
        ChkTaxSerivce.Value = vbChecked
        Me.TxtTaxServiceValue.text = Rs("TaxServiceValue").Value
    End If
End If
TxtBillComment.text = IIf(IsNull(Rs("TransactionComment").Value), "", (Rs("TransactionComment").Value))
FG.Rows = 2
FG.Clear flexClearScrollable, flexClearEverything
FG.Refresh
StrSQL = "SELECT TblItems.HaveSerial, * FROM TblItems INNER JOIN Transaction_Details " & _
"ON TblItems.ItemID = Transaction_Details.Item_ID INNER JOIN dbo.TblUnites ON dbo.Transaction_Details.UnitID = dbo.TblUnites.UnitID"
StrSQL = StrSQL + " where Transaction_ID=" & Val(Rs("Transaction_ID").Value)

Set RsDetails = New ADODB.Recordset
RsDetails.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText
XPTxtSum.text = ""
If Not (RsDetails.EOF Or RsDetails.BOF) Then
    FG.Rows = RsDetails.RecordCount + 1
    For I = 1 To RsDetails.RecordCount
        FG.Cell(flexcpPicture, I, FG.ColIndex("Ser")) = ""
        FG.Cell(flexcpData, I, FG.ColIndex("Ser")) = ""
        FG.TextMatrix(I, FG.ColIndex("Code")) = IIf(IsNull(RsDetails("Item_ID")), "", (RsDetails("Item_ID").Value))
        FG.TextMatrix(I, FG.ColIndex("Name")) = IIf(IsNull(RsDetails("Item_ID")), "", Trim$(RsDetails("Item_ID").Value))
        FG.TextMatrix(I, FG.ColIndex("Serial")) = IIf(IsNull(RsDetails("ItemSerial")), "", Trim(RsDetails("ItemSerial").Value))
        If RsDetails("HaveSerial") = True Then
            FG.TextMatrix(I, FG.ColIndex("HaveSerial")) = True
            '«·»ÕÀ ⁄‰ ⁄„·Ì«  «·«” »œ«· «·Œ«’… »«·›« Ê—…
            If (RsDetails("Item_ID")) <> "" And RsDetails("ItemSerial") <> "" Then
                StrSQL = "select * From ReplacedItems where ReturnID=" & XPTxtBillID.text
                StrSQL = StrSQL + " and ItemID=" & RsDetails("Item_ID")
                StrSQL = StrSQL + " and ItemSerial='" & RsDetails("ItemSerial") & "'"
                Set RsReplace = New ADODB.Recordset
                RsReplace.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText
                If Not (RsReplace.EOF Or RsReplace.BOF) Then
                    FG.Cell(flexcpPicture, I, FG.ColIndex("Ser")) = MDIFrmMain.ImgLstTree.ListImages("Request").Picture
                    FG.Cell(flexcpData, I, FG.ColIndex("Ser")) = "X"
                End If
            End If
        End If
        FG.TextMatrix(I, FG.ColIndex("ItemType")) = _
             IIf(IsNull(RsDetails("ItemType").Value), "", (RsDetails("ItemType").Value))
        If RsDetails("ItemType").Value = 1 Then
            FG.Cell(flexcpPicture, I, FG.ColIndex("Ser")) = _
                MDIFrmMain.ImgLstTree.ListImages("Maintenance").Picture
            
        End If
        FG.TextMatrix(I, FG.ColIndex("Count")) = IIf(IsNull(RsDetails("Quantity")), "", (RsDetails("Quantity").Value))
        FG.TextMatrix(I, FG.ColIndex("Price")) = IIf(IsNull(RsDetails("Price")), "", (RsDetails("Price").Value))
        
        If SystemOptions.SysDataBaseType = SQLServerDataBase Then
            FG.TextMatrix(I, FG.ColIndex("ItemCase")) = IIf(IsNull(RsDetails("ItemCase")), "", (RsDetails("ItemCase").Value))
        Else
            FG.TextMatrix(I, FG.ColIndex("ItemCase")) = IIf(IsNull(RsDetails("Transaction_Details.ItemCase")), "", (RsDetails("Transaction_Details.ItemCase").Value))
        End If
        
        FG.TextMatrix(I, FG.ColIndex("DiscountType")) = IIf(IsNull(RsDetails("ItemDiscountType")), "", (RsDetails("ItemDiscountType").Value))
        FG.TextMatrix(I, FG.ColIndex("DiscountVal")) = IIf(IsNull(RsDetails("ItemDiscount")), "", (RsDetails("ItemDiscount").Value))
        FG.TextMatrix(I, FG.ColIndex("guaranteeTime")) = IIf(IsNull(RsDetails("guaranteeTime")), "", (RsDetails("guaranteeTime").Value))
        
        FG.TextMatrix(I, FG.ColIndex("ItemCostPrice")) = IIf(IsNull(RsDetails("CostPrice")), "", (RsDetails("CostPrice").Value))
        FG.TextMatrix(I, FG.ColIndex("PofTransID")) = IIf(IsNull(RsDetails("CostTransID")), "", (RsDetails("CostTransID").Value))
        FG.TextMatrix(I, FG.ColIndex("ItemProfit")) = IIf(IsNull(RsDetails("ItemProfit")), "", (RsDetails("ItemProfit").Value))
        FG.TextMatrix(I, FG.ColIndex("Remarks")) = IIf(IsNull(RsDetails("Remarks")), "", (RsDetails("Remarks").Value))
        
        FG.TextMatrix(I, FG.ColIndex("ColorID")) = IIf(IsNull(RsDetails("ColorID")), 1, (RsDetails("ColorID").Value))
        FG.TextMatrix(I, FG.ColIndex("ItemSize")) = IIf(IsNull(RsDetails("ItemSize")), 0, (RsDetails("ItemSize").Value))
        
        If Val(FG.TextMatrix(I, FG.ColIndex("ItemProfit"))) = 0 Then
            Me.FG.Cell(flexcpBackColor, I, 1, I, FG.Cols - 1) = vbYellow
        ElseIf Val(FG.TextMatrix(I, FG.ColIndex("ItemProfit"))) < 0 Then
             Me.FG.Cell(flexcpBackColor, I, 1, I, FG.Cols - 1) = vbRed
        Else
             Me.FG.Cell(flexcpBackColor, I, 1, I, FG.Cols - 1) = 0
        End If
        FG.Cell(flexcpData, I, FG.ColIndex("UnitID")) = IIf(IsNull(RsDetails("UnitID")), "", (RsDetails("UnitID").Value))
        FG.TextMatrix(I, FG.ColIndex("UnitID")) = IIf(IsNull(RsDetails("UnitName")), "", (RsDetails("UnitName").Value))

        RsDetails.MoveNext
        
        If FG.Rows > 10 Then
            If I = 8 Then FG.Refresh
        End If
    Next I
    '----------------------------
    Me.LblInvProfit.Caption = FG.Aggregate(flexSTSum, FG.FixedRows, FG.ColIndex("ItemProfit"), FG.Rows - 1, FG.ColIndex("ItemProfit"))
    If Val(Me.LblInvProfit.Caption) > 0 Then
        Me.LblInvProfit.ForeColor = &H4000&
    ElseIf Val(Me.LblInvProfit.Caption) = 0 Then
        Me.LblInvProfit.ForeColor = vbBlack
    ElseIf Val(Me.LblInvProfit.Caption) < 0 Then
        Me.LblInvProfit.ForeColor = vbRed
    End If
    '---------------------------
    FG.AutoSize 0, FG.Cols - 1, False
End If
XPChkPayType(0).Value = Unchecked
XPChkPayType(1).Value = Unchecked
XPChkPayType(2).Value = Unchecked
XPTxtValue(0).text = ""
XPTxtValue(1).text = ""
XPTxtSerial(0).text = ""
XPTxtSerial(1).text = ""
XPTxtValue(1).Tag = ""
DtpDelayDate.Value = Date
'----------------------------------------------------------------------------------------
StrSQL = "Select * From Notes Where Transaction_ID=" & Val(Rs("Transaction_ID").Value)
RsNotes.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText
If Not (RsNotes.EOF Or RsNotes.BOF) Then
    For I = 1 To RsNotes.RecordCount
        If RsNotes("NoteType").Value = 0 Then
            XPChkPayType(0).Value = Checked
            XPChkPayType_Click (0)
            XPTxtValue(0).text = IIf(IsNull(RsNotes("Note_Value").Value), "", (RsNotes("Note_Value").Value))
            XPTxtSerial(0).text = IIf(IsNull(RsNotes("NoteSerial").Value), "", Trim$(RsNotes("NoteSerial").Value))
            Me.DcboBox.BoundText = IIf(IsNull(RsNotes("BoxID").Value), "", RsNotes("BoxID").Value)
        End If
        If RsNotes("NoteType").Value = 1 Then
            XPChkPayType(1).Value = Checked
            XPChkPayType_Click (1)
            XPTxtValue(1).text = IIf(IsNull(RsNotes("Note_Value").Value), "", (RsNotes("Note_Value").Value))
            XPTxtValue(1).Tag = IIf(IsNull(RsNotes("NoteID").Value), "", (RsNotes("NoteID").Value))
            XPTxtSerial(1).text = IIf(IsNull(RsNotes("NoteSerial").Value), "", Trim$(RsNotes("NoteSerial").Value))
            DtpDelayDate.Value = IIf(IsNull(RsNotes("DueDate").Value), "", (RsNotes("DueDate").Value))
        End If
        If RsNotes("NoteType").Value = 2 Then
            XPChkPayType(2).Value = Checked
            XPChkPayType_Click (2)
        End If
        RsNotes.MoveNext
    Next I
End If
Set RsNotes = New ADODB.Recordset
StrSQL = "SELECT Notes.NoteID, Notes.NoteDate, Notes.NoteType, Notes.NoteSerial," & _
"Notes.Note_Value, Notes.BankID,BanksData.BankName , Notes.ChqueNum, Notes.DueDate "
StrSQL = StrSQL + " FROM Notes INNER JOIN BanksData ON Notes.BankID = BanksData.BankID "
StrSQL = StrSQL + " Where NoteType=2 AND NOTES.Transaction_ID=" & Val(Rs("Transaction_ID").Value)
StrSQL = StrSQL + " Order BY Notes.NoteID"
RsNotes.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText
With Me.FgCheques
    .Rows = .FixedRows
    If Not (RsNotes.BOF Or RsNotes.EOF) Then
        .Rows = .FixedRows + RsNotes.RecordCount
        For I = .FixedRows To .Rows - 1
            .TextMatrix(I, .ColIndex("CheckValue")) = IIf(IsNull(RsNotes("Note_Value").Value), "", RsNotes("Note_Value").Value)
            .TextMatrix(I, .ColIndex("CheckNumber")) = IIf(IsNull(RsNotes("ChqueNum").Value), "", RsNotes("ChqueNum").Value)
            .TextMatrix(I, .ColIndex("BankID")) = IIf(IsNull(RsNotes("BankID").Value), "", RsNotes("BankID").Value)
            .TextMatrix(I, .ColIndex("BankName")) = IIf(IsNull(RsNotes("BankName").Value), "", RsNotes("BankName").Value)
            If Not IsNull(RsNotes("DueDate").Value) Then
                .TextMatrix(I, .ColIndex("DueDate")) = DisplayDate(RsNotes("DueDate").Value)
            Else
                .TextMatrix(I, .ColIndex("DueDate")) = ""
            End If
            RsNotes.MoveNext
        Next I
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
        LngPartID = RsTest("PartID").Value
        Me.LblPrecenType.Tag = RsTest("InterestType").Value
        If RsTest("InterestType").Value = 0 Then
            LblPrecenType.Caption = "‰”»… „∆ÊÌ…"
        ElseIf RsTest("InterestType").Value = 1 Then
            LblPrecenType.Caption = "ﬁÌ„… À«» …"
        ElseIf RsTest("InterestType").Value = 2 Then
            LblPrecenType.Caption = "·«ÌÊÃœ"
        End If
        Me.LblPrecenValue.Caption = RsTest("InterestVal").Value
        Me.LblInstallTotal.Caption = RsTest("Total").Value
        Me.LblInstallCount.Caption = RsTest("InstallCount").Value
        Me.LblFirstInstallDate.Caption = DisplayDate(RsTest("FirstInstallDate").Value)
        Me.LblInstallmentType.Tag = RsTest("InstallmentType").Value
        If RsTest("InstallmentType").Value = 0 Then
            LblInstallmentType.Caption = "ÌÊ„"
        ElseIf RsTest("InstallmentType").Value = 1 Then
            LblInstallmentType.Caption = "‘Â—"
        ElseIf RsTest("InstallmentType").Value = 2 Then
            LblInstallmentType.Caption = "”‰…"
        End If
        Me.LblInstallSeprator.Caption = RsTest("InstallSeprator").Value
        Me.LblStartValue.Caption = IIf(IsNull(RsTest("StartValue").Value), "", RsTest("StartValue").Value)
        Set RsPartDetails = New ADODB.Recordset
        StrSQL = "Select * From InstallMentDetails Where PartID=" & LngPartID
        RsPartDetails.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText
        If Not (RsPartDetails.BOF Or RsPartDetails.EOF) Then
            RsPartDetails.MoveFirst
            With Me.FgInstallments
                .Rows = .FixedRows + RsPartDetails.RecordCount
                For I = .FixedRows To .Rows - 1
                    .TextMatrix(I, .ColIndex("QestID")) = IIf(IsNull(RsPartDetails("QestID").Value), "", RsPartDetails("QestID").Value)
                    .TextMatrix(I, .ColIndex("Value")) = IIf(IsNull(RsPartDetails("Value").Value), "", RsPartDetails("Value").Value)
                    If Not IsNull(RsPartDetails("DueDate").Value) Then
                        .TextMatrix(I, .ColIndex("Due_Date")) = DisplayDate(RsPartDetails("DueDate").Value)
                    Else
                         .TextMatrix(I, .ColIndex("Due_Date")) = ""
                    End If
                    RsPartDetails.MoveNext
                Next I
            End With
        End If
    Else
        CmdINSTALLMENT.Enabled = False
        CmdINSTALLMENT.Caption = " ﬁ”Ìÿ «·ﬁÌ„… «·¬Ã·…"
    End If
End If
TxtFillData.text = "F"
'-----------------------------------------------------------------------------------------------
Dim SngRelatedNotesValues As Single
Me.CmdNotes.Visible = ShowRelatedNotes(Val(Me.XPTxtBillID.text), 0, SngRelatedNotesValues)
Me.CmdNotes.Tag = SngRelatedNotesValues

SngRelatedNotesValues = 0
Me.CmdRetruns.Visible = ShowRelatedTransactions(Val(Me.XPTxtBillID.text), 0, SngRelatedNotesValues)
Me.CmdRetruns.Tag = SngRelatedNotesValues

'-----------------------------------------------------------------------------------------------
Screen.MousePointer = vbDefault
XPTxtCurrent.Caption = Rs.AbsolutePosition
XPTxtCount.Caption = Rs.RecordCount
Exit Sub
ErrTrap:
Resume
Screen.MousePointer = vbDefault
End Sub
Private Sub Undo()
Dim Msg As String

On Error GoTo ErrTrap

Select Case TxtModFlg.text
    Case "N"
        Msg = "”Ê› Ì „ «· —«Ã⁄ ›Ï  ”ÃÌ· Â–Â «·›« Ê—… .."
        Msg = Msg & Chr(13) & "›Â· «‰  „ «ﬂœ „‰ «·√” „—«— ..!!"
        If MsgBox(Msg, vbQuestion + vbYesNo + vbMsgBoxRight + vbMsgBoxRtlReading + vbDefaultButton2, App.Title) = vbYes Then
            clear_all Me
            Me.TxtModFlg.text = "R"
            XPBtnMove_Click (1)
        End If
    Case "E"
        Msg = "”Ê› Ì „ «· —«Ã⁄ ›Ï  ⁄œÌ· Â–Â «·›« Ê—… .."
        Msg = Msg & Chr(13) & "›Â· «‰  „ «ﬂœ „‰ «·√” „—«— ..!!"
        If MsgBox(Msg, vbQuestion + vbYesNo + vbMsgBoxRight + vbMsgBoxRtlReading + vbDefaultButton2, App.Title) = vbYes Then
            Rs.Find "Transaction_ID='" & Val(XPTxtBillID.text) & "'", , adSearchForward, adBookmarkFirst
            If Rs.EOF Or Rs.BOF Then
                Me.TxtModFlg.text = "R"
                Exit Sub
            End If
            If Not Rs.EOF Or Rs.BOF Then
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
Dim RsTest As ADODB.Recordset
Dim StrSQL As String
Dim IntRes As Integer
Dim BegainTrans As Boolean
On Error GoTo ErrTrap
If XPTxtBillID.text = "" Then
    clear_all Me
    Msg = "Â–Â «·⁄„·Ì… €Ì— „ «Õ… ÕÌÀ √‰Â ·«ÌÊÃœ √Ì ”Ã·« "
    MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + _
    vbMsgBoxRtlReading, App.Title
    TxtModFlg_Change
    Exit Sub
End If
If AvailableDeal = False Then
    Exit Sub
End If
'«·√ﬁ”«ÿ «·„”œœ… ⁄·Ï «·›« Ê—…
If XPTxtValue(1).Tag <> "" Then
    StrSQL = "select * From ReceiptQestForBill Where NoteID=" & XPTxtValue(1).Tag
    Set RsTest = New ADODB.Recordset
    RsTest.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText
    If Not (RsTest.EOF Or RsTest.BOF) Then
        Msg = "·ﬁœ  „  Õ’Ì· »⁄÷ «·√ﬁ”«ÿ «·„”Ã·… ⁄·Ï Â–Â «·›« Ê—…" & Chr(13)
        Msg = Msg + "Ê·« Ì„ﬂ‰ Õ–› »Ì«‰« Â«" & Chr(13)
        Msg = Msg + "≈–« ﬂ‰   —€» ›Ì Õ–› »Ì«‰«  Â–Â «·›« Ê—…" & Chr(13)
        Msg = Msg + "ÌÃ» Õ–› ⁄„·Ì«  «· Õ’Ì· «·Œ«’… »Â«"
        MsgBox Msg, vbOKOnly + vbQuestion + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
        Exit Sub
    End If
End If
'⁄„·Ì«  «·’Ì«‰… «·„— »ÿ… »«·›« Ê—…
StrSQL = "select * From MaintenanceJuncTransaction Where Transaction_ID=" & _
Trim(XPTxtBillID.text)
Set RsTest = New ADODB.Recordset
RsTest.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText
If Not (RsTest.EOF Or RsTest.BOF) Then
    Msg = "·ﬁœ  „ ≈Ã—«¡ »⁄÷ ⁄„·Ì«  «·’Ì«‰… ⁄·Ï Â–Â «·›« Ê—… Ê·« Ì„ﬂ‰ Õ–›Â«"
    Msg = Msg + "≈–« ﬂ‰   —€» ›Ì Õ–› »Ì«‰«  Â–Â «·›« Ê—…" & Chr(13)
    Msg = Msg + "ÌÃ» Õ–› ⁄„·Ì«  «·’Ì«‰… «·Œ«’… »Â«"
    MsgBox Msg, vbOKOnly + vbQuestion + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
    Exit Sub
End If
If Me.CboPayMentType.ListIndex = 0 Then
    '›« Ê—… ‰ﬁœÌ…
    If CheckBoxAccount(Me.DcboBox.BoundText, Val(Me.XPTxtValue(0).text), XPDtbBill.Value, False) = False Then
        Msg = "·‰ Ì„ﬂ‰ «·”„«Õ »Õ–› Â–« «·⁄„·Ì…..!!!"
        Msg = Msg & Chr(13) & "ÕÌÀ «‰Â« ”Ê› Ì‰ Ã ⁄‰Â« Œÿ« ›Ï Õ”«»«  «·Œ“‰…"
        MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
        Exit Sub
    End If
End If
Msg = "”Ì „ Õ–› »Ì«‰«  «·⁄„·Ì… —ﬁ„ " & Chr(13)
Msg = Msg + (TxtTransSerial.text) & Chr(13)
Msg = Msg + " Â·  —€» ›Ì Õ–› Â–Â «·»Ì«‰« ø"
IntRes = MsgBox(Msg, vbYesNo + vbQuestion + vbMsgBoxRight + _
    vbMsgBoxRtlReading, App.Title)
If IntRes = vbYes Then
    If Not Rs.RecordCount < 1 Then
        Cn.BeginTrans
        BegainTrans = True
        StrSQL = "Delete From DOUBLE_ENTREY_VOUCHERS  " & _
        "Where DOUBLE_ENTREY_VOUCHERS.Transaction_ID=" & Rs("Transaction_ID").Value
        Cn.Execute StrSQL, , adExecuteNoRecords
        Rs.Delete
        Cn.CommitTrans
        BegainTrans = False
        Msg = " „  ⁄„·Ì… «·Õ–› "
        MsgBox Msg, vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
        Rs.MoveFirst
        If Rs.RecordCount < 1 Then
            clear_all Me
            TxtModFlg_Change
            XPTxtCurrent.Caption = 0
            XPTxtCount.Caption = 0
        Else
            Retrive
        End If
    End If
End If
TxtModFlg_Change
Exit Sub
ErrTrap:

Msg = "·«Ì„ﬂ‰ Õ–› Â–« «·”Ã· · ﬂ«„· «·»Ì«‰«  " & _
Chr(13) & " ÊÃœ »Ì«‰«  „— »ÿ… »Â–« «·”Ã· "
Msg = Msg & Chr(13) & Err.Number
Msg = Msg & Chr(13) & Err.Description
Msg = Msg & Chr(13) & Err.Source
MsgBox Msg, vbMsgBoxRight + vbMsgBoxRtlReading + _
        vbExclamation, App.Title
    

If BegainTrans = True Then
    Rs.CancelUpdate
    Cn.RollbackTrans
    BegainTrans = False
End If
End Sub
Private Sub AddTip()
Dim Wrap As String
Dim BolRtl As Boolean

On Error GoTo ErrTrap
Wrap = Chr(13) + Chr(10)
Set TTP = New clstooltip
If SystemOptions.UserInterface = ArabicInterface Then
    BolRtl = True
    With TTP
       .Create Me.hwnd, "»Ì«‰«  ›« Ê—… «·»Ì⁄", 1, 15204351, -2147483630, BolRtl
       .MaxWidth = 4000
       .VisibleTime = 9000
       .DelayTime = 600
       .AddControl Cmd(0), _
        "ÃœÌœ ..." & Wrap & _
        "·«÷«›… »Ì«‰«  ⁄„·Ì… »Ì⁄ ÃœÌœ…" & Wrap & _
        "  ≈÷€ÿ Â‰«" & Wrap & _
        "„›« ÌÕ «·«Œ ’«— F12 OR Enter", BolRtl
    End With
    With TTP
       .Create Me.hwnd, "»Ì«‰«  ›« Ê—… «·»Ì⁄", 1, 15204351, -2147483630
       .MaxWidth = 4000
       .VisibleTime = 9000
       .DelayTime = 600
       .AddControl Cmd(7), _
        "ÿ»«⁄… ..." & Wrap & _
        "·⁄—÷ «·»Ì«‰«  «·Õ«·Ì… ›Ì  ﬁ—Ì— " & Wrap & _
        " Ì„ﬂ‰ ÿ»«⁄ Â ⁄‰ ÿ—Ìﬁ «·ÿ«»⁄…" & Wrap & _
        "„›« ÌÕ «·«Œ ’«— F6", BolRtl
    End With
    With TTP
       .Create Me.hwnd, "»Ì«‰«  ›« Ê—… «·»Ì⁄", 1, 15204351, -2147483630
       .MaxWidth = 4000
       .VisibleTime = 9000
       .DelayTime = 600
       .AddControl Cmd(1), _
        " ⁄œÌ· ..." & Wrap & _
        "· ⁄œÌ· »Ì«‰«  ⁄„·Ì… «·»Ì⁄" & Wrap & _
        "  ≈÷€ÿ Â‰«" & Wrap & _
        "„›« ÌÕ «·«Œ ’«— F11", BolRtl
    End With
    With TTP
       .Create Me.hwnd, "»Ì«‰«  ›« Ê—… «·»Ì⁄", 1, 15204351, -2147483630
       .MaxWidth = 4000
       .VisibleTime = 9000
       .DelayTime = 600
       .AddControl Cmd(2), _
        "Õ›Ÿ ..." & Wrap & _
        "·Õ›Ÿ »Ì«‰«  ⁄„·Ì… «·»Ì⁄ «·ÃœÌœ…" & Wrap & _
         "·Õ›Ÿ «· ⁄œÌ·« " & Wrap & _
        "  ≈÷€ÿ Â‰«" & Wrap & _
        "„›« ÌÕ «·«Œ ’«— F10", BolRtl
    End With
    With TTP
       .Create Me.hwnd, "»Ì«‰«  ›« Ê—… «·»Ì⁄", 1, 15204351, -2147483630
       .MaxWidth = 4000
       .VisibleTime = 9000
       .DelayTime = 600
       .AddControl Cmd(3), _
        " —«Ã⁄ ..." & Wrap & _
        "·· —«Ã⁄ ⁄‰ ⁄„·Ì… «·»Ì⁄" & Wrap & _
         "··· —«Ã⁄ ⁄‰ ⁄„·Ì… «· ⁄œÌ·" & Wrap & _
        "  ≈÷€ÿ Â‰«" & Wrap & _
        "„›« ÌÕ «·«Œ ’«— F9", BolRtl
    End With
     With TTP
       .Create Me.hwnd, "»Ì«‰«  ›« Ê—… «·»Ì⁄", 1, 15204351, -2147483630
       .MaxWidth = 4000
       .VisibleTime = 9000
       .DelayTime = 600
       .AddControl Cmd(4), _
        "Õ–› ..." & Wrap & _
        "·Õ–› »Ì«‰«  ⁄„·Ì… »Ì⁄" & Wrap & _
        "  ≈÷€ÿ Â‰«" & Wrap & _
        "„›« ÌÕ «·«Œ ’«— F8", BolRtl
    End With
    With TTP
       .Create Me.hwnd, "»Ì«‰«  ›« Ê—… «·»Ì⁄", 1, 15204351, -2147483630
       .MaxWidth = 4000
       .VisibleTime = 9000
       .DelayTime = 600
       .AddControl Cmd(5), _
        "»ÕÀ ..." & Wrap & _
        "···»ÕÀ ⁄‰ ⁄„·Ì… »Ì⁄" & Wrap & _
        "Ì‰ÿ»ﬁ ⁄·ÌÂ« ‘—Êÿ „⁄Ì‰…" & Wrap & _
        "  ≈÷€ÿ Â‰«" & Wrap & _
        "„›« ÌÕ «·«Œ ’«— F7", BolRtl
    End With
    With TTP
       .Create Me.hwnd, "»Ì«‰«  ›« Ê—… «·»Ì⁄", 1, 15204351, -2147483630
       .MaxWidth = 4000
       .VisibleTime = 9000
       .DelayTime = 600
       .AddControl Cmd(6), _
        "Œ—ÊÃ ..." & Wrap & _
        "·«€·«ﬁ Â–Â «·‰«›–…" & Wrap & _
        "  ≈÷€ÿ Â‰«" & Wrap & _
        "„›« ÌÕ «·«Œ ’«— Ctrl + X", BolRtl
    End With
    
    
    With TTP
       .Create Me.hwnd, "»Ì«‰«  ›« Ê—… «·»Ì⁄", 1, 15204351, -2147483630
       .MaxWidth = 4000
       .VisibleTime = 9000
       .DelayTime = 600
       .AddControl XPBtnNewClients, _
        "≈÷«›… ⁄„Ì· ÃœÌœ ..." & Wrap & _
        "· ”ÃÌ· »Ì«‰«  ⁄„Ì· ÃœÌœ" & Wrap & _
        " «÷€ÿ Â‰«" & Wrap & _
        "„›« ÌÕ «·«Œ ’«— F5", BolRtl
    End With
    
    With TTP
       .Create Me.hwnd, "»Ì«‰«  ›« Ê—… «·»Ì⁄", 1, 15204351, -2147483630
       .MaxWidth = 4000
       .VisibleTime = 9000
       .DelayTime = 600
       .AddControl XPBtnMove(1), _
        "«·√Ê· ..." & Wrap & _
        "··«‰ ﬁ«· «·Ï √Ê· ”Ã·" & Wrap & _
        " ›ﬁÿ ≈÷€ÿ Â‰«", BolRtl
    End With
    With TTP
       .Create Me.hwnd, "»Ì«‰«  ›« Ê—… «·»Ì⁄", 1, 15204351, -2147483630
       .MaxWidth = 4000
       .VisibleTime = 9000
       .DelayTime = 600
       .AddControl XPBtnMove(0), _
        "«·”«»ﬁ ..." & Wrap & _
        "··«‰ ﬁ«· «·Ï «·”Ã· «·”«»ﬁ" & Wrap & _
        " ›ﬁÿ ≈÷€ÿ Â‰«", BolRtl
    End With
    With TTP
       .Create Me.hwnd, "»Ì«‰«  ›« Ê—… «·»Ì⁄", 1, 15204351, -2147483630
       .MaxWidth = 4000
       .VisibleTime = 9000
       .DelayTime = 600
       .AddControl XPBtnMove(3), _
        "«· «·Ì ..." & Wrap & _
        "··«‰ ﬁ«· «·Ï «·”Ã· «· «·Ì" & Wrap & _
        " ›ﬁÿ ≈÷€ÿ Â‰«", BolRtl
    End With
    With TTP
       .Create Me.hwnd, "»Ì«‰«  ›« Ê—… «·»Ì⁄", 1, 15204351, -2147483630
       .MaxWidth = 4000
       .VisibleTime = 9000
       .DelayTime = 600
       .AddControl XPBtnMove(2), _
        "«·√ŒÌ— ..." & Wrap & _
        "··«‰ ﬁ«· «·Ï ¬Œ— ”Ã·" & Wrap & _
        " ›ﬁÿ ≈÷€ÿ Â‰«", BolRtl
    End With
    With TTP
       .Create Me.hwnd, "»Ì«‰«  ›« Ê—… «·»Ì⁄", 1, 15204351, -2147483630
       .MaxWidth = 4000
       .VisibleTime = 9000
       .DelayTime = 600
       .AddControl CmdHelp, _
        "„”«⁄œ… ..." & Wrap & _
        "·· ⁄—› ⁄·Ï ÊŸÌ›… Â–Â «·‰«›–…" & Wrap & _
        "ÊﬂÌ›Ì… «· ⁄«„· „⁄Â«" & Wrap & _
        "≈÷€ÿ Â‰«" & Wrap, BolRtl
    End With
ElseIf SystemOptions.UserInterface = EnglishInterface Then
    BolRtl = False
    With TTP
       .Create Me.hwnd, "Bill Invoice", 1, 15204351, -2147483630, BolRtl
       .MaxWidth = 4000
       .VisibleTime = 9000
       .DelayTime = 600
       .AddControl Cmd(0), _
        "New..." & Wrap & _
        "Click here to add new Bill Invoice" & Wrap & _
        "" & Wrap & _
        "Shortcut (Enter Or F12)", BolRtl
    End With
    With TTP
       .Create Me.hwnd, "Bill Invoice", 1, 15204351, -2147483630, BolRtl
       .MaxWidth = 4000
       .VisibleTime = 9000
       .DelayTime = 600
       .AddControl Cmd(7), _
        "Print..." & Wrap & _
        "Print this Bill Invoice" & Wrap & _
        "" & Wrap & _
        "Shortcut (F6)", BolRtl
    End With
    With TTP
       .Create Me.hwnd, "Bill Invoice", 1, 15204351, -2147483630, BolRtl
       .MaxWidth = 4000
       .VisibleTime = 9000
       .DelayTime = 600
       .AddControl Cmd(1), _
        "Edit..." & Wrap & _
        "Edit this Bill Invoice Record" & Wrap & _
        "  " & Wrap & _
        "Shortcut (F11)", BolRtl
    End With
    With TTP
       .Create Me.hwnd, "Bill Invoice", 1, 15204351, -2147483630, BolRtl
       .MaxWidth = 4000
       .VisibleTime = 9000
       .DelayTime = 600
       .AddControl Cmd(2), _
        "Save..." & Wrap & _
        "Save the New Bill Invoice Or Save the edit" & Wrap & _
         "in the current Bill Invoice" & Wrap & _
        "" & Wrap & _
        "Shortcut (F10)", BolRtl
    End With
    With TTP
       .Create Me.hwnd, "Bill Invoice", 1, 15204351, -2147483630, BolRtl
       .MaxWidth = 4000
       .VisibleTime = 9000
       .DelayTime = 600
       .AddControl Cmd(3), _
        "Undo..." & Wrap & _
        "Undo in the New Bill Invoice" & Wrap & _
        "Or Undo in the Editing" & Wrap & _
        "" & Wrap & _
        "Shortcut (F9)", BolRtl
    End With
     With TTP
       .Create Me.hwnd, "Bill Invoice", 1, 15204351, -2147483630, BolRtl
       .MaxWidth = 4000
       .VisibleTime = 9000
       .DelayTime = 600
       .AddControl Cmd(4), _
        "Delete..." & Wrap & _
        "Delete this current Bill Invoice" & Wrap & _
        "" & Wrap & _
        "Shortcut (F8)", BolRtl
    End With
    With TTP
       .Create Me.hwnd, "Bill Invoice", 1, 15204351, -2147483630, BolRtl
       .MaxWidth = 4000
       .VisibleTime = 9000
       .DelayTime = 600
       .AddControl Cmd(5), _
        "Search..." & Wrap & _
        "Click here to display the search" & Wrap & _
        "Screen" & Wrap & _
        "Shortcut (F7)", BolRtl
    End With
    With TTP
       .Create Me.hwnd, "Bill Invoice", 1, 15204351, -2147483630, BolRtl
       .MaxWidth = 4000
       .VisibleTime = 9000
       .DelayTime = 600
       .AddControl Cmd(6), _
        "Exit..." & Wrap & _
        "Close this Window", BolRtl
    End With
    
    
    With TTP
       .Create Me.hwnd, "Bill Invoice", 1, 15204351, -2147483630, BolRtl
       .MaxWidth = 4000
       .VisibleTime = 9000
       .DelayTime = 600
       .AddControl XPBtnNewClients, _
        "Add New Customer...." & Wrap & _
        "To add New Customer Click here..." & Wrap & _
        "Shortcut (F5)", BolRtl
    End With
    
    With TTP
       .Create Me.hwnd, "Bill Invoice", 1, 15204351, -2147483630, BolRtl
       .MaxWidth = 4000
       .VisibleTime = 9000
       .DelayTime = 600
       .AddControl XPBtnMove(1), _
        "First..." & Wrap & _
        "Move to first Record" & Wrap & _
        "", BolRtl
    End With
    With TTP
       .Create Me.hwnd, "Bill Invoice", 1, 15204351, -2147483630, BolRtl
       .MaxWidth = 4000
       .VisibleTime = 9000
       .DelayTime = 600
       .AddControl XPBtnMove(0), _
        "Previous..." & Wrap & _
        "Move to Previous Record" & Wrap & _
        " , BolRTL"
    End With
    With TTP
       .Create Me.hwnd, "Bill Invoice", 1, 15204351, -2147483630, BolRtl
       .MaxWidth = 4000
       .VisibleTime = 9000
       .DelayTime = 600
       .AddControl XPBtnMove(3), _
        "Next..." & Wrap & _
        "Move to Next Record" & Wrap & _
        "", BolRtl
    End With
    With TTP
       .Create Me.hwnd, "Bill Invoice", 1, 15204351, -2147483630, BolRtl
       .MaxWidth = 4000
       .VisibleTime = 9000
       .DelayTime = 600
       .AddControl XPBtnMove(2), _
        "Last..." & Wrap & _
        "Move to Last Record" & Wrap & _
        "", BolRtl
    End With
    With TTP
       .Create Me.hwnd, "Bill Invoice", 1, 15204351, -2147483630, BolRtl
       .MaxWidth = 4000
       .VisibleTime = 9000
       .DelayTime = 600
       .AddControl CmdHelp, _
        "Help..." & Wrap & _
        "to View Help Files" & Wrap & _
        "click Here" & Wrap & _
        "Shortcut(F1)" & Wrap, BolRtl
    End With
End If
Exit Sub
ErrTrap:
End Sub
Private Sub SaveData()
Dim Msg As String
Dim RowNum As Integer
Dim RSTransDetails As ADODB.Recordset
Dim RsNotes As ADODB.Recordset
Dim RsTemp      As New ADODB.Recordset
Dim RsTest      As New ADODB.Recordset
Dim RsRepeat    As ADODB.Recordset
Dim RsDetalis   As ADODB.Recordset
Dim StrSQL      As String
Dim StrSqlDel   As String
Dim Note_ID As Long
Dim TransBegine As Boolean
Dim BolTemp As Boolean
Dim LnItemID As Long
Dim I As Integer
Dim DblNotesTotal As Double
Dim SngTemp As Single
On Error GoTo ErrTrap

Me.FG.FinishEditing True
DoEvents

Screen.MousePointer = vbArrowHourglass

If Trim(Me.TxtTransSerial.text) = "" Then
    Msg = "ÌÃ» ≈œŒ«· —ﬁ„ «·›« Ê—…...!!"
    MsgBox Msg, vbOKOnly + vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
    TxtTransSerial.SetFocus
    Screen.MousePointer = vbDefault
    Exit Sub
Else
    If Me.TxtModFlg.text = "N" Then
        BolTemp = UniqueTransSerial(Trim(Me.TxtTransSerial.text), 2)
    ElseIf Me.TxtModFlg.text = "E" Then
         BolTemp = UniqueTransSerial(Trim(Me.TxtTransSerial.text), 2, Val(Me.XPTxtBillID.text))
    End If
    If BolTemp = False Then
        Msg = "—ﬁ„ «·›« Ê—… „”Ã· „”»ﬁ« ›Ï «·»—‰«„Ã.." & Chr(13)
        Msg = Msg & "Ê·«Ì„ﬂ‰  ﬂ—«— —ﬁ„ «·›« Ê—…"
        MsgBox Msg, vbOKOnly + vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
        TxtTransSerial.SetFocus
        Screen.MousePointer = vbDefault
        Exit Sub
    End If
End If
If Val(DBCboClientName.BoundText) = 0 Then
    Msg = "„‰ ›÷·ﬂ √œŒ· «”„ «·⁄„Ì·"
    MsgBox Msg, vbOKOnly + vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
    DBCboClientName.SetFocus
    SendKeys "{F4}"
    Screen.MousePointer = vbDefault
    Exit Sub
End If
If DCboStoreName.text = "" Then
    Msg = "ÌÃ»  ÕœÌœ «·„Œ“‰"
    MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
    DCboStoreName.SetFocus
    SendKeys "{F4}"
    Screen.MousePointer = vbDefault
    Exit Sub
End If
If Trim(DcboEmp.BoundText) = "" Then
    Msg = "ÌÃ»  ÕœÌœ «”„ «·„ÊŸ›..!!!"
    MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
    DcboEmp.SetFocus
    SendKeys "{F4}"
    Screen.MousePointer = vbDefault
    Exit Sub
End If
If XPDtbBill.Value = "" Then
    Msg = "ÌÃ»  ÕœÌœ  «—ÌŒ «·»Ì⁄"
    MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
    XPDtbBill.SetFocus
    SendKeys "{F4}"
    Screen.MousePointer = vbDefault
    Exit Sub
End If
If CboPayMentType.ListIndex = -1 Then
    Msg = "ÌÃ»  ÕœÌœ ÿ—Ìﬁ… «·œ›⁄"
    MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
    CboPayMentType.SetFocus
    SendKeys "{F4}"
    Screen.MousePointer = vbDefault
    Exit Sub
End If
If XPChkPayType(0).Value = vbChecked Then
    If Me.DcboBox.BoundText = "" Then
        MsgBox "ÌÃ»  ÕœÌœ «”„ «·Œ“‰…...!!!", vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
        Screen.MousePointer = vbDefault
        Exit Sub
    End If
End If
'----------------------------------------------
If Val(Me.XPTxtValue(1).text) > 0 Then
    If ChkInstall.Value = vbChecked Then
        If Val(Me.LblInstallTotal.Caption) = 0 Then
            Msg = "ÌÃ» Õ”«» «·√ﬁ”«ÿ ﬁ»· ⁄„·Ì… «·Õ›Ÿ..!!!"
            MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
            Me.XPTab301.CurrTab = 1
            Screen.MousePointer = vbDefault
            Exit Sub
        End If
        If Val(Me.LblInstallTotal.Caption) <> Val(Me.XPTxtValue(1).text) Then
            Me.XPTxtValue(1).text = Val(Me.LblInstallTotal.Caption)
        End If
    End If
End If
'-----------------------------------------
If XPChkPayType(2).Value = vbChecked Then
    If Val(Me.lbl(18).Caption) = 0 Then
        Msg = "ÌÃ» ≈œŒ«· «·‘Ìﬂ«  ﬁ»· ⁄„·Ì… «·Õ›Ÿ..!!!"
        MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
        Me.XPTab301.CurrTab = 1
        Screen.MousePointer = vbDefault
        Exit Sub
    End If
End If
If XPChkTAX.Value = Checked Then
    If XPTxtTaxValue.text = "" Then
        Msg = "ÌÃ» «œŒ«· ﬁÌ„… ÷—Ì»… «·„»Ì⁄« "
        MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
        XPTxtTaxValue.SetFocus
        FG.SetFocus
        Screen.MousePointer = vbDefault
        Exit Sub
    End If
End If
If XPCboDiscountType.ListIndex = 1 Or XPCboDiscountType.ListIndex = 2 Then
    If XPTxtDiscountVal.text = "" Then
        Msg = "≈–« ﬂ«‰ Â‰«ﬂ Œ’„ ⁄·Ï «·›« Ê—… " & Chr(13)
        Msg = Msg + "ÌÃ»  ÕœÌœ ﬁÌ„… Â–« «·Œ’„ " & Chr(13)
        Msg = Msg + "√Ê √Œ Ì«— ·« ÌÊÃœ Œ’„ "
        MsgBox Msg, vbOKOnly + vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
        XPCboDiscountType.SetFocus
        SendKeys "{F4}"
        Screen.MousePointer = vbDefault
        Exit Sub
    End If
End If
'--------------------------------
'«·ﬂ‘› ⁄·Ï „œÌÊ‰Ì… «·⁄„Ì· '
If Val(Me.DBCboClientName.BoundText) <> 1 Or Val(Me.DBCboClientName.BoundText <> 2) Then
    If Me.CboPayMentType.ListIndex = 1 Then
        If Val(Me.XPTxtValue(1).text) > 0 Then
            If CheckCusCredit(Val(Me.DBCboClientName.BoundText), _
                Val(Me.XPTxtValue(1).text), 0) = False Then
                Screen.MousePointer = vbDefault
                Exit Sub
            End If
        End If
    End If
End If
'--------------------------------
Me.XPTab301.CurrTab = 0
If NewGrid.CheckDataEntered = False Then
    Exit Sub
End If
'-------------------------------
If NewGrid.Calculate(1, , False, True) = False Then
    Screen.MousePointer = vbDefault
    Exit Sub
End If
'-------------------------------
If Me.XPChkPayType(0).Value = vbChecked Then
    DblNotesTotal = Val(Me.XPTxtValue(0).text)
End If
If Me.XPChkPayType(1).Value = vbChecked Then
    DblNotesTotal = DblNotesTotal + Val(Me.XPTxtValue(1).text)
End If
If Me.XPChkPayType(2).Value = vbChecked Then
    DblNotesTotal = DblNotesTotal + Val(Me.lbl(18).Caption)
End If
If DblNotesTotal <> Val(LblTotal.Caption) Then
    Msg = "≈Ã„«·Ï «·√Ê—«ﬁ «·„«·Ì… €Ì— „ ”«ÊÏ „⁄ ≈Ã„«·Ï «·›« Ê—…...!!!"
    MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
    Exit Sub
End If
'---------------------------------
Set RSTransDetails = New ADODB.Recordset
RSTransDetails.Open "[Transaction_Details]", Cn, adOpenStatic, adLockOptimistic, adCmdTable
Set RsNotes = New ADODB.Recordset
RsNotes.Open "[Notes]", Cn, adOpenStatic, adLockOptimistic, adCmdTable
If SystemOptions.SysRegisterState <> Registered And SystemOptions.SysRegisterState <> DevelopVersion Then
    If Rs.RecordCount > 50 Then
        'Exit Sub
    End If
End If
Screen.MousePointer = vbArrowHourglass
Cn.BeginTrans
TransBegine = True
If Me.TxtModFlg.text = "N" Then
    XPTxtBillID.text = CStr(new_id("Transactions", "Transaction_ID", "", True))
    Rs.AddNew
    Rs("Transaction_ID").Value = Val(XPTxtBillID.text)
ElseIf Me.TxtModFlg.text = "E" Then
    StrSqlDel = "delete From Transaction_Details where Transaction_ID=" & Val(Rs("Transaction_ID").Value)
    Cn.Execute StrSqlDel, , adExecuteNoRecords
    StrSqlDel = "delete From Notes where Transaction_ID=" & Val(Rs("Transaction_ID").Value)
    Cn.Execute StrSqlDel, , adExecuteNoRecords
    StrSQL = "Delete From DOUBLE_ENTREY_VOUCHERS Where Transaction_ID=" & Val(Me.XPTxtBillID.text)
    Cn.Execute StrSQL, , adExecuteNoRecords
End If
Rs("Transaction_Serial").Value = IIf(Trim(Me.TxtTransSerial.text) = "", "", Trim(Me.TxtTransSerial.text))
Rs("Transaction_Date").Value = XPDtbBill.Value
Rs("Transaction_Type").Value = 2
Rs("UserID").Value = User_ID

If XPCboDiscountType.ListIndex = -1 Then
    Rs("Trans_DiscountType").Value = 0
Else
    Rs("Trans_DiscountType").Value = Val(XPCboDiscountType.ListIndex)
End If
Rs("Trans_Discount").Value = IIf(XPTxtDiscountVal.text = "", Null, Val(XPTxtDiscountVal.text))
Rs("CusID").Value = IIf(DBCboClientName.BoundText = "", Null, Val(DBCboClientName.BoundText))
Rs("StoreID").Value = IIf(DCboStoreName.BoundText = "", Null, Val(DCboStoreName.BoundText))
If CboPayMentType.ListIndex = -1 Then
    Rs("PaymentType").Value = 0
Else
    Rs("PaymentType").Value = Val(CboPayMentType.ListIndex)
End If

Rs("TaxFound").Value = IIf(XPChkTAX.Value = Checked, True, False)
Rs("TaxValue").Value = IIf(XPTxtTaxValue.text = "", Null, Val(XPTxtTaxValue.text))
Rs("Emp_ID").Value = IIf(DcboEmp.BoundText = "", Null, DcboEmp.BoundText)
If Me.CboSaleType.ListIndex = 0 Or Me.CboSaleType.ListIndex = -1 Then
    Rs("SaleType").Value = 0
Else
    Rs("SaleType").Value = 1
End If
If Trim$(Me.TxtCashCustomerName.text) <> "" Then
    Rs("CashCustomerName").Value = Trim$(Me.TxtCashCustomerName.text)
Else
    Rs("CashCustomerName").Value = Null
End If
Rs("TransactionComment").Value = IIf(Trim$(TxtBillComment.text) = "", Null, Trim$(TxtBillComment.text))
'÷—»Ì… Œ’„ Ê≈÷«›…
If ChkTaxAdd.Value = vbChecked And Val(Me.TxtTaxAddValue.text) > 0 Then
    Rs("TaxAddValue").Value = Val(Me.TxtTaxAddValue.text)
Else
    Rs("TaxAddValue").Value = 0
End If
'÷—»Ì… œ„€…
If ChkTaxStamp.Value = vbChecked And Val(Me.TxtTaxStampValue.text) > 0 Then
    Rs("TaxStampValue").Value = Val(Me.TxtTaxStampValue.text)
Else
    Rs("TaxStampValue").Value = 0
End If
'÷—»Ì… Œœ„…
If ChkTaxSerivce.Value = vbChecked And Val(Me.TxtTaxServiceValue.text) > 0 Then
    Rs("TaxServiceValue").Value = Val(Me.TxtTaxServiceValue.text)
Else
    Rs("TaxServiceValue").Value = 0
End If
Rs.update


For RowNum = 1 To FG.Rows - 1
    If FG.TextMatrix(RowNum, FG.ColIndex("Code")) <> "" Then
         'Check Repeat Serial
        If FG.TextMatrix(RowNum, FG.ColIndex("Serial")) <> "" Then
            StrSQL = "select * From Transaction_Details where ItemSerial='" & FG.TextMatrix(RowNum, FG.ColIndex("Serial")) & "'"
            StrSQL = StrSQL + " and Transaction_ID =" & XPTxtBillID.text
            Set RsTemp = New ADODB.Recordset
            RsTemp.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText
            If Not (RsTemp.EOF Or RsTemp.BOF) Then
                Msg = "«·”Ì—Ì«· «·Œ«’ »«·’‰›" & Chr(13)
                Msg = Msg + FG.Cell(flexcpTextDisplay, RowNum, FG.ColIndex("name")) & Chr(13)
                Msg = Msg + " „ √œŒ«·Â ·ﬁÿ⁄… √Œ—Ï ›Ì Â–Â «·›« Ê—…"
                MsgBox Msg, vbOKOnly + vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
                RsTemp.Close
                XPTab301.CurrTab = 0
                FG.Row = RowNum
                FG.Col = FG.ColIndex("name")
                FG.ShowCell RowNum, FG.ColIndex("name")
                FG.SetFocus
                
                TransBegine = False
                Cn.RollbackTrans
                Screen.MousePointer = vbDefault
                Exit Sub
            End If
            RsTemp.Close
        End If
        If IsEmpty(Me.FG.Cell(flexcpData, RowNum, FG.ColIndex("UnitID"))) Then
            If Val(Me.FG.Cell(flexcpData, RowNum, FG.ColIndex("UnitID"))) = 0 Then
                Msg = " ÌÃ»  ÕœÌœ ÊÕœ… «·ﬂ„Ì… «·Œ«’… »«·’‰›" & Chr(13)
                Msg = Msg + FG.Cell(flexcpTextDisplay, RowNum, FG.ColIndex("Name")) & Chr(13)
                MsgBox Msg, vbOKOnly + vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
                XPTab301.CurrTab = 0
                FG.Row = RowNum
                FG.Col = FG.ColIndex("UnitID")
                FG.ShowCell RowNum, FG.ColIndex("UnitID")
                FG.SetFocus
                Screen.MousePointer = vbDefault
                Exit Sub
            End If
        End If

        RSTransDetails.AddNew
        RSTransDetails("Transaction_ID").Value = Val(XPTxtBillID.text)
        RSTransDetails("Item_ID").Value = IIf((FG.TextMatrix(RowNum, FG.ColIndex("Code")) = ""), Null, Val(FG.TextMatrix(RowNum, FG.ColIndex("Code"))))
        'RSTransDetails("Quantity").Value = IIf((FG.TextMatrix(RowNum, FG.ColIndex("Count")) = ""), Null, Val(FG.TextMatrix(RowNum, FG.ColIndex("Count"))))
'            RSTransDetails("ItemName").Value = IIf((FG.TextMatrix(RowNum, FG.ColIndex("Name")) = ""), Null, Val(FG.TextMatrix(RowNum, FG.ColIndex("Name"))))
        If Not FG.TextMatrix(RowNum, FG.ColIndex("Name")) = "" Then
            StrSQL = "select * From TblItems where ItemID=" & FG.TextMatrix(RowNum, FG.ColIndex("Name"))
            RsTemp.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText
            If Not (RsTemp.EOF Or RsTemp.BOF) Then
                If RsTemp("HaveSerial").Value = True Then
                    RSTransDetails("ItemSerial").Value = IIf((FG.TextMatrix(RowNum, FG.ColIndex("Serial")) = ""), Null, FG.TextMatrix(RowNum, FG.ColIndex("Serial")))
                End If
            End If
            RsTemp.Close
        End If
        RSTransDetails("Price").Value = IIf((FG.TextMatrix(RowNum, FG.ColIndex("Price")) = ""), Null, Val(FG.TextMatrix(RowNum, FG.ColIndex("Price"))))
        RSTransDetails("ItemDiscountType").Value = IIf((FG.TextMatrix(RowNum, FG.ColIndex("DiscountType")) = ""), Null, Val(FG.TextMatrix(RowNum, FG.ColIndex("DiscountType"))))
        RSTransDetails("ItemCase").Value = IIf((FG.TextMatrix(RowNum, FG.ColIndex("ItemCase")) = ""), Null, Val(FG.TextMatrix(RowNum, FG.ColIndex("ItemCase"))))
        RSTransDetails("ItemDiscount").Value = IIf((FG.TextMatrix(RowNum, FG.ColIndex("DiscountVal")) = ""), 0, Val(FG.TextMatrix(RowNum, FG.ColIndex("DiscountVal"))))
        RSTransDetails("guaranteeTime").Value = _
        IIf((FG.TextMatrix(RowNum, FG.ColIndex("guaranteeTime")) = ""), Null, _
        Val(FG.TextMatrix(RowNum, FG.ColIndex("guaranteeTime"))))
        RSTransDetails("CostPrice").Value = _
                IIf((FG.TextMatrix(RowNum, FG.ColIndex("ItemCostPrice")) = ""), Null, Val(FG.TextMatrix(RowNum, FG.ColIndex("ItemCostPrice"))))
        RSTransDetails("CostTransID").Value = _
                IIf((FG.TextMatrix(RowNum, FG.ColIndex("PofTransID")) = ""), Null, Val(FG.TextMatrix(RowNum, FG.ColIndex("PofTransID"))))
        RSTransDetails("ItemProfit").Value = _
                IIf((FG.TextMatrix(RowNum, FG.ColIndex("ItemProfit")) = ""), Null, Val(FG.TextMatrix(RowNum, FG.ColIndex("ItemProfit"))))
        
        RSTransDetails("UnitID").Value = _
         IIf(FG.Cell(flexcpData, RowNum, FG.ColIndex("UnitID")) = "", Null, (FG.Cell(flexcpData, RowNum, FG.ColIndex("UnitID"))))
       RSTransDetails("ShowQty").Value = IIf((FG.TextMatrix(RowNum, FG.ColIndex("Count")) = ""), Null, Val(FG.TextMatrix(RowNum, FG.ColIndex("Count"))))

        RSTransDetails("Remarks").Value = _
                IIf((FG.TextMatrix(RowNum, FG.ColIndex("Remarks")) = ""), Null, Trim$(FG.TextMatrix(RowNum, FG.ColIndex("Remarks"))))
                
        RSTransDetails("ColorID").Value = _
            IIf((FG.TextMatrix(RowNum, FG.ColIndex("ColorID")) = ""), 1, Val(FG.TextMatrix(RowNum, FG.ColIndex("ColorID"))))
        RSTransDetails("ItemSize").Value = _
            IIf((FG.TextMatrix(RowNum, FG.ColIndex("ItemSize")) = ""), "", Trim$(FG.TextMatrix(RowNum, FG.ColIndex("ItemSize"))))
           
            Dim RsUnitData As ADODB.Recordset
            Dim LngCurItemID As Long
            Dim LngUnitID As Long
            Dim DblQty As Double

        
            LngCurItemID = Val(FG.TextMatrix(RowNum, FG.ColIndex("Code")))
            LngUnitID = Val(FG.Cell(flexcpData, RowNum, FG.ColIndex("UnitID")))
            DblQty = Val(FG.TextMatrix(RowNum, FG.ColIndex("Count")))

            StrSQL = "Select * From TblItemsUnits Where ItemID=" & LngCurItemID
            StrSQL = StrSQL + " AND UnitID=" & LngUnitID
            Set RsUnitData = New ADODB.Recordset
            RsUnitData.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText
            If Not (Rs.BOF Or Rs.EOF) Then
                RSTransDetails("QtyBySmalltUnit").Value = RsUnitData("FactorBySmallUnit").Value
                RSTransDetails("Quantity").Value = RSTransDetails("QtyBySmalltUnit").Value * RSTransDetails("showqty").Value
            End If

         
        RSTransDetails.update
        '-------------
        
    End If
Next RowNum

If Me.XPChkPayType(0).Value = Checked Then
    RsNotes.AddNew
    RsNotes("NoteID").Value = CStr(new_id("Notes", "NoteID", "", True))
    If Me.TxtModFlg.text = "N" Then
        RsNotes("NoteSerial").Value = CStr(new_id("Notes", "NoteSerial", "", True))
        XPTxtSerial(0).text = RsNotes("NoteSerial").Value
    ElseIf Trim(XPTxtSerial(0).text) <> "" Then
        RsNotes("NoteSerial").Value = Trim(XPTxtSerial(0).text)
    Else
        RsNotes("NoteSerial").Value = CStr(new_id("Notes", "NoteSerial", "", True))
        XPTxtSerial(0).text = RsNotes("NoteSerial").Value
    End If
    RsNotes("Transaction_ID").Value = Val(XPTxtBillID.text)
    RsNotes("NoteDate").Value = XPDtbBill.Value
    RsNotes("NoteType").Value = 0
    RsNotes("Note_Value").Value = IIf(XPTxtValue(0).text = "", Null, Val(XPTxtValue(0).text))
    RsNotes("Member_ID").Value = IIf(DBCboClientName.BoundText = "", Null, Val(DBCboClientName.BoundText))
    RsNotes("BankID").Value = Null
    RsNotes("BoxID").Value = IIf(DcboBox.BoundText = "", Null, Val(DcboBox.BoundText))
    RsNotes("CUSID").Value = Null
    
    RsNotes.update
End If
'«·ﬁÌ„ «·¬Ã·…
If Me.XPChkPayType(1).Value = Checked Then
    RsNotes.AddNew
    RsNotes("NoteID").Value = CStr(new_id("Notes", "NoteID", "", True))
    XPTxtValue(1).Tag = IIf(IsNull(RsNotes("NoteID").Value), "", (RsNotes("NoteID").Value))
    Note_ID = RsNotes("NoteID").Value
    RsNotes("NoteDate").Value = XPDtbBill.Value
    If Me.TxtModFlg.text = "N" Then
        RsNotes("NoteSerial").Value = CStr(new_id("Notes", "NoteSerial", "", True))
        XPTxtSerial(1).text = RsNotes("NoteSerial").Value
    ElseIf Trim(XPTxtSerial(1).text) <> "" Then
        RsNotes("NoteSerial").Value = Trim(XPTxtSerial(1).text)
    Else
        RsNotes("NoteSerial").Value = CStr(new_id("Notes", "NoteSerial", "", True))
        XPTxtSerial(1).text = RsNotes("NoteSerial").Value
    End If
    RsNotes("Transaction_ID").Value = Val(XPTxtBillID.text)
    RsNotes("NoteType").Value = 1
    RsNotes("Note_Value").Value = IIf(XPTxtValue(1).text = "", Null, Val(XPTxtValue(1).text))
    RsNotes("Member_ID").Value = IIf(DBCboClientName.BoundText = "", Null, Val(DBCboClientName.BoundText))
    RsNotes("BankID").Value = Null
    RsNotes("CUSID").Value = Null
    RsNotes("DueDate").Value = DtpDelayDate.Value
    RsNotes.update
End If
If Me.XPChkPayType(2).Value = Checked Then
    With Me.FgCheques
        For I = .FixedRows To .Rows - 1
            RsNotes.AddNew
                RsNotes("NoteID").Value = CStr(new_id("Notes", "NoteID", "", True))
                RsNotes("NoteDate").Value = XPDtbBill.Value
                RsNotes("Transaction_ID").Value = Val(XPTxtBillID.text)
                RsNotes("NoteType").Value = 2
                RsNotes("Note_Value").Value = Val(.TextMatrix(I, .ColIndex("CheckValue")))
                RsNotes("BankID").Value = Val(.TextMatrix(I, .ColIndex("BankID")))
                RsNotes("ChqueNum").Value = Trim$(.TextMatrix(I, .ColIndex("CheckNumber")))
                RsNotes("DueDate").Value = CDate(Trim$(.TextMatrix(I, .ColIndex("DueDate"))))
                RsNotes("Member_ID").Value = Val(Me.DBCboClientName.BoundText)
                RsNotes("CUSID").Value = Val(Me.DBCboClientName.BoundText)
            RsNotes.update
        Next I
    End With
End If
'Õ›Ÿ «·√›”«ÿ
If Me.XPChkPayType(1).Value = Checked Then
    If ChkInstall.Value = vbChecked Then
        'Save installment Data
        Set RsTemp = New ADODB.Recordset
        RsTemp.Open "InstallMent", Cn, adOpenStatic, adLockOptimistic, adCmdTable
        RsTemp.AddNew
            RsTemp("PartID").Value = CStr(new_id("InstallMent", "PartID", "", True))
            RsTemp("NoteID").Value = Note_ID
            RsTemp("BasicAmmount").Value = IIf(XPTxtValue(1).text = "", 0, Val(XPTxtValue(1).text))
            RsTemp("InterestType").Value = Val(Me.LblPrecenType.Tag)
            RsTemp("InterestVal").Value = Val(LblPrecenValue.Caption)
            RsTemp("Total").Value = Val(LblInstallTotal.Caption)
            RsTemp("InstallCount").Value = Val(LblInstallCount.Caption)
            RsTemp("FirstInstallDate").Value = CDate(Me.LblFirstInstallDate.Caption)
            If Val(LblInstallmentType.Tag) = 0 Then
                RsTemp("InstallmentType").Value = 0
            ElseIf Val(LblInstallmentType.Tag) = 1 Then
                RsTemp("InstallmentType").Value = 1
            ElseIf Val(LblInstallmentType.Tag) = 2 Then
                RsTemp("InstallmentType").Value = 2
            End If
            RsTemp("InstallSeprator").Value = Val(Me.LblInstallSeprator.Caption)
            RsTemp("StartValue").Value = IIf(Val(Me.LblStartValue.Caption) = 0, Null, Val(Me.LblStartValue.Caption))
            RsTemp("CustID").Value = IIf(DBCboClientName.BoundText = "", Null, Val(DBCboClientName.BoundText))
            RsTemp("Type").Value = 0
        RsTemp.update
        'save installment Details
        Set RsDetalis = New ADODB.Recordset
        RsDetalis.Open "InstallMentDetails", Cn, adOpenStatic, adLockOptimistic, adCmdTable
        With Me.FgInstallments
            For RowNum = 1 To .Rows - 1
                RsDetalis.AddNew
                    RsDetalis("QestID").Value = CStr(new_id("InstallMentDetails", "QestID", "", True))
                    RsDetalis("PartID").Value = RsTemp("PartID").Value
                    RsDetalis("QeqtNum").Value = IIf(.TextMatrix(RowNum, .ColIndex("Serial")) = "", "", .TextMatrix(RowNum, .ColIndex("Serial")))
                    RsDetalis("Value").Value = IIf(.TextMatrix(RowNum, .ColIndex("Value")) = "", "", Val(.TextMatrix(RowNum, .ColIndex("Value"))))
                    RsDetalis("DueDate").Value = IIf(.TextMatrix(RowNum, .ColIndex("Due_Date")) = "", "", .TextMatrix(RowNum, .ColIndex("Due_Date")))
                    RsDetalis("Receipt").Value = False
                RsDetalis.update
            Next RowNum
        End With
    End If
End If

Dim LngDevID As Long
Dim LngDevNO  As Integer
Dim StrTempAccountCode As String
Dim StrTempDes As String
LngDevID = new_id("DOUBLE_ENTREY_VOUCHERS", "Double_Entry_Vouchers_ID", "")
'----------------
SngTemp = NewGrid.GetItemsCostTotal
If SngTemp > 0 Then
    StrTempAccountCode = "a3a2" ' ﬂ·›… «·„»Ì⁄« 
    StrTempDes = "›« Ê—… »Ì⁄ —ﬁ„ " & Me.TxtTransSerial.text
    LngDevNO = LngDevNO + 1
    If ModAccounts.AddNewDev(LngDevID, LngDevNO, StrTempAccountCode, SngTemp, _
        0, StrTempDes, , , , , Me.XPDtbBill.Value, Me.DCboUserName.BoundText, Val(Me.XPTxtBillID.text)) = False Then
        GoTo ErrTrap
    End If
    StrTempAccountCode = "a1a2a5" '„Õ“Ê‰ «·»÷«⁄…
    StrTempDes = "›« Ê—… »Ì⁄ —ﬁ„ " & Me.TxtTransSerial.text
    LngDevNO = LngDevNO + 1
    If ModAccounts.AddNewDev(LngDevID, LngDevNO, StrTempAccountCode, SngTemp, _
        1, StrTempDes, , , , , Me.XPDtbBill.Value, Me.DCboUserName.BoundText, Val(Me.XPTxtBillID.text)) = False Then
        GoTo ErrTrap
    End If
    '----------------
    LngDevID = LngDevID + 1
    LngDevNO = 0
End If

If Me.XPChkPayType(0).Value = vbChecked Then
    StrTempAccountCode = GetMyAccountCode("TblBoxesData", "BoxID", Val(Me.DcboBox.BoundText))
    StrTempDes = "›« Ê—… »Ì⁄ —ﬁ„ " & Me.TxtTransSerial.text
    LngDevNO = LngDevNO + 1
    If ModAccounts.AddNewDev(LngDevID, LngDevNO, StrTempAccountCode, Val(Me.XPTxtValue(0).text), _
        0, StrTempDes, , , , , Me.XPDtbBill.Value, Me.DCboUserName.BoundText, Val(Me.XPTxtBillID.text)) = False Then
        GoTo ErrTrap
    End If
End If
If Me.XPChkPayType(1).Value = vbChecked Then
    StrTempAccountCode = GetMyAccountCode("TblCustemers", "CusID", Val(Me.DBCboClientName.BoundText))
    StrTempDes = "›« Ê—… »Ì⁄ —ﬁ„ " & Me.TxtTransSerial.text
    LngDevNO = LngDevNO + 1
    If ModAccounts.AddNewDev(LngDevID, LngDevNO, StrTempAccountCode, Val(Me.XPTxtValue(1).text), _
        0, StrTempDes, , , , , Me.XPDtbBill.Value, Me.DCboUserName.BoundText, Val(Me.XPTxtBillID.text)) = False Then
        GoTo ErrTrap
    End If
End If
If Me.XPChkPayType(2).Value = vbChecked Then
    StrTempAccountCode = "a1a2a4"
    StrTempDes = "⁄œœ " & Me.lbl(19).Caption & "  ‘Ìﬂ«  " & Chr(13)
    StrTempDes = "›« Ê—… »Ì⁄ —ﬁ„ " & Me.TxtTransSerial.text
    LngDevNO = LngDevNO + 1
    If ModAccounts.AddNewDev(LngDevID, LngDevNO, StrTempAccountCode, Val(Me.lbl(18).Caption), _
        0, StrTempDes, , , , , Me.XPDtbBill.Value, Me.DCboUserName.BoundText, Val(Me.XPTxtBillID.text)) = False Then
        GoTo ErrTrap
    End If
End If
If Val(Me.LblDiscountsTotal.Caption) > 0 Then
    StrTempAccountCode = "a3a5" '«·Œ’„ «·„”„ÊÕ »Â
    StrTempDes = "›« Ê—… »Ì⁄ —ﬁ„ " & Me.TxtTransSerial.text
    LngDevNO = LngDevNO + 1
    If ModAccounts.AddNewDev(LngDevID, LngDevNO, StrTempAccountCode, Val(Me.LblDiscountsTotal.Caption), _
        0, StrTempDes, , , , , Me.XPDtbBill.Value, Me.DCboUserName.BoundText, Val(Me.XPTxtBillID.text)) = False Then
        GoTo ErrTrap
    End If
End If
If Me.ChkTaxAdd.Value = vbChecked Then
    StrTempAccountCode = "a2a5a4" '÷—»Ì… √—»«Õ  Ã«—Ì… (Œ’„ Ê≈÷«›…
    StrTempDes = "›« Ê—… »Ì⁄ —ﬁ„ " & Me.TxtTransSerial.text
    SngTemp = Val(Me.lbl(52).Caption)
    LngDevNO = LngDevNO + 1
    If ModAccounts.AddNewDev(LngDevID, LngDevNO, StrTempAccountCode, SngTemp, _
        0, StrTempDes, , , , , Me.XPDtbBill.Value, Me.DCboUserName.BoundText, Val(Me.XPTxtBillID.text)) = False Then
        GoTo ErrTrap
    End If
End If
If Me.ChkTaxStamp.Value = vbChecked Then
    StrTempAccountCode = "a3a9" 'œ„€«  ÕﬂÊ„Ì…
    StrTempDes = "›« Ê—… »Ì⁄ —ﬁ„ " & Me.TxtTransSerial.text
    SngTemp = Val(Me.lbl(53).Caption)
    LngDevNO = LngDevNO + 1
    If ModAccounts.AddNewDev(LngDevID, LngDevNO, StrTempAccountCode, SngTemp, _
        0, StrTempDes, , , , , Me.XPDtbBill.Value, Me.DCboUserName.BoundText, Val(Me.XPTxtBillID.text)) = False Then
        GoTo ErrTrap
    End If
End If
'«·œ«∆‰
SngTemp = NewGrid.GetItemsTotal(ItemsGoodType)
If SngTemp > 0 Then
    StrTempAccountCode = "a4a1" '«·„»Ì⁄« 
    StrTempDes = "›« Ê—… »Ì⁄ —ﬁ„ " & Me.TxtTransSerial.text
    LngDevNO = LngDevNO + 1
    If ModAccounts.AddNewDev(LngDevID, LngDevNO, StrTempAccountCode, SngTemp, _
        1, StrTempDes, , , , , Me.XPDtbBill.Value, Me.DCboUserName.BoundText, Val(Me.XPTxtBillID.text)) = False Then
        GoTo ErrTrap
    End If
End If
SngTemp = NewGrid.GetItemsTotal(ItemsServiceType)
If SngTemp > 0 Then
    StrTempAccountCode = "a4a7" '≈Ì—«œ«  «·Œœ„« 
    StrTempDes = "›« Ê—… »Ì⁄ —ﬁ„ " & Me.TxtTransSerial.text
    LngDevNO = LngDevNO + 1
    If ModAccounts.AddNewDev(LngDevID, LngDevNO, StrTempAccountCode, SngTemp, _
        1, StrTempDes, , , , , Me.XPDtbBill.Value, Me.DCboUserName.BoundText, Val(Me.XPTxtBillID.text)) = False Then
        GoTo ErrTrap
    End If
End If
'
If XPChkTAX.Value = vbChecked Then
    StrTempAccountCode = "a1a3a5" '÷—»Ì… „»Ì⁄«  „œÌ‰…
    SngTemp = Val(Me.lbl(51).Caption)
    StrTempDes = "›« Ê—… »Ì⁄ —ﬁ„ " & Me.TxtTransSerial.text
    LngDevNO = LngDevNO + 1
    If ModAccounts.AddNewDev(LngDevID, LngDevNO, StrTempAccountCode, SngTemp, _
        1, StrTempDes, , , , , Me.XPDtbBill.Value, Me.DCboUserName.BoundText, Val(Me.XPTxtBillID.text)) = False Then
        GoTo ErrTrap
    End If
End If
If ChkTaxSerivce.Value = vbChecked Then
    StrTempAccountCode = "a4a9" '÷—»Ì… Œœ„… „»Ì⁄« 
    SngTemp = Val(Me.lbl(54).Caption)
    StrTempDes = "›« Ê—… »Ì⁄ —ﬁ„ " & Me.TxtTransSerial.text
    LngDevNO = LngDevNO + 1
    If ModAccounts.AddNewDev(LngDevID, LngDevNO, StrTempAccountCode, SngTemp, _
        1, StrTempDes, , , , , Me.XPDtbBill.Value, Me.DCboUserName.BoundText, Val(Me.XPTxtBillID.text)) = False Then
        GoTo ErrTrap
    End If
End If
Cn.CommitTrans
TransBegine = False
XPTxtCurrent.Caption = Rs.AbsolutePosition
XPTxtCount.Caption = Rs.RecordCount
Select Case Me.TxtModFlg.text
    Case "N"
        Msg = "  „ Õ›Ÿ »Ì«‰«  Â–Â «·⁄„·Ì…" & Chr(13)
        Msg = Msg + "Â·  —€» ›Ì ≈÷«›… »Ì«‰«  √Œ—Ì"
        If MsgBox(Msg, vbYesNo + vbQuestion + vbMsgBoxRight + _
            vbMsgBoxRtlReading + vbDefaultButton1, App.Title) = vbYes Then
            Cmd_Click (0)
            Screen.MousePointer = vbDefault
        Else
            TxtModFlg.text = "R"
        End If
        If MDIFrmMain.MnuInvPrintSave.Checked = True Then
            Cmd_Click (7)
        End If
    Case "E"
        MsgBox " „ Õ›Ÿ Â–Â «· ⁄œÌ·« ", vbInformation + vbOKOnly + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
        TxtModFlg.text = "R"
End Select

Screen.MousePointer = vbDefault
Exit Sub
ErrTrap:
If TransBegine = True Then
    TransBegine = False
    Cn.RollbackTrans
End If
'Resume
If Rs.EditMode <> adEditNone Then
    Rs.CancelUpdate
End If
If Not RsNotes Is Nothing Then
    If RsNotes.EditMode <> adEditNone Then
        RsNotes.CancelUpdate
    End If
End If
If Not RSTransDetails Is Nothing Then
    If RSTransDetails.EditMode <> adEditNone Then
        RSTransDetails.CancelUpdate
    End If
End If
Screen.MousePointer = vbDefault
If Err.Number = -2147217900 Then
    Msg = "·« Ì„ﬂ‰ Õ›Ÿ Â–Â «·»Ì«‰«  " & Chr(13)
    Msg = Msg + "·ﬁœ  „ «œŒ«· ﬁÌ„ €Ì— ’«·Õ… " & Chr(13)
    Msg = Msg + " √ﬂœ „‰ œﬁ… «·»Ì«‰«  Ê√⁄œ «·„Õ«Ê·…"
    Msg = Msg & Chr(13) & Err.Description
    Msg = Msg & Chr(13) & Err.Number
    Msg = Msg & Chr(13) & Err.Source
    Msg = Msg & Chr(13) & Err.LastDllError
    MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
    Exit Sub
End If
Msg = "⁄›Ê«...ÕœÀ Œÿ√ „« √À‰«¡ Õ›Ÿ Â–Â «·»Ì«‰«  " & Chr(13)
Msg = Msg & Chr(13) & Err.Description
Msg = Msg & Chr(13) & Err.Number
Msg = Msg & Chr(13) & Err.Source
Msg = Msg & Chr(13) & Err.LastDllError
MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
End Sub
Private Sub XPBtnNewClients_Click()
On Error GoTo ErrTrap
With FrmAddNewCustemer
    .DealingForm = InvoiceTransaction
    FrmAddNewCustemer.AddType = 1
    .Caption = "≈÷«›… ⁄„Ì· ÃœÌœ"
    .lbl(1).Caption = "ﬂÊœ «·⁄„Ì·"
    .lbl(0).Caption = "«”„ «·⁄„Ì·"
    Set .DcboCustomers = DBCboClientName
    .Show vbModal
    cSearchDcbo(0).Refresh
End With
Exit Sub
ErrTrap:
End Sub

Private Sub XPCboDiscountType_Change()
XPCboDiscountType_Click
End Sub
Private Sub XPCboDiscountType_Click()
On Error GoTo ErrTrap
If XPCboDiscountType.ListIndex = 0 Or XPCboDiscountType.ListIndex = 3 Or XPCboDiscountType.ListIndex = -1 Then
    
    XPTxtDiscountVal.Enabled = False
    XPTxtDiscountVal.text = ""
Else
    
    XPTxtDiscountVal.Enabled = True
    XPTxtDiscountVal.text = ""
End If
If Me.TxtModFlg.text = "N" Or Me.TxtModFlg.text = "E" Then
    If FG.TextMatrix(1, FG.ColIndex("Code")) <> "" Then
        NewGrid.Calculate 1, , , True
    End If
End If
Exit Sub
ErrTrap:
End Sub
Private Sub XPChkPayType_Click(Index As Integer)
On Error GoTo ErrTrap
Select Case Index
    Case 0
        If XPChkPayType(0).Value = Checked Then
            If Me.TxtModFlg.text = "N" Then
                XPTxtValue(0).text = ""
                XPTxtSerial(0).text = ""
            End If
            If Me.TxtModFlg.text = "N" Or Me.TxtModFlg.text = "E" Then
                XPTxtValue(0).Enabled = True
'                XPTxtSerial(0).Enabled = True
                XPTxtValue(0).Locked = False
'                XPTxtSerial(0).Locked = False
            End If
        Else
            XPTxtValue(0).Enabled = False
            XPTxtValue(0).text = ""
'            XPTxtSerial(0).Enabled = False
        End If
    Case 1
        If XPChkPayType(1).Value = Checked Then
            If Me.TxtModFlg.text = "N" Then
                XPTxtValue(1).text = ""
                XPTxtSerial(1).text = ""
                DtpDelayDate.Value = Date
                XPTxtSerial(1).text = CStr(new_id("Notes", "NoteSerial", "", True))
            End If
            If Me.TxtModFlg.text = "N" Or Me.TxtModFlg.text = "E" Then
                XPTxtValue(1).Enabled = True
                XPTxtValue(1).Locked = False
                DtpDelayDate.Enabled = True
            Else
                DtpDelayDate.Enabled = False
            End If
            Me.ChkInstall.Enabled = True
        Else
            XPTxtValue(1).Enabled = False
            XPTxtSerial(1).Enabled = False
            XPTxtValue(1).text = ""
            Me.ChkInstall.Enabled = False
        End If
    Case 2
        If XPChkPayType(2).Value = Checked And Me.TxtModFlg.text <> "R" Then
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

If XPChkTAX.Value = Checked Then
    XPTxtTaxValue.Enabled = True
    lbl(4).Enabled = True
    lbl(45).Enabled = True
Else
    XPTxtTaxValue.text = ""
    XPTxtTaxValue.Enabled = False
    lbl(4).Enabled = False
    lbl(45).Enabled = False
End If
Exit Sub
ErrTrap:
End Sub

Private Sub XPTxtDiscountVal_Change()
Dim Msg As String
On Error GoTo ErrTrap

If Me.TxtModFlg.text = "N" Or Me.TxtModFlg.text = "E" Then
    NewGrid.Calculate 1, , , True
End If
Exit Sub
ErrTrap:
End Sub
Private Sub PrintReport(Optional PrinterTarget As Boolean = False)

Dim ShowType As Integer
'Dim clrep As ClsReportProp
Dim StrPath As String
Dim Msg As String
Dim P_Target As PrintTarget

On Error GoTo ErrTrap

'If MDIFrmMain.MnuInvPrintDirect.Checked = True Then
'    P_Target = PrinterTarget
'
'End If

ShowType = GetSetting(StrAppRegPath, "View_Type", "SallReportType", 1)

If ShowType = 1 Then
    If XPTxtBillID.text <> "" Then
        Set SaleReport = New ClsSaleReport
        SaleReport.ShowSallingData XPTxtBillID.text
        If MDIFrmMain.MnuInvPrintReceipt.Checked = True Then
            SaleReport.PrintInvoiceReceipt Val(XPTxtBillID.text), P_Target
        End If
    End If
ElseIf (ShowType = 2) Or (ShowType = 4) Then
    P_Target = IIf(MDIFrmMain.MnuInvPrintSave.Checked = True, PrintTarget.PrinterTarget, PrintTarget.WindowTarget)
    If XPTxtBillID.text <> "" Then
        Set SaleReport = New ClsSaleReport
        'SaleReport.ShowSallingDataShort XPTxtBillID.text, P_Target
        SaleReport.ShowSallingData XPTxtBillID.text, 3
        'ÿ»«⁄… ≈Ì’«· ≈” ·«„ «·‰ﬁœÌ…
        If MDIFrmMain.MnuInvPrintReceipt.Checked = True Then
            SaleReport.PrintInvoiceReceipt Val(XPTxtBillID.text), P_Target
        End If
    End If
ElseIf ShowType = 3 Then
    If XPTxtBillID.text <> "" Then
        StrPath = GetSetting(StrAppRegPath, "PrintReport", "ReportPath", App.Path & "\Bill_Template\SaleMain.drp")
        If StrPath = "" Then
            Msg = "⁄›Ê« : Â‰«ﬂ Œÿ√„« ›Ì „”«— «· ﬁ—Ì— "
            MsgBox Msg, vbOKOnly + vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
            Exit Sub
        End If
        Set crep = New ClsReportProp
        crep.OpenFile = StrPath
        crep.LoadFile StrPath, FrmPreview
        crep.InvoID = XPTxtBillID.text
        crep.ShowReport
        FrmPreview.Show vbModal
        Set crep = Nothing
    End If
ElseIf ShowType = 5 Then
    If XPTxtBillID.text <> "" Then
        Set SaleReport = New ClsSaleReport
        SaleReport.ShowSallingData Val(XPTxtBillID.text), 1, Val(Me.DBCboClientName.BoundText)
        If MDIFrmMain.MnuInvPrintReceipt.Checked = True Then
            SaleReport.PrintInvoiceReceipt Val(XPTxtBillID.text), P_Target
        End If
    End If
ElseIf ShowType = 6 Then
    If XPTxtBillID.text <> "" Then
        Set SaleReport = New ClsSaleReport
        SaleReport.ShowSallingData Val(XPTxtBillID.text), 2, Val(Me.DBCboClientName.BoundText)
        
            SaleReport.PrintInvoiceReceipt Val(XPTxtBillID.text), P_Target
       
    End If
End If
Exit Sub
ErrTrap:
End Sub

Private Sub XPTxtDiscountVal_KeyPress(KeyAscii As Integer)
KeyAscii = KeyAscii_Num(KeyAscii, Me.XPTxtDiscountVal.text, 0)
End Sub

Private Sub XPTxtSum_Change()
On Error GoTo ErrTrap
If CboPayMentType.ListIndex = 0 Then
    XPChkPayType(0).Value = Checked
    XPTxtValue(0).text = XPTxtSum.text
End If
Me.LblTotal.Caption = XPTxtSum.text
CalculateInvPrecent
Exit Sub
ErrTrap:
End Sub
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
Dim IntResult As String
Dim StrMSG As String

On Error GoTo ErrTrap
If Me.TxtModFlg.text <> "R" Then
Select Case Me.TxtModFlg.text
    Case "N"
        StrMSG = "”Ê› Ì „ €·ﬁ «·‘«‘… Ê·„  ‰ Â „‰  ”ÃÌ·" & Chr(13)
        StrMSG = StrMSG & " «·»Ì«‰«  «·ÃœÌœ… «·Õ«·Ì… " & Chr(13)
        StrMSG = StrMSG & " Â·  —Ìœ «·Õ›Ÿ ﬁ»· «·Œ—ÊÃ" & Chr(13)
        StrMSG = StrMSG & "‰⁄„" & "-" & "Ì „ Õ›Ÿ «·»Ì«‰«  «·ÃœÌœ…" & Chr(13)
        StrMSG = StrMSG & "·«" & "-" & "·‰ Ì „ «·Õ›Ÿ" & Chr(13)
        StrMSG = StrMSG & "≈·€«¡ «·√„—" & "-" & "≈·€«¡ ⁄„·Ì… «·Œ—ÊÃ" & Chr(13)
    Case "E"
        StrMSG = "”Ê› Ì „ €·ﬁ «·‘«‘… Ê·„  ‰ Â „‰  ”ÃÌ·" & Chr(13)
        StrMSG = StrMSG & " «· ⁄œÌ·«  «·ÃœÌœ… ⁄·Ï «·”Ã· «·Õ«·Ï " & Chr(13)
        StrMSG = StrMSG & " Â·  —Ìœ «·Õ›Ÿ ﬁ»· «·Œ—ÊÃ" & Chr(13)
        StrMSG = StrMSG & "‰⁄„" & "-" & "Ì „ Õ›Ÿ «· ⁄œÌ·«   «·ÃœÌœ…" & Chr(13)
        StrMSG = StrMSG & "·«" & "-" & "·‰ Ì „ «·Õ›Ÿ" & Chr(13)
        StrMSG = StrMSG & "≈·€«¡ «·√„—" & "-" & "≈·€«¡ ⁄„·Ì… «·Œ—ÊÃ" & Chr(13)
End Select
IntResult = MsgBox(StrMSG, vbMsgBoxRight + vbYesNoCancel + vbMsgBoxRtlReading + vbQuestion, App.Title)
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
Public Sub Convert()
Cmd_Click (0)
End Sub
Public Sub Cala()
NewGrid.Calculate 1, , , True
End Sub
Private Sub DBCboClientName_Change()
Dim Msg As String
Dim RsTemp  As ADODB.Recordset
Dim StrSQL As String

On Error GoTo ErrTrap

If Val(DBCboClientName.BoundText) <> 0 Then
    If (DBCboClientName.BoundText = 1 Or DBCboClientName.BoundText = 2) And _
        Me.TxtModFlg.text <> "R" Then
        CboPayMentType.Locked = True
        CboPayMentType.ListIndex = 0
        Me.TxtCashCustomerName.Enabled = True
        Me.CmdCash(0).Enabled = True
        Me.CmdCash(1).Enabled = True
    Else
        CboPayMentType.Locked = False
        Me.TxtCashCustomerName.Enabled = False
        Me.CmdCash(0).Enabled = False
        Me.CmdCash(1).Enabled = False
    End If
    If Me.TxtModFlg.text = "N" Or Me.TxtModFlg.text = "E" Then
        StrSQL = "Select * From TblCustemers Where CusID=" & _
        Val(DBCboClientName.BoundText)
        Set RsTemp = New ADODB.Recordset
        RsTemp.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText
        If Not (RsTemp.BOF Or RsTemp.EOF) Then
            If Not (IsNull(RsTemp("SaleType").Value)) Then
                If RsTemp("SaleType").Value = 0 Then
                    Me.CboSaleType.ListIndex = 0
                ElseIf RsTemp("SaleType").Value = 1 Then
                    Me.CboSaleType.ListIndex = 1
                End If
            Else
                Me.CboSaleType.ListIndex = -1
            End If
            If Not (IsNull(RsTemp("Trans_DiscountType").Value)) Then
                If RsTemp("Trans_DiscountType").Value = 0 Then
                    Me.XPCboDiscountType.ListIndex = 0
                    Me.XPTxtDiscountVal.text = 0
                ElseIf RsTemp("Trans_DiscountType").Value = 1 Then
                    Me.XPCboDiscountType.ListIndex = 1
                    Me.XPTxtDiscountVal.text = IIf(IsNull(RsTemp("Trans_Discount").Value), "", RsTemp("Trans_Discount").Value)
                ElseIf RsTemp("Trans_DiscountType").Value = 2 Then
                    Me.XPCboDiscountType.ListIndex = 2
                    Me.XPTxtDiscountVal.text = IIf(IsNull(RsTemp("Trans_Discount").Value), "", RsTemp("Trans_Discount").Value)
                End If
            Else
                Me.XPCboDiscountType.ListIndex = 0
                Me.XPTxtDiscountVal.text = 0
            End If
        Else
            Me.CboSaleType.ListIndex = -1
            Me.XPCboDiscountType.ListIndex = 0
            Me.XPTxtDiscountVal.text = 0
        End If
        RsTemp.Close
        Set RsTemp = Nothing
    End If
End If
Exit Sub
ErrTrap:
Msg = Err.Description & Chr(13) & ""
Msg = Msg & Err.Source & Chr(13) & ""
Msg = Msg & Me.name & " DBCboClientName_Change:" & Chr(13) & ""
MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
End Sub

Private Sub DBCboClientName_Click(Area As Integer)
DBCboClientName_Change
End Sub

Private Sub XPTxtValue_Change(Index As Integer)
On Error GoTo ErrTrap
If Me.TxtModFlg.text = "N" Or Me.TxtModFlg.text = "E" Then
    If XPTxtValue(1).text <> "" Then
        If Val(Me.XPTxtValue(1).text) > 0 Then
            ChkInstall.Enabled = True
        End If
    End If
End If
Exit Sub
ErrTrap:
End Sub
Public Sub ReplacementData()
Dim Msg As String
On Error GoTo ErrTrap
Dim StrSQL As String
Dim RsReplace As ADODB.Recordset
If Me.TxtModFlg.text <> "R" Then Exit Sub
'«·»ÕÀ ⁄‰ ⁄„·Ì«  «·«” »œ«· «·Œ«’… »«·›« Ê—…
If FG.TextMatrix(FG.Row, FG.ColIndex("Code")) <> "" And FG.TextMatrix(FG.Row, FG.ColIndex("Serial")) <> "" Then
    StrSQL = "select * From ReplacedItems where ReturnID=" & XPTxtBillID.text
    StrSQL = StrSQL + " and ItemID=" & FG.TextMatrix(FG.Row, FG.ColIndex("Code"))
    StrSQL = StrSQL + " and ItemSerial='" & FG.TextMatrix(FG.Row, FG.ColIndex("Serial")) & "'"
    Set RsReplace = New ADODB.Recordset
    RsReplace.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText
    If Not (RsReplace.EOF Or RsReplace.BOF) Then
        Msg = "·ﬁœ  „ «” »œ«· «·ﬁÿ⁄… : " & FG.Cell(flexcpTextDisplay, FG.Row, FG.ColIndex("Name")) & Chr(13)
        Msg = Msg + "–«  «·”Ì—Ì«· : " & FG.TextMatrix(FG.Row, FG.ColIndex("Serial")) & Chr(13)
        Msg = Msg + " »«·ﬁÿ⁄… –«  «·”Ì—Ì«· : " & RsReplace("newSerial").Value & Chr(13)
        Msg = Msg + "›Ì ⁄„·Ì… ’Ì«‰…"
        MsgBox Msg, vbOKOnly + vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, "ﬁÿ⁄…  „ «” »œ«·Â«"
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

For RowNum = 1 To FG.Rows - 1
    If FG.TextMatrix(RowNum, FG.ColIndex("Code")) <> "" Then
        StrSQL = "select * From QryDelPurchase where Transaction_Date >=" & SQLDate(XPDtbBill.Value, True) & ""
        StrSQL = StrSQL + " and Item_ID=" & FG.TextMatrix(RowNum, FG.ColIndex("Code"))
        StrSQL = StrSQL + " and Transaction_Type=9"

        If FG.Cell(flexcpChecked, RowNum, FG.ColIndex("HaveSerial")) = flexChecked Then
            If FG.TextMatrix(RowNum, FG.ColIndex("Serial")) <> "" Then
                StrSQL = StrSQL + " and ItemSerial='" & FG.TextMatrix(RowNum, FG.ColIndex("Serial")) & "'"
            End If
        End If
        Set RsSalle = New ADODB.Recordset
        RsSalle.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText
        If Not (RsSalle.EOF Or RsSalle.BOF) Then
            If FG.Cell(flexcpChecked, RowNum, FG.ColIndex("HaveSerial")) = flexChecked Then
'                StrSql = "select * From QryGardComplete where ItemID=" & FG.TextMatrix(RowNum, FG.ColIndex("Code"))
'                StrSql = StrSql + " AND ItemSerial='" & FG.TextMatrix(RowNum, FG.ColIndex("Serial")) & "'"
'                StrSql = StrSql + " AND StoreID=" & DCboStoreName.BoundText
'                Set RsTemp = New ADODB.Recordset
'                RsTemp.Open StrSql, Cn, adOpenStatic, adLockOptimistic, adCmdText
'                If RsTemp.EOF Or RsTemp.BOF Then
                    With FrmAlarm
                        .DealingForm = InvoiceTransaction
                        .Show vbModal
                    End With
                    AvailableDeal = False
                    Exit Function
'                End If
                RsTemp.Close
            Else
                LngItemID = Val(FG.TextMatrix(RowNum, FG.ColIndex("Code")))
                Set RsTemp = New ADODB.Recordset
                Set RsTemp = GetItemQuantityStock(LngItemID, Me.DCboStoreName.BoundText, Me.XPDtbBill.Value, Val(Me.XPTxtBillID.text))
                If Not (RsTemp.EOF Or RsTemp.BOF) Then
                    If Val(RsTemp("QTY").Value) < Val(FG.TextMatrix(RowNum, FG.ColIndex("Count"))) Then
                        With FrmAlarm
                            .DealingForm = InvoiceTransaction
                            .Show vbModal
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



Private Sub SetDefaults()
Dim StrTemp As String
Dim RsTemp As ADODB.Recordset

Me.CboSaleType.ListIndex = 0
If SystemOptions.SysPurDateTakeType = InvDateFromLocalCompuer Then
     XPDtbBill.Value = Date
ElseIf SystemOptions.SysPurDateTakeType = InvDateFromServerComputer Then
    StrTemp = "select Getdate() as ServerDate"
    Set RsTemp = New ADODB.Recordset
    RsTemp.Open StrTemp, Cn, adOpenStatic, adLockReadOnly, adCmdText
    If Not (RsTemp.BOF Or RsTemp.EOF) Then
           If Not IsNull(RsTemp("ServerDate").Value) Then
               XPDtbBill.Value = Format(RsTemp("ServerDate").Value, "yyyy/M/d")
           End If
           'XPDtbBill.Value = IIf(IsNull(RsTemp("ServerDate").Value), Date, (RsTemp("ServerDate").Value))
       End If
    RsTemp.Close
    Set RsTemp = Nothing
End If

If Not (Rs.BOF Or Rs.EOF) Then
    Rs.MoveLast
    If SystemOptions.SysPurDateTakeType = InvDateFromLastInvDate Then
        XPDtbBill.Value = IIf(IsNull(Rs("Transaction_Date").Value), Date, (Rs("Transaction_Date").Value))
    End If
    Me.DcboEmp.BoundText = IIf(IsNull(Rs("Emp_ID").Value), "", Rs("Emp_ID").Value)
    If Not IsNull(Rs("Transaction_Serial").Value) Then
        StrTemp = Rs("Transaction_Serial").Value
        StrTemp = Val(StrTemp) + 1
        TxtTransSerial.text = StrTemp
    Else
        TxtTransSerial.text = 1
    End If
Else
     TxtTransSerial.text = 1
End If
End Sub

Private Sub CalculateInvPrecent()
Dim DblInvTotal As Double
Dim DblInvProfit As Double
Dim DblRes As Double

DblInvProfit = Val(Me.LblInvProfit.Caption)
DblInvTotal = Val(Me.XPTxtSum.text)
If DblInvProfit = 0 Or DblInvTotal = 0 Then
    DblRes = 0
Else
    DblRes = 100 * (DblInvProfit / DblInvTotal)
End If
Me.lblInvPrecent.Caption = "%" & CStr(Int(DblRes)) 'Format(DblRes, SystemOptions.SysDefCurrencyForamt)
End Sub

Private Sub LoadCombosData()
Dcombos.GetEmployees Me.DcboEmp
Dcombos.GetUsers Me.DCboUserName
Dcombos.GetBoxes Me.DcboBox
Dcombos.GetCustomersSuppliers 0, Me.DBCboClientName
Dcombos.GetStores Me.DCboStoreName

Set cSearchDcbo(0) = New clsDCboSearch
Set cSearchDcbo(0).Client = Me.DBCboClientName
cSearchDcbo(0).SetBuddyText Me.TxtCusID

Set cSearchDcbo(1) = New clsDCboSearch
Set cSearchDcbo(1).Client = Me.DCboStoreName
cSearchDcbo(1).SetBuddyText Me.TxtStoreID



Set cSearchDcbo(3) = New clsDCboSearch
Set cSearchDcbo(3).Client = Me.DcboEmp
cSearchDcbo(3).SetBuddyText Me.TxtEmployeeID
End Sub

Private Sub ChangeLang()
Dim XPic As IPictureDisp
Set XPic = Me.XPBtnMove(1).ButtonImage
Set Me.XPBtnMove(1).ButtonImage = Me.XPBtnMove(2).ButtonImage
Set Me.XPBtnMove(2).ButtonImage = XPic

Set XPic = Me.XPBtnMove(0).ButtonImage
Set Me.XPBtnMove(0).ButtonImage = Me.XPBtnMove(3).ButtonImage
Set Me.XPBtnMove(3).ButtonImage = XPic

Me.LblShortcutKeys.Caption = "(New F12 OR Enter) ,(Edit F11),(Save F10),(Undo F9),(Delete F8),(Search F7)"
Me.Caption = "Bill Invoice"
Ele(9).Caption = Me.Caption
lbl(5).Caption = "Invoice ID"
lbl(6).Caption = "Invoice Date"
lbl(7).Caption = "Customer Name"
lbl(24).Caption = "Store Name"
lbl(25).Caption = "Employee Name"
lbl(9).Caption = "Payment Type"
lbl(10).Caption = "Discount Type"
lbl(8).Caption = "Discount Value"
lbl(22).Caption = "Profit Value"
lbl(23).Caption = "Profit Perce"

lbl(3).Caption = "Invoice Total:"
lbl(1).Caption = "Record By:"
lbl(2).Caption = "Records Count:"

lbl(31).Caption = "Item Code"
lbl(30).Caption = "Item Name"
lbl(29).Caption = "Item Case"
lbl(28).Caption = "Item Serial"
lbl(27).Caption = "Quantity"
lbl(26).Caption = "Price"
lbl(32).Caption = "Sales Type"
lbl(33).Caption = "Cash CustomerName"
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
    lbl(11).Caption = "Box Name"
    lbl(21).Caption = "Due Date"
    
    lbl(18).Caption = "Check NO."
    lbl(17).Caption = "Bank Name"
    lbl(19).Caption = "Due Date"
    CmdINSTALLMENT.Caption = "INSTALLMENT"
Me.XPTab301.TabCaption(2) = "Comment On Invoice"
    Me.Ele(15).Caption = "Write any Comments about this Invoice"
End Sub

Private Sub XPTxtValue_KeyPress(Index As Integer, KeyAscii As Integer)
KeyAscii = KeyAscii_Num(KeyAscii, Me.XPTxtValue(Index).text, 0)
End Sub

Private Function CheckCashCustomer() As Boolean
Dim Rs As ADODB.Recordset
Dim StrSQL As String
If Trim$(Me.TxtCashCustomerName.text) = "" Then
    CheckCashCustomer = True
Else
    StrSQL = "Select * From Transactions Where CashCustomerName='" & _
    Trim$(Me.TxtCashCustomerName.text) & "'"
    
End If
End Function

Private Sub XPTxtValue_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
If Val(Me.XPTxtValue(Index).text) <> 0 Then
    Me.XPTxtValue(Index).ToolTipText = WriteNo(Me.XPTxtValue(Index).text, 1, True)
Else
    Me.XPTxtValue(Index).ToolTipText = ""
End If
End Sub



Private Sub SumChecks()
With Me.FgCheques
    If .Rows > 1 Then
        Me.lbl(19).Caption = .Aggregate(flexSTCount, .FixedRows, .ColIndex("CheckNumber"), .Rows - 1, _
        .ColIndex("CheckNumber"))
        Me.lbl(18).Caption = .Aggregate(flexSTSum, .FixedRows, .ColIndex("CheckValue"), .Rows - 1, _
        .ColIndex("CheckValue"))
    Else
        Me.lbl(19).Caption = 0
        Me.lbl(18).Caption = 0
    End If
End With
End Sub
Private Sub ClearNotes()

LblPrecenType.Caption = 0
LblPrecenValue.Caption = 0
LblInstallTotal.Caption = 0
LblInstallCount.Caption = 0
LblFirstInstallDate.Caption = ""
LblInstallSeprator.Caption = ""
LblInstallmentType.Caption = ""
LblStartValue.Caption = ""
lbl(19).Caption = ""
lbl(18).Caption = ""
End Sub
