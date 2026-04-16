VERSION 5.00
Object = "{C115893A-A3BF-43AF-B28D-69DB846077F3}#1.0#0"; "vsflex8u.ocx"
Object = "{0AFE7BE0-11B7-4A3E-978D-D4501E9A57FE}#1.0#0"; "c1sizer.ocx"
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{FE5DCFAD-BC1D-11D2-94CF-004005455FAA}#1.4#0"; "ImpulseButton.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#12.0#0"; "SuiteCtrls.ocx"
Object = "{D155F1AE-D9A4-458C-8CEE-498CB717DB7B}#1.0#0"; "DBPix20.ocx"
Begin VB.Form FrmPO10 
   Caption         =   "    «„— «·‘—«¡ "
   ClientHeight    =   9825
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   17430
   HelpContextID   =   340
   Icon            =   "FrmPO10.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   9825
   ScaleWidth      =   17430
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   8
      Charset         =   178
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Begin C1SizerLibCtl.C1Elastic EleMain 
      Height          =   9825
      Left            =   0
      TabIndex        =   22
      TabStop         =   0   'False
      Top             =   0
      Width           =   17430
      _cx             =   30745
      _cy             =   17330
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
      Begin C1SizerLibCtl.C1Elastic C1Elastic4 
         Height          =   615
         Left            =   0
         TabIndex        =   191
         TabStop         =   0   'False
         Top             =   9210
         Width           =   17430
         _cx             =   30745
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
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         FloodColor      =   6553600
         ForeColorDisabled=   -2147483631
         Caption         =   ""
         Align           =   2
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
            Height          =   405
            Index           =   12
            Left            =   5505
            TabIndex        =   202
            Top             =   120
            Width           =   1230
            _ExtentX        =   2170
            _ExtentY        =   714
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
         Begin ImpulseButton.ISButton Cmd 
            Height          =   405
            Index           =   0
            Left            =   14325
            TabIndex        =   192
            Top             =   120
            Width           =   1020
            _ExtentX        =   1799
            _ExtentY        =   714
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
            Height          =   405
            Index           =   1
            Left            =   12555
            TabIndex        =   193
            Top             =   120
            Width           =   1020
            _ExtentX        =   1799
            _ExtentY        =   714
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
            Height          =   405
            Index           =   2
            Left            =   11160
            TabIndex        =   15
            Top             =   120
            Width           =   1020
            _ExtentX        =   1799
            _ExtentY        =   714
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
            Height          =   405
            Index           =   3
            Left            =   9600
            TabIndex        =   194
            Top             =   120
            Width           =   1230
            _ExtentX        =   2170
            _ExtentY        =   714
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
            Height          =   405
            Index           =   4
            Left            =   8130
            TabIndex        =   195
            Top             =   120
            Width           =   1230
            _ExtentX        =   2170
            _ExtentY        =   714
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
            Height          =   405
            Index           =   5
            Left            =   6840
            TabIndex        =   196
            Top             =   120
            Width           =   1020
            _ExtentX        =   1799
            _ExtentY        =   714
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
            Height          =   405
            Index           =   7
            Left            =   5325
            TabIndex        =   197
            Top             =   30
            Visible         =   0   'False
            Width           =   1365
            _ExtentX        =   2408
            _ExtentY        =   714
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
            Height          =   405
            Left            =   3960
            TabIndex        =   198
            Top             =   120
            Width           =   1230
            _ExtentX        =   2170
            _ExtentY        =   714
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
         Begin ImpulseButton.ISButton Cmd 
            Height          =   390
            Index           =   10
            Left            =   0
            TabIndex        =   199
            Top             =   -735
            Width           =   1365
            _ExtentX        =   2408
            _ExtentY        =   688
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
         Begin ImpulseButton.ISButton Cmd 
            Height          =   405
            Index           =   6
            Left            =   0
            TabIndex        =   200
            Top             =   120
            Width           =   1230
            _ExtentX        =   2170
            _ExtentY        =   714
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
         Begin ImpulseButton.ISButton ISButton2 
            Height          =   405
            Left            =   960
            TabIndex        =   201
            Top             =   120
            Width           =   1230
            _ExtentX        =   2170
            _ExtentY        =   714
            ButtonStyle     =   1
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
            ButtonImage     =   "FrmPO10.frx":038A
            ColorButton     =   14871017
            ColorHighlight  =   16777215
            ColorHoverText  =   16711680
            ColorShadow     =   4210752
            ColorOutline    =   0
            DrawFocusRectangle=   0   'False
            ColorToggledHoverText=   16711680
            ColorTextShadow =   4210752
         End
         Begin ImpulseButton.ISButton cmdCopyToNew 
            Height          =   330
            Left            =   2535
            TabIndex        =   233
            Top             =   165
            Width           =   1005
            _ExtentX        =   1773
            _ExtentY        =   582
            ButtonStyle     =   1
            ButtonPositionImage=   1
            Caption         =   "‰”ŒÂ „„«À·Â"
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
         Height          =   465
         Index           =   3
         Left            =   15
         TabIndex        =   23
         TabStop         =   0   'False
         Top             =   8775
         Width           =   17445
         _cx             =   30771
         _cy             =   820
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
            Alignment       =   1  'Right Justify
            BackColor       =   &H00C0FFFF&
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Left            =   15300
            Locked          =   -1  'True
            TabIndex        =   24
            TabStop         =   0   'False
            Top             =   -210
            Visible         =   0   'False
            Width           =   1095
         End
         Begin MSDataListLib.DataCombo DCboUserName 
            Height          =   315
            Left            =   4680
            TabIndex        =   25
            Top             =   120
            Width           =   1290
            _ExtentX        =   2275
            _ExtentY        =   556
            _Version        =   393216
            Enabled         =   0   'False
            Text            =   ""
            RightToLeft     =   -1  'True
         End
         Begin VB.Label LblFinal 
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
            Height          =   465
            Left            =   6600
            TabIndex        =   223
            Top             =   0
            Width           =   1815
         End
         Begin VB.Label Label13 
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
            Height          =   360
            Left            =   6600
            TabIndex        =   222
            Top             =   480
            Visible         =   0   'False
            Width           =   1860
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "«·«Ã„«·Ì"
            Height          =   285
            Index           =   53
            Left            =   8505
            RightToLeft     =   -1  'True
            TabIndex        =   221
            Top             =   75
            Width           =   585
         End
         Begin VB.Label LblFinal2 
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
            Left            =   6600
            TabIndex        =   220
            Top             =   0
            Width           =   1815
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "«·ﬁÌ„… «·„÷«›…"
            Height          =   285
            Index           =   52
            Left            =   10320
            RightToLeft     =   -1  'True
            TabIndex        =   219
            Top             =   75
            Width           =   945
         End
         Begin VB.Label LblVat 
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
            Height          =   465
            Left            =   9300
            TabIndex        =   218
            Top             =   0
            Width           =   975
         End
         Begin VB.Label LblTotalView 
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
            Height          =   465
            Left            =   11340
            TabIndex        =   140
            Top             =   0
            Width           =   1170
         End
         Begin VB.Label lbl 
            Alignment       =   2  'Center
            BackColor       =   &H00E2E9E9&
            Caption         =   "«·’«›Ì"
            Height          =   345
            Index           =   49
            Left            =   12390
            TabIndex        =   142
            Top             =   120
            Width           =   915
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
            Height          =   405
            Left            =   11400
            TabIndex        =   141
            Top             =   30
            Visible         =   0   'False
            Width           =   1035
         End
         Begin VB.Label LblDiscountsTotalView 
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
            Height          =   465
            Left            =   13710
            TabIndex        =   137
            Top             =   0
            Width           =   1155
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
            Height          =   405
            Left            =   13365
            TabIndex        =   139
            Top             =   0
            Visible         =   0   'False
            Width           =   900
         End
         Begin VB.Label lbl 
            Alignment       =   2  'Center
            BackColor       =   &H00E2E9E9&
            Caption         =   "Œ’Ê„« "
            Height          =   345
            Index           =   50
            Left            =   14520
            TabIndex        =   138
            Top             =   120
            Width           =   675
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "«Ã„«·Ì «·ﬂ„ÌÂ"
            Height          =   330
            Index           =   63
            Left            =   17325
            TabIndex        =   85
            Top             =   135
            Visible         =   0   'False
            Width           =   930
         End
         Begin VB.Label LblTotalQty 
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
            Height          =   390
            Left            =   17370
            TabIndex        =   84
            Top             =   0
            Visible         =   0   'False
            Width           =   2415
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
            Height          =   465
            Left            =   15255
            TabIndex        =   83
            Top             =   0
            Width           =   1275
         End
         Begin VB.Label lbl 
            Alignment       =   2  'Center
            BackColor       =   &H00E2E9E9&
            Caption         =   "«·≈Ã„«·Ï"
            Height          =   465
            Index           =   25
            Left            =   16290
            TabIndex        =   82
            Top             =   120
            Width           =   1275
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "≈Ã„«·Ì «·ÿ·»"
            Height          =   285
            Index           =   3
            Left            =   16560
            TabIndex        =   31
            Top             =   75
            Visible         =   0   'False
            Width           =   1695
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "«·”Ã· «·Õ«·Ì:"
            Height          =   345
            Index           =   0
            Left            =   3360
            TabIndex        =   30
            Top             =   120
            Width           =   1050
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "⁄œœ «·”Ã·« :"
            Height          =   345
            Index           =   2
            Left            =   1050
            TabIndex        =   29
            Top             =   120
            Width           =   930
         End
         Begin VB.Label XPTxtCurrent 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Height          =   345
            Left            =   2220
            TabIndex        =   28
            Top             =   120
            Width           =   1050
         End
         Begin VB.Label XPTxtCount 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Height          =   345
            Left            =   120
            TabIndex        =   27
            Top             =   120
            Width           =   810
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Õ—— »Ê«”ÿ… : "
            Height          =   345
            Index           =   1
            Left            =   5730
            TabIndex        =   26
            Top             =   120
            Width           =   1290
         End
      End
      Begin C1SizerLibCtl.C1Elastic Ele 
         Height          =   3105
         Index           =   0
         Left            =   0
         TabIndex        =   32
         TabStop         =   0   'False
         Top             =   480
         Width           =   17400
         _cx             =   30692
         _cy             =   5477
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
         Begin VB.TextBox txtMrNo 
            Height          =   345
            Left            =   1560
            TabIndex        =   234
            Top             =   1170
            Width           =   1305
         End
         Begin VB.CommandButton cmdInsertItems 
            Caption         =   "«·«’‰«›"
            Height          =   375
            Left            =   9150
            TabIndex        =   231
            Top             =   900
            Width           =   735
         End
         Begin VB.CheckBox chkTaxExempt 
            Alignment       =   1  'Right Justify
            Caption         =   "„⁄›Ì"
            Height          =   315
            Left            =   8400
            TabIndex        =   230
            Top             =   2760
            Width           =   975
         End
         Begin VB.TextBox txtContainerNo 
            BackColor       =   &H0000FFFF&
            Height          =   345
            Left            =   9900
            TabIndex        =   5
            Top             =   510
            Width           =   1725
         End
         Begin VB.ComboBox CBoBasedON 
            BackColor       =   &H0000FFFF&
            Height          =   315
            ItemData        =   "FrmPO10.frx":6BEC
            Left            =   11280
            List            =   "FrmPO10.frx":6BEE
            RightToLeft     =   -1  'True
            Style           =   2  'Dropdown List
            TabIndex        =   186
            Top             =   165
            Width           =   1305
         End
         Begin VB.ComboBox Combo1 
            Height          =   315
            ItemData        =   "FrmPO10.frx":6BF0
            Left            =   840
            List            =   "FrmPO10.frx":6BF2
            RightToLeft     =   -1  'True
            Style           =   2  'Dropdown List
            TabIndex        =   184
            Top             =   1920
            Width           =   6555
         End
         Begin VB.TextBox TxtNoteSerial 
            Alignment       =   1  'Right Justify
            Height          =   315
            Left            =   240
            TabIndex        =   180
            Top             =   840
            Width           =   2100
         End
         Begin VB.OptionButton opt 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "√„—"
            Height          =   255
            Index           =   1
            Left            =   10815
            TabIndex        =   179
            Top             =   -240
            Visible         =   0   'False
            Width           =   975
         End
         Begin VB.OptionButton opt 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "ÿ·»"
            Height          =   255
            Index           =   0
            Left            =   11655
            TabIndex        =   178
            Top             =   -240
            Visible         =   0   'False
            Width           =   1170
         End
         Begin VB.TextBox TxtPO6 
            Alignment       =   1  'Right Justify
            BackColor       =   &H0000FFFF&
            Height          =   315
            Left            =   9660
            TabIndex        =   1
            Top             =   150
            Width           =   1620
         End
         Begin VB.TextBox TxtPayment 
            Alignment       =   1  'Right Justify
            Height          =   315
            Left            =   960
            TabIndex        =   169
            Top             =   765
            Visible         =   0   'False
            Width           =   1155
         End
         Begin VB.TextBox TxtModeSupply 
            Alignment       =   1  'Right Justify
            Height          =   420
            Left            =   9900
            TabIndex        =   162
            Top             =   4800
            Visible         =   0   'False
            Width           =   5835
         End
         Begin VB.TextBox TxtModeRecept 
            Alignment       =   1  'Right Justify
            Height          =   420
            Left            =   360
            TabIndex        =   160
            Top             =   3240
            Visible         =   0   'False
            Width           =   7035
         End
         Begin VB.TextBox TxtCashCustomerName 
            Alignment       =   1  'Right Justify
            Height          =   420
            Left            =   3870
            TabIndex        =   155
            Top             =   1035
            Width           =   3525
         End
         Begin VB.TextBox XPTxtDiscountVal 
            Alignment       =   1  'Right Justify
            Height          =   315
            Left            =   10335
            TabIndex        =   151
            Top             =   2160
            Width           =   1305
         End
         Begin VB.ComboBox XPCboDiscountType 
            Height          =   315
            Left            =   13185
            Style           =   2  'Dropdown List
            TabIndex        =   150
            Top             =   2220
            Width           =   2550
         End
         Begin VB.ComboBox CboPayMentType 
            Height          =   315
            Left            =   255
            Style           =   2  'Dropdown List
            TabIndex        =   149
            Top             =   -540
            Visible         =   0   'False
            Width           =   2100
         End
         Begin VB.TextBox TxtPONo 
            Alignment       =   1  'Right Justify
            Height          =   315
            Left            =   9900
            TabIndex        =   145
            Top             =   -315
            Visible         =   0   'False
            Width           =   2970
         End
         Begin VB.TextBox TxtEmployeeID 
            Alignment       =   1  'Right Justify
            Height          =   315
            Left            =   6435
            TabIndex        =   135
            Top             =   660
            Width           =   960
         End
         Begin VB.ComboBox CboType 
            Height          =   315
            Left            =   270
            Style           =   2  'Dropdown List
            TabIndex        =   99
            Top             =   3360
            Visible         =   0   'False
            Width           =   2085
         End
         Begin VB.TextBox oldtxtNoteSerial1 
            Alignment       =   1  'Right Justify
            Height          =   390
            Left            =   0
            TabIndex        =   97
            Top             =   4605
            Visible         =   0   'False
            Width           =   1665
         End
         Begin VB.TextBox TxtNoteSerial1 
            Alignment       =   1  'Right Justify
            BackColor       =   &H0000FFFF&
            Height          =   315
            Left            =   13995
            TabIndex        =   0
            Top             =   165
            Width           =   1740
         End
         Begin VB.TextBox TxtBillComment 
            Alignment       =   1  'Right Justify
            Height          =   690
            Left            =   855
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   94
            Top             =   2325
            Width           =   6540
         End
         Begin VB.TextBox TxtStoreID 
            Alignment       =   1  'Right Justify
            BackColor       =   &H0000FFFF&
            Height          =   315
            Left            =   14700
            TabIndex        =   93
            Top             =   1365
            Width           =   1035
         End
         Begin VB.TextBox TxtSearchCode 
            Alignment       =   1  'Right Justify
            BackColor       =   &H0000FFFF&
            Height          =   315
            Left            =   14700
            TabIndex        =   92
            Top             =   915
            Width           =   1035
         End
         Begin VB.TextBox Txt_order_no 
            Alignment       =   1  'Right Justify
            Height          =   660
            Left            =   2370
            TabIndex        =   81
            Top             =   3315
            Visible         =   0   'False
            Width           =   1320
         End
         Begin VB.Frame Frame3 
            Caption         =   "»Ì«‰«  «·«⁄ „«œ"
            Height          =   840
            Left            =   -2070
            TabIndex        =   67
            Top             =   -3540
            Visible         =   0   'False
            Width           =   5055
            Begin VB.TextBox TxtLcNo 
               Alignment       =   1  'Right Justify
               Height          =   285
               Left            =   600
               TabIndex        =   68
               Top             =   240
               Width           =   2175
            End
            Begin MSComCtl2.DTPicker DTPicker2 
               Height          =   315
               Left            =   4080
               TabIndex        =   69
               Top             =   600
               Width           =   2100
               _ExtentX        =   3704
               _ExtentY        =   556
               _Version        =   393216
               Format          =   107479041
               CurrentDate     =   38784
            End
            Begin MSComCtl2.DTPicker DTPicker3 
               Height          =   315
               Left            =   4560
               TabIndex        =   70
               Top             =   960
               Width           =   1620
               _ExtentX        =   2858
               _ExtentY        =   556
               _Version        =   393216
               Format          =   107479041
               CurrentDate     =   38784
            End
            Begin MSComCtl2.DTPicker DTPicker4 
               Height          =   315
               Left            =   120
               TabIndex        =   71
               Top             =   960
               Width           =   2100
               _ExtentX        =   3704
               _ExtentY        =   556
               _Version        =   393216
               Format          =   107479041
               CurrentDate     =   38784
            End
            Begin MSComCtl2.DTPicker DTPicker5 
               Height          =   315
               Left            =   4560
               TabIndex        =   72
               Top             =   1320
               Width           =   1620
               _ExtentX        =   2858
               _ExtentY        =   556
               _Version        =   393216
               Format          =   107479041
               CurrentDate     =   38784
            End
            Begin MSComCtl2.DTPicker DTPicker6 
               Height          =   315
               Left            =   120
               TabIndex        =   73
               Top             =   1320
               Width           =   2100
               _ExtentX        =   3704
               _ExtentY        =   556
               _Version        =   393216
               Format          =   107479041
               CurrentDate     =   38784
            End
            Begin ImpulseButton.ISButton ISButton1 
               Height          =   285
               Left            =   120
               TabIndex        =   86
               Top             =   240
               Width           =   435
               _ExtentX        =   767
               _ExtentY        =   503
               ButtonStyle     =   1
               ButtonPositionImage=   1
               Caption         =   "⁄—÷"
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
            Begin VB.Label Label9 
               Alignment       =   1  'Right Justify
               Caption         =   "„·«ÕŸ« "
               Height          =   375
               Left            =   2400
               TabIndex        =   80
               Top             =   1080
               Width           =   975
            End
            Begin VB.Label Label8 
               Alignment       =   1  'Right Justify
               Caption         =   " «—ÌŒ «·Ê’Ê· «·„ Êﬁ⁄"
               Height          =   255
               Left            =   2280
               TabIndex        =   79
               Top             =   1440
               Width           =   1575
            End
            Begin VB.Label Label7 
               Alignment       =   1  'Right Justify
               Caption         =   " «—ÌŒ «· √ŒÌ—"
               Height          =   255
               Left            =   6480
               TabIndex        =   78
               Top             =   1440
               Width           =   1215
            End
            Begin VB.Label Label6 
               Alignment       =   1  'Right Justify
               Caption         =   "«· «—ÌŒ «·›⁄·Ì"
               Height          =   375
               Left            =   2640
               TabIndex        =   77
               Top             =   1080
               Width           =   1215
            End
            Begin VB.Label Label5 
               Alignment       =   1  'Right Justify
               Caption         =   "«· «—ÌŒ «·„ Êﬁ⁄"
               Height          =   375
               Left            =   6480
               TabIndex        =   76
               Top             =   1080
               Width           =   1215
            End
            Begin VB.Label Label3 
               Alignment       =   1  'Right Justify
               Caption         =   "«· «—ÌŒ"
               Height          =   255
               Left            =   6360
               TabIndex        =   75
               Top             =   600
               Width           =   1335
            End
            Begin VB.Label Label1 
               Alignment       =   1  'Right Justify
               Caption         =   "—ﬁ„ «·«⁄ „«œ"
               Height          =   255
               Index           =   0
               Left            =   2640
               TabIndex        =   74
               Top             =   240
               Width           =   1095
            End
         End
         Begin VB.Frame Frame2 
            Height          =   2700
            Left            =   2670
            TabIndex        =   54
            Top             =   3660
            Visible         =   0   'False
            Width           =   7725
            Begin VB.TextBox Text7 
               Alignment       =   1  'Right Justify
               Height          =   285
               Left            =   240
               TabIndex        =   57
               Top             =   600
               Width           =   3855
            End
            Begin VB.TextBox Text3 
               Alignment       =   1  'Right Justify
               Height          =   285
               Left            =   2640
               TabIndex        =   56
               Top             =   1320
               Width           =   1455
            End
            Begin VB.TextBox Text2 
               Alignment       =   1  'Right Justify
               Height          =   285
               Left            =   240
               TabIndex        =   55
               Top             =   960
               Width           =   1335
            End
            Begin MSComCtl2.DTPicker DTPicker1 
               Height          =   315
               Left            =   240
               TabIndex        =   58
               Top             =   1320
               Width           =   1320
               _ExtentX        =   2328
               _ExtentY        =   556
               _Version        =   393216
               Format          =   107544577
               CurrentDate     =   38784
            End
            Begin MSDataListLib.DataCombo DataCombo9 
               Height          =   315
               Left            =   1920
               TabIndex        =   59
               Top             =   240
               Width           =   2145
               _ExtentX        =   3784
               _ExtentY        =   556
               _Version        =   393216
               ListField       =   "6"
               Text            =   ""
               RightToLeft     =   -1  'True
            End
            Begin MSDataListLib.DataCombo DataCombo11 
               Height          =   315
               Left            =   2640
               TabIndex        =   60
               Top             =   960
               Width           =   1425
               _ExtentX        =   2514
               _ExtentY        =   556
               _Version        =   393216
               ListField       =   "6"
               Text            =   ""
               RightToLeft     =   -1  'True
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               Caption         =   " «·«‰ Â«¡"
               Height          =   285
               Index           =   24
               Left            =   1680
               TabIndex        =   66
               Top             =   1320
               Width           =   975
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               Caption         =   "«·ﬁÌ„…"
               Height          =   285
               Index           =   23
               Left            =   1560
               TabIndex        =   65
               Top             =   960
               Width           =   975
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               Caption         =   "—ﬁ„ «·Õ”«»"
               Height          =   285
               Index           =   22
               Left            =   4320
               TabIndex        =   64
               Top             =   1320
               Width           =   1215
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               Caption         =   "«·⁄„·…"
               Height          =   285
               Index           =   21
               Left            =   4320
               TabIndex        =   63
               Top             =   960
               Width           =   1215
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               Caption         =   "«·»‰ﬂ"
               Height          =   285
               Index           =   20
               Left            =   4320
               TabIndex        =   62
               Top             =   600
               Width           =   1215
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               Caption         =   "‰Ê⁄ «·«„—"
               Height          =   285
               Index           =   19
               Left            =   4320
               TabIndex        =   61
               Top             =   240
               Width           =   1215
            End
         End
         Begin VB.Frame Frame1 
            Height          =   2520
            Left            =   -2595
            TabIndex        =   43
            Top             =   -5685
            Visible         =   0   'False
            Width           =   8805
            Begin VB.CheckBox chkshipped 
               Alignment       =   1  'Right Justify
               Caption         =   " „ «·‘Õ‰"
               Height          =   195
               Left            =   120
               TabIndex        =   95
               Top             =   1320
               Width           =   1815
            End
            Begin VB.TextBox Text1 
               Alignment       =   1  'Right Justify
               Height          =   285
               Left            =   120
               TabIndex        =   44
               Top             =   600
               Width           =   1935
            End
            Begin MSDataListLib.DataCombo DataCombo4 
               Height          =   315
               Left            =   3120
               TabIndex        =   45
               Top             =   960
               Width           =   2145
               _ExtentX        =   3784
               _ExtentY        =   556
               _Version        =   393216
               ListField       =   "6"
               Text            =   ""
               RightToLeft     =   -1  'True
            End
            Begin MSDataListLib.DataCombo DataCombo5 
               Height          =   315
               Left            =   3120
               TabIndex        =   46
               Top             =   1320
               Width           =   2145
               _ExtentX        =   3784
               _ExtentY        =   556
               _Version        =   393216
               ListField       =   "6"
               Text            =   ""
               RightToLeft     =   -1  'True
            End
            Begin MSDataListLib.DataCombo DataCombo6 
               Height          =   315
               Left            =   120
               TabIndex        =   47
               Top             =   960
               Width           =   1905
               _ExtentX        =   3360
               _ExtentY        =   556
               _Version        =   393216
               ListField       =   "6"
               Text            =   ""
               RightToLeft     =   -1  'True
            End
            Begin MSDataListLib.DataCombo DataCombo7 
               Height          =   315
               Left            =   3120
               TabIndex        =   48
               Top             =   240
               Width           =   2145
               _ExtentX        =   3784
               _ExtentY        =   556
               _Version        =   393216
               ListField       =   "6"
               Text            =   ""
               RightToLeft     =   -1  'True
            End
            Begin MSDataListLib.DataCombo DataCombo1 
               Height          =   315
               Left            =   3120
               TabIndex        =   88
               Top             =   600
               Width           =   2130
               _ExtentX        =   3757
               _ExtentY        =   556
               _Version        =   393216
               ListField       =   "6"
               Text            =   ""
               RightToLeft     =   -1  'True
            End
            Begin MSDataListLib.DataCombo DataCombo2 
               Height          =   315
               Left            =   120
               TabIndex        =   90
               Top             =   240
               Width           =   1890
               _ExtentX        =   3334
               _ExtentY        =   556
               _Version        =   393216
               ListField       =   "6"
               Text            =   ""
               RightToLeft     =   -1  'True
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               BackStyle       =   0  'Transparent
               Caption         =   "«·„‘—Ê⁄"
               Height          =   270
               Index           =   11
               Left            =   2130
               TabIndex        =   91
               Top             =   240
               Width           =   855
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               BackStyle       =   0  'Transparent
               Caption         =   "„—ﬂ“ «· ﬂ·›…"
               Height          =   285
               Index           =   10
               Left            =   5370
               TabIndex        =   89
               Top             =   600
               Width           =   1095
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               Caption         =   "«·ﬁÌ„…"
               Height          =   285
               Index           =   17
               Left            =   2040
               TabIndex        =   53
               Top             =   600
               Width           =   975
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               Caption         =   "«· ’‰Ì›"
               Height          =   285
               Index           =   16
               Left            =   5400
               TabIndex        =   52
               Top             =   240
               Width           =   1095
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               Caption         =   "ÿ—Ìﬁ… «·œ›⁄"
               Height          =   285
               Index           =   15
               Left            =   2040
               TabIndex        =   51
               Top             =   960
               Width           =   975
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               Caption         =   "ÿ—Ìﬁ… «·‘Õ‰"
               Height          =   285
               Index           =   14
               Left            =   5280
               TabIndex        =   50
               Top             =   1320
               Width           =   1215
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               Caption         =   "«·»·œ"
               Height          =   285
               Index           =   13
               Left            =   5280
               TabIndex        =   49
               Top             =   960
               Width           =   1215
            End
         End
         Begin VB.ComboBox CboPriceType 
            Enabled         =   0   'False
            Height          =   315
            Left            =   5040
            Style           =   2  'Dropdown List
            TabIndex        =   17
            Top             =   -330
            Visible         =   0   'False
            Width           =   3105
         End
         Begin VB.TextBox TxtTransSerial 
            Alignment       =   1  'Right Justify
            Height          =   450
            Left            =   13995
            TabIndex        =   16
            Top             =   -330
            Visible         =   0   'False
            Width           =   2190
         End
         Begin VB.TextBox TxtModFlg 
            Alignment       =   1  'Right Justify
            Height          =   450
            Left            =   3870
            TabIndex        =   34
            Top             =   -525
            Visible         =   0   'False
            Width           =   870
         End
         Begin VB.TextBox TxtFillData 
            Alignment       =   1  'Right Justify
            Height          =   375
            Left            =   2610
            TabIndex        =   33
            Top             =   -1650
            Visible         =   0   'False
            Width           =   1215
         End
         Begin VB.TextBox XPTxtBillID 
            Alignment       =   1  'Right Justify
            Height          =   420
            Left            =   30
            Locked          =   -1  'True
            TabIndex        =   21
            Top             =   -450
            Visible         =   0   'False
            Width           =   2565
         End
         Begin MSDataListLib.DataCombo DBCboClientName 
            Height          =   315
            Left            =   9900
            TabIndex        =   6
            Top             =   915
            Width           =   4875
            _ExtentX        =   8599
            _ExtentY        =   556
            _Version        =   393216
            BackColor       =   65535
            ListField       =   "6"
            Text            =   ""
            RightToLeft     =   -1  'True
         End
         Begin MSDataListLib.DataCombo DCboStoreName 
            Height          =   315
            Left            =   9900
            TabIndex        =   7
            Top             =   1365
            Width           =   4755
            _ExtentX        =   8387
            _ExtentY        =   556
            _Version        =   393216
            BackColor       =   65535
            ListField       =   "7"
            Text            =   ""
            RightToLeft     =   -1  'True
         End
         Begin MSComCtl2.DTPicker XPDtbBill 
            Height          =   315
            Left            =   13995
            TabIndex        =   4
            Top             =   540
            Width           =   1740
            _ExtentX        =   3069
            _ExtentY        =   556
            _Version        =   393216
            CalendarBackColor=   65535
            Format          =   251265025
            CurrentDate     =   38784
         End
         Begin ImpulseButton.ISButton XPBtnNewClients 
            Height          =   615
            Left            =   9420
            TabIndex        =   35
            TabStop         =   0   'False
            Top             =   1740
            Visible         =   0   'False
            Width           =   555
            _ExtentX        =   979
            _ExtentY        =   1085
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
            ButtonImage     =   "FrmPO10.frx":6BF4
            ColorButton     =   14871017
            ColorHighlight  =   16777215
            ColorShadow     =   -2147483631
            ColorOutline    =   -2147483631
            DrawFocusRectangle=   0   'False
         End
         Begin ImpulseButton.ISButton CmdTemplate 
            Height          =   855
            Left            =   2055
            TabIndex        =   36
            Top             =   3600
            Visible         =   0   'False
            Width           =   2295
            _ExtentX        =   4048
            _ExtentY        =   1508
            ButtonStyle     =   1
            ButtonPositionImage=   1
            Caption         =   "≈œ—«Ã ⁄—÷ Ã«Â“"
            BackColor       =   12632256
            ForeColor       =   16711680
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
         Begin C1SizerLibCtl.C1Elastic Ele 
            Height          =   870
            Index           =   4
            Left            =   19635
            TabIndex        =   18
            TabStop         =   0   'False
            Top             =   2835
            Width           =   4965
            _cx             =   8758
            _cy             =   1535
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
            Begin VB.CheckBox XPChkTAX 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "÷—»Ì»… «·„»Ì⁄« "
               Height          =   330
               Left            =   1860
               TabIndex        =   19
               Top             =   210
               Width           =   1815
            End
            Begin VB.TextBox XPTxtTaxValue 
               Alignment       =   1  'Right Justify
               Height          =   390
               Left            =   30
               TabIndex        =   20
               Top             =   150
               Width           =   915
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "«·ﬁÌ„…"
               Enabled         =   0   'False
               Height          =   240
               Index           =   4
               Left            =   990
               TabIndex        =   41
               Top             =   285
               Width           =   720
            End
         End
         Begin MSDataListLib.DataCombo Dccurrency 
            Height          =   315
            Left            =   240
            TabIndex        =   3
            Top             =   165
            Width           =   2100
            _ExtentX        =   3704
            _ExtentY        =   556
            _Version        =   393216
            BackColor       =   65535
            ListField       =   "6"
            Text            =   ""
            RightToLeft     =   -1  'True
         End
         Begin MSDataListLib.DataCombo dcBranch 
            Height          =   315
            Left            =   3915
            TabIndex        =   2
            Top             =   180
            Width           =   3435
            _ExtentX        =   6059
            _ExtentY        =   556
            _Version        =   393216
            BackColor       =   65535
            Text            =   ""
            RightToLeft     =   -1  'True
         End
         Begin MSDataListLib.DataCombo DcboEmp 
            Height          =   315
            Left            =   3885
            TabIndex        =   136
            Top             =   660
            Width           =   2505
            _ExtentX        =   4419
            _ExtentY        =   556
            _Version        =   393216
            ListField       =   "7"
            Text            =   ""
            RightToLeft     =   -1  'True
         End
         Begin ImpulseButton.ISButton CmdConvert 
            Height          =   615
            Left            =   0
            TabIndex        =   158
            TabStop         =   0   'False
            Top             =   0
            Visible         =   0   'False
            Width           =   360
            _ExtentX        =   635
            _ExtentY        =   1085
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
            ButtonImage     =   "FrmPO10.frx":6F8E
            ColorButton     =   14871017
            ColorHighlight  =   16777215
            ColorShadow     =   -2147483631
            ColorOutline    =   -2147483631
            DrawFocusRectangle=   0   'False
         End
         Begin MSDataListLib.DataCombo DcbDetpartment 
            Height          =   315
            Left            =   9900
            TabIndex        =   8
            Top             =   1800
            Width           =   5835
            _ExtentX        =   10292
            _ExtentY        =   556
            _Version        =   393216
            BackColor       =   65535
            ListField       =   "7"
            Text            =   ""
            RightToLeft     =   -1  'True
         End
         Begin MSDataListLib.DataCombo DcbPayment 
            Height          =   315
            Left            =   240
            TabIndex        =   170
            Top             =   480
            Width           =   2100
            _ExtentX        =   3704
            _ExtentY        =   556
            _Version        =   393216
            ListField       =   "7"
            Text            =   ""
            RightToLeft     =   -1  'True
         End
         Begin MSDataListLib.DataCombo DcbShiping 
            Height          =   315
            Left            =   840
            TabIndex        =   9
            Top             =   1560
            Width           =   6555
            _ExtentX        =   11562
            _ExtentY        =   556
            _Version        =   393216
            BackColor       =   65535
            ListField       =   "7"
            Text            =   ""
            RightToLeft     =   -1  'True
         End
         Begin MSComCtl2.DTPicker SippingDate 
            Height          =   315
            Left            =   13650
            TabIndex        =   172
            Top             =   2640
            Width           =   2100
            _ExtentX        =   3704
            _ExtentY        =   556
            _Version        =   393216
            Format          =   251199489
            CurrentDate     =   38784
         End
         Begin MSComCtl2.DTPicker DeliverDate 
            Height          =   315
            Left            =   9900
            TabIndex        =   174
            Top             =   2670
            Width           =   1905
            _ExtentX        =   3360
            _ExtentY        =   556
            _Version        =   393216
            Format          =   251199489
            CurrentDate     =   38784
         End
         Begin ImpulseButton.ISButton Cmd 
            CausesValidation=   0   'False
            Height          =   375
            Index           =   11
            Left            =   -210
            TabIndex        =   183
            Top             =   1200
            Width           =   1620
            _ExtentX        =   2858
            _ExtentY        =   661
            ButtonStyle     =   1
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
            ColorShadow     =   4210752
            ColorOutline    =   0
            DrawFocusRectangle=   0   'False
            ColorToggledHoverText=   16711680
            ColorTextShadow =   4210752
         End
         Begin VSFlex8Ctl.VSFlexGrid tmpGrd 
            Height          =   3045
            Left            =   -5220
            TabIndex        =   225
            Top             =   4530
            Visible         =   0   'False
            Width           =   6795
            _cx             =   11986
            _cy             =   5371
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
            BackColor       =   8421631
            ForeColor       =   -2147483640
            BackColorFixed  =   -2147483633
            ForeColorFixed  =   -2147483630
            BackColorSel    =   -2147483635
            ForeColorSel    =   -2147483634
            BackColorBkg    =   -2147483636
            BackColorAlternate=   8421631
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
            Cols            =   40
            FixedRows       =   1
            FixedCols       =   1
            RowHeightMin    =   0
            RowHeightMax    =   0
            ColWidthMin     =   0
            ColWidthMax     =   0
            ExtendLastCol   =   0   'False
            FormatString    =   ""
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
         Begin VSFlex8Ctl.VSFlexGrid grdExcel 
            Height          =   2040
            Index           =   1
            Left            =   -9690
            TabIndex        =   226
            Top             =   2160
            Visible         =   0   'False
            Width           =   12255
            _cx             =   21616
            _cy             =   3598
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
            GridLines       =   3
            GridLinesFixed  =   2
            GridLineWidth   =   5
            Rows            =   2
            Cols            =   16
            FixedRows       =   1
            FixedCols       =   1
            RowHeightMin    =   300
            RowHeightMax    =   0
            ColWidthMin     =   0
            ColWidthMax     =   0
            ExtendLastCol   =   -1  'True
            FormatString    =   $"FrmPO10.frx":7328
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
            ExplorerBar     =   3
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
         Begin VB.Label Label14 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "MR NO "
            ForeColor       =   &H00000000&
            Height          =   270
            Index           =   1
            Left            =   2880
            TabIndex        =   235
            Top             =   1260
            Width           =   1020
         End
         Begin VB.Label Label14 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "Q.Ref"
            ForeColor       =   &H00000000&
            Height          =   270
            Index           =   0
            Left            =   11610
            TabIndex        =   224
            Top             =   600
            Width           =   1590
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "„Êﬁ› «·ÿ·»"
            Height          =   240
            Index           =   47
            Left            =   7590
            TabIndex        =   185
            Top             =   1920
            Width           =   1545
         End
         Begin VB.Label lbl 
            Alignment       =   2  'Center
            BackColor       =   &H00E2E9E9&
            Caption         =   "—ﬁ„ «·ﬁÌœ"
            Height          =   285
            Index           =   45
            Left            =   2355
            TabIndex        =   181
            Top             =   840
            Width           =   1410
         End
         Begin VB.Label lbl 
            Alignment       =   2  'Center
            BackColor       =   &H00E2E9E9&
            Caption         =   "»‰«¡ ⁄·Ï  "
            Height          =   285
            Index           =   44
            Left            =   12690
            TabIndex        =   176
            Top             =   120
            Width           =   1200
         End
         Begin VB.Label lbl 
            Alignment       =   2  'Center
            BackColor       =   &H00E2E9E9&
            Caption         =   " «—ÌŒ «· Ê’Ì·"
            Height          =   270
            Index           =   43
            Left            =   11775
            TabIndex        =   175
            Top             =   2640
            Width           =   1800
         End
         Begin VB.Label lbl 
            Alignment       =   2  'Center
            BackColor       =   &H00E2E9E9&
            Caption         =   " «—ÌŒ «·‘Õ‰"
            Height          =   270
            Index           =   42
            Left            =   15765
            TabIndex        =   173
            Top             =   2640
            Width           =   1560
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "ÿ—Ìﬁ… «·‘Õ‰ Ê«· Ê—Ìœ"
            Height          =   375
            Index           =   41
            Left            =   7425
            TabIndex        =   171
            Top             =   1560
            Width           =   1725
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "«ÿ—Ìﬁ… ««” ·«„ «·„Ê«œ"
            Height          =   375
            Index           =   38
            Left            =   21750
            TabIndex        =   167
            Top             =   1440
            Visible         =   0   'False
            Width           =   1560
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "„·«ÕŸ« "
            Height          =   255
            Index           =   28
            Left            =   7350
            TabIndex        =   166
            Top             =   2400
            Width           =   1740
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "«”„ «·„‰œÊ»"
            Height          =   390
            Index           =   32
            Left            =   7740
            TabIndex        =   165
            Top             =   600
            Width           =   1395
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "«”„ «·„Ê—œ «·‰ﬁœÌ"
            Height          =   240
            Index           =   36
            Left            =   7440
            TabIndex        =   164
            Top             =   1035
            Width           =   1695
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "ÿ—Ìﬁ… «· Ê—Ìœ"
            Height          =   375
            Index           =   39
            Left            =   17160
            TabIndex        =   163
            Top             =   2280
            Visible         =   0   'False
            Width           =   1440
         End
         Begin VB.Label lbl 
            Alignment       =   2  'Center
            BackColor       =   &H00E2E9E9&
            Caption         =   "«·«œ«—… «·ÿ«·»…"
            Height          =   375
            Index           =   37
            Left            =   15765
            TabIndex        =   161
            Top             =   1800
            Width           =   1560
         End
         Begin VB.Label lbl 
            Alignment       =   2  'Center
            BackColor       =   &H00E2E9E9&
            Caption         =   "ﬁÌ„…"
            Height          =   330
            Index           =   35
            Left            =   11775
            TabIndex        =   154
            Top             =   2160
            Width           =   1275
         End
         Begin VB.Label lbl 
            Alignment       =   2  'Center
            BackColor       =   &H00E2E9E9&
            Caption         =   "‰Ê⁄ «·Œ’„"
            Height          =   330
            Index           =   34
            Left            =   15660
            TabIndex        =   153
            Top             =   2220
            Width           =   1755
         End
         Begin VB.Label lbl 
            Alignment       =   2  'Center
            BackColor       =   &H00E2E9E9&
            Caption         =   "%"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Index           =   55
            Left            =   9660
            TabIndex        =   152
            Top             =   2280
            Width           =   540
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "»‰«¡ ⁄·Ï"
            Height          =   330
            Index           =   33
            Left            =   12720
            TabIndex        =   146
            Top             =   -315
            Visible         =   0   'False
            Width           =   1035
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "”Ì«”… «·ÿ·»Ì…"
            Height          =   330
            Index           =   18
            Left            =   2370
            TabIndex        =   98
            Top             =   3360
            Visible         =   0   'False
            Width           =   1530
         End
         Begin VB.Label Label2 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "«·›—⁄"
            ForeColor       =   &H00000000&
            Height          =   270
            Left            =   7275
            TabIndex        =   96
            Top             =   120
            Width           =   1890
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "«·⁄„·Â"
            Height          =   390
            Index           =   12
            Left            =   2310
            TabIndex        =   87
            Top             =   165
            Width           =   1110
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "ÿ—Ìﬁ… «·œ›⁄"
            Height          =   330
            Index           =   9
            Left            =   2355
            TabIndex        =   42
            Top             =   480
            Width           =   1155
         End
         Begin VB.Label lbl 
            Alignment       =   2  'Center
            BackColor       =   &H00E2E9E9&
            Caption         =   "—ﬁ„ «·«„—"
            Height          =   375
            Index           =   5
            Left            =   15885
            TabIndex        =   40
            Top             =   120
            Width           =   1440
         End
         Begin VB.Label lbl 
            Alignment       =   2  'Center
            BackColor       =   &H00E2E9E9&
            Caption         =   " «—ÌŒ «·«„—"
            Height          =   270
            Index           =   6
            Left            =   15885
            TabIndex        =   39
            Top             =   540
            Width           =   1440
         End
         Begin VB.Label lbl 
            Alignment       =   2  'Center
            BackColor       =   &H00E2E9E9&
            Caption         =   "«·„Ê—œ"
            Height          =   390
            Index           =   7
            Left            =   15765
            TabIndex        =   38
            Top             =   915
            Width           =   1560
         End
         Begin VB.Label lbl 
            Alignment       =   2  'Center
            BackColor       =   &H00E2E9E9&
            Caption         =   "«”„ «·„Œ“‰"
            Height          =   375
            Index           =   8
            Left            =   15645
            TabIndex        =   37
            Top             =   1320
            Width           =   1680
         End
      End
      Begin C1SizerLibCtl.C1Tab XPTab301 
         Height          =   5895
         Left            =   0
         TabIndex        =   100
         Top             =   2805
         Width           =   17415
         _cx             =   30718
         _cy             =   10398
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
         Caption         =   "«·√’‰«›|Õ«·Â «·«⁄ „«œ|⁄—Ê÷ «·«”⁄«—|«·ÿ·»«  «·œ«Œ·Ì…|»Ì«‰«  «·ﬁÌ„… «·„÷«›…"
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
         Picture(0)      =   "FrmPO10.frx":759F
         Begin C1SizerLibCtl.C1Elastic C1Elastic1 
            Height          =   5430
            Left            =   18060
            TabIndex        =   127
            TabStop         =   0   'False
            Top             =   45
            Width           =   17325
            _cx             =   30559
            _cy             =   9578
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
            Begin VSFlex8UCtl.VSFlexGrid GRID2 
               Height          =   4230
               Left            =   2520
               TabIndex        =   128
               Tag             =   "1"
               Top             =   960
               Width           =   13230
               _cx             =   23336
               _cy             =   7461
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
               AllowUserResizing=   0
               SelectionMode   =   0
               GridLines       =   1
               GridLinesFixed  =   2
               GridLineWidth   =   1
               Rows            =   3
               Cols            =   8
               FixedRows       =   1
               FixedCols       =   1
               RowHeightMin    =   300
               RowHeightMax    =   0
               ColWidthMin     =   0
               ColWidthMax     =   0
               ExtendLastCol   =   -1  'True
               FormatString    =   $"FrmPO10.frx":7939
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
            Begin VB.Label Label11 
               Alignment       =   2  'Center
               BackColor       =   &H00FFFFC0&
               Caption         =   "„ÿ·Ê» «⁄ „«œ… Õ«·Ì«"
               Height          =   255
               Left            =   9960
               TabIndex        =   143
               Top             =   4560
               Width           =   3375
            End
         End
         Begin C1SizerLibCtl.C1Elastic Ele 
            Height          =   5430
            Index           =   15
            Left            =   45
            TabIndex        =   101
            TabStop         =   0   'False
            Top             =   45
            Width           =   17325
            _cx             =   30559
            _cy             =   9578
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
            GridRows        =   1
            GridCols        =   1
            Frame           =   3
            FrameStyle      =   0
            FrameWidth      =   1
            FrameColor      =   -2147483628
            FrameShadow     =   -2147483632
            FloodStyle      =   1
            _GridInfo       =   $"FrmPO10.frx":7A7C
            AccessibleName  =   ""
            AccessibleDescription=   ""
            AccessibleValue =   ""
            AccessibleRole  =   9
            Begin C1SizerLibCtl.C1Elastic Ele 
               Height          =   5400
               Index           =   16
               Left            =   15
               TabIndex        =   102
               TabStop         =   0   'False
               Top             =   15
               Width           =   17295
               _cx             =   30506
               _cy             =   9525
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
               Begin C1SizerLibCtl.C1Elastic Ele 
                  Height          =   8235
                  Index           =   5
                  Left            =   0
                  TabIndex        =   111
                  TabStop         =   0   'False
                  Top             =   -480
                  Width           =   17370
                  _cx             =   30639
                  _cy             =   14526
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
                  Begin VB.TextBox txtFileName 
                     Alignment       =   2  'Center
                     Height          =   225
                     Left            =   0
                     TabIndex        =   229
                     Top             =   0
                     Width           =   1290
                  End
                  Begin VB.CommandButton cmdLoadFile 
                     Caption         =   " Õ„Ì· «·„·›..."
                     Height          =   264
                     Left            =   12510
                     TabIndex        =   228
                     Top             =   5535
                     Width           =   1440
                  End
                  Begin VB.CommandButton cmdSelectFile 
                     Caption         =   " ÕœÌœ «·„·›..."
                     Height          =   240
                     Left            =   14040
                     RightToLeft     =   -1  'True
                     TabIndex        =   227
                     Top             =   5565
                     Width           =   1275
                  End
                  Begin VB.Frame Frame4 
                     BorderStyle     =   0  'None
                     Height          =   915
                     Left            =   600
                     TabIndex        =   112
                     Top             =   5130
                     Visible         =   0   'False
                     Width           =   2130
                     Begin DBPIXLib.DBPix20 DBPix202 
                        Height          =   855
                        Left            =   480
                        TabIndex        =   113
                        Top             =   -390
                        Width           =   2415
                        _Version        =   131072
                        _ExtentX        =   4260
                        _ExtentY        =   1508
                        _StockProps     =   1
                        _Image          =   "FrmPO10.frx":7AB2
                        ImageResampleWidth=   100
                        ImageResampleHeight=   100
                        ImageResampleMode=   1
                        ImageSaveFormat =   0
                        JPEGQuality     =   75
                        JPEGEncoding    =   0
                        JPEGColorMode   =   0
                        JPEGNoRecompress=   -1  'True
                        JPEGRotateWarning=   0
                        PNGColorDepth   =   0
                        PNGCompression  =   0
                        PNGFilter       =   0
                        PNGInterlace    =   1
                        ImageDitherMethod=   3
                        ImagePaletteMethod=   4
                        ImagePreviewMode=   0   'False
                        ImageKeepMetaData=   -1  'True
                        UseAmbientBackcolor=   -1  'True
                        ViewAsyncDecoding=   -1  'True
                        ViewEnableMouseZoom=   -1  'True
                        ViewInitialZoom =   1
                        ViewHAlign      =   1
                        ViewVAlign      =   1
                        ViewMenuMode    =   0
                     End
                     Begin VB.Label LblPostedPerson 
                        Alignment       =   2  'Center
                        BackStyle       =   0  'Transparent
                        Caption         =   "."
                        Height          =   255
                        Left            =   3600
                        TabIndex        =   116
                        Top             =   240
                        Width           =   1695
                     End
                     Begin VB.Label Label10 
                        Alignment       =   1  'Right Justify
                        BackStyle       =   0  'Transparent
                        Caption         =   "«·‰ÊﬁÌ⁄"
                        Height          =   255
                        Left            =   2640
                        TabIndex        =   115
                        Top             =   240
                        Width           =   855
                     End
                     Begin VB.Label Label4 
                        Alignment       =   1  'Right Justify
                        BackStyle       =   0  'Transparent
                        Caption         =   "Ì⁄ „œ"
                        Height          =   255
                        Left            =   5160
                        TabIndex        =   114
                        Top             =   240
                        Width           =   735
                     End
                  End
                  Begin C1SizerLibCtl.C1Elastic Ele 
                     Height          =   1140
                     Index           =   2
                     Left            =   0
                     TabIndex        =   117
                     TabStop         =   0   'False
                     Top             =   1245
                     Width           =   17220
                     _cx             =   30374
                     _cy             =   2011
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
                     Begin VB.TextBox txtItemCodeSearch 
                        BackColor       =   &H0000FFFF&
                        Height          =   270
                        Left            =   14640
                        TabIndex        =   232
                        Top             =   300
                        Width           =   1890
                     End
                     Begin VB.TextBox TxtSerial 
                        Alignment       =   1  'Right Justify
                        BackColor       =   &H00FFFFFF&
                        Enabled         =   0   'False
                        Height          =   345
                        Left            =   -6675
                        MaxLength       =   20
                        TabIndex        =   177
                        Top             =   0
                        Visible         =   0   'False
                        Width           =   4020
                     End
                     Begin VB.ComboBox CboItemCase 
                        BackColor       =   &H0000FFFF&
                        Height          =   315
                        Left            =   3405
                        Style           =   2  'Dropdown List
                        TabIndex        =   118
                        Top             =   300
                        Width           =   1665
                     End
                     Begin VB.TextBox TxtQuantity 
                        Alignment       =   1  'Right Justify
                        BackColor       =   &H0000FFFF&
                        Height          =   315
                        Left            =   2130
                        MaxLength       =   10
                        TabIndex        =   12
                        Top             =   300
                        Width           =   1245
                     End
                     Begin VB.TextBox TxtPrice 
                        Alignment       =   1  'Right Justify
                        BackColor       =   &H0000FFFF&
                        Height          =   315
                        Left            =   690
                        MaxLength       =   10
                        TabIndex        =   13
                        Top             =   330
                        Width           =   1380
                     End
                     Begin MSDataListLib.DataCombo DCboItemsName 
                        Height          =   315
                        Left            =   5280
                        TabIndex        =   11
                        Top             =   300
                        Width           =   6555
                        _ExtentX        =   11562
                        _ExtentY        =   556
                        _Version        =   393216
                        BackColor       =   65535
                        Text            =   ""
                        RightToLeft     =   -1  'True
                     End
                     Begin MSDataListLib.DataCombo DCboItemsCode 
                        Height          =   315
                        Left            =   11955
                        TabIndex        =   10
                        Top             =   300
                        Width           =   2460
                        _ExtentX        =   4339
                        _ExtentY        =   556
                        _Version        =   393216
                        BackColor       =   65535
                        Text            =   ""
                        RightToLeft     =   -1  'True
                     End
                     Begin ImpulseButton.ISButton CmdAdd 
                        Height          =   480
                        Left            =   -150
                        TabIndex        =   14
                        Top             =   255
                        Width           =   840
                        _ExtentX        =   1482
                        _ExtentY        =   847
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
                        ButtonImage     =   "FrmPO10.frx":7ACA
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
                     Begin ImpulseButton.ISButton ISButton3 
                        Height          =   375
                        Left            =   0
                        TabIndex        =   203
                        Top             =   240
                        Visible         =   0   'False
                        Width           =   840
                        _ExtentX        =   1482
                        _ExtentY        =   661
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
                        ButtonImage     =   "FrmPO10.frx":7E64
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
                     Begin ImpulseButton.ISButton ISButton4 
                        Height          =   375
                        Left            =   150
                        TabIndex        =   204
                        Top             =   240
                        Width           =   540
                        _ExtentX        =   953
                        _ExtentY        =   661
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
                        ButtonImage     =   "FrmPO10.frx":81FE
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
                     Begin XtremeSuiteControls.CheckBox ChAuto 
                        Height          =   255
                        Left            =   840
                        TabIndex        =   205
                        Top             =   720
                        Width           =   1515
                        _Version        =   786432
                        _ExtentX        =   2672
                        _ExtentY        =   450
                        _StockProps     =   79
                        Caption         =   " Õ„Ì· «·„Ê«œ «·Ì«"
                        BackColor       =   14871017
                        UseVisualStyle  =   -1  'True
                        TextAlignment   =   1
                        RightToLeft     =   -1  'True
                     End
                     Begin MSDataListLib.DataCombo dcopr2 
                        Height          =   315
                        Left            =   7275
                        TabIndex        =   206
                        Top             =   720
                        Width           =   4185
                        _ExtentX        =   7382
                        _ExtentY        =   556
                        _Version        =   393216
                        Appearance      =   0
                        Text            =   ""
                        RightToLeft     =   -1  'True
                     End
                     Begin MSDataListLib.DataCombo dcproject1 
                        Height          =   315
                        Left            =   12210
                        TabIndex        =   207
                        Top             =   720
                        Width           =   4410
                        _ExtentX        =   7779
                        _ExtentY        =   556
                        _Version        =   393216
                        Appearance      =   0
                        Text            =   ""
                        RightToLeft     =   -1  'True
                     End
                     Begin MSDataListLib.DataCombo DcbProcess1 
                        Height          =   315
                        Left            =   2505
                        TabIndex        =   208
                        Top             =   720
                        Width           =   4095
                        _ExtentX        =   7223
                        _ExtentY        =   556
                        _Version        =   393216
                        Appearance      =   0
                        Text            =   ""
                        RightToLeft     =   -1  'True
                     End
                     Begin VB.Label lbl 
                        Alignment       =   1  'Right Justify
                        BackColor       =   &H00E2E9E9&
                        Caption         =   "«·⁄„·ÌÂ"
                        Height          =   195
                        Index           =   51
                        Left            =   5760
                        RightToLeft     =   -1  'True
                        TabIndex        =   211
                        Top             =   720
                        Width           =   1440
                     End
                     Begin VB.Label lbl 
                        Alignment       =   1  'Right Justify
                        BackColor       =   &H00E2E9E9&
                        Caption         =   "«·„‘—Ê⁄"
                        Height          =   315
                        Index           =   48
                        Left            =   15930
                        RightToLeft     =   -1  'True
                        TabIndex        =   210
                        Top             =   720
                        Width           =   1290
                     End
                     Begin VB.Label lbl 
                        Alignment       =   1  'Right Justify
                        BackColor       =   &H00E2E9E9&
                        Caption         =   "«·»‰œ"
                        Height          =   195
                        Index           =   46
                        Left            =   10470
                        RightToLeft     =   -1  'True
                        TabIndex        =   209
                        Top             =   720
                        Width           =   1590
                     End
                     Begin VB.Label lbl 
                        Alignment       =   2  'Center
                        BackColor       =   &H00E2E9E9&
                        Caption         =   "ﬂÊœ «·’‰›"
                        Height          =   255
                        Index           =   31
                        Left            =   11880
                        TabIndex        =   123
                        Top             =   -30
                        Width           =   3180
                     End
                     Begin VB.Label lbl 
                        Alignment       =   2  'Center
                        BackColor       =   &H00E2E9E9&
                        Caption         =   "≈”„ «·’‰›"
                        Height          =   255
                        Index           =   30
                        Left            =   8220
                        TabIndex        =   122
                        Top             =   30
                        Width           =   3045
                     End
                     Begin VB.Label lbl 
                        Alignment       =   2  'Center
                        BackColor       =   &H00E2E9E9&
                        Caption         =   "Õ«·… «·’‰›"
                        Height          =   255
                        Index           =   29
                        Left            =   3855
                        TabIndex        =   121
                        Top             =   0
                        Width           =   1215
                     End
                     Begin VB.Label lbl 
                        Alignment       =   2  'Center
                        BackColor       =   &H00E2E9E9&
                        Caption         =   "«·ﬂ„Ì…"
                        Height          =   255
                        Index           =   27
                        Left            =   2355
                        TabIndex        =   120
                        Top             =   0
                        Width           =   930
                     End
                     Begin VB.Label lbl 
                        Alignment       =   2  'Center
                        BackColor       =   &H00E2E9E9&
                        Caption         =   "«·”⁄—"
                        Height          =   255
                        Index           =   26
                        Left            =   765
                        TabIndex        =   119
                        Top             =   0
                        Width           =   1035
                     End
                  End
                  Begin VSFlex8UCtl.VSFlexGrid FG 
                     Height          =   2910
                     Left            =   -30
                     TabIndex        =   124
                     Top             =   2370
                     Width           =   17070
                     _cx             =   30110
                     _cy             =   5133
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
                     Cols            =   21
                     FixedRows       =   1
                     FixedCols       =   1
                     RowHeightMin    =   300
                     RowHeightMax    =   0
                     ColWidthMin     =   0
                     ColWidthMax     =   0
                     ExtendLastCol   =   -1  'True
                     FormatString    =   $"FrmPO10.frx":8598
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
                     Height          =   630
                     Left            =   600
                     TabIndex        =   125
                     Top             =   5220
                     Width           =   3945
                     _ExtentX        =   6959
                     _ExtentY        =   1111
                     ButtonWidth     =   609
                     ButtonHeight    =   1005
                     Appearance      =   1
                     _Version        =   393216
                  End
                  Begin ImpulseButton.ISButton Accredit 
                     Height          =   540
                     Left            =   4620
                     TabIndex        =   144
                     Top             =   5445
                     Width           =   2355
                     _ExtentX        =   4154
                     _ExtentY        =   953
                     ButtonStyle     =   1
                     ButtonPositionImage=   1
                     Caption         =   "«—”«· ··«⁄ „«œ"
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
                  Begin VB.Label lbl 
                     Alignment       =   1  'Right Justify
                     BackColor       =   &H00E2E9E9&
                     Caption         =   "„·«ÕŸ« "
                     Height          =   390
                     Index           =   40
                     Left            =   7890
                     TabIndex        =   168
                     Top             =   6015
                     Width           =   1140
                  End
                  Begin VB.Label LblItemsCount 
                     Alignment       =   2  'Center
                     BackColor       =   &H00404040&
                     ForeColor       =   &H0000FFFF&
                     Height          =   300
                     Left            =   0
                     TabIndex        =   126
                     Top             =   5355
                     Width           =   600
                  End
               End
               Begin VB.Label Label12 
                  Alignment       =   1  'Right Justify
                  Caption         =   "Label12"
                  Height          =   885
                  Left            =   3795
                  TabIndex        =   110
                  Top             =   240
                  Width           =   915
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Enabled         =   0   'False
                  Height          =   3285
                  Index           =   62
                  Left            =   3645
                  TabIndex        =   103
                  Top             =   1395
                  Width           =   525
               End
            End
            Begin C1SizerLibCtl.C1Elastic Ele 
               Height          =   5400
               Index           =   9
               Left            =   15
               TabIndex        =   104
               TabStop         =   0   'False
               Top             =   15
               Width           =   17295
               _cx             =   30506
               _cy             =   9525
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
               Begin VB.CheckBox Check1 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "÷—»Ì»… «·„»Ì⁄« "
                  Height          =   2880
                  Left            =   5910
                  TabIndex        =   106
                  Top             =   1395
                  Width           =   1140
               End
               Begin VB.TextBox Text8 
                  Alignment       =   1  'Right Justify
                  Height          =   4335
                  Left            =   4470
                  MaxLength       =   4
                  TabIndex        =   105
                  Top             =   900
                  Width           =   765
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
                  Height          =   3345
                  Index           =   69
                  Left            =   4095
                  TabIndex        =   109
                  Top             =   1395
                  Width           =   375
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "«·ﬁÌ„…"
                  Enabled         =   0   'False
                  Height          =   2835
                  Index           =   68
                  Left            =   5235
                  TabIndex        =   108
                  Top             =   1680
                  Width           =   375
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Enabled         =   0   'False
                  Height          =   2880
                  Index           =   67
                  Left            =   3570
                  TabIndex        =   107
                  Top             =   1395
                  Width           =   525
               End
            End
         End
         Begin C1SizerLibCtl.C1Elastic C1Elastic2 
            Height          =   5430
            Left            =   18360
            TabIndex        =   147
            TabStop         =   0   'False
            Top             =   45
            Width           =   17325
            _cx             =   30559
            _cy             =   9578
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
            Begin VSFlex8UCtl.VSFlexGrid Grid3 
               Height          =   3390
               Left            =   2040
               TabIndex        =   148
               Tag             =   "1"
               Top             =   1320
               Width           =   13230
               _cx             =   23336
               _cy             =   5980
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
               AllowUserResizing=   0
               SelectionMode   =   0
               GridLines       =   1
               GridLinesFixed  =   2
               GridLineWidth   =   1
               Rows            =   3
               Cols            =   7
               FixedRows       =   1
               FixedCols       =   1
               RowHeightMin    =   300
               RowHeightMax    =   0
               ColWidthMin     =   0
               ColWidthMax     =   0
               ExtendLastCol   =   -1  'True
               FormatString    =   $"FrmPO10.frx":889B
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
         End
         Begin C1SizerLibCtl.C1Elastic C1Elastic3 
            Height          =   5430
            Left            =   18660
            TabIndex        =   156
            TabStop         =   0   'False
            Top             =   45
            Width           =   17325
            _cx             =   30559
            _cy             =   9578
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
            Begin VSFlex8UCtl.VSFlexGrid VSFlexGrid1 
               Height          =   3390
               Left            =   2280
               TabIndex        =   157
               Tag             =   "1"
               Top             =   1080
               Width           =   13230
               _cx             =   23336
               _cy             =   5980
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
               AllowUserResizing=   0
               SelectionMode   =   0
               GridLines       =   1
               GridLinesFixed  =   2
               GridLineWidth   =   1
               Rows            =   3
               Cols            =   7
               FixedRows       =   1
               FixedCols       =   1
               RowHeightMin    =   300
               RowHeightMax    =   0
               ColWidthMin     =   0
               ColWidthMax     =   0
               ExtendLastCol   =   -1  'True
               FormatString    =   $"FrmPO10.frx":89B8
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
         End
         Begin C1SizerLibCtl.C1Elastic C1Elastic6 
            Height          =   5430
            Left            =   18960
            TabIndex        =   212
            TabStop         =   0   'False
            Top             =   45
            Width           =   17325
            _cx             =   30559
            _cy             =   9578
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
            Begin VB.TextBox TxtValueAdded 
               Alignment       =   1  'Right Justify
               Height          =   372
               Left            =   8505
               TabIndex        =   214
               Top             =   4980
               Width           =   2715
            End
            Begin VB.CheckBox ChecVAT 
               Alignment       =   1  'Right Justify
               Caption         =   " ÕœÌœ «·ﬂ·"
               Height          =   285
               Left            =   15675
               RightToLeft     =   -1  'True
               TabIndex        =   213
               Top             =   840
               Width           =   1200
            End
            Begin VSFlex8UCtl.VSFlexGrid VatGrid 
               Height          =   3630
               Left            =   180
               TabIndex        =   215
               Tag             =   "1"
               Top             =   1215
               Width           =   17100
               _cx             =   30162
               _cy             =   6403
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
               AllowUserResizing=   0
               SelectionMode   =   0
               GridLines       =   1
               GridLinesFixed  =   2
               GridLineWidth   =   1
               Rows            =   3
               Cols            =   8
               FixedRows       =   1
               FixedCols       =   1
               RowHeightMin    =   300
               RowHeightMax    =   0
               ColWidthMin     =   0
               ColWidthMax     =   0
               ExtendLastCol   =   -1  'True
               FormatString    =   $"FrmPO10.frx":8AD5
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
            Begin VB.Label Label22 
               Alignment       =   1  'Right Justify
               Caption         =   "«’‰«› «·ﬁÌ„… «·„÷«›…"
               ForeColor       =   &H00FF0000&
               Height          =   255
               Left            =   11895
               TabIndex        =   217
               Top             =   840
               Width           =   3390
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               Caption         =   " «·«Ã„«·Ì"
               Height          =   285
               Index           =   104
               Left            =   11595
               TabIndex        =   216
               Top             =   5040
               Width           =   1215
            End
         End
      End
      Begin C1SizerLibCtl.C1Elastic Ele 
         Height          =   615
         Index           =   6
         Left            =   450
         TabIndex        =   129
         TabStop         =   0   'False
         Top             =   90
         Width           =   17475
         _cx             =   30824
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
         Caption         =   ""
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
         Begin VB.TextBox TXTNoteID 
            Alignment       =   1  'Right Justify
            Height          =   375
            Left            =   5910
            TabIndex        =   182
            Top             =   240
            Visible         =   0   'False
            Width           =   375
         End
         Begin ImpulseButton.ISButton XPBtnMove 
            Height          =   345
            Index           =   0
            Left            =   1965
            TabIndex        =   130
            Top             =   105
            Width           =   735
            _ExtentX        =   1296
            _ExtentY        =   609
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
            ButtonImage     =   "FrmPO10.frx":8BE9
            ColorButton     =   16777215
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
            Left            =   1215
            TabIndex        =   131
            Top             =   105
            Width           =   630
            _ExtentX        =   1111
            _ExtentY        =   609
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
            ButtonImage     =   "FrmPO10.frx":8F83
            ColorButton     =   16777215
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
            Left            =   2745
            TabIndex        =   132
            Top             =   105
            Width           =   900
            _ExtentX        =   1588
            _ExtentY        =   609
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
            ButtonImage     =   "FrmPO10.frx":931D
            ColorButton     =   16777215
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
            Left            =   180
            TabIndex        =   133
            Top             =   105
            Width           =   690
            _ExtentX        =   1217
            _ExtentY        =   609
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
            ButtonImage     =   "FrmPO10.frx":96B7
            ColorButton     =   16777215
            ColorHighlight  =   4194304
            ColorHoverText  =   16777215
            ColorShadow     =   -2147483631
            ColorOutline    =   -2147483631
            DrawFocusRectangle=   0   'False
            DisabledImageStyle=   1
            ColorToggledHoverText=   16777215
            ColorTextShadow =   16777215
         End
         Begin C1SizerLibCtl.C1Elastic Ele 
            Height          =   75
            Index           =   1
            Left            =   0
            TabIndex        =   189
            TabStop         =   0   'False
            Top             =   480
            Visible         =   0   'False
            Width           =   20145
            _cx             =   35534
            _cy             =   132
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
               Height          =   45
               Index           =   8
               Left            =   -360
               TabIndex        =   190
               Top             =   15
               Visible         =   0   'False
               Width           =   1260
               _ExtentX        =   2223
               _ExtentY        =   79
               ButtonStyle     =   1
               ButtonPositionImage=   1
               Caption         =   "ÿ»«⁄Â ÿ·» ‘—«¡ "
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
         Begin MSComDlg.CommonDialog CommonDialog1 
            Left            =   0
            Top             =   0
            _ExtentX        =   847
            _ExtentY        =   847
            _Version        =   393216
         End
         Begin VB.Label LBLGross 
            Alignment       =   1  'Right Justify
            Caption         =   "Label13"
            Height          =   375
            Left            =   0
            RightToLeft     =   -1  'True
            TabIndex        =   236
            Top             =   0
            Visible         =   0   'False
            Width           =   2055
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "√„— «·‘—«¡ "
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   14.25
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   375
            Index           =   1
            Left            =   11760
            RightToLeft     =   -1  'True
            TabIndex        =   188
            Top             =   2160
            Visible         =   0   'False
            Width           =   4350
         End
         Begin VB.Image Image2 
            Height          =   495
            Left            =   16230
            Picture         =   "FrmPO10.frx":9A51
            Stretch         =   -1  'True
            Top             =   0
            Visible         =   0   'False
            Width           =   1020
         End
         Begin VB.Image ImgFavorites 
            Height          =   390
            Left            =   6270
            Picture         =   "FrmPO10.frx":ABBD
            Stretch         =   -1  'True
            Top             =   120
            Width           =   525
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            BackStyle       =   0  'Transparent
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000C0&
            Height          =   195
            Index           =   64
            Left            =   3840
            TabIndex        =   134
            Top             =   360
            Width           =   7770
         End
      End
   End
   Begin ImpulseButton.ISButton Cmd 
      Height          =   390
      Index           =   9
      Left            =   0
      TabIndex        =   159
      Top             =   0
      Width           =   1035
      _ExtentX        =   1826
      _ExtentY        =   688
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
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "ÿ·» «·‘—«¡"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   2
      Left            =   0
      RightToLeft     =   -1  'True
      TabIndex        =   187
      Top             =   0
      Width           =   3930
   End
   Begin VB.Image Image1 
      Height          =   495
      Left            =   4530
      Picture         =   "FrmPO10.frx":E825
      Stretch         =   -1  'True
      Top             =   0
      Width           =   810
   End
End
Attribute VB_Name = "FrmPO10"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rs As ADODB.Recordset
Dim TTP As clstooltip
Dim NewGrid As New ClsGrid
Dim SaleReport As ClsSaleReport
Dim cSearchDcbo(3)   As clsDCboSearch
 Dim CurrentTransactionType As Integer
Public project_id1 As Integer
Dim mCopyToNew As Boolean

Private Sub chkTaxExempt_Click()
  Dim i As Integer
If Me.TxtModFlg.text <> "R" Then
If chkTaxExempt.value = vbChecked Then
    ChecVAT.value = vbUnchecked
Else
    ChecVAT.value = vbChecked
End If
    If ChecVAT.value = vbChecked Then

        With Me.VatGrid
 
            For i = 1 To .rows - 1
        
                .TextMatrix(i, .ColIndex("select")) = True
            Next i

        End With

    Else

        With Me.VatGrid

            For i = 1 To .rows - 1
        
                .TextMatrix(i, .ColIndex("select")) = False
            Next i

        End With

    End If
    RelinVatGrid
    End If
End Sub

Private Sub cmdCopyToNew_Click()
    mCopyToNew = True
    TxtModFlg.text = "N"
    Me.XPTxtBillID.text = ""
 
    Me.DCboUserName.BoundText = user_id
    'Me.DcBranch.BoundText = Current_branch
    TxtNoteSerial.text = ""
    TxtNoteSerial1.text = ""
    NewGrid.Calculate 1, , , True
    
    mCopyToNew = False
    
End Sub

Private Sub cmdInsertItems_Click()
    
        '*******************
        
        Dim StrSQL  As String
        Dim rs2  As ADODB.Recordset
        Dim mUnitPurPrice As Double
         Dim mUnitId As Long
         Dim LngItemID As Long
        Dim mName As String
        Dim mUnitName As String
        StrSQL = " SELECT TblItemsUnits.UnitPurPrice,TblItemsUnits.UnitSalesPrice,dbo.TblUnites.UnitNamee, dbo.TblUnites.UnitName, dbo.TblItems.Fullcode, dbo.TblUnites.UnitID, dbo.TblItemsUnits.ItemID,TblItems.* "
        StrSQL = StrSQL & "    FROM         dbo.TblItems INNER JOIN"
        StrSQL = StrSQL & "                   dbo.TblItemsUnits ON dbo.TblItems.ItemID = dbo.TblItemsUnits.ItemID INNER JOIN"
        StrSQL = StrSQL & "                      dbo.TblUnites ON dbo.TblItemsUnits.UnitID = dbo.TblUnites.UnitID"
        StrSQL = StrSQL & "  Where DefaultSupplier = " & val(DBCboClientName.BoundText) & " And TblItemsUnits.DefaultUnit = 1"

        Set rs2 = New ADODB.Recordset
        rs2.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText

        Do While Not rs2.EOF
            LngItemID = IIf(IsNull(rs2("ItemID").value), 0, rs2("ItemID").value)
            mUnitId = IIf(IsNull(rs2("UnitID").value), 0, rs2("UnitID").value)
            mUnitPurPrice = IIf(IsNull(rs2("UnitPurPrice").value), 0, rs2("UnitPurPrice").value)
            mName = IIf(IsNull(rs2("ItemName").value), 0, rs2("ItemName").value)
            mUnitName = IIf(IsNull(rs2("UnitName").value), 0, rs2("UnitName").value)
        

        '         ParrtNoCode = "" 'IIf(IsNull(RsTemp("ParrtNoCode").value), "", RsTemp("ParrtNoCode").value)
        '        ItemDetailedCode = "" 'IIf(IsNull(RsTemp("ItemDetailedCode").value), "", RsTemp("ItemDetailedCode").value)
        
        If LngItemID = 0 Then GoTo NextRow
        ' If mCode = "" and  Then GoTo NextRow
         
        With Me.FG

            If .TextMatrix(.rows - 1, .ColIndex("Code")) <> "" Then
                .rows = .rows + 1
            End If
               ' NewGrid.FillGrid

            .TextMatrix(.rows - 1, FG.ColIndex("Code")) = LngItemID
            .TextMatrix(.rows - 1, FG.ColIndex("Name")) = LngItemID
            .TextMatrix(.rows - 1, FG.ColIndex("Name")) = mName
            .TextMatrix(.rows - 1, FG.ColIndex("Count")) = 1
            .TextMatrix(.rows - 1, FG.ColIndex("DiscountType")) = 1
        
            .TextMatrix(.rows - 1, FG.ColIndex("Serial")) = "" ' IIf(IsNull(RsDetails("ItemSerial").value), "", RsDetails("ItemSerial").value)
            .TextMatrix(.rows - 1, FG.ColIndex("HaveSerial")) = "" ' IIf(IsNull(RsDetails("HaveSerial").value), "", RsDetails("HaveSerial").value)
            FG.TextMatrix(.rows - 1, FG.ColIndex("ItemCase")) = "" ' IIf(IsNull(RsDetails("ItemCase")), "", (RsDetails("ItemCase").value))
            FG.TextMatrix(.rows - 1, FG.ColIndex("ColorID")) = 1 ' IIf(IsNull(RsDetails("ColorID")), 1, (RsDetails("ColorID").value))
            FG.TextMatrix(.rows - 1, FG.ColIndex("ItemSize")) = 1 ' IIf(IsNull(RsDetails("ItemSize")), 1, (RsDetails("ItemSize").value))
            FG.TextMatrix(.rows - 1, FG.ColIndex("ClassID")) = 1 ' IIf(IsNull(RsDetails("ClassID")), 1, (RsDetails("ClassID").value))
            FG.cell(flexcpData, .rows - 1, FG.ColIndex("UnitID")) = mUnitId ' IIf(IsNull(RsDetails("UnitID")), "", (RsDetails("UnitID").value))
            'FG.TextMatrix(.Rows - 1, FG.ColIndex("ParrtNoCode")) = ParrtNoCode
            'FG.TextMatrix(.Rows - 1, FG.ColIndex("ItemDetailedCode")) = ItemDetailedCode
            .TextMatrix(.rows - 1, FG.ColIndex("Price")) = mUnitPurPrice
            ' .TextMatrix(.Rows - 1, FG.ColIndex("ShowPrice")) = mUnitPurPrice
            .TextMatrix(.rows - 1, FG.ColIndex("Valu")) = val(.TextMatrix(.rows - 1, .ColIndex("Price"))) * val(.TextMatrix(.rows - 1, .ColIndex("Count")))

            If SystemOptions.UserInterface = ArabicInterface Then
                FG.TextMatrix(.rows - 1, FG.ColIndex("UnitID")) = mUnitName
            Else
                FG.TextMatrix(.rows - 1, FG.ColIndex("UnitID")) = mUnitName
            End If

        End With

NextRow:
        rs2.MoveNext
    Loop

    's = " UPDATE Groups SET ParentID = 1  WHERE ISNULL(ParentID,0) = 0  and
End Sub

Private Sub cmdLoadFile_Click()
    'ExportToExcel Me, Grd, "TT", , "grdExcel"
    tmpGrd.rows = 1

    Dim i                   As Long
    Dim s                   As String
    Dim mIndex              As Long

    Dim rsDummy             As New ADODB.Recordset
    Dim rsDummy2            As New ADODB.Recordset
    Dim mCode               As String
    Dim mGroupID            As Long
    Dim mUnitId             As Long
    Dim mUnitPurPrice       As Double
    Dim mUnitSalesPrice     As Double
    Dim mRatePur            As Double
    Dim mRateSale           As Double
    Dim mNewCode            As String
    Dim mMaxId              As Long
    Dim mName               As String
    Dim mbarCode            As String
    Dim mUnitWholeSalePrice As Double
    Dim mUnitName           As String
    Dim rsDummyUnit         As New ADODB.Recordset
    Dim mQty                As Double
    Dim StrSQL              As String
    mIndex = 1
    Dim rs2       As New ADODB.Recordset
    Dim LngItemID As Long
    GrdExcel(1).rows = 1
    FromExcel GrdExcel(1), tmpGrd, Me, , , txtFileName.text, "TblEmployee"
    'Dim StrSQL As String
    Dim valid As Boolean
    valid = True
    Dim GroupName
    For i = 1 To GrdExcel(mIndex).rows - 1
        mCode = Trim(GrdExcel(mIndex).TextMatrix(i, GrdExcel(mIndex).ColIndex("Fullcode")))
        mName = Trim(GrdExcel(mIndex).TextMatrix(i, GrdExcel(mIndex).ColIndex("ItemName")))
        GroupName = Trim(GrdExcel(mIndex).TextMatrix(i, GrdExcel(mIndex).ColIndex("Group")))
        If Trim(mCode) = "" And Trim(mName) = "" And Trim(GroupName) = "" Then
            '            MsgBox i & " ﬂÊœ «·’‰› Ê«”„ «·’‰› ›«—€ ·« Ì„ﬂ‰ «ﬂ„«· «·⁄„·ÌÂ "
            '            valid = False
            GoTo mnextrow
        End If
        If Trim(mCode) = "" And Trim(mName) = "" Then
            MsgBox i & " ﬂÊœ «·’‰› Ê«”„ «·’‰› ›«—€ ·« Ì„ﬂ‰ «ﬂ„«· «·⁄„·ÌÂ "
            valid = False
            Exit Sub
        End If
        If Trim(GroupName) = "" Then
            MsgBox i & "    «·„Ã„Ê⁄Â ›«—€Â ·« Ì„ﬂ‰ «ﬂ„«· «·⁄„·ÌÂ "
            valid = False
            Exit Sub
        End If
        
        GroupName = Trim(GrdExcel(mIndex).TextMatrix(i, GrdExcel(mIndex).ColIndex("Group")))
        StrSQL = "SELECT Groups.GroupID, "
        StrSQL = StrSQL & "       Groups.Code GroupCode, "
        StrSQL = StrSQL & "       Groups.ParentID, "
        StrSQL = StrSQL & "       maingrp.Code MainCode "
        StrSQL = StrSQL & "FROM dbo.Groups "
        StrSQL = StrSQL & "    LEFT OUTER JOIN Groups maingrp "
        StrSQL = StrSQL & "        ON Groups.ParentID = maingrp.GroupID "
        StrSQL = StrSQL & "   Where Groups.GroupName Like N'%" & GroupName & "%'"
        Set rs2 = New ADODB.Recordset
        rs2.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

        If rs2.EOF Then
           
            MsgBox i & "«·„Ã„Ê⁄Â €Ì— ’ÕÌÕÂ ·« Ì„ﬂ‰ «ﬂ„«· «·⁄„·ÌÂ "
            valid = False
            Exit Sub
        End If
        rs2.Close
mnextrow:
    Next
    If valid = False Then
        Exit Sub
     
    End If
    For i = 1 To GrdExcel(mIndex).rows - 1
        mCode = Trim(GrdExcel(mIndex).TextMatrix(i, GrdExcel(mIndex).ColIndex("Fullcode")))
        mbarCode = Trim(GrdExcel(mIndex).TextMatrix(i, GrdExcel(mIndex).ColIndex("barCodeNO")))
        mQty = val(GrdExcel(mIndex).TextMatrix(i, GrdExcel(mIndex).ColIndex("ShowQty")))
        mUnitName = Trim(GrdExcel(mIndex).TextMatrix(i, GrdExcel(mIndex).ColIndex("UnitName")))
        mUnitWholeSalePrice = val(GrdExcel(mIndex).TextMatrix(i, GrdExcel(mIndex).ColIndex("UnitWholeSalePrice")))
        Set rsDummyUnit = New ADODB.Recordset
        s = "Select UnitName,UnitId from TblUnites Where UnitName Like '%" & Trim(mUnitName) & "%'"
        rsDummyUnit.Open s, Cn, adOpenStatic, adLockReadOnly

        If Not rsDummyUnit.EOF Then
            mUnitId = val(rsDummyUnit!UnitID & "")
        End If
    
        mUnitPurPrice = val(GrdExcel(mIndex).TextMatrix(i, GrdExcel(mIndex).ColIndex("UnitPurPrice")))
        mName = Trim(GrdExcel(mIndex).TextMatrix(i, GrdExcel(mIndex).ColIndex("ItemName")))
    
        mUnitSalesPrice = val(GrdExcel(mIndex).TextMatrix(i, GrdExcel(mIndex).ColIndex("UnitSalesPrice")))
        If mName = "" Then Exit Sub
        '********try to get item **********
        
        StrSQL = "       SELECT * FROM dbo.TblItems "
        StrSQL = StrSQL & "       WHERE (TblItems.Fullcode  = N'" & mCode & "' Or TblItems.ItemName = N'" & mName & "'); "
        Dim rsCheckData As New ADODB.Recordset
        rsCheckData.Open StrSQL, Cn, adOpenStatic, adLockOptimistic

        If rsCheckData.EOF Then
            'Insert New Item
            If Not SaveItemsExcelMeth2(GrdExcel(mIndex), i) Then
                GoTo NextRow
            End If
            mCode = Trim(GrdExcel(mIndex).TextMatrix(i, GrdExcel(mIndex).ColIndex("Fullcode")))
        Else
            mCode = Trim(rsCheckData!Fullcode & "")
        End If
        rsCheckData.Close
        
        '*******************
        StrSQL = "  SELECT     dbo.TblUnites.UnitNamee, dbo.TblUnites.UnitName, dbo.TblItems.Fullcode, dbo.TblUnites.UnitID, dbo.TblItemsUnits.ItemID"
        StrSQL = StrSQL & "    FROM         dbo.TblItems INNER JOIN"
        StrSQL = StrSQL & "                   dbo.TblItemsUnits ON dbo.TblItems.ItemID = dbo.TblItemsUnits.ItemID INNER JOIN"
        StrSQL = StrSQL & "                      dbo.TblUnites ON dbo.TblItemsUnits.UnitID = dbo.TblUnites.UnitID"
        StrSQL = StrSQL & "  WHERE     (TblItems.Fullcode  = N'" & mCode & "' Or TblItems.ItemName = N'" & mName & "') AND (dbo.TblItemsUnits.UnitID = " & mUnitId & ")"
        StrSQL = StrSQL & "  GROUP BY dbo.TblUnites.UnitNamee, dbo.TblUnites.UnitName, dbo.TblItems.Fullcode, dbo.TblUnites.UnitID, dbo.TblItemsUnits.ItemID"
        Set rs2 = New ADODB.Recordset
        rs2.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText

        If rs2.RecordCount > 0 Then
            LngItemID = IIf(IsNull(rs2("ItemID").value), 0, rs2("ItemID").value)
            mUnitId = IIf(IsNull(rs2("UnitID").value), 0, rs2("UnitID").value)
        End If

        '         ParrtNoCode = "" 'IIf(IsNull(RsTemp("ParrtNoCode").value), "", RsTemp("ParrtNoCode").value)
        '        ItemDetailedCode = "" 'IIf(IsNull(RsTemp("ItemDetailedCode").value), "", RsTemp("ItemDetailedCode").value)
        
        If LngItemID = 0 Then GoTo NextRow
        ' If mCode = "" and  Then GoTo NextRow
         
        With Me.FG

            If .TextMatrix(.rows - 1, .ColIndex("Code")) <> "" Then
                .rows = .rows + 1
            End If
            ' NewGrid.FillGrid

            'FG.TextMatrix(Num, FG.ColIndex("Code")) = IIf(IsNull(RsDetails("ItemCode")), "", (RsDetails("ItemCode").value))

            FG.cell(flexcpData, .rows - 1, FG.ColIndex("Code")) = LngItemID
            FG.cell(flexcpData, .rows - 1, FG.ColIndex("Name")) = LngItemID
            .TextMatrix(.rows - 1, FG.ColIndex("Code")) = LngItemID
            .TextMatrix(.rows - 1, FG.ColIndex("Name")) = LngItemID
            .TextMatrix(.rows - 1, FG.ColIndex("Name")) = mName
            .TextMatrix(.rows - 1, FG.ColIndex("Count")) = mQty
            .TextMatrix(.rows - 1, FG.ColIndex("DiscountType")) = 1
        
            .TextMatrix(.rows - 1, FG.ColIndex("Serial")) = "" ' IIf(IsNull(RsDetails("ItemSerial").value), "", RsDetails("ItemSerial").value)
            .TextMatrix(.rows - 1, FG.ColIndex("HaveSerial")) = "" ' IIf(IsNull(RsDetails("HaveSerial").value), "", RsDetails("HaveSerial").value)
            FG.TextMatrix(.rows - 1, FG.ColIndex("ItemCase")) = "" ' IIf(IsNull(RsDetails("ItemCase")), "", (RsDetails("ItemCase").value))
            FG.TextMatrix(.rows - 1, FG.ColIndex("ColorID")) = 1 ' IIf(IsNull(RsDetails("ColorID")), 1, (RsDetails("ColorID").value))
            FG.TextMatrix(.rows - 1, FG.ColIndex("ItemSize")) = 1 ' IIf(IsNull(RsDetails("ItemSize")), 1, (RsDetails("ItemSize").value))
            FG.TextMatrix(.rows - 1, FG.ColIndex("ClassID")) = 1 ' IIf(IsNull(RsDetails("ClassID")), 1, (RsDetails("ClassID").value))
            FG.cell(flexcpData, .rows - 1, FG.ColIndex("UnitID")) = mUnitId ' IIf(IsNull(RsDetails("UnitID")), "", (RsDetails("UnitID").value))
            'FG.TextMatrix(.Rows - 1, FG.ColIndex("ParrtNoCode")) = ParrtNoCode
            'FG.TextMatrix(.Rows - 1, FG.ColIndex("ItemDetailedCode")) = ItemDetailedCode
            .TextMatrix(.rows - 1, FG.ColIndex("Price")) = mUnitPurPrice
            ' .TextMatrix(.Rows - 1, FG.ColIndex("ShowPrice")) = mUnitPurPrice
            .TextMatrix(.rows - 1, FG.ColIndex("Valu")) = val(.TextMatrix(.rows - 1, .ColIndex("Price"))) * val(.TextMatrix(.rows - 1, .ColIndex("Count")))

            If SystemOptions.UserInterface = ArabicInterface Then
                FG.TextMatrix(.rows - 1, FG.ColIndex("UnitID")) = mUnitName
            Else
                FG.TextMatrix(.rows - 1, FG.ColIndex("UnitID")) = mUnitName
            End If

        End With
        LngItemID = 0
        mCode = ""
NextRow:
    Next
ReLineGrid

    's = " UPDATE Groups SET ParentID = 1  WHERE ISNULL(ParentID,0) = 0  and GroupID <> 1"
    ' Cn.Execute s
    '
    ' s = " UPDATE groups SET Code = fullCode WHERE ISNULL(code,'') = ''"
    ' Cn.Execute s
 
End Sub

Private Function SaveItemsExcelMeth2(mygrd, i) As Boolean
    
    Dim s                   As String
    Dim rsDummy             As New ADODB.Recordset
    Dim rsDummy2            As New ADODB.Recordset
    Dim mCode               As String
    Dim mGroupID            As Long
    Dim mUnitId             As Long
    Dim mUnitPurPrice       As Double
    Dim mUnitSalesPrice     As Double
    Dim mRatePur            As Double
    Dim mRateSale           As Double
    Dim mNewCode            As String
    Dim mMaxId              As Long
    Dim mUnitName           As String
    Dim mName               As String
    Dim mbarCode            As String
    Dim mUnitWholeSalePrice As Double
    
    Dim rsDummySupp         As New ADODB.Recordset
    Dim mIndex              As Long
    Dim rsDummyUnit         As New ADODB.Recordset
    Dim mMinSelingPrice     As Double
    Dim mMaxSelingPrice     As Double
    Dim mSelingPriceDestr   As Double
    Dim mDefaultSupplier    As String
    Dim mDefaultSupplierID  As Integer
    Dim GroupName           As String
    Dim MainGroupName       As String
    Dim mGroupCode2 As String
    Dim rsgg As ADODB.Recordset
    '***************************
    Dim StrSQL As String
    SaveItemsExcelMeth2 = True
    '******************************
    
    GroupName = Trim(mygrd.TextMatrix(i, mygrd.ColIndex("Group")))
    MainGroupName = Trim(mygrd.TextMatrix(i, mygrd.ColIndex("MainGroup")))
      
    mCode = Trim(mygrd.TextMatrix(i, mygrd.ColIndex("Fullcode")))
    mbarCode = Trim(mygrd.TextMatrix(i, mygrd.ColIndex("barCodeNO")))
    
    mUnitName = Trim(mygrd.TextMatrix(i, mygrd.ColIndex("UnitName")))
    mUnitWholeSalePrice = val(mygrd.TextMatrix(i, mygrd.ColIndex("UnitWholeSalePrice")))
    mDefaultSupplier = "" ' Trim(mygrd.TextMatrix(i, mygrd.ColIndex("DefaultSupplier")))
    mMinSelingPrice = 0 ' val(mygrd.TextMatrix(i, mygrd.ColIndex("MinSelingPrice")))
    
    mMaxSelingPrice = 0 '  val(mygrd.TextMatrix(i, mygrd.ColIndex("MaxSelingPrice")))
    mSelingPriceDestr = 0 ' val(mygrd.TextMatrix(i, mygrd.ColIndex("SelingPriceDestr")))

    Set rsDummyUnit = New ADODB.Recordset
    s = "Select UnitName,UnitId from TblUnites Where UnitName Like '%" & Trim(mUnitName) & "%'"
    rsDummyUnit.Open s, Cn, adOpenStatic, adLockReadOnly

    If Not rsDummyUnit.EOF Then
        mUnitId = val(rsDummyUnit!UnitID & "")
    End If
    
    '    Set rsDummySupp = New ADODB.Recordset
    '    s = "SELECT CusID FROM TblCustemers Where (CusName Like '%" & Trim(mDefaultSupplier) & "%'     Or CusNamee Like '%" & Trim(mDefaultSupplier) & "%')"
    '    rsDummySupp.Open s, Cn, adOpenStatic, adLockReadOnly
    '
    '    If Not rsDummySupp.EOF Then
    '        mDefaultSupplierID = val(rsDummySupp!CusID & "")
    '    End If
    
    mUnitPurPrice = val(mygrd.TextMatrix(i, mygrd.ColIndex("UnitPurPrice")))
    mName = Trim(mygrd.TextMatrix(i, mygrd.ColIndex("ItemName")))
    
    mUnitSalesPrice = val(mygrd.TextMatrix(i, mygrd.ColIndex("UnitSalesPrice")))
  
    If Trim(mCode) = "" And Trim(mName) = "" Then
       ' Err.Raise vbObjectError + 1000, Me.Name, "ﬂÊœ «·’‰› Ê«”„ «·’‰› ›«—€ ·« Ì„ﬂ‰ «ﬂ„«· «·⁄„·ÌÂ"
       SaveItemsExcelMeth2 = False
       Exit Function
    End If
    
    s = "Select * from tblItems where  1 = 2 "  ' GroupId = " & mGroupID & " and FullCode = N'" & mCode & "' and ItemName =N'" & mName & "'"
    Set rsDummy = New ADODB.Recordset
    rsDummy.Open s, Cn, adOpenKeyset, adLockOptimistic

    Dim ItemID       As Integer
    Dim ItemCodeOnly As String
   
    rsDummy.AddNew
    
    mNewCode = GetNewItemCode(GroupName, MainGroupName, ItemID, ItemCodeOnly, mGroupID)
       
        StrSQL = "SELECT Groups.GroupID, "
    StrSQL = StrSQL & "       Groups.Code GroupCode, "
    StrSQL = StrSQL & "       Groups.ParentID, "
    StrSQL = StrSQL & "       maingrp.Code MainCode "
    StrSQL = StrSQL & "FROM dbo.Groups "
    StrSQL = StrSQL & "    LEFT OUTER JOIN Groups maingrp "
    StrSQL = StrSQL & "        ON Groups.ParentID = maingrp.GroupID "
    StrSQL = StrSQL & "   Where Groups.GroupName Like N'%" & GroupName & "%'"
    Set rsgg = New ADODB.Recordset
    rsgg.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText
    If Not rsgg.EOF Then
        mGroupCode2 = Trim(rsgg!GroupCode & "")
    End If
    
   If mNewCode = "" Then
   SaveItemsExcelMeth2 = False
   Exit Function
   End If
   
    rsDummy!ItemID = ItemID
    rsDummy!ItemName = IIf(mName = "", mCode, mName)
        
    rsDummy!HaveSerial = 0
    rsDummy!HaveGuarantee = 0
    rsDummy!DealerPrice = 0
    rsDummy!GuaranteeValue = 0
    rsDummy!GuaranteeType = 0
    rsDummy!DefaultSupplier = mDefaultSupplierID
        
    rsDummy!IsArchive = 0
    rsDummy!ItemType = 0
    rsDummy!AssbliedItem = 0
    rsDummy!RelatedItem = 0
    rsDummy!ItemCase = 1
    rsDummy!AssbliedItem = 0
    rsDummy!prifix = mGroupCode2
    If mCode = "" Then mCode = mNewCode
    rsDummy!Fullcode = IIf(mCode = "", mNewCode, mCode)
    rsDummy!itemcode = ItemCodeOnly ' mCode 'IIf(mCode = "", ItemCodeOnly, mCode)
    'rsDummy!itemcode = mNewCode ' IIf(mCode = "", ItemCodeOnly, mCode)
    rsDummy!barCodeNO = IIf(mbarCode = "", mCode, mbarCode)
    rsDummy!code = ItemCodeOnly ' IIf(mCode = "", ItemCodeOnly, mCode)
    rsDummy!GroupID = mGroupID

    rsDummy!SallingPrice = mUnitSalesPrice
    rsDummy.update

    If Trim(mCode = "") Then
        mygrd.TextMatrix(i, mygrd.ColIndex("Fullcode")) = mNewCode
    End If

    s = "Select * from TblItemsUnits where ItemId = " & val(rsDummy!ItemID & "") & " and UnitId = " & mUnitId
    Set rsDummy2 = New ADODB.Recordset
    rsDummy2.Open s, Cn, adOpenKeyset, adLockOptimistic

    If rsDummy2.EOF Then
        rsDummy2.AddNew
        rsDummy2!ItemID = val(rsDummy!ItemID & "")
        rsDummy2!UnitID = mUnitId
        rsDummy2!UnitFactor = 1
        rsDummy2!FactorByDefaultUnit = 1
        rsDummy2!FactorBySmallUnit = 1
        rsDummy2!DefaultUnit = 1
        rsDummy2!MinSelingPrice = mMinSelingPrice
        rsDummy2!MaxSelingPrice = mMaxSelingPrice
        rsDummy2!SelingPriceDestr = mSelingPriceDestr

    End If

    rsDummy2!UnitPurPrice = mUnitPurPrice
    rsDummy2!UnitSalesPrice = (mUnitSalesPrice)
    rsDummy2!UnitWholeSalePrice = (mUnitWholeSalePrice)
    rsDummy2.update

End Function

Private Function GetNewItemCode(gName, MainName, ItemID, itemcode, GroupID) As String
'GroupName, MainGroupName, itemId, ItemCodeOnly, mGroupID
    Dim rs2            As ADODB.Recordset
    Dim StrSQL        As String
    
    Dim MainGroupcode As String
    Dim GroupCode     As String
    Dim MainID        As Integer
    

    On Error GoTo ErrTrap
    StrSQL = "SELECT Groups.GroupID, "
    StrSQL = StrSQL & "       Groups.Code GroupCode, "
    StrSQL = StrSQL & "       Groups.ParentID, "
    StrSQL = StrSQL & "       maingrp.Code MainCode "
    StrSQL = StrSQL & "FROM dbo.Groups "
    StrSQL = StrSQL & "    LEFT OUTER JOIN Groups maingrp "
    StrSQL = StrSQL & "        ON Groups.ParentID = maingrp.GroupID "
    StrSQL = StrSQL & "   Where Groups.GroupName Like N'%" & gName & "%'"
    Set rs2 = New ADODB.Recordset
    rs2.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

    If Not rs2.EOF Then
       
        MainGroupcode = rs2!MainCode & ""
        MainID = val(rs2!ParentID & "")
        GroupCode = rs2!GroupCode & ""
        GroupID = val(rs2!GroupID & "")
    Else

        If Trim(MainName) = "" Then
           ' Err.Raise vbObjectError + 1000, Me.Name, "«·„Ã„Ê⁄Â «·—∆Ì”ÌÂ ›«—€Â ·« Ì„ﬂ‰ «ﬂ„«· «·⁄„·ÌÂ "
           GetNewItemCode = ""
           Exit Function
        End If

        StrSQL = "SELECT Groups.GroupID, "
        StrSQL = StrSQL & "       Groups.Code GroupCode, "
        StrSQL = StrSQL & "       Groups.ParentID, "
        StrSQL = StrSQL & "       maingrp.Code MainCode "
        StrSQL = StrSQL & "FROM dbo.Groups "
        StrSQL = StrSQL & "    LEFT OUTER JOIN Groups maingrp "
        StrSQL = StrSQL & "        ON Groups.ParentID = maingrp.GroupID "
        StrSQL = StrSQL & "   Where Groups.GroupName Like N'%" & MainName & "%'"
        rs2.Close
        rs2.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

        If rs2.EOF Then
            If rs2!GroupCode & "" = "" Then
                'Err.Raise vbObjectError + 1000, Me.Name, "«·„Ã„Ê⁄Â «·—∆Ì”ÌÂ €Ì— ’ÕÌÕÂ ·« Ì„ﬂ‰ «ﬂ„«· «·⁄„·ÌÂ "
                GetNewItemCode = ""
                Exit Function
            End If

        Else
            MainGroupcode = rs2!GroupCode & ""
            MainID = val(rs2!GroupID & "")
            Dim GID   As Integer
            Dim GCode As String
            rs2.Close
            rs2.Open " SELECT MAX(GroupID) id FROM dbo.Groups", Cn, adOpenForwardOnly, adLockReadOnly
            GID = (rs2!ID) + 1
            rs2.Close
            rs2.Open " SELECT MAX(Code) id FROM dbo.Groups  WHERE ParentID = " & MainID & "  ", Cn, adOpenForwardOnly, adLockReadOnly
            GCode = Format(val(Replace((rs2!ID & ""), MainGroupcode, "", 1)) + 1, "00")
            StrSQL = "  SELECT * FROM dbo.Groups Where 1 = 2 "
            rs2.Close
            rs2.Open StrSQL, Cn, adOpenKeyset, adLockOptimistic
            rs2.AddNew
            rs2!GroupCode = MainGroupcode & GCode
            rs2!GroupID = GID
            rs2!GroupName = gName
            rs2!ParentID = MainID
            rs2!code = MainGroupcode & GCode
            rs2.update
            GroupCode = rs2!GroupCode & ""
            GroupID = val(rs2!GroupID & "")
            
        End If
         
    End If

    Dim rsITem   As New ADODB.Recordset
     
   
     
    rsITem.Open " SELECT MAX(ItemID) id  FROM dbo.TblItems ", Cn, adOpenForwardOnly, adLockReadOnly
    ItemID = val(rsITem!ID & "") + 1
    rsITem.Close
    rsITem.Open " SELECT MAX(ItemCode) id  FROM dbo.TblItems  WHERE GroupID = " & GroupID & " ", Cn, adOpenForwardOnly, adLockReadOnly
    itemcode = Format(val(rsITem!ID & "") + 1, "0000")
    
    GetNewItemCode = GroupCode & itemcode
    Exit Function
ErrTrap:
End Function


Public Sub FromExcel(ByRef mGrid As Object, _
                     ByRef mtmpGrd As Object, _
                     Frm As Form, _
                     Optional MainFormName As String = "", _
                     Optional ProgressBar As Object = Nothing, Optional ByVal XlsFileName As String = "", Optional ByVal MainTableName As String = "")


    ' If Not i Then Exit Sub
       Dim cProgress As ClsProgress
       Dim Hide As Integer
       Dim i As Long
       Dim j As Long
       Dim jj As Long
       Dim H As Long
    '    Dim mtmpGrd As VSFlexGrid
    If XlsFileName = "" Then
    MsgBox "Õœœ «·„·› «Ê·«", vbCritical
    Exit Sub
        'XlsFileName = GetGridFileName(mGrid, MainFormName)
    End If
    If FileExists(XlsFileName) Then

        mtmpGrd.FixedCols = 0
        mtmpGrd.FixedRows = 0

        mtmpGrd.loadgrid XlsFileName, flexFileExcel

        mtmpGrd.backcolor = &HFFFFFF
        mtmpGrd.BackColorAlternate = &HE9E9E9
        mtmpGrd.BackColorBkg = &H8000000C
        mtmpGrd.BackColorFixed = &H8000000F
        mtmpGrd.BackColorFrozen = &HC0FFFF
        mtmpGrd.BackColorSel = &H8000000D
        mtmpGrd.ForeColor = &H80000008
        mtmpGrd.ForeColorFixed = &HFF0000
        mtmpGrd.ForeColorSel = &H8000000E
        mtmpGrd.GridColor = &H8000000F
        mtmpGrd.GridColorFixed = &H80000010
        mtmpGrd.FixedCols = 1
        mtmpGrd.FixedRows = 1
        '·«‰ Loaded ÌŒ ›Ì
        mtmpGrd.Cols = mGrid.Cols + 1
        mtmpGrd.ColKey(mtmpGrd.Cols - 1) = "Loaded"
        mtmpGrd.ColHidden(mtmpGrd.Cols - 1) = True
        mtmpGrd.AutoSize 0, mtmpGrd.Cols - 1
    End If
    mGrid.rows = 1
    mGrid.rows = mtmpGrd.rows - 4

    '********************************
    If Not ProgressBar Is Nothing Then
        ProgressBar.Min = 1
        ProgressBar.Max = IIf(mGrid.rows > 2, mGrid.rows - 1, 2)    ' mGrid.Rows - 1
        ProgressBar.Visible = True
        '********************************
    End If
        Set cProgress = New ClsProgress
       cProgress.ProgressType = Waiting
    

    



       
        For i = 1 To mtmpGrd.rows - 1
        '********************************
        If Not ProgressBar Is Nothing Then
            ProgressBar.value = i
            DoEvents
            ProgressBar.Refresh
        End If
        cProgress.StartProgress
       DoEvents
        '********************************
        jj = 0
        For j = 1 To mGrid.Cols - 1
            If j = 18 Then
                j = 18
            End If
            If Not mGrid.ColHidden(j) Then
                jj = jj + 1
                       If mGrid.ColKey(j) = "MainGroumName" Then
                    j = j
                End If
                If i > mGrid.rows - 1 Then
                    mGrid.rows = mGrid.rows + 1
                End If
                
                Debug.Print i & " " & mGrid.TextMatrix(i, j)
                If InStr(1, mGrid.ColComboList(j), "#") Then
                    Hide = 0
                    For H = j - 1 To 1 Step -1
                        Hide = Hide + IIf(mGrid.ColHidden(H), 1, 0)
                    Next
                    mGrid.TextMatrix(i, j) = mtmpGrd.TextMatrix(i, j - Hide)
                    'Replace(Trim(mtmpGrd.TextMatrix(i, jj)), "'", "")
                Else
                    mGrid.TextMatrix(i, j) = Replace(Trim(mtmpGrd.TextMatrix(i, jj)), "'", "")
                End If
                If Trim(mGrid.ColEditMask(j)) = "Date" Then
                    GetFieldID mGrid.ColEditMask(j), i, j, mGrid
                End If
                'pValue = Split(G.ColComboList(j), ";")
            Else
                j = j
                If j = 34 Then
                j = j
                End If
                If Trim(mGrid.ColEditMask(j)) <> "" Then
                    GetFieldID mGrid.ColEditMask(j), i, j, mGrid, MainTableName
                End If
                If Trim(mGrid.ColComboList(j)) <> "" Then
                    GetIDCombo Trim(mGrid.ColComboList(j)), i, j, mGrid
                End If
            End If
            If Trim(Replace(Trim(mtmpGrd.TextMatrix(i, 1)), "'", "")) = "" _
            And Trim(Replace(Trim(mtmpGrd.TextMatrix(i, 3)), "'", "")) = "" Then
                mGrid.rows = i + 1:  Exit Sub
            End If
        Next
        ' DisplayOrderTotals
NextRow:
    Next
    '********************************
    If Not ProgressBar Is Nothing Then
        ProgressBar.Visible = False
    End If
           DoEvents
    cProgress.FinishProgress
    cProgress.StopProgess
    Set cProgress = Nothing
    MsgBox " „ «·«œ—«Ã"
    '********************************
End Sub

Function CheckAcconts() As Boolean
CheckAcconts = False
Dim Account_Code_dynamic101 As String
Dim Account_Code_dynamic102 As String

            Account_Code_dynamic101 = get_account_code_branch(101, my_branch)
            Account_Code_dynamic102 = get_account_code_branch(102, my_branch)
             If Account_Code_dynamic101 = "NO account" Then
                                            If SystemOptions.UserInterface = ArabicInterface Then
                                                MsgBox "·„ Ì „  ÕœÌœ Õ”«»   «·„œÌ‰ ·«Ê«„— «·‘—«¡  ›Ì «·›—⁄ ·Â–… «·⁄„·Ì…", vbCritical
                                            Else
                                                MsgBox "Sales Cost Account Not Defined in this Branch", vbCritical
                                            End If
            
                                GoTo ErrTrap
              End If
              
              
              
                  If Account_Code_dynamic102 = "NO account" Then
                                            If SystemOptions.UserInterface = ArabicInterface Then
                                                MsgBox "·„ Ì „  ÕœÌœ Õ”«»   «·œ«∆‰ ·«Ê«„— «·‘—«¡  ›Ì «·›—⁄ ·Â–… «·⁄„·Ì…", vbCritical
                                            Else
                                                MsgBox "Sales Cost Account Not Defined in this Branch", vbCritical
                                            End If
            
                                GoTo ErrTrap
              End If
              
              
  
   CheckAcconts = True
   Exit Function
ErrTrap:
      CheckAcconts = False
End Function

Private Sub CBoBasedON_Change()
If Me.TxtModFlg <> "R" Then
TxtPO6.text = ""
End If

End Sub

Private Sub CBoBasedON_Click()
CBoBasedON_Change
End Sub

Private Sub ChAuto_Click()
If Me.ChAuto.value = vbChecked Then
ISButton4.Visible = False
ISButton3.Visible = True
Else
ISButton3.Visible = False
ISButton4.Visible = True
End If
End Sub

Private Sub Cmd_Click(index As Integer)
    Dim intDef As Integer
    '   On Error GoTo ErrTrap
    
    
    
        Dim Transaction_Type As Integer
        Dim Sanad_No As Integer

        If Me.CboPriceType.ListIndex = 0 Then
            Transaction_Type = CurrentTransactionType
            Sanad_No = CurrentTransactionType
  
 
         
        End If



 

        

    If Voucher_coding(val(dcBranch.BoundText), XPDtbBill.value, Sanad_No, 0, , Transaction_Type, , , , , , val(DCboUserName.BoundText)) = "" Then
        TxtNoteSerial1.locked = False
    Else
        TxtNoteSerial1.locked = True
 
    End If
    
    
    Select Case index

        Case 0

            If DoPremis(Do_New, Me.Name, True) = False Then
                Exit Sub
            End If

            clear_all Me
            TxtModFlg.text = "N"
            NewGrid.GridDefaultValue 1
            Me.DCboUserName.BoundText = user_id
            intDef = GetSetting(StrAppRegPath, "DefaultOptions", "DefaultClient", 2)
            DBCboClientName.BoundText = intDef
            intDef = GetSetting(StrAppRegPath, "DefaultOptions", "DefaultSaleStore", 1)
            DCboStoreName.BoundText = intDef
            Dccurrency.BoundText = 1
            '    FG.SetFocus
            FG.Col = FG.ColIndex("Code")
            FG.row = FG.rows - 1
            Me.CboPriceType.ListIndex = 0
            CboPayMentType.ListIndex = 0
                   
            Dim dstore       As Integer
            Dim dBox         As Integer
            Dim usertype     As Integer
            Dim EmpID        As Integer
            Dim userbranchid As Integer
            'GetBranchData branch_id, dstore, dBox
                 
            GetUserData user_id, usertype, userbranchid, dstore, dBox, , EmpID
     
            If usertype <> 0 Then 'admin
                dcBranch.Enabled = False
 
                DCboStoreName.Enabled = True
                '  TxtStoreID.Enabled = False
                Me.DCboStoreName.BoundText = dstore
            Else
                dcBranch.Enabled = True
 
                DCboStoreName.Enabled = True
 
                Me.dcBranch.BoundText = ""
                Me.DCboStoreName.BoundText = ""
                '                TxtStoreID.Enabled = True
            End If

            If SystemOptions.usertype <> UserAdminAll Then
                If checkmanyBranches = False Then
                    Me.dcBranch.Enabled = True
                Else
                    Me.dcBranch.Enabled = True
                End If
                    
                If checkmanyStores = False Then
                    Me.DCboStoreName.Enabled = True
                                    
                Else
                    Me.DCboStoreName.Enabled = True
 
                End If
                                  
            End If
                   
            Me.dcBranch.BoundText = Current_branch
            DBPix202.ImageClear
            Accredit.Enabled = True
            If SystemOptions.UserInterface = ArabicInterface Then
                Accredit.Caption = "   «·«—”«· ··«⁄ „«œ "
            Else
                Accredit.Caption = " send to Approval   "
            End If
            FillOrderGrid
            FillOrderGrid2
            opt(1).value = True

        Case 1

            If DoPremis(Do_Edit, Me.Name, True) = False Then
                Exit Sub
            End If

            If ChekClodePeriod(Me.XPDtbBill.value) = True Then
                If SystemOptions.UserInterface = ArabicInterface Then
                    MsgBox "› —Â „€·ﬁ… "
                Else
                    MsgBox "Please Change Date Becouse This is Period is Closed"
                End If
                Exit Sub
            End If
              
            Dim StrSQL As String
            Dim rs2    As New ADODB.Recordset
   
            StrSQL = "SELECT NoteSerial1 FROM Transactions where Transaction_Type = 22 and IsNull(order_no,0)  = '" & Trim(TxtNoteSerial1.text) & "' "
            rs2.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText
            If Not rs2.EOF Then
                MsgBox "Â–« «·«„— ·« Ì„ﬂ‰ «· ⁄œÌ· ⁄·ÌÂ ›ﬁœ «” Œœ„ ›Ï ›« Ê—… „‘ —Ì«  —ﬁ„ " & rs2!NoteSerial1 & ""
            
                Exit Sub
            End If

            If ScreenAproved(val(Me.XPTxtBillID.text), Me.Name) = True Then
                If SystemOptions.UserInterface = ArabicInterface Then
                    MsgBox "«·Õ—ﬂÂ „— »ÿÂ »«·«⁄ „«œ« "
                Else
                    MsgBox "Can not edit.This process associated with approvals"
                End If
                Exit Sub
            End If

            TxtModFlg.text = "E"
            CuurentLogdata
            Me.DCboUserName.BoundText = user_id

        Case 2
            Dim Msg As String

            If SystemOptions.POMustentryAndBillMustEntry = True And (TxtPO6.text = "" Or CBoBasedON.ListIndex = 0) Then
                MsgBox "·«Ì„ﬂ‰ «·Õ›Ÿ ·⁄œ„ «Œ Ì«— »‰«¡ ⁄·Ì Ê ÕœÌœ «·—ﬁ„", vbCritical
                Exit Sub
            End If

            If Trim(dcBranch.BoundText) = "" Then
                If SystemOptions.UserInterface = EnglishInterface Then
                    Msg = "Specify Departement"
                Else
                    Msg = "Õœœ «·›—⁄ «Ê·« "
                End If
              
                MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
                dcBranch.SetFocus
                Sendkeys "{F4}"
                Screen.MousePointer = vbDefault
                Exit Sub
            End If
            'ma If SystemOptions.PoCreateVoucher = True Then
            'ma  If CheckAcconts = False Then Exit Sub
            'ma End If

            SaveData

        Case 3
            Undo

        Case 4

            If DoPremis(Do_Delete, Me.Name, True) = False Then
                Exit Sub
            End If

            Del_TransAction

        Case 5

            If DoPremis(Do_Search, Me.Name, True) = False Then
                Exit Sub
            End If

            If ChekClodePeriod(Me.XPDtbBill.value) = True Then
                If SystemOptions.UserInterface = ArabicInterface Then
                    MsgBox "› —Â „€·ﬁ… "
                Else
                    MsgBox "Please Change Date Becouse This is Period is Closed"
                End If
                Exit Sub
            End If

            If ScreenAproved(val(Me.XPTxtBillID.text), Me.Name) = True Then
                If SystemOptions.UserInterface = ArabicInterface Then
                    MsgBox "«·Õ—ﬂÂ „— »ÿÂ »«·«⁄ „«œ« "
                Else
                    MsgBox "Can not edit.This process associated with approvals"
                End If
                Exit Sub
            End If
            
            FrmBuySearch.DealingForm = GridTransType.purchaseOrderApproved
            FrmBuySearch.Caption = "«·»ÕÀ ⁄‰ «Ê«„— «·‘—«¡ "
            FrmBuySearch.show vbModal

        Case 7

            If DoPremis(Do_Print, Me.Name, True) = False Then
                Exit Sub
            End If

            PrintReport
        Case 12
            print_report
         
        Case 11
            ShowGL_cc TxtNoteSerial.text, , 200
        
        Case 8
            On Error GoTo ErrTrap

            If XPTxtBillID.text <> "" Then
                Set SaleReport = New ClsSaleReport
                SaleReport.ShowPrice XPTxtBillID.text, 6, DcboEmp.text, val(DBCboClientName.BoundText)
            End If

            '        PrintReport1 (Txt_order_no.text)
        Case 6
            Unload Me
            
    End Select

    Exit Sub
ErrTrap:
End Sub

Function PrintReport1(order_no As String)
    
    Dim MySQL As String
    Dim RsData As New ADODB.Recordset
    Dim xApp As New CRAXDRT.Application
    Dim xReport As CRAXDRT.Report
    Dim CViewer As ClsReportViewer
    Dim StrReportTitle As String
    Dim StrFileName As String
    Dim Msg As String

    MySQL = "Select * From QRY_items_orders_data where order_no='" & order_no & "'"

    If SystemOptions.UserInterface = ArabicInterface Then
        StrFileName = App.path & "\Reports\" & "Order_status.rpt"
    Else
        StrFileName = App.path & "\Reports\" & "Order_status.rpt"
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
        MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
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
        'xReport.ParameterFields(2).AddCurrentValue RPTComment_Arabic
        xReport.ParameterFields(4).AddCurrentValue get_branch_name(val(my_branch))
        StrReportTitle = "Order status" '& StrAccountName
 
    Else
 
        xReport.ParameterFields(1).AddCurrentValue cCompanyInfo.ArabCompanyName ' RPTCompany_Name_Eng
        'xReport.ParameterFields(2).AddCurrentValue RPTComment_Eng
        xReport.ParameterFields(4).AddCurrentValue get_branch_name(val(my_branch))
        StrReportTitle = "Order status"
 
    End If

    xReport.ParameterFields(3).AddCurrentValue user_name
    xReport.reporttitle = StrReportTitle
    xReport.EnableParameterPrompting = False
    xReport.ApplicationName = App.Title
    xReport.ReportAuthor = App.Title
    Set CViewer = New ClsReportViewer
    CViewer.FireReport xReport, WindowTarget, ""

    RsData.Close
    Set RsData = Nothing
    Screen.MousePointer = vbDefault

End Function

Private Sub CmdConvert_Click()
    Dim RowNum As Integer
    Dim Frm As Form
    On Error GoTo ErrTrap
    Screen.MousePointer = vbArrowHourglass

    If Me.CboPriceType.ListIndex = 0 Then
        Set Frm = New frmsalebill
    ElseIf Me.CboPriceType.ListIndex = 1 Then
        Set Frm = New FrmBillBuy
    End If

    With Frm
        .Convert
        '    .XPTxtBillID.Text = XPTxtBillID.Text
        .XPDtbBill.value = XPDtbBill.value
        .DBCboClientName.BoundText = DBCboClientName.BoundText
        .DCboStoreName.BoundText = DCboStoreName.BoundText
        .Dccurrency.BoundText = Me.Dccurrency.BoundText

        For RowNum = 1 To FG.rows - 1

            If .FG.TextMatrix(.FG.rows - 1, .FG.ColIndex("Code")) <> "" Then
                .FG.rows = .FG.rows + 1
            End If

            .FG.TextMatrix(.FG.rows - 1, .FG.ColIndex("Name")) = IIf(FG.TextMatrix(RowNum, FG.ColIndex("Name")) = "", "", FG.TextMatrix(RowNum, FG.ColIndex("Name")))
            .FG.TextMatrix(.FG.rows - 1, .FG.ColIndex("Code")) = IIf(FG.TextMatrix(RowNum, FG.ColIndex("Code")) = "", "", FG.TextMatrix(RowNum, FG.ColIndex("Code")))
            .FG.TextMatrix(.FG.rows - 1, .FG.ColIndex("ItemCase")) = IIf(FG.TextMatrix(RowNum, FG.ColIndex("ItemCase")) = "", "", FG.TextMatrix(RowNum, FG.ColIndex("ItemCase")))
            .FG.TextMatrix(.FG.rows - 1, .FG.ColIndex("HaveSerial")) = IIf(FG.TextMatrix(RowNum, FG.ColIndex("HaveSerial")) = "", "", FG.TextMatrix(RowNum, FG.ColIndex("HaveSerial")))
            .FG.TextMatrix(.FG.rows - 1, .FG.ColIndex("Count")) = IIf(FG.TextMatrix(RowNum, FG.ColIndex("Count")) = "", "", FG.TextMatrix(RowNum, FG.ColIndex("Count")))
            .FG.TextMatrix(.FG.rows - 1, .FG.ColIndex("Price")) = IIf(FG.TextMatrix(RowNum, FG.ColIndex("Price")) = "", "", FG.TextMatrix(RowNum, FG.ColIndex("Price")))
            .FG.TextMatrix(.FG.rows - 1, .FG.ColIndex("DiscountType")) = IIf(FG.TextMatrix(RowNum, FG.ColIndex("DiscountType")) = "", "", FG.TextMatrix(RowNum, FG.ColIndex("DiscountType")))
            Dim StrSQL As String
            Dim RsUnit As New ADODB.Recordset
        
            StrSQL = "SELECT dbo.Transactions.Transaction_Type, dbo.Transaction_Details.UnitId, dbo.TblUnites.UnitName, dbo.Transactions.Transaction_Serial FROM dbo.Transaction_Details INNER JOIN dbo.Transactions ON dbo.Transaction_Details.Transaction_ID = dbo.Transactions.Transaction_ID INNER JOIN dbo.TblUnites ON dbo.Transaction_Details.UnitId = dbo.TblUnites.UnitID WHERE (dbo.Transactions.Transaction_Type = 6) AND (dbo.Transactions.Transaction_Serial = '" & TxtTransSerial & "')"
            Set RsUnit = New ADODB.Recordset
            RsUnit.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText
        
            .FG.cell(flexcpData, .FG.rows - 1, FG.ColIndex("UnitID")) = IIf(IsNull(RsUnit("UnitID")), "", (RsUnit("UnitID").value))
            .FG.TextMatrix(.FG.rows - 1, FG.ColIndex("UnitID")) = IIf(IsNull(RsUnit("UnitName")), "", (RsUnit("UnitName").value))
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
    SystemOptions.SysHelp.HHDisplayTopicID Me.hWnd
End Sub

Private Sub CMDSelectFile_Click()
CommonDialog1.ShowOpen
txtFileName.text = CommonDialog1.FileName

End Sub

Private Sub CmdTemplate_Click()
    Dim Frm  As FrmBuySearch
    On Error GoTo ErrTrap
    Set Frm = New FrmBuySearch

    With Frm
        .DealingForm = InsertTemplate
        .Caption = "«·⁄—Ê÷ «·Ã«Â“…"
        '    .MDIChild = True
        .BorderStyle = 0
        '  .MinButton = True
        .show vbModeless, mdifrmmain
        .Visible = True
    End With

    Exit Sub
ErrTrap:
End Sub

Private Sub DBCboClientName_Change()
    TxtSearchCode.text = ""

    Dim DefaultSalesPersonId As Integer
    Dim Fullcode As String

    GetCustomersDetail val(DBCboClientName.BoundText), DefaultSalesPersonId, Fullcode

    TxtSearchCode.text = Fullcode

    If Me.TxtModFlg.text = "N" Or Me.TxtModFlg.text = "E" Then
 
        GetCustomersDetail val(DBCboClientName.BoundText), DefaultSalesPersonId

        If Not DefaultSalesPersonId = 0 Then

            Me.DcboEmp.BoundText = DefaultSalesPersonId
        End If
        FillOrderGrid
        FillOrderGrid2
    End If
 
End Sub

Private Sub DBCboClientName_KeyUp(KeyCode As Integer, _
                                  Shift As Integer)

    If KeyCode = vbKeyF3 Then
         FrmCompanySearch.lblSearchtype.Caption = 801
          FrmCompanySearch.show vbModal
        
        
    End If
          
    If KeyCode = vbKeyF5 Then
        Dim Dcombos As ClsDataCombos
        
        Set Dcombos = New ClsDataCombos

       
            Dcombos.GetCustomersSuppliers 2, Me.DBCboClientName, True
      
    End If

End Sub
 
Private Sub DcboEmp_KeyUp(KeyCode As Integer, _
                          Shift As Integer)

    If KeyCode = vbKeyF5 Then
        Dim Dcombos As ClsDataCombos
        
        Set Dcombos = New ClsDataCombos
        Dcombos.GetSalesRepData Me.DcboEmp

    End If

End Sub

Private Sub DCboItemsCode_Change()
    '
    On Error Resume Next
    TxtQuantity = GetItemResultValue(val(DCboItemsCode.BoundText))
End Sub
Function GetItemResultValue(itmid) As Double
    Dim s As String
    s = ""
    s = s & "SELECT  Top 1       TblItemShowDitailses.ResultValue "
    s = s & "FROM dbo.TblItemShows "
    s = s & "    JOIN TblItemShowDitailses "
    s = s & "        ON ID2 = TblItemShows.ID "
    s = s & "           AND TblItemShowDitailses.TransType = TblItemShows.TransType "
    s = s & " JOIN dbo.TblItems ON TblItemShowDitailses.ItemID = dbo.TblItems.ItemID "
    s = s & "WHERE '" & Format(XPDtbBill.value, "yyyy-MM-dd") & "' "
    s = s & "      BETWEEN StartSDate AND EndDate "
    s = s & "      AND TblItemShows.TransType = 1 "
  
'    If bycode = 1 Then
'        s = s & "  And  (  TblItems.code = '" & itmid & "'  OR dbo.TblItems.Fullcode = '" & itmid & "' ) "
'    Else
        s = s & "AND  TblItemShowDitailses.ItemID = " & itmid
    'End If
    s = s & "  ORDER BY TblItemShowDitailses.ID DESC  "
    Dim mrs As New ADODB.Recordset
    mrs.Open s, Cn, adOpenForwardOnly, adLockReadOnly
    If mrs.EOF Then
        GetItemResultValue = 0
        Exit Function
    End If
    GetItemResultValue = val(mrs!ResultValue & "")
    mrs.Close
    Exit Function
End Function
Private Sub DCboItemsCode_KeyUp(KeyCode As Integer, _
                                Shift As Integer)

    If KeyCode = vbKeyF9 Then
                    
        FrmSearchSerial.XPTxtCode.text = DCboItemsCode.text
        FrmSearchSerial.show
        FrmSearchSerial.Cmd_Click (0)
                    
    End If

  '  If KeyCode = vbKeyF3 Then
  '
  '      Load FrmItemSearch
  '      FrmItemSearch.RetrunType = 22
  '      FrmItemSearch.show vbModal
  '  End If


  If KeyCode = vbKeyF3 Then
        
        Load FrmItemSearch
        FrmItemSearch.RetrunType = 30719
        FrmItemSearch.show vbModal
    End If


End Sub

Private Sub DCboItemsName_KeyUp(KeyCode As Integer, _
                                Shift As Integer)

    If KeyCode = vbKeyF9 Then
                    
           FrmSearchSerial.XPTxtCode.text = DCboItemsCode.text
          FrmSearchSerial.show
           FrmSearchSerial.Cmd_Click (0)
                    
    End If

    If KeyCode = vbKeyF3 Then
        
        Load FrmItemSearch
        FrmItemSearch.RetrunType = 30719
        FrmItemSearch.show vbModal
    End If

End Sub

Private Sub DCboStoreName_Change()
 TxtStoreID.text = getStoreCoding(val(DCboStoreName.BoundText))
   If Me.TxtModFlg = "N" Or Me.TxtModFlg = "E" Then

                 If CheckStoreCoding(val(dcBranch.BoundText), 29) = True Then
                ' TxtNoteSerial.text = ""
                TxtNoteSerial1.text = ""
            
                 End If
     
    End If
End Sub

Private Sub DCboStoreName_Click(Area As Integer)
'DCboStoreName_Change
End Sub

Private Sub DCboStoreName_KeyUp(KeyCode As Integer, _
                                Shift As Integer)

    If KeyCode = vbKeyF5 Then
        Dim Dcombos As ClsDataCombos
        
        Set Dcombos = New ClsDataCombos
        Dcombos.GetStores Me.DCboStoreName

    End If
        
End Sub

Private Sub Dcbranch_Change()
  If Me.TxtModFlg = "N" Or Me.TxtModFlg = "E" Then
TxtNoteSerial1.text = ""
TxtNoteSerial.text = ""
End If



End Sub

Private Sub Dcbranch_Click(Area As Integer)
'TxtNoteSerial.text = ""
'TxtNoteSerial1.text = ""
End Sub

Private Sub dcBranch_KeyUp(KeyCode As Integer, _
                           Shift As Integer)

    If KeyCode = vbKeyF5 Then
        Dim Dcombos As ClsDataCombos
        
        Set Dcombos = New ClsDataCombos
        Dcombos.GetBranches dcBranch
    End If

End Sub

Private Sub DCCurrency_KeyUp(KeyCode As Integer, _
                             Shift As Integer)

    If KeyCode = vbKeyF5 Then
        Dim My_SQL As String
        My_SQL = " select id,code from currency"
 
        fill_combo Me.Dccurrency, My_SQL

    End If

End Sub


Private Sub dcopr2_Change()
dcopr2_Click (0)
End Sub

Private Sub dcopr2_Click(Area As Integer)
On Error Resume Next
 Dim Dcombos As ClsDataCombos
 Dim project_id As Integer
       Set Dcombos = New ClsDataCombos
  If dcproject1.BoundText <> "" Then
         'project_id = get_project_id(dcproject1.BoundText, "Material_account")
         If Me.dcopr2.BoundText <> "" Then
         Dcombos.GetProcessOfProjedt DcbProcess1, val(dcproject1.BoundText), , dcopr2.BoundText, 2
         End If
       
    End If
End Sub

Private Sub dcproject1_Change()
dcproject1_Click (0)
End Sub
Function fillterms1(project_id As Integer)
    Dim My_SQL As String
 
    My_SQL = " select oprid,des from dbo.projects_des where project_id=" & project_id

  
        fill_combo Me.dcopr2, My_SQL
        
   ' dcopr.ReFill
End Function
Private Sub dcproject1_Click(Area As Integer)
    On Error Resume Next
    If dcproject1.BoundText <> "" Then
       ' project_id1 = get_project_id(dcproject1.BoundText, "Material_account")
         fillterms1 (val(dcproject1.BoundText))
    End If
End Sub

Private Sub Ele_Click(index As Integer)

    Select Case index

        Case 6
            On Error GoTo ErrTrap
            '        If Me.WindowState = vbNormal Then
            '            Me.WindowState = vbMaximized
            '        Else
            '            Me.WindowState = vbNormal
            '        End If
    End Select

    Exit Sub
ErrTrap:
End Sub

Function showComm()
    Me.LblFinal = val(LblTotal.Caption) + val(TxtValueAdded.text)
   ' Me.LblFinal.Caption = Format(val(LblFinal.Caption), "#,###." & String(Abs(SystemOptions.Count_ACCOUNT_digit), "0"))
    Me.LblFinal2.Caption = Format(val(LblFinal.Caption), "#,###." & String(Abs(SystemOptions.Count_ACCOUNT_digit), "0"))
End Function
Private Sub FG_AfterEdit(ByVal row As Long, _
   ByVal Col As Long)
    Dim rs             As New ADODB.Recordset
    Dim StrSQL         As String
    Dim StrAccountType As String
    Dim StrComboList   As String
    Dim Msg            As String
    Dim LngRow         As Double
    Dim StrAccountCode As String
    Dim Rs1            As ADODB.Recordset

    Dim ClsAcc         As New ClsAccounts
  
    Dim sql            As String
    Set Rs1 = New ADODB.Recordset
    With FG

        Select Case .ColKey(Col)
            Case "project"
                StrAccountCode = .ComboData
                LngRow = .FindRow(StrAccountCode, .FixedRows, .ColIndex("projectid"), False, True)
                .TextMatrix(row, .ColIndex("projectid")) = StrAccountCode
                .TextMatrix(row, .ColIndex("operaid")) = 0
                .TextMatrix(row, .ColIndex("pandid")) = 0
                .TextMatrix(row, .ColIndex("pand")) = ""
                .TextMatrix(row, .ColIndex("opera")) = ""
            Case "pand"
                StrAccountCode = .ComboData
                LngRow = .FindRow(StrAccountCode, .FixedRows, .ColIndex("pandid"), False, True)
                .TextMatrix(row, .ColIndex("pandid")) = StrAccountCode
                .TextMatrix(row, .ColIndex("operaid")) = 0
                .TextMatrix(row, .ColIndex("opera")) = ""
            Case "opera"
                StrAccountCode = .ComboData
                LngRow = .FindRow(StrAccountCode, .FixedRows, .ColIndex("operaid"), False, True)
                .TextMatrix(row, .ColIndex("operaid")) = StrAccountCode
        End Select
    End With
    ' With FG
    '        Select Case .ColKey(Col)
    '            Case "countris"
    '                StrAccountCode = .ComboData
    '                LngRow = .FindRow(StrAccountCode, .FixedRows, .ColIndex("countrisid"), False, True)
    '                .TextMatrix(Row, .ColIndex("countrisid")) = StrAccountCode
    '         End Select
    ' End With
   If Col = FG.ColIndex("Code") Or Col = FG.ColIndex("Name") Then
   Dim mLngItemID
     mLngItemID = val(FG.ComboData)
        FG.TextMatrix(row, FG.ColIndex("Count")) = val(GetItemResultValue(mLngItemID))
    End If
    If Me.TxtModFlg <> "E" Then Exit Sub
    
  

    '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\LogFile Saving//////////////////////////////////////////////////////////////////////////
    If Col = FG.ColIndex("Code") Or Col = FG.ColIndex("Name") Then
        RegisterItemData Me.Name, Me.TxtModFlg, FG.cell(flexcpTextDisplay, row, FG.ColIndex("Code")), FG.cell(flexcpTextDisplay, row, FG.ColIndex("Name")), , , , , , , , , , , Me.TXT_order_no
    
    ElseIf Col = FG.ColIndex("UnitID") Then
        RegisterItemData Me.Name, Me.TxtModFlg, FG.cell(flexcpTextDisplay, row, FG.ColIndex("Code")), FG.cell(flexcpTextDisplay, row, FG.ColIndex("Name")), FG.cell(flexcpTextDisplay, row, FG.ColIndex("UnitID")), , , , , , , , , , Me.TXT_order_no
    ElseIf Col = FG.ColIndex("Count") Then
        RegisterItemData Me.Name, Me.TxtModFlg, FG.cell(flexcpTextDisplay, row, FG.ColIndex("Code")), FG.cell(flexcpTextDisplay, row, FG.ColIndex("Name")), , (FG.TextMatrix(row, FG.ColIndex("Count"))), , , , , , , , , Me.TXT_order_no
    ElseIf Col = FG.ColIndex("Price") Then
        RegisterItemData Me.Name, Me.TxtModFlg, FG.cell(flexcpTextDisplay, row, FG.ColIndex("Code")), FG.cell(flexcpTextDisplay, row, FG.ColIndex("Name")), , , (FG.TextMatrix(row, FG.ColIndex("Price"))), , , , , , , , Me.TXT_order_no
    ElseIf Col = FG.ColIndex("ColorID") Then
        RegisterItemData Me.Name, Me.TxtModFlg, FG.cell(flexcpTextDisplay, row, FG.ColIndex("Code")), FG.cell(flexcpTextDisplay, row, FG.ColIndex("Name")), , , , , FG.cell(flexcpTextDisplay, row, FG.ColIndex("ColorID")), , , , , , Me.TXT_order_no
    ElseIf Col = FG.ColIndex("ItemSize") Then
        RegisterItemData Me.Name, Me.TxtModFlg, FG.cell(flexcpTextDisplay, row, FG.ColIndex("Code")), FG.cell(flexcpTextDisplay, row, FG.ColIndex("Name")), , , , , , FG.cell(flexcpTextDisplay, row, FG.ColIndex("ItemSize")), , , , , Me.TXT_order_no
    ElseIf Col = FG.ColIndex("ClassId") Then
        RegisterItemData Me.Name, Me.TxtModFlg, FG.cell(flexcpTextDisplay, row, FG.ColIndex("Code")), FG.cell(flexcpTextDisplay, row, FG.ColIndex("Name")), , , , , , , FG.cell(flexcpTextDisplay, row, FG.ColIndex("ClassId")), , , , Me.TXT_order_no
    ElseIf Col = FG.ColIndex("DiscountType") Then
        RegisterItemData Me.Name, Me.TxtModFlg, FG.cell(flexcpTextDisplay, row, FG.ColIndex("Code")), FG.cell(flexcpTextDisplay, row, FG.ColIndex("Name")), , , , , , , , FG.cell(flexcpTextDisplay, row, FG.ColIndex("DiscountType")), , , Me.TXT_order_no
    ElseIf Col = FG.ColIndex("DiscountVal") Then
        RegisterItemData Me.Name, Me.TxtModFlg, FG.cell(flexcpTextDisplay, row, FG.ColIndex("Code")), FG.cell(flexcpTextDisplay, row, FG.ColIndex("Name")), , , , , , , , , FG.TextMatrix(row, FG.ColIndex("DiscountVal")), , Me.TXT_order_no

    End If

    '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\LogFile Saving//////////////////////////////////////////////////////////////////////////
    ReLineGrid
End Sub

Private Sub FG_CellButtonClick(ByVal row As Long, _
                               ByVal Col As Long)

    If Me.TxtModFlg.text = "N" Or Me.TxtModFlg.text = "E" Then
        '    FrmAddNewItem.Tag = "xx"
    '    FrmAddNewItem.DealingForm = ShowPrice
    '    FrmAddNewItem.show vbModal
    End If

End Sub

Private Sub fg_StartEdit(ByVal row As Long, ByVal Col As Long, Cancel As Boolean)
Dim rs As New ADODB.Recordset
    Dim StrSQL  As String
    Dim StrAccountType As String
    Dim StrComboList As String
    Dim Msg As String
Dim StrAccountCode As String

'With FG
'  Select Case .ColKey(Col)
'
'            Case "countris"
'
'                StrSQL = "select * from TblCountriesData"
'                rs.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText
'
'                    StrComboList = FG.BuildComboList(rs, "CountryName", "CountryID")
'
'                If StrComboList <> "" Then
'                    StrComboList = "|" & StrComboList
'                End If
'                 .ComboList = StrComboList
'            End Select
'
'End With
 
 
    With Me.FG

        Select Case .ColKey(Col)
      Case "project"
               StrSQL = "  SELECT     id, Project_name, Project_nameE"
               StrSQL = StrSQL & "     From dbo.Projects"
             If SystemOptions.UserInterface = ArabicInterface Then
              StrSQL = StrSQL & " Where (Not (Project_name Is Null)) and Project_name<>N'""'"
              Else
              StrSQL = StrSQL & " Where (Not (Project_nameE Is Null))and Project_nameE <>N'""'"
             End If
       Set rs = New ADODB.Recordset
                rs.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

                If SystemOptions.UserInterface = ArabicInterface Then
                    StrComboList = .BuildComboList(rs, "Project_name", "id")
                Else
                    StrComboList = .BuildComboList(rs, "Project_nameE", "id")
                End If

                If StrComboList <> "" Then
                    StrComboList = "|" & StrComboList
                End If

                .ComboList = StrComboList
                rs.Close
                Set rs = Nothing
         ''//////////////
              Case "pand"
               StrSQL = " SELECT     oprid, des"
               StrSQL = StrSQL & "          From dbo.projects_des"
               StrSQL = StrSQL & "          Where (project_id = " & val(.TextMatrix(row, .ColIndex("ProjectID"))) & " and project_id<>0)"
        
       Set rs = New ADODB.Recordset
                rs.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText
               StrComboList = .BuildComboList(rs, "des", "oprid")
                If StrComboList <> "" Then
                    StrComboList = "|" & StrComboList
                End If

                .ComboList = StrComboList
                rs.Close
                Set rs = Nothing
                

          Case "opera"
               StrSQL = "       SELECT     dbo.terms_operations.OPRIDD, dbo.TblProcessDEF.ProcessName, dbo.TblProcessDEF.ProcessNameE"
               StrSQL = StrSQL & "       FROM         dbo.terms_operations LEFT OUTER JOIN"
               StrSQL = StrSQL & "       dbo.TblProcessDEF ON dbo.terms_operations.OPRIDD = dbo.TblProcessDEF.TblProcessDEFID"
               StrSQL = StrSQL & "   Where (dbo.terms_operations.project_id = " & val(.TextMatrix(row, .ColIndex("ProjectID"))) & ") And (dbo.terms_operations.ProjectDes_ID = " & val(.TextMatrix(row, .ColIndex("PandID"))) & ")"
       Set rs = New ADODB.Recordset
                rs.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText
                If SystemOptions.UserInterface = ArabicInterface Then
                   StrComboList = .BuildComboList(rs, "ProcessName", "OPRIDD")
                Else
                   StrComboList = .BuildComboList(rs, "ProcessNameE", "OPRIDD")
                End If
                If StrComboList <> "" Then
                    StrComboList = "|" & StrComboList
                End If

                .ComboList = StrComboList
                rs.Close
                Set rs = Nothing
   End Select
   End With
ReLineGrid
End Sub
Private Sub ReLineGrid()
    Dim i As Integer
    Dim IntCounter As Integer
   ''''///
   Dim SUM As Double
   IntCounter = 0
   SUM = 0
     With FG

        For i = .FixedRows To .rows - 1

            If .TextMatrix(i, .ColIndex("Name")) <> "" Then
                IntCounter = IntCounter + 1
                .TextMatrix(i, .ColIndex("Ser")) = IntCounter
                            End If

        Next i

    End With
    End Sub
Private Sub Form_Activate()
    'XPTxtBillID.SetFocus
End Sub

 

Private Sub ImgFavorites_Click()
AddTofaforites Me.Name, Me.Caption, Me.Caption
End Sub

Private Sub ISButton1_Click()
    FrmLC.show
    FrmLC.Retrive Trim(Me.TxtLcNo.text)
    'Frame3.Visible = True
End Sub

Private Sub ISButton2_Click()
    On Error Resume Next
ShowAttachments TxtNoteSerial1, "060520152"
 
End Sub

Private Sub ISButton3_Click()
If ChAuto.value = vbChecked Then
FillGrid
ChAuto.value = vbUnchecked
Exit Sub
End If
End Sub
Sub FillGrid()
Dim i As Integer
Dim rs2 As ADODB.Recordset
Set rs2 = New ADODB.Recordset
Dim sql As String
Dim current_row As Integer
sql = " SELECT     dbo.TblMatrials.Pand, dbo.TblMatrials.ProjectID, dbo.TblMatrials.monthly, dbo.TblMatrials.catalogID, dbo.TblMatrials.OperCode, dbo.TblMatrials.priceapro, "
sql = sql & "                       dbo.TblMatrials.Quntapro, dbo.TblMatrials.Price, dbo.TblMatrials.[Count], dbo.TblMatrials.ItemID, dbo.TblItems.ItemName, dbo.TblItems.ItemNamee,"
sql = sql & "                       dbo.TblItems.fullcode , dbo.terms_operations.OPRIDD"
sql = sql & "  FROM         dbo.TblMatrials RIGHT OUTER JOIN"
sql = sql & "                       dbo.terms_operations ON dbo.TblMatrials.Opr = dbo.terms_operations.id LEFT OUTER JOIN"
sql = sql & "                       dbo.TblItems ON dbo.TblMatrials.ItemID = dbo.TblItems.ItemID"
sql = sql & "  Where (dbo.TblMatrials.ProjectID =" & val(dcproject1.BoundText) & ") "
If dcopr2.text <> "" And val(dcopr2.BoundText) <> 0 Then
sql = sql & " And dbo.TblMatrials.Pand = " & val(dcopr2.BoundText) & ""
End If
If DcbProcess1.text <> "" And val(DcbProcess1.BoundText) <> 0 Then
sql = sql & " And dbo.terms_operations.OPRIDD  =" & val(DcbProcess1.BoundText) & ""
End If
rs2.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
If rs2.RecordCount > 0 Then
    With FG
    rs2.MoveFirst
    If FG.rows = 2 And val(FG.TextMatrix(1, FG.ColIndex("Code"))) = 0 Then
    FG.rows = FG.rows - 1
    End If
     For i = 1 To rs2.RecordCount
    FG.rows = FG.rows + 1
        current_row = FG.rows - 1
    .TextMatrix(current_row, .ColIndex("operaid")) = DcbProcess1.BoundText
     .TextMatrix(current_row, .ColIndex("pandid")) = Me.dcopr2.BoundText
     .TextMatrix(current_row, .ColIndex("projectid")) = val(dcproject1.BoundText)
     .TextMatrix(current_row, .ColIndex("project")) = Me.dcproject1.text
     .TextMatrix(current_row, .ColIndex("pand")) = Me.dcopr2.text
     .TextMatrix(current_row, .ColIndex("opera")) = DcbProcess1.text
    ' .TextMatrix(i, .ColIndex("Count")) = IIf(IsNull(Rs2("Quntapro").value), 0, Rs2("Quntapro").value)
    ' .TextMatrix(i, .ColIndex("Price")) = IIf(IsNull(Rs2("priceapro").value), 0, Rs2("priceapro").value)
    ' .TextMatrix(i, .ColIndex("Valu")) = val(.TextMatrix(i, .ColIndex("Count"))) * val(.TextMatrix(i, .ColIndex("Price")))
    ' .TextMatrix(i, .ColIndex("opera")) = IIf(IsNull(Rs2("ItemID").value), 0, Rs2("ItemID").value)
    ' .TextMatrix(i, .ColIndex("Code")) = IIf(IsNull(Rs2("ItemID").value), "", Rs2("ItemID").value)
    ' .TextMatrix(i, .ColIndex("Name")) = IIf(IsNull(Rs2("ItemName").value), "", Rs2("ItemName").value)
    
     
     ' DCboItemsCode.BoundText = IIf(IsNull(Rs2("ItemID").value), 0, Rs2("ItemID").value)
      DCboItemsName.BoundText = IIf(IsNull(rs2("ItemID").value), "", rs2("ItemID").value)
       TxtQuantity.text = IIf(IsNull(rs2("Quntapro").value), 0, rs2("Quntapro").value)
      TxtPrice.text = IIf(IsNull(rs2("priceapro").value), 0, rs2("priceapro").value)
    NewGrid.CmdAddData_Click
    rs2.MoveNext
  Next i
  
    End With
 End If
End Sub
Private Sub ISButton4_Click()
If ChAuto.value = vbUnchecked Then
'FillGrid
'ChAuto.value = vbUnchecked
Exit Sub
End If

If Me.dcproject1.text = "" Then
If SystemOptions.UserInterface = ArabicInterface Then
MsgBox "Ì—ÃÏ «Œ Ì«— «·„‘—Ê⁄"
Else
MsgBox "Please Select Project"
End If
dcproject1.SetFocus
Exit Sub
End If
If DCboItemsName.text = "" Or val(DCboItemsName.BoundText) = 0 Then
If SystemOptions.UserInterface = ArabicInterface Then
MsgBox "Ì—ÃÏ «Œ Ì«— «·’‰› "
Else
MsgBox "Please Select Item"
End If
DCboItemsName.SetFocus
Exit Sub
End If
If val(TxtQuantity.text) = 0 Then
If SystemOptions.UserInterface = ArabicInterface Then
MsgBox "Ì—ÃÏ  ÕœÌœ «·ﬂ„Ì… "
Else
MsgBox "Please Eneter Qty"
End If
TxtQuantity.SetFocus
Exit Sub
End If
    Dim current_row As Integer

    If FG.rows = 1 Then
        FG.rows = FG.rows + 1
        current_row = 1
    Else
        
        FG.rows = FG.rows + 1
        current_row = FG.rows - 1
    End If

    With FG
    .TextMatrix(current_row, .ColIndex("operaid")) = DcbProcess1.BoundText
     .TextMatrix(current_row, .ColIndex("pandid")) = Me.dcopr2.BoundText
    .TextMatrix(current_row, .ColIndex("projectid")) = val(dcproject1.BoundText)
    .TextMatrix(current_row, .ColIndex("project")) = Me.dcproject1.text
        .TextMatrix(current_row, .ColIndex("pand")) = Me.dcopr2.text
    .TextMatrix(current_row, .ColIndex("opera")) = DcbProcess1.text
    End With
 NewGrid.CmdAddData_Click
End Sub

Private Sub Label10_Click()
    Frame3.Visible = False
End Sub
 
Private Sub Accredit_Click()
'    Dim sql As String
'    Dim BeginTrans As Boolean
'    'sql = "update  Transactions  set Posted=" & user_id & "  where Transaction_ID=" & Val(XPTxtBillID.text)
'    'Cn.Execute sql
'
'    Cn.BeginTrans
'    BeginTrans = True
'
'    If IsNull(rs("Posted")) Then
'        rs("Posted") = user_id
'        rs("PostedDate") = Time
'    Else
'        rs("Posted") = Null
'       rs("PostedDate") = Time
'    End If
'
'    rs.update
' If SystemOptions.UserInterface = ArabicInterface Then
'    Accredit.Caption = " „ «·«—”«· ··«⁄ „«œ"
'Else
'Accredit.Caption = "Sent To approval "
'End If
'
'    Cn.CommitTrans
'    BeginTrans = False
'FillApprovedTable
'  '  Retrive (val(XPTxtBillID.text))


Dim BeginTrans As Boolean
If val(XPTxtBillID.text) = 0 Then
     If SystemOptions.UserInterface = ArabicInterface Then
     MsgBox "«Õ›Ÿ «·”‰œ «Ê·«", vbCritical
     Else
     MsgBox "Save Doc First", vbCritical
     End If
      
      Exit Sub
      End If
 
 
    SendTopost Me.Name, "Transactions", "Transaction_ID", 0, val(dcBranch.BoundText), val(XPTxtBillID.text), TxtNoteSerial1.text
  rs.Resync
 If SystemOptions.UserInterface = ArabicInterface Then
    Accredit.Caption = " „ «·«—”«· ··«⁄ „«œ"
Else
    Accredit.Caption = "Sent To Approval "
End If
    Retrive (val(Me.XPTxtBillID.text))

End Sub
Function FillApprovedTable()

   Exit Function

SendTopost Me.Name, "Transactions", "Transaction_ID", 0, val(dcBranch.BoundText), val(XPTxtBillID.text), TxtNoteSerial1.text

 Dim RSApproval  As New ADODB.Recordset
   Set RSApproval = New ADODB.Recordset
   Dim currentdate As Date
   RSApproval.Open "[ApprovalData]", Cn, adOpenStatic, adLockOptimistic, adCmdTable


 Dim sql As String
  Dim Rs1 As New ADODB.Recordset
 Dim i As Integer
    sql = "SELECT     TOP 100 PERCENT dbo.TblApprovalDef.ScreenName, dbo.TblApprovalDefDetails.PlainMessageID AS levelo, dbo.TbllevelWorker.EmpID, "
  sql = sql & " dbo.TblApprovalDefDetails.id AS levelorder, dbo.TbllevelWorker.id AS currorder"
  sql = sql & " FROM         dbo.TblApprovalDef INNER JOIN"
  sql = sql & " dbo.TblApprovalDefDetails ON dbo.TblApprovalDef.id = dbo.TblApprovalDefDetails.lMessageDefID INNER JOIN"
  sql = sql & "  dbo.TbllevelWorker ON dbo.TblApprovalDefDetails.PlainMessageID = dbo.TbllevelWorker.LevelID"
sql = sql & " WHERE     (dbo.TblApprovalDef.ScreenName = N'" & Me.Name & "')"
sql = sql & " ORDER BY dbo.TblApprovalDefDetails.id, dbo.TbllevelWorker.id  "

    Rs1.Open sql, Cn, adOpenStatic, adLockReadOnly, adCmdText

    If Rs1.RecordCount > 0 Then
            currentdate = Now
            For i = 1 To Rs1.RecordCount
              RSApproval.AddNew
                RSApproval("ScreenName").value = Me.Name
                RSApproval("levelo").value = IIf(IsNull(Rs1("levelo").value), Null, Rs1("levelo").value)
               RSApproval("EmpID").value = IIf(IsNull(Rs1("EmpID").value), Null, Rs1("EmpID").value)
                RSApproval("levelorder").value = IIf(IsNull(Rs1("levelorder").value), Null, Rs1("levelorder").value)
                 RSApproval("currorder").value = IIf(IsNull(Rs1("currorder").value), Null, Rs1("currorder").value)
                  RSApproval("Transaction_ID").value = val(XPTxtBillID.text)
                  RSApproval("NoteSerial").value = TxtNoteSerial1.text
                RSApproval("Transaction_Date").value = Date
                
                  RSApproval("ExpectedtimeTime").value = DateAdd("N", GetTimeforTransaction(Me.Name), currentdate)
               RSApproval("SendTime").value = currentdate

                 If i = 1 Then
                        RSApproval("Currcursor").value = 1
                         RSApproval("FromUser").value = user_name
                End If
                
                RSApproval.update
                Rs1.MoveNext
            Next i

    End If
    
    

End Function
Public Sub RetriveOrder(Optional order_no As String = "", _
                        Optional Transaction_Type As Integer = 0)
    Dim RsDetails As New ADODB.Recordset
    Dim StrSQL As String
    Dim RsNotes As New ADODB.Recordset
    Dim RsTemp As ADODB.Recordset
    Dim rs As ADODB.Recordset
    Dim Num As Long
    On Error GoTo ErrTrap
    FG.Clear flexClearScrollable, flexClearEverything
    FG.rows = 2
    FG.Clear flexClearScrollable, flexClearEverything
    FG.Refresh

   
        StrSQL = "Select * from transactions where  Transaction_Type=43 and Order_no='" & order_no & "'"
 

    Set rs = New ADODB.Recordset
    rs.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

    If rs.RecordCount < 1 Then
 
        Exit Sub
    Else
        DBCboClientName.BoundText = IIf(IsNull(rs("CusID").value), "", rs("CusID").value)
        Me.Dccurrency.BoundText = IIf(IsNull(rs("Currency_id").value), "", rs("Currency_id").value)
        Me.DCboStoreName.BoundText = IIf(IsNull(rs("storeid").value), "", rs("storeid").value)
        Me.dcBranch.BoundText = IIf(IsNull(rs("Branchid").value), "", rs("Branchid").value)

        'txtContainerNo_Currency_rate.text = IIf(IsNull(rs("Currency_rate").value), 1, (rs("Currency_rate").value))
    End If

    If rs.EOF Or rs.BOF Then
        Exit Sub
    End If

    Screen.MousePointer = vbArrowHourglass

    StrSQL = "SELECT TblItems.HaveSerial, * FROM TblItems INNER JOIN Transaction_Details " & "ON TblItems.ItemID = Transaction_Details.Item_ID INNER JOIN dbo.TblUnites ON dbo.Transaction_Details.UnitID = dbo.TblUnites.UnitID"
    StrSQL = StrSQL + " where Transaction_ID=" & val(rs("Transaction_ID").value)

    RsDetails.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText
    XPTxtSum.text = ""

    If Not (RsDetails.EOF Or RsDetails.BOF) Then
        FG.rows = RsDetails.RecordCount + 1

        For Num = 1 To RsDetails.RecordCount
            FG.TextMatrix(Num, FG.ColIndex("Code")) = IIf(IsNull(RsDetails("Item_ID")), "", (RsDetails("Item_ID").value))
            FG.TextMatrix(Num, FG.ColIndex("Name")) = IIf(IsNull(RsDetails("Item_ID")), "", Trim(RsDetails("Item_ID").value))
            FG.TextMatrix(Num, FG.ColIndex("Count")) = IIf(IsNull(RsDetails("showqty")), "", (RsDetails("showqty").value))

            'FG.TextMatrix(Num, FG.ColIndex("Price")) = IIf(IsNull(RsDetails("Price")), "", (RsDetails("Price").value))
            If Transaction_Type = 0 Then
                FG.TextMatrix(Num, FG.ColIndex("Price")) = IIf(IsNull(RsDetails("ShowPrice")), 0, (RsDetails("ShowPrice").value)) ' GET_COST_PRICE_FOR_PRODUCT_ITEM(Val(FG.TextMatrix(Num, FG.ColIndex("Code"))))
            End If
      
            '  FG.TextMatrix(Num, FG.ColIndex("Expenses")) = IIf(IsNull(RsDetails("Lineexpenses")), "", (RsDetails("Lineexpenses").value))
         
            FG.TextMatrix(Num, FG.ColIndex("ItemCase")) = IIf(IsNull(RsDetails("ItemCase")), "", (RsDetails("ItemCase").value))
            FG.TextMatrix(Num, FG.ColIndex("DiscountType")) = IIf(IsNull(RsDetails("ItemDiscountType")), "", (RsDetails("ItemDiscountType").value))
            FG.TextMatrix(Num, FG.ColIndex("DiscountVal")) = IIf(IsNull(RsDetails("ItemDiscount")), "", (RsDetails("ItemDiscount").value))
            FG.TextMatrix(Num, FG.ColIndex("ColorID")) = IIf(IsNull(RsDetails("ColorID")), 1, (RsDetails("ColorID").value))
            FG.TextMatrix(Num, FG.ColIndex("ItemSize")) = IIf(IsNull(RsDetails("ItemSize")), 1, (RsDetails("ItemSize").value))
            FG.TextMatrix(Num, FG.ColIndex("ClassID")) = IIf(IsNull(RsDetails("ClassID")), 1, (RsDetails("ClassID").value))
            FG.TextMatrix(Num, FG.ColIndex("ItemType")) = IIf(IsNull(RsDetails("ItemType")), 0, (RsDetails("ItemType").value))
         
            If RsDetails("HaveSerial") = True Then
                FG.TextMatrix(Num, FG.ColIndex("HaveSerial")) = True
            End If
        
            FG.cell(flexcpData, Num, FG.ColIndex("UnitID")) = IIf(IsNull(RsDetails("UnitID")), "", (RsDetails("UnitID").value))
            FG.TextMatrix(Num, FG.ColIndex("UnitID")) = IIf(IsNull(RsDetails("UnitName")), "", (RsDetails("UnitName").value))
        
            RsDetails.MoveNext
            Debug.Print Num

            If FG.rows > 10 Then
                If Num = 8 Then FG.Refresh
            End If

        Next Num

    End If

    TxtFillData.text = "F"
    Screen.MousePointer = vbDefault
    XPTxtCurrent.Caption = rs.AbsolutePosition
    XPTxtCount.Caption = rs.RecordCount
    Exit Sub
ErrTrap:
    Screen.MousePointer = vbDefault
End Sub
Sub RetriveoOrderPO6(Optional TransID As Integer = 0, Optional Notserial As String = "", Optional Transaction_Type As Integer)
Dim StrSQL As String
Dim RsDetails As ADODB.Recordset
Dim Rs1 As ADODB.Recordset
Set RsDetails = New ADODB.Recordset
Dim Num As Integer
   FG.Clear flexClearScrollable, flexClearEverything
    FG.rows = 2
    FG.Clear flexClearScrollable, flexClearEverything
    FG.Refresh


If TransID <> 0 Then
StrSQL = "SELECT * FROM Transactions WHERE Transaction_ID=" & TransID
Else
StrSQL = "SELECT * FROM Transactions WHERE NoteSerial1='" & Notserial & " '"

End If
 StrSQL = StrSQL + " and Transaction_Type =" & Transaction_Type & "  "
    Set Rs1 = New ADODB.Recordset
    Rs1.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText
    If Rs1.RecordCount > 0 Then
    TxtBillComment.text = IIf(IsNull(Rs1("TransactionComment")), "", (Rs1("TransactionComment").value))
    DCboStoreName.BoundText = IIf(IsNull(Rs1("StoreID")), 0, (Rs1("StoreID").value))
     DBCboClientName.BoundText = IIf(IsNull(Rs1("StoreID")), 0, (Rs1("StoreID").value))
'     DBCboClientName.BoundText = IIf(IsNull(Rs1("CusID")), 0, (Rs1("CusID").value))
     DBCboClientName.BoundText = IIf(IsNull(Rs1("CusID")), 0, (Rs1("CusID").value))
     DcbDetpartment.BoundText = IIf(IsNull(Rs1("DepartementID")), 0, (Rs1("DepartementID").value))
    
   FG.Clear flexClearScrollable, flexClearEverything
    FG.rows = 2
    FG.Clear flexClearScrollable, flexClearEverything
    FG.Refresh
    StrSQL = " SELECT   dbo.Transaction_Details.unitid as itmemunitid,   dbo.TblItems.HaveSerial AS Expr1, *"
    StrSQL = StrSQL + "  FROM         dbo.TblItems INNER JOIN"
    StrSQL = StrSQL + "                   dbo.Transaction_Details ON dbo.TblItems.ItemID = dbo.Transaction_Details.Item_ID INNER JOIN"
    StrSQL = StrSQL + "                  dbo.TblUnites ON dbo.Transaction_Details.UnitId = dbo.TblUnites.UnitID LEFT OUTER JOIN"
    StrSQL = StrSQL + "                   dbo.TblProcessDEF ON dbo.Transaction_Details.Oper_ID = dbo.TblProcessDEF.TblProcessDEFID LEFT OUTER JOIN"
     StrSQL = StrSQL + "                  dbo.projects_des ON dbo.Transaction_Details.Pand_ID = dbo.projects_des.oprid LEFT OUTER JOIN"
     StrSQL = StrSQL + "                  dbo.projects ON dbo.Transaction_Details.project_ID1 = dbo.projects.id"

    StrSQL = StrSQL + " where Transaction_ID=" & Rs1("Transaction_ID").value
    StrSQL = StrSQL + " order by Transaction_Details.id "
    
'order by id
    RsDetails.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText
    
    If Not (RsDetails.EOF Or RsDetails.BOF) Then
        FG.rows = RsDetails.RecordCount + 1

        For Num = 1 To RsDetails.RecordCount
        ''//
       '  Fg.TextMatrix(Num, Fg.ColIndex("projectid")) = IIf(IsNull(RsDetails("project_ID1")), "", (RsDetails("project_ID1").value))
       '  Fg.TextMatrix(Num, Fg.ColIndex("pandid")) = IIf(IsNull(RsDetails("Pand_ID")), "", (RsDetails("Pand_ID").value))
       '  Fg.TextMatrix(Num, Fg.ColIndex("operaid")) = IIf(IsNull(RsDetails("Oper_ID")), "", (RsDetails("Oper_ID").value))
       '  Fg.TextMatrix(Num, Fg.ColIndex("pand")) = IIf(IsNull(RsDetails("des")), "", (RsDetails("des").value))
       '  If SystemOptions.UserInterface = ArabicInterface Then
       '  Fg.TextMatrix(Num, Fg.ColIndex("project")) = IIf(IsNull(RsDetails("Project_name")), "", (RsDetails("Project_name").value))
       '  Fg.TextMatrix(Num, Fg.ColIndex("opera")) = IIf(IsNull(RsDetails("ProcessName")), "", (RsDetails("ProcessName").value))
       '  Else
       '  Fg.TextMatrix(Num, Fg.ColIndex("project")) = IIf(IsNull(RsDetails("Project_nameE")), "", (RsDetails("Project_nameE").value))
       '  Fg.TextMatrix(Num, Fg.ColIndex("opera")) = IIf(IsNull(RsDetails("ProcessNameE")), "", (RsDetails("ProcessNameE").value))
       '  End If
        ''//
    
        
            FG.TextMatrix(Num, FG.ColIndex("Code")) = IIf(IsNull(RsDetails("Item_ID")), "", (RsDetails("Item_ID").value))
            FG.TextMatrix(Num, FG.ColIndex("Name")) = IIf(IsNull(RsDetails("Item_ID")), "", Trim(RsDetails("Item_ID").value))
           ' Fg.TextMatrix(Num, Fg.ColIndex("Count")) = IIf(IsNull(RsDetails("showqty")), "", (RsDetails("showqty").value))
            
         If SystemOptions.poWithatotalQty = False Then
             FG.TextMatrix(Num, FG.ColIndex("Count")) = IIf(IsNull(RsDetails("showqty")), 0, (RsDetails("showqty").value)) - IIf(IsNull(RsDetails("ItemBalance")), 0, (RsDetails("ItemBalance").value))
          Else
          FG.TextMatrix(Num, FG.ColIndex("Count")) = IIf(IsNull(RsDetails("showqty")), 0, (RsDetails("showqty").value))
          End If
            
            
            
            FG.TextMatrix(Num, FG.ColIndex("Price")) = IIf(IsNull(RsDetails("showPrice")), 0, (RsDetails("showPrice").value))
        
            FG.TextMatrix(Num, FG.ColIndex("ItemCase")) = IIf(IsNull(RsDetails("ItemCase")), "", (RsDetails("ItemCase").value))
            FG.TextMatrix(Num, FG.ColIndex("DiscountType")) = IIf(IsNull(RsDetails("ItemDiscountType")), "", (RsDetails("ItemDiscountType").value))
            FG.TextMatrix(Num, FG.ColIndex("DiscountVal")) = IIf(IsNull(RsDetails("ItemDiscount")), "", (RsDetails("ItemDiscount").value))
        
            If RsDetails("HaveSerial") = True Then
                FG.TextMatrix(Num, FG.ColIndex("HaveSerial")) = True
            End If
        
            FG.cell(flexcpData, Num, FG.ColIndex("UnitID")) = IIf(IsNull(RsDetails("itmemunitid")), "", (RsDetails("itmemunitid").value))
             FG.TextMatrix(Num, FG.ColIndex("UnitID")) = IIf(IsNull(RsDetails("UnitName")), "", (RsDetails("UnitName").value))
            FG.TextMatrix(Num, FG.ColIndex("ColorID")) = IIf(IsNull(RsDetails("ColorID")), 1, (RsDetails("ColorID").value))
            FG.TextMatrix(Num, FG.ColIndex("ItemSize")) = IIf(IsNull(RsDetails("ItemSize")), 1, (RsDetails("ItemSize").value))
            FG.TextMatrix(Num, FG.ColIndex("ClassID")) = IIf(IsNull(RsDetails("ClassID")), 1, (RsDetails("ClassID").value))
        

 ' FG.TextMatrix(Num, FG.ColIndex("Price")) = ModItemCostPrice.GetCostItemPrice(FG.TextMatrix(Num, FG.ColIndex("Code")), 0, , , SystemOptions.SysMainStockCostMethod, , , XPDtbBill.value, val(Me.XPTxtBillID), val(FG.Cell(flexcpData, Num, FG.ColIndex("UnitID"))))
        



     '   FG.TextMatrix(Num, FG.ColIndex("RequestLimit")) = IIf(IsNull(RsDetails("RequestLimit")), 0, (RsDetails("RequestLimit").value))
     '   FG.TextMatrix(Num, FG.ColIndex("LastPurchaseDate")) = IIf(IsNull(RsDetails("LastPurchaseDate")), "", (RsDetails("LastPurchaseDate").value))
        FG.TextMatrix(Num, FG.ColIndex("LastPurchasePrice")) = IIf(IsNull(RsDetails("LastPurchasePrice")), 0, (RsDetails("LastPurchasePrice").value))
    '    FG.TextMatrix(Num, FG.ColIndex("LastPurchaseqty")) = IIf(IsNull(RsDetails("LastPurchaseqty")), 0, (RsDetails("LastPurchaseqty").value))
   '     FG.TextMatrix(Num, FG.ColIndex("AverageIssue")) = IIf(IsNull(RsDetails("AverageIssue")), 0, (RsDetails("AverageIssue").value))
        
            RsDetails.MoveNext
            Debug.Print Num

            If FG.rows > 10 Then
                If Num = 8 Then FG.Refresh
            End If

        Next Num
End If
    End If
End Sub

Private Sub TxtPO6_Change()
Dim transactiontype As Integer
If Me.TxtModFlg.text <> "R" And Me.TxtModFlg.text <> "" Then

        If CBoBasedON.ListIndex = 0 Then
               transactiontype = 0
        ElseIf CBoBasedON.ListIndex = 1 Then
        transactiontype = 38
        ElseIf CBoBasedON.ListIndex = 2 Then
        transactiontype = 47
           ElseIf CBoBasedON.ListIndex = 3 Then
        transactiontype = 6
        
        End If
  RetriveoOrderPO6 , Me.TxtPO6.text, transactiontype
  NewGrid.Calculate 1, , , True
  
End If



End Sub

Private Sub TxtPO6_KeyUp(KeyCode As Integer, Shift As Integer)



If Me.TxtModFlg.text <> "R" Then

If CBoBasedON.ListIndex = 1 Then
                If KeyCode = vbKeyF3 Then
                FrmBuySearch.DealingForm = GridTransType.internalorder
                  FrmBuySearch.index = 15
                    FrmBuySearch.Caption = "«·»ÕÀ ⁄‰ ÿ·»«   œ«Œ·Ì…"
                   FrmBuySearch.show vbModal
                   End If
       
 ElseIf CBoBasedON.ListIndex = 2 Then
 
                If KeyCode = vbKeyF3 Then
               
               FrmBuySearch.DealingForm = GridTransType.purchaserequest
                  FrmBuySearch.index = 16
                    FrmBuySearch.Caption = "«·»ÕÀ ⁄‰  ÿ·»«  «·‘—«¡"
                   FrmBuySearch.show vbModal
               
                   End If
                   
  ElseIf CBoBasedON.ListIndex = 3 Then
  
    If KeyCode = vbKeyF3 Then
            Dim transactionName As String
                      If SystemOptions.UserInterface = ArabicInterface Then
                          transactionName = "»ÕÀ ⁄‰ «Ê«„— «·»Ì⁄"
                        Else
                        transactionName = "Search  Sales Order"
                        End If
            

    Order_no_search.show
         Order_no_search.RetrunType = 20
        Order_no_search.Label1(2).Caption = transactionName
        Order_no_search.lblSpecificsearch = 6
        
    End If
    
 
 End If
 
 
       End If
       
       
       
End Sub

Private Sub TxtPONo_Change()
  If Me.TxtModFlg <> "R" And Me.TxtModFlg <> "" Then
        RetriveOrder Me.TxtPONo
    End If
End Sub

Private Sub TxtPONo_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF3 Then
        Order_no_search4.show
        Order_no_search4.RetrunType = 43

        If val(Me.DBCboClientName.BoundText) <> 2 Then
        
            Order_no_search4.DBCboClientName.BoundText = Me.DBCboClientName.BoundText
        End If
    End If
End Sub

Private Sub TxtSearchCode_KeyPress(KeyAscii As Integer)
    Dim CUSTID As Integer

    If KeyAscii = vbKeyReturn Then
        GetCustomersDetail CUSTID, , TxtSearchCode.text, 1
        DBCboClientName.BoundText = CUSTID
    End If

End Sub

Private Sub TxtFillData_Change()

    If TxtFillData.text = "F" Then
        NewGrid.Calculate 1, , , True
    End If

End Sub

Private Sub TxtLcNo_KeyUp(KeyCode As Integer, _
                          Shift As Integer)

    If KeyCode = vbKeyF3 Then
        Order_no_search3.show
        Order_no_search3.RetrunType = 1
         
    End If
        
End Sub

Private Sub TxtStoreID_KeyDown(KeyCode As Integer, Shift As Integer)
 Dim StoreID As Integer

    If KeyCode = vbKeyReturn Then
    StoreID = getStoreInformatin(TxtStoreID)
        DCboStoreName.BoundText = StoreID
    End If
End Sub

Private Sub TxtValueAdded_Change()
RelinVatGrid
End Sub

Private Sub VatGrid_Click()
RelinVatGrid
End Sub

Private Sub VSFlexGrid1_Click()
    With FG
        .Clear flexClearScrollable, flexClearEverything
        .rows = 1
       
    End With
 
    fillOrders (1)
    
End Sub

Private Sub XPBtnMove_Click(index As Integer)
    On Error GoTo ErrTrap

    If Me.TxtModFlg.text = "N" Then
        clear_all Me
        Me.TxtModFlg.text = "R"
        XPBtnMove_Click (1)
    End If

    Select Case index

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
            '        Cmd_Click (0)
        Else
            '        SendKeys "{TAB}"
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
        '    If Cmd(3).Enabled = False Then Exit Sub
        '    Cmd_Click (3)
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
        
        End If
    End If

    If KeyCode = vbKeyF5 Then
        If TxtModFlg.text = "N" Or TxtModFlg.text = "E" Then
            XPBtnNewClients_Click
        End If
    End If

    If KeyCode = vbKeyF3 Then
        If TxtModFlg.text = "N" Or TxtModFlg.text = "E" Then
       
        End If
    End If

    If Shift = 2 Then
        If KeyCode = vbKeySpace Then
            If TxtModFlg.text = "N" Or TxtModFlg.text = "E" Then
            
            End If
        End If
    End If

    If Shift = 2 Then
        If KeyCode = vbKeyX Then
            Cmd_Click (6)
        End If
    End If

    Exit Sub
ErrTrap:
End Sub

Private Sub LblDiscountsTotal_Change()
    LblDiscountsTotalView.Caption = Format(val(LblDiscountsTotal.Caption), "#,###." & String(Abs(SystemOptions.SysDefCurrencyForamt), "0"))
End Sub

Private Sub LblTotal_Change()
    LblTotalView.Caption = Format(val(LblTotal.Caption), "#,###." & String(Abs(SystemOptions.SysDefCurrencyForamt), "0"))
    showComm
End Sub
Sub RelinVatGrid()
 'Salim1503
Dim i As Integer
Dim k As Integer
Dim SmValu As Double
SmValu = 0
With VatGrid
For i = 1 To .rows - 1
If .cell(flexcpChecked, i, .ColIndex("select")) = flexChecked Then
For k = FG.FixedRows To FG.rows - 1
If k = i And val(FG.TextMatrix(k, FG.ColIndex("Code"))) = val(.TextMatrix(i, .ColIndex("ItemID"))) And val(FG.TextMatrix(k, FG.ColIndex("Valu"))) = val(.TextMatrix(i, .ColIndex("Valu"))) Then
FG.TextMatrix(k, FG.ColIndex("Vat")) = val(.TextMatrix(i, .ColIndex("Vat")))
FG.TextMatrix(k, FG.ColIndex("Vatyo")) = val(.TextMatrix(i, .ColIndex("Vatyo")))
End If
Next k
SmValu = SmValu + val(.TextMatrix(i, .ColIndex("Vat")))
Else
For k = FG.FixedRows To FG.rows - 1
If k = i And val(FG.TextMatrix(k, FG.ColIndex("Code"))) = val(.TextMatrix(i, .ColIndex("ItemID"))) And val(FG.TextMatrix(k, FG.ColIndex("Valu"))) = val(.TextMatrix(i, .ColIndex("Valu"))) Then
If FG.ColIndex("Vat") <> -1 Then 'salim1503
FG.TextMatrix(k, FG.ColIndex("Vat")) = 0
End If
If FG.ColIndex("Vatyo") <> -1 Then 'salim1503
FG.TextMatrix(k, FG.ColIndex("Vatyo")) = 0
End If
If FG.ColIndex("TypeVAT") <> -1 Then 'salim1503
FG.TextMatrix(k, FG.ColIndex("TypeVAT")) = 0
End If
End If
Next k
End If
Next i
End With
TxtValueAdded.text = SmValu
LblVat.Caption = SmValu
showComm

End Sub
Private Sub ChecVAT_Click()
  Dim i As Integer
If Me.TxtModFlg.text <> "R" Then
    If ChecVAT.value = vbChecked Then

        With Me.VatGrid
 
            For i = 1 To .rows - 1
        
                .TextMatrix(i, .ColIndex("select")) = True
            Next i

        End With

    Else

        With Me.VatGrid

            For i = 1 To .rows - 1
        
                .TextMatrix(i, .ColIndex("select")) = False
            Next i

        End With

    End If
    RelinVatGrid
    End If
End Sub
Private Sub Form_Load()
    Dim RsClients As New ADODB.Recordset
    Dim StrSQL    As String
    Dim Num       As Integer
    Dim StrList   As String
    Dim BGround   As New ClsBackGroundPic
    Dim RsNote    As New ADODB.Recordset
    Dim ShowTax   As Boolean
    Dim Dcombos   As ClsDataCombos
    Dim My_SQL    As String
    On Error GoTo ErrTrap
    
    '  If SystemOptions.UserInterface = ArabicInterface Then
    '                FG.ColComboList(FG.ColIndex("Shipping")) = "#1;  «·„‘ —Ì|#2; «·»«∆⁄"
    '            ElseIf SystemOptions.UserInterface = EnglishInterface Then
    '               FG.ColComboList(FG.ColIndex("Shipping")) = "#1;Buyer |#2;Seller "
    '            End If
            
    ' If GeneralPriceType = 0 Then
    If SystemOptions.POMustentryAndBillMustEntry = True Then
        TxtPO6.locked = True
    End If

    ScreenNameArabic = "  √Ê«„— «·‘—«¡ "
    ScreenNameEnglish = "Purchase  Order "
    CurrentTransactionType = 29
  
    With Me.CBoBasedON
        If SystemOptions.UserInterface = ArabicInterface Then
            .Clear
            .AddItem "»·«"
            .AddItem "ÿ·» œ«Œ·Ì"
            .AddItem " ÿ·» ‘—«¡"
            .AddItem " «„—  »Ì⁄"
        Else
            .Clear
            .AddItem ("With out")
            .AddItem ("Internal Order")
            .AddItem ("External Order")
            .AddItem ("Sales Order")
        
        End If
    End With
    
    ' End If
    '///////////
    With Combo1
        If SystemOptions.UserInterface = ArabicInterface Then
            .Clear
            .AddItem ("„ﬁ»Ê·")
            .AddItem ("„—›Ê÷")
        Else
            .Clear
            .AddItem ("Accepted")
            .AddItem ("Refused")
        End If
    End With
    If SystemOptions.IsHiddenTransportInv Then
        Label14(0).Caption = "—ﬁ„ ÿ·» «—«„ﬂÊ"
    End If
    RegisterLogInOut Me.Name, ScreenNameArabic, ScreenNameEnglish, "1"
    
    '  Me.Caption = ScreenNameArabic
    ' Ele(6).Caption = ScreenNameArabic
    ''//////////
    My_SQL = "    select oprid,des from dbo.projects_des"

    fill_combo Me.dcopr2, My_SQL
    If SystemOptions.UserInterface = ArabicInterface Then
        My_SQL = " select  id,Project_name from projects where Project_name<>N'""' And Not (Project_name Is Null)"
    Else
        My_SQL = " select  id,Project_nameE from projects where Project_nameE<>N'""' And Not (Project_nameE Is Null)"
    End If
    fill_combo dcproject1, My_SQL

    If SystemOptions.UserInterface = ArabicInterface Then
        Me.Caption = ScreenNameArabic
    Else
   
        Me.Caption = ScreenNameEnglish
    End If
    Ele(6).Caption = Me.Caption
   
    If SystemOptions.UserInterface = EnglishInterface Then
        SetInterface Me
        ChangeLang

    End If

    ShowTax = GetSetting(StrAppRegPath, "SallBill", "HaveTaxOnSalles", False)
    Ele(4).Visible = ShowTax
    Set Cmd(0).ButtonImage = mdifrmmain.ImgLstTree.ListImages("New").Picture
    Set Cmd(1).ButtonImage = mdifrmmain.ImgLstTree.ListImages("Edit").Picture
    Set Cmd(2).ButtonImage = mdifrmmain.ImgLstTree.ListImages("save").Picture
    Set Cmd(3).ButtonImage = mdifrmmain.ImgLstTree.ListImages("Undo").Picture
    Set Cmd(4).ButtonImage = mdifrmmain.ImgLstTree.ListImages("Del").Picture
    Set Cmd(5).ButtonImage = mdifrmmain.ImgLstTree.ListImages("Search").Picture
    Set Cmd(6).ButtonImage = mdifrmmain.ImgLstTree.ListImages("Exit").Picture
    Set Cmd(7).ButtonImage = mdifrmmain.ImgLstTree.ListImages("Print").Picture
    Set CmdHelp.ButtonImage = mdifrmmain.ImgLstTree.ListImages("Help").Picture
    Set CmdConvert.ButtonImage = mdifrmmain.ImgLstTree.ListImages("Excute").Picture
    Set CmdTemplate.ButtonImage = mdifrmmain.ImgLstTree.ListImages("Excute").Picture
    Set Accredit.ButtonImage = mdifrmmain.ImgLstTree.ListImages("Required").Picture
    Set NewGrid.Grid = FG
    NewGrid.GridTrans = GridTransType.purchaseOrderApproved
        '********************
        Set NewGrid.txtItemCodeSearch = txtItemCodeSearch
 
NewGrid.frmname = Me.Name
    '********************
    Set NewGrid.VatGrid = VatGrid
    Set NewGrid.TxtValueAdded = TxtValueAdded
    Set NewGrid.TxtModFlag = TxtModFlg
    Set NewGrid.txtTotal = XPTxtSum
    Set NewGrid.TxtFillData = TxtFillData
    Set NewGrid.TxtTaxValue = Me.XPTxtTaxValue
    Set NewGrid.XPTxtDiscountVal = Me.XPTxtDiscountVal
    
    Set NewGrid.GrdTBar = Me.TBar
    Set NewGrid.LblItemsCount = Me.LblItemsCount
    'Set NewGrid.LblItemsCount = Me.LblItemsCount
    Set NewGrid.LblTotalAll = Me.LblTotalAll
    Set NewGrid.LblTotalQty = Me.LblTotalQty
    Set NewGrid.LblDiscountsTotal = Me.LblDiscountsTotal
    Set NewGrid.Customer = Me.DBCboClientName
    Set NewGrid.TxtSerial = TxtSerial
    Set NewGrid.CboDiscount_Type = XPCboDiscountType
    Set NewGrid.TxtDiscount_Val = XPTxtDiscountVal
    
    ' ⁄»∆… »Ì«‰«  «·√’‰«›
    Set NewGrid.DCboItemName = DCboItemsName
    Set NewGrid.DCboItemCode = DCboItemsCode
    Set NewGrid.StoreName = DCboStoreName
    Set NewGrid.CboItemCase = CboItemCase
 Set NewGrid.LBLGross = LBLGross
    Set NewGrid.DtpBillDate = Me.XPDtbBill
        
    'Set NewGrid.TxtSerial = TxtSerial
    Set NewGrid.TxtQuantity = TxtQuantity
    Set NewGrid.TxtPrice = TxtPrice
    Set NewGrid.CmdAddData = Me.CmdAdd
    NewGrid.frmname = Me.Name

    NewGrid.FillGrid

    Resize_Form Me, TransactionSize
    Me.left = (mdifrmmain.Width - Me.Width) / 2
    Me.top = (mdifrmmain.Height - Me.Height) / 2 - 500
    FG.WallPaper = BGround.Picture
    AddTip
    XPDtbBill.value = Date
    Set Dcombos = New ClsDataCombos
   
    Dcombos.GetCustomersSuppliers 2, Me.DBCboClientName, True '  2 supplier  1 customer
    Dcombos.GetProcessOfProjedt Me.DcbProcess1
    Dcombos.GetPaymentMathods Me.DcbPayment
    Dcombos.GetShipingMathods Me.DcbShiping
    Dcombos.GetStores Me.DCboStoreName
    Dcombos.GetUsers Me.DCboUserName
    Dcombos.GetBranches Me.dcBranch
    Dcombos.GetEmpDepartments Me.DcbDetpartment
    Set cSearchDcbo(0) = New clsDCboSearch
    Set cSearchDcbo(0).Client = Me.DBCboClientName

    Set cSearchDcbo(1) = New clsDCboSearch
    Set cSearchDcbo(1).Client = Me.DCboStoreName

    Dcombos.GetSalesRepDatapurchase Me.DcboEmp
 
    Set cSearchDcbo(3) = New clsDCboSearch
    Set cSearchDcbo(3).Client = Me.DcboEmp
    cSearchDcbo(3).SetBuddyText Me.TxtEmployeeID

    NewGrid.FillGrid

    With CboPayMentType
        .Clear
        .AddItem "‰ﬁœ«"
        .AddItem "¬Ã·"
    End With
        
    With XPCboDiscountType
        If SystemOptions.UserInterface = ArabicInterface Then
            .Clear
            .AddItem "·«ÌÊÃœ Œ’„"
            .AddItem "Œ’„ »ﬁÌ„…"
            .AddItem "Œ’„ »‰”»…"
        Else
            .Clear
            .AddItem ("No Discount")
            .AddItem ("Discount With Value")
            .AddItem ("Discount With Percentage")
        End If
    End With
      
    With Me.CboPriceType

        If SystemOptions.UserInterface = ArabicInterface Then
            .Clear
            .AddItem " ÿ·»«  «Ê«„— «·»Ì⁄ "
       
        Else
            .Clear
            .AddItem " Sales Order "
 
        End If

        .ListIndex = 0
    End With

    With Me.CboType
        .Clear

        If SystemOptions.UserInterface = ArabicInterface Then
            .AddItem "   ÌœÊÌ "
            .AddItem "«·Ì ÿ»ﬁ« ·Õœ «·ÿ·» "
     
        Else
            .AddItem "Manual"
            .AddItem "Auto "
     
        End If

        .ListIndex = 0
    End With

    'StrSQL = "SELECT * FROM Transactions WHERE (Transaction_Type=6 or Transaction_Type=29  or Transaction_Type=17)" 'OR Transaction_Type=17
   
    My_SQL = " select id,code from currency"
 
    fill_combo Me.Dccurrency, My_SQL
    fill_combo Me.DataCombo11, My_SQL

    My_SQL = " select code,account_name from markaas_taklefa"
 
    fill_combo Me.DataCombo1, My_SQL

    My_SQL = " select id,Project_name from projects"
 
    fill_combo Me.DataCombo2, My_SQL

    My_SQL = " select CountryID,CountryName from TblCountriesData"
 
    fill_combo Me.DataCombo4, My_SQL

    My_SQL = " select id,name from Shipment_mode"
 
    fill_combo Me.DataCombo5, My_SQL
    CboPriceType.ListIndex = GeneralPriceType

    StrSQL = "SELECT * FROM Transactions   WHERE     Transaction_Type=" & CurrentTransactionType
    StrSQL = StrSQL & "  AND      BranchId in(" & Current_branchSql & ")"
    StrSQL = StrSQL + " Order By Transaction_ID"
    Set rs = New ADODB.Recordset
    rs.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText
    
    XPBtnMove_Click 2
    Me.TxtModFlg.text = "R"

    If OPEN_NEW_SCREEN = True Then
        Cmd_Click (0)
    End If
    Exit Sub
ErrTrap:
End Sub
Sub RetriveValueAdded()
Dim sql As String
Dim i As Integer
Dim rs2 As ADODB.Recordset
Set rs2 = New ADODB.Recordset
    VatGrid.Clear flexClearScrollable, flexClearEverything
    VatGrid.rows = 1
sql = " SELECT     dbo.TransactionValueAdded.Transaction_Type, dbo.TransactionValueAdded.Transaction_ID, dbo.TransactionValueAdded.Vat, dbo.TransactionValueAdded.Vatyo,"
sql = sql & " dbo.TransactionValueAdded.ItemID , dbo.TblItems.itemname, dbo.TblItems.Fullcode, dbo.TblItems.ItemNamee ,dbo.TransactionValueAdded.selectd ,dbo.TransactionValueAdded.Valu "
sql = sql & " FROM         dbo.TransactionValueAdded LEFT OUTER JOIN"
sql = sql & "                      dbo.TblItems ON dbo.TransactionValueAdded.ItemID = dbo.TblItems.ItemID"
sql = sql & " Where (dbo.TransactionValueAdded.Transaction_Type = 29) And (dbo.TransactionValueAdded.Transaction_ID = " & val(XPTxtBillID.text) & ")"
rs2.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
If rs2.RecordCount > 0 Then
With Me.VatGrid
rs2.MoveFirst
.rows = .rows + rs2.RecordCount
For i = 1 To .rows - 1
 .TextMatrix(i, .ColIndex("index")) = i
.TextMatrix(i, .ColIndex("ItemID")) = IIf(IsNull(rs2("ItemID").value), "", rs2("ItemID").value)
.TextMatrix(i, .ColIndex("Vat")) = IIf(IsNull(rs2("Vat").value), "", rs2("Vat").value)
.TextMatrix(i, .ColIndex("Vatyo")) = IIf(IsNull(rs2("Vatyo").value), "", rs2("Vatyo").value)
.TextMatrix(i, .ColIndex("Code")) = IIf(IsNull(rs2("Fullcode").value), "", rs2("Fullcode").value)
.TextMatrix(i, .ColIndex("select")) = IIf(IsNull(rs2("selectd").value), 0, rs2("selectd").value)
.TextMatrix(i, .ColIndex("Valu")) = IIf(IsNull(rs2("Valu").value), 0, rs2("Valu").value)

If SystemOptions.UserInterface = ArabicInterface Then
.TextMatrix(i, .ColIndex("Name")) = IIf(IsNull(rs2("ItemName").value), "", rs2("ItemName").value)
Else
.TextMatrix(i, .ColIndex("Name")) = IIf(IsNull(rs2("ItemNamee").value), "", rs2("ItemNamee").value)
End If
rs2.MoveNext
Next i
End With
End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    On Error GoTo ErrTrap
    Dim i As Integer
    RegisterLogInOut Me.Name, ScreenNameArabic, ScreenNameEnglish

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
    Set SaleReport = Nothing
    Exit Sub
ErrTrap:
End Sub

Function CuurentLogdata(Optional Currentmode As String)
    LogTextA = "    ‘«‘… " & ScreenNameArabic & CHR(13) & " —ﬁ„ «·”‰œ   " & TXT_order_no.text & CHR(13) & " «· «—ÌŒ " & XPDtbBill.value & CHR(13) & "«‰Ê⁄ «·”‰œ  " & CboPriceType.text & CHR(13) & " «·„Œ“‰  " & DCboStoreName.text & CHR(13) & "  «·⁄„Ì· / «·„Ê—œ   " & DBCboClientName.text & CHR(13) & " —ﬁ„ «·«⁄ „«œ    " & TxtLcNo
                     
    LogTexte = "    Screen  " & ScreenNameEnglish & CHR(13) & "Vchr . No   " & TXT_order_no.text & CHR(13) & " Date " & XPDtbBill.value & CHR(13) & " Type  " & CboPriceType.text & CHR(13) & " Store  " & DCboStoreName.text & CHR(13) & " Customer/ Supplier " & DBCboClientName.text & CHR(13) & " Lc NO    " & TxtLcNo
                     
    If Currentmode <> "D" Then
        AddToLogFile CInt(user_id), 0, Date, Time, LogTextA, LogTexte, Me.Name, Me.TxtModFlg, "", , , Me.TXT_order_no
    Else
        AddToLogFile CInt(user_id), 0, Date, Time, LogTextA, LogTexte, Me.Name, "D", "", , , Me.TXT_order_no
    End If
    
End Function

Private Sub TxtModFlg_Change()
    On Error GoTo ErrTrap

    Select Case Me.TxtModFlg.text

        Case "R"
            ' Me.Caption = "⁄—÷ √”⁄«—"
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
        
            Me.XPDtbBill.Enabled = False
            Me.DBCboClientName.locked = True
            Me.DCboStoreName.locked = True
            FG.Editable = flexEDNone
      '      Accredit.Enabled = True
            CmdConvert.Enabled = True
            '   CmdConvert.Visible = True
            CmdTemplate.Visible = False

            If rs.RecordCount < 1 Then
                Me.XPBtnMove(0).Enabled = False
                Me.XPBtnMove(1).Enabled = False
                Me.XPBtnMove(2).Enabled = False
                Me.XPBtnMove(3).Enabled = False
                Me.Cmd(1).Enabled = False
                Me.Cmd(4).Enabled = False
                Me.Cmd(5).Enabled = False
                Me.Cmd(7).Enabled = False
                CmdConvert.Enabled = False
                Accredit.Enabled = False
            End If

            Ele(2).Enabled = False

        Case "N"
            ' Me.Caption = "⁄—÷ √”⁄«—( ÃœÌœ )"
            Accredit.Enabled = True
            Me.Cmd(2).Enabled = True
            Me.Cmd(3).Enabled = True
            Me.Cmd(0).Enabled = False
            Me.Cmd(1).Enabled = False
            Me.Cmd(4).Enabled = False
            Me.Cmd(5).Enabled = False
            Me.Cmd(7).Enabled = False
            Accredit.Enabled = False
            '   Me.XPBtnMove(0).Enabled = False
            '   Me.XPBtnMove(1).Enabled = False
            '   Me.XPBtnMove(2).Enabled = False
            '   Me.XPBtnMove(3).Enabled = False
            XPBtnNewClients.Enabled = True
            FG.Enabled = True
            If Not mCopyToNew Then
                FG.rows = 2
            End If
            Me.XPDtbBill.Enabled = True
            XPDtbBill.value = Date
            Me.DBCboClientName.locked = False
            Me.DCboStoreName.locked = False
            FG.Editable = flexEDKbdMouse
        
            CmdConvert.Visible = False
            CmdTemplate.Enabled = True
            '  CmdTemplate.Visible = True
            Ele(2).Enabled = True
            CboItemCase.ListIndex = 0

        Case "E"
            ' Me.Caption = "⁄—÷ √”⁄«—(  ⁄œÌ· )"
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
            Me.XPDtbBill.Enabled = True
            Me.DBCboClientName.locked = False
            Me.DCboStoreName.locked = False
            FG.Editable = flexEDKbdMouse
            XPBtnNewClients.Enabled = True
        
            Accredit.Enabled = False
            CmdConvert.Visible = False
            CmdTemplate.Visible = False
            Ele(2).Enabled = True
    End Select

    Exit Sub
ErrTrap:
End Sub

Public Sub Retrive(Optional Lngid As Long = 0)
    Dim RsDetails As New ADODB.Recordset
    Dim StrSQL As String
    Dim RsNotes As New ADODB.Recordset
    Dim RsTemp As ADODB.Recordset
    Dim Num As Long
    Dim Dusername As String
    On Error GoTo ErrTrap

    If rs.RecordCount < 1 Then
        XPTxtCurrent.Caption = 0
        XPTxtCount.Caption = 0
        Exit Sub
    End If

    If rs.EOF Or rs.BOF Then
        Exit Sub
    End If

    If Lngid <> 0 Then
        rs.Find "Transaction_ID=" & Lngid, , adSearchForward, adBookmarkFirst

        If rs.BOF Or rs.EOF Then
            Exit Sub
        End If
    End If

    TxtFillData.text = "T"
    Screen.MousePointer = vbArrowHourglass
    XPTxtBillID.text = IIf(IsNull(rs("Transaction_ID").value), "", val(rs("Transaction_ID").value))
    
  'ma If SystemOptions.PoCreateVoucher = True Then
   'ma  Me.TXTNoteID.text = IIf(IsNull(rs.Fields("NoteID").value), "", rs.Fields("NoteID").value)
'ma  Me.TxtNoteSerial.text = IIf(IsNull(rs.Fields("NoteSerial").value), "", rs.Fields("NoteSerial").value)
 'ma  End If



    Me.Combo1.ListIndex = IIf(IsNull(rs("Shipping_Pos").value), 0, rs("Shipping_Pos").value)

    Me.DcboEmp.BoundText = IIf(IsNull(rs("Emp_ID").value), "", rs("Emp_ID").value)
    TXT_order_no.text = IIf(IsNull(rs("order_no").value), "", rs("order_no").value)
    TxtPONo.text = IIf(IsNull(rs("PONo").value), "", rs("PONo").value)
    TxtBillComment.text = IIf(IsNull(rs("TransactionComment").value), "", (rs("TransactionComment").value))
      CboPayMentType.ListIndex = IIf(IsNull(rs("PaymentType").value), 0, rs("PaymentType").value)
Me.TxtPayment.text = IIf(IsNull(rs("PaymentT").value), "", (rs("PaymentT").value))
XPCboDiscountType.ListIndex = IIf(IsNull(rs("Trans_DiscountType").value), -1, val(rs("Trans_DiscountType").value))
  XPTxtDiscountVal.text = IIf(IsNull(rs("Trans_Discount").value), "", (rs("Trans_Discount").value))
   CBoBasedON.ListIndex = IIf(IsNull(rs("CBoBasedON").value), 1, (rs("CBoBasedON").value))
   TxtValueAdded.text = IIf(IsNull(rs("VAT").value), 0, (rs("VAT").value))
    If rs("shipped").value = True Then
        chkshipped.value = vbChecked
    Else
        chkshipped.value = Unchecked
    End If

    Me.DataCombo4.BoundText = IIf(IsNull(rs("countryid").value), "", rs("countryid").value)

  If Not (IsNull(rs("CashCustomerName").value)) Then
        Me.TxtCashCustomerName.text = rs("CashCustomerName").value
    Else
        Me.TxtCashCustomerName.text = ""
    End If
    
    TxtTransSerial.text = IIf(IsNull(rs("Transaction_Serial").value), "", (rs("Transaction_Serial").value))
    XPDtbBill.value = IIf(IsNull(rs("Transaction_Date").value), "", (rs("Transaction_Date").value))
    Me.DBCboClientName.BoundText = IIf(IsNull(rs("CusID").value), "", rs("CusID").value)
    
    
    txtContainerNo = IIf(IsNull(rs("ContainerNo").value), "", rs("ContainerNo").value)
    'Me.DcFixedAssets.BoundText = IIf(IsNull(rs("FixesAssetsID").value), "", rs("FixesAssetsID").value)
    
    
    Dccurrency.BoundText = IIf(IsNull(rs("Currency_id").value), "", rs("Currency_id").value)
    'If rs("Transaction_Type").value = 6 Then
    '    Me.CboPriceType.ListIndex = 1
    'ElseIf rs("Transaction_Type").value = 17 Then '17
    '    Me.CboPriceType.ListIndex = 0
    'ElseIf rs("Transaction_Type").value = 29 Then
    'Me.CboPriceType.ListIndex = 2
    'End If

 
   Me.CboPriceType.ListIndex = 0
    Me.DCboUserName.BoundText = IIf(IsNull(rs("UserID").value), "", rs("UserID").value)
    Me.DCboStoreName.BoundText = IIf(IsNull(rs("StoreID").value), "", rs("StoreID").value)
''//
 Me.DcbDetpartment.BoundText = IIf(IsNull(rs("DeptID").value), "", rs("DeptID").value)
 Me.TxtModeRecept.text = IIf(IsNull(rs("ModeReceptEq").value), "", (rs("ModeReceptEq").value))
 Me.TxtModeSupply.text = IIf(IsNull(rs("ModeSupply").value), "", (rs("ModeSupply").value))
''//

    XPTxtTaxValue.text = IIf(IsNull(rs("TaxValue").value), "", (rs("TaxValue").value))
    TxtLcNo.text = IIf(IsNull(rs("LcNo").value), "", (rs("LcNo").value))
    XPChkTAX.value = IIf(rs("TaxFound") = True, Checked, Unchecked)
    dcBranch.BoundText = IIf(IsNull(rs("BranchId").value), "", rs("BranchId").value)

    Me.TxtNoteSerial1.text = IIf(IsNull(rs("NoteSerial1").value), "", (rs("NoteSerial1").value))
    Me.oldtxtNoteSerial1.text = IIf(IsNull(rs("OldNoteSerial1").value), IIf(IsNull(rs("NoteSerial1").value), "", rs("NoteSerial1").value), rs("OldNoteSerial1").value)
''/// 11 05 2015
Me.DcbPayment.BoundText = IIf(IsNull(rs("PaymentID").value), "", rs("PaymentID").value)
Me.DcbShiping.BoundText = IIf(IsNull(rs("ShipingID").value), "", rs("ShipingID").value)
''//
    If TXT_order_no <> "" Then
        Me.TxtNoteSerial1.text = TXT_order_no
    End If
''// 25 05 2015
SippingDate.value = IIf(IsNull(rs("SippingDate").value), Date, (rs("SippingDate").value))
DeliverDate.value = IIf(IsNull(rs("DeliverDate").value), Date, (rs("DeliverDate").value))
TxtPO6.text = IIf(IsNull(rs("NotSeialPO6").value), "", (rs("NotSeialPO6").value))

txtMrNo.text = IIf(IsNull(rs("MrNo").value), "", (rs("MrNo").value))
 If IsNull(rs("chkTaxExempt").value) Then
        Me.chkTaxExempt.value = vbUnchecked
    Else
        Me.chkTaxExempt.value = IIf(rs("chkTaxExempt").value = 0, vbUnchecked, vbChecked)
    End If

    'Txt_order_no
If IsNull(rs("requestOrOrder").value) Then

opt(0).value = True
Else
        If rs("requestOrOrder").value = 0 Then
            opt(0).value = True
        Else
            opt(1).value = True
        End If

End If

 
        
    lbl(64).Caption = showLabel(TxtNoteSerial1, oldtxtNoteSerial1)

'    DBPix202.ImageClear

'    If Dir(App.path & "\images\sign\sign" & rs("posted").value & ".JPG") <> "" Then
'
'        DBPix202.ImageLoadFile (App.path & "\images\sign\sign" & user_id & ".JPG")
'    End If

   If IsNull(rs("posted").value) Then
                                                   If SystemOptions.UserInterface = ArabicInterface Then
                                                    Accredit.Caption = "   «·«—”«· ··«⁄ „«œ "
                                                  Else
                                                    Accredit.Caption = " send to Approval   "
                                               End If
                                               Accredit.Enabled = True
  Else
                                                   If SystemOptions.UserInterface = ArabicInterface Then
                                                    Accredit.Caption = "  „ «·«—”«· ··«⁄ „«œ "
                                                  Else
                                                    Accredit.Caption = " sent to Approval   "
                                               End If
                                               Accredit.Enabled = False
   End If
   
    'If Not IsNull(rs("posted").value) Then
    '    Frame4.Visible = True
    '    GetUserData val(rs("posted").value), , , , , , , Dusername
    '    LblPostedPerson = Dusername

    '                If user_id = rs("posted").value Then
    '                                If CheckOrderNotInTransaction(21, TxtNoteSerial1) = False Then
    '                                                If SystemOptions.UserInterface = ArabicInterface Then
    '                                                    Accredit.Caption = "«·€«¡ «·«⁄ „«œ "
    '                                                Else
    '                                                    Accredit.Caption = "Cancel Accredit   "
    '                                                End If
    '
    '                                Else
    '
    '                                                If SystemOptions.UserInterface = ArabicInterface Then
    '                                                    Accredit.Caption = "  «—”«· ··«⁄ „«œ "
    '                                                Else
    '                                                    Accredit.Caption = " send to accredit   "
     '                                               End If
    '
    '                                End If
    '
    '                End If

    'Else
    '    Frame4.Visible = False
    '    Accredit.Caption = "     «—”«· ··«⁄ „«œ "
    'End If
  
    FG.Clear flexClearScrollable, flexClearEverything
    FG.rows = 2
    FG.Clear flexClearScrollable, flexClearEverything
    FG.Refresh
    StrSQL = " SELECT     TOP 100 PERCENT dbo.TblItems.HaveSerial AS Expr1, *, dbo.projects.Project_name, dbo.projects.Project_nameE, dbo.projects_des.des, "
    StrSQL = StrSQL & "                  dbo.TblProcessDEF.ProcessName , dbo.TblProcessDEF.ProcessNameE"
    StrSQL = StrSQL & "     FROM         dbo.TblItems INNER JOIN"
    StrSQL = StrSQL & "                  dbo.Transaction_Details ON dbo.TblItems.ItemID = dbo.Transaction_Details.Item_ID INNER JOIN"
    StrSQL = StrSQL & "                  dbo.TblUnites ON dbo.Transaction_Details.UnitId = dbo.TblUnites.UnitID LEFT OUTER JOIN"
    StrSQL = StrSQL & "                  dbo.TblProcessDEF ON dbo.Transaction_Details.Oper_ID = dbo.TblProcessDEF.TblProcessDEFID LEFT OUTER JOIN"
    StrSQL = StrSQL & "                  dbo.projects_des ON dbo.Transaction_Details.Pand_ID = dbo.projects_des.oprid LEFT OUTER JOIN"
    StrSQL = StrSQL & "                  dbo.projects ON dbo.Transaction_Details.project_ID1 = dbo.projects.id LEFT OUTER JOIN"
    StrSQL = StrSQL & "                  dbo.TblCountriesData ON dbo.Transaction_Details.countrisid = dbo.TblCountriesData.CountryID"
  StrSQL = StrSQL & " Where (dbo.Transaction_Details.Transaction_ID =" & val(rs("Transaction_ID").value) & ")"
  
 ' StrSQL = " SELECT   dbo.Transaction_Details.unitid as itemunitid , dbo.TblItems.HaveSerial,  dbo.TblUnites.UnitName, dbo.TblUnites.UnitNamee, dbo.Transaction_Details.project_ID1 AS project_ID11, dbo.projects.Project_name, "
 ' StrSQL = StrSQL + "                    dbo.Transaction_Details.Pand_ID AS Pand_ID1, dbo.projects_des.des, dbo.Transaction_Details.Oper_ID AS Oper_ID1, dbo.TblProcessDEF.ProcessName,"
 ' StrSQL = StrSQL + "                    dbo.TblProcessDEF.ProcessNameE, dbo.Transaction_Details.*"
 ' StrSQL = StrSQL + " FROM         dbo.TblProcessDEF RIGHT OUTER JOIN"
 ' StrSQL = StrSQL + "                    dbo.Transaction_Details ON dbo.TblProcessDEF.TblProcessDEFID = dbo.Transaction_Details.Oper_ID LEFT OUTER JOIN"
 ' StrSQL = StrSQL + "                    dbo.projects_des ON dbo.Transaction_Details.Pand_ID = dbo.projects_des.oprid and dbo.projects_des.oprid <> 0 LEFT OUTER JOIN"
 ' StrSQL = StrSQL + "                    dbo.TblUnites ON dbo.Transaction_Details.UnitId = dbo.TblUnites.UnitID LEFT OUTER JOIN"
 ' StrSQL = StrSQL + "                    dbo.projects ON dbo.Transaction_Details.project_ID1 = dbo.projects.id LEFT OUTER JOIN"
 ' StrSQL = StrSQL + "                    dbo.TblItems ON dbo.Transaction_Details.Item_ID = dbo.TblItems.ItemID"
 '' StrSQL = StrSQL + " where Transaction_ID=" & val(rs("Transaction_ID").value)
  'StrSQL = StrSQL + " order by Transaction_Details.id "

    RsDetails.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText
    XPTxtSum.text = ""

    If Not (RsDetails.EOF Or RsDetails.BOF) Then
        FG.rows = RsDetails.RecordCount + 1

        For Num = 1 To RsDetails.RecordCount
        ''//
            FG.TextMatrix(Num, FG.ColIndex("projectid")) = IIf(IsNull(RsDetails("project_ID1")), "", (RsDetails("project_ID1").value))
            FG.TextMatrix(Num, FG.ColIndex("project")) = IIf(IsNull(RsDetails("Project_name")), "", Trim(RsDetails("Project_name").value))
            FG.TextMatrix(Num, FG.ColIndex("pandid")) = IIf(IsNull(RsDetails("Pand_ID")), "", Trim(RsDetails("Pand_ID").value))
            FG.TextMatrix(Num, FG.ColIndex("pand")) = IIf(IsNull(RsDetails("des")), "", (RsDetails("des").value))
            FG.TextMatrix(Num, FG.ColIndex("operaid")) = IIf(IsNull(RsDetails("Oper_ID")), "", Trim(RsDetails("Oper_ID").value))
             If SystemOptions.UserInterface = ArabicInterface Then
             FG.TextMatrix(Num, FG.ColIndex("opera")) = IIf(IsNull(RsDetails("ProcessName")), "", (RsDetails("ProcessName").value))
             Else
             FG.TextMatrix(Num, FG.ColIndex("opera")) = IIf(IsNull(RsDetails("ProcessNameE")), "", (RsDetails("ProcessNameE").value))
             End If
            FG.TextMatrix(Num, FG.ColIndex("ProductionDate")) = IIf(IsNull(RsDetails("ProductionDate")), "", (RsDetails("ProductionDate").value))
            FG.TextMatrix(Num, FG.ColIndex("ExpiryDate")) = IIf(IsNull(RsDetails("ExpiryDate")), "", (RsDetails("ExpiryDate").value))
            FG.TextMatrix(Num, FG.ColIndex("LotNO")) = IIf(IsNull(RsDetails("LotNO")), "", (RsDetails("LotNO").value))
            FG.TextMatrix(Num, FG.ColIndex("catalog")) = IIf(IsNull(RsDetails("catalog").value), "", (RsDetails("catalog").value))
         '   FG.TextMatrix(Num, FG.ColIndex("countris")) = IIf(IsNull(RsDetails("CountryName")), "", Trim(RsDetails("CountryName").value))
         '   FG.TextMatrix(Num, FG.ColIndex("Shipping")) = IIf(IsNull(RsDetails("Shipping")), "", (RsDetails("Shipping").value))
        
           FG.TextMatrix(Num, FG.ColIndex("countris")) = IIf(IsNull(RsDetails("countrisid")), "", Trim(RsDetails("countrisid").value))
            FG.TextMatrix(Num, FG.ColIndex("Shipping")) = IIf(IsNull(RsDetails("Shipping")), "", (RsDetails("Shipping").value))
 
 
        ''//
            FG.TextMatrix(Num, FG.ColIndex("Code")) = IIf(IsNull(RsDetails("ItemCode")), "", (RsDetails("ItemCode").value))
           
            FG.TextMatrix(Num, FG.ColIndex("Count")) = IIf(IsNull(RsDetails("showqty")), "", (RsDetails("showqty").value))
            FG.TextMatrix(Num, FG.ColIndex("Price")) = IIf(IsNull(RsDetails("showPrice")), "", (RsDetails("showPrice").value))
        
            FG.TextMatrix(Num, FG.ColIndex("ItemCase")) = IIf(IsNull(RsDetails("ItemCase")), "", (RsDetails("ItemCase").value))
            FG.TextMatrix(Num, FG.ColIndex("DiscountType")) = IIf(IsNull(RsDetails("ItemDiscountType")), "", (RsDetails("ItemDiscountType").value))
            FG.TextMatrix(Num, FG.ColIndex("DiscountVal")) = IIf(IsNull(RsDetails("ItemDiscount")), "", (RsDetails("ItemDiscount").value))
            FG.TextMatrix(Num, FG.ColIndex("Vat")) = IIf(IsNull(RsDetails("Vat")), "", (RsDetails("Vat").value))
            FG.TextMatrix(Num, FG.ColIndex("Vatyo")) = IIf(IsNull(RsDetails("Vatyo")), "", (RsDetails("Vatyo").value))
            FG.TextMatrix(Num, FG.ColIndex("TypeVAT")) = IIf(IsNull(RsDetails("TypeVAT")), "", (RsDetails("TypeVAT").value))
            If RsDetails("HaveSerial") = True Then
                FG.TextMatrix(Num, FG.ColIndex("HaveSerial")) = True
            End If
            FG.TextMatrix(Num, FG.ColIndex("UnitID")) = IIf(IsNull(RsDetails("UnitName")), "", (RsDetails("UnitName").value))
        If SystemOptions.UserInterface = ArabicInterface Then
           FG.TextMatrix(Num, FG.ColIndex("Name")) = IIf(IsNull(RsDetails("ItemName")), "", Trim(RsDetails("ItemName").value))
           Else
           FG.TextMatrix(Num, FG.ColIndex("Name")) = IIf(IsNull(RsDetails("ItemNamee")), "", Trim(RsDetails("ItemNamee").value))
           End If
            FG.cell(flexcpData, Num, FG.ColIndex("UnitID")) = IIf(IsNull(RsDetails("UnitID")), "", (RsDetails("UnitID").value))
           
           
               FG.TextMatrix(Num, FG.ColIndex("Name")) = IIf(IsNull(RsDetails("ItemID")), "", Trim(RsDetails("ItemID").value))
 FG.TextMatrix(Num, FG.ColIndex("Code")) = IIf(IsNull(RsDetails("ItemID")), "", (RsDetails("ItemID").value))
 
 
 
            FG.TextMatrix(Num, FG.ColIndex("ColorID")) = IIf(IsNull(RsDetails("ColorID")), 1, (RsDetails("ColorID").value))
            FG.TextMatrix(Num, FG.ColIndex("ItemSize")) = IIf(IsNull(RsDetails("ItemSize")), 1, (RsDetails("ItemSize").value))
            FG.TextMatrix(Num, FG.ColIndex("ClassID")) = IIf(IsNull(RsDetails("ClassID")), 1, (RsDetails("ClassID").value))
        
            RsDetails.MoveNext
            Debug.Print Num

            If FG.rows > 10 Then
                If Num = 8 Then FG.Refresh
            End If

        Next Num

    End If
RetriveValueAdded
RelinVatGrid
fillapprovData
FillOrderGrid
FillOrderGrid2
ReLineGrid
showComm
    TxtFillData.text = "F"
    Screen.MousePointer = vbDefault
    XPTxtCurrent.Caption = rs.AbsolutePosition
    XPTxtCount.Caption = rs.RecordCount
    Exit Sub
ErrTrap:
    Screen.MousePointer = vbDefault
End Sub
Private Sub GRID3_Click()

    With FG
        .Clear flexClearScrollable, flexClearEverything
        .rows = 1
       
    End With
 
    fillOrders (0)

End Sub


Function fillOrders(Optional gridno As Integer = 0)
    Dim i As Integer
If gridno = 0 Then
    With Grid3

        For i = 1 To .rows - 1

            If .cell(flexcpChecked, i, .ColIndex("select")) = flexChecked Then
                Retrive_orders_data (val(.TextMatrix(i, .ColIndex("Transaction_ID"))))
            
            End If

        Next i

    End With

Else


    With VSFlexGrid1

        For i = 1 To .rows - 1

            If .cell(flexcpChecked, i, .ColIndex("select")) = flexChecked Then
                Retrive_orders_data (val(.TextMatrix(i, .ColIndex("Transaction_ID"))))
            
            End If

        Next i

    End With
    
End If
End Function



Function Retrive_orders_data(Transaction_ID As Integer)
    Dim StrSQL  As String
    Dim RsDetails As ADODB.Recordset
    Set RsDetails = New ADODB.Recordset
    Dim row_count As Integer
    Dim Num As Integer
    StrSQL = "SELECT TblItems.HaveSerial, * FROM TblItems INNER JOIN Transaction_Details " & "ON TblItems.ItemID = Transaction_Details.Item_ID INNER JOIN dbo.TblUnites ON dbo.Transaction_Details.UnitID = dbo.TblUnites.UnitID"
    StrSQL = StrSQL + " where Transaction_ID=" & Transaction_ID

    RsDetails.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText
    XPTxtSum.text = ""

    If Not (RsDetails.EOF Or RsDetails.BOF) Then
        row_count = FG.rows
    
        If FG.TextMatrix(row_count - 1, FG.ColIndex("Code")) = "" Then
            row_count = row_count - 1
        End If
     
        FG.rows = RsDetails.RecordCount + row_count

        For Num = row_count To FG.rows - 1 'RsDetails.RecordCount
    
'            FG.TextMatrix(Num, FG.ColIndex("order_no")) = IIf(IsNull(RsDetails("order_no")), "", (RsDetails("order_no").value))
'            FG.TextMatrix(Num, FG.ColIndex("OrderArrivalDate")) = IIf(IsNull(RsDetails("OrderArrivalDate")), "", (RsDetails("OrderArrivalDate").value))
'            FG.TextMatrix(Num, FG.ColIndex("FoxyNo")) = IIf(IsNull(RsDetails("FoxyNo")), "", (RsDetails("FoxyNo").value))
        
            FG.TextMatrix(Num, FG.ColIndex("Code")) = IIf(IsNull(RsDetails("Item_ID")), "", (RsDetails("Item_ID").value))
            FG.TextMatrix(Num, FG.ColIndex("Name")) = IIf(IsNull(RsDetails("Item_ID")), "", Trim(RsDetails("Item_ID").value))
        
            '   FG.TextMatrix(Num, FG.ColIndex("Count")) = items_qty_not_recieved_in_order(FG.TextMatrix(Num, FG.ColIndex("Code")), FG.TextMatrix(Num, FG.ColIndex("order_no")))
            FG.TextMatrix(Num, FG.ColIndex("Count")) = IIf(IsNull(RsDetails("Quantity")), "", (RsDetails("Quantity").value))
        
            FG.TextMatrix(Num, FG.ColIndex("Price")) = IIf(IsNull(RsDetails("showprice")), "", (RsDetails("showprice").value))
            FG.TextMatrix(Num, FG.ColIndex("ItemCase")) = IIf(IsNull(RsDetails("ItemCase")), "", (RsDetails("ItemCase").value))
            FG.TextMatrix(Num, FG.ColIndex("DiscountType")) = IIf(IsNull(RsDetails("ItemDiscountType")), "", (RsDetails("ItemDiscountType").value))
            FG.TextMatrix(Num, FG.ColIndex("DiscountVal")) = IIf(IsNull(RsDetails("ItemDiscount")), "", (RsDetails("ItemDiscount").value))
            FG.TextMatrix(Num, FG.ColIndex("ColorID")) = IIf(IsNull(RsDetails("ColorID")), 1, (RsDetails("ColorID").value))
        
            If RsDetails("HaveSerial") = True Then
                FG.TextMatrix(Num, FG.ColIndex("HaveSerial")) = True
            End If
        
            FG.cell(flexcpData, Num, FG.ColIndex("UnitID")) = IIf(IsNull(RsDetails("UnitID")), "", (RsDetails("UnitID").value))
            FG.TextMatrix(Num, FG.ColIndex("UnitID")) = IIf(IsNull(RsDetails("UnitName")), "", (RsDetails("UnitName").value))
        
            RsDetails.MoveNext
            ' Debug.Print Num
            ' If FG.Rows > 10 Then
            '     If Num = 8 Then FG.Refresh
            ' End If
        Next Num

    End If

End Function
Sub SaveValueAdded()
Dim i As Integer
Dim sql As String
Dim rs2 As ADODB.Recordset
Set rs2 = New ADODB.Recordset

sql = "Select * from  TransactionValueAdded where 1=-1"
rs2.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
With Me.VatGrid
For i = 1 To .rows - 1
If val(.TextMatrix(i, .ColIndex("ItemID"))) <> 0 Then
rs2.AddNew
rs2("Transaction_ID").value = val(Me.XPTxtBillID.text)
rs2("Transaction_Type").value = 29
rs2("ItemID").value = val(.TextMatrix(i, .ColIndex("ItemID")))
rs2("Vatyo").value = val(.TextMatrix(i, .ColIndex("Vatyo")))
rs2("Vat").value = val(.TextMatrix(i, .ColIndex("Vat")))
rs2("Valu").value = val(.TextMatrix(i, .ColIndex("Valu")))
If .cell(flexcpChecked, i, .ColIndex("select")) = flexChecked Then
rs2("selectd").value = 1
Else
rs2("selectd").value = 0
End If
rs2.update
End If
Next i
End With
End Sub
Function fillapprovData()
Dim Num As Integer
 Dim RsDetails As New ADODB.Recordset
 Dim StrSQL As String
'    StrSQL = "SELECT     TOP 100 PERCENT  dbo.ApprovalData.Currcursor, dbo.ApprovalData.ScreenName, dbo.ApprovalData.levelo, dbo.ApprovalData.EmpID, dbo.ApprovalData.levelorder, dbo.ApprovalData.currorder, "
'StrSQL = StrSQL + "  dbo.ApprovalData.Transaction_ID, dbo.ApprovalData.NoteID, dbo.ApprovalData.ApprovDate, dbo.ApprovalData.Remarks, dbo.TblEmployee.Emp_Code,"
'StrSQL = StrSQL + "   dbo.TblEmployee.emp_name , dbo.TblEmployee.Emp_Namee, dbo.TbLLevels.name, dbo.TbLLevels.namee"
'StrSQL = StrSQL + " FROM         dbo.ApprovalData INNER JOIN"
'StrSQL = StrSQL + "   dbo.TblEmployee ON dbo.ApprovalData.EmpID = dbo.TblEmployee.Emp_ID INNER JOIN"
'StrSQL = StrSQL + "   dbo.TbLLevels ON dbo.ApprovalData.levelo = dbo.TbLLevels.LevelID"
'StrSQL = StrSQL + " WHERE     (dbo.ApprovalData.Transaction_ID = " & val(XPTxtBillID.text) & ") AND (dbo.ApprovalData.ScreenName = N'" & Me.name & "')"
'StrSQL = StrSQL + "  ORDER BY dbo.ApprovalData.levelorder"
 
 StrSQL = "SELECT     TOP 100 PERCENT dbo.ApprovalData.Currcursor, dbo.ApprovalData.ScreenName, dbo.ApprovalData.levelo, dbo.ApprovalData.EmpID, dbo.ApprovalData.levelorder, "
StrSQL = StrSQL + " dbo.ApprovalData.currorder, dbo.ApprovalData.Transaction_ID, dbo.ApprovalData.NoteID, dbo.ApprovalData.ApprovDate, dbo.ApprovalData.Remarks,"
StrSQL = StrSQL + " dbo.TbLLevels.name , dbo.TbLLevels.namee, dbo.TblUsers.UserID, dbo.TblUsers.UserName"
StrSQL = StrSQL + " FROM         dbo.ApprovalData INNER JOIN"
StrSQL = StrSQL + " dbo.TbLLevels ON dbo.ApprovalData.levelo = dbo.TbLLevels.LevelID INNER JOIN"
StrSQL = StrSQL + " dbo.TblUsers ON dbo.ApprovalData.EmpID = dbo.TblUsers.UserID"
StrSQL = StrSQL + " WHERE     (dbo.ApprovalData.Transaction_ID = " & val(XPTxtBillID.text) & ") AND (dbo.ApprovalData.ScreenName = N'" & Me.Name & "')"
'StrSQL = StrSQL + " and ApprovalData.empid in (Select tblusers.UserID from tblusers where tblusers.BranchId = " & val(dcBranch.BoundText) & "  )"
StrSQL = StrSQL + " ORDER BY dbo.ApprovalData.levelorder"

    RsDetails.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText

 If Not (RsDetails.EOF Or RsDetails.BOF) Then
        Grid2.rows = RsDetails.RecordCount + 1
 

        For Num = 1 To RsDetails.RecordCount
        
       Grid2.TextMatrix(Num, Grid2.ColIndex("Currcursor")) = IIf(IsNull(RsDetails("Currcursor")), "", RsDetails("Currcursor"))
    If Grid2.TextMatrix(Num, Grid2.ColIndex("Currcursor")) = "1" Then
   Grid2.cell(flexcpBackColor, Num, 1, Num, 7) = &HFFFFC0
   Else
    Grid2.cell(flexcpBackColor, Num, 1, Num, 7) = vbWhite
    End If
    
        Grid2.TextMatrix(Num, Grid2.ColIndex("Approved")) = IIf(IsNull(RsDetails("ApprovDate")), "", flexChecked)
           If SystemOptions.UserInterface = ArabicInterface Then
            Grid2.TextMatrix(Num, Grid2.ColIndex("levelName")) = IIf(IsNull(RsDetails("Name")), "", Trim(RsDetails("Name").value))
          Else
             Grid2.TextMatrix(Num, Grid2.ColIndex("levelName")) = IIf(IsNull(RsDetails("Namee")), "", Trim(RsDetails("Namee").value))
          End If
            If SystemOptions.UserInterface = ArabicInterface Then
            Grid2.TextMatrix(Num, Grid2.ColIndex("EmpName")) = IIf(IsNull(RsDetails("UserName")), "", (RsDetails("UserName").value))
            Else
            Grid2.TextMatrix(Num, Grid2.ColIndex("EmpName")) = IIf(IsNull(RsDetails("UserName")), "", (RsDetails("UserName").value))
            End If
            Grid2.TextMatrix(Num, Grid2.ColIndex("ApprovDate")) = IIf(IsNull(RsDetails("ApprovDate")), "", (RsDetails("ApprovDate").value))
          Grid2.TextMatrix(Num, Grid2.ColIndex("REMARKS")) = IIf(IsNull(RsDetails("REMARKS")), "", (RsDetails("REMARKS").value))
 
 
RsDetails.MoveNext
If Num = RsDetails.RecordCount Then

        If Grid2.TextMatrix(Num, Grid2.ColIndex("Approved")) <> "" Then
                                If SystemOptions.UserInterface = ArabicInterface Then
                                      Label11.Caption = " „ «·«⁄ „«œ ··„” ‰œ »«·ﬂ«„·"
                                 Else
                                       Label11.Caption = "Approved"
                                 End If
                            Label11.backcolor = &H80FF80
        Else
                             If SystemOptions.UserInterface = ArabicInterface Then
                                     Label11.Caption = "„ÿ·Ê» «⁄ „«œ… Õ«·Ì«"
                            Else
                                     Label11.Caption = "Currently required Approve"
                            End If
                 Label11.backcolor = &HFFFFC0
        End If

End If

        Next Num
Else
 Grid2.rows = 1
    End If
RsDetails.Close
Accredit.Caption = Label11.Caption
End Function



Private Sub XPTxtDiscountVal_Change()
    Dim Msg As String
    On Error GoTo ErrTrap

    If Me.TxtModFlg.text = "N" Or Me.TxtModFlg.text = "E" Then
        NewGrid.Calculate 1, , , True
    End If

    Exit Sub
ErrTrap:
End Sub

Private Sub XPTxtSum_Change()
    On Error GoTo ErrTrap
 
    Me.LblTotal.Caption = XPTxtSum.text
 
    Exit Sub
ErrTrap:
End Sub

Private Sub Undo()
    Dim Msg As String

    On Error GoTo ErrTrap

    Select Case TxtModFlg.text

        Case "N"
            Msg = "”Ê› Ì „ «· —«Ã⁄ ›Ï  ”ÃÌ· Â–« «·”‰œ   .."
            Msg = Msg & CHR(13) & "›Â· «‰  „ «ﬂœ „‰ «·√” „—«— ..!!"

            If MsgBox(Msg, vbQuestion + vbYesNo + vbMsgBoxRight + vbMsgBoxRtlReading + vbDefaultButton2, App.Title) = vbYes Then
                clear_all Me
                Me.TxtModFlg.text = "R"
                XPBtnMove_Click (1)
            End If

        Case "E"
            Msg = "”Ê› Ì „ «· —«Ã⁄ ›Ï  ⁄œÌ·    Â–« «·”‰œ .."
            Msg = Msg & CHR(13) & "›Â· «‰  „ «ﬂœ „‰ «·√” „—«— ..!!"

            If MsgBox(Msg, vbQuestion + vbYesNo + vbMsgBoxRight + vbMsgBoxRtlReading + vbDefaultButton2, App.Title) = vbYes Then
                rs.Find "Transaction_ID='" & val(XPTxtBillID.text) & "'", , adSearchForward, adBookmarkFirst

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
    On Error GoTo ErrTrap
    Dim Msg  As String

    If XPTxtBillID.text <> "" Then
        Msg = "”Ì „ Õ–› »Ì«‰«  «·⁄„·Ì… —ﬁ„ " & CHR(13)
        Msg = Msg + (XPTxtBillID.text) & CHR(13)
        Msg = Msg + " Â·  —€» ›Ì Õ–› Â–Â «·»Ì«‰« ø"

        If MsgBox(Msg, vbYesNo + vbQuestion + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title) = vbYes Then
            If Not rs.RecordCount < 1 Then
                CuurentLogdata ("D")
                          
    On Error Resume Next
    Dim i As Integer
    Dim sql As String
 
     

    With Grid3

        For i = 1 To .rows - 1
     
 
                sql = "update transactions set closed=0 ,nots='' ,nots2='' where  Transaction_ID=" & val(.TextMatrix(i, .ColIndex("Transaction_ID")))
               
        
       
            Cn.Execute sql
 
        Next
       
    End With
    
                
      With VSFlexGrid1

        For i = 1 To .rows - 1
     
 
                sql = "update transactions set closed=0 ,nots='' ,nots2='' where  Transaction_ID=" & val(.TextMatrix(i, .ColIndex("Transaction_ID")))
               
        
       
            Cn.Execute sql
 
        Next
       
    End With
    Cn.Execute "Delete from TransactionValueAdded where Transaction_ID=" & val(Me.XPTxtBillID.text) & ""
                rs.delete
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

    Else
        clear_all Me
        Msg = "Â–Â «·⁄„·Ì… €Ì— „ «Õ… ÕÌÀ √‰Â ·«ÌÊÃœ √Ì ”Ã·« "
        MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
        TxtModFlg_Change
        Exit Sub
    End If

    TxtModFlg_Change
    Exit Sub
ErrTrap:

    If Err.Number = -2147217887 Then
        Msg = "·«Ì„ﬂ‰ Õ–› Â–« «·”Ã· · ﬂ«„· «·»Ì«‰«  " & CHR(13) & " ÊÃœ »Ì«‰«  „— »ÿ… »Â–« «·„Ê—œ "
        MsgBox Msg, vbMsgBoxRight + vbMsgBoxRtlReading + vbExclamation, App.Title
        rs.CancelUpdate
    End If

End Sub

Private Sub AddTip()
    Dim Wrap As String
    On Error GoTo ErrTrap
    Wrap = CHR(13) + CHR(10)
    Set TTP = New clstooltip

    With TTP
        .Create Me.hWnd, "⁄—÷ √”⁄«—", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
        .AddControl Cmd(0), "ÃœÌœ ..." & Wrap & "·«÷«›… »Ì«‰«  ⁄—÷ ”⁄— ÃœÌœ" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
    End With

    With TTP
        .Create Me.hWnd, "⁄—÷ √”⁄«—", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
        .AddControl Cmd(7), "ÿ»«⁄… ..." & Wrap & "·⁄—÷  ﬁ—Ì— »«·»Ì«‰«  «·Õ«·Ì… " & Wrap & " Ì„ﬂ‰ ÿ»«⁄ Â ⁄‰ ÿ—Ìﬁ «·ÿ«»⁄…", True
    End With

    With TTP
        .Create Me.hWnd, "⁄—÷ √”⁄«—", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
        .AddControl Cmd(1), " ⁄œÌ· ..." & Wrap & "· ⁄œÌ· »Ì«‰«  «·⁄—÷ «·Õ«·Ì" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
    End With

    With TTP
        .Create Me.hWnd, "⁄—÷ √”⁄«—", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
        .AddControl Cmd(2), "Õ›Ÿ ..." & Wrap & "·Õ›Ÿ »Ì«‰«  ⁄—÷ «·”⁄— «·ÃœÌœ" & Wrap & "·Õ›Ÿ «· ⁄œÌ·« " & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
    End With

    With TTP
        .Create Me.hWnd, "⁄—÷ √”⁄«—", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
        .AddControl Cmd(3), " —«Ã⁄ ..." & Wrap & "·· —«Ã⁄ ⁄‰ ⁄„·Ì… «·≈÷«›…" & Wrap & "··· —«Ã⁄ ⁄‰ ⁄„·Ì… «· ⁄œÌ·" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
    End With

    With TTP
        .Create Me.hWnd, "⁄—÷ √”⁄«—", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
        .AddControl Cmd(4), "Õ–› ..." & Wrap & "·Õ–› »Ì«‰«  «·⁄—÷ «·Õ«·Ì" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
    End With

    With TTP
        .Create Me.hWnd, "⁄—÷ √”⁄«—", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
        .AddControl Cmd(5), "»ÕÀ ..." & Wrap & "···»ÕÀ ⁄‰ ⁄—÷ ”⁄—" & Wrap & "Ì‰ÿ»ﬁ ⁄·ÌÂ« ‘—Êÿ „⁄Ì‰…" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
    End With

    With TTP
        .Create Me.hWnd, "⁄—÷ √”⁄«—", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
        .AddControl Cmd(6), "Œ—ÊÃ ..." & Wrap & "·«€·«ﬁ Â–Â «·‰«›–…" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
    End With

    With TTP
        .Create Me.hWnd, "⁄—÷ √”⁄«—", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
        .AddControl XPBtnNewClients, "≈÷«›… ⁄„Ì· ÃœÌœ ..." & Wrap & "· ”ÃÌ· »Ì«‰«  ⁄„Ì· ÃœÌœ" & Wrap & " «÷€ÿ Â‰«", True
    End With

    With TTP
        .Create Me.hWnd, "⁄—÷ √”⁄«—", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
        .AddControl XPBtnMove(1), "«·√Ê· ..." & Wrap & "··«‰ ﬁ«· «·Ï √Ê· ”Ã·" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
    End With

    With TTP
        .Create Me.hWnd, "⁄—÷ √”⁄«—", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
        .AddControl XPBtnMove(0), "«·”«»ﬁ ..." & Wrap & "··«‰ ﬁ«· «·Ï «·”Ã· «·”«»ﬁ" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
    End With

    With TTP
        .Create Me.hWnd, "⁄—÷ √”⁄«—", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
        .AddControl XPBtnMove(3), "«· «·Ì ..." & Wrap & "··«‰ ﬁ«· «·Ï «·”Ã· «· «·Ì" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
    End With

    With TTP
        .Create Me.hWnd, "⁄—÷ √”⁄«—", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
        .AddControl XPBtnMove(2), "«·√ŒÌ— ..." & Wrap & "··«‰ ﬁ«· «·Ï ¬Œ— ”Ã·" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
    End With

    With TTP
        .Create Me.hWnd, "⁄—÷ √”⁄«—", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
        .AddControl CmdHelp, "„”«⁄œ… ..." & Wrap & "·· ⁄—› ⁄·Ï ÊŸÌ›… Â–Â «·‰«›–…" & Wrap & "ÊﬂÌ›Ì… «· ⁄«„· „⁄Â«" & Wrap & "≈÷€ÿ Â‰«" & Wrap, True
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

    Me.lbl(55).Visible = (Me.XPCboDiscountType.ListIndex = 2)

    'Me.lbl(21).Visible = (Me.XPCboDiscountType.ListIndex = 2)
    If XPCboDiscountType.ListIndex = 0 Then
      '  lbl(8).Visible = False
        XPTxtDiscountVal.Visible = False
      '  lbl(8).Visible = False
    Else
     '   lbl(8).Visible = True
        XPTxtDiscountVal.Visible = True
      '  lbl(8).Visible = True
    End If

    Exit Sub
ErrTrap:
End Sub


Private Function CheckCashCustomer() As Boolean
    Dim rs As ADODB.Recordset
    Dim StrSQL As String

    If Trim$(Me.TxtCashCustomerName.text) = "" Then
        CheckCashCustomer = True
    Else
        StrSQL = "Select * From Transactions Where CashCustomerName='" & Trim$(Me.TxtCashCustomerName.text) & "'"
    
    End If

End Function


Private Sub SaveData()
    Dim Msg As String
    Dim RowNum As Integer
    Dim RSTransDetails As ADODB.Recordset
    'Dim RsNotes As ADODB.Recordset
    Dim RsTemp  As New ADODB.Recordset
    Dim RsTest As New ADODB.Recordset
    Dim RsRepeat As ADODB.Recordset
    Dim StrSQL As String
    Dim StrSqlDel As String
    Dim BeginTrans As Boolean
'    On Error GoTo ErrTrap
    Screen.MousePointer = vbArrowHourglass

    If Me.TxtModFlg.text <> "R" Then
        If DBCboClientName.text = "" And opt(1).value = True Then
            If SystemOptions.UserInterface = ArabicInterface Then
                Msg = "„‰ ›÷·ﬂ √œŒ· «”„ «·„Ê—œ"
            Else
                Msg = "Please Select Vendor"
            End If

            MsgBox Msg, vbOKOnly + vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
            DBCboClientName.SetFocus
           'ma  SendKeys "{F4}"
            Screen.MousePointer = vbDefault
            Exit Sub
        End If

        If DCboStoreName.text = "" Then
            If SystemOptions.UserInterface = ArabicInterface Then
                Msg = "ÌÃ»  ÕœÌœ «·„Œ“‰"
            Else
                Msg = "Select Inventory"
            End If

            MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
            DCboStoreName.SetFocus
            'sa SendKeys "{F4}"
            Screen.MousePointer = vbDefault
            Exit Sub
        End If
    
        If Dccurrency.text = "" Then
            If SystemOptions.UserInterface = ArabicInterface Then
                Msg = "Õœœ «·⁄„·…"
            Else
                Msg = "Select Currency"
            End If

            MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
            Dccurrency.SetFocus
            Sendkeys "{F4}"
            Screen.MousePointer = vbDefault
            Exit Sub
        End If
    
    
    If CboPayMentType.ListIndex = -1 Then
        If SystemOptions.UserInterface = ArabicInterface Then
            Msg = "ÌÃ»  ÕœÌœ ÿ—Ìﬁ… «·œ›⁄"
        Else
            Msg = "Specify Payment Method"
        End If

        MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
        CboPayMentType.SetFocus
        Sendkeys "{F4}"
        Screen.MousePointer = vbDefault
        Exit Sub
    End If
    
        If Me.CboPriceType.ListIndex = -1 Then
            If SystemOptions.UserInterface = ArabicInterface Then
                Msg = "ÌÃ»  ÕœÌœ ‰Ê⁄    «·«„—  ( )...!!!"
            Else
                Msg = "Specify Order Type"
            End If

            MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
            CboPriceType.SetFocus
            Sendkeys "{F4}"
            Screen.MousePointer = vbDefault
            Exit Sub
        End If



 If XPCboDiscountType.ListIndex = 1 Or XPCboDiscountType.ListIndex = 2 Then
        If XPTxtDiscountVal.text = "" Then
            If SystemOptions.UserInterface = ArabicInterface Then
                Msg = "≈–« ﬂ«‰ Â‰«ﬂ Œ’„ ⁄·Ï «·«„— " & CHR(13)
                Msg = Msg + "ÌÃ»  ÕœÌœ ﬁÌ„… Â–« «·Œ’„ " & CHR(13)
                Msg = Msg + "√Ê √Œ Ì«— ·« ÌÊÃœ Œ’„ "
            Else
                Msg = Msg + " Must Enter Discount Value " & CHR(13)
            End If

            MsgBox Msg, vbOKOnly + vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
            XPCboDiscountType.SetFocus
            Sendkeys "{F4}"
            Screen.MousePointer = vbDefault
            Exit Sub
        End If
    End If




        If XPChkTAX.value = Checked Then
            If XPTxtTaxValue.text = "" Then
                If SystemOptions.UserInterface = ArabicInterface Then
                    Msg = "ÌÃ» «œŒ«· ﬁÌ„… ÷—Ì»… «·„»Ì⁄« "
                Else
                    Msg = "Insert Sales Tax"
                End If

                MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
                XPTxtTaxValue.SetFocus
                FG.SetFocus
                Screen.MousePointer = vbDefault
                Exit Sub
            End If
        End If
    
 
    
        If NewGrid.CheckDataEntered = False Then
            Exit Sub
        End If
Dim sql As String
        Set RSTransDetails = New ADODB.Recordset
        sql = "Select * from Transaction_Details where 1=-1 "
        RSTransDetails.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
        Dim Transaction_Type As Integer
        Dim Sanad_No As Integer

        If Me.CboPriceType.ListIndex = 0 Then
            Transaction_Type = CurrentTransactionType
            Sanad_No = CurrentTransactionType
  
 
         
        End If

        my_branch = val(dcBranch.BoundText)

        If TxtNoteSerial1.text = "" Then
            If Voucher_coding(val(my_branch), XPDtbBill.value, Sanad_No, 0, , Transaction_Type, , val(DCboStoreName.BoundText)) = "error" Then
                MsgBox " ·« Ì„ﬂ‰ «÷«›…   Â–« «·”‰œ ·«‰ﬂ  ⁄œÌ  «·Õœ «·„”„ÊÕ »… „‰ «·”‰œ«   ": Exit Sub
            Else
                       
                If Voucher_coding(val(my_branch), XPDtbBill.value, Sanad_No, 0, , Transaction_Type, , val(DCboStoreName.BoundText)) = "" Then
                    MsgBox " ·«»œ „‰ ﬂ «»… —ﬁ„ «·”‰œ ÌœÊÌ« ﬂ„« Õœœ    " & CHR(13) & " Enter Vchr No": Exit Sub
                Else
                    TxtNoteSerial1.text = Voucher_coding(val(my_branch), XPDtbBill.value, Sanad_No, 170, , Transaction_Type, , val(DCboStoreName.BoundText))
                End If
            End If
        End If
 
        TXT_order_no = Me.TxtNoteSerial1.text
 
        Cn.BeginTrans
        BeginTrans = True
    
        If Me.TxtModFlg.text = "N" Then
            XPTxtBillID.text = CStr(new_id("Transactions", "Transaction_ID", "", True))
            TxtTransSerial.text = CStr(new_id("Transactions", "Transaction_Serial", "", True, "Transaction_Type=6"))
            
            Me.oldtxtNoteSerial1.text = Trim$(Me.TxtNoteSerial1.text)
            rs.AddNew
         Else
           If SystemOptions.PoCreateVoucher = True Then
                StrSQL = "Delete From DOUBLE_ENTREY_VOUCHERS Where Notes_ID=" & val(Me.TXTNoteID.text)
                  Cn.Execute StrSQL, , adExecuteNoRecords

         End If
        End If

        Screen.MousePointer = vbArrowHourglass
       
'////////
rs("Shipping_Pos").value = IIf(Combo1.ListIndex = -1, Null, Combo1.ListIndex)

       rs("NoteSerial1").value = IIf(Trim(Me.TxtNoteSerial1.text) = "", Null, Trim(Me.TxtNoteSerial1.text))
        rs("OldNoteSerial1").value = Trim$(Me.oldtxtNoteSerial1.text) '
        rs("branchID").value = val(Me.dcBranch.BoundText)
     rs("TransactionComment").value = IIf(Trim$(TxtBillComment.text) = "", Null, Trim$(TxtBillComment.text))
       rs("Transaction_ID").value = val(XPTxtBillID.text)
        rs("order_no").value = TXT_order_no.text
        
         If CBoBasedON.ListIndex = -1 Then
        rs("CBoBasedON").value = 0
    Else
        rs("CBoBasedON").value = val(CBoBasedON.ListIndex)
    End If


        If chkshipped.value = vbChecked Then
            rs("shipped").value = 1
        Else
            rs("shipped").value = 0
        End If
    
    
       If opt(0).value = True Then
            rs("requestOrOrder").value = 0
        Else
            rs("requestOrOrder").value = 1
        End If
        
        
        rs("Transaction_Date").value = XPDtbBill.value
        rs("Transaction_Serial").value = TxtTransSerial.text

      rs("PONO").value = IIf(TxtPONo.text = "", Null, (TxtPONo.text))
rs("Transaction_Type").value = CurrentTransactionType

        rs("UserID").value = user_id
        rs("CusID").value = IIf(DBCboClientName.BoundText = "", Null, val(DBCboClientName.BoundText))
        'rs("FixesAssetsID").value = IIf(DcFixedAssets.BoundText = "", Null, val(DcFixedAssets.BoundText))
        
        rs("ContainerNo").value = IIf(txtContainerNo.text = "", Null, Trim(txtContainerNo.text))
            
        
        
        rs("countryid").value = IIf(DataCombo4.BoundText = "", Null, val(DataCombo4.BoundText))
    
        rs("Currency_id").value = IIf(Dccurrency.BoundText = "", Null, val(Dccurrency.BoundText))
    
       If Trim$(Me.TxtCashCustomerName.text) <> "" Then
        rs("CashCustomerName").value = Trim$(Me.TxtCashCustomerName.text)
    Else
        rs("CashCustomerName").value = Null
    End If
    
       If CboPayMentType.ListIndex = -1 Then
        rs("PaymentType").value = 0
    Else
        rs("PaymentType").value = val(CboPayMentType.ListIndex)
    End If
    rs("PaymentT").value = Trim$(Me.TxtPayment.text)
    
       If XPCboDiscountType.ListIndex = -1 Then
        rs("Trans_DiscountType").value = 0
    Else
        rs("Trans_DiscountType").value = val(XPCboDiscountType.ListIndex)
    End If
  rs("Trans_Discount").value = IIf(XPTxtDiscountVal.text = "", Null, val(XPTxtDiscountVal.text))
 ''//
   rs("DeptID").value = IIf(DcbDetpartment.BoundText = "", Null, val(Me.DcbDetpartment.BoundText))
   rs("ModeReceptEq").value = IIf(Me.TxtModeRecept.text = "", Null, TxtModeRecept.text)
   rs("ModeSupply").value = IIf(Me.TxtModeSupply.text = "", Null, TxtModeSupply.text)

 ''//
 
        rs("Emp_ID").value = IIf(DcboEmp.BoundText = "", Null, DcboEmp.BoundText)
        rs("StoreID").value = IIf(DCboStoreName.BoundText = "", Null, val(DCboStoreName.BoundText))
        rs("TaxFound").value = IIf(XPChkTAX.value = Checked, True, False)
        rs("TaxValue").value = IIf(XPTxtTaxValue.text = "", Null, val(XPTxtTaxValue.text))
        rs("total").value = IIf(XPTxtSum.text = "", Null, val(XPTxtSum.text))
        rs("LcNo").value = IIf(TxtLcNo.text = "", Null, (TxtLcNo.text))
        
    
    ''//11 05 2015
    rs("ShipingID").value = val(Me.DcbShiping.BoundText)
    rs("PaymentID").value = val(Me.DcbPayment.BoundText)
    ''//
   ''//25 05 2015
   rs("SippingDate").value = SippingDate.value
   rs("DeliverDate").value = DeliverDate.value
   rs("NotSeialPO6").value = Me.TxtPO6.text
   rs("MrNo").value = Me.txtMrNo.text
   
   rs("VAT").value = val(TxtValueAdded.text)
      If chkTaxExempt.value = vbChecked Then
            rs("chkTaxExempt").value = 1
        Else
            rs("chkTaxExempt").value = 0
        End If

        rs.update
    
       CuurentLogdata
  
        If Me.TxtModFlg.text = "E" Then
           Cn.Execute "Delete from TransactionValueAdded where Transaction_ID=" & val(Me.XPTxtBillID.text) & ""
            StrSqlDel = "delete From Transaction_Details where Transaction_ID=" & val(rs("Transaction_ID").value)
            Cn.Execute StrSqlDel, , adExecuteNoRecords
        End If
Dim TotalShahnPerLine As Double
        For RowNum = 1 To FG.rows - 1

            If FG.TextMatrix(RowNum, FG.ColIndex("Code")) <> "" Then
                RSTransDetails.AddNew
                RSTransDetails("Transaction_ID").value = val(XPTxtBillID.text)
                RSTransDetails("order_id").value = val(XPTxtBillID.text)
             
                RSTransDetails("order_no").value = TXT_order_no.text
                ''//
               ' RSTransDetails("countrisid").value = IIf((FG.TextMatrix(RowNum, FG.ColIndex("countrisid")) = ""), Null, val(FG.TextMatrix(RowNum, FG.ColIndex("countrisid"))))
               ' RSTransDetails("Shipping").value = IIf((FG.TextMatrix(RowNum, FG.ColIndex("Shipping")) = ""), Null, val(FG.TextMatrix(RowNum, FG.ColIndex("Shipping"))))
                RSTransDetails("countrisid").value = IIf((FG.TextMatrix(RowNum, FG.ColIndex("countris")) = ""), Null, val(FG.TextMatrix(RowNum, FG.ColIndex("countris"))))
                RSTransDetails("Shipping").value = IIf((FG.TextMatrix(RowNum, FG.ColIndex("Shipping")) = ""), Null, val(FG.TextMatrix(RowNum, FG.ColIndex("Shipping"))))
                RSTransDetails("Oper_ID").value = IIf((FG.TextMatrix(RowNum, FG.ColIndex("operaid")) = ""), Null, val(FG.TextMatrix(RowNum, FG.ColIndex("operaid"))))
                RSTransDetails("Pand_ID").value = IIf((FG.TextMatrix(RowNum, FG.ColIndex("pandid")) = ""), Null, val(FG.TextMatrix(RowNum, FG.ColIndex("pandid"))))
                RSTransDetails("project_ID1").value = IIf((FG.TextMatrix(RowNum, FG.ColIndex("projectid")) = ""), Null, val(FG.TextMatrix(RowNum, FG.ColIndex("projectid"))))
                RSTransDetails("Vat").value = IIf((FG.TextMatrix(RowNum, FG.ColIndex("Vat")) = ""), Null, (FG.TextMatrix(RowNum, FG.ColIndex("Vat"))))
                RSTransDetails("Vatyo").value = IIf((FG.TextMatrix(RowNum, FG.ColIndex("Vatyo")) = ""), Null, (FG.TextMatrix(RowNum, FG.ColIndex("Vatyo"))))
                RSTransDetails("TypeVAT").value = IIf(FG.TextMatrix(RowNum, FG.ColIndex("TypeVAT")) = "", Null, FG.TextMatrix(RowNum, FG.ColIndex("TypeVAT")))
                RSTransDetails("catalog").value = IIf((FG.TextMatrix(RowNum, FG.ColIndex("catalog")) = ""), Null, val(FG.TextMatrix(RowNum, FG.ColIndex("catalog"))))
                RSTransDetails("Item_ID").value = IIf((FG.TextMatrix(RowNum, FG.ColIndex("Code")) = ""), Null, val(FG.TextMatrix(RowNum, FG.ColIndex("Code"))))
                RSTransDetails("Quantity").value = IIf((FG.TextMatrix(RowNum, FG.ColIndex("Count")) = ""), Null, val(FG.TextMatrix(RowNum, FG.ColIndex("Count"))))
                RSTransDetails("ShowPrice").value = IIf((FG.TextMatrix(RowNum, FG.ColIndex("Price")) = ""), Null, val(FG.TextMatrix(RowNum, FG.ColIndex("Price"))))
                RSTransDetails("ItemDiscountType").value = IIf((FG.TextMatrix(RowNum, FG.ColIndex("DiscountType")) = ""), Null, val(FG.TextMatrix(RowNum, FG.ColIndex("DiscountType"))))
                RSTransDetails("ItemCase").value = IIf((FG.TextMatrix(RowNum, FG.ColIndex("ItemCase")) = ""), Null, val(FG.TextMatrix(RowNum, FG.ColIndex("ItemCase"))))
                
                
                            

       
            
            
            
                
                RSTransDetails("ItemDiscount").value = IIf((FG.TextMatrix(RowNum, FG.ColIndex("DiscountVal")) = ""), Null, val(FG.TextMatrix(RowNum, FG.ColIndex("DiscountVal"))))
            
                RSTransDetails("ColorID").value = IIf((FG.TextMatrix(RowNum, FG.ColIndex("ColorID")) = ""), 1, val(FG.TextMatrix(RowNum, FG.ColIndex("ColorID"))))
                RSTransDetails("ItemSize").value = IIf((FG.TextMatrix(RowNum, FG.ColIndex("ItemSize")) = ""), "", Trim$(FG.TextMatrix(RowNum, FG.ColIndex("ItemSize"))))
                RSTransDetails("ClassId").value = IIf((FG.TextMatrix(RowNum, FG.ColIndex("ClassId")) = ""), 1, val(FG.TextMatrix(RowNum, FG.ColIndex("ClassId"))))
            
                RSTransDetails("UnitID").value = IIf(FG.cell(flexcpData, RowNum, FG.ColIndex("UnitID")) = "", Null, (FG.cell(flexcpData, RowNum, FG.ColIndex("UnitID"))))
                RSTransDetails("ShowQty").value = IIf((FG.TextMatrix(RowNum, FG.ColIndex("Count")) = ""), Null, val(FG.TextMatrix(RowNum, FG.ColIndex("Count"))))
         RSTransDetails("ProductionDate").value = IIf((FG.TextMatrix(RowNum, FG.ColIndex("ProductionDate")) = ""), Null, Format((FG.TextMatrix(RowNum, FG.ColIndex("ProductionDate"))), "DD/mm/YYYY"))
        RSTransDetails("ExpiryDate").value = IIf((FG.TextMatrix(RowNum, FG.ColIndex("ExpiryDate")) = ""), Null, Format((FG.TextMatrix(RowNum, FG.ColIndex("ExpiryDate"))), "DD/mm/YYYY"))
        RSTransDetails("LotNO").value = IIf(FG.TextMatrix(RowNum, FG.ColIndex("LotNO")) = "", Null, FG.TextMatrix(RowNum, FG.ColIndex("LotNO")))


                Dim RsUnitData As ADODB.Recordset
                Dim LngCurItemID As Long
                Dim LngUnitID As Long
                Dim DblQty As Double
        
                LngCurItemID = val(FG.TextMatrix(RowNum, FG.ColIndex("Code")))
                LngUnitID = val(FG.cell(flexcpData, RowNum, FG.ColIndex("UnitID")))
                DblQty = val(FG.TextMatrix(RowNum, FG.ColIndex("Count")))

                StrSQL = "Select * From TblItemsUnits Where ItemID=" & LngCurItemID
                StrSQL = StrSQL + " AND UnitID=" & LngUnitID
                Set RsUnitData = New ADODB.Recordset
                RsUnitData.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

                If Not (RsUnitData.BOF Or RsUnitData.EOF) Then
                    RSTransDetails("QtyBySmalltUnit").value = RsUnitData("UnitFactor").value
                    RSTransDetails("Quantity").value = RSTransDetails("QtyBySmalltUnit").value * RSTransDetails("showqty").value
                    'RSTransDetails("Price").value = Val(IIf((Fg.TextMatrix(RowNum, Fg.ColIndex("Price")) = ""), Null, Val(Fg.TextMatrix(RowNum, Fg.ColIndex("Price"))))) / RSTransDetails("Quantity").value
                    RSTransDetails("Price").value = val(IIf((FG.TextMatrix(RowNum, FG.ColIndex("Price")) = ""), 0, val(FG.TextMatrix(RowNum, FG.ColIndex("Price"))))) / RSTransDetails("QtyBySmalltUnit").value
                End If


            If SystemOptions.DiscountByQtyOnly = True Then
            
                TotalShahnPerLine = ((((IIf(IsNull(RSTransDetails("price")), 0, 1) * IIf(IsNull(RSTransDetails("Quantity")), 0, RSTransDetails("Quantity")) / (LblTotalQty.Caption))) * val(XPTxtDiscountVal.text)) / IIf(IsNull(RSTransDetails("Quantity")), 0, RSTransDetails("Quantity")))
            
            Else
                If val(LblTotalAll.Caption) * val(XPTxtDiscountVal.text) <> 0 Then
                    TotalShahnPerLine = ((((IIf(IsNull(RSTransDetails("price")), 0, RSTransDetails("price")) * IIf(IsNull(RSTransDetails("Quantity")), 0, RSTransDetails("Quantity")) / (LblTotalAll.Caption))) * val(XPTxtDiscountVal.text)) / IIf(IsNull(RSTransDetails("Quantity")), 0, RSTransDetails("Quantity")))
                Else
                    TotalShahnPerLine = 0
                End If
            End If
Dim Quantity As Double
Dim Price As Double
Dim TotalValue As Double
Dim TotalQty As Double
Dim discountvalue As Double

' «·Õ’Ê· ⁄·Ï «·ﬁÌ„
            Quantity = IIf(IsNull(RSTransDetails("Quantity")), 0, RSTransDetails("Quantity"))
            Price = IIf(IsNull(RSTransDetails("price")), 0, RSTransDetails("price"))
            TotalValue = LblTotalAll.Caption
            TotalQty = LblTotalQty.Caption
            discountvalue = val(XPTxtDiscountVal.text)
            
            '  ÿ»Ìﬁ «·‘—ÿ
            If SystemOptions.DiscountByQtyOnly = False Then
                '  Ê“Ì⁄ «·Œ’„ »‰«¡ ⁄·Ï «·ﬂ„Ì… ›ﬁÿ
                TotalShahnPerLine = ((Quantity / TotalQty) * discountvalue)
            Else
                '  Ê“Ì⁄ «·Œ’„ »‰«¡ ⁄·Ï «·ﬂ„Ì… * «·”⁄—
                TotalShahnPerLine = (((Price * Quantity) / TotalValue) * discountvalue)
            End If
If SystemOptions.DiscountByQtyOnly = True And val(TotalShahnPerLine) <> 0 Then
        RSTransDetails("ItemDiscountType").value = 2
        RSTransDetails("ItemDiscount").value = TotalShahnPerLine
End If



                RSTransDetails.update
            End If

        Next RowNum
        
        Cn.Execute "Update tblItems set DefaultSupplier =  " & DBCboClientName.BoundText & " Where ItemId In (SELECT Item_ID FROM Transaction_Details WHERE Transaction_ID  = " & val(val(XPTxtBillID.text)) & ")"
        Dim sss As String
sss = "Update TblItemsUnits set UnitPurPrice =  "
sss = sss & " (SELECT Top 1 Transaction_Details.ShowPrice FROM Transaction_Details WHERE Transaction_ID  = " & val(val(XPTxtBillID.text)) & " "
sss = sss & " and Transaction_Details.Item_Id =TblItemsUnits.ItemID and TblItemsUnits.UnitId = Transaction_Details.UnitId )"
sss = sss & " Where ItemId In (SELECT Item_ID FROM Transaction_Details WHERE Transaction_ID  = " & val(val(XPTxtBillID.text)) & ")"
Cn.Execute sss


    Closeorders (0)
   Closeorders (1)
   If SystemOptions.PoCreateVoucher = True Then
       createVoucher
      updateNotesValueAndNobytext (val(TXTNoteID.text))
   End If
   SaveValueAdded
   
   
        Cn.CommitTrans
        BeginTrans = False
        XPTxtCurrent.Caption = rs.AbsolutePosition
        XPTxtCount.Caption = rs.RecordCount
        lbl(64).Caption = showLabel(TxtNoteSerial1, oldtxtNoteSerial1)


Accredit_Click
rs.Resync adAffectCurrent
        Select Case Me.TxtModFlg.text

            Case "N"

                If SystemOptions.UserInterface = ArabicInterface Then
                    Msg = "  „ Õ›Ÿ »Ì«‰«  Â–Â «·⁄„·Ì…" & CHR(13)
                    Msg = Msg + "Â·  —€» ›Ì ≈÷«›… »Ì«‰«  √Œ—Ì"
                Else
                    Msg = " Saved Successfully" & CHR(13)
                    Msg = Msg + "do you new Operation?"
        
                End If

                If MsgBox(Msg, vbYesNo + vbQuestion + vbMsgBoxRight + vbMsgBoxRtlReading + vbDefaultButton2, App.Title) = vbYes Then
                    Cmd_Click (0)
                    Screen.MousePointer = vbDefault
                    Exit Sub
                End If
            
            Case "E"

                If SystemOptions.UserInterface = ArabicInterface Then
                    MsgBox " „ Õ›Ÿ Â–Â «· ⁄œÌ·« ", vbInformation + vbOKOnly + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
                Else
                    MsgBox "Saved Changes Successfully", vbInformation + vbOKOnly + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
                End If

        End Select

        TxtModFlg.text = "R"
  FillOrderGrid
  FillOrderGrid2
    End If

    Screen.MousePointer = vbDefault
    Exit Sub
ErrTrap:
    Screen.MousePointer = vbDefault

    If BeginTrans = True Then
        BeginTrans = False
        Cn.RollbackTrans
    End If

    If rs.EditMode <> adEditNone Then
        rs.CancelUpdate
    End If

    If Err.Number = -2147217900 Then
        If SystemOptions.UserInterface = ArabicInterface Then
            Msg = "·« Ì„ﬂ‰ Õ›Ÿ Â–Â «·»Ì«‰«  " & CHR(13)
            Msg = Msg + "·ﬁœ  „ «œŒ«· ﬁÌ„ €Ì— ’«·Õ… " & CHR(13)
            Msg = Msg + " √ﬂœ „‰ œﬁ… «·»Ì«‰«  Ê√⁄œ «·„Õ«Ê·…"
        Else
    
            Msg = "Cant Save Error"
        End If

        MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
        Exit Sub
    End If

    If SystemOptions.UserInterface = ArabicInterface Then
        Msg = "⁄›Ê«...ÕœÀ Œÿ√ „« √À‰«¡ Õ›Ÿ Â–Â «·»Ì«‰«  " & CHR(13)
    Else
        Msg = "Sorry... Error During Saving " & CHR(13)
    End If

    MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
End Sub
Public Function CREATE_VOUCHER_GE(general_noteid As Long, BranchID As Integer, UserID As Long _
, NoteDate As Date)

 Dim Notevalue As Single
    Dim LngDevID As Long
    Dim LngDevNO  As Integer
    Dim StrTempAccountCode As String
    Dim StrTempDes As String
    Dim SngTemp  As Variant
    Dim Account_Code_dynamic As String
    Dim i As Integer
 
 Dim StrSQL As String
 
         StrSQL = "Delete From DOUBLE_ENTREY_VOUCHERS Where Notes_ID=" & general_noteid
        Cn.Execute StrSQL, , adExecuteNoRecords
        

 LngDevNO = 0

    LngDevID = new_id("DOUBLE_ENTREY_VOUCHERS", "Double_Entry_Vouchers_ID", "")
    '«·ÿ—› «·„Ì‰
     
    my_branch = BranchID

  
  
            StrTempDes = " √„— «·‘—«¡ —ﬁ„ " & TxtNoteSerial1 & "  ··„Ê—œ   " & DBCboClientName.text & " „·«ÕŸ«  " & TxtBillComment.text
            LngDevNO = LngDevNO + 1
 
Notevalue = (NewGrid.GetItemsTotal(ItemsGoodType) - val(LblDiscountsTotal.Caption))
  
 Dim Account_Code_dynamic101 As String
  Dim Account_Code_dynamic102 As String
 
   Account_Code_dynamic101 = get_account_code_branch(101, my_branch)
            Account_Code_dynamic102 = get_account_code_branch(102, my_branch)
              
'll:
   LngDevNO = 0
  
  
 If Notevalue > 0 Then
       ' «·„œÌ‰
      
   LngDevNO = LngDevNO + 1
   StrTempAccountCode = Account_Code_dynamic101
  
            If ModAccounts.AddNewDev(LngDevID, LngDevNO, StrTempAccountCode, Notevalue, 0, StrTempDes, general_noteid, , , , NoteDate, UserID, , , , , , , , , , , , , , , , , , BranchID) = False Then
                GoTo ErrTrap
            
            
            End If
            
        LngDevNO = LngDevNO + 1
        StrTempAccountCode = Account_Code_dynamic102
   
              
              
              If ModAccounts.AddNewDev(LngDevID, LngDevNO, StrTempAccountCode, Notevalue, 1, StrTempDes, general_noteid, , , , NoteDate, UserID, , , , , , , , , , , , , , , , , , BranchID) = False Then
                GoTo ErrTrap
            End If
  

            
            
  End If
  
   
 
  
ErrTrap:
End Function

Function createVoucher()
Dim NoteID As Long
Dim NoteDate As Date
Dim NoteSerial As String
Dim Notevalue As Double
Dim des As String
des = "«„— ‘—«— —ﬁ„  " & TxtNoteSerial1 & " „‰ «·„Ê—œ " & DBCboClientName.text
Dim tablename As String
Dim Filedname As String
Dim ContNo As Long
Dim sql As String
tablename = "Transactions"
Filedname = "Transaction_ID"
ContNo = Me.XPTxtBillID.text
Notevalue = LblTotalView.Caption


                     If Me.TxtModFlg = "N" Then
                                'ma  CreateNotes NoteID, (XPDtbBill.value), val(dcBranch.BoundText), 8064, Notevalue, NoteSerial, TxtNoteSerial1, tablename, Filedname, ContNo, des, ToHijriDate(XPDtbBill.value)
                                     TXTNoteID.text = NoteID
                                    TxtNoteSerial.text = NoteSerial
                    Else
                                      If TXTNoteID.text = "" Or TxtNoteSerial.text = "" Then
                                   'ma  CreateNotes NoteID, (XPDtbBill.value), val(dcBranch.BoundText), 8064, Notevalue, NoteSerial, TxtNoteSerial1, tablename, Filedname, ContNo, des, ToHijriDate(XPDtbBill.value)
                                                       TXTNoteID.text = NoteID
                                                  TxtNoteSerial.text = NoteSerial
                                    Else
                                                  sql = "update notes  set Note_Value=" & Notevalue & ",note_value_by_characters='" & WriteNo(val(Notevalue), 0, True) & "'"
                                                  sql = sql & ",remark='" & TxtNoteSerial1 & "'"
                                                    sql = sql & " where NoteID=" & val(TXTNoteID.text)
                                                     Cn.Execute sql
                                                     
                                       End If
                         
                    End If
ReLineGrid
CREATE_VOUCHER_GE val(TXTNoteID.text), val(dcBranch.BoundText), user_id, XPDtbBill.value



End Function



Private Sub XPBtnNewClients_Click()

    'With FrmAddNewCustemer
    '    .DealingForm = ShowPrice
    '    .show vbModal
    '    .Caption = "≈÷«›… ⁄„Ì· ÃœÌœ"
    '    .lbl(1).Caption = "ﬂÊœ «·⁄„Ì·"
    '    .lbl(0).Caption = "«”„ «·⁄„Ì·"
    'End With

End Sub

Private Sub XPChkTAX_Click()
    On Error GoTo ErrTrap

    If XPChkTAX.value = Checked Then
        XPTxtTaxValue.Enabled = True
        XPTxtTaxValue.locked = False
        lbl(4).Enabled = True
    Else
        XPTxtTaxValue.text = ""
        XPTxtTaxValue.Enabled = False
        lbl(4).Enabled = False
    End If

    Exit Sub
ErrTrap:
End Sub
Function print_report(Optional NoteSerial As String)
     
    Dim MySQL          As String
    Dim RsData         As New ADODB.Recordset
    Dim xApp           As New CRAXDRT.Application
    Dim xReport        As CRAXDRT.Report
    Dim CViewer        As ClsReportViewer
    Dim StrReportTitle As String
    Dim StrFileName    As String
    Dim Msg            As String

    MySQL = MySQL & "    SELECT        dbo.Transaction_Details.Vat, dbo.Transaction_Details.Vatyo, dbo.Transactions.SippingDate, dbo.Transactions.DeliverDate, dbo.Transactions.NoteSerial, dbo.Transactions.NoteSerial1, dbo.Transactions.Transaction_ID,"
    MySQL = MySQL & "                            dbo.Transaction_Details.ItemDiscountType, dbo.Transaction_Details.ItemDiscount, dbo.Transactions.order_no, dbo.Transaction_Details.Item_ID, dbo.Transaction_Details.ShowQty, dbo.Transaction_Details.showPrice,"
    MySQL = MySQL & "                            dbo.Transaction_Details.ItemSize, dbo.Transaction_Details.ColorID, dbo.Transaction_Details.UnitId, dbo.Transaction_Details.ClassId, dbo.TblItems.ItemCode, dbo.TblItems.ItemName, dbo.TblItems.ItemNamee,"
    MySQL = MySQL & "                            dbo.TblItemsColors.ColorName, dbo.TblItemsSizes.SizeName, dbo.TblUnites.UnitName, dbo.TblItemsclasses.SizeName AS ClassName, dbo.Transactions.Transaction_Date, dbo.TblCustemers.CusName,"
    MySQL = MySQL & "                            dbo.TblCustemers.CusNamee,dbo.TblCustemers.ResponsibleContact, dbo.Transactions.Transaction_Type, dbo.TblCustemers.Fullcode, dbo.TblCustemers.Fullcode as CusFullCode, dbo.TblCustemers.E_mail, dbo.TblCustemers.Cus_Phone, dbo.TblCustemers.Cus_mobile, dbo.TblCustemers.FaxNumber,"
    MySQL = MySQL & "                            dbo.Transaction_Details.ParrtNoCode, dbo.TblUnites.UnitNamee, dbo.Transactions.ModeReceptEq, dbo.Transactions.ModeSupply, dbo.Transactions.DeptID, dbo.TblEmpDepartments.DepartmentName,"
    MySQL = MySQL & "                            dbo.TblEmpDepartments.DepartmentNamee, dbo.Transactions.PaymentType, dbo.Transaction_Details.ID AS IDTr, dbo.Transactions.PaymentT, dbo.Transactions.ShipingID, dbo.TblShipingData.Name AS ShipName,"
    MySQL = MySQL & "                            dbo.TblShipingData.NameE AS ShipNameE, dbo.Transactions.PaymentID, dbo.TblPaymetData.Name AS PaymName, dbo.TblPaymetData.NameE AS PaymNameE, dbo.Transactions.Shipping_Pos,"
    MySQL = MySQL & "                          dbo.Transaction_Details.LotNO, dbo.Transaction_Details.ExpiryDate, dbo.Transaction_Details.ProductionDate, dbo.Transactions.Currency_id, dbo.currency.name AS Currname, dbo.currency.code AS Currcode,"
    MySQL = MySQL & "                           dbo.currency.nameE AS CurrnameE , Transactions.MrNo, dbo.Transactions.Trans_Discount, dbo.Transactions.Trans_DiscountType, dbo.Transactions.chkTaxExempt ,TblStore.StoreName,TblStore.StoreNamee,Transactions.TransactionComment,TblItems.ItemComment,TblCustemers.VATNO,"
    MySQL = MySQL & "                           projects.*,Transactions.ContainerNo  ,Transactions.VAT VatValue"
    MySQL = MySQL & "   FROM            dbo.TblItemsSizes RIGHT OUTER JOIN"
    MySQL = MySQL & "                           dbo.TblUnites RIGHT OUTER JOIN"
    MySQL = MySQL & "                            dbo.currency RIGHT OUTER JOIN"
    MySQL = MySQL & "                            dbo.Transactions INNER JOIN"
    MySQL = MySQL & "                            dbo.Transaction_Details ON dbo.Transactions.Transaction_ID = dbo.Transaction_Details.Transaction_ID INNER JOIN"
    MySQL = MySQL & "                            dbo.TblItems ON dbo.Transaction_Details.Item_ID = dbo.TblItems.ItemID Left outer JOIN"
    MySQL = MySQL & "                            dbo.TblItemsclasses ON dbo.Transaction_Details.ClassId = dbo.TblItemsclasses.SizeId ON dbo.currency.id = dbo.Transactions.Currency_id LEFT OUTER JOIN"
    MySQL = MySQL & "                            dbo.TblPaymetData ON dbo.Transactions.PaymentID = dbo.TblPaymetData.ID LEFT OUTER JOIN"
    MySQL = MySQL & "                            dbo.TblShipingData ON dbo.Transactions.ShipingID = dbo.TblShipingData.ID LEFT OUTER JOIN"
    MySQL = MySQL & "                            dbo.TblEmpDepartments ON dbo.Transactions.DeptID = dbo.TblEmpDepartments.DeparmentID LEFT OUTER JOIN"
    MySQL = MySQL & "                            dbo.TblCustemers ON dbo.Transactions.CusID = dbo.TblCustemers.CusID ON dbo.TblUnites.UnitID = dbo.Transaction_Details.UnitId ON dbo.TblItemsSizes.SizeId = dbo.Transaction_Details.ItemSize LEFT OUTER JOIN"
    MySQL = MySQL & "                            dbo.TblItemsColors ON dbo.Transaction_Details.ColorID = dbo.TblItemsColors.ColorID"
    MySQL = MySQL & "                                LEFT OUTER JOIN"
    MySQL = MySQL & "                                dbo.TblStore ON dbo.Transactions.StoreID = dbo.TblStore.StoreID"
    MySQL = MySQL & "                                Left outer join projects  ON dbo.Transaction_Details.project_id1 = dbo.projects.ID"
    
    MySQL = MySQL & "   WHERE      (dbo.Transactions.Transaction_ID =" & val(XPTxtBillID.text) & ")"

    If SystemOptions.UserInterface = ArabicInterface Then
        '      StrFileName = App.path & "\Reports\Inventory\PerformaInvoices7Sh.rpt"
        StrFileName = App.path & "\Special\" & SystemOptions.Reportpath & "\PerformaInvoices7Sh.rpt"
    Else
        '   StrFileName = App.path & "\Reports\Inventory\PerformaInvoices7Sh.rpt"
        StrFileName = App.path & "\Special\" & SystemOptions.Reportpath & "\PerformaInvoices7ShEng.rpt"
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
        MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
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
        'If Me.DTPickerAccFrom.value <> Empty Or Me.DTPickerAccFrom.value <> Null Then
        '    StrReportTitle = StrReportTitle + " »œ«Ì… „‰ " & Format(Me.DTPickerAccFrom.value, "yyyy/M/d") & ""
        'End If
        'If Me.DTPickerAccTo.value <> Empty Or Me.DTPickerAccTo.value <> Null Then
        '    StrReportTitle = StrReportTitle + " ≈·Ï " & Format(Me.DTPickerAccTo.value, "yyyy/M/d") & " "
        'End If
        xReport.ParameterFields(12).AddCurrentValue TxtBillComment.text  'RPTCompany_Name_Arabic
        'TxtBillComment
        
    Else
 
        xReport.ParameterFields(1).AddCurrentValue cCompanyInfo.ArabCompanyName ' RPTCompany_Name_Eng
        '  xReport.ParameterFields(2).AddCurrentValue RPTComment_Eng
        xReport.ParameterFields(4).AddCurrentValue get_branch_name(val(my_branch))
        StrReportTitle = ""
        'If Me.DTPickerAccFrom.value <> Empty Or Me.DTPickerAccFrom.value <> Null Then
        '    StrReportTitle = StrReportTitle + " From Date " & (Me.DTPickerAccFrom.value) & ""
        'End If
        'If Me.DTPickerAccTo.value <> Empty Or Me.DTPickerAccTo.value <> Null Then
        '    StrReportTitle = StrReportTitle + " To Date :  " & (Me.DTPickerAccTo.value) & ""
        'End If
    End If

    xReport.ParameterFields(3).AddCurrentValue user_name
    xReport.ParameterFields(9).AddCurrentValue LblTotalAll.Caption
    xReport.ParameterFields(10).AddCurrentValue LblDiscountsTotalView.Caption
    xReport.ParameterFields(11).AddCurrentValue LblTotalView.Caption
    '    xReport.ParameterFields(4).AddCurrentValue WriteNo(Format(val(TxtAdvanceValue.text), "0.00"), 0, True, ".")
    '
    '   xReport.ParameterFields(7).AddCurrentValue DBIssueDate.value
    ' xReport.ParameterFields(8).AddCurrentValue IIf(IsNumeric(Fg.TextMatrix(Me.Fg.FixedRows, Fg.ColIndex("PartValue"))), val(Fg.TextMatrix(Me.Fg.FixedRows, Fg.ColIndex("PartValue"))), 0)
    ' xReport.ParameterFields(9).AddCurrentValue val(lbl(22).Caption)
    '  xReport.ParameterFields(10).AddCurrentValue val(TxtDiscount.text)
    '  xReport.ParameterFields(11).AddCurrentValue txtDiscountDES.text
   
    '    xReport.ParameterFields(5).AddCurrentValue ToHijriDate(RsData("notedate").value)
    xReport.reporttitle = StrReportTitle
    xReport.EnableParameterPrompting = False
    xReport.ApplicationName = App.Title
    xReport.ReportAuthor = App.Title
    Set CViewer = New ClsReportViewer
    CViewer.FireReport xReport, WindowTarget, "", , , , StrFileName, , MySQL

    RsData.Close
    Set RsData = Nothing
    Screen.MousePointer = vbDefault
 
End Function
Private Sub PrintReport()
    On Error GoTo ErrTrap

    If XPTxtBillID.text <> "" Then
        Set SaleReport = New ClsSaleReport
        SaleReport.ShowPrice XPTxtBillID.text, 7, DcboEmp.text
    End If

    Exit Sub
ErrTrap:
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, _
                             UnloadMode As Integer)
    Dim StrMSG As String
    Dim IntResult As String
    On Error GoTo ErrTrap

    If Me.TxtModFlg.text <> "R" Then

        Select Case Me.TxtModFlg.text

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

Public Sub Cala()
    NewGrid.Calculate 1
End Sub

Private Sub XPDtbBill_Change()

    If Trim(TxtNoteSerial1.text) <> "" Then
        oldtxtNoteSerial1.text = TxtNoteSerial1.text
    End If

    TxtNoteSerial1.text = ""
 TxtNoteSerial.text = ""
 
End Sub

Private Sub XPTxtTaxValue_Change()

    If Me.TxtModFlg.text = "N" Or Me.TxtModFlg.text = "E" Then
        NewGrid.Calculate 1, , , True
    End If

End Sub
Private Sub ChangeLang()
    Dim XPic As IPictureDisp
    Set XPic = Me.XPBtnMove(1).ButtonImage
    Set Me.XPBtnMove(1).ButtonImage = Me.XPBtnMove(2).ButtonImage
    Set Me.XPBtnMove(2).ButtonImage = XPic
    Set XPic = Me.XPBtnMove(0).ButtonImage
    Set Me.XPBtnMove(0).ButtonImage = Me.XPBtnMove(3).ButtonImage
    Set Me.XPBtnMove(3).ButtonImage = XPic
    chkshipped.Caption = "shipped"
    Me.Caption = "Purchase Order"
    lbl(48).Caption = "Project"
    Me.Label1(1).Caption = "Purchase Order"
    Me.XPTab301.TabCaption(0) = "Items"
    Me.XPTab301.TabCaption(1) = "Approved Status"
    Me.XPTab301.TabCaption(2) = "Quotations"
    Me.XPTab301.TabCaption(3) = "Internal Orders"
    lbl(18).Caption = "Type"
    Label4.Caption = "ACC. BY"
    Label10.Caption = "Signature"
    lbl(32).Caption = "Sales Person"
    Accredit.Caption = "Accredit"
    Cmd(8).Caption = "Print Pur. Order"
    lbl(50).Caption = "Discounts"
    lbl(49).Caption = "Net"
    ''''''''''''''''''''''''''''
    ISButton2.Caption = "Attachments"
    lbl(35).Caption = "Offer End"
    lbl(42).Caption = "Shipment Date"
    lbl(43).Caption = "Delivery Date"
    lbl(45).Caption = "Delivery Date"
    lbl(47).Caption = "Entry NO."
    lbl(44).Caption = "internal Request"
    lbl(37).Caption = "Management"
    lbl(36).Caption = "Supplier Name"
    lbl(35).Caption = "Value"
    lbl(41).Caption = "Shipping Method "
    lbl(34).Caption = "Discount Type"
    Cmd(12).Caption = "Print"
        Me.XPTab301.TabCaption(4) = "VAT"
Label22.Caption = "Data of VAT"
lbl(104).Caption = "Total"
lbl(52).Caption = "VAT"
lbl(53).Caption = "Total"
ChecVAT.RightToLeft = False
ChecVAT.Caption = "Select All"
With VatGrid
.TextMatrix(0, .ColIndex("index")) = "Serial"
.TextMatrix(0, .ColIndex("select")) = "Select"
.TextMatrix(0, .ColIndex("Code")) = "Item Code"
.TextMatrix(0, .ColIndex("Name")) = "Item Name"
.TextMatrix(0, .ColIndex("Vatyo")) = "Percentage"
.TextMatrix(0, .ColIndex("Vat")) = "Value"
.TextMatrix(0, .ColIndex("Valu")) = "Item Value"
End With

    '''''''''''''''''''''''''
     With XPCboDiscountType
            .Clear
            .AddItem "No Discount"
            .AddItem "Value Discount"
            .AddItem "Precetage Discount"
        End With


    lbl(5).Caption = "Ord/P INV. No"
    Frame3.Caption = "LC Data"
    ISButton1.Caption = "View"
    lbl(25).Caption = "Total"
    lbl(63).Caption = "Qty"
    Label2.Caption = "Branch"
    lbl(6).Caption = "Date"
    lbl(7).Caption = "Supplier"
    lbl(8).Caption = "Store"
    lbl(9).Caption = "Type"
    lbl(10).Caption = "Cost Center"
    Label14(0).Caption = "Q.Ref"
    chkTaxExempt.Caption = "With out vat"
    cmdSelectFile.Caption = "Select File"
    cmdLoadFile.Caption = "Load"
    lbl(11).Caption = "Project"
    lbl(16).Caption = "Article Section"
    lbl(12).Caption = "Currency"
    lbl(13).Caption = "Country"
    lbl(51).Caption = "Operation"
    lbl(46).Caption = "Item"
    lbl(14).Caption = "Shipment Mode"
    lbl(17).Caption = "Value"
    lbl(15).Caption = "Payment M"
    lbl(28).Caption = "  Remarks"
    lbl(19).Caption = "Kind Of Order"
    lbl(24).Caption = "Expiry Date"
    lbl(20).Caption = "Credit Bank"
    lbl(21).Caption = "Credit Curr."
    lbl(22).Caption = "Credit No."
    lbl(23).Caption = "Value"
    ChAuto.Caption = "Auto"
    'ISButton1.Caption = "Show Port Data"
   ' Label1.Caption = "LC NO:"
    'Label2.Caption = "Supp info No."
    Label3.Caption = "Supp info Date"
    Label5.Caption = "Exp Del Date"
    Label6.Caption = "Act Del Date"
    Label7.Caption = "Late Date"
    Label8.Caption = "Exp Arrival Date"
    Label9.Caption = "Comments"

    lbl(31).Caption = "Item Code"
    lbl(30).Caption = "item name"

    lbl(29).Caption = "Status"
    lbl(27).Caption = "Qty"
    lbl(26).Caption = "Price"

    lbl(3).Caption = "Total"
    lbl(1).Caption = "By"
    lbl(0).Caption = "Currenr rec."
    lbl(2).Caption = "Total rec."

    Cmd(0).Caption = "New"
    Cmd(1).Caption = "Edit"
    Cmd(2).Caption = "Save"
    Cmd(3).Caption = "Undo"
    Cmd(4).Caption = "Delete"
    Cmd(5).Caption = "Search"
    Cmd(7).Caption = "Print"
    Cmd(10).Caption = "Print"
    Cmd(11).Caption = "Entry Print "
    Cmd(6).Caption = "Exit"
    CmdHelp.Caption = "Help"

    CmdConvert.Caption = "Convert To Bill"
    CmdTemplate.Caption = "Insert template"

    With Me.Grid3
        .TextMatrix(0, .ColIndex("Select")) = "Select"
        .TextMatrix(0, .ColIndex("order_no")) = "order_no"
        .TextMatrix(0, .ColIndex("Transaction_Date")) = "Transaction_Date"
        .TextMatrix(0, .ColIndex("BranchName")) = "BranchNo"
        .TextMatrix(0, .ColIndex("StoreName")) = "StoreName"
    End With
    With Me.Grid2
        .TextMatrix(0, .ColIndex("Approved")) = "Approved"
        .TextMatrix(0, .ColIndex("levelName")) = "Level Name"
        .TextMatrix(0, .ColIndex("EmpName")) = "Emp Name"
        .TextMatrix(0, .ColIndex("ApprovDate")) = "Approv Date"
        .TextMatrix(0, .ColIndex("Remarks")) = "Remarks"
    End With
    With Me.FG
        .TextMatrix(0, .ColIndex("FoxyNo")) = "Program NO."
        .TextMatrix(0, .ColIndex("Shipping")) = "Shipping"
        .TextMatrix(0, .ColIndex("pand")) = "Item"
        .TextMatrix(0, .ColIndex("project")) = "Project"
        .TextMatrix(0, .ColIndex("opera")) = "Operation"
    End With
    With Me.VSFlexGrid1
        .TextMatrix(0, .ColIndex("Select")) = "Select"
        .TextMatrix(0, .ColIndex("order_no")) = "order_no"
        .TextMatrix(0, .ColIndex("Transaction_Date")) = "Transaction_Date"
        .TextMatrix(0, .ColIndex("BranchName")) = "BranchNo"
        .TextMatrix(0, .ColIndex("StoreName")) = "StoreName"
    End With
    With CboPayMentType
            .Clear
            .AddItem "Cash"
            .AddItem "Credit"
        End With
 
 
End Sub

Function Closeorders(Optional gridno As Integer = 0)
   ' On Error Resume Next
    Dim i As Integer
    Dim sql As String
 
     
If gridno = 0 Then
    With Grid3

        For i = 1 To .rows - 1
     
            If .cell(flexcpChecked, i, .ColIndex("select")) = flexChecked Then
           
                sql = "update transactions set closed=1" & ",nots=" & val(Me.XPTxtBillID.text) & ",nots2=" & Me.TxtNoteSerial1.text & " where  Transaction_ID= " & val(.TextMatrix(i, .ColIndex("Transaction_ID")))
            Else
                sql = "update transactions set closed=0 ,nots='' ,nots2='' where  Transaction_ID=" & val(.TextMatrix(i, .ColIndex("Transaction_ID")))
               
            End If
       
            Cn.Execute sql
 
        Next
       
    End With
       
  Else
  
  With VSFlexGrid1

        For i = 1 To .rows - 1
     
            If .cell(flexcpChecked, i, .ColIndex("select")) = flexChecked Then
           
                sql = "update transactions set closed=1" & ",nots=" & val(Me.XPTxtBillID.text) & ",nots2='" & Me.TxtNoteSerial1.text & "'  where  Transaction_ID= " & val(.TextMatrix(i, .ColIndex("Transaction_ID")))
            Else
                sql = "update transactions set closed=0 ,nots='' ,nots2='' where  Transaction_ID=" & val(.TextMatrix(i, .ColIndex("Transaction_ID")))
               
            End If
       
            Cn.Execute sql
 
        Next
       
    End With
   
  End If
End Function
Function FillOrderGrid2()
    ' ⁄»∆… «Ê«„— «·‘—«¡ Ê «·»Ì⁄

    With Me.VSFlexGrid1
        .rows = .FixedRows
        .ExtendLastCol = True
        .RowHeightMin = 300
        .Editable = flexEDKbdMouse
        .ExplorerBar = flexExSortShowAndMove

        '    .AutoSize 0, .Cols - 1, False
    End With

    Dim i As Integer
    Dim RsExp As ADODB.Recordset
    Dim My_SQL As String

    Set RsExp = New ADODB.Recordset
  '  My_SQL = "SELECT dbo.Transactions.closed,dbo.Transactions.Transaction_ID,dbo.Transactions.NoteSerial1 , dbo.Transactions.Transaction_Date,dbo.Transactions.CusID, dbo.TblCustemers.CusName FROM dbo.Transactions  INNER JOIN dbo.TblCustemers ON dbo.Transactions.CusID = dbo.TblCustemers.CusID Where  Transaction_Type=38  AND  dbo.Transactions.Approved = 1   AND CLOSED= 0 and   dbo.TblCustemers.CusID=" & val(DBCboClientName.BoundText)

If Me.TxtModFlg = "N" Then
My_SQL = "SELECT dbo.Transactions.Transaction_ID,    dbo.Transactions.NoteSerial1, dbo.Transactions.Transaction_Date, dbo.TblStore.StoreName, dbo.TblStore.StoreNamee, dbo.TblBranchesData.branch_name, "
My_SQL = My_SQL & "  dbo.TblBranchesData.branch_nameE , dbo.TblBranchesData.branch_id, dbo.Transactions.Closed, dbo.Transactions.Approved"
My_SQL = My_SQL & " FROM         dbo.TblBranchesData INNER JOIN"
My_SQL = My_SQL & " dbo.Transactions INNER JOIN"
My_SQL = My_SQL & " dbo.TblStore ON dbo.Transactions.StoreID = dbo.TblStore.StoreID ON dbo.TblBranchesData.branch_id = dbo.Transactions.BranchId"
'My_SQL = My_SQL & " WHERE     (dbo.Transactions.Transaction_Type = 38) AND (dbo.Transactions.Approved = 1)  and CLOSED= 0 and   dbo.Transactions.CusID=" & val(DBCboClientName.BoundText)
'My_SQL = My_SQL & " WHERE     (dbo.Transactions.Transaction_Type = 38) AND (dbo.Transactions.Approved = 1)  and CLOSED= 0" '  and   dbo.Transactions.CusID=" & val(DBCboClientName.BoundText)
My_SQL = My_SQL & " WHERE     (dbo.Transactions.Transaction_Type = 38 and OrderType=3 )   and CLOSED= 0  " '  and   dbo.Transactions.CusID=" & val(DBCboClientName.BoundText)
Else
My_SQL = "SELECT dbo.Transactions.Transaction_ID,    dbo.Transactions.NoteSerial1, dbo.Transactions.Transaction_Date, dbo.TblStore.StoreName, dbo.TblStore.StoreNamee, dbo.TblBranchesData.branch_name, "
My_SQL = My_SQL & "  dbo.TblBranchesData.branch_nameE , dbo.TblBranchesData.branch_id, dbo.Transactions.Closed, dbo.Transactions.Approved"
My_SQL = My_SQL & " FROM         dbo.TblBranchesData INNER JOIN"
My_SQL = My_SQL & " dbo.Transactions INNER JOIN"
My_SQL = My_SQL & " dbo.TblStore ON dbo.Transactions.StoreID = dbo.TblStore.StoreID ON dbo.TblBranchesData.branch_id = dbo.Transactions.BranchId"
 'My_SQL = My_SQL & " WHERE     (dbo.Transactions.Transaction_Type = 38) AND (dbo.Transactions.Approved = 1)  and CLOSED= 0 "
My_SQL = My_SQL & "  WHERE    nots ='" & val(Me.XPTxtBillID.text) & "'"

End If

    RsExp.Open My_SQL, Cn, adOpenKeyset, adLockReadOnly, adCmdText

    With Me.VSFlexGrid1
        .rows = 2
        .Clear flexClearScrollable

        If RsExp.RecordCount > 0 Then
            .rows = RsExp.RecordCount + 1
            RsExp.MoveFirst

            For i = 1 To .rows - 1
                .TextMatrix(i, .ColIndex("Select")) = IIf(IsNull(RsExp.Fields("closed").value), 0, RsExp.Fields("closed").value)
         
                .TextMatrix(i, .ColIndex("order_no")) = IIf(IsNull(RsExp.Fields("NoteSerial1").value), "", RsExp.Fields("NoteSerial1").value)
               
                .TextMatrix(i, .ColIndex("Transaction_Date")) = IIf(IsNull(RsExp.Fields("Transaction_Date").value), "", RsExp.Fields("Transaction_Date").value)
           

                .TextMatrix(i, .ColIndex("Transaction_ID")) = IIf(IsNull(RsExp.Fields("Transaction_ID").value), "", RsExp.Fields("Transaction_ID").value)

If SystemOptions.UserInterface = ArabicInterface Then
   .TextMatrix(i, .ColIndex("StoreName")) = IIf(IsNull(RsExp.Fields("StoreName").value), "", RsExp.Fields("StoreName").value)
    .TextMatrix(i, .ColIndex("BranchName")) = IIf(IsNull(RsExp.Fields("branch_name").value), "", RsExp.Fields("branch_name").value)
                                
Else
.TextMatrix(i, .ColIndex("StoreName")) = IIf(IsNull(RsExp.Fields("StoreNamee").value), "", RsExp.Fields("StoreNamee").value)
    .TextMatrix(i, .ColIndex("BranchName")) = IIf(IsNull(RsExp.Fields("branch_namee").value), "", RsExp.Fields("branch_namee").value)
                                
End If

                RsExp.MoveNext
            Next

            RsExp.Close
        End If

        .RowHeight(-1) = 300
    End With

    'GRID2.Visible = True

End Function


Function FillOrderGrid()
    ' ⁄»∆… «Ê«„— «·‘—«¡ Ê «·»Ì⁄

    With Me.Grid3
        .rows = .FixedRows
        .ExtendLastCol = True
        .RowHeightMin = 300
        .Editable = flexEDKbdMouse
        .ExplorerBar = flexExSortShowAndMove

        '    .AutoSize 0, .Cols - 1, False
    End With

    Dim i As Integer
    Dim RsExp As ADODB.Recordset
    Dim My_SQL As String

    Set RsExp = New ADODB.Recordset
  '  My_SQL = "SELECT dbo.Transactions.closed,dbo.Transactions.Transaction_ID,dbo.Transactions.NoteSerial1 , dbo.Transactions.Transaction_Date,dbo.Transactions.CusID, dbo.TblCustemers.CusName FROM dbo.Transactions  INNER JOIN dbo.TblCustemers ON dbo.Transactions.CusID = dbo.TblCustemers.CusID Where  Transaction_Type=38  AND  dbo.Transactions.Approved = 1   AND CLOSED= 0 and   dbo.TblCustemers.CusID=" & val(DBCboClientName.BoundText)

If Me.TxtModFlg = "N" Then
My_SQL = "SELECT dbo.Transactions.Transaction_ID,    dbo.Transactions.NoteSerial1, dbo.Transactions.Transaction_Date, dbo.TblStore.StoreName, dbo.TblStore.StoreNamee, dbo.TblBranchesData.branch_name, "
My_SQL = My_SQL & "  dbo.TblBranchesData.branch_nameE , dbo.TblBranchesData.branch_id, dbo.Transactions.Closed, dbo.Transactions.Approved"
My_SQL = My_SQL & " FROM         dbo.TblBranchesData INNER JOIN"
My_SQL = My_SQL & " dbo.Transactions INNER JOIN"
My_SQL = My_SQL & " dbo.TblStore ON dbo.Transactions.StoreID = dbo.TblStore.StoreID ON dbo.TblBranchesData.branch_id = dbo.Transactions.BranchId"
'My_SQL = My_SQL & " WHERE     (dbo.Transactions.Transaction_Type = 46) AND (dbo.Transactions.Approved = 1)  and CLOSED= 0 and   dbo.Transactions.CusID=" & val(DBCboClientName.BoundText)
My_SQL = My_SQL & " WHERE     (dbo.Transactions.Transaction_Type = 46)   and CLOSED= 0 and   dbo.Transactions.CusID=" & val(DBCboClientName.BoundText)

Else
My_SQL = "SELECT dbo.Transactions.Transaction_ID,    dbo.Transactions.NoteSerial1, dbo.Transactions.Transaction_Date, dbo.TblStore.StoreName, dbo.TblStore.StoreNamee, dbo.TblBranchesData.branch_name, "
My_SQL = My_SQL & "  dbo.TblBranchesData.branch_nameE , dbo.TblBranchesData.branch_id, dbo.Transactions.Closed, dbo.Transactions.Approved"
My_SQL = My_SQL & " FROM         dbo.TblBranchesData INNER JOIN"
My_SQL = My_SQL & " dbo.Transactions INNER JOIN"
My_SQL = My_SQL & " dbo.TblStore ON dbo.Transactions.StoreID = dbo.TblStore.StoreID ON dbo.TblBranchesData.branch_id = dbo.Transactions.BranchId"
'My_SQL = My_SQL & " WHERE     (dbo.Transactions.Transaction_Type = 38) AND (dbo.Transactions.Approved = 1)  and CLOSED= 0 "
My_SQL = My_SQL & "  WHERE    nots ='" & val(Me.XPTxtBillID.text) & "'"

End If

    RsExp.Open My_SQL, Cn, adOpenKeyset, adLockReadOnly, adCmdText

    With Me.Grid3
        .rows = 2
        .Clear flexClearScrollable

        If RsExp.RecordCount > 0 Then
            .rows = RsExp.RecordCount + 1
            RsExp.MoveFirst

            For i = 1 To .rows - 1
                .TextMatrix(i, .ColIndex("Select")) = IIf(IsNull(RsExp.Fields("closed").value), 0, RsExp.Fields("closed").value)
         
                .TextMatrix(i, .ColIndex("order_no")) = IIf(IsNull(RsExp.Fields("NoteSerial1").value), "", RsExp.Fields("NoteSerial1").value)
               
                .TextMatrix(i, .ColIndex("Transaction_Date")) = IIf(IsNull(RsExp.Fields("Transaction_Date").value), "", RsExp.Fields("Transaction_Date").value)
           

                .TextMatrix(i, .ColIndex("Transaction_ID")) = IIf(IsNull(RsExp.Fields("Transaction_ID").value), "", RsExp.Fields("Transaction_ID").value)

If SystemOptions.UserInterface = ArabicInterface Then
   .TextMatrix(i, .ColIndex("StoreName")) = IIf(IsNull(RsExp.Fields("StoreName").value), "", RsExp.Fields("StoreName").value)
    .TextMatrix(i, .ColIndex("BranchName")) = IIf(IsNull(RsExp.Fields("branch_name").value), "", RsExp.Fields("branch_name").value)
                                
Else
.TextMatrix(i, .ColIndex("StoreName")) = IIf(IsNull(RsExp.Fields("StoreNamee").value), "", RsExp.Fields("StoreNamee").value)
    .TextMatrix(i, .ColIndex("BranchName")) = IIf(IsNull(RsExp.Fields("branch_namee").value), "", RsExp.Fields("branch_namee").value)
                                
End If

                RsExp.MoveNext
            Next

            RsExp.Close
        End If

        .RowHeight(-1) = 300
    End With

    Grid2.Visible = True

End Function




Private Sub GetFieldID(ByVal mTableColName As String, ByVal mRow As Long, ByVal mCol As Long, ByVal mGrid As Object, Optional ByVal MainTableName As String = "")
    Dim mTableName As String
    Dim mFieldIDName As String
    Dim mFieldName As String
    Dim xx As Variant
    Dim mValue As String
    Dim rsDummy As New ADODB.Recordset
    Dim rsDummy2 As New ADODB.Recordset
    If mCol = 67 Then
        mCol = 67
    End If
    If mGrid.ColKey(mCol) = "NationlID" Then
        mCol = mCol
    End If

End Sub

Private Function SearchInGrid(ByVal mGrd As Object, ByVal mTxt As String, ByVal mFldName As String) As String
Dim i As Long
For i = 1 To mGrd.rows - 1
    If Trim(mGrd.TextMatrix(i, mGrd.ColIndex(mFldName))) = mTxt Then
        SearchInGrid = i
        Exit Function
    End If
Next
SearchInGrid = ""
End Function


Private Sub GetIDCombo(ByVal mTableColID As String, ByVal mRow As Long, ByVal mCol As Long, ByVal mGrid As Object)
Dim mTxt As String
mTxt = Trim(mGrid.TextMatrix(mRow, mCol - 1))
Select Case mTableColID
Case "sexID"
    If mTxt = "Male" Or mTxt = "–ﬂ—" Then
        mTxt = 1
    Else
        mTxt = 2
    End If
Case "MaritalStatusID"
'    DcbMatrial.AddItem "√⁄“»"
'      DcbMatrial.AddItem "„ “ÊÃ"
    If mTxt = "√⁄“»" Or mTxt = "Single" Then
        mTxt = 0
    ElseIf mTxt = "„ “ÊÃ" Or UCase(mTxt) = "MARRIED" Then
        mTxt = 1
    ElseIf mTxt = "„ÿ·ﬁ/„ÿ·›…" Or UCase(mTxt) = "DIVORCED" Then
        mTxt = 2
    ElseIf mTxt = "«—„·/√—„·…" Or UCase(mTxt) = "WIDOWED" Then
        mTxt = 3
        
    End If
Case "Emp_Name1.Emp_Name2.Emp_Name3.Emp_Name4"
    mTxt = mGrid.TextMatrix(mRow, mCol - 4) + " " + mGrid.TextMatrix(mRow, mCol - 3) + " " + mGrid.TextMatrix(mRow, mCol - 2) + " " + mGrid.TextMatrix(mRow, mCol - 1)
Case ""
End Select
mGrid.TextMatrix(mRow, mCol) = mTxt
End Sub



Public Function CheckDateIsHij(ByVal mDate As String) As Integer
    If Not IsDate(mDate) Then CheckDateIsHij = 3: Exit Function
    
    If Trim(mDate) = "" Then CheckDateIsHij = 3: Exit Function
    
    If year(mDate) < 1800 Then
        CheckDateIsHij = 1
    Else
        CheckDateIsHij = 2
    End If
End Function


