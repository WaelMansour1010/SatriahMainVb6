VERSION 5.00
Object = "{C115893A-A3BF-43AF-B28D-69DB846077F3}#1.0#0"; "vsflex8u.ocx"
Object = "{0AFE7BE0-11B7-4A3E-978D-D4501E9A57FE}#1.0#0"; "c1sizer.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{FE5DCFAD-BC1D-11D2-94CF-004005455FAA}#1.4#0"; "ImpulseButton.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{D155F1AE-D9A4-458C-8CEE-498CB717DB7B}#1.0#0"; "DBPix20.ocx"
Begin VB.Form FrmShowPriceOriginal 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "«Ê«„— «·»Ì⁄ Ê «·‘—«¡ Ê«·›Ê« Ì— «·„»œ∆Ì…"
   ClientHeight    =   9150
   ClientLeft      =   -15
   ClientTop       =   375
   ClientWidth     =   13605
   HelpContextID   =   340
   Icon            =   "FrmShowPrice.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   RightToLeft     =   -1  'True
   ScaleHeight     =   9150
   ScaleWidth      =   13605
   Begin C1SizerLibCtl.C1Elastic EleMain 
      Height          =   9150
      Left            =   0
      TabIndex        =   16
      TabStop         =   0   'False
      Top             =   0
      Width           =   13605
      _cx             =   23998
      _cy             =   16140
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
      AutoSizeChildren=   0
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
      Begin C1SizerLibCtl.C1Elastic Ele 
         Height          =   555
         Index           =   1
         Left            =   15
         TabIndex        =   17
         TabStop         =   0   'False
         Top             =   8580
         Width           =   16080
         _cx             =   28363
         _cy             =   979
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
            Height          =   390
            Index           =   0
            Left            =   12090
            TabIndex        =   18
            Top             =   90
            Width           =   930
            _ExtentX        =   1640
            _ExtentY        =   688
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
            Height          =   390
            Index           =   1
            Left            =   10650
            TabIndex        =   19
            Top             =   90
            Width           =   945
            _ExtentX        =   1667
            _ExtentY        =   688
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
            Height          =   390
            Index           =   2
            Left            =   8775
            TabIndex        =   20
            Top             =   90
            Width           =   1050
            _ExtentX        =   1852
            _ExtentY        =   688
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
            Height          =   390
            Index           =   3
            Left            =   7410
            TabIndex        =   21
            Top             =   90
            Width           =   885
            _ExtentX        =   1561
            _ExtentY        =   688
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
            Height          =   390
            Index           =   4
            Left            =   6240
            TabIndex        =   22
            Top             =   90
            Width           =   915
            _ExtentX        =   1614
            _ExtentY        =   688
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
            Height          =   390
            Index           =   5
            Left            =   5085
            TabIndex        =   23
            Top             =   90
            Width           =   975
            _ExtentX        =   1720
            _ExtentY        =   688
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
            Height          =   390
            Index           =   6
            Left            =   990
            TabIndex        =   24
            Top             =   90
            Width           =   900
            _ExtentX        =   1588
            _ExtentY        =   688
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
            Height          =   390
            Index           =   7
            Left            =   3960
            TabIndex        =   25
            Top             =   90
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
         Begin ImpulseButton.ISButton CmdHelp 
            Height          =   390
            Left            =   2025
            TabIndex        =   26
            Top             =   90
            Width           =   885
            _ExtentX        =   1561
            _ExtentY        =   688
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
         Begin ImpulseButton.ISButton ISButton2 
            Height          =   390
            Left            =   -240
            TabIndex        =   122
            Top             =   -120
            Visible         =   0   'False
            Width           =   885
            _ExtentX        =   1561
            _ExtentY        =   688
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
            ColorButton     =   14871017
            ColorHighlight  =   16777215
            ColorHoverText  =   16711680
            ColorShadow     =   4210752
            ColorOutline    =   0
            DrawFocusRectangle=   0   'False
            ColorToggledHoverText=   16711680
            ColorTextShadow =   4210752
         End
         Begin ImpulseButton.ISButton Accredit 
            Height          =   390
            Left            =   2880
            TabIndex        =   123
            Top             =   120
            Width           =   885
            _ExtentX        =   1561
            _ExtentY        =   688
            ButtonStyle     =   1
            ButtonPositionImage=   1
            Caption         =   "«⁄ „«œ"
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
         Height          =   435
         Index           =   3
         Left            =   15
         TabIndex        =   27
         TabStop         =   0   'False
         Top             =   8130
         Width           =   13560
         _cx             =   23918
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
         AutoSizeChildren=   0
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
            Left            =   13710
            Locked          =   -1  'True
            RightToLeft     =   -1  'True
            TabIndex        =   28
            TabStop         =   0   'False
            Top             =   30
            Width           =   1230
         End
         Begin MSDataListLib.DataCombo DCboUserName 
            Height          =   315
            Left            =   4050
            TabIndex        =   29
            Top             =   45
            Width           =   1455
            _ExtentX        =   2566
            _ExtentY        =   556
            _Version        =   393216
            Enabled         =   0   'False
            Text            =   ""
            RightToLeft     =   -1  'True
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "«Ã„«·Ì «·ﬂ„ÌÂ"
            Height          =   300
            Index           =   63
            Left            =   9120
            TabIndex        =   104
            Top             =   135
            Visible         =   0   'False
            Width           =   960
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
            Height          =   360
            Left            =   6480
            TabIndex        =   103
            Top             =   0
            Visible         =   0   'False
            Width           =   2580
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
            Height          =   405
            Left            =   10200
            RightToLeft     =   -1  'True
            TabIndex        =   102
            Top             =   0
            Width           =   2460
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "«·≈Ã„«·Ï"
            Height          =   285
            Index           =   25
            Left            =   12720
            RightToLeft     =   -1  'True
            TabIndex        =   101
            Top             =   60
            Width           =   630
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "≈Ã„«·Ì «·ÿ·»"
            Height          =   255
            Index           =   3
            Left            =   13950
            RightToLeft     =   -1  'True
            TabIndex        =   35
            Top             =   75
            Width           =   1875
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "«·”Ã· «·Õ«·Ì:"
            Height          =   240
            Index           =   0
            Left            =   2850
            RightToLeft     =   -1  'True
            TabIndex        =   34
            Top             =   120
            Width           =   1035
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "⁄œœ «·”Ã·« :"
            Height          =   240
            Index           =   2
            Left            =   1050
            RightToLeft     =   -1  'True
            TabIndex        =   33
            Top             =   120
            Width           =   930
         End
         Begin VB.Label XPTxtCurrent 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Height          =   270
            Left            =   2175
            RightToLeft     =   -1  'True
            TabIndex        =   32
            Top             =   105
            Width           =   690
         End
         Begin VB.Label XPTxtCount 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Height          =   240
            Left            =   90
            RightToLeft     =   -1  'True
            TabIndex        =   31
            Top             =   135
            Width           =   615
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Õ—— »Ê«”ÿ… : "
            Height          =   315
            Index           =   1
            Left            =   5370
            RightToLeft     =   -1  'True
            TabIndex        =   30
            Top             =   75
            Width           =   1020
         End
      End
      Begin C1SizerLibCtl.C1Elastic Ele 
         Height          =   4815
         Index           =   5
         Left            =   0
         TabIndex        =   36
         TabStop         =   0   'False
         Top             =   3300
         Width           =   13560
         _cx             =   23918
         _cy             =   8493
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
         Begin VB.Frame Frame4 
            BorderStyle     =   0  'None
            Height          =   855
            Left            =   7320
            RightToLeft     =   -1  'True
            TabIndex        =   124
            Top             =   3840
            Width           =   6255
            Begin DBPIXLib.DBPix20 DBPix202 
               Height          =   855
               Left            =   240
               TabIndex        =   125
               Top             =   0
               Width           =   2415
               _Version        =   131072
               _ExtentX        =   4260
               _ExtentY        =   1508
               _StockProps     =   1
               _Image          =   "FrmShowPrice.frx":038A
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
            Begin VB.Label Label4 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "Ì⁄ „œ"
               Height          =   255
               Left            =   5160
               RightToLeft     =   -1  'True
               TabIndex        =   128
               Top             =   240
               Width           =   735
            End
            Begin VB.Label Label10 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "«·‰ÊﬁÌ⁄"
               Height          =   255
               Left            =   2520
               RightToLeft     =   -1  'True
               TabIndex        =   127
               Top             =   240
               Width           =   735
            End
            Begin VB.Label LblPostedPerson 
               Alignment       =   2  'Center
               BackStyle       =   0  'Transparent
               Caption         =   "."
               Height          =   255
               Left            =   3600
               RightToLeft     =   -1  'True
               TabIndex        =   126
               Top             =   240
               Width           =   1695
            End
         End
         Begin C1SizerLibCtl.C1Elastic Ele 
            Height          =   690
            Index           =   2
            Left            =   30
            TabIndex        =   46
            TabStop         =   0   'False
            Top             =   270
            Width           =   13500
            _cx             =   23813
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
            Begin VB.TextBox TxtPrice 
               Alignment       =   1  'Right Justify
               Height          =   300
               Left            =   780
               MaxLength       =   10
               RightToLeft     =   -1  'True
               TabIndex        =   14
               Top             =   300
               Width           =   2025
            End
            Begin VB.TextBox TxtQuantity 
               Alignment       =   1  'Right Justify
               Enabled         =   0   'False
               Height          =   300
               Left            =   2820
               MaxLength       =   10
               RightToLeft     =   -1  'True
               TabIndex        =   13
               Top             =   300
               Width           =   2160
            End
            Begin VB.ComboBox CboItemCase 
               Height          =   315
               Left            =   5040
               RightToLeft     =   -1  'True
               Style           =   2  'Dropdown List
               TabIndex        =   12
               Top             =   300
               Width           =   1890
            End
            Begin MSDataListLib.DataCombo DCboItemsName 
               Height          =   315
               Left            =   6945
               TabIndex        =   11
               Top             =   300
               Width           =   3255
               _ExtentX        =   5741
               _ExtentY        =   556
               _Version        =   393216
               Text            =   ""
               RightToLeft     =   -1  'True
            End
            Begin MSDataListLib.DataCombo DCboItemsCode 
               Height          =   315
               Left            =   10260
               TabIndex        =   10
               Top             =   300
               Width           =   3195
               _ExtentX        =   5636
               _ExtentY        =   556
               _Version        =   393216
               Text            =   ""
               RightToLeft     =   -1  'True
            End
            Begin ImpulseButton.ISButton CmdAdd 
               Height          =   375
               Left            =   75
               TabIndex        =   15
               Top             =   270
               Width           =   630
               _ExtentX        =   1111
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
               ButtonImage     =   "FrmShowPrice.frx":03A2
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
               Caption         =   "«·”⁄—"
               Height          =   255
               Index           =   26
               Left            =   855
               RightToLeft     =   -1  'True
               TabIndex        =   51
               Top             =   0
               Width           =   1950
            End
            Begin VB.Label lbl 
               Alignment       =   2  'Center
               BackColor       =   &H00E2E9E9&
               Caption         =   "«·ﬂ„Ì…"
               Height          =   255
               Index           =   27
               Left            =   3060
               RightToLeft     =   -1  'True
               TabIndex        =   50
               Top             =   0
               Width           =   1890
            End
            Begin VB.Label lbl 
               Alignment       =   2  'Center
               BackColor       =   &H00E2E9E9&
               Caption         =   "Õ«·… «·’‰›"
               Height          =   255
               Index           =   29
               Left            =   5280
               RightToLeft     =   -1  'True
               TabIndex        =   49
               Top             =   0
               Width           =   1680
            End
            Begin VB.Label lbl 
               Alignment       =   2  'Center
               BackColor       =   &H00E2E9E9&
               Caption         =   "≈”„ «·’‰›"
               Height          =   255
               Index           =   30
               Left            =   7260
               RightToLeft     =   -1  'True
               TabIndex        =   48
               Top             =   0
               Width           =   3000
            End
            Begin VB.Label lbl 
               Alignment       =   2  'Center
               BackColor       =   &H00E2E9E9&
               Caption         =   "ﬂÊœ «·’‰›"
               Height          =   255
               Index           =   31
               Left            =   10440
               RightToLeft     =   -1  'True
               TabIndex        =   47
               Top             =   0
               Width           =   3015
            End
         End
         Begin VSFlex8UCtl.VSFlexGrid FG 
            Height          =   2790
            Left            =   30
            TabIndex        =   9
            Top             =   1095
            Width           =   13500
            _cx             =   23812
            _cy             =   4921
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
            FormatString    =   $"FrmShowPrice.frx":073C
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
            Left            =   495
            TabIndex        =   53
            Top             =   4140
            Width           =   6795
            _ExtentX        =   11986
            _ExtentY        =   1111
            ButtonWidth     =   609
            ButtonHeight    =   1005
            Appearance      =   1
            _Version        =   393216
         End
         Begin VB.Label LblItemsCount 
            Alignment       =   2  'Center
            BackColor       =   &H00404040&
            ForeColor       =   &H0000FFFF&
            Height          =   285
            Left            =   30
            RightToLeft     =   -1  'True
            TabIndex        =   54
            Top             =   4140
            Width           =   450
         End
      End
      Begin C1SizerLibCtl.C1Elastic Ele 
         Height          =   2715
         Index           =   0
         Left            =   15
         TabIndex        =   37
         TabStop         =   0   'False
         Top             =   645
         Width           =   13545
         _cx             =   23892
         _cy             =   4789
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
         Begin VB.ComboBox CboType 
            Height          =   315
            Left            =   3840
            RightToLeft     =   -1  'True
            Style           =   2  'Dropdown List
            TabIndex        =   130
            Top             =   120
            Width           =   1530
         End
         Begin VB.TextBox oldtxtNoteSerial1 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   0
            RightToLeft     =   -1  'True
            TabIndex        =   121
            Top             =   840
            Visible         =   0   'False
            Width           =   1215
         End
         Begin VB.TextBox TxtNoteSerial1 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   10560
            RightToLeft     =   -1  'True
            TabIndex        =   120
            Top             =   120
            Width           =   1695
         End
         Begin VB.TextBox Text6 
            Alignment       =   1  'Right Justify
            Height          =   570
            Left            =   7320
            MultiLine       =   -1  'True
            RightToLeft     =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   115
            Top             =   1920
            Width           =   4905
         End
         Begin VB.TextBox Text5 
            Alignment       =   1  'Right Justify
            Height          =   330
            Left            =   11400
            RightToLeft     =   -1  'True
            TabIndex        =   113
            Top             =   1200
            Width           =   825
         End
         Begin VB.TextBox TxtSearchCode 
            Alignment       =   1  'Right Justify
            Height          =   330
            Left            =   11400
            RightToLeft     =   -1  'True
            TabIndex        =   112
            Top             =   840
            Width           =   825
         End
         Begin VB.TextBox Txt_order_no 
            Alignment       =   1  'Right Justify
            Height          =   330
            Left            =   1800
            RightToLeft     =   -1  'True
            TabIndex        =   100
            Top             =   840
            Visible         =   0   'False
            Width           =   945
         End
         Begin VB.Frame Frame3 
            Caption         =   "»Ì«‰«  «·«⁄ „«œ"
            Height          =   615
            Left            =   0
            RightToLeft     =   -1  'True
            TabIndex        =   80
            Top             =   120
            Width           =   3855
            Begin VB.TextBox TxtLcNo 
               Alignment       =   1  'Right Justify
               Height          =   285
               Left            =   600
               RightToLeft     =   -1  'True
               TabIndex        =   81
               Top             =   240
               Width           =   2175
            End
            Begin MSComCtl2.DTPicker DTPicker2 
               Height          =   315
               Left            =   4080
               TabIndex        =   82
               Top             =   600
               Width           =   2100
               _ExtentX        =   3704
               _ExtentY        =   556
               _Version        =   393216
               Format          =   94830593
               CurrentDate     =   38784
            End
            Begin MSComCtl2.DTPicker DTPicker3 
               Height          =   315
               Left            =   4560
               TabIndex        =   83
               Top             =   960
               Width           =   1620
               _ExtentX        =   2858
               _ExtentY        =   556
               _Version        =   393216
               Format          =   94830593
               CurrentDate     =   38784
            End
            Begin MSComCtl2.DTPicker DTPicker4 
               Height          =   315
               Left            =   120
               TabIndex        =   84
               Top             =   960
               Width           =   2100
               _ExtentX        =   3704
               _ExtentY        =   556
               _Version        =   393216
               Format          =   94830593
               CurrentDate     =   38784
            End
            Begin MSComCtl2.DTPicker DTPicker5 
               Height          =   315
               Left            =   4560
               TabIndex        =   85
               Top             =   1320
               Width           =   1620
               _ExtentX        =   2858
               _ExtentY        =   556
               _Version        =   393216
               Format          =   94830593
               CurrentDate     =   38784
            End
            Begin MSComCtl2.DTPicker DTPicker6 
               Height          =   315
               Left            =   120
               TabIndex        =   86
               Top             =   1320
               Width           =   2100
               _ExtentX        =   3704
               _ExtentY        =   556
               _Version        =   393216
               Format          =   94830593
               CurrentDate     =   38784
            End
            Begin ImpulseButton.ISButton ISButton1 
               Height          =   285
               Left            =   120
               TabIndex        =   105
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
               RightToLeft     =   -1  'True
               TabIndex        =   93
               Top             =   1080
               Width           =   975
            End
            Begin VB.Label Label8 
               Alignment       =   1  'Right Justify
               Caption         =   " «—ÌŒ «·Ê’Ê· «·„ Êﬁ⁄"
               Height          =   255
               Left            =   2280
               RightToLeft     =   -1  'True
               TabIndex        =   92
               Top             =   1440
               Width           =   1575
            End
            Begin VB.Label Label7 
               Alignment       =   1  'Right Justify
               Caption         =   " «—ÌŒ «· √ŒÌ—"
               Height          =   255
               Left            =   6480
               RightToLeft     =   -1  'True
               TabIndex        =   91
               Top             =   1440
               Width           =   1215
            End
            Begin VB.Label Label6 
               Alignment       =   1  'Right Justify
               Caption         =   "«· «—ÌŒ «·›⁄·Ì"
               Height          =   375
               Left            =   2640
               RightToLeft     =   -1  'True
               TabIndex        =   90
               Top             =   1080
               Width           =   1215
            End
            Begin VB.Label Label5 
               Alignment       =   1  'Right Justify
               Caption         =   "«· «—ÌŒ «·„ Êﬁ⁄"
               Height          =   375
               Left            =   6480
               RightToLeft     =   -1  'True
               TabIndex        =   89
               Top             =   1080
               Width           =   1215
            End
            Begin VB.Label Label3 
               Alignment       =   1  'Right Justify
               Caption         =   "«· «—ÌŒ"
               Height          =   255
               Left            =   6360
               RightToLeft     =   -1  'True
               TabIndex        =   88
               Top             =   600
               Width           =   1335
            End
            Begin VB.Label Label1 
               Alignment       =   1  'Right Justify
               Caption         =   "—ﬁ„ «·«⁄ „«œ"
               Height          =   255
               Left            =   2640
               RightToLeft     =   -1  'True
               TabIndex        =   87
               Top             =   240
               Width           =   1095
            End
         End
         Begin VB.Frame Frame2 
            Height          =   1815
            Left            =   2040
            RightToLeft     =   -1  'True
            TabIndex        =   67
            Top             =   2520
            Visible         =   0   'False
            Width           =   5700
            Begin VB.TextBox Text7 
               Alignment       =   1  'Right Justify
               Height          =   285
               Left            =   240
               RightToLeft     =   -1  'True
               TabIndex        =   70
               Top             =   600
               Width           =   3855
            End
            Begin VB.TextBox Text3 
               Alignment       =   1  'Right Justify
               Height          =   285
               Left            =   2640
               RightToLeft     =   -1  'True
               TabIndex        =   69
               Top             =   1320
               Width           =   1455
            End
            Begin VB.TextBox Text2 
               Alignment       =   1  'Right Justify
               Height          =   285
               Left            =   240
               RightToLeft     =   -1  'True
               TabIndex        =   68
               Top             =   960
               Width           =   1335
            End
            Begin MSComCtl2.DTPicker DTPicker1 
               Height          =   315
               Left            =   240
               TabIndex        =   71
               Top             =   1320
               Width           =   1320
               _ExtentX        =   2328
               _ExtentY        =   556
               _Version        =   393216
               Format          =   94830593
               CurrentDate     =   38784
            End
            Begin MSDataListLib.DataCombo DataCombo9 
               Height          =   315
               Left            =   1920
               TabIndex        =   72
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
               TabIndex        =   73
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
               RightToLeft     =   -1  'True
               TabIndex        =   79
               Top             =   1320
               Width           =   975
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               Caption         =   "«·ﬁÌ„…"
               Height          =   285
               Index           =   23
               Left            =   1560
               RightToLeft     =   -1  'True
               TabIndex        =   78
               Top             =   960
               Width           =   975
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               Caption         =   "—ﬁ„ «·Õ”«»"
               Height          =   285
               Index           =   22
               Left            =   4320
               RightToLeft     =   -1  'True
               TabIndex        =   77
               Top             =   1320
               Width           =   1215
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               Caption         =   "«·⁄„·…"
               Height          =   285
               Index           =   21
               Left            =   4320
               RightToLeft     =   -1  'True
               TabIndex        =   76
               Top             =   960
               Width           =   1215
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               Caption         =   "«·»‰ﬂ"
               Height          =   285
               Index           =   20
               Left            =   4320
               RightToLeft     =   -1  'True
               TabIndex        =   75
               Top             =   600
               Width           =   1215
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               Caption         =   "‰Ê⁄ «·«„—"
               Height          =   285
               Index           =   19
               Left            =   4320
               RightToLeft     =   -1  'True
               TabIndex        =   74
               Top             =   240
               Width           =   1215
            End
         End
         Begin VB.Frame Frame1 
            Height          =   1695
            Left            =   120
            RightToLeft     =   -1  'True
            TabIndex        =   56
            Top             =   840
            Width           =   6615
            Begin VB.CheckBox chkshipped 
               Alignment       =   1  'Right Justify
               Caption         =   " „ «·‘Õ‰"
               Height          =   195
               Left            =   120
               RightToLeft     =   -1  'True
               TabIndex        =   116
               Top             =   1320
               Width           =   1815
            End
            Begin VB.TextBox Text1 
               Alignment       =   1  'Right Justify
               Height          =   285
               Left            =   120
               RightToLeft     =   -1  'True
               TabIndex        =   57
               Top             =   600
               Width           =   1935
            End
            Begin MSDataListLib.DataCombo DataCombo4 
               Height          =   315
               Left            =   3120
               TabIndex        =   58
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
               TabIndex        =   59
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
               TabIndex        =   60
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
               TabIndex        =   61
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
               TabIndex        =   108
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
               TabIndex        =   110
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
               RightToLeft     =   -1  'True
               TabIndex        =   111
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
               RightToLeft     =   -1  'True
               TabIndex        =   109
               Top             =   600
               Width           =   1095
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               Caption         =   "«·ﬁÌ„…"
               Height          =   285
               Index           =   17
               Left            =   2040
               RightToLeft     =   -1  'True
               TabIndex        =   66
               Top             =   600
               Width           =   975
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               Caption         =   "«· ’‰Ì›"
               Height          =   285
               Index           =   16
               Left            =   5400
               RightToLeft     =   -1  'True
               TabIndex        =   65
               Top             =   240
               Width           =   1095
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               Caption         =   "ÿ—Ìﬁ… «·œ›⁄"
               Height          =   285
               Index           =   15
               Left            =   2040
               RightToLeft     =   -1  'True
               TabIndex        =   64
               Top             =   960
               Width           =   975
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               Caption         =   "ÿ—Ìﬁ… «·‘Õ‰"
               Height          =   285
               Index           =   14
               Left            =   5280
               RightToLeft     =   -1  'True
               TabIndex        =   63
               Top             =   1320
               Width           =   1215
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               Caption         =   "«·»·œ"
               Height          =   285
               Index           =   13
               Left            =   5280
               RightToLeft     =   -1  'True
               TabIndex        =   62
               Top             =   960
               Width           =   1215
            End
         End
         Begin VB.ComboBox CboPriceType 
            Height          =   315
            Left            =   7290
            RightToLeft     =   -1  'True
            Style           =   2  'Dropdown List
            TabIndex        =   4
            Top             =   120
            Width           =   2250
         End
         Begin VB.TextBox TxtTransSerial 
            Alignment       =   1  'Right Justify
            Height          =   330
            Left            =   10560
            RightToLeft     =   -1  'True
            TabIndex        =   0
            Top             =   -240
            Visible         =   0   'False
            Width           =   1665
         End
         Begin VB.TextBox TxtModFlg 
            Alignment       =   1  'Right Justify
            Height          =   330
            Left            =   2880
            RightToLeft     =   -1  'True
            TabIndex        =   39
            Top             =   -210
            Visible         =   0   'False
            Width           =   705
         End
         Begin VB.TextBox TxtFillData 
            Alignment       =   1  'Right Justify
            Height          =   270
            Left            =   1965
            RightToLeft     =   -1  'True
            TabIndex        =   38
            Top             =   -150
            Visible         =   0   'False
            Width           =   870
         End
         Begin VB.TextBox XPTxtBillID 
            Alignment       =   1  'Right Justify
            Height          =   300
            Left            =   30
            Locked          =   -1  'True
            RightToLeft     =   -1  'True
            TabIndex        =   8
            Top             =   -150
            Visible         =   0   'False
            Width           =   1920
         End
         Begin MSDataListLib.DataCombo DBCboClientName 
            Height          =   315
            Left            =   7305
            TabIndex        =   2
            Top             =   840
            Width           =   4110
            _ExtentX        =   7250
            _ExtentY        =   556
            _Version        =   393216
            ListField       =   "6"
            Text            =   ""
            RightToLeft     =   -1  'True
         End
         Begin MSDataListLib.DataCombo DCboStoreName 
            Height          =   315
            Left            =   7305
            TabIndex        =   3
            Top             =   1230
            Width           =   4110
            _ExtentX        =   7250
            _ExtentY        =   556
            _Version        =   393216
            ListField       =   "7"
            Text            =   ""
            RightToLeft     =   -1  'True
         End
         Begin MSComCtl2.DTPicker XPDtbBill 
            Height          =   315
            Left            =   10560
            TabIndex        =   1
            Top             =   480
            Width           =   1695
            _ExtentX        =   2990
            _ExtentY        =   556
            _Version        =   393216
            Format          =   94830593
            CurrentDate     =   38784
         End
         Begin ImpulseButton.ISButton XPBtnNewClients 
            Height          =   450
            Left            =   6255
            TabIndex        =   40
            TabStop         =   0   'False
            Top             =   1830
            Width           =   60
            _ExtentX        =   106
            _ExtentY        =   794
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
            ButtonImage     =   "FrmShowPrice.frx":0928
            ColorButton     =   14871017
            ColorHighlight  =   16777215
            ColorShadow     =   -2147483631
            ColorOutline    =   -2147483631
            DrawFocusRectangle=   0   'False
         End
         Begin ImpulseButton.ISButton CmdTemplate 
            Height          =   480
            Left            =   585
            TabIndex        =   41
            Top             =   1515
            Visible         =   0   'False
            Width           =   1740
            _ExtentX        =   3069
            _ExtentY        =   847
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
            Height          =   630
            Index           =   4
            Left            =   12960
            TabIndex        =   5
            TabStop         =   0   'False
            Top             =   1920
            Width           =   3795
            _cx             =   6694
            _cy             =   1111
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
               RightToLeft     =   -1  'True
               TabIndex        =   6
               Top             =   210
               Width           =   1815
            End
            Begin VB.TextBox XPTxtTaxValue 
               Alignment       =   1  'Right Justify
               Height          =   390
               Left            =   30
               RightToLeft     =   -1  'True
               TabIndex        =   7
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
               RightToLeft     =   -1  'True
               TabIndex        =   52
               Top             =   285
               Width           =   720
            End
         End
         Begin ImpulseButton.ISButton CmdConvert 
            Height          =   285
            Left            =   10080
            TabIndex        =   99
            Top             =   2400
            Visible         =   0   'False
            Width           =   2115
            _ExtentX        =   3731
            _ExtentY        =   503
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
         Begin MSDataListLib.DataCombo Dccurrency 
            Height          =   315
            Left            =   7320
            TabIndex        =   106
            Top             =   1560
            Width           =   4905
            _ExtentX        =   8652
            _ExtentY        =   556
            _Version        =   393216
            ListField       =   "6"
            Text            =   ""
            RightToLeft     =   -1  'True
         End
         Begin MSDataListLib.DataCombo dcBranch 
            Height          =   315
            Left            =   7320
            TabIndex        =   117
            Top             =   480
            Width           =   2280
            _ExtentX        =   4022
            _ExtentY        =   556
            _Version        =   393216
            BackColor       =   16777215
            Text            =   ""
            RightToLeft     =   -1  'True
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "”Ì«”… «·ÿ·»Ì…"
            Height          =   240
            Index           =   18
            Left            =   5640
            RightToLeft     =   -1  'True
            TabIndex        =   129
            Top             =   120
            Width           =   1095
         End
         Begin VB.Label Label2 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "«·›—⁄"
            ForeColor       =   &H00000000&
            Height          =   375
            Left            =   9585
            TabIndex        =   118
            Top             =   480
            Width           =   735
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "„·«ÕŸ« "
            Height          =   270
            Index           =   28
            Left            =   12360
            RightToLeft     =   -1  'True
            TabIndex        =   114
            Top             =   2040
            Width           =   945
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "«·⁄„·Â"
            Height          =   285
            Index           =   12
            Left            =   12360
            RightToLeft     =   -1  'True
            TabIndex        =   107
            Top             =   1680
            Width           =   855
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "‰Ê⁄ «·«„—"
            Height          =   240
            Index           =   9
            Left            =   9540
            RightToLeft     =   -1  'True
            TabIndex        =   55
            Top             =   120
            Width           =   855
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "—ﬁ„ «·«„—"
            Height          =   270
            Index           =   5
            Left            =   12210
            RightToLeft     =   -1  'True
            TabIndex        =   45
            Top             =   120
            Width           =   1170
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   " «—ÌŒ «·«„—"
            Height          =   195
            Index           =   6
            Left            =   12600
            RightToLeft     =   -1  'True
            TabIndex        =   44
            Top             =   480
            Width           =   810
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "«·⁄„Ì· / «·„Ê—œ"
            Height          =   240
            Index           =   7
            Left            =   12315
            RightToLeft     =   -1  'True
            TabIndex        =   43
            Top             =   840
            Width           =   1065
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "«”„ «·„Œ“‰"
            Height          =   270
            Index           =   8
            Left            =   12435
            RightToLeft     =   -1  'True
            TabIndex        =   42
            Top             =   1200
            Width           =   945
         End
      End
      Begin C1SizerLibCtl.C1Elastic Ele 
         Height          =   615
         Index           =   6
         Left            =   15
         TabIndex        =   94
         TabStop         =   0   'False
         Top             =   15
         Width           =   13500
         _cx             =   23813
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
         Caption         =   "«Ê«„— «·»Ì⁄ Ê «·‘—«¡ Ê«·›Ê« Ì— «·„»œ∆Ì…"
         Align           =   0
         AutoSizeChildren=   0
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
         Begin ImpulseButton.ISButton XPBtnMove 
            Height          =   345
            Index           =   0
            Left            =   1860
            TabIndex        =   95
            Top             =   105
            Width           =   750
            _ExtentX        =   1323
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
            ButtonImage     =   "FrmShowPrice.frx":0CC2
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
            Left            =   1005
            TabIndex        =   96
            Top             =   105
            Width           =   735
            _ExtentX        =   1296
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
            ButtonImage     =   "FrmShowPrice.frx":105C
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
            Left            =   2670
            TabIndex        =   97
            Top             =   105
            Width           =   735
            _ExtentX        =   1296
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
            ButtonImage     =   "FrmShowPrice.frx":13F6
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
            Left            =   165
            TabIndex        =   98
            Top             =   105
            Width           =   750
            _ExtentX        =   1323
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
            ButtonImage     =   "FrmShowPrice.frx":1790
            ColorHighlight  =   4194304
            ColorHoverText  =   16777215
            ColorShadow     =   -2147483631
            ColorOutline    =   -2147483631
            DrawFocusRectangle=   0   'False
            DisabledImageStyle=   1
            ColorToggledHoverText=   16777215
            ColorTextShadow =   16777215
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
            Left            =   3600
            RightToLeft     =   -1  'True
            TabIndex        =   119
            Top             =   0
            Width           =   3915
         End
      End
   End
End
Attribute VB_Name = "FrmShowPriceOriginal"
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

Private Sub Cmd_Click(Index As Integer)
    Dim intDef As Integer
    On Error GoTo ErrTrap

    Select Case Index

        Case 0

            If DoPremis(Do_New, Me.Name, True) = False Then
                Exit Sub
            End If

            clear_all Me
            TxtModFlg.Text = "N"
            XPTxtBillID.Text = CStr(new_id("Transactions", "Transaction_ID", "", True))
            TxtTransSerial.Text = CStr(new_id("Transactions", "Transaction_Serial", "", True, "Transaction_Type=6"))
            NewGrid.GridDefaultValue 1
            Me.DCboUserName.BoundText = user_id
            intDef = GetSetting(StrAppRegPath, "DefaultOptions", "DefaultClient", 2)
            DBCboClientName.BoundText = intDef
            intDef = GetSetting(StrAppRegPath, "DefaultOptions", "DefaultSaleStore", 1)
            DCboStoreName.BoundText = intDef
            DcCurrency.BoundText = 1
            FG.SetFocus
            FG.Col = FG.ColIndex("Code")
            FG.Row = FG.Rows - 1
            Me.CBOPriceType.ListIndex = GeneralPriceType
            Me.dcBranch.BoundText = branch_id
            DBPix202.ImageClear

        Case 1

            If DoPremis(Do_Edit, Me.Name, True) = False Then
                Exit Sub
            End If

            TxtModFlg.Text = "E"
            CuurentLogdata
            Me.DCboUserName.BoundText = user_id

        Case 2
            Dim Msg  As String

            If Trim(dcBranch.BoundText) = "" Then
                If SystemOptions.UserInterface = EnglishInterface Then
                    Msg = "Specify Departement"
                Else
                    Msg = "Õœœ «·›—⁄ «Ê·« "
                End If
              
                MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
                dcBranch.SetFocus
                SendKeys "{F4}"
                Screen.MousePointer = vbDefault
                Exit Sub
            End If

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

            FrmBuySearch.DealingForm = GridTransType.ShowPrice
            FrmBuySearch.Caption = "«·»ÕÀ ⁄‰ ⁄—÷ ”⁄—"
            FrmBuySearch.show vbModal

        Case 7

            If DoPremis(Do_Print, Me.Name, True) = False Then
                Exit Sub
            End If

            PrintReport

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
    xReport.ApplicationName = App.title
    xReport.ReportAuthor = App.title
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

    If Me.CBOPriceType.ListIndex = 0 Then
        Set Frm = New frmsalebill
    ElseIf Me.CBOPriceType.ListIndex = 1 Then
        Set Frm = New FrmBillBuy
    End If

    With Frm
        .Convert
        '    .XPTxtBillID.Text = XPTxtBillID.Text
        .XPDtbBill.Value = XPDtbBill.Value
        .DBCboClientName.BoundText = DBCboClientName.BoundText
        .DCboStoreName.BoundText = DCboStoreName.BoundText
        .DcCurrency.BoundText = Me.DcCurrency.BoundText

        For RowNum = 1 To FG.Rows - 1

            If .FG.TextMatrix(.FG.Rows - 1, .FG.ColIndex("Code")) <> "" Then
                .FG.Rows = .FG.Rows + 1
            End If

            .FG.TextMatrix(.FG.Rows - 1, .FG.ColIndex("Name")) = IIf(FG.TextMatrix(RowNum, FG.ColIndex("Name")) = "", "", FG.TextMatrix(RowNum, FG.ColIndex("Name")))
            .FG.TextMatrix(.FG.Rows - 1, .FG.ColIndex("Code")) = IIf(FG.TextMatrix(RowNum, FG.ColIndex("Code")) = "", "", FG.TextMatrix(RowNum, FG.ColIndex("Code")))
            .FG.TextMatrix(.FG.Rows - 1, .FG.ColIndex("ItemCase")) = IIf(FG.TextMatrix(RowNum, FG.ColIndex("ItemCase")) = "", "", FG.TextMatrix(RowNum, FG.ColIndex("ItemCase")))
            .FG.TextMatrix(.FG.Rows - 1, .FG.ColIndex("HaveSerial")) = IIf(FG.TextMatrix(RowNum, FG.ColIndex("HaveSerial")) = "", "", FG.TextMatrix(RowNum, FG.ColIndex("HaveSerial")))
            .FG.TextMatrix(.FG.Rows - 1, .FG.ColIndex("Count")) = IIf(FG.TextMatrix(RowNum, FG.ColIndex("Count")) = "", "", FG.TextMatrix(RowNum, FG.ColIndex("Count")))
            .FG.TextMatrix(.FG.Rows - 1, .FG.ColIndex("Price")) = IIf(FG.TextMatrix(RowNum, FG.ColIndex("Price")) = "", "", FG.TextMatrix(RowNum, FG.ColIndex("Price")))
            .FG.TextMatrix(.FG.Rows - 1, .FG.ColIndex("DiscountType")) = IIf(FG.TextMatrix(RowNum, FG.ColIndex("DiscountType")) = "", "", FG.TextMatrix(RowNum, FG.ColIndex("DiscountType")))
            Dim StrSQL As String
            Dim RsUnit As New ADODB.Recordset
        
            StrSQL = "SELECT dbo.Transactions.Transaction_Type, dbo.Transaction_Details.UnitId, dbo.TblUnites.UnitName, dbo.Transactions.Transaction_Serial FROM dbo.Transaction_Details INNER JOIN dbo.Transactions ON dbo.Transaction_Details.Transaction_ID = dbo.Transactions.Transaction_ID INNER JOIN dbo.TblUnites ON dbo.Transaction_Details.UnitId = dbo.TblUnites.UnitID WHERE (dbo.Transactions.Transaction_Type = 6) AND (dbo.Transactions.Transaction_Serial = '" & TxtTransSerial & "')"
            Set RsUnit = New ADODB.Recordset
            RsUnit.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText
        
            .FG.Cell(flexcpData, .FG.Rows - 1, FG.ColIndex("UnitID")) = IIf(IsNull(RsUnit("UnitID")), "", (RsUnit("UnitID").Value))
            .FG.TextMatrix(.FG.Rows - 1, FG.ColIndex("UnitID")) = IIf(IsNull(RsUnit("UnitName")), "", (RsUnit("UnitName").Value))
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
    TxtSearchCode.Text = ""

    Dim DefaultSalesPersonId As Integer
    Dim Fullcode As String

    GetCustomersDetail val(DBCboClientName.BoundText), DefaultSalesPersonId, Fullcode

    TxtSearchCode.Text = Fullcode

End Sub

Private Sub DBCboClientName_KeyUp(KeyCode As Integer, _
                                  Shift As Integer)

    If KeyCode = vbKeyF3 Then
        FrmCustemerSearch.searchtype = 7
        FrmCustemerSearch.show vbModal
    End If

End Sub
 
Private Sub Ele_Click(Index As Integer)

    Select Case Index

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

Private Sub FG_AfterEdit(ByVal Row As Long, _
                         ByVal Col As Long)

    If Me.TxtModFlg <> "E" Then Exit Sub

    '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\LogFile Saving//////////////////////////////////////////////////////////////////////////
    If Col = FG.ColIndex("Code") Or Col = FG.ColIndex("Name") Then
        RegisterItemData Me.Name, Me.TxtModFlg, FG.Cell(flexcpTextDisplay, Row, FG.ColIndex("Code")), FG.Cell(flexcpTextDisplay, Row, FG.ColIndex("Name")), , , , , , , , , , , Me.TXT_order_no
    ElseIf Col = FG.ColIndex("UnitID") Then
        RegisterItemData Me.Name, Me.TxtModFlg, FG.Cell(flexcpTextDisplay, Row, FG.ColIndex("Code")), FG.Cell(flexcpTextDisplay, Row, FG.ColIndex("Name")), FG.Cell(flexcpTextDisplay, Row, FG.ColIndex("UnitID")), , , , , , , , , , Me.TXT_order_no
    ElseIf Col = FG.ColIndex("Count") Then
        RegisterItemData Me.Name, Me.TxtModFlg, FG.Cell(flexcpTextDisplay, Row, FG.ColIndex("Code")), FG.Cell(flexcpTextDisplay, Row, FG.ColIndex("Name")), , (FG.TextMatrix(Row, FG.ColIndex("Count"))), , , , , , , , , Me.TXT_order_no
    ElseIf Col = FG.ColIndex("Price") Then
        RegisterItemData Me.Name, Me.TxtModFlg, FG.Cell(flexcpTextDisplay, Row, FG.ColIndex("Code")), FG.Cell(flexcpTextDisplay, Row, FG.ColIndex("Name")), , , (FG.TextMatrix(Row, FG.ColIndex("Price"))), , , , , , , , Me.TXT_order_no
    ElseIf Col = FG.ColIndex("ColorID") Then
        RegisterItemData Me.Name, Me.TxtModFlg, FG.Cell(flexcpTextDisplay, Row, FG.ColIndex("Code")), FG.Cell(flexcpTextDisplay, Row, FG.ColIndex("Name")), , , , , FG.Cell(flexcpTextDisplay, Row, FG.ColIndex("ColorID")), , , , , , Me.TXT_order_no
    ElseIf Col = FG.ColIndex("ItemSize") Then
        RegisterItemData Me.Name, Me.TxtModFlg, FG.Cell(flexcpTextDisplay, Row, FG.ColIndex("Code")), FG.Cell(flexcpTextDisplay, Row, FG.ColIndex("Name")), , , , , , FG.Cell(flexcpTextDisplay, Row, FG.ColIndex("ItemSize")), , , , , Me.TXT_order_no
    ElseIf Col = FG.ColIndex("ClassId") Then
        RegisterItemData Me.Name, Me.TxtModFlg, FG.Cell(flexcpTextDisplay, Row, FG.ColIndex("Code")), FG.Cell(flexcpTextDisplay, Row, FG.ColIndex("Name")), , , , , , , FG.Cell(flexcpTextDisplay, Row, FG.ColIndex("ClassId")), , , , Me.TXT_order_no
    ElseIf Col = FG.ColIndex("DiscountType") Then
        RegisterItemData Me.Name, Me.TxtModFlg, FG.Cell(flexcpTextDisplay, Row, FG.ColIndex("Code")), FG.Cell(flexcpTextDisplay, Row, FG.ColIndex("Name")), , , , , , , , FG.Cell(flexcpTextDisplay, Row, FG.ColIndex("DiscountType")), , , Me.TXT_order_no
    ElseIf Col = FG.ColIndex("DiscountVal") Then
        RegisterItemData Me.Name, Me.TxtModFlg, FG.Cell(flexcpTextDisplay, Row, FG.ColIndex("Code")), FG.Cell(flexcpTextDisplay, Row, FG.ColIndex("Name")), , , , , , , , , FG.TextMatrix(Row, FG.ColIndex("DiscountVal")), , Me.TXT_order_no

    End If

    '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\LogFile Saving//////////////////////////////////////////////////////////////////////////

End Sub

Private Sub Fg_CellButtonClick(ByVal Row As Long, _
                               ByVal Col As Long)

    If Me.TxtModFlg.Text = "N" Or Me.TxtModFlg.Text = "E" Then
        '    FrmAddNewItem.Tag = "xx"
'        FrmAddNewItem.DealingForm = ShowPrice
'        FrmAddNewItem.show vbModal
    End If

End Sub

Private Sub Form_Activate()
    'XPTxtBillID.SetFocus
End Sub

Private Sub ISButton1_Click()
    FrmLC.show
    FrmLC.Retrive Trim(Me.TxtLcNo.Text)
    'Frame3.Visible = True
End Sub

Private Sub Label10_Click()
    Frame3.Visible = False
End Sub
 
Private Sub Accredit_Click()
    Dim Sql As String
    Dim BeginTrans As Boolean
    'sql = "update  Transactions  set Posted=" & user_id & "  where Transaction_ID=" & Val(XPTxtBillID.text)
    'Cn.Execute sql

    Cn.BeginTrans
    BeginTrans = True

    If IsNull(rs("Posted")) Then
        rs("Posted") = user_id
        rs("PostedDate") = Date
    Else
        rs("Posted") = Null
        rs("PostedDate") = Date
    End If
   
    rs.update

    Cn.CommitTrans
    BeginTrans = False

    Retrive (val(XPTxtBillID.Text))

End Sub

Private Sub TxtSearchCode_KeyPress(KeyAscii As Integer)
    Dim CUSTID As Integer

    If KeyAscii = vbKeyReturn Then
        GetCustomersDetail CUSTID, , TxtSearchCode.Text, 1
        DBCboClientName.BoundText = CUSTID
    End If

End Sub

Private Sub TxtFillData_Change()

    If TxtFillData.Text = "F" Then
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

Private Sub XPBtnMove_Click(Index As Integer)
    On Error GoTo ErrTrap

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

Private Sub Form_KeyDown(KeyCode As Integer, _
                         Shift As Integer)
    On Error GoTo ErrTrap

    If KeyCode = vbKeyReturn Then
        If Me.TxtModFlg.Text = "R" Then
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
        If TxtModFlg.Text = "N" Or TxtModFlg.Text = "E" Then
        
        End If
    End If

    If KeyCode = vbKeyF5 Then
        If TxtModFlg.Text = "N" Or TxtModFlg.Text = "E" Then
            XPBtnNewClients_Click
        End If
    End If

    If KeyCode = vbKeyF3 Then
        If TxtModFlg.Text = "N" Or TxtModFlg.Text = "E" Then
       
        End If
    End If

    If Shift = 2 Then
        If KeyCode = vbKeySpace Then
            If TxtModFlg.Text = "N" Or TxtModFlg.Text = "E" Then
            
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

Private Sub Form_Load()
    Dim RsClients As New ADODB.Recordset
    Dim StrSQL As String
    Dim Num As Integer
    Dim StrList As String
    Dim BGround As New ClsBackGroundPic
    Dim RsNote As New ADODB.Recordset
    Dim ShowTax As Boolean
    Dim Dcombos As ClsDataCombos

    On Error GoTo ErrTrap
    ScreenNameArabic = "«Ê«„— «·»Ì⁄ Ê «·‘—«¡ Ê«·›Ê« Ì— «·„»œ∆Ì…"
    ScreenNameEnglish = "Sales Order / Purcahse Order / Performa Invoice"
    RegisterLogInOut Me.Name, ScreenNameArabic, ScreenNameEnglish, "1"

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
    Set NewGrid.Grid = FG
    NewGrid.GridTrans = GridTransType.ShowPrice
    Set NewGrid.TxtModFlag = TxtModFlg
    Set NewGrid.TxtTotal = XPTxtSum
    Set NewGrid.TxtFillData = TxtFillData
    Set NewGrid.TxtTaxValue = Me.XPTxtTaxValue
    Set NewGrid.GrdTBar = Me.TBar
    Set NewGrid.LblItemsCount = Me.LblItemsCount
    'Set NewGrid.LblItemsCount = Me.LblItemsCount
    Set NewGrid.LblTotalAll = Me.LblTotalAll
    Set NewGrid.LblTotalQty = Me.LblTotalQty

    ' ⁄»∆… »Ì«‰«  «·√’‰«›
    Set NewGrid.DCboItemName = DCboItemsName
    Set NewGrid.DCboItemCode = DCboItemsCode
    Set NewGrid.CboItemCase = CboItemCase
    Set NewGrid.CmdAddData = CmdAdd
    'Set NewGrid.TxtSerial = TxtSerial
    Set NewGrid.TxtQuantity = TxtQuantity
    Set NewGrid.TxtPrice = TxtPrice
    ' Resize_Form Me, TransactionSize
    Me.left = (mdifrmmain.Width - Me.Width) / 2
    Me.top = (mdifrmmain.Height - Me.Height) / 2 - 500
    FG.WallPaper = BGround.Picture
    AddTip
    XPDtbBill.Value = Date
    Set Dcombos = New ClsDataCombos
    Dcombos.GetCustomersSuppliers 0, Me.DBCboClientName, True
    Dcombos.GetStores Me.DCboStoreName
    Dcombos.GetUsers Me.DCboUserName
    Dcombos.GetBranches Me.dcBranch

    Set cSearchDcbo(0) = New clsDCboSearch
    Set cSearchDcbo(0).Client = Me.DBCboClientName

    Set cSearchDcbo(1) = New clsDCboSearch
    Set cSearchDcbo(1).Client = Me.DCboStoreName
    NewGrid.FillGrid

    With Me.CBOPriceType
        .Clear

        If SystemOptions.UserInterface = ArabicInterface Then
            .AddItem " «„— »Ì⁄ "
            .AddItem "«„—‘—«¡/ ÿ·»Ì…  "
            .AddItem "  ›« Ê—… „»œ∆ÌÂ"
            .AddItem "«·ÿ·»«    «·œ«Œ·Ì… "
        Else
            .AddItem "Sales  Order"
            .AddItem "Purchases   Order"
            .AddItem "PerForma   Invoices"
            .AddItem "Internal Order "
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

    StrSQL = "SELECT * FROM Transactions WHERE (Transaction_Type=6 or Transaction_Type=29  or Transaction_Type=17)" 'OR Transaction_Type=17
    StrSQL = StrSQL + " Order By Transaction_ID"
    Set rs = New ADODB.Recordset
    rs.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText
    Dim My_SQL As String
    My_SQL = " select id,code from currency"
 
    fill_combo Me.DcCurrency, My_SQL
    fill_combo Me.DataCombo11, My_SQL

    My_SQL = " select code,account_name from markaas_taklefa"
 
    fill_combo Me.DataCombo1, My_SQL

    My_SQL = " select id,Project_name from projects"
 
    fill_combo Me.DataCombo2, My_SQL

    My_SQL = " select CountryID,CountryName from TblCountriesData"
 
    fill_combo Me.DataCombo4, My_SQL

    My_SQL = " select id,name from Shipment_mode"
 
    fill_combo Me.DataCombo5, My_SQL

    XPBtnMove_Click 2
    Me.TxtModFlg.Text = "R"

    If OPEN_NEW_SCREEN = True Then
        Cmd_Click (0)
    End If

    Exit Sub
ErrTrap:
End Sub

Private Sub Form_Unload(Cancel As Integer)
    On Error GoTo ErrTrap
    Dim I As Integer
    RegisterLogInOut Me.Name, ScreenNameArabic, ScreenNameEnglish

    If rs.State = adStateOpen Then
        If Not (rs.EOF Or rs.BOF) Then
            If rs.EditMode <> adEditNone Then
                rs.CancelUpdate
            End If
        End If

        rs.Close
    End If

    For I = LBound(cSearchDcbo) To UBound(cSearchDcbo)
        Set cSearchDcbo(I) = Nothing
    Next I

    Set rs = Nothing
    Set TTP = Nothing
    NewGrid.Class_Terminate
    Set NewGrid = Nothing
    Set SaleReport = Nothing
    Exit Sub
ErrTrap:
End Sub

Function CuurentLogdata(Optional Currentmode As String)
    LogTextA = "    ‘«‘… " & ScreenNameArabic & Chr(13) & " —ﬁ„ «·”‰œ   " & TXT_order_no.Text & Chr(13) & " «· «—ÌŒ " & XPDtbBill.Value & Chr(13) & "«‰Ê⁄ «·”‰œ  " & CBOPriceType.Text & Chr(13) & " «·„Œ“‰  " & DCboStoreName.Text & Chr(13) & "  «·⁄„Ì· / «·„Ê—œ   " & DBCboClientName.Text & Chr(13) & " —ﬁ„ «·«⁄ „«œ    " & TxtLcNo
                     
    LogTextE = "    Screen  " & ScreenNameEnglish & Chr(13) & "Vchr . No   " & TXT_order_no.Text & Chr(13) & " Date " & XPDtbBill.Value & Chr(13) & " Type  " & CBOPriceType.Text & Chr(13) & " Store  " & DCboStoreName.Text & Chr(13) & " Customer/ Supplier " & DBCboClientName.Text & Chr(13) & " Lc NO    " & TxtLcNo
                     
    If Currentmode <> "D" Then
        AddToLogFile CInt(user_id), 0, Date, Time, LogTextA, LogTextE, Me.Name, Me.TxtModFlg, "", , , Me.TXT_order_no
    Else
        AddToLogFile CInt(user_id), 0, Date, Time, LogTextA, LogTextE, Me.Name, "D", "", , , Me.TXT_order_no
    End If
    
End Function

Private Sub TxtModFlg_Change()
    On Error GoTo ErrTrap

    Select Case Me.TxtModFlg.Text

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
            Accredit.Enabled = True
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
            FG.Rows = 2
            Me.XPDtbBill.Enabled = True
            XPDtbBill.Value = Date
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
        rs.find "Transaction_ID=" & Lngid, , adSearchForward, adBookmarkFirst

        If rs.BOF Or rs.EOF Then
            Exit Sub
        End If
    End If

    TxtFillData.Text = "T"
    Screen.MousePointer = vbArrowHourglass
    XPTxtBillID.Text = IIf(IsNull(rs("Transaction_ID").Value), "", val(rs("Transaction_ID").Value))

    TXT_order_no.Text = IIf(IsNull(rs("order_no").Value), "", rs("order_no").Value)

    If rs("shipped").Value = True Then
        chkshipped.Value = vbChecked
    Else
        chkshipped.Value = Unchecked
    End If

    Me.DataCombo4.BoundText = IIf(IsNull(rs("countryid").Value), "", rs("countryid").Value)

    TxtTransSerial.Text = IIf(IsNull(rs("Transaction_Serial").Value), "", (rs("Transaction_Serial").Value))
    XPDtbBill.Value = IIf(IsNull(rs("Transaction_Date").Value), "", (rs("Transaction_Date").Value))
    Me.DBCboClientName.BoundText = IIf(IsNull(rs("CusID").Value), "", rs("CusID").Value)
    DcCurrency.BoundText = IIf(IsNull(rs("Currency_id").Value), "", rs("Currency_id").Value)
    'If rs("Transaction_Type").value = 6 Then
    '    Me.CboPriceType.ListIndex = 1
    'ElseIf rs("Transaction_Type").value = 17 Then '17
    '    Me.CboPriceType.ListIndex = 0
    'ElseIf rs("Transaction_Type").value = 29 Then
    'Me.CboPriceType.ListIndex = 2
    'End If

    If rs("Transaction_Type").Value = 6 Then
        Me.CBOPriceType.ListIndex = 0
    ElseIf rs("Transaction_Type").Value = 17 Then '17
        Me.CBOPriceType.ListIndex = 1
    ElseIf rs("Transaction_Type").Value = 29 Then
        Me.CBOPriceType.ListIndex = 2
    End If

    Me.DCboUserName.BoundText = IIf(IsNull(rs("UserID").Value), "", rs("UserID").Value)
    Me.DCboStoreName.BoundText = IIf(IsNull(rs("StoreID").Value), "", rs("StoreID").Value)

    XPTxtTaxValue.Text = IIf(IsNull(rs("TaxValue").Value), "", (rs("TaxValue").Value))
    TxtLcNo.Text = IIf(IsNull(rs("LcNo").Value), "", (rs("LcNo").Value))
    XPChkTAX.Value = IIf(rs("TaxFound") = True, Checked, Unchecked)
    dcBranch.BoundText = IIf(IsNull(rs("BranchId").Value), "", rs("BranchId").Value)

    Me.TxtNoteSerial1.Text = IIf(IsNull(rs("NoteSerial1").Value), "", (rs("NoteSerial1").Value))
    Me.oldtxtNoteSerial1.Text = IIf(IsNull(rs("OldNoteSerial1").Value), IIf(IsNull(rs("NoteSerial1").Value), "", rs("NoteSerial1").Value), rs("OldNoteSerial1").Value)

    If TXT_order_no <> "" Then
        Me.TxtNoteSerial1.Text = TXT_order_no
    End If

    'Txt_order_no

    lbl(64).Caption = showLabel(TxtNoteSerial1, oldtxtNoteSerial1)

    DBPix202.ImageClear

    If Dir(App.path & "\images\sign" & rs("posted").Value & ".JPG") <> "" Then
        
        DBPix202.ImageLoadFile (App.path & "\images\sign" & user_id & ".JPG")
    End If

    If Not IsNull(rs("posted").Value) Then
        Frame4.Visible = True
        GetUserData val(rs("posted").Value), , , , , , , Dusername
        LblPostedPerson = Dusername

        If user_id = rs("posted").Value Then
            If CheckOrderNotInTransaction(21, TxtNoteSerial1) = False Then
                Accredit.Caption = "«·€«¡ «·«⁄ „«œ "
            Else
                Accredit.Caption = "   «⁄ „«œ "
            End If
         
        End If

    Else
        Frame4.Visible = False
        Accredit.Caption = "   «⁄ „«œ "
    End If
  
    FG.Clear flexClearScrollable, flexClearEverything
    FG.Rows = 2
    FG.Clear flexClearScrollable, flexClearEverything
    FG.Refresh
    StrSQL = "SELECT TblItems.HaveSerial, * FROM TblItems INNER JOIN Transaction_Details " & "ON TblItems.ItemID = Transaction_Details.Item_ID INNER JOIN dbo.TblUnites ON dbo.Transaction_Details.UnitID = dbo.TblUnites.UnitID"
    StrSQL = StrSQL + " where Transaction_ID=" & val(rs("Transaction_ID").Value)

    RsDetails.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText
    XPTxtSum.Text = ""

    If Not (RsDetails.EOF Or RsDetails.BOF) Then
        FG.Rows = RsDetails.RecordCount + 1

        For Num = 1 To RsDetails.RecordCount
            FG.TextMatrix(Num, FG.ColIndex("Code")) = IIf(IsNull(RsDetails("Item_ID")), "", (RsDetails("Item_ID").Value))
            FG.TextMatrix(Num, FG.ColIndex("Name")) = IIf(IsNull(RsDetails("Item_ID")), "", Trim(RsDetails("Item_ID").Value))
            FG.TextMatrix(Num, FG.ColIndex("Count")) = IIf(IsNull(RsDetails("showqty")), "", (RsDetails("showqty").Value))
            FG.TextMatrix(Num, FG.ColIndex("Price")) = IIf(IsNull(RsDetails("showPrice")), "", (RsDetails("showPrice").Value))
        
            FG.TextMatrix(Num, FG.ColIndex("ItemCase")) = IIf(IsNull(RsDetails("ItemCase")), "", (RsDetails("ItemCase").Value))
            FG.TextMatrix(Num, FG.ColIndex("DiscountType")) = IIf(IsNull(RsDetails("ItemDiscountType")), "", (RsDetails("ItemDiscountType").Value))
            FG.TextMatrix(Num, FG.ColIndex("DiscountVal")) = IIf(IsNull(RsDetails("ItemDiscount")), "", (RsDetails("ItemDiscount").Value))
        
            If RsDetails("HaveSerial") = True Then
                FG.TextMatrix(Num, FG.ColIndex("HaveSerial")) = True
            End If
        
            FG.Cell(flexcpData, Num, FG.ColIndex("UnitID")) = IIf(IsNull(RsDetails("UnitID")), "", (RsDetails("UnitID").Value))
            FG.TextMatrix(Num, FG.ColIndex("UnitID")) = IIf(IsNull(RsDetails("UnitName")), "", (RsDetails("UnitName").Value))
            FG.TextMatrix(Num, FG.ColIndex("ColorID")) = IIf(IsNull(RsDetails("ColorID")), 1, (RsDetails("ColorID").Value))
            FG.TextMatrix(Num, FG.ColIndex("ItemSize")) = IIf(IsNull(RsDetails("ItemSize")), 1, (RsDetails("ItemSize").Value))
            FG.TextMatrix(Num, FG.ColIndex("ClassID")) = IIf(IsNull(RsDetails("ClassID")), 1, (RsDetails("ClassID").Value))
        
            RsDetails.MoveNext
            Debug.Print Num

            If FG.Rows > 10 Then
                If Num = 8 Then FG.Refresh
            End If

        Next Num

    End If

    TxtFillData.Text = "F"
    Screen.MousePointer = vbDefault
    XPTxtCurrent.Caption = rs.AbsolutePosition
    XPTxtCount.Caption = rs.RecordCount
    Exit Sub
ErrTrap:
    Screen.MousePointer = vbDefault
End Sub

Private Sub Undo()
    Dim Msg As String

    On Error GoTo ErrTrap

    Select Case TxtModFlg.Text

        Case "N"
            Msg = "”Ê› Ì „ «· —«Ã⁄ ›Ï  ”ÃÌ· Â–Â «·›« Ê—… .."
            Msg = Msg & Chr(13) & "›Â· «‰  „ «ﬂœ „‰ «·√” „—«— ..!!"

            If MsgBox(Msg, vbQuestion + vbYesNo + vbMsgBoxRight + vbMsgBoxRtlReading + vbDefaultButton2, App.title) = vbYes Then
                clear_all Me
                Me.TxtModFlg.Text = "R"
                XPBtnMove_Click (1)
            End If

        Case "E"
            Msg = "”Ê› Ì „ «· —«Ã⁄ ›Ï  ⁄œÌ· Â–Â «·›« Ê—… .."
            Msg = Msg & Chr(13) & "›Â· «‰  „ «ﬂœ „‰ «·√” „—«— ..!!"

            If MsgBox(Msg, vbQuestion + vbYesNo + vbMsgBoxRight + vbMsgBoxRtlReading + vbDefaultButton2, App.title) = vbYes Then
                rs.find "Transaction_ID='" & val(XPTxtBillID.Text) & "'", , adSearchForward, adBookmarkFirst

                If rs.EOF Or rs.BOF Then
                    Me.TxtModFlg.Text = "R"
                    Exit Sub
                End If

                If Not rs.EOF Or rs.BOF Then
                    Me.TxtModFlg.Text = "R"
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

    If XPTxtBillID.Text <> "" Then
        Msg = "”Ì „ Õ–› »Ì«‰«  «·⁄„·Ì… —ﬁ„ " & Chr(13)
        Msg = Msg + (XPTxtBillID.Text) & Chr(13)
        Msg = Msg + " Â·  —€» ›Ì Õ–› Â–Â «·»Ì«‰« ø"

        If MsgBox(Msg, vbYesNo + vbQuestion + vbMsgBoxRight + vbMsgBoxRtlReading, App.title) = vbYes Then
            If Not rs.RecordCount < 1 Then
                CuurentLogdata ("D")
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
        MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
        TxtModFlg_Change
        Exit Sub
    End If

    TxtModFlg_Change
    Exit Sub
ErrTrap:

    If Err.Number = -2147217887 Then
        Msg = "·«Ì„ﬂ‰ Õ–› Â–« «·”Ã· · ﬂ«„· «·»Ì«‰«  " & Chr(13) & " ÊÃœ »Ì«‰«  „— »ÿ… »Â–« «·„Ê—œ "
        MsgBox Msg, vbMsgBoxRight + vbMsgBoxRtlReading + vbExclamation, App.title
        rs.CancelUpdate
    End If

End Sub

Private Sub AddTip()
    Dim Wrap As String
    On Error GoTo ErrTrap
    Wrap = Chr(13) + Chr(10)
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
    On Error GoTo ErrTrap
    Screen.MousePointer = vbArrowHourglass

    If Me.TxtModFlg.Text <> "R" Then
        If DBCboClientName.Text = "" Then
            If SystemOptions.UserInterface = ArabicInterface Then
                Msg = "„‰ ›÷·ﬂ √œŒ· «”„ «·⁄„Ì·"
            Else
                Msg = "Please Select Vendor"
            End If

            MsgBox Msg, vbOKOnly + vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
            DBCboClientName.SetFocus
            SendKeys "{F4}"
            Screen.MousePointer = vbDefault
            Exit Sub
        End If

        If DCboStoreName.Text = "" Then
            If SystemOptions.UserInterface = ArabicInterface Then
                Msg = "ÌÃ»  ÕœÌœ «·„Œ“‰"
            Else
                Msg = "Select Inventory"
            End If

            MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
            DCboStoreName.SetFocus
            SendKeys "{F4}"
            Screen.MousePointer = vbDefault
            Exit Sub
        End If
    
        If DcCurrency.Text = "" Then
            If SystemOptions.UserInterface = ArabicInterface Then
                Msg = "Õœœ «·⁄„·…"
            Else
                Msg = "Select Currency"
            End If

            MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
            DcCurrency.SetFocus
            SendKeys "{F4}"
            Screen.MousePointer = vbDefault
            Exit Sub
        End If
    
        If Me.CBOPriceType.ListIndex = -1 Then
            If SystemOptions.UserInterface = ArabicInterface Then
                Msg = "ÌÃ»  ÕœÌœ ‰Ê⁄ ⁄—÷ «·”⁄— (⁄—÷ ”⁄— ›« Ê—… »Ì⁄ «Ê ÿ·»Ì… ‘—«¡)...!!!"
            Else
                Msg = "Specify Order Type"
            End If

            MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
            CBOPriceType.SetFocus
            SendKeys "{F4}"
            Screen.MousePointer = vbDefault
            Exit Sub
        End If

        If XPChkTAX.Value = Checked Then
            If XPTxtTaxValue.Text = "" Then
                If SystemOptions.UserInterface = ArabicInterface Then
                    Msg = "ÌÃ» «œŒ«· ﬁÌ„… ÷—Ì»… «·„»Ì⁄« "
                Else
                    Msg = "Insert Sales Tax"
                End If

                MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
                XPTxtTaxValue.SetFocus
                FG.SetFocus
                Screen.MousePointer = vbDefault
                Exit Sub
            End If
        End If
    
        '            If Txt_order_no.text = "" Then
        '        If SystemOptions.UserInterface = ArabicInterface Then
        '            Msg = "ÌÃ» «œŒ«·     —ﬁ„ «·ÿ·»ÌÂ «Ê —ﬁ„ «·›« Ê—… «·„»œ∆ÌÂ "
        '        Else
        '        Msg = "Insert  Order no or Performa INvoice NO"
        '        End If
        '            MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
        '            Txt_order_no.SetFocus
        '
        '
        '            Exit Sub
        '        End If
    
        If NewGrid.CheckDataEntered = False Then
            Exit Sub
        End If

        Set RSTransDetails = New ADODB.Recordset
     '   RSTransDetails.Open "[Transaction_Details]", Cn, adOpenStatic, adLockOptimistic, adCmdTable
       StrSQL = "SELECT     dbo.Transaction_Details.* from dbo.Transaction_Details Where (Transaction_ID = -1)"
   RSTransDetails.Open StrSQL, Cn, adOpenKeyset, adLockOptimistic, adCmdText
   
        Dim Transaction_Type As Integer
        Dim Sanad_No As Integer

        If Me.CBOPriceType.ListIndex = 0 Then
            Transaction_Type = 6
            Sanad_No = 29
        ElseIf Me.CBOPriceType.ListIndex = 1 Then
            Transaction_Type = 17
            Sanad_No = 30
        ElseIf Me.CBOPriceType.ListIndex = 2 Then
            Transaction_Type = 29
            Sanad_No = 31
        End If

        If TxtNoteSerial1.Text = "" Then
            If Voucher_coding(val(my_branch), XPDtbBill.Value, Sanad_No, 170, , Transaction_Type) = "error" Then
                MsgBox " ·« Ì„ﬂ‰ «÷«›…   Â–« «·”‰œ ·«‰ﬂ  ⁄œÌ  «·Õœ «·„”„ÊÕ »… „‰ «·”‰œ«   ": Exit Sub
            Else
                       
                If Voucher_coding(val(my_branch), XPDtbBill.Value, Sanad_No, 170, , Transaction_Type) = "" Then
                    MsgBox " ·«»œ „‰ ﬂ «»… —ﬁ„ «·”‰œ ÌœÊÌ« ﬂ„« Õœœ    " & Chr(13) & " Enter Vchr No": Exit Sub
                Else
                    TxtNoteSerial1.Text = Voucher_coding(val(my_branch), XPDtbBill.Value, Sanad_No, 170, , Transaction_Type)
                End If
            End If
        End If
 
        TXT_order_no = Me.TxtNoteSerial1.Text
 
        Cn.BeginTrans
        BeginTrans = True
    
        If Me.TxtModFlg.Text = "N" Then
            Me.oldtxtNoteSerial1.Text = Trim$(Me.TxtNoteSerial1.Text)
            rs.AddNew
        End If

        Screen.MousePointer = vbArrowHourglass
        rs("NoteSerial1").Value = IIf(Trim(Me.TxtNoteSerial1.Text) = "", Null, Trim(Me.TxtNoteSerial1.Text))
        rs("OldNoteSerial1").Value = Trim$(Me.oldtxtNoteSerial1.Text) '
        rs("branchID").Value = val(Me.dcBranch.BoundText)
   
        rs("Transaction_ID").Value = val(XPTxtBillID.Text)
        rs("order_no").Value = TXT_order_no.Text
    
        If chkshipped.Value = vbChecked Then
            rs("shipped").Value = 1
        Else
            rs("shipped").Value = 0
        End If
    
        rs("Transaction_Date").Value = XPDtbBill.Value
        rs("Transaction_Serial").Value = TxtTransSerial.Text

        If Me.CBOPriceType.ListIndex = 0 Then
            rs("Transaction_Type").Value = 6
        ElseIf Me.CBOPriceType.ListIndex = 1 Then
            rs("Transaction_Type").Value = 17
        ElseIf Me.CBOPriceType.ListIndex = 2 Then
            rs("Transaction_Type").Value = 29
        End If

        rs("UserID").Value = user_id
        rs("CusID").Value = IIf(DBCboClientName.BoundText = "", Null, val(DBCboClientName.BoundText))
        rs("countryid").Value = IIf(DataCombo4.BoundText = "", Null, val(DataCombo4.BoundText))
    
        rs("Currency_id").Value = IIf(DcCurrency.BoundText = "", Null, val(DcCurrency.BoundText))
    
        rs("StoreID").Value = IIf(DCboStoreName.BoundText = "", Null, val(DCboStoreName.BoundText))
        rs("TaxFound").Value = IIf(XPChkTAX.Value = Checked, True, False)
        rs("TaxValue").Value = IIf(XPTxtTaxValue.Text = "", Null, val(XPTxtTaxValue.Text))
        rs("total").Value = IIf(XPTxtSum.Text = "", Null, val(XPTxtSum.Text))
        rs("LcNo").Value = IIf(TxtLcNo.Text = "", Null, (TxtLcNo.Text))
    
        rs.update
    
        CuurentLogdata
  
        If Me.TxtModFlg.Text = "E" Then
            StrSqlDel = "delete From Transaction_Details where Transaction_ID=" & val(rs("Transaction_ID").Value)
            Cn.Execute StrSqlDel, , adExecuteNoRecords
        End If

        For RowNum = 1 To FG.Rows - 1

            If FG.TextMatrix(RowNum, FG.ColIndex("Code")) <> "" Then
                RSTransDetails.AddNew
                RSTransDetails("Transaction_ID").Value = val(XPTxtBillID.Text)
                RSTransDetails("order_id").Value = val(XPTxtBillID.Text)
             
                RSTransDetails("order_no").Value = TXT_order_no.Text
             
                RSTransDetails("Item_ID").Value = IIf((FG.TextMatrix(RowNum, FG.ColIndex("Code")) = ""), Null, val(FG.TextMatrix(RowNum, FG.ColIndex("Code"))))
                RSTransDetails("Quantity").Value = IIf((FG.TextMatrix(RowNum, FG.ColIndex("Count")) = ""), Null, val(FG.TextMatrix(RowNum, FG.ColIndex("Count"))))
                RSTransDetails("ShowPrice").Value = IIf((FG.TextMatrix(RowNum, FG.ColIndex("Price")) = ""), Null, val(FG.TextMatrix(RowNum, FG.ColIndex("Price"))))
                RSTransDetails("ItemDiscountType").Value = IIf((FG.TextMatrix(RowNum, FG.ColIndex("DiscountType")) = ""), Null, val(FG.TextMatrix(RowNum, FG.ColIndex("DiscountType"))))
                RSTransDetails("ItemCase").Value = IIf((FG.TextMatrix(RowNum, FG.ColIndex("ItemCase")) = ""), Null, val(FG.TextMatrix(RowNum, FG.ColIndex("ItemCase"))))
                RSTransDetails("ItemDiscount").Value = IIf((FG.TextMatrix(RowNum, FG.ColIndex("DiscountVal")) = ""), Null, val(FG.TextMatrix(RowNum, FG.ColIndex("DiscountVal"))))
            
                RSTransDetails("ColorID").Value = IIf((FG.TextMatrix(RowNum, FG.ColIndex("ColorID")) = ""), 1, val(FG.TextMatrix(RowNum, FG.ColIndex("ColorID"))))
                RSTransDetails("ItemSize").Value = IIf((FG.TextMatrix(RowNum, FG.ColIndex("ItemSize")) = ""), "", Trim$(FG.TextMatrix(RowNum, FG.ColIndex("ItemSize"))))
                RSTransDetails("ClassId").Value = IIf((FG.TextMatrix(RowNum, FG.ColIndex("ClassId")) = ""), 1, val(FG.TextMatrix(RowNum, FG.ColIndex("ClassId"))))
            
                RSTransDetails("UnitID").Value = IIf(FG.Cell(flexcpData, RowNum, FG.ColIndex("UnitID")) = "", Null, (FG.Cell(flexcpData, RowNum, FG.ColIndex("UnitID"))))
                RSTransDetails("ShowQty").Value = IIf((FG.TextMatrix(RowNum, FG.ColIndex("Count")) = ""), Null, val(FG.TextMatrix(RowNum, FG.ColIndex("Count"))))
 
                Dim RsUnitData As ADODB.Recordset
                Dim LngCurItemID As Long
                Dim LngUnitID As Long
                Dim DblQty As Double
        
                LngCurItemID = val(FG.TextMatrix(RowNum, FG.ColIndex("Code")))
                LngUnitID = val(FG.Cell(flexcpData, RowNum, FG.ColIndex("UnitID")))
                DblQty = val(FG.TextMatrix(RowNum, FG.ColIndex("Count")))

                StrSQL = "Select * From TblItemsUnits Where ItemID=" & LngCurItemID
                StrSQL = StrSQL + " AND UnitID=" & LngUnitID
                Set RsUnitData = New ADODB.Recordset
                RsUnitData.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

                If Not (rs.BOF Or rs.EOF) Then
                    RSTransDetails("QtyBySmalltUnit").Value = RsUnitData("UnitFactor").Value
                    RSTransDetails("Quantity").Value = RSTransDetails("QtyBySmalltUnit").Value * RSTransDetails("showqty").Value
                    'RSTransDetails("Price").value = Val(IIf((Fg.TextMatrix(RowNum, Fg.ColIndex("Price")) = ""), Null, Val(Fg.TextMatrix(RowNum, Fg.ColIndex("Price"))))) / RSTransDetails("Quantity").value
                    RSTransDetails("Price").Value = val(IIf((FG.TextMatrix(RowNum, FG.ColIndex("Price")) = ""), Null, val(FG.TextMatrix(RowNum, FG.ColIndex("Price"))))) / RSTransDetails("QtyBySmalltUnit").Value
                End If

                RSTransDetails.update
            End If

        Next RowNum

        Cn.CommitTrans
        BeginTrans = False
        XPTxtCurrent.Caption = rs.AbsolutePosition
        XPTxtCount.Caption = rs.RecordCount
        lbl(64).Caption = showLabel(TxtNoteSerial1, oldtxtNoteSerial1)
    
        Select Case Me.TxtModFlg.Text

            Case "N"

                If SystemOptions.UserInterface = ArabicInterface Then
                    Msg = "  „ Õ›Ÿ »Ì«‰«  Â–Â «·⁄„·Ì…" & Chr(13)
                    Msg = Msg + "Â·  —€» ›Ì ≈÷«›… »Ì«‰«  √Œ—Ì"
                Else
                    Msg = " Saved Successfully" & Chr(13)
                    Msg = Msg + "do you new Operation?"
        
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
                    MsgBox "Saved Changes Successfully", vbInformation + vbOKOnly + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
                End If

        End Select

        TxtModFlg.Text = "R"
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
            Msg = "·« Ì„ﬂ‰ Õ›Ÿ Â–Â «·»Ì«‰«  " & Chr(13)
            Msg = Msg + "·ﬁœ  „ «œŒ«· ﬁÌ„ €Ì— ’«·Õ… " & Chr(13)
            Msg = Msg + " √ﬂœ „‰ œﬁ… «·»Ì«‰«  Ê√⁄œ «·„Õ«Ê·…"
        Else
    
            Msg = "Cant Save Error"
        End If

        MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
        Exit Sub
    End If

    If SystemOptions.UserInterface = ArabicInterface Then
        Msg = "⁄›Ê«...ÕœÀ Œÿ√ „« √À‰«¡ Õ›Ÿ Â–Â «·»Ì«‰«  " & Chr(13)
    Else
        Msg = "Sorry... Error During Saving " & Chr(13)
    End If

    MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
End Sub

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

    If XPChkTAX.Value = Checked Then
        XPTxtTaxValue.Enabled = True
        XPTxtTaxValue.locked = False
        lbl(4).Enabled = True
    Else
        XPTxtTaxValue.Text = ""
        XPTxtTaxValue.Enabled = False
        lbl(4).Enabled = False
    End If

    Exit Sub
ErrTrap:
End Sub

Private Sub PrintReport()
    On Error GoTo ErrTrap

    If XPTxtBillID.Text <> "" Then
        Set SaleReport = New ClsSaleReport
        SaleReport.ShowPrice XPTxtBillID.Text
    End If

    Exit Sub
ErrTrap:
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, _
                             UnloadMode As Integer)
    Dim StrMSG As String
    Dim IntResult As String
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

Public Sub Cala()
    NewGrid.Calculate 1
End Sub

Private Sub XPDtbBill_Change()

    If Trim(TxtNoteSerial1.Text) <> "" Then
        oldtxtNoteSerial1.Text = TxtNoteSerial1.Text
    End If

    TxtNoteSerial1.Text = ""
 
End Sub

Private Sub XPTxtTaxValue_Change()

    If Me.TxtModFlg.Text = "N" Or Me.TxtModFlg.Text = "E" Then
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
    Me.Caption = "Order Request/Proforma   Invoice"

    Ele(6).Caption = Me.Caption
    lbl(5).Caption = "Ord/P INV. No"
    Frame3.Caption = "LC Data"
    ISButton1.Caption = "View"
    lbl(25).Caption = "Total"
    lbl(63).Caption = "Qty"
    Label2.Caption = "Branch"
    lbl(6).Caption = "Date"
    lbl(7).Caption = "Client"
    lbl(8).Caption = "Store"
    lbl(9).Caption = "Type"
    lbl(10).Caption = "Cost Center"
    lbl(11).Caption = "Project"
    lbl(16).Caption = "Article Section"
    lbl(12).Caption = "Currency"
    lbl(13).Caption = "Country"
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
    'ISButton1.Caption = "Show Port Data"
    Label1.Caption = "LC NO:"
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
    Cmd(6).Caption = "Exit"
    CmdHelp.Caption = "Help"

    CmdConvert.Caption = "Convert To Bill"
    CmdTemplate.Caption = "Insert template"

End Sub
