VERSION 5.00
Object = "{C115893A-A3BF-43AF-B28D-69DB846077F3}#1.0#0"; "vsflex8u.ocx"
Object = "{0AFE7BE0-11B7-4A3E-978D-D4501E9A57FE}#1.0#0"; "c1sizer.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "msdatlst.ocx"
Object = "{FE5DCFAD-BC1D-11D2-94CF-004005455FAA}#1.4#0"; "ImpulseButton.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form FrmGard1 
   BackColor       =   &H00E2E9E9&
   Caption         =   "Ã—œ «·„Œ«“‰"
   ClientHeight    =   8610
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11835
   HelpContextID   =   170
   Icon            =   "FrmGard1.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   RightToLeft     =   -1  'True
   ScaleHeight     =   8610
   ScaleWidth      =   11835
   Begin C1SizerLibCtl.C1Elastic EleMain 
      Height          =   8610
      Left            =   0
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   0
      Width           =   11835
      _cx             =   20876
      _cy             =   15187
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
      GridRows        =   5
      GridCols        =   4
      Frame           =   3
      FrameStyle      =   0
      FrameWidth      =   1
      FrameColor      =   -2147483628
      FrameShadow     =   -2147483632
      FloodStyle      =   1
      _GridInfo       =   $"FrmGard1.frx":038A
      AccessibleName  =   ""
      AccessibleDescription=   ""
      AccessibleValue =   ""
      AccessibleRole  =   9
      Begin C1SizerLibCtl.C1Elastic EleFooter 
         Height          =   945
         Left            =   30
         TabIndex        =   10
         TabStop         =   0   'False
         Top             =   7635
         Width           =   11775
         _cx             =   20770
         _cy             =   1667
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
         Caption         =   "0"
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
         Begin MSDataListLib.DataCombo DCboUserName 
            Height          =   315
            Left            =   90
            TabIndex        =   11
            Top             =   90
            Width           =   1545
            _ExtentX        =   2725
            _ExtentY        =   556
            _Version        =   393216
            Enabled         =   0   'False
            Text            =   ""
            RightToLeft     =   -1  'True
         End
         Begin ImpulseButton.ISButton CmdHelp 
            Height          =   375
            Left            =   1455
            TabIndex        =   12
            Top             =   480
            Width           =   1695
            _ExtentX        =   2990
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
            Index           =   0
            Left            =   10560
            TabIndex        =   13
            Top             =   480
            Width           =   1185
            _ExtentX        =   2090
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
            Index           =   2
            Left            =   9180
            TabIndex        =   14
            Top             =   480
            Width           =   1155
            _ExtentX        =   2037
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
            Left            =   7710
            TabIndex        =   15
            Top             =   480
            Width           =   1275
            _ExtentX        =   2249
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
            Left            =   6300
            TabIndex        =   16
            Top             =   480
            Width           =   1170
            _ExtentX        =   2064
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
            Index           =   5
            Left            =   4815
            TabIndex        =   17
            Top             =   480
            Width           =   1260
            _ExtentX        =   2223
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
            Index           =   6
            Left            =   60
            TabIndex        =   18
            Top             =   480
            Width           =   1200
            _ExtentX        =   2117
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
         Begin ImpulseButton.ISButton Cmd 
            Height          =   375
            Index           =   7
            Left            =   3330
            TabIndex        =   19
            Top             =   480
            Width           =   1275
            _ExtentX        =   2249
            _ExtentY        =   661
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
            ColorShadow     =   -2147483637
            ColorOutline    =   0
            DrawFocusRectangle=   0   'False
            ColorToggledHoverText=   16711680
            ColorTextShadow =   -2147483637
         End
         Begin VB.Label Label200 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Height          =   270
            Index           =   3
            Left            =   3060
            RightToLeft     =   -1  'True
            TabIndex        =   43
            Top             =   120
            Width           =   660
         End
         Begin VB.Label Label2 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "«Ã„«·Ì «·ﬂ„ÌÂ"
            Height          =   270
            Index           =   2
            Left            =   3915
            RightToLeft     =   -1  'True
            TabIndex        =   42
            Top             =   120
            Width           =   1065
         End
         Begin VB.Label Label2 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "«·”Ã· «·Õ«·Ì:"
            Height          =   345
            Index           =   0
            Left            =   9945
            RightToLeft     =   -1  'True
            TabIndex        =   24
            Top             =   90
            Width           =   1740
         End
         Begin VB.Label Label3 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "⁄œœ «·”Ã·« :"
            Height          =   315
            Left            =   7185
            RightToLeft     =   -1  'True
            TabIndex        =   23
            Top             =   120
            Width           =   1080
         End
         Begin VB.Label XPTxtCurrent 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Height          =   315
            Left            =   8940
            RightToLeft     =   -1  'True
            TabIndex        =   22
            Top             =   120
            Width           =   1005
         End
         Begin VB.Label XPTxtCount 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Height          =   315
            Left            =   6630
            RightToLeft     =   -1  'True
            TabIndex        =   21
            Top             =   120
            Width           =   555
         End
         Begin VB.Label Label2 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Õ—— »Ê«”ÿ… : "
            Height          =   270
            Index           =   1
            Left            =   1620
            RightToLeft     =   -1  'True
            TabIndex        =   20
            Top             =   120
            Width           =   1020
         End
      End
      Begin C1SizerLibCtl.C1Elastic EleTop 
         Height          =   1125
         Left            =   30
         TabIndex        =   8
         TabStop         =   0   'False
         Top             =   750
         Width           =   11775
         _cx             =   20770
         _cy             =   1984
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
         Begin VB.Frame Frame2 
            BackColor       =   &H00E2E9E9&
            Caption         =   "›Ì «·› —…"
            ForeColor       =   &H00FF0000&
            Height          =   1125
            Index           =   0
            Left            =   0
            RightToLeft     =   -1  'True
            TabIndex        =   47
            Top             =   0
            Width           =   2535
            Begin MSComCtl2.DTPicker DTPFrom 
               Height          =   345
               Left            =   120
               TabIndex        =   48
               Top             =   240
               Width           =   1665
               _ExtentX        =   2937
               _ExtentY        =   609
               _Version        =   393216
               CheckBox        =   -1  'True
               CustomFormat    =   "dd/m/yyyy"
               DateIsNull      =   -1  'True
               Format          =   100073473
               CurrentDate     =   36494
            End
            Begin MSComCtl2.DTPicker DTPTo 
               Height          =   345
               Left            =   120
               TabIndex        =   49
               Top             =   630
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
               Index           =   23
               Left            =   1950
               RightToLeft     =   -1  'True
               TabIndex        =   51
               Top             =   675
               Width           =   345
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "„‰"
               Height          =   285
               Index           =   24
               Left            =   1980
               RightToLeft     =   -1  'True
               TabIndex        =   50
               Top             =   255
               Width           =   285
            End
         End
         Begin VB.TextBox txtorder_no 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   5040
            RightToLeft     =   -1  'True
            TabIndex        =   44
            Top             =   240
            Width           =   1575
         End
         Begin VB.ComboBox Combo1 
            Height          =   315
            Left            =   210
            RightToLeft     =   -1  'True
            TabIndex        =   39
            Text            =   "Combo1"
            Top             =   570
            Visible         =   0   'False
            Width           =   1470
         End
         Begin VB.TextBox TxtTransSerial 
            Alignment       =   1  'Right Justify
            Enabled         =   0   'False
            Height          =   375
            Left            =   9825
            RightToLeft     =   -1  'True
            TabIndex        =   36
            Top             =   60
            Width           =   885
         End
         Begin ImpulseButton.ISButton CmdDo 
            Height          =   435
            Left            =   2490
            TabIndex        =   32
            Top             =   570
            Width           =   1395
            _ExtentX        =   2461
            _ExtentY        =   767
            Caption         =   "»œ¡ ⁄„·Ì… «·Ã—œ"
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
         End
         Begin C1SizerLibCtl.C1Elastic Ele 
            Height          =   525
            Left            =   8460
            TabIndex        =   29
            TabStop         =   0   'False
            Top             =   390
            Width           =   3285
            _cx             =   5794
            _cy             =   926
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   178
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Enabled         =   0   'False
            Appearance      =   4
            MousePointer    =   0
            Version         =   801
            BackColor       =   14871017
            ForeColor       =   -2147483630
            FloodColor      =   6553600
            ForeColorDisabled=   -2147483631
            Caption         =   "‰Ê⁄ «·Ã—œ"
            Align           =   0
            AutoSizeChildren=   0
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
            Begin VB.OptionButton XPOptShowType 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "ÿ»ﬁ« ··‘Õ‰Â"
               CausesValidation=   0   'False
               Height          =   225
               Index           =   1
               Left            =   180
               RightToLeft     =   -1  'True
               TabIndex        =   31
               Top             =   240
               Width           =   1215
            End
            Begin VB.OptionButton XPOptShowType 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   " ›’Ì·Ì"
               Height          =   255
               Index           =   0
               Left            =   1260
               RightToLeft     =   -1  'True
               TabIndex        =   30
               Top             =   210
               Value           =   -1  'True
               Width           =   1455
            End
         End
         Begin MSDataListLib.DataCombo DCboStoreName 
            Height          =   315
            Left            =   3930
            TabIndex        =   25
            Top             =   570
            Width           =   2670
            _ExtentX        =   4710
            _ExtentY        =   556
            _Version        =   393216
            Text            =   ""
            RightToLeft     =   -1  'True
         End
         Begin MSComCtl2.DTPicker XPDtbGard 
            Height          =   345
            Left            =   7695
            TabIndex        =   26
            Top             =   60
            Width           =   1455
            _ExtentX        =   2566
            _ExtentY        =   609
            _Version        =   393216
            Format          =   100073473
            CurrentDate     =   38784
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "«·‘Õ‰Â"
            Height          =   315
            Index           =   9
            Left            =   6600
            RightToLeft     =   -1  'True
            TabIndex        =   46
            Top             =   240
            Width           =   990
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "ÿ—Ìﬁ… Õ”«» «· ﬂ·›…"
            Height          =   375
            Index           =   7
            Left            =   810
            RightToLeft     =   -1  'True
            TabIndex        =   40
            Top             =   510
            Visible         =   0   'False
            Width           =   990
         End
         Begin VB.Image Img 
            Height          =   480
            Left            =   60
            Picture         =   "FrmGard1.frx":0417
            Top             =   30
            Visible         =   0   'False
            Width           =   480
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H00C0FFFF&
            BorderStyle     =   1  'Fixed Single
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
            Height          =   465
            Index           =   4
            Left            =   90
            RightToLeft     =   -1  'True
            TabIndex        =   38
            Top             =   30
            Visible         =   0   'False
            Width           =   1110
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "„·ÕÊŸ… Â«„…:"
            Height          =   225
            Index           =   3
            Left            =   1050
            RightToLeft     =   -1  'True
            TabIndex        =   37
            Top             =   90
            Visible         =   0   'False
            Width           =   1185
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "—ﬁ„ «·⁄„·Ì…"
            Height          =   255
            Index           =   2
            Left            =   10830
            RightToLeft     =   -1  'True
            TabIndex        =   35
            Top             =   120
            Width           =   825
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "«· «—ÌŒ"
            Height          =   315
            Index           =   6
            Left            =   9240
            RightToLeft     =   -1  'True
            TabIndex        =   28
            Top             =   90
            Width           =   525
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "«”„ «·„Œ“‰"
            Height          =   315
            Index           =   5
            Left            =   6690
            RightToLeft     =   -1  'True
            TabIndex        =   27
            Top             =   690
            Width           =   990
         End
      End
      Begin C1SizerLibCtl.C1Elastic EleHeader 
         Height          =   705
         Left            =   30
         TabIndex        =   1
         TabStop         =   0   'False
         Top             =   30
         Width           =   11775
         _cx             =   20770
         _cy             =   1244
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial (Arabic)"
            Size            =   24
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
         BackColor       =   16777215
         ForeColor       =   4210688
         FloodColor      =   6553600
         ForeColorDisabled=   -2147483631
         Caption         =   "Ã—œ «·„Œ«“‰"
         Align           =   0
         AutoSizeChildren=   0
         BorderWidth     =   6
         ChildSpacing    =   4
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
            Height          =   345
            Left            =   3780
            RightToLeft     =   -1  'True
            TabIndex        =   3
            Top             =   90
            Visible         =   0   'False
            Width           =   855
         End
         Begin VB.TextBox XPTxtBillID 
            Alignment       =   1  'Right Justify
            Height          =   315
            Left            =   3690
            RightToLeft     =   -1  'True
            TabIndex        =   2
            Top             =   120
            Visible         =   0   'False
            Width           =   945
         End
         Begin ImpulseButton.ISButton XPBtnMove 
            Height          =   345
            Index           =   0
            Left            =   1215
            TabIndex        =   4
            Top             =   150
            Width           =   495
            _ExtentX        =   873
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
            ButtonImage     =   "FrmGard1.frx":10E1
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
            Left            =   150
            TabIndex        =   5
            Top             =   150
            Width           =   495
            _ExtentX        =   873
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
            ButtonImage     =   "FrmGard1.frx":147B
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
            Left            =   1740
            TabIndex        =   6
            Top             =   150
            Width           =   495
            _ExtentX        =   873
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
            ButtonImage     =   "FrmGard1.frx":1815
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
            Left            =   675
            TabIndex        =   7
            Top             =   150
            Width           =   495
            _ExtentX        =   873
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
            ButtonImage     =   "FrmGard1.frx":1BAF
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
      Begin VSFlex8UCtl.VSFlexGrid FG 
         Height          =   5730
         Left            =   30
         TabIndex        =   9
         Top             =   1890
         Width           =   11775
         _cx             =   20770
         _cy             =   10107
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
         Cols            =   11
         FixedRows       =   1
         FixedCols       =   0
         RowHeightMin    =   300
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   -1  'True
         FormatString    =   $"FrmGard1.frx":1F49
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
         WallPaperAlignment=   0
         AccessibleName  =   ""
         AccessibleDescription=   ""
         AccessibleValue =   ""
         AccessibleRole  =   24
      End
      Begin MSComctlLib.ProgressBar PrgBr 
         Height          =   5730
         Left            =   30
         TabIndex        =   41
         Top             =   1890
         Width           =   11775
         _ExtentX        =   20770
         _ExtentY        =   10107
         _Version        =   393216
         Appearance      =   0
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "ﬁÌ„… «·„Œ“Ê‰:"
         ForeColor       =   &H000000FF&
         Height          =   285
         Index           =   1
         Left            =   30
         RightToLeft     =   -1  'True
         TabIndex        =   34
         Top             =   7335
         Width           =   4470
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "⁄œœ «·√’‰«›:"
         ForeColor       =   &H000000FF&
         Height          =   285
         Index           =   0
         Left            =   8865
         RightToLeft     =   -1  'True
         TabIndex        =   33
         Top             =   7335
         Width           =   2940
      End
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   "«”„ «·„Œ“‰"
      Height          =   315
      Index           =   8
      Left            =   0
      RightToLeft     =   -1  'True
      TabIndex        =   45
      Top             =   0
      Width           =   990
   End
End
Attribute VB_Name = "FrmGard1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rs As ADODB.Recordset
Dim TTP As clstooltip
Dim VReport As ClsGardReport
Dim cSearchDcbo  As clsDCboSearch

Private Sub CmdDo_Click()

    DoStockCount
End Sub

Private Sub CmdHelp_Click()
    SystemOptions.SysHelp.HHTopicID = Me.HelpContextID
    SystemOptions.SysHelp.HHDisplayTopicID Me.hWnd
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
    'CmdConvert.Caption = "Convert to bill"
    Dim XPic As IPictureDisp
    Set XPic = Me.XPBtnMove(1).ButtonImage
    Set Me.XPBtnMove(1).ButtonImage = Me.XPBtnMove(2).ButtonImage
    Set Me.XPBtnMove(2).ButtonImage = XPic

    Set XPic = Me.XPBtnMove(0).ButtonImage
    Set Me.XPBtnMove(0).ButtonImage = Me.XPBtnMove(3).ButtonImage
    Set Me.XPBtnMove(3).ButtonImage = XPic

    Me.Caption = "Inventory Count"
    EleHeader.Caption = Me.Caption

    lbl(2).Caption = "ID"
    lbl(0).Caption = "Items Count"
    lbl(1).Caption = "Value"

    lbl(6).Caption = "Date"
    Ele.Caption = "Type"
    lbl(5).Caption = "Store "
    XPOptShowType(0).Caption = "Detailed"
    XPOptShowType(1).Caption = "Total"
    lbl(7).Caption = "Cost cal."
    CmdDo.Caption = "Start"

    Label2(0).Caption = "Curr Rec."
    Label2(1).Caption = "By"
    Label3.Caption = "Rec. Count:"
    
    Me.Cmd(0).Caption = "New"
    'Me.Cmd(1).Caption = "Edit"
    Me.Cmd(2).Caption = "Save"
    Me.Cmd(3).Caption = "Undo"
    Me.Cmd(4).Caption = "Delete"
    Me.Cmd(5).Caption = "Search"
    Me.Cmd(6).Caption = "Exit"
    Me.Cmd(7).Caption = "Print"
    Me.CmdHelp.Caption = "Help"
     
    With FG
        .TextMatrix(0, .ColIndex("Index")) = "Index"
        .TextMatrix(0, .ColIndex("ItemID")) = "ItemID"
        .TextMatrix(0, .ColIndex("ItemCode")) = "ItemCode"
        .TextMatrix(0, .ColIndex("Name")) = "Name"
        .TextMatrix(0, .ColIndex("Count")) = "Count"
        .TextMatrix(0, .ColIndex("CostPrice")) = "CostPrice"
        .TextMatrix(0, .ColIndex("ItemValue")) = "ItemValue"
    End With
    
End Sub

Private Sub Form_Load()
    Dim StrSQL As String
    Dim BGround As New ClsBackGroundPic
    Dim RsItems As New ADODB.Recordset
    Dim Dcombos As ClsDataCombos
    Dim Msg As String

    On Error GoTo ErrTrap
    Set Cmd(0).ButtonImage = mdifrmmain.ImgLstTree.ListImages("New").Picture
    Set Cmd(2).ButtonImage = mdifrmmain.ImgLstTree.ListImages("save").Picture
    Set Cmd(3).ButtonImage = mdifrmmain.ImgLstTree.ListImages("Undo").Picture
    Set Cmd(4).ButtonImage = mdifrmmain.ImgLstTree.ListImages("Del").Picture
    Set Cmd(5).ButtonImage = mdifrmmain.ImgLstTree.ListImages("Search").Picture
    Set Cmd(6).ButtonImage = mdifrmmain.ImgLstTree.ListImages("Exit").Picture
    Set Cmd(7).ButtonImage = mdifrmmain.ImgLstTree.ListImages("Print").Picture
    Set CmdHelp.ButtonImage = mdifrmmain.ImgLstTree.ListImages("Help").Picture
    SetDtpickerDate Me.XPDtbGard
    AddTip
    FG.WallPaper = BGround.Picture
    Resize_Form Me, TransactionSize
    Set Dcombos = New ClsDataCombos
    Dcombos.GetStores Me.DCboStoreName
    Dcombos.GetUsers DCboUserName

    StrSQL = "Select * From TblItems Order BY ItemCode"
    RsItems.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText
    StrList = FG.BuildComboList(RsItems, "ItemName", "ItemID")

    If StrList <> "" Then
        FG.ColComboList(FG.ColIndex("Name")) = "|" & StrList
    End If

    Set cSearchDcbo = New clsDCboSearch
    Set cSearchDcbo.Client = Me.DCboStoreName
    StrSQL = "SELECT * FROM Transactions WHERE Transaction_Type=4"
    Set rs = New ADODB.Recordset
    rs.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText
    Retrive
    Me.TxtModFlg.text = "R"
    '--------
    Me.PrgBr.Visible = False
    '-------
    Msg = "ÌÃ» „·«ÕŸ… «‰  ﬂ·›… «·„Œ“Ê‰ „„ﬂ‰ «‰  Œ ·› ›Ï Õ«·… «·Ã—œ "
    Msg = Msg & "«· ›’Ì·Ì ⁄‰ «·Ã—œ «·√Ã„«·Ï Ê–·ﬂ ·«‰Â ›Ï «·‰Ê⁄ «·√Ê· ›«‰ «·»—‰«„Ã ÌﬁÊ„ "
    Msg = Msg & "»Õ”«» ”⁄—  ﬂ·›… «·„Œ“Ê‰( »‰«¡ ⁄·Ï ”⁄— «·‘—«¡) ··’‰› »œ·«·… «·”Ì—Ì«· ‰„»— «·Œ«’ »ﬂ· ﬁÿ⁄… „‰ ﬁÿ⁄ «·’‰› "
    Msg = Msg & "Ê·ﬂ‰ ›Ï Õ«·… «·Ã—œ «·√Ã„«·Ï ›«‰  ﬂ·›… «·„Œ“Ê‰  Õ”» »‰«¡ ⁄·Ï «Œ— ”⁄— ‘—«¡ ”Ã· ··’‰›"
    lbl(4).Caption = Msg
    Resize_Form Me, TransactionSize

    'Me.Width = 10695
    'Me.Height = 9120
    If SystemOptions.UserInterface = EnglishInterface Then
        SetInterface Me
        ChangeLang
    End If

    If OPEN_NEW_SCREEN = True Then
        Cmd_Click (0)
    End If

    Exit Sub
ErrTrap:
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

Private Sub Cmd_Click(Index As Integer)
    On Error GoTo ErrTrap

    Select Case Index

        Case 0
            clear_all Me
            TxtModFlg.text = "N"
            Me.DCboUserName.BoundText = user_id
            XPOptShowType(1).value = True

        Case 7
            printing

        Case 2
            SaveData

        Case 3
            Call Undo

        Case 4
            Del_TransAction

        Case 5
            FrmGardSearch.Show vbModal

        Case 6
            Unload Me
    End Select

    Exit Sub
ErrTrap:
End Sub

Private Sub Form_Unload(Cancel As Integer)
    On Error GoTo ErrTrap

    If rs.State = adStateOpen Then
        If Not (rs.EOF Or rs.BOF) Then
            If rs.EditMode <> adEditNone Then
                rs.CancelUpdate
            End If
        End If

        rs.Close
    End If

    Set cSearchDcbo = Nothing
    Set rs = Nothing
    Set TTP = Nothing
    Set VReport = Nothing
    Exit Sub
ErrTrap:
End Sub

Private Sub TxtModFlg_Change()
    On Error GoTo ErrTrap

    Select Case Me.TxtModFlg.text

        Case "R"

            If SystemOptions.UserInterface = ArabicInterface Then
                Me.Caption = "Ã—œ «·„Œ«“‰"
            Else
                Me.Caption = "Stock count"
            End If
        
            Me.Cmd(2).Enabled = False
            Me.Cmd(3).Enabled = False
        
            Me.Cmd(0).Enabled = True
            Me.Cmd(4).Enabled = True
            Me.Cmd(5).Enabled = True
            Me.Cmd(7).Enabled = True
            Me.XPBtnMove(0).Enabled = True
            Me.XPBtnMove(1).Enabled = True
            Me.XPBtnMove(2).Enabled = True
            Me.XPBtnMove(3).Enabled = True
        
            '        FG.Editable = flexEDNone
            XPDtbGard.Enabled = False
            DCboStoreName.locked = True
            XPOptShowType(0).Enabled = False
            XPOptShowType(1).Enabled = False

            If rs.RecordCount < 1 Then
                Me.XPBtnMove(0).Enabled = False
                Me.XPBtnMove(1).Enabled = False
                Me.XPBtnMove(2).Enabled = False
                Me.XPBtnMove(3).Enabled = False
                Me.Cmd(4).Enabled = False
                Me.Cmd(5).Enabled = False
                Me.Cmd(7).Enabled = False
            End If

        Case "N"

            If SystemOptions.UserInterface = ArabicInterface Then
                Me.Caption = "Ã—œ «·„Œ«“‰( ÃœÌœ )"
            Else
                Me.Caption = "Stock count(New)"
            End If
        
            Me.Cmd(2).Enabled = True
            Me.Cmd(3).Enabled = True
        
            Me.Cmd(0).Enabled = False
            Me.Cmd(4).Enabled = False
            Me.Cmd(5).Enabled = False
            Me.Cmd(7).Enabled = False
            ' Me.XPBtnMove(0).Enabled = False
            ' Me.XPBtnMove(1).Enabled = False
            ' Me.XPBtnMove(2).Enabled = False
            ' Me.XPBtnMove(3).Enabled = False
        
            '        FG.Editable = flexEDKbdMouse
            XPOptShowType(1).value = True
            XPDtbGard.Enabled = True
            DCboStoreName.locked = False
            XPOptShowType(0).Enabled = True
            XPOptShowType(1).Enabled = True
            XPOptShowType(0).value = True
            XPDtbGard.value = Date

        Case "E"

            If SystemOptions.UserInterface = ArabicInterface Then
                Me.Caption = "Ã—œ «·„Œ«“‰(  ⁄œÌ· )"
            Else
                Me.Caption = "Stock count(edit)"
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
        
            '        FG.Editable = flexEDKbdMouse
            XPDtbGard.Enabled = True
            DCboStoreName.locked = False
            XPOptShowType(0).Enabled = True
            XPOptShowType(1).Enabled = True
    End Select

    Exit Sub
ErrTrap:
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
    On Error GoTo ErrTrap

    If XPTxtBillID.text <> "" Then
        Msg = "”Ì „ Õ–› »Ì«‰«  «·⁄„·Ì… —ﬁ„ " & Chr(13)
        Msg = Msg + (TxtTransSerial.text) & Chr(13)
        Msg = Msg + " Â·  —€» ›Ì Õ–› Â–Â «·»Ì«‰« ø"

        If MsgBox(Msg, vbYesNo + vbQuestion + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title) = vbYes Then
            If Not rs.RecordCount < 1 Then
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
        Msg = "·«Ì„ﬂ‰ Õ–› Â–« «·”Ã· · ﬂ«„· «·»Ì«‰«  " & Chr(13) & " ÊÃœ »Ì«‰«  „— »ÿ… »Â–« «·„Ê—œ "
        MsgBox Msg, vbMsgBoxRight + vbMsgBoxRtlReading + vbExclamation, App.Title
        rs.CancelUpdate
    End If

End Sub

Private Sub AddTip()
    Dim Wrap As String
    On Error GoTo ErrTrap
    Wrap = Chr(13) + Chr(10)
    Set TTP = New clstooltip

    With TTP
        .Create Me.hWnd, "⁄„·Ì«  «·Ã—œ", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
        .AddControl Cmd(0), "ÃœÌœ ..." & Wrap & "·»œ¡ ⁄„·Ì… Ã—œ ÃœÌœ…" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
    End With

    With TTP
        .Create Me.hWnd, "⁄„·Ì«  «·Ã—œ", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
        .AddControl Cmd(7), "ÿ»«⁄… ..." & Wrap & "·⁄—÷ «·»Ì«‰«  «·Õ«·Ì… ›Ì  ﬁ—Ì— " & Wrap & " Ì„ﬂ‰ ÿ»«⁄ Â ⁄‰ ÿ—Ìﬁ «·ÿ«»⁄…", True
    End With

    'With TTP
    '   .Create Me.hwnd, "⁄„·Ì«  «·Ã—œ", 1, 15204351, -2147483630
    '   .MaxWidth = 4000
    '   .VisibleTime = 9000
    '   .DelayTime = 600
    '   .AddControl Cmd(1), _
    '    " ⁄œÌ· ..." & Wrap & _
    '    "· ⁄œÌ· »Ì«‰«  ⁄„·Ì… «·Ã—œ" & Wrap & _
    '    " ›ﬁÿ ≈÷€ÿ Â‰«", True
    'End With
    With TTP
        .Create Me.hWnd, "⁄„·Ì«  «·Ã—œ", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
        .AddControl Cmd(2), "Õ›Ÿ ..." & Wrap & "·Õ›Ÿ »Ì«‰«  ⁄„·Ì… «·Ã—œ «·ÃœÌœ…" & Wrap & "·Õ›Ÿ «· ⁄œÌ·« " & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
    End With

    With TTP
        .Create Me.hWnd, "⁄„·Ì«  «·Ã—œ", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
        .AddControl Cmd(3), " —«Ã⁄ ..." & Wrap & "·· —«Ã⁄ ⁄‰ ⁄„·Ì… «·Ã—œ" & Wrap & "··· —«Ã⁄ ⁄‰ ⁄„·Ì… «· ⁄œÌ·" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
    End With

    With TTP
        .Create Me.hWnd, "⁄„·Ì«  «·Ã—œ", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
        .AddControl Cmd(4), "Õ–› ..." & Wrap & "·Õ–› »Ì«‰«  ⁄„·Ì… «·Ã—œ" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
    End With

    With TTP
        .Create Me.hWnd, "⁄„·Ì«  «·Ã—œ", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
        .AddControl Cmd(5), "»ÕÀ ..." & Wrap & "···»ÕÀ ⁄‰ ⁄„·Ì… Ã—œ" & Wrap & "Ì‰ÿ»ﬁ ⁄·ÌÂ« ‘—Êÿ „⁄Ì‰…" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
    End With

    With TTP
        .Create Me.hWnd, "⁄„·Ì«  «·Ã—œ", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
        .AddControl Cmd(6), "Œ—ÊÃ ..." & Wrap & "·«€·«ﬁ Â–Â «·‰«›–…" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
    End With

    With TTP
        .Create Me.hWnd, "⁄„·Ì«  «·Ã—œ", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
        .AddControl XPBtnMove(1), "«·√Ê· ..." & Wrap & "··«‰ ﬁ«· «·Ï √Ê· ”Ã·" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
    End With

    With TTP
        .Create Me.hWnd, "⁄„·Ì«  «·Ã—œ", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
        .AddControl XPBtnMove(0), "«·”«»ﬁ ..." & Wrap & "··«‰ ﬁ«· «·Ï «·”Ã· «·”«»ﬁ" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
    End With

    With TTP
        .Create Me.hWnd, "⁄„·Ì«  «·Ã—œ", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
        .AddControl XPBtnMove(3), "«· «·Ì ..." & Wrap & "··«‰ ﬁ«· «·Ï «·”Ã· «· «·Ì" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
    End With

    With TTP
        .Create Me.hWnd, "⁄„·Ì«  «·Ã—œ", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
        .AddControl XPBtnMove(2), "«·√ŒÌ— ..." & Wrap & "··«‰ ﬁ«· «·Ï ¬Œ— ”Ã·" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
    End With

    With TTP
        .Create Me.hWnd, "⁄„·Ì«  «·Ã—œ", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
        .AddControl CmdHelp, "„”«⁄œ… ..." & Wrap & "·· ⁄—› ⁄·Ï ÊŸÌ›… Â–Â «·‰«›–…" & Wrap & "ÊﬂÌ›Ì… «· ⁄«„· „⁄Â«" & Wrap & "≈÷€ÿ Â‰«" & Wrap, True
    End With

    Exit Sub
ErrTrap:
End Sub

Private Sub SaveData()
    Dim RsDetails As New ADODB.Recordset
    Dim RowNum As Integer
    Dim RsTemp As New ADODB.Recordset
    Dim StrSQL As String
    Dim Msg As String
    Dim BeginTrans As Boolean
    On Error GoTo ErrTrap
    Screen.MousePointer = vbArrowHourglass

    If Me.TxtModFlg.text <> "R" Then
        If DCboStoreName.text = "" Then
            Msg = "Õœœ «·„Œ“‰ «·–Ì  —€» ›Ì Ã—œÂ"
            MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
            DCboStoreName.SetFocus
            SendKeys "{F4}"
            Exit Sub
        End If

        If TxtModFlg.text = "N" Then
            XPTxtBillID.text = CStr(new_id("Transactions", "Transaction_ID", "", True))
            TxtTransSerial.text = CStr(new_id("Transactions", "Transaction_Serial", "", True, "Transaction_Type=4"))
            rs.AddNew
        End If

        RsDetails.Open "[Transaction_Details]", Cn, adOpenStatic, adLockOptimistic, adCmdTable
        Cn.BeginTrans
        BeginTrans = True
        rs("Transaction_ID").value = val(XPTxtBillID.text)
        rs("order_no").value = TXTOrDer_no.text
        
        rs("Transaction_Serial").value = TxtTransSerial.text
        rs("Transaction_Date").value = XPDtbGard.value
        rs("Transaction_Type").value = 4
        rs("UserID").value = user_id
        rs("StoreID").value = IIf(DCboStoreName.BoundText = "", Null, val(DCboStoreName.BoundText))

        If XPOptShowType(1).value = True Then
            rs("ExtraOptions").value = 0
        Else
            rs("ExtraOptions").value = 1
        End If

        rs.update

        'SAVING DETAILS
        For RowNum = 1 To FG.Rows - 1
            RsDetails.AddNew

            With FG
                RsDetails("Transaction_ID").value = val(XPTxtBillID.text)
                
                RsDetails("order_no").value = IIf(.TextMatrix(RowNum, .ColIndex("order_no")) = "", Null, .TextMatrix(RowNum, .ColIndex("order_no")))
                RsDetails("Item_ID").value = IIf(.TextMatrix(RowNum, .ColIndex("Name")) = "", Null, .TextMatrix(RowNum, .ColIndex("Name")))
                RsDetails("Quantity").value = IIf(.TextMatrix(RowNum, .ColIndex("Count")) = "", Null, .TextMatrix(RowNum, .ColIndex("Count")))
                RsDetails("Price").value = val(.TextMatrix(RowNum, .ColIndex("CostPrice")))

                If XPOptShowType(0).value = True Then
                    If .Rowdata(RowNum) <> "" Then
                        RsDetails("ItemSerial").value = IIf(.TextMatrix(RowNum, .ColIndex("Serial")) = "", Null, .TextMatrix(RowNum, .ColIndex("Serial")))
                    End If
                End If

            End With

            RsDetails.update
        Next RowNum

        Cn.CommitTrans
        BeginTrans = False
        XPTxtCurrent.Caption = rs.AbsolutePosition
        XPTxtCount.Caption = rs.RecordCount

        Select Case Me.TxtModFlg.text

            Case "N"
                Msg = "  „ Õ›Ÿ »Ì«‰«  Â–Â «·⁄„·Ì…" & Chr(13)
                Msg = Msg + "Â·  —€» ›Ì ≈÷«›… »Ì«‰«  √Œ—Ì"

                If MsgBox(Msg, vbYesNo + vbQuestion + vbMsgBoxRight + vbMsgBoxRtlReading + vbDefaultButton2, App.Title) = vbYes Then
                    Cmd_Click (0)
                    Exit Sub
                End If

            Case "E"
                MsgBox " „ Õ›Ÿ Â–Â «· ⁄œÌ·« ", vbInformation + vbOKOnly + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
        End Select

        TxtModFlg.text = "R"
    End If

    Screen.MousePointer = vbDefault
    Exit Sub
ErrTrap:

    If BeginTrans = True Then
        BeginTrans = False
        Cn.RollbackTrans
    End If

    Screen.MousePointer = vbDefault

    If Err.Number = -2147217900 Then
        Msg = "·« Ì„ﬂ‰ Õ›Ÿ Â–Â «·»Ì«‰«  " & Chr(13)
        Msg = Msg + "·ﬁœ  „ «œŒ«· ﬁÌ„ €Ì— ’«·Õ… " & Chr(13)
        Msg = Msg + " √ﬂœ „‰ œﬁ… «·»Ì«‰«  Ê√⁄œ «·„Õ«Ê·…"
        MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
        Exit Sub
    End If

    Msg = "⁄›Ê«...ÕœÀ Œÿ√ „« √À‰«¡ Õ›Ÿ Â–Â «·»Ì«‰«  " & Chr(13)
    MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
End Sub

Public Sub Retrive(Optional Lngid As Long = 0)
    Dim RsDetails As New ADODB.Recordset
    Dim RsTemp As ADODB.Recordset
    Dim StrSQL As String
    Dim Num As Integer
    On Error GoTo ErrTrap
    Screen.MousePointer = vbArrowHourglass

    If rs.RecordCount < 1 Then
        XPTxtCurrent.Caption = 0
        XPTxtCount.Caption = 0
        Screen.MousePointer = vbDefault
        Exit Sub
    End If

    If rs.EOF Or rs.BOF Then
        Screen.MousePointer = vbDefault
        Exit Sub
    End If

    If Lngid <> 0 Then
        rs.find "Transaction_ID=" & Lngid, , adSearchForward, adBookmarkFirst

        If rs.BOF Or rs.EOF Then
            Screen.MousePointer = vbDefault
            Exit Sub
        End If
    End If

    XPTxtBillID.text = IIf(IsNull(rs("Transaction_ID").value), "", (rs("Transaction_ID").value))
    TxtTransSerial.text = IIf(IsNull(rs("Transaction_Serial").value), "", rs("Transaction_Serial").value)
    TXTOrDer_no.text = IIf(IsNull(rs("order_no").value), "", rs("order_no").value)

    XPDtbGard.value = IIf(IsNull(rs("Transaction_Date").value), "", (rs("Transaction_Date").value))
    Me.DCboStoreName.BoundText = IIf(IsNull(rs("StoreID").value), "", rs("StoreID").value)
    Me.DCboUserName.BoundText = IIf(IsNull(rs("UserID").value), "", rs("UserID").value)

    If IsNull(rs("ExtraOptions").value) Then
        XPOptShowType(1).value = True
    Else

        If rs("ExtraOptions").value = 0 Then
            XPOptShowType(1).value = True
            XPOptShowType(0).value = False
        Else
            XPOptShowType(0).value = True
            XPOptShowType(1).value = False
        End If
    
    End If

    FG.Rows = 2
    FG.Clear flexClearScrollable, flexClearEverything
    StrSQL = "SELECT Transaction_Details.*, TblItems.ItemCode, TblItems.HaveSerial " & "FROM TblItems INNER JOIN Transaction_Details ON TblItems.ItemID =" & "Transaction_Details.Item_ID WHere Transaction_ID=" & val(rs("Transaction_ID").value)

    RsDetails.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText

    If Not (RsDetails.EOF Or RsDetails.BOF) Then
        FG.Rows = RsDetails.RecordCount + 1
        FG.ColHidden(FG.ColIndex("Serial")) = False

        For Num = 1 To RsDetails.RecordCount

            With FG
                .TextMatrix(Num, .ColIndex("Index")) = Num
            
                .TextMatrix(Num, .ColIndex("order_no")) = IIf(IsNull(RsDetails("order_no")), "", RsDetails("order_no").value)
                .TextMatrix(Num, .ColIndex("ItemID")) = IIf(IsNull(RsDetails("Item_ID")), "", RsDetails("Item_ID").value)
                .TextMatrix(Num, .ColIndex("ItemCode")) = IIf(IsNull(RsDetails("ItemCode")), "", RsDetails("ItemCode").value)
                .TextMatrix(Num, .ColIndex("Name")) = IIf(IsNull(RsDetails("Item_ID")), "", Trim(RsDetails("Item_ID").value))
                .TextMatrix(Num, .ColIndex("Count")) = IIf(IsNull(RsDetails("Quantity")), "", (RsDetails("Quantity").value))
                .TextMatrix(Num, .ColIndex("CostPrice")) = IIf(IsNull(RsDetails("Price")), "", (RsDetails("Price").value))
                .TextMatrix(Num, .ColIndex("ItemValue")) = val(.TextMatrix(Num, .ColIndex("Count"))) * val(.TextMatrix(Num, .ColIndex("CostPrice")))

                If RsDetails("HaveSerial").value = True Then
                    .TextMatrix(Num, .ColIndex("Serial")) = IIf(IsNull(RsDetails("ItemSerial")), "", Trim(RsDetails("ItemSerial").value))
                Else
                    .TextMatrix(Num, .ColIndex("Serial")) = "·Ì” ·Â ”Ì—Ì«·"
                End If

            End With

            RsDetails.MoveNext
        Next Num
    
        Me.lbl(1).Caption = "ﬁÌ„… «·„Œ“Ê‰: " & FG.Aggregate(flexSTSum, FG.FixedRows, FG.ColIndex("ItemValue"), FG.Rows - 1, FG.ColIndex("ItemValue"))
        Me.Label200(3).Caption = FG.Aggregate(flexSTSum, FG.FixedRows, FG.ColIndex("Count"), FG.Rows - 1, FG.ColIndex("Count"))
    End If

    XPTxtCurrent.Caption = rs.AbsolutePosition
    XPTxtCount.Caption = rs.RecordCount
    Screen.MousePointer = vbDefault
    Exit Sub
ErrTrap:
    Screen.MousePointer = vbDefault
End Sub

Private Sub XPOptShowType_Click(Index As Integer)

    Select Case Index

        Case 0
            XPOptShowType(0).value = True
            XPOptShowType(1).value = False

        Case 1
            XPOptShowType(1).value = True
            XPOptShowType(0).value = False
    End Select

End Sub

Private Sub printing()
    On Error GoTo ErrTrap

    If XPTxtBillID.text <> "" Then
        Set VReport = New ClsGardReport

        If Me.XPOptShowType(0).value = True Then
            VReport.ShowGardData XPTxtBillID.text, 0
        Else
            VReport.ShowGardData XPTxtBillID.text, 1
        End If
    End If

    Exit Sub
ErrTrap:
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, _
                             UnloadMode As Integer)
    Dim IntResult As String
    Dim StrMSG As String
    On Error GoTo ErrTrap

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

Private Sub DoStockCount()
    Dim RsTemp As New ADODB.Recordset
    Dim StrSQL As String
    Dim RowNum As Integer
    Dim DblItemCost As Double
    Dim LngItemID As Long
    Dim DblQty As Double
    Dim LngOldItemID As Long
    Dim LngItemCount As Long
    Dim StrItemSerial As String

    On Error GoTo ErrTrap

    If Me.TxtModFlg.text <> "N" Then
        Exit Sub
    End If

    FG.Rows = 2
    FG.Clear flexClearScrollable, flexClearEverything

    If DCboStoreName.BoundText = "" Then
        Screen.MousePointer = vbDefault
        Exit Sub
    End If

    DoEvents
    Screen.MousePointer = vbArrowHourglass

    If XPOptShowType(1).value = True Then
        FG.ColHidden(FG.ColIndex("Serial")) = True

        If SystemOptions.SysDataBaseType = AccessDataBase Then
        
            StrSQL = "Select * From QryGARDShort where StoreID=" & DCboStoreName.BoundText
        ElseIf SystemOptions.SysDataBaseType = SQLServerDataBase Then
            '
            '      StrSQL = "Select * From dbo.QryGARDShort()QryGARDShort where StoreID=" & DCboStoreName.BoundText
            ' StrSQL = " Select *  from QryGardWithOrderNo where StoreID=" & DCboStoreName.BoundText
            StrSQL = "SELECT     SUM(dbo.Transaction_Details.Quantity * dbo.TransactionTypes.StockEffect) AS SUMQTY, dbo.Transaction_Details.Item_ID AS ItemID, dbo.Transactions.StoreID, "
            StrSQL = StrSQL & "  dbo.TblStore.StoreName, dbo.Transaction_Details.order_no, dbo.Transaction_Details.UnitId, dbo.TblItems.ItemCode, dbo.TblItems.ItemName,"
            StrSQL = StrSQL & " dbo.TblUnites.unitname"
            StrSQL = StrSQL & " FROM         dbo.Transactions INNER JOIN"
            StrSQL = StrSQL & " dbo.Transaction_Details ON dbo.Transactions.Transaction_ID = dbo.Transaction_Details.Transaction_ID INNER JOIN"
            StrSQL = StrSQL & " dbo.TransactionTypes ON dbo.Transactions.Transaction_Type = dbo.TransactionTypes.Transaction_Type INNER JOIN"
            StrSQL = StrSQL & " dbo.TblItems ON dbo.Transaction_Details.Item_ID = dbo.TblItems.ItemID INNER JOIN"
            StrSQL = StrSQL & " dbo.TblStore ON dbo.Transactions.StoreID = dbo.TblStore.StoreID LEFT OUTER JOIN"
            StrSQL = StrSQL & " dbo.TblUnites ON dbo.Transaction_Details.UnitId = dbo.TblUnites.UnitID LEFT OUTER JOIN"
            StrSQL = StrSQL & " dbo.TblItemsColors ON dbo.Transaction_Details.ColorID = dbo.TblItemsColors.ColorID"
            StrSQL = StrSQL & "  WHERE    (dbo.Transactions.StoreID = " & val(DCboStoreName.BoundText) & ") "
            BolBegain = True

            If Not IsNull(DTPFrom.value) Then
             
                If BolBegain = True Then
                    StrSQL = StrSQL + " and Transactions.Transaction_Date >=" & SQLDate(Me.DTPFrom.value, True) & ""
                Else
                    BolBegain = True
                    StrSQL = StrSQL + " Where Transactions.Transaction_Date >=" & SQLDate(Me.DTPFrom.value, True) & ""
                End If
            End If
        
            If Not IsNull(DTPTo.value) Then
 
                If BolBegain = True Then
                    StrSQL = StrSQL + " and Transactions.Transaction_Date <=" & SQLDate(Me.DTPTo.value, True) & ""
                Else
                    StrSQL = StrSQL + " Where Transactions.Transaction_Date <=" & SQLDate(Me.DTPTo.value, True) & ""
                End If
            End If
        
            If TXTOrDer_no.text <> "" Then
                StrSQL = StrSQL & " and  Transaction_Details.order_no='" & TXTOrDer_no.text & "'"
            End If
      
            StrSQL = StrSQL & " GROUP BY dbo.Transaction_Details.Item_ID, dbo.Transactions.StoreID, dbo.TblStore.StoreName, dbo.Transaction_Details.order_no, dbo.Transaction_Details.UnitId,"
            StrSQL = StrSQL & " dbo.TblItems.ItemCode , dbo.TblItems.itemname, dbo.TblUnites.unitname"
            StrSQL = StrSQL & " HAVING      (SUM(dbo.Transaction_Details.Quantity * dbo.TransactionTypes.StockEffect) <> 0)"
    
        End If
    
        If SystemOptions.SysDataBaseType = SQLServerDataBase Then
            StrSQL = StrSQL + " Order By   Transaction_Details.order_no,ItemID"
        End If
    
        RsTemp.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

        If RsTemp.EOF Or RsTemp.BOF Then
            Screen.MousePointer = vbDefault
            Exit Sub
        End If

        'Me.EleMain.Grid(gsRowHeight, 3) = 315
        PrgBr.Visible = True
        PrgBr.value = 0
        PrgBr.Max = RsTemp.RecordCount
        Me.lbl(0).Caption = "⁄œœ «·√’‰«›: " & RsTemp.RecordCount
        FG.Rows = RsTemp.RecordCount + 1

        For RowNum = 1 To RsTemp.RecordCount

            With FG
                .TextMatrix(RowNum, .ColIndex("Index")) = RowNum
                .TextMatrix(RowNum, .ColIndex("order_no")) = IIf(IsNull(RsTemp("order_no").value), "", RsTemp("order_no").value)
                .TextMatrix(RowNum, .ColIndex("ItemID")) = IIf(IsNull(RsTemp("ItemID").value), "", RsTemp("ItemID").value)
                .TextMatrix(RowNum, .ColIndex("ItemCode")) = IIf(IsNull(RsTemp("ItemCode").value), "", RsTemp("ItemCode").value)
                .TextMatrix(RowNum, .ColIndex("Name")) = IIf(IsNull(RsTemp("ItemID").value), "", RsTemp("ItemID").value)
                .TextMatrix(RowNum, .ColIndex("Count")) = IIf(IsNull(RsTemp("sumQTY").value), "", RsTemp("sumQTY").value)
                LngItemID = val(RsTemp("ItemID").value)
                DblQty = val(.TextMatrix(RowNum, .ColIndex("Count")))

                If val(.TextMatrix(RowNum, .ColIndex("ItemCode"))) = 2002 Then
                    'Stop
                End If

                DblItemCost = GetCostItemPrice(LngItemID, 2)
                .TextMatrix(RowNum, .ColIndex("CostPrice")) = DblItemCost
                .TextMatrix(RowNum, .ColIndex("ItemValue")) = DblItemCost * DblQty
                PrgBr.value = RowNum

                DoEvents
            End With

            RsTemp.MoveNext
        Next RowNum

        RsTemp.Close
        PrgBr.value = 0
        PrgBr.Visible = False
        Me.lbl(1).Caption = "ﬁÌ„… «·„Œ“Ê‰: " & FG.Aggregate(flexSTSum, FG.FixedRows, FG.ColIndex("ItemValue"), FG.Rows - 1, FG.ColIndex("ItemValue"))
    
        'Me.EleMain.Grid(gsRowHeight, 3) = 0
        DoEvents
    End If

    If XPOptShowType(0).value = True Then
    
        FG.ColHidden(FG.ColIndex("Serial")) = False

        If SystemOptions.SysDataBaseType = AccessDataBase Then
            StrSQL = "select * From QryGardComplete where StoreID=" & DCboStoreName.BoundText
        Else
            StrSQL = "select * From dbo.QryGardComplete(0)QryGardComplete where StoreID=" & DCboStoreName.BoundText
        End If

        If SystemOptions.SysDataBaseType = SQLServerDataBase Then
            StrSQL = StrSQL + " Order By GroupID,ItemID"
        End If

        RsTemp.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

        If RsTemp.EOF Or RsTemp.BOF Then
            Screen.MousePointer = vbDefault
            Exit Sub
        End If

        PrgBr.Visible = True
        PrgBr.value = 0
        PrgBr.Max = RsTemp.RecordCount
    
        FG.Rows = RsTemp.RecordCount + 1

        For RowNum = 1 To RsTemp.RecordCount

            With FG
                .TextMatrix(RowNum, .ColIndex("Index")) = RowNum
                .TextMatrix(RowNum, .ColIndex("ItemID")) = IIf(IsNull(RsTemp("ItemID").value), "", RsTemp("ItemID").value)
                .TextMatrix(RowNum, .ColIndex("ItemCode")) = IIf(IsNull(RsTemp("ItemCode").value), "", RsTemp("ItemCode").value)
                .TextMatrix(RowNum, .ColIndex("Name")) = IIf(IsNull(RsTemp("ItemID").value), "", RsTemp("ItemID").value)
                .TextMatrix(RowNum, .ColIndex("Count")) = IIf(IsNull(RsTemp("QTY").value), "", RsTemp("QTY").value)
                LngItemID = val(RsTemp("ItemID").value)
            
                If LngOldItemID <> LngItemID Then
                    LngOldItemID = LngItemID
                    LngItemCount = LngItemCount + 1
                End If

                DblQty = val(.TextMatrix(RowNum, .ColIndex("Count")))

                If RsTemp("HaveSerial").value = True Then
                    .TextMatrix(RowNum, .ColIndex("Serial")) = IIf(IsNull(RsTemp("ItemSerial")), "", Trim(RsTemp("ItemSerial").value))
                    StrItemSerial = Trim(.TextMatrix(RowNum, .ColIndex("Serial")))

                    If StrItemSerial <> "" Then
                        DblItemCost = GetCostItemPrice(LngItemID, 1, StrItemSerial)
                        .Rowdata(RowNum) = "NoSerial"
                    Else
                        DblItemCost = GetCostItemPrice(LngItemID, 2)
                    End If

                Else
                    .TextMatrix(RowNum, .ColIndex("Serial")) = "·Ì” ·Â ”Ì—Ì«·"
                    DblItemCost = GetCostItemPrice(LngItemID, 2)
                End If

                .TextMatrix(RowNum, .ColIndex("CostPrice")) = DblItemCost
                .TextMatrix(RowNum, .ColIndex("ItemValue")) = DblItemCost * DblQty
                PrgBr.value = RowNum

                DoEvents
            End With

            RsTemp.MoveNext
        Next RowNum

        RsTemp.Close
        Me.lbl(0).Caption = "⁄œœ «·√’‰«›: " & LngItemCount
        PrgBr.value = 0
        PrgBr.Visible = False
        Me.lbl(1).Caption = "ﬁÌ„… «·„Œ“Ê‰: " & FG.Aggregate(flexSTSum, FG.FixedRows, FG.ColIndex("ItemValue"), FG.Rows - 1, FG.ColIndex("ItemValue"))
    End If

    Me.Label200(3).Caption = FG.Aggregate(flexSTSum, FG.FixedRows, FG.ColIndex("Count"), FG.Rows - 1, FG.ColIndex("Count"))
    FG.AutoSize 0, FG.Cols - 1, False
    Screen.MousePointer = vbDefault
    Exit Sub
ErrTrap:
    Screen.MousePointer = vbDefault
End Sub
