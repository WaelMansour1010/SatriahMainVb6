VERSION 5.00
Object = "{C115893A-A3BF-43AF-B28D-69DB846077F3}#1.0#0"; "vsflex8u.ocx"
Object = "{0AFE7BE0-11B7-4A3E-978D-D4501E9A57FE}#1.0#0"; "c1sizer.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{FE5DCFAD-BC1D-11D2-94CF-004005455FAA}#1.4#0"; "ImpulseButton.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form FrmTemplate 
   Caption         =   "⁄—Ê÷ «·«’‰«›"
   ClientHeight    =   7005
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   8580
   HelpContextID   =   350
   Icon            =   "FrmTemplate.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   RightToLeft     =   -1  'True
   ScaleHeight     =   7005
   ScaleWidth      =   8580
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
      Height          =   7005
      Left            =   0
      TabIndex        =   15
      TabStop         =   0   'False
      Top             =   0
      Width           =   8580
      _cx             =   15134
      _cy             =   12356
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
      GridRows        =   6
      GridCols        =   4
      Frame           =   3
      FrameStyle      =   0
      FrameWidth      =   1
      FrameColor      =   -2147483628
      FrameShadow     =   -2147483632
      FloodStyle      =   1
      _GridInfo       =   $"FrmTemplate.frx":038A
      AccessibleName  =   ""
      AccessibleDescription=   ""
      AccessibleValue =   ""
      AccessibleRole  =   9
      Begin C1SizerLibCtl.C1Elastic Ele 
         Height          =   435
         Index           =   3
         Left            =   15
         TabIndex        =   33
         TabStop         =   0   'False
         Top             =   6000
         Width           =   8550
         _cx             =   15081
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
            Left            =   6330
            Locked          =   -1  'True
            RightToLeft     =   -1  'True
            TabIndex        =   34
            TabStop         =   0   'False
            Top             =   30
            Width           =   960
         End
         Begin MSDataListLib.DataCombo DCboUserName 
            Height          =   315
            Left            =   4020
            TabIndex        =   35
            Top             =   45
            Width           =   1230
            _ExtentX        =   2170
            _ExtentY        =   556
            _Version        =   393216
            Enabled         =   0   'False
            Text            =   ""
            RightToLeft     =   -1  'True
         End
         Begin VB.Label Lbl 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Õ—— »Ê«”ÿ… : "
            Height          =   315
            Index           =   1
            Left            =   5280
            RightToLeft     =   -1  'True
            TabIndex        =   41
            Top             =   75
            Width           =   1005
         End
         Begin VB.Label XPTxtCount 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Height          =   240
            Left            =   75
            RightToLeft     =   -1  'True
            TabIndex        =   40
            Top             =   135
            Width           =   585
         End
         Begin VB.Label XPTxtCurrent 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Height          =   270
            Left            =   1890
            RightToLeft     =   -1  'True
            TabIndex        =   39
            Top             =   105
            Width           =   690
         End
         Begin VB.Label Lbl 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "⁄œœ «·”Ã·« :"
            Height          =   240
            Index           =   2
            Left            =   810
            RightToLeft     =   -1  'True
            TabIndex        =   38
            Top             =   120
            Width           =   1005
         End
         Begin VB.Label Lbl 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "«·”Ã· «·Õ«·Ì:"
            Height          =   240
            Index           =   0
            Left            =   2760
            RightToLeft     =   -1  'True
            TabIndex        =   37
            Top             =   120
            Width           =   1170
         End
         Begin VB.Label Lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "≈Ã„«·Ì «·›« Ê—…"
            Height          =   255
            Index           =   3
            Left            =   7305
            RightToLeft     =   -1  'True
            TabIndex        =   36
            Top             =   75
            Width           =   1230
         End
      End
      Begin C1SizerLibCtl.C1Elastic Ele 
         Height          =   540
         Index           =   1
         Left            =   15
         TabIndex        =   23
         TabStop         =   0   'False
         Top             =   6450
         Width           =   8550
         _cx             =   15081
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
            Left            =   7635
            TabIndex        =   24
            Top             =   90
            Width           =   855
            _ExtentX        =   1508
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
            Left            =   6675
            TabIndex        =   25
            Top             =   90
            Width           =   855
            _ExtentX        =   1508
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
            Left            =   5700
            TabIndex        =   26
            Top             =   90
            Width           =   885
            _ExtentX        =   1561
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
            Left            =   4770
            TabIndex        =   27
            Top             =   90
            Width           =   840
            _ExtentX        =   1482
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
            Left            =   3825
            TabIndex        =   28
            Top             =   90
            Width           =   855
            _ExtentX        =   1508
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
            Left            =   2880
            TabIndex        =   29
            Top             =   90
            Width           =   855
            _ExtentX        =   1508
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
            TabIndex        =   30
            Top             =   90
            Width           =   840
            _ExtentX        =   1482
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
            Left            =   1890
            TabIndex        =   31
            Top             =   90
            Width           =   885
            _ExtentX        =   1561
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
            Left            =   960
            TabIndex        =   32
            Top             =   90
            Width           =   855
            _ExtentX        =   1508
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
         Height          =   4290
         Index           =   5
         Left            =   15
         TabIndex        =   16
         TabStop         =   0   'False
         Top             =   1695
         Width           =   8550
         _cx             =   15081
         _cy             =   7567
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
         GridRows        =   2
         GridCols        =   4
         Frame           =   3
         FrameStyle      =   0
         FrameWidth      =   1
         FrameColor      =   -2147483628
         FrameShadow     =   -2147483632
         FloodStyle      =   1
         _GridInfo       =   $"FrmTemplate.frx":0421
         AccessibleName  =   ""
         AccessibleDescription=   ""
         AccessibleValue =   ""
         AccessibleRole  =   9
         Begin C1SizerLibCtl.C1Elastic Ele 
            Height          =   690
            Index           =   2
            Left            =   30
            TabIndex        =   46
            TabStop         =   0   'False
            Top             =   3570
            Width           =   8490
            _cx             =   14975
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
               Left            =   555
               MaxLength       =   10
               RightToLeft     =   -1  'True
               TabIndex        =   12
               Top             =   300
               Width           =   1230
            End
            Begin VB.TextBox TxtQuantity 
               Alignment       =   1  'Right Justify
               Enabled         =   0   'False
               Height          =   300
               Left            =   1845
               MaxLength       =   10
               RightToLeft     =   -1  'True
               TabIndex        =   11
               Top             =   300
               Width           =   1230
            End
            Begin VB.ComboBox CboItemCase 
               Height          =   315
               Left            =   3135
               RightToLeft     =   -1  'True
               Style           =   2  'Dropdown List
               TabIndex        =   10
               Top             =   300
               Width           =   1110
            End
            Begin MSDataListLib.DataCombo DCboItemsName 
               Height          =   315
               Left            =   4305
               TabIndex        =   9
               Top             =   300
               Width           =   2040
               _ExtentX        =   3598
               _ExtentY        =   556
               _Version        =   393216
               Text            =   ""
               RightToLeft     =   -1  'True
            End
            Begin MSDataListLib.DataCombo DCboItemsCode 
               Height          =   315
               Left            =   6405
               TabIndex        =   8
               Top             =   300
               Width           =   1980
               _ExtentX        =   3493
               _ExtentY        =   556
               _Version        =   393216
               Text            =   ""
               RightToLeft     =   -1  'True
            End
            Begin ImpulseButton.ISButton CmdAdd 
               Height          =   375
               Left            =   60
               TabIndex        =   13
               Top             =   263
               Width           =   390
               _ExtentX        =   688
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
               ButtonImage     =   "FrmTemplate.frx":0489
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
            Begin VB.Label Lbl 
               Alignment       =   2  'Center
               BackColor       =   &H00E2E9E9&
               Caption         =   "«·”⁄—"
               Height          =   255
               Index           =   26
               Left            =   615
               RightToLeft     =   -1  'True
               TabIndex        =   51
               Top             =   0
               Width           =   1170
            End
            Begin VB.Label Lbl 
               Alignment       =   2  'Center
               BackColor       =   &H00E2E9E9&
               Caption         =   "«·ﬂ„Ì…"
               Height          =   255
               Index           =   27
               Left            =   1935
               RightToLeft     =   -1  'True
               TabIndex        =   50
               Top             =   0
               Width           =   1080
            End
            Begin VB.Label Lbl 
               Alignment       =   2  'Center
               BackColor       =   &H00E2E9E9&
               Caption         =   "Õ«·… «·’‰›"
               Height          =   255
               Index           =   29
               Left            =   3225
               RightToLeft     =   -1  'True
               TabIndex        =   49
               Top             =   0
               Width           =   990
            End
            Begin VB.Label Lbl 
               Alignment       =   2  'Center
               BackColor       =   &H00E2E9E9&
               Caption         =   "≈”„ «·’‰›"
               Height          =   255
               Index           =   30
               Left            =   4455
               RightToLeft     =   -1  'True
               TabIndex        =   48
               Top             =   0
               Width           =   1890
            End
            Begin VB.Label Lbl 
               Alignment       =   2  'Center
               BackColor       =   &H00E2E9E9&
               Caption         =   "ﬂÊœ «·’‰›"
               Height          =   255
               Index           =   31
               Left            =   6525
               RightToLeft     =   -1  'True
               TabIndex        =   47
               Top             =   0
               Width           =   1860
            End
         End
         Begin VSFlex8UCtl.VSFlexGrid FG 
            Height          =   3525
            Left            =   495
            TabIndex        =   7
            Top             =   30
            Width           =   8025
            _cx             =   14155
            _cy             =   6218
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
            Cols            =   12
            FixedRows       =   1
            FixedCols       =   1
            RowHeightMin    =   300
            RowHeightMax    =   0
            ColWidthMin     =   0
            ColWidthMax     =   0
            ExtendLastCol   =   -1  'True
            FormatString    =   $"FrmTemplate.frx":0823
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
         Begin C1SizerLibCtl.C1Elastic C1Elastic2 
            Height          =   3525
            Left            =   30
            TabIndex        =   42
            TabStop         =   0   'False
            Top             =   30
            Width           =   450
            _cx             =   794
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
            Begin ImpulseButton.ISButton XPBtnAdd 
               Height          =   315
               Left            =   0
               TabIndex        =   43
               TabStop         =   0   'False
               Top             =   210
               Width           =   390
               _ExtentX        =   688
               _ExtentY        =   556
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
               ButtonImage     =   "FrmTemplate.frx":09E6
               ColorButton     =   14871017
               ColorHighlight  =   16777215
               ColorShadow     =   -2147483631
               ColorOutline    =   -2147483631
               DrawFocusRectangle=   0   'False
               LowerToggledContent=   0   'False
            End
            Begin ImpulseButton.ISButton XPBtnRemove 
               Height          =   315
               Left            =   0
               TabIndex        =   44
               TabStop         =   0   'False
               Top             =   750
               Width           =   390
               _ExtentX        =   688
               _ExtentY        =   556
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
               ButtonImage     =   "FrmTemplate.frx":0D80
               ColorButton     =   14871017
               ColorHighlight  =   16777215
               ColorShadow     =   -2147483631
               ColorOutline    =   -2147483631
               DrawFocusRectangle=   0   'False
               LowerToggledContent=   0   'False
            End
            Begin ImpulseButton.ISButton XPFillData 
               Height          =   315
               Left            =   0
               TabIndex        =   45
               TabStop         =   0   'False
               Top             =   1290
               Width           =   390
               _ExtentX        =   688
               _ExtentY        =   556
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
               ButtonImage     =   "FrmTemplate.frx":111A
               ColorButton     =   14871017
               ColorHighlight  =   16777215
               ColorShadow     =   -2147483631
               ColorOutline    =   -2147483631
               DrawFocusRectangle=   0   'False
               LowerToggledContent=   0   'False
            End
         End
      End
      Begin C1SizerLibCtl.C1Elastic Ele 
         Height          =   1035
         Index           =   0
         Left            =   15
         TabIndex        =   17
         TabStop         =   0   'False
         Top             =   645
         Width           =   8550
         _cx             =   15081
         _cy             =   1826
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
         Begin VB.Frame Fram 
            BackColor       =   &H00E2E9E9&
            Caption         =   "„œ… «·⁄—÷"
            ForeColor       =   &H00FF0000&
            Height          =   600
            Left            =   2640
            RightToLeft     =   -1  'True
            TabIndex        =   14
            Top             =   435
            Width           =   5835
            Begin VB.OptionButton XPOpt 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "”«—Ì Õ Ï"
               Height          =   210
               Index           =   0
               Left            =   2130
               RightToLeft     =   -1  'True
               TabIndex        =   4
               Top             =   255
               Width           =   1245
            End
            Begin VB.OptionButton XPOpt 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "„› ÊÕ Õ Ï Ì „ Õ–›Â"
               Height          =   315
               Index           =   1
               Left            =   3510
               RightToLeft     =   -1  'True
               TabIndex        =   3
               Top             =   210
               Value           =   -1  'True
               Width           =   1725
            End
            Begin MSComCtl2.DTPicker DtpExpire 
               Height          =   345
               Left            =   90
               TabIndex        =   5
               Top             =   165
               Width           =   1395
               _ExtentX        =   2461
               _ExtentY        =   609
               _Version        =   393216
               Enabled         =   0   'False
               Format          =   96468993
               CurrentDate     =   38784
            End
         End
         Begin VB.TextBox TxtCode 
            Alignment       =   1  'Right Justify
            Height          =   315
            Left            =   3075
            MaxLength       =   50
            RightToLeft     =   -1  'True
            TabIndex        =   1
            Top             =   90
            Width           =   1620
         End
         Begin VB.TextBox XPTxtBillID 
            Alignment       =   1  'Right Justify
            Height          =   315
            Left            =   5760
            Locked          =   -1  'True
            MaxLength       =   10
            RightToLeft     =   -1  'True
            TabIndex        =   0
            Top             =   90
            Width           =   1650
         End
         Begin VB.TextBox TxtFillData 
            Alignment       =   1  'Right Justify
            Height          =   195
            Left            =   3330
            RightToLeft     =   -1  'True
            TabIndex        =   19
            Top             =   -195
            Visible         =   0   'False
            Width           =   570
         End
         Begin VB.TextBox TxtModFlg 
            Alignment       =   1  'Right Justify
            Height          =   195
            Left            =   4080
            RightToLeft     =   -1  'True
            TabIndex        =   18
            Top             =   -195
            Visible         =   0   'False
            Width           =   705
         End
         Begin MSComCtl2.DTPicker XPDtbBill 
            Height          =   315
            Left            =   150
            TabIndex        =   2
            Top             =   90
            Width           =   1500
            _ExtentX        =   2646
            _ExtentY        =   556
            _Version        =   393216
            Format          =   96468993
            CurrentDate     =   38784
         End
         Begin ImpulseButton.ISButton CmdConvert 
            Height          =   390
            Left            =   -30
            TabIndex        =   6
            Top             =   555
            Width           =   2595
            _ExtentX        =   4577
            _ExtentY        =   688
            ButtonStyle     =   1
            ButtonPositionImage=   1
            Caption         =   " ÕÊÌ· ≈·Ï ›« Ê—… »Ì⁄"
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
         Begin VB.Label Lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "«”„ «·⁄—÷"
            Height          =   195
            Index           =   4
            Left            =   4755
            RightToLeft     =   -1  'True
            TabIndex        =   22
            Top             =   150
            Width           =   960
         End
         Begin VB.Label Lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   " «—ÌŒ «·⁄—÷"
            Height          =   225
            Index           =   6
            Left            =   1665
            RightToLeft     =   -1  'True
            TabIndex        =   21
            Top             =   135
            Width           =   1140
         End
         Begin VB.Label Lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "ﬂÊœ «·⁄—÷"
            Height          =   210
            Index           =   5
            Left            =   7470
            RightToLeft     =   -1  'True
            TabIndex        =   20
            Top             =   135
            Width           =   930
         End
      End
      Begin C1SizerLibCtl.C1Elastic C1Elastic6 
         Height          =   615
         Left            =   15
         TabIndex        =   52
         TabStop         =   0   'False
         Top             =   15
         Width           =   8550
         _cx             =   15081
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
         Caption         =   "⁄—Ê÷ «·«’‰«›"
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
         Begin ImpulseButton.ISButton XPBtnMove 
            Height          =   345
            Index           =   0
            Left            =   1155
            TabIndex        =   53
            Top             =   120
            Width           =   480
            _ExtentX        =   847
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
            ButtonImage     =   "FrmTemplate.frx":14B4
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
            Left            =   630
            TabIndex        =   54
            Top             =   120
            Width           =   480
            _ExtentX        =   847
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
            ButtonImage     =   "FrmTemplate.frx":184E
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
            TabIndex        =   55
            Top             =   120
            Width           =   465
            _ExtentX        =   820
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
            ButtonImage     =   "FrmTemplate.frx":1BE8
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
            Left            =   105
            TabIndex        =   56
            Top             =   120
            Width           =   480
            _ExtentX        =   847
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
            ButtonImage     =   "FrmTemplate.frx":1F82
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
   End
End
Attribute VB_Name = "FrmTemplate"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rs As ADODB.Recordset
Dim TTP As clstooltip
Dim NewGrid As New ClsGrid
Dim SaleReport As ClsSaleReport
Dim cSearchDcbo(3)  As clsDCboSearch

Private Sub C1Elastic6_DblClick()
    On Error GoTo ErrTrap

    If Me.WindowState = vbNormal Then
        Me.WindowState = vbMaximized
    Else
        Me.WindowState = vbNormal
    End If

    Exit Sub
ErrTrap:
End Sub

Private Sub Cmd_Click(Index As Integer)

    'On Error GoTo ErrTrap
    Select Case Index

        Case 0

            If DoPremis(Do_New, Me.name, True) = False Then
                Exit Sub
            End If

            clear_all Me
            TxtModFlg.text = "N"
            XPTxtBillID.text = CStr(new_id("Templates", "TemplateID", "", True))
            NewGrid.GridDefaultValue 1
            Me.DCboUserName.BoundText = user_id
            XPOpt(1).value = True
            Fg.SetFocus
            Fg.Col = Fg.ColIndex("Code")
            Fg.Row = Fg.Rows - 1

        Case 1

            If DoPremis(Do_Edit, Me.name, True) = False Then
                Exit Sub
            End If

            TxtModFlg.text = "E"
            Me.DCboUserName.BoundText = user_id

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

            Load FrmBuySearch

            With FrmBuySearch
                .DealingForm = Template
                .Caption = "«·»ÕÀ ⁄‰ «·⁄—Ê÷ «·Ã«Â“…"
            
                .XPChkSearchType.Visible = False
                .TxtVal.Visible = True
                .XPLbl(2).Visible = True
                .XPLbl(1).Visible = False
                .XPLbl(0).Visible = False
                .XPLbl(3).Visible = True
                .XPLbl(4).Visible = True
                .Fg.TextMatrix(0, .Fg.ColIndex("Transaction_ID")) = "ﬂÊœ «·⁄—÷"
                .Fg.TextMatrix(0, .Fg.ColIndex("BillDate")) = "«”„ «·⁄—÷"
                .Fg.TextMatrix(0, .Fg.ColIndex("ClientNmae")) = " «—ÌŒ «·⁄—÷"
                .Fg.TextMatrix(0, .Fg.ColIndex("StorName")) = "ﬁÌ„… «·⁄—÷"
                .show vbModal
            End With

        Case 7

            If DoPremis(Do_Print, Me.name, True) = False Then
                Exit Sub
            End If

            PrintReport

        Case 6
            Unload Me
    End Select

    Exit Sub
ErrTrap:
End Sub

Private Sub CmdConvert_Click()
    On Error GoTo ErrTrap
    Dim Msg As String
    Screen.MousePointer = vbArrowHourglass
    Dim RowNum As Integer

    '„‰⁄  ÕÊÌ· «·⁄—÷ ≈·Ï ›« Ê—… »Ì⁄ »⁄œ «‰ Â«¡ „œ… «·⁄—÷
    If XPOpt(0).value = True Then
        If SQLDate(Date) > SQLDate(DtpExpire.value) Then
            Msg = "⁄›Ê« : ·ﬁœ «‰ Â  „œ… Â–« «·⁄—÷"
            MsgBox Msg, vbOKOnly + vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
            Exit Sub
        End If

        If XPOpt(0).value = True Then
            If SQLDate(Date) = SQLDate(DtpExpire.value) Then
                Msg = " ‰»ÌÂ : ”Ì‰ ÂÌ «·⁄„· »Â–« «·⁄—÷ «·ÌÊ„"
                MsgBox Msg, vbOKOnly + vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
            End If
        End If
    End If

    With frmsalebill
        .Convert
        .XPDtbBill.value = XPDtbBill.value

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
            .Fg.TextMatrix(.Fg.Rows - 1, .Fg.ColIndex("DiscountVal")) = IIf(Fg.TextMatrix(RowNum, Fg.ColIndex("DiscountVal")) = "", "", Fg.TextMatrix(RowNum, Fg.ColIndex("DiscountVal")))
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

Private Sub TxtFillData_Change()

    If TxtFillData.text = "F" Then
        NewGrid.Calculate 1
    End If

End Sub

Private Sub XPBtnAdd_Click()

    If Fg.TextMatrix(Fg.Rows - 1, Fg.ColIndex("Code")) <> "" Then
        Fg.Rows = Fg.Rows + 1
        NewGrid.GridDefaultValue Fg.Rows - 1
        Fg.Row = Fg.Rows - 1
        Fg.Col = Fg.ColIndex("Code")
        Fg.ShowCell Fg.Rows - 1, Fg.ColIndex("Code")
        Fg.SetFocus
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
            XPBtnAdd_Click
        End If
    End If

    If KeyCode = vbKeyF3 Then
        If TxtModFlg.text = "N" Or TxtModFlg.text = "E" Then
            XPBtnRemove_Click
        End If
    End If

    If Shift = 2 Then
        If KeyCode = vbKeySpace Then
            If TxtModFlg.text = "N" Or TxtModFlg.text = "E" Then
                XPFillData_Click
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
    Dim Dcombos As ClsDataCombos
    On Error GoTo ErrTrap
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
    Set NewGrid.Grid = Fg
    NewGrid.GridTrans = Template
    Set NewGrid.TxtModFlag = TxtModFlg
    Set NewGrid.TXTTotal = XPTxtSum
    Set NewGrid.TxtFillData = TxtFillData

    ' ⁄»∆… »Ì«‰«  «·√’‰«›
    Set NewGrid.DcboItemName = DCboItemsName
    Set NewGrid.DCboItemCode = DCboItemsCode
    Set NewGrid.CboItemCase = CboItemCase
    Set NewGrid.CmdAddData = CmdAdd
    'Set NewGrid.TxtSerial = TxtSerial
    Set NewGrid.TxtQuantity = TxtQuantity
    Set NewGrid.TxtPrice = TxtPrice
    NewGrid.FillGrid
    Resize_Form Me, TransactionSize
    Fg.WallPaper = BGround.Picture
    AddTip
    SetDtpickerDate Me.XPDtbBill
    SetDtpickerDate Me.DtpExpire

    Set Dcombos = New ClsDataCombos
    Dcombos.GetUsers Me.DCboUserName

    StrSQL = "SELECT * FROM Templates"
    Set rs = New ADODB.Recordset
    rs.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText

    If SystemOptions.UserInterface = EnglishInterface Then
        SetInterface Me
        ChangeLang
    End If

    XPBtnMove_Click 2

    Me.TxtModFlg.text = "R"

    If OPEN_NEW_SCREEN = True Then
        Cmd_Click (0)
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

    Me.Caption = "Templates "
    C1Elastic6.Caption = Me.Caption
    lbl(5).Caption = "ID"
    lbl(4).Caption = "Name"
    lbl(6).Caption = "Date"
    Fram.Caption = "Valid until"
    XPOpt(1).Caption = "Delete"
    XPOpt(0).Caption = "Date "
    CmdConvert.Caption = "Convert to sales invoice"

    lbl(31).Caption = "Item Code"
    lbl(30).Caption = "Item Name"
    lbl(29).Caption = "Item Case"
 
    lbl(27).Caption = "Quantity"
    lbl(26).Caption = "Price"

    Me.Cmd(0).Caption = "New"
    Me.Cmd(1).Caption = "Edit"
    Me.Cmd(2).Caption = "Save"
    Me.Cmd(3).Caption = "Undo"
    Me.Cmd(4).Caption = "Delete"
    Me.Cmd(5).Caption = "Search"
    Me.Cmd(6).Caption = "Exit"
    Me.Cmd(7).Caption = "Print"
    Me.CmdHelp.Caption = "Help"
 
    lbl(3).Caption = "Total"
    lbl(1).Caption = "User"
    lbl(0).Caption = "Curr rec."
    lbl(2).Caption = "rec. count"

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
    Set SaleReport = Nothing
    Exit Sub
ErrTrap:
End Sub

Private Sub TxtModFlg_Change()
    On Error GoTo ErrTrap

    Select Case Me.TxtModFlg.text

        Case "R"

            If SystemOptions.UserInterface = ArabicInterface Then
                Me.Caption = "⁄—Ê÷ Ã«Â“…"
            Else
                Me.Caption = "Price Templates"
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
            Txtcode.locked = True
            XPBtnAdd.Enabled = False
            XPBtnRemove.Enabled = False
            Me.XPDtbBill.Enabled = False
            Fg.Editable = flexEDNone
            XPBtnAdd.Enabled = False
            XPBtnRemove.Enabled = False
            XPFillData.Enabled = False
            CmdConvert.Enabled = True
            Fram.Enabled = False

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
            End If

            Ele(2).Enabled = False

        Case "N"
         
            If SystemOptions.UserInterface = ArabicInterface Then
                Me.Caption = "⁄—Ê÷ Ã«Â“…( ÃœÌœ )"
            Else
                Me.Caption = "Price Templates (New)"
            End If
    
            Me.Cmd(2).Enabled = True
            Me.Cmd(3).Enabled = True
        
            Me.Cmd(0).Enabled = False
            Me.Cmd(1).Enabled = False
            Me.Cmd(4).Enabled = False
            Me.Cmd(5).Enabled = False
            Me.Cmd(7).Enabled = False
        
            '       Me.XPBtnMove(0).Enabled = False
            '       Me.XPBtnMove(1).Enabled = False
            '       Me.XPBtnMove(2).Enabled = False
            '       Me.XPBtnMove(3).Enabled = False
            Txtcode.locked = False
            Fg.Enabled = True
            Fg.Rows = 2
            Me.XPDtbBill.Enabled = True
            XPDtbBill.value = Date
            DtpExpire.value = Date + 1
            Fg.Editable = flexEDKbdMouse
            XPBtnAdd.Enabled = True
            XPBtnRemove.Enabled = True
            XPFillData.Enabled = True
            CmdConvert.Enabled = False
            Fram.Enabled = True
            Ele(2).Enabled = True
            CboItemCase.ListIndex = 0

        Case "E"

            If SystemOptions.UserInterface = ArabicInterface Then
                Me.Caption = "⁄—Ê÷ Ã«Â“…(  ⁄œÌ· )"
            Else
                Me.Caption = "Price Templates (edit)"
            End If
        
            Me.Cmd(2).Enabled = True
            Me.Cmd(3).Enabled = True
        
            Me.Cmd(0).Enabled = False
            Me.Cmd(1).Enabled = False
            Me.Cmd(4).Enabled = False
            Me.Cmd(5).Enabled = False
            Me.Cmd(7).Enabled = False
            Txtcode.locked = False
            Me.XPBtnMove(0).Enabled = False
            Me.XPBtnMove(1).Enabled = False
            Me.XPBtnMove(2).Enabled = False
            Me.XPBtnMove(3).Enabled = False
            XPBtnRemove.Enabled = True
            Fg.Enabled = True
            Me.XPDtbBill.Enabled = True
            Fg.Editable = flexEDKbdMouse
            XPBtnAdd.Enabled = True
            XPBtnRemove.Enabled = True
            XPFillData.Enabled = True
            CmdConvert.Enabled = False
            Fram.Enabled = True
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
        rs.find "TemplateID=" & Lngid, , adSearchForward, adBookmarkFirst

        If rs.BOF Or rs.EOF Then
            Exit Sub
        End If
    End If

    TxtFillData.text = "T"
    Screen.MousePointer = vbArrowHourglass
    XPTxtBillID.text = IIf(IsNull(rs("TemplateID").value), "", val(rs("TemplateID").value))
    Me.Txtcode.text = IIf(IsNull(rs("TemplateName").value), "", Trim(rs("TemplateName").value))
    XPDtbBill.value = IIf(IsNull(rs("Date").value), "", (rs("Date").value))
    Me.DCboUserName.BoundText = IIf(IsNull(rs("User_ID").value), "", rs("User_ID").value)

    If Not IsNull(rs("TemplateType").value) Then
        If rs("TemplateType").value = 1 Then
            XPOpt(1).value = True
        ElseIf rs("TemplateType").value = 0 Then
            XPOpt(0).value = True
            DtpExpire.value = IIf(IsNull(rs("TemplateTime").value), Date, rs("TemplateTime").value)
        End If
    End If

    Fg.Rows = 2
    Fg.Clear flexClearScrollable, flexClearEverything
    StrSQL = "SELECT Template_Details.TemplateID, Template_Details.ItemID, " & "TblItems.HaveSerial, Template_Details.Quantity, Template_Details.Price, " & "Template_Details.ItemDiscountType, Template_Details.ItemDiscount, " & "Template_Details.ItemCase FROM TblItems INNER JOIN Template_Details ON " & "TblItems.ItemID = Template_Details.ItemID"
    StrSQL = StrSQL + " where TemplateID=" & val(rs("TemplateID").value)
    RsDetails.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText
    XPTxtSum.text = ""

    If Not (RsDetails.EOF Or RsDetails.BOF) Then
        Fg.Rows = RsDetails.RecordCount + 1

        For Num = 1 To RsDetails.RecordCount
            Fg.TextMatrix(Num, Fg.ColIndex("Code")) = IIf(IsNull(RsDetails("ItemID")), "", (RsDetails("ItemID").value))
            Fg.TextMatrix(Num, Fg.ColIndex("Name")) = IIf(IsNull(RsDetails("ItemID")), "", Trim(RsDetails("ItemID").value))
            Fg.TextMatrix(Num, Fg.ColIndex("Count")) = IIf(IsNull(RsDetails("Quantity")), "", (RsDetails("Quantity").value))
            Fg.TextMatrix(Num, Fg.ColIndex("Price")) = IIf(IsNull(RsDetails("Price")), "", (RsDetails("Price").value))
            Fg.TextMatrix(Num, Fg.ColIndex("ItemCase")) = IIf(IsNull(RsDetails("ItemCase")), "", (RsDetails("ItemCase").value))
            Fg.TextMatrix(Num, Fg.ColIndex("DiscountType")) = IIf(IsNull(RsDetails("ItemDiscountType")), "", (RsDetails("ItemDiscountType").value))
            Fg.TextMatrix(Num, Fg.ColIndex("DiscountVal")) = IIf(IsNull(RsDetails("ItemDiscount")), "", (RsDetails("ItemDiscount").value))

            If RsDetails("HaveSerial") = True Then
                Fg.TextMatrix(Num, Fg.ColIndex("HaveSerial")) = True
            End If

            RsDetails.MoveNext

            If Fg.Rows > 10 Then
                If Num = 8 Then Fg.Refresh
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
    On Error GoTo ErrTrap
    Dim Msg  As String

    If XPTxtBillID.text <> "" Then
        Msg = "”Ì „ Õ–› »Ì«‰«  «·⁄„·Ì… —ﬁ„ " & Chr(13)
        Msg = Msg + (XPTxtBillID.text) & Chr(13)
        Msg = Msg + " Â·  —€» ›Ì Õ–› Â–Â «·»Ì«‰« ø"

        If MsgBox(Msg, vbYesNo + vbQuestion + vbMsgBoxRight + vbMsgBoxRtlReading, App.title) = vbYes Then
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
        .Create Me.hwnd, "⁄—Ê÷ Ã«Â“…", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
        .AddControl Cmd(0), "ÃœÌœ ..." & Wrap & "·«÷«›… »Ì«‰«  ⁄—÷ ÃœÌœ" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
    End With

    With TTP
        .Create Me.hwnd, "⁄—Ê÷ Ã«Â“…", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
        .AddControl Cmd(7), "ÿ»«⁄… ..." & Wrap & "·⁄—÷  ﬁ—Ì— »«·»Ì«‰«  «·Õ«·Ì… " & Wrap & " Ì„ﬂ‰ ÿ»«⁄ Â ⁄‰ ÿ—Ìﬁ «·ÿ«»⁄…", True
    End With

    With TTP
        .Create Me.hwnd, "⁄—Ê÷ Ã«Â“…", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
        .AddControl Cmd(1), " ⁄œÌ· ..." & Wrap & "· ⁄œÌ· »Ì«‰«  «·⁄—÷ «·Õ«·Ì" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
    End With

    With TTP
        .Create Me.hwnd, "⁄—Ê÷ Ã«Â“…", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
        .AddControl Cmd(2), "Õ›Ÿ ..." & Wrap & "·Õ›Ÿ »Ì«‰«  «·⁄—÷ «·ÃœÌœ" & Wrap & "·Õ›Ÿ «· ⁄œÌ·« " & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
    End With

    With TTP
        .Create Me.hwnd, "⁄—Ê÷ Ã«Â“…", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
        .AddControl Cmd(3), " —«Ã⁄ ..." & Wrap & "·· —«Ã⁄ ⁄‰ ⁄„·Ì… «·≈÷«›…" & Wrap & "··· —«Ã⁄ ⁄‰ ⁄„·Ì… «· ⁄œÌ·" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
    End With

    With TTP
        .Create Me.hwnd, "⁄—Ê÷ Ã«Â“…", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
        .AddControl Cmd(4), "Õ–› ..." & Wrap & "·Õ–› »Ì«‰«  «·⁄—÷ «·Õ«·Ì" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
    End With

    With TTP
        .Create Me.hwnd, "⁄—Ê÷ Ã«Â“…", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
        .AddControl Cmd(5), "»ÕÀ ..." & Wrap & "···»ÕÀ ⁄‰ ⁄—÷ " & Wrap & "Ì‰ÿ»ﬁ ⁄·ÌÂ« ‘—Êÿ „⁄Ì‰…" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
    End With

    With TTP
        .Create Me.hwnd, "⁄—Ê÷ Ã«Â“…", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
        .AddControl Cmd(6), "Œ—ÊÃ ..." & Wrap & "·«€·«ﬁ Â–Â «·‰«›–…" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
    End With

    With TTP
        .Create Me.hwnd, "⁄—Ê÷ Ã«Â“…", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
        .AddControl XPBtnAdd, "≈÷«›… «·√’‰«› ..." & Wrap & " ·«÷«›… ’‰› ÃœÌœ" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
    End With

    With TTP
        .Create Me.hwnd, "⁄—Ê÷ Ã«Â“…", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
        .AddControl XPBtnRemove, "Õ–› ’‰› ..." & Wrap & "·Õ–› √Õœ «·√’‰«›" & Wrap & " ÕœœÂ Ê«÷€ÿ Â‰«", True
    End With

    With TTP
        .Create Me.hwnd, "⁄—Ê÷ Ã«Â“…", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
        .AddControl XPBtnMove(1), "«·√Ê· ..." & Wrap & "··«‰ ﬁ«· «·Ï √Ê· ”Ã·" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
    End With

    With TTP
        .Create Me.hwnd, "⁄—Ê÷ Ã«Â“…", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
        .AddControl XPBtnMove(0), "«·”«»ﬁ ..." & Wrap & "··«‰ ﬁ«· «·Ï «·”Ã· «·”«»ﬁ" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
    End With

    With TTP
        .Create Me.hwnd, "⁄—Ê÷ Ã«Â“…", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
        .AddControl XPBtnMove(3), "«· «·Ì ..." & Wrap & "··«‰ ﬁ«· «·Ï «·”Ã· «· «·Ì" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
    End With

    With TTP
        .Create Me.hwnd, "⁄—Ê÷ Ã«Â“…", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
        .AddControl XPBtnMove(2), "«·√ŒÌ— ..." & Wrap & "··«‰ ﬁ«· «·Ï ¬Œ— ”Ã·" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
    End With

    With TTP
        .Create Me.hwnd, "⁄—Ê÷ Ã«Â“…", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
        .AddControl XPFillData, " ⁄»∆… »Ì«‰«  «·√’‰«›" & Wrap & "· ⁄»∆… »Ì«‰«  «·√’‰«› ›Ì" & Wrap & "›Ì ‰«›–… ÕÊ«—" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
    End With

    With TTP
        .Create Me.hwnd, "⁄—Ê÷ Ã«Â“…", 1, 15204351, -2147483630
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
    Dim RsTemp  As New ADODB.Recordset
    Dim RsTest As New ADODB.Recordset
    Dim RsRepeat As ADODB.Recordset
    Dim StrSQL As String
    Dim StrSqlDel As String
    Dim BeginTrans As Boolean
    'On Error GoTo ErrTrap

    Screen.MousePointer = vbArrowHourglass

    If Me.TxtModFlg.text <> "R" Then
        If Txtcode.text = "" Then
            Msg = "ÌÃ»  ”ÃÌ· «”„ «·⁄—÷"
            MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
            Txtcode.SetFocus
            Screen.MousePointer = vbDefault
            Exit Sub
        End If

        If DateDiff("d", Date, Me.DtpExpire) <= 0 Then
            Msg = "ÌÃ» √‰ ÌﬂÊ‰  «—ÌŒ ≈‰ Â«¡ «·⁄—÷ .... ›Ï «·„” ﬁ»·...!!!"
            MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
            DtpExpire.SetFocus
            SendKeys "{F4}"
            Exit Sub
        End If

        Select Case Me.TxtModFlg.text

            Case "N"
                Set RsRepeat = New ADODB.Recordset
                StrSQL = "select * From Templates where TemplateName='" & Trim(Txtcode.text) & "'"
                RsRepeat.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText

                If Not (RsRepeat.EOF Or RsRepeat.BOF) Then
                    Msg = " „  ”ÃÌ· Â–« «·«”„ ·√Õœ «·⁄—Ê÷ «·”«»ﬁ…"
                    MsgBox Msg, vbOKOnly + vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
                    Txtcode.SetFocus
                    Screen.MousePointer = vbDefault
                    Exit Sub
                End If

            Case "E"
                Set RsRepeat = New ADODB.Recordset
                StrSQL = "select * From Templates where TemplateName='" & Trim(Txtcode.text) & "'"
                RsRepeat.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText

                If Not (RsRepeat.EOF Or RsRepeat.BOF) Then
                    If RsRepeat("TemplateID").value <> XPTxtBillID.text Then
                        Msg = " „  ”ÃÌ· Â–« «·«”„ ·√Õœ «·⁄—Ê÷ «·”«»ﬁ…"
                        MsgBox Msg, vbOKOnly + vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
                        Txtcode.SetFocus
                        Screen.MousePointer = vbDefault
                        Exit Sub
                    End If
                End If

        End Select

        For RowNum = 1 To Fg.Rows - 1

            If Fg.TextMatrix(RowNum, Fg.ColIndex("Code")) = "" Then
                Msg = "ÌÃ» «œŒ«· √”„«¡ ﬂ«›… «·√’‰«›" & Chr(13)
                Fg.Row = RowNum
                Fg.Col = Fg.ColIndex("Name")
                Fg.ShowCell RowNum, Fg.ColIndex("Name")
                MsgBox Msg, vbOKOnly + vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
                Fg.SetFocus
                Screen.MousePointer = vbDefault
                Exit Sub
            End If

            If IsNumeric(Trim(Fg.TextMatrix(RowNum, Fg.ColIndex("Code")))) Then
                StrSQL = "select * From TblItems where ItemID=" & Trim(Fg.TextMatrix(RowNum, Fg.ColIndex("Code")))
                RsTemp.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText

                If RsTemp.EOF Or RsTemp.BOF Then
                    Msg = "·« ÊÃœ √Ì »Ì«‰«  ⁄‰  «·’‰›" & Chr(13)
                    Msg = Msg + Trim(Fg.TextMatrix(RowNum, Fg.ColIndex("Code"))) & Chr(13)
                    Msg = Msg + "≈–« ﬂ«‰ ·„ Ì „  ”ÃÌ·Â" & Chr(13)
                    Msg = Msg + "ÌÃ» ≈œŒ«· »Ì«‰« Â „‰ ‘«‘… «·√’‰«›" & Chr(13)
                    MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
                    Fg.Row = RowNum
                    Fg.Col = Fg.ColIndex("Name")
                    Fg.ShowCell RowNum, Fg.ColIndex("Name")
                    Fg.SetFocus
                    Screen.MousePointer = vbDefault
                    Exit Sub
                End If

                RsTemp.Close
            Else
                Msg = "·« ÊÃœ √Ì »Ì«‰«  ⁄‰  «·’‰›" & Chr(13)
                Msg = Msg + Trim(Fg.TextMatrix(RowNum, Fg.ColIndex("Code"))) & Chr(13)
                Msg = Msg + "≈–« ﬂ«‰ ·„ Ì „  ”ÃÌ·Â" & Chr(13)
                Msg = Msg + "ÌÃ» ≈œŒ«· »Ì«‰« Â „‰ ‘«‘… «·√’‰«›" & Chr(13)
                MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
                Fg.Row = RowNum
                Fg.Col = Fg.ColIndex("Name")
                Fg.ShowCell RowNum, Fg.ColIndex("Name")
                Fg.SetFocus
                Screen.MousePointer = vbDefault
                Exit Sub
            End If

            If Fg.TextMatrix(RowNum, Fg.ColIndex("Code")) <> "" Then
                If val(Fg.TextMatrix(RowNum, Fg.ColIndex("Count"))) <= 0 Then
                    Msg = " ÌÃ»  ÕœÌœ «·ﬂ„Ì… «·Œ«’… »«·’‰›" & Chr(13)
                    Msg = Msg + Fg.Cell(flexcpTextDisplay, RowNum, Fg.ColIndex("Name")) & Chr(13)
                    MsgBox Msg, vbOKOnly + vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
                    Fg.Row = RowNum
                    Fg.Col = Fg.ColIndex("Count")
                    Fg.ShowCell RowNum, Fg.ColIndex("Count")
                    Fg.SetFocus
                    Screen.MousePointer = vbDefault
                    Exit Sub
                End If
            End If

            If Fg.TextMatrix(RowNum, Fg.ColIndex("Code")) <> "" Then
                If val(Fg.TextMatrix(RowNum, Fg.ColIndex("Price"))) <= 0 Then
                    If val(Fg.TextMatrix(RowNum, Fg.ColIndex("DiscountType"))) <> 4 Then
                        Msg = "·„ Ì „  ÕœÌœ À„‰ «·’‰›" & Chr(13)
                        Msg = Msg + Fg.Cell(flexcpTextDisplay, RowNum, Fg.ColIndex("Name")) & Chr(13)
                        MsgBox Msg, vbOKOnly + vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
                        Fg.Row = RowNum
                        Fg.Col = Fg.ColIndex("Price")
                        Fg.ShowCell RowNum, Fg.ColIndex("Price")
                        Fg.SetFocus
                        Screen.MousePointer = vbDefault
                        Exit Sub
                    End If
                End If
            End If

        Next RowNum

        For RowNum = 1 To Fg.Rows - 1

            If Fg.TextMatrix(RowNum, Fg.ColIndex("Code")) <> "" Then
                If (Fg.TextMatrix(RowNum, Fg.ColIndex("DiscountType"))) = "" Then
                    Msg = "„‰ ›÷·ﬂ Õœœ ‰Ê⁄ «·Œ’„ «·Œ«’ »«·’‰›" & Chr(13)
                    Msg = Msg + Fg.Cell(flexcpTextDisplay, RowNum, Fg.ColIndex("Name")) & Chr(13)
                    MsgBox Msg, vbOKOnly + vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
                    Fg.Row = RowNum
                    Fg.Col = Fg.ColIndex("DiscountType")
                    Fg.ShowCell RowNum, Fg.ColIndex("DiscountType")
                    Fg.SetFocus
                    Screen.MousePointer = vbDefault
                    Exit Sub
                End If
            End If

        Next RowNum

        Set RSTransDetails = New ADODB.Recordset
        RSTransDetails.Open "[Template_Details]", Cn, adOpenStatic, adLockOptimistic, adCmdTable

        If Me.TxtModFlg.text = "N" Then
            rs.AddNew
        End If

        Cn.BeginTrans
        BeginTrans = True
        Screen.MousePointer = vbArrowHourglass
        rs("TemplateID").value = val(XPTxtBillID.text)
        rs("TemplateName").value = IIf(Txtcode.text = "", "", Trim(Txtcode.text))
        rs("Date").value = XPDtbBill.value
        rs("User_ID").value = user_id

        If XPOpt(1).value = True Then
            rs("TemplateType").value = 1
        ElseIf XPOpt(0).value = True Then
            rs("TemplateType").value = 0
            rs("TemplateTime").value = DtpExpire.value
        End If

        rs.update

        If Me.TxtModFlg.text = "E" Then
            StrSqlDel = "delete From Template_Details where TemplateID=" & val(rs("TemplateID").value)
            Cn.Execute StrSqlDel, , adExecuteNoRecords
        End If

        For RowNum = 1 To Fg.Rows - 1

            If Fg.TextMatrix(RowNum, Fg.ColIndex("Code")) <> "" Then
                RSTransDetails.AddNew
                RSTransDetails("TemplateID").value = val(XPTxtBillID.text)
                RSTransDetails("ItemID").value = IIf((Fg.TextMatrix(RowNum, Fg.ColIndex("Code")) = ""), Null, val(Fg.TextMatrix(RowNum, Fg.ColIndex("Code"))))
                RSTransDetails("Quantity").value = IIf((Fg.TextMatrix(RowNum, Fg.ColIndex("Count")) = ""), Null, val(Fg.TextMatrix(RowNum, Fg.ColIndex("Count"))))
                RSTransDetails("Price").value = IIf((Fg.TextMatrix(RowNum, Fg.ColIndex("Price")) = ""), Null, val(Fg.TextMatrix(RowNum, Fg.ColIndex("Price"))))
                RSTransDetails("ItemDiscountType").value = IIf((Fg.TextMatrix(RowNum, Fg.ColIndex("DiscountType")) = ""), Null, val(Fg.TextMatrix(RowNum, Fg.ColIndex("DiscountType"))))
                RSTransDetails("ItemCase").value = IIf((Fg.TextMatrix(RowNum, Fg.ColIndex("ItemCase")) = ""), Null, val(Fg.TextMatrix(RowNum, Fg.ColIndex("ItemCase"))))
                RSTransDetails("ItemDiscount").value = IIf((Fg.TextMatrix(RowNum, Fg.ColIndex("DiscountVal")) = ""), Null, val(Fg.TextMatrix(RowNum, Fg.ColIndex("DiscountVal"))))
            
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
            
                RSTransDetails.update
        
            End If

        Next RowNum

        Cn.CommitTrans
        BeginTrans = False
        XPTxtCurrent.Caption = rs.AbsolutePosition
        XPTxtCount.Caption = rs.RecordCount

        Select Case Me.TxtModFlg.text

            Case "N"
                Msg = "  „ Õ›Ÿ »Ì«‰«  Â–Â «·⁄„·Ì…" & Chr(13)
                Msg = Msg + "Â·  —€» ›Ì ≈÷«›… »Ì«‰«  √Œ—Ì"

                If MsgBox(Msg, vbYesNo + vbQuestion + vbMsgBoxRight + vbMsgBoxRtlReading + vbDefaultButton2, App.title) = vbYes Then
                    Cmd_Click (0)
                    Screen.MousePointer = vbDefault
                    Exit Sub
                End If
            
            Case "E"
                MsgBox " „ Õ›Ÿ Â–Â «· ⁄œÌ·« ", vbInformation + vbOKOnly + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
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
        MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
        Exit Sub
    End If

    Msg = "⁄›Ê«...ÕœÀ Œÿ√ „« √À‰«¡ Õ›Ÿ Â–Â «·»Ì«‰«  " & Chr(13)
    MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
End Sub

Private Sub XPBtnRemove_Click()
    On Error GoTo ErrTrap

    If Fg.Rows > 1 Then
        If Fg.Rows = 2 Then
            Fg.Clear flexClearScrollable, flexClearEverything
            NewGrid.Calculate 1, True
        Else

            If Fg.Rows > 1 Then
                If Fg.Row <> Fg.FixedRows - 1 Then
                    Fg.RemoveItem (Fg.Row)
                End If
            End If

            NewGrid.Calculate 1
        End If
    End If

    Exit Sub
ErrTrap:
End Sub

Private Sub XPFillData_Click()
    On Error GoTo ErrTrap

    If Fg.TextMatrix(Fg.Rows - 1, Fg.ColIndex("Code")) <> "" Then
        Fg.Rows = Fg.Rows + 1
        NewGrid.GridDefaultValue Fg.Rows - 1
    End If

    With FrmFillItems
        .DealingForm = Template
        .lbl(0).Visible = False
        .TxtGuaranteeTime.Visible = False
        '.C1ElasticSerial.top = FrmFillItems.C1ElasticDiscount.top
        .Height = 3930
        .show vbModal
        NewGrid.Calculate 1
    End With

    Exit Sub
ErrTrap:
End Sub

Private Sub PrintReport()
    On Error GoTo ErrTrap

    If XPTxtBillID.text <> "" Then
        Set SaleReport = New ClsSaleReport
        SaleReport.PrintTemplate XPTxtBillID.text
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

Private Sub XPOpt_Click(Index As Integer)
    DtpExpire.Enabled = XPOpt(0).value
End Sub

