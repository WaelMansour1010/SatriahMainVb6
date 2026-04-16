VERSION 5.00
Object = "{C115893A-A3BF-43AF-B28D-69DB846077F3}#1.0#0"; "vsflex8u.ocx"
Object = "{0AFE7BE0-11B7-4A3E-978D-D4501E9A57FE}#1.0#0"; "c1sizer.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{FE5DCFAD-BC1D-11D2-94CF-004005455FAA}#1.4#0"; "ImpulseButton.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form FrmShowPrice1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "ÿ·»Ì… ‘—«¡"
   ClientHeight    =   7260
   ClientLeft      =   -15
   ClientTop       =   375
   ClientWidth     =   13575
   HelpContextID   =   340
   Icon            =   "FrmShowPrice1.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   RightToLeft     =   -1  'True
   ScaleHeight     =   7260
   ScaleWidth      =   13575
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
      Height          =   7260
      Left            =   0
      TabIndex        =   16
      TabStop         =   0   'False
      Top             =   0
      Width           =   13575
      _cx             =   23945
      _cy             =   12806
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
         Left            =   2535
         TabIndex        =   17
         TabStop         =   0   'False
         Top             =   6540
         Width           =   13560
         _cx             =   23918
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
            Left            =   8490
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
            Left            =   7410
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
            Left            =   6255
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
            Left            =   5250
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
            Left            =   4200
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
            Left            =   3165
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
            Left            =   30
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
            Left            =   2040
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
            Left            =   1065
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
      End
      Begin C1SizerLibCtl.C1Elastic Ele 
         Height          =   435
         Index           =   3
         Left            =   15
         TabIndex        =   27
         TabStop         =   0   'False
         Top             =   6090
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
            Left            =   9990
            Locked          =   -1  'True
            RightToLeft     =   -1  'True
            TabIndex        =   28
            TabStop         =   0   'False
            Top             =   30
            Width           =   1590
         End
         Begin MSDataListLib.DataCombo DCboUserName 
            Height          =   315
            Left            =   6330
            TabIndex        =   29
            Top             =   45
            Width           =   2055
            _ExtentX        =   3625
            _ExtentY        =   556
            _Version        =   393216
            Enabled         =   0   'False
            Text            =   ""
            RightToLeft     =   -1  'True
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "≈Ã„«·Ì «·ÿ·»"
            Height          =   255
            Index           =   3
            Left            =   11670
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
            Left            =   4290
            RightToLeft     =   -1  'True
            TabIndex        =   34
            Top             =   120
            Width           =   1635
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "⁄œœ «·”Ã·« :"
            Height          =   240
            Index           =   2
            Left            =   1290
            RightToLeft     =   -1  'True
            TabIndex        =   33
            Top             =   120
            Width           =   1530
         End
         Begin VB.Label XPTxtCurrent 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Height          =   270
            Left            =   3015
            RightToLeft     =   -1  'True
            TabIndex        =   32
            Top             =   105
            Width           =   1050
         End
         Begin VB.Label XPTxtCount 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Height          =   240
            Left            =   90
            RightToLeft     =   -1  'True
            TabIndex        =   31
            Top             =   135
            Width           =   1095
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Õ—— »Ê«”ÿ… : "
            Height          =   315
            Index           =   1
            Left            =   8490
            RightToLeft     =   -1  'True
            TabIndex        =   30
            Top             =   75
            Width           =   1380
         End
      End
      Begin C1SizerLibCtl.C1Elastic Ele 
         Height          =   3855
         Index           =   5
         Left            =   15
         TabIndex        =   36
         TabStop         =   0   'False
         Top             =   2340
         Width           =   13560
         _cx             =   23918
         _cy             =   6800
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
         Begin C1SizerLibCtl.C1Elastic Ele 
            Height          =   690
            Index           =   2
            Left            =   30
            TabIndex        =   46
            TabStop         =   0   'False
            Top             =   30
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
               ButtonImage     =   "FrmShowPrice1.frx":038A
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
            Height          =   2670
            Left            =   30
            TabIndex        =   9
            Top             =   735
            Width           =   13500
            _cx             =   23812
            _cy             =   4710
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
            FormatString    =   $"FrmShowPrice1.frx":0724
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
            Top             =   3420
            Width           =   12195
            _ExtentX        =   21511
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
            Height          =   405
            Left            =   30
            RightToLeft     =   -1  'True
            TabIndex        =   54
            Top             =   3420
            Width           =   450
         End
      End
      Begin C1SizerLibCtl.C1Elastic Ele 
         Height          =   1755
         Index           =   0
         Left            =   15
         TabIndex        =   37
         TabStop         =   0   'False
         Top             =   645
         Width           =   13545
         _cx             =   23892
         _cy             =   3096
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
         Begin VB.TextBox Txt_order_no 
            Alignment       =   1  'Right Justify
            Height          =   330
            Left            =   10560
            RightToLeft     =   -1  'True
            TabIndex        =   107
            Top             =   0
            Width           =   1665
         End
         Begin VB.Frame Frame3 
            Height          =   1815
            Left            =   15600
            RightToLeft     =   -1  'True
            TabIndex        =   88
            Top             =   1680
            Width           =   7815
            Begin VB.TextBox Text5 
               Alignment       =   1  'Right Justify
               Height          =   285
               Left            =   4080
               RightToLeft     =   -1  'True
               TabIndex        =   89
               Top             =   600
               Width           =   2295
            End
            Begin MSComCtl2.DTPicker DTPicker2 
               Height          =   315
               Left            =   120
               TabIndex        =   90
               Top             =   600
               Width           =   2100
               _ExtentX        =   3704
               _ExtentY        =   556
               _Version        =   393216
               Format          =   91881473
               CurrentDate     =   38784
            End
            Begin MSComCtl2.DTPicker DTPicker3 
               Height          =   315
               Left            =   4800
               TabIndex        =   91
               Top             =   960
               Width           =   1620
               _ExtentX        =   2858
               _ExtentY        =   556
               _Version        =   393216
               Format          =   91881473
               CurrentDate     =   38784
            End
            Begin MSComCtl2.DTPicker DTPicker4 
               Height          =   315
               Left            =   120
               TabIndex        =   92
               Top             =   960
               Width           =   2100
               _ExtentX        =   3704
               _ExtentY        =   556
               _Version        =   393216
               Format          =   91881473
               CurrentDate     =   38784
            End
            Begin MSComCtl2.DTPicker DTPicker5 
               Height          =   315
               Left            =   4800
               TabIndex        =   93
               Top             =   1320
               Width           =   1620
               _ExtentX        =   2858
               _ExtentY        =   556
               _Version        =   393216
               Format          =   91881473
               CurrentDate     =   38784
            End
            Begin MSComCtl2.DTPicker DTPicker6 
               Height          =   315
               Left            =   120
               TabIndex        =   94
               Top             =   1320
               Width           =   2100
               _ExtentX        =   3704
               _ExtentY        =   556
               _Version        =   393216
               Format          =   91881473
               CurrentDate     =   38784
            End
            Begin VB.Label Label8 
               Alignment       =   1  'Right Justify
               Caption         =   " «—ÌŒ «·Ê’Ê· «·„ Êﬁ⁄"
               Height          =   255
               Left            =   2280
               RightToLeft     =   -1  'True
               TabIndex        =   100
               Top             =   1440
               Width           =   1575
            End
            Begin VB.Label Label7 
               Alignment       =   1  'Right Justify
               Caption         =   " «—ÌŒ «· √ŒÌ—"
               Height          =   255
               Left            =   6480
               RightToLeft     =   -1  'True
               TabIndex        =   99
               Top             =   1440
               Width           =   1215
            End
            Begin VB.Label Label6 
               Alignment       =   1  'Right Justify
               Caption         =   "«· «—ÌŒ «·›⁄·Ì"
               Height          =   375
               Left            =   2640
               RightToLeft     =   -1  'True
               TabIndex        =   98
               Top             =   1200
               Width           =   1215
            End
            Begin VB.Label Label5 
               Alignment       =   1  'Right Justify
               Caption         =   "«· «—ÌŒ «·„ Êﬁ⁄"
               Height          =   375
               Left            =   6480
               RightToLeft     =   -1  'True
               TabIndex        =   97
               Top             =   1080
               Width           =   1215
            End
            Begin VB.Label Label3 
               Alignment       =   1  'Right Justify
               Caption         =   "«· «—ÌŒ"
               Height          =   375
               Left            =   2520
               RightToLeft     =   -1  'True
               TabIndex        =   96
               Top             =   720
               Width           =   1335
            End
            Begin VB.Label Label2 
               Alignment       =   1  'Right Justify
               Caption         =   "«·—ﬁ„"
               Height          =   375
               Left            =   6720
               RightToLeft     =   -1  'True
               TabIndex        =   95
               Top             =   720
               Width           =   975
            End
         End
         Begin VB.Frame Frame2 
            Height          =   1815
            Left            =   15240
            RightToLeft     =   -1  'True
            TabIndex        =   75
            Top             =   1680
            Width           =   5700
            Begin VB.TextBox Text7 
               Alignment       =   1  'Right Justify
               Height          =   285
               Left            =   5400
               RightToLeft     =   -1  'True
               TabIndex        =   78
               Top             =   600
               Width           =   3855
            End
            Begin VB.TextBox Text3 
               Alignment       =   1  'Right Justify
               Height          =   285
               Left            =   2640
               RightToLeft     =   -1  'True
               TabIndex        =   77
               Top             =   1320
               Width           =   1455
            End
            Begin VB.TextBox Text2 
               Alignment       =   1  'Right Justify
               Height          =   285
               Left            =   240
               RightToLeft     =   -1  'True
               TabIndex        =   76
               Top             =   960
               Width           =   1335
            End
            Begin MSComCtl2.DTPicker DTPicker1 
               Height          =   315
               Left            =   240
               TabIndex        =   79
               Top             =   1320
               Width           =   1320
               _ExtentX        =   2328
               _ExtentY        =   556
               _Version        =   393216
               Format          =   91881473
               CurrentDate     =   38784
            End
            Begin MSDataListLib.DataCombo DataCombo9 
               Height          =   315
               Left            =   1920
               TabIndex        =   80
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
               TabIndex        =   81
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
               TabIndex        =   87
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
               TabIndex        =   86
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
               TabIndex        =   85
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
               TabIndex        =   84
               Top             =   960
               Width           =   1215
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               Caption         =   "»‰«¡ ⁄·Ï"
               ForeColor       =   &H000000FF&
               Height          =   285
               Index           =   20
               Left            =   9600
               RightToLeft     =   -1  'True
               TabIndex        =   83
               Top             =   480
               Width           =   1095
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               Caption         =   "‰Ê⁄ «·«„—"
               Height          =   285
               Index           =   19
               Left            =   4440
               RightToLeft     =   -1  'True
               TabIndex        =   82
               Top             =   240
               Width           =   1095
            End
         End
         Begin VB.Frame Frame1 
            Height          =   1695
            Left            =   0
            RightToLeft     =   -1  'True
            TabIndex        =   60
            Top             =   0
            Width           =   8895
            Begin VB.ComboBox CboPayMentType 
               Height          =   315
               Left            =   5520
               RightToLeft     =   -1  'True
               Style           =   2  'Dropdown List
               TabIndex        =   113
               Top             =   600
               Width           =   2145
            End
            Begin VB.TextBox txtRemark 
               Alignment       =   1  'Right Justify
               Height          =   645
               Left            =   360
               MultiLine       =   -1  'True
               RightToLeft     =   -1  'True
               ScrollBars      =   2  'Vertical
               TabIndex        =   111
               Top             =   960
               Width           =   7335
            End
            Begin VB.TextBox TxtShipmentArae 
               Alignment       =   1  'Right Justify
               Height          =   285
               Left            =   360
               RightToLeft     =   -1  'True
               TabIndex        =   109
               Top             =   600
               Width           =   3735
            End
            Begin VB.CheckBox chkshipped 
               Alignment       =   1  'Right Justify
               Caption         =   " „ «·‘Õ‰"
               Height          =   195
               Left            =   120
               RightToLeft     =   -1  'True
               TabIndex        =   108
               Top             =   240
               Width           =   975
            End
            Begin VB.TextBox txtShipmentPrice 
               Alignment       =   1  'Right Justify
               Height          =   285
               Left            =   1920
               RightToLeft     =   -1  'True
               TabIndex        =   61
               Top             =   240
               Width           =   2175
            End
            Begin MSDataListLib.DataCombo Dccurrency 
               Height          =   315
               Left            =   9000
               TabIndex        =   62
               Top             =   600
               Width           =   2145
               _ExtentX        =   3784
               _ExtentY        =   556
               _Version        =   393216
               ListField       =   "6"
               Text            =   ""
               RightToLeft     =   -1  'True
            End
            Begin MSDataListLib.DataCombo DataCombo4 
               Height          =   315
               Left            =   9000
               TabIndex        =   63
               Top             =   960
               Width           =   2145
               _ExtentX        =   3784
               _ExtentY        =   556
               _Version        =   393216
               ListField       =   "6"
               Text            =   ""
               RightToLeft     =   -1  'True
            End
            Begin MSDataListLib.DataCombo DcshipmentMethod 
               Height          =   315
               Left            =   5520
               TabIndex        =   64
               Top             =   240
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
               Left            =   9360
               TabIndex        =   65
               Top             =   600
               Width           =   2145
               _ExtentX        =   3784
               _ExtentY        =   556
               _Version        =   393216
               ListField       =   "6"
               Text            =   ""
               RightToLeft     =   -1  'True
            End
            Begin MSDataListLib.DataCombo DataCombo7 
               Height          =   315
               Left            =   9000
               TabIndex        =   66
               Top             =   240
               Width           =   2145
               _ExtentX        =   3784
               _ExtentY        =   556
               _Version        =   393216
               ListField       =   "6"
               Text            =   ""
               RightToLeft     =   -1  'True
            End
            Begin MSDataListLib.DataCombo DataCombo8 
               Height          =   315
               Left            =   120
               TabIndex        =   67
               Top             =   2040
               Width           =   1905
               _ExtentX        =   3360
               _ExtentY        =   556
               _Version        =   393216
               ListField       =   "6"
               Text            =   ""
               RightToLeft     =   -1  'True
            End
            Begin VB.Label Label9 
               Alignment       =   1  'Right Justify
               Caption         =   "„·«ÕŸ« "
               ForeColor       =   &H00000000&
               Height          =   375
               Left            =   7560
               RightToLeft     =   -1  'True
               TabIndex        =   112
               Top             =   960
               Width           =   975
            End
            Begin VB.Label Label1 
               Alignment       =   1  'Right Justify
               Caption         =   "ÃÂ… «· ”·Ì„"
               ForeColor       =   &H00000000&
               Height          =   375
               Left            =   4200
               RightToLeft     =   -1  'True
               TabIndex        =   110
               Top             =   600
               Width           =   855
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               Caption         =   "«· ”⁄Ì—"
               Height          =   285
               Index           =   18
               Left            =   2040
               RightToLeft     =   -1  'True
               TabIndex        =   74
               Top             =   2040
               Width           =   975
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               Caption         =   "ﬁÌ„… «·‘Õ‰"
               ForeColor       =   &H00000000&
               Height          =   285
               Index           =   17
               Left            =   4080
               RightToLeft     =   -1  'True
               TabIndex        =   73
               Top             =   240
               Width           =   975
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               Caption         =   "«· ’‰Ì›"
               Height          =   285
               Index           =   16
               Left            =   9960
               RightToLeft     =   -1  'True
               TabIndex        =   72
               Top             =   240
               Width           =   1215
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               Caption         =   "ÿ—Ìﬁ… «·œ›⁄"
               ForeColor       =   &H00000000&
               Height          =   285
               Index           =   15
               Left            =   7680
               RightToLeft     =   -1  'True
               TabIndex        =   71
               Top             =   600
               Width           =   975
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               Caption         =   "ÿ—Ìﬁ… «·‘Õ‰"
               ForeColor       =   &H00000000&
               Height          =   285
               Index           =   14
               Left            =   7440
               RightToLeft     =   -1  'True
               TabIndex        =   70
               Top             =   240
               Width           =   1215
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               Caption         =   "«·»·œ"
               Height          =   285
               Index           =   13
               Left            =   9720
               RightToLeft     =   -1  'True
               TabIndex        =   69
               Top             =   960
               Width           =   1335
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               Caption         =   "«·⁄„·Â"
               Height          =   285
               Index           =   12
               Left            =   10080
               RightToLeft     =   -1  'True
               TabIndex        =   68
               Top             =   600
               Width           =   1335
            End
         End
         Begin VB.ComboBox CboPriceType 
            Height          =   315
            Left            =   14250
            RightToLeft     =   -1  'True
            Style           =   2  'Dropdown List
            TabIndex        =   4
            Top             =   360
            Width           =   2250
         End
         Begin VB.TextBox TxtTransSerial 
            Alignment       =   1  'Right Justify
            Height          =   330
            Left            =   10560
            RightToLeft     =   -1  'True
            TabIndex        =   0
            Top             =   -240
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
            Left            =   8985
            TabIndex        =   2
            Top             =   840
            Width           =   3270
            _ExtentX        =   5768
            _ExtentY        =   556
            _Version        =   393216
            ListField       =   "6"
            Text            =   ""
            RightToLeft     =   -1  'True
         End
         Begin MSDataListLib.DataCombo DCboStoreName 
            Height          =   315
            Left            =   9825
            TabIndex        =   3
            Top             =   1950
            Visible         =   0   'False
            Width           =   2430
            _ExtentX        =   4286
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
            Format          =   91881473
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
            ButtonImage     =   "FrmShowPrice1.frx":08EA
            ColorButton     =   14871017
            ColorHighlight  =   16777215
            ColorShadow     =   -2147483631
            ColorOutline    =   -2147483631
            DrawFocusRectangle=   0   'False
         End
         Begin ImpulseButton.ISButton CmdTemplate 
            Height          =   480
            Left            =   3465
            TabIndex        =   41
            Top             =   -405
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
            Left            =   5280
            TabIndex        =   5
            TabStop         =   0   'False
            Top             =   -480
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
         Begin MSDataListLib.DataCombo DataCombo1 
            Height          =   315
            Left            =   14250
            TabIndex        =   56
            Top             =   720
            Width           =   2250
            _ExtentX        =   3969
            _ExtentY        =   556
            _Version        =   393216
            ListField       =   "6"
            Text            =   ""
            RightToLeft     =   -1  'True
         End
         Begin MSDataListLib.DataCombo DataCombo2 
            Height          =   315
            Left            =   14250
            TabIndex        =   59
            Top             =   1050
            Width           =   2250
            _ExtentX        =   3969
            _ExtentY        =   556
            _Version        =   393216
            ListField       =   "6"
            Text            =   ""
            RightToLeft     =   -1  'True
         End
         Begin ImpulseButton.ISButton CmdConvert 
            Height          =   285
            Left            =   11280
            TabIndex        =   106
            Top             =   1440
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
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "«·„‘—Ê⁄"
            Height          =   270
            Index           =   11
            Left            =   8580
            RightToLeft     =   -1  'True
            TabIndex        =   58
            Top             =   810
            Width           =   1095
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "„—ﬂ“ «· ﬂ·›…"
            Height          =   285
            Index           =   10
            Left            =   16380
            RightToLeft     =   -1  'True
            TabIndex        =   57
            Top             =   720
            Width           =   1095
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "‰Ê⁄ «·«„—"
            Height          =   240
            Index           =   9
            Left            =   16500
            RightToLeft     =   -1  'True
            TabIndex        =   55
            Top             =   360
            Width           =   855
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "—ﬁ„ «·ÿ·»Ì…"
            ForeColor       =   &H00000000&
            Height          =   390
            Index           =   5
            Left            =   12570
            RightToLeft     =   -1  'True
            TabIndex        =   45
            Top             =   120
            Width           =   810
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   " «—ÌŒ «·ÿ·»Ì…"
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   6
            Left            =   12360
            RightToLeft     =   -1  'True
            TabIndex        =   44
            Top             =   600
            Width           =   1050
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "«·⁄„Ì· / «·„Ê—œ"
            ForeColor       =   &H00000000&
            Height          =   240
            Index           =   7
            Left            =   12315
            RightToLeft     =   -1  'True
            TabIndex        =   43
            Top             =   960
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
            Top             =   1920
            Visible         =   0   'False
            Width           =   945
         End
      End
      Begin C1SizerLibCtl.C1Elastic Ele 
         Height          =   615
         Index           =   6
         Left            =   15
         TabIndex        =   101
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
         Caption         =   "ÿ·»Ì… ‘—«¡"
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
            TabIndex        =   102
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
            ButtonImage     =   "FrmShowPrice1.frx":0C84
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
            TabIndex        =   103
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
            ButtonImage     =   "FrmShowPrice1.frx":101E
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
            TabIndex        =   104
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
            ButtonImage     =   "FrmShowPrice1.frx":13B8
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
            TabIndex        =   105
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
            ButtonImage     =   "FrmShowPrice1.frx":1752
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
Attribute VB_Name = "FrmShowPrice1"
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
        If DoPremis(Do_New, Me.name, True) = False Then
            Exit Sub
        End If
        clear_all Me
        TxtModFlg.text = "N"
        XPTxtBillID.text = CStr(new_id("Transactions", "Transaction_ID", "", True))
        TxtTransSerial.text = CStr(new_id("Transactions", "Transaction_Serial", "", True, "Transaction_Type=25"))
        NewGrid.GridDefaultValue 1
        Me.DCboUserName.BoundText = user_id
        intDef = GetSetting(StrAppRegPath, "DefaultOptions", "DefaultClient", 2)
        DBCboClientName.BoundText = intDef
        intDef = GetSetting(StrAppRegPath, "DefaultOptions", "DefaultSaleStore", 1)
        DCboStoreName.BoundText = intDef
        FG.SetFocus
        FG.Col = FG.ColIndex("Code")
        FG.Row = FG.Rows - 1
        Me.CboPriceType.ListIndex = 0
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
        FrmBuySearch.DealingForm = GridTransType.ShowPrice
        FrmBuySearch.Caption = "«·»ÕÀ ⁄‰ ⁄—÷ ”⁄—"
        FrmBuySearch.Show vbModal
    Case 7
        If DoPremis(Do_Print, Me.name, True) = False Then
            Exit Sub
        End If
        'PrintReport
        PrintReport1 (Txt_order_no.text)
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
    StrFileName = App.Path & "\Reports\" & "Order_status.rpt"
Else
    StrFileName = App.Path & "\Reports\" & "Order_status.rpt"
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
    xReport.ParameterFields(4).AddCurrentValue get_branch_name(Val(my_branch))
    StrReportTitle = "Order status" '& StrAccountName
 
Else
 
    xReport.ParameterFields(1).AddCurrentValue cCompanyInfo.ArabCompanyName ' RPTCompany_Name_Eng
    'xReport.ParameterFields(2).AddCurrentValue RPTComment_Eng
     xReport.ParameterFields(4).AddCurrentValue get_branch_name(Val(my_branch))
    StrReportTitle = "Order status"
 
End If
xReport.ParameterFields(3).AddCurrentValue user_name
xReport.ReportTitle = StrReportTitle
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

        
        .FG.Cell(flexcpData, .FG.Rows - 1, FG.ColIndex("UnitID")) = IIf(IsNull(RsUnit("UnitID")), "", (RsUnit("UnitID").value))
        .FG.TextMatrix(.FG.Rows - 1, FG.ColIndex("UnitID")) = IIf(IsNull(RsUnit("UnitName")), "", (RsUnit("UnitName").value))
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
    .Show vbModeless, MDIFrmMain
    .Visible = True
End With
Exit Sub
ErrTrap:
End Sub

Private Sub DCboItemsCode_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyF3 Then
        Load FrmItemSearch
        FrmItemSearch.RetrunType = 2
        FrmItemSearch.Show vbModal
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

Private Sub FG_CellButtonClick(ByVal Row As Long, ByVal Col As Long)
If Me.TxtModFlg.text = "N" Or Me.TxtModFlg.text = "E" Then
'    FrmAddNewItem.Tag = "xx"
    FrmAddNewItem.DealingForm = ShowPrice
    FrmAddNewItem.Show vbModal
End If
End Sub

Private Sub Form_Activate()
'XPTxtBillID.SetFocus
End Sub

Private Sub ISButton1_Click()
'Frame3.Visible = True
End Sub

Private Sub Label10_Click()
Frame3.Visible = False
End Sub

Private Sub TxtFillData_Change()
If TxtFillData.text = "F" Then
    NewGrid.Calculate 1, , , True
End If
End Sub

Private Sub XPBtnMove_Click(Index As Integer)
On Error GoTo ErrTrap
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
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
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
If SystemOptions.UserInterface = EnglishInterface Then
    SetInterface Me
    ChangeLang
End If
ShowTax = GetSetting(StrAppRegPath, "SallBill", "HaveTaxOnSalles", False)
Ele(4).Visible = ShowTax
Set Cmd(0).ButtonImage = MDIFrmMain.ImgLstTree.ListImages("New").Picture
Set Cmd(1).ButtonImage = MDIFrmMain.ImgLstTree.ListImages("Edit").Picture
Set Cmd(2).ButtonImage = MDIFrmMain.ImgLstTree.ListImages("save").Picture
Set Cmd(3).ButtonImage = MDIFrmMain.ImgLstTree.ListImages("Undo").Picture
Set Cmd(4).ButtonImage = MDIFrmMain.ImgLstTree.ListImages("Del").Picture
Set Cmd(5).ButtonImage = MDIFrmMain.ImgLstTree.ListImages("Search").Picture
Set Cmd(6).ButtonImage = MDIFrmMain.ImgLstTree.ListImages("Exit").Picture
Set Cmd(7).ButtonImage = MDIFrmMain.ImgLstTree.ListImages("Print").Picture
Set CmdHelp.ButtonImage = MDIFrmMain.ImgLstTree.ListImages("Help").Picture
Set CmdConvert.ButtonImage = MDIFrmMain.ImgLstTree.ListImages("Excute").Picture
Set CmdTemplate.ButtonImage = MDIFrmMain.ImgLstTree.ListImages("Excute").Picture
Set NewGrid.Grid = FG
NewGrid.GridTrans = GridTransType.ShowPrice
Set NewGrid.TxtModFlag = TxtModFlg
Set NewGrid.txttotal = XPTxtSum
Set NewGrid.TxtFillData = TxtFillData
Set NewGrid.TxtTaxValue = Me.XPTxtTaxValue
Set NewGrid.GrdTBar = Me.TBar
Set NewGrid.LblItemsCount = Me.LblItemsCount
' ⁄»∆… »Ì«‰«  «·√’‰«›
Set NewGrid.DcboItemName = DCboItemsName
Set NewGrid.DCboItemCode = DCboItemsCode
Set NewGrid.CboItemCase = CboItemCase
Set NewGrid.CmdAddData = CmdAdd
'Set NewGrid.TxtSerial = TxtSerial
Set NewGrid.TxtQuantity = TxtQuantity
Set NewGrid.TxtPrice = TxtPrice
' Resize_Form Me, TransactionSize
 Me.left = (MDIFrmMain.Width - Me.Width) / 2
    Me.top = (MDIFrmMain.Height - Me.Height) / 2 - 500
FG.WallPaper = BGround.Picture
AddTip
XPDtbBill.value = Date
Set Dcombos = New ClsDataCombos
Dcombos.GetCustomersSuppliers 0, Me.DBCboClientName, True
Dcombos.GetStores Me.DCboStoreName
Dcombos.GetUsers Me.DCboUserName

Set cSearchDcbo(0) = New clsDCboSearch
Set cSearchDcbo(0).Client = Me.DBCboClientName

Set cSearchDcbo(1) = New clsDCboSearch
Set cSearchDcbo(1).Client = Me.DCboStoreName
NewGrid.FillGrid

With Me.CboPriceType
    .Clear
    If SystemOptions.UserInterface = ArabicInterface Then
        .AddItem "⁄—÷ √”⁄«— ·›Ê« Ì— «·»Ì⁄"
        .AddItem "ÿ·»Ì… ‘—«¡"
    Else
        .AddItem "Sales  Order"
        .AddItem "Purchases   Order"
    End If
    .ListIndex = 0
End With

    With CboPaymentType
        .Clear
        .AddItem "‰ﬁœ«"
        .AddItem "«Ã·"
    End With
    
StrSQL = "SELECT * FROM Transactions WHERE (Transaction_Type=25 )" 'OR Transaction_Type=17
StrSQL = StrSQL + " Order By Transaction_ID"
Set rs = New ADODB.Recordset
rs.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText
Dim My_SQL As String
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
 
fill_combo Me.DcshipmentMethod, My_SQL




XPBtnMove_Click 2
Me.TxtModFlg.text = "R"
If OPEN_NEW_SCREEN = True Then
Cmd_Click (0)
End If

Exit Sub
ErrTrap:
End Sub
Private Sub Form_Unload(Cancel As Integer)
On Error GoTo ErrTrap
Dim I As Integer
If rs.state = adStateOpen Then
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
Private Sub TxtModFlg_Change()
On Error GoTo ErrTrap
Select Case Me.TxtModFlg.text
    Case "R"
         Me.Caption = "ÿ·»Ì… ‘—«¡"
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
        
        CmdConvert.Enabled = True
       ' CmdConvert.Visible = True
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
        End If
        Ele(2).Enabled = False
    Case "N"
         Me.Caption = "ÿ·»Ì… ‘—«¡( ÃœÌœ )"
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
        FG.Rows = 2
        Me.XPDtbBill.Enabled = True
        XPDtbBill.value = Date
        Me.DBCboClientName.locked = False
        Me.DCboStoreName.locked = False
        FG.Editable = flexEDKbdMouse
        
     '   CmdConvert.Visible = False
        CmdTemplate.Enabled = True
        CmdTemplate.Visible = True
        Ele(2).Enabled = True
        CboItemCase.ListIndex = 0
    Case "E"
         Me.Caption = "ÿ·»Ì… ‘—«¡ (  ⁄œÌ· )"
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
        
        
       
       ' CmdConvert.Visible = False
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
TxtFillData.text = "T"
Screen.MousePointer = vbArrowHourglass
XPTxtBillID.text = IIf(IsNull(rs("Transaction_ID").value), "", Val(rs("Transaction_ID").value))

Txt_order_no.text = IIf(IsNull(rs("order_no").value), "", rs("order_no").value)
If rs("shipped").value = True Then
    chkshipped.value = vbChecked
Else
    chkshipped.value = Unchecked
End If
  
'Me.DataCombo4.BoundText = IIf(IsNull(rs("countryid").value), "", rs("countryid").value)

TxtTransSerial.text = IIf(IsNull(rs("Transaction_Serial").value), "", (rs("Transaction_Serial").value))
XPDtbBill.value = IIf(IsNull(rs("Transaction_Date").value), "", (rs("Transaction_Date").value))
Me.DBCboClientName.BoundText = IIf(IsNull(rs("CusID").value), "", rs("CusID").value)
Me.DcshipmentMethod.BoundText = IIf(IsNull(rs("shipmentMethod").value), "", rs("shipmentMethod").value)
txtShipmentPrice.text = IIf(Not IsNumeric(rs("ShipmentPrice").value), 0, (rs("ShipmentPrice").value))
TxtShipmentArae.text = IIf(IsNull(rs("ShipmentArae").value), "", (rs("ShipmentArae").value))
txtremark.text = IIf(IsNull(rs("Remark").value), "", (rs("Remark").value))
'Dccurrency.BoundText = IIf(IsNull(rs("Currency_id").value), "", rs("Currency_id").value)
'If rs("Transaction_Type").value = 6 Then
'    Me.CboPriceType.ListIndex = 1
'ElseIf rs("Transaction_Type").value = 17 Then '17
'    Me.CboPriceType.ListIndex = 0
'End If
Me.DCboUserName.BoundText = IIf(IsNull(rs("UserID").value), "", rs("UserID").value)
'Me.DCboStoreName.BoundText = IIf(IsNull(rs("StoreID").value), "", rs("StoreID").value)
CboPaymentType.ListIndex = IIf(IsNull(rs("PaymentType").value), 0, rs("PaymentType").value)

XPTxtTaxValue.text = IIf(IsNull(rs("TaxValue").value), "", (rs("TaxValue").value))
XPChkTAX.value = IIf(rs("TaxFound") = True, Checked, Unchecked)

FG.Clear flexClearScrollable, flexClearEverything
FG.Rows = 2
FG.Clear flexClearScrollable, flexClearEverything
FG.Refresh
StrSQL = "SELECT TblItems.HaveSerial, * FROM TblItems INNER JOIN Transaction_Details " & _
"ON TblItems.ItemID = Transaction_Details.Item_ID INNER JOIN dbo.TblUnites ON dbo.Transaction_Details.UnitID = dbo.TblUnites.UnitID"
StrSQL = StrSQL + " where Transaction_ID=" & Val(rs("Transaction_ID").value)

RsDetails.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText
XPTxtSum.text = ""
If Not (RsDetails.EOF Or RsDetails.BOF) Then
    FG.Rows = RsDetails.RecordCount + 1
    For Num = 1 To RsDetails.RecordCount
        FG.TextMatrix(Num, FG.ColIndex("Code")) = IIf(IsNull(RsDetails("Item_ID")), "", (RsDetails("Item_ID").value))
        FG.TextMatrix(Num, FG.ColIndex("Name")) = IIf(IsNull(RsDetails("Item_ID")), "", Trim(RsDetails("Item_ID").value))
        FG.TextMatrix(Num, FG.ColIndex("Count")) = IIf(IsNull(RsDetails("Quantity")), "", (RsDetails("Quantity").value))
        FG.TextMatrix(Num, FG.ColIndex("Price")) = IIf(IsNull(RsDetails("Price")), "", (RsDetails("Price").value))
        FG.TextMatrix(Num, FG.ColIndex("ItemCase")) = IIf(IsNull(RsDetails("ItemCase")), "", (RsDetails("ItemCase").value))
        FG.TextMatrix(Num, FG.ColIndex("DiscountType")) = IIf(IsNull(RsDetails("ItemDiscountType")), "", (RsDetails("ItemDiscountType").value))
        FG.TextMatrix(Num, FG.ColIndex("DiscountVal")) = IIf(IsNull(RsDetails("ItemDiscount")), "", (RsDetails("ItemDiscount").value))
        
        If RsDetails("HaveSerial") = True Then
            FG.TextMatrix(Num, FG.ColIndex("HaveSerial")) = True
        End If
        
        FG.Cell(flexcpData, Num, FG.ColIndex("UnitID")) = IIf(IsNull(RsDetails("UnitID")), "", (RsDetails("UnitID").value))
        FG.TextMatrix(Num, FG.ColIndex("UnitID")) = IIf(IsNull(RsDetails("UnitName")), "", (RsDetails("UnitName").value))
        
        RsDetails.MoveNext
        Debug.Print Num
        If FG.Rows > 10 Then
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
            rs.find "Transaction_ID='" & Val(XPTxtBillID.text) & "'", , adSearchForward, adBookmarkFirst
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
    If MsgBox(Msg, vbYesNo + vbQuestion + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title) = vbYes Then
         If Not rs.RecordCount < 1 Then
            rs.Delete
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
    MsgBox Msg, vbMsgBoxRight + vbMsgBoxRtlReading + _
            vbExclamation, App.Title
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
   .AddControl Cmd(0), _
    "ÃœÌœ ..." & Wrap & _
    "·«÷«›… »Ì«‰«  ⁄—÷ ”⁄— ÃœÌœ" & Wrap & _
    " ›ﬁÿ ≈÷€ÿ Â‰«", True
End With
With TTP
   .Create Me.hWnd, "⁄—÷ √”⁄«—", 1, 15204351, -2147483630
   .MaxWidth = 4000
   .VisibleTime = 9000
   .DelayTime = 600
   .AddControl Cmd(7), _
    "ÿ»«⁄… ..." & Wrap & _
    "·⁄—÷  ﬁ—Ì— »«·»Ì«‰«  «·Õ«·Ì… " & Wrap & _
    " Ì„ﬂ‰ ÿ»«⁄ Â ⁄‰ ÿ—Ìﬁ «·ÿ«»⁄…", True
End With
With TTP
   .Create Me.hWnd, "⁄—÷ √”⁄«—", 1, 15204351, -2147483630
   .MaxWidth = 4000
   .VisibleTime = 9000
   .DelayTime = 600
   .AddControl Cmd(1), _
    " ⁄œÌ· ..." & Wrap & _
    "· ⁄œÌ· »Ì«‰«  «·⁄—÷ «·Õ«·Ì" & Wrap & _
    " ›ﬁÿ ≈÷€ÿ Â‰«", True
End With
With TTP
   .Create Me.hWnd, "⁄—÷ √”⁄«—", 1, 15204351, -2147483630
   .MaxWidth = 4000
   .VisibleTime = 9000
   .DelayTime = 600
   .AddControl Cmd(2), _
    "Õ›Ÿ ..." & Wrap & _
    "·Õ›Ÿ »Ì«‰«  ⁄—÷ «·”⁄— «·ÃœÌœ" & Wrap & _
     "·Õ›Ÿ «· ⁄œÌ·« " & Wrap & _
    " ›ﬁÿ ≈÷€ÿ Â‰«", True
End With
With TTP
   .Create Me.hWnd, "⁄—÷ √”⁄«—", 1, 15204351, -2147483630
   .MaxWidth = 4000
   .VisibleTime = 9000
   .DelayTime = 600
   .AddControl Cmd(3), _
    " —«Ã⁄ ..." & Wrap & _
    "·· —«Ã⁄ ⁄‰ ⁄„·Ì… «·≈÷«›…" & Wrap & _
     "··· —«Ã⁄ ⁄‰ ⁄„·Ì… «· ⁄œÌ·" & Wrap & _
    " ›ﬁÿ ≈÷€ÿ Â‰«", True
End With
 With TTP
   .Create Me.hWnd, "⁄—÷ √”⁄«—", 1, 15204351, -2147483630
   .MaxWidth = 4000
   .VisibleTime = 9000
   .DelayTime = 600
   .AddControl Cmd(4), _
    "Õ–› ..." & Wrap & _
    "·Õ–› »Ì«‰«  «·⁄—÷ «·Õ«·Ì" & Wrap & _
    " ›ﬁÿ ≈÷€ÿ Â‰«", True
End With
With TTP
   .Create Me.hWnd, "⁄—÷ √”⁄«—", 1, 15204351, -2147483630
   .MaxWidth = 4000
   .VisibleTime = 9000
   .DelayTime = 600
   .AddControl Cmd(5), _
    "»ÕÀ ..." & Wrap & _
    "···»ÕÀ ⁄‰ ⁄—÷ ”⁄—" & Wrap & _
    "Ì‰ÿ»ﬁ ⁄·ÌÂ« ‘—Êÿ „⁄Ì‰…" & Wrap & _
    " ›ﬁÿ ≈÷€ÿ Â‰«", True
End With
With TTP
   .Create Me.hWnd, "⁄—÷ √”⁄«—", 1, 15204351, -2147483630
   .MaxWidth = 4000
   .VisibleTime = 9000
   .DelayTime = 600
   .AddControl Cmd(6), _
    "Œ—ÊÃ ..." & Wrap & _
    "·«€·«ﬁ Â–Â «·‰«›–…" & Wrap & _
    " ›ﬁÿ ≈÷€ÿ Â‰«", True
End With


With TTP
   .Create Me.hWnd, "⁄—÷ √”⁄«—", 1, 15204351, -2147483630
   .MaxWidth = 4000
   .VisibleTime = 9000
   .DelayTime = 600
   .AddControl XPBtnNewClients, _
    "≈÷«›… ⁄„Ì· ÃœÌœ ..." & Wrap & _
    "· ”ÃÌ· »Ì«‰«  ⁄„Ì· ÃœÌœ" & Wrap & _
    " «÷€ÿ Â‰«", True
End With

With TTP
   .Create Me.hWnd, "⁄—÷ √”⁄«—", 1, 15204351, -2147483630
   .MaxWidth = 4000
   .VisibleTime = 9000
   .DelayTime = 600
   .AddControl XPBtnMove(1), _
    "«·√Ê· ..." & Wrap & _
    "··«‰ ﬁ«· «·Ï √Ê· ”Ã·" & Wrap & _
    " ›ﬁÿ ≈÷€ÿ Â‰«", True
End With
With TTP
   .Create Me.hWnd, "⁄—÷ √”⁄«—", 1, 15204351, -2147483630
   .MaxWidth = 4000
   .VisibleTime = 9000
   .DelayTime = 600
   .AddControl XPBtnMove(0), _
    "«·”«»ﬁ ..." & Wrap & _
    "··«‰ ﬁ«· «·Ï «·”Ã· «·”«»ﬁ" & Wrap & _
    " ›ﬁÿ ≈÷€ÿ Â‰«", True
End With
With TTP
   .Create Me.hWnd, "⁄—÷ √”⁄«—", 1, 15204351, -2147483630
   .MaxWidth = 4000
   .VisibleTime = 9000
   .DelayTime = 600
   .AddControl XPBtnMove(3), _
    "«· «·Ì ..." & Wrap & _
    "··«‰ ﬁ«· «·Ï «·”Ã· «· «·Ì" & Wrap & _
    " ›ﬁÿ ≈÷€ÿ Â‰«", True
End With
With TTP
   .Create Me.hWnd, "⁄—÷ √”⁄«—", 1, 15204351, -2147483630
   .MaxWidth = 4000
   .VisibleTime = 9000
   .DelayTime = 600
   .AddControl XPBtnMove(2), _
    "«·√ŒÌ— ..." & Wrap & _
    "··«‰ ﬁ«· «·Ï ¬Œ— ”Ã·" & Wrap & _
    " ›ﬁÿ ≈÷€ÿ Â‰«", True
End With


With TTP
   .Create Me.hWnd, "⁄—÷ √”⁄«—", 1, 15204351, -2147483630
   .MaxWidth = 4000
   .VisibleTime = 9000
   .DelayTime = 600
   .AddControl CmdHelp, _
    "„”«⁄œ… ..." & Wrap & _
    "·· ⁄—› ⁄·Ï ÊŸÌ›… Â–Â «·‰«›–…" & Wrap & _
    "ÊﬂÌ›Ì… «· ⁄«„· „⁄Â«" & Wrap & _
    "≈÷€ÿ Â‰«" & Wrap, True
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
If Me.TxtModFlg.text <> "R" Then
    If DBCboClientName.text = "" Then
      If SystemOptions.UserInterface = ArabicInterface Then
        Msg = "„‰ ›÷·ﬂ √œŒ· «”„ «·⁄„Ì·"
      Else
      Msg = "Please Select Vendor"
      End If
        MsgBox Msg, vbOKOnly + vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
        DBCboClientName.SetFocus
        SendKeys "{F4}"
        Screen.MousePointer = vbDefault
        Exit Sub
    End If
'    If DCboStoreName.text = "" Then
'     If SystemOptions.UserInterface = ArabicInterface Then
'        Msg = "ÌÃ»  ÕœÌœ «·„Œ“‰"
'    Else
'    Msg = "Select Inventory"
'    End If
'        MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
'        DCboStoreName.SetFocus
'        SendKeys "{F4}"
'        Screen.MousePointer = vbDefault
'        Exit Sub
'    End If
    
'        If Dccurrency.text = "" Then
'        If SystemOptions.UserInterface = ArabicInterface Then
'        Msg = "Õœœ «·⁄„·…"
'        Else
'        Msg = "Select Currency"
'        End If
'        MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
'        Dccurrency.SetFocus
'        SendKeys "{F4}"
'        Screen.MousePointer = vbDefault
'        Exit Sub
'    End If
    
    
'    If Me.CboPriceType.ListIndex = -1 Then
'     If SystemOptions.UserInterface = ArabicInterface Then
'        Msg = "ÌÃ»  ÕœÌœ ‰Ê⁄ ⁄—÷ «·”⁄— (⁄—÷ ”⁄— ›« Ê—… »Ì⁄ «Ê ÿ·»Ì… ‘—«¡)...!!!"
'    Else
'    Msg = "Specify Order Type"
'    End If
'        MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
'        CboPriceType.SetFocus
'        SendKeys "{F4}"
'        Screen.MousePointer = vbDefault
'        Exit Sub
'    End If
'    If XPChkTAX.value = Checked Then
'        If XPTxtTaxValue.text = "" Then
'        If SystemOptions.UserInterface = ArabicInterface Then
'            Msg = "ÌÃ» «œŒ«· ﬁÌ„… ÷—Ì»… «·„»Ì⁄« "
'        Else
'        Msg = "Insert Sales Tax"
'        End If
'            MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
'            XPTxtTaxValue.SetFocus
'            FG.SetFocus
'            Screen.MousePointer = vbDefault
'            Exit Sub
'        End If
'    End If
    
    If CboPaymentType.ListIndex = -1 Then
    If SystemOptions.UserInterface = ArabicInterface Then
        Msg = "ÌÃ»  ÕœÌœ ÿ—Ìﬁ… «·œ›⁄"
    Else
    Msg = "Specify Payment Method"
    End If
     End If
    
    If NewGrid.CheckDataEntered = False Then
        Exit Sub
    End If
    Set RSTransDetails = New ADODB.Recordset
    RSTransDetails.Open "[Transaction_Details]", Cn, adOpenStatic, adLockOptimistic, adCmdTable
    Cn.BeginTrans
    BeginTrans = True
    If Me.TxtModFlg.text = "N" Then
        rs.AddNew
    End If
    Screen.MousePointer = vbArrowHourglass
    rs("Transaction_ID").value = Val(XPTxtBillID.text)
    rs("order_no").value = Txt_order_no.text
    
    If chkshipped.value = vbChecked Then
        rs("shipped").value = 1
    Else
        rs("shipped").value = 0
    End If
    
    rs("Transaction_Date").value = XPDtbBill.value
    rs("Transaction_Serial").value = TxtTransSerial.text
    'If Me.CboPriceType.ListIndex = 0 Then
    '    rs("Transaction_Type").value = 6
    'ElseIf Me.CboPriceType.ListIndex = 1 Then
    '    rs("Transaction_Type").value = 6 '17
    'End If
     rs("Transaction_Type").value = 25
       If CboPaymentType.ListIndex = -1 Then
        rs("PaymentType").value = 0
    Else
        rs("PaymentType").value = Val(CboPaymentType.ListIndex)
    End If
    
    rs("UserID").value = user_id
    rs("CusID").value = IIf(DBCboClientName.BoundText = "", Null, Val(DBCboClientName.BoundText))
    rs("shipmentMethod").value = IIf(DcshipmentMethod.BoundText = "", Null, Val(DcshipmentMethod.BoundText))
      rs("ShipmentPrice").value = IIf(txtShipmentPrice.text = "", 0, Val(txtShipmentPrice.text))
      rs("ShipmentArae").value = IIf(TxtShipmentArae.text = "", Null, TxtShipmentArae.text)
      rs("Remark").value = IIf(txtremark.text = "", Null, txtremark.text)
    
'    rs("countryid").value = IIf(DataCombo4.BoundText = "", Null, Val(DataCombo4.BoundText))
    
'    rs("Currency_id").value = IIf(Dccurrency.BoundText = "", Null, Val(Dccurrency.BoundText))
    
    
'    rs("StoreID").value = IIf(DCboStoreName.BoundText = "", Null, Val(DCboStoreName.BoundText))
    rs("TaxFound").value = IIf(XPChkTAX.value = Checked, True, False)
    rs("TaxValue").value = IIf(XPTxtTaxValue.text = "", Null, Val(XPTxtTaxValue.text))
    rs("total").value = IIf(XPTxtSum.text = "", Null, Val(XPTxtSum.text))
    
    rs.update
    If Me.TxtModFlg.text = "E" Then
       StrSqlDel = "delete From Transaction_Details where Transaction_ID=" & Val(rs("Transaction_ID").value)
       Cn.Execute StrSqlDel, , adExecuteNoRecords
    End If
    For RowNum = 1 To FG.Rows - 1
        If FG.TextMatrix(RowNum, FG.ColIndex("Code")) <> "" Then
            RSTransDetails.AddNew
            RSTransDetails("Transaction_ID").value = Val(XPTxtBillID.text)
            RSTransDetails("order_id").value = Val(XPTxtBillID.text)
             
             RSTransDetails("order_no").value = Txt_order_no.text
             
            RSTransDetails("Item_ID").value = IIf((FG.TextMatrix(RowNum, FG.ColIndex("Code")) = ""), Null, Val(FG.TextMatrix(RowNum, FG.ColIndex("Code"))))
            RSTransDetails("Quantity").value = IIf((FG.TextMatrix(RowNum, FG.ColIndex("Count")) = ""), Null, Val(FG.TextMatrix(RowNum, FG.ColIndex("Count"))))
            RSTransDetails("Price").value = IIf((FG.TextMatrix(RowNum, FG.ColIndex("Price")) = ""), Null, Val(FG.TextMatrix(RowNum, FG.ColIndex("Price"))))
            RSTransDetails("ItemDiscountType").value = IIf((FG.TextMatrix(RowNum, FG.ColIndex("DiscountType")) = ""), Null, Val(FG.TextMatrix(RowNum, FG.ColIndex("DiscountType"))))
            RSTransDetails("ItemCase").value = IIf((FG.TextMatrix(RowNum, FG.ColIndex("ItemCase")) = ""), Null, Val(FG.TextMatrix(RowNum, FG.ColIndex("ItemCase"))))
            RSTransDetails("ItemDiscount").value = IIf((FG.TextMatrix(RowNum, FG.ColIndex("DiscountVal")) = ""), Null, Val(FG.TextMatrix(RowNum, FG.ColIndex("DiscountVal"))))
            
             RSTransDetails("UnitID").value = _
         IIf(FG.Cell(flexcpData, RowNum, FG.ColIndex("UnitID")) = "", Null, (FG.Cell(flexcpData, RowNum, FG.ColIndex("UnitID"))))
       RSTransDetails("ShowQty").value = IIf((FG.TextMatrix(RowNum, FG.ColIndex("Count")) = ""), Null, Val(FG.TextMatrix(RowNum, FG.ColIndex("Count"))))
 

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
        If SystemOptions.UserInterface = ArabicInterface Then
            Msg = "  „ Õ›Ÿ »Ì«‰«  Â–Â «·⁄„·Ì…" & Chr(13)
            Msg = Msg + "Â·  —€» ›Ì ≈÷«›… »Ì«‰«  √Œ—Ì"
        Else
          Msg = " Saved Successfully" & Chr(13)
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
        MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
        Exit Sub
    End If
     If SystemOptions.UserInterface = ArabicInterface Then
    Msg = "⁄›Ê«...ÕœÀ Œÿ√ „« √À‰«¡ Õ›Ÿ Â–Â «·»Ì«‰«  " & Chr(13)
    Else
    Msg = "Sorry... Error During Saving " & Chr(13)
    End If
    MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
End Sub
Private Sub XPBtnNewClients_Click()
With FrmAddNewCustemer
    .DealingForm = ShowPrice
    .Show vbModal
    .Caption = "≈÷«›… ⁄„Ì· ÃœÌœ"
    .lbl(1).Caption = "ﬂÊœ «·⁄„Ì·"
    .lbl(0).Caption = "«”„ «·⁄„Ì·"
End With
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

Private Sub PrintReport()
On Error GoTo ErrTrap
If XPTxtBillID.text <> "" Then
    Set SaleReport = New ClsSaleReport
    SaleReport.ShowPrice XPTxtBillID.text
End If
Exit Sub
ErrTrap:
End Sub
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
Dim StrMSG As String
Dim IntResult As String
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
Public Sub Cala()
NewGrid.Calculate 1
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
Me.Caption = "Order Request"
    With CboPaymentType
        .Clear
        .AddItem "Cash"
        .AddItem "Credit"
    End With
Ele(6).Caption = Me.Caption
lbl(5).Caption = "No"
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
lbl(18).Caption = "Order Price"
lbl(19).Caption = "Kind Of Order"
lbl(24).Caption = "Expiry Date"
lbl(20).Caption = "Credit Bank"
lbl(21).Caption = "Credit Curr."
lbl(22).Caption = "Credit No."
lbl(23).Caption = "Value"
'ISButton1.Caption = "Show Port Data"
Label1.Caption = "Port"
Label2.Caption = "Supp info No."
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
