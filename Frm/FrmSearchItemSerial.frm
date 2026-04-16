VERSION 5.00
Object = "{C115893A-A3BF-43AF-B28D-69DB846077F3}#1.0#0"; "vsflex8u.ocx"
Object = "{0AFE7BE0-11B7-4A3E-978D-D4501E9A57FE}#1.0#0"; "c1sizer.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{FE5DCFAD-BC1D-11D2-94CF-004005455FAA}#1.4#0"; "ImpulseButton.ocx"
Begin VB.Form FrmSerialData 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00E2E9E9&
   Caption         =   "«·«” ⁄·«„ ⁄‰ ”Ì—Ì«·"
   ClientHeight    =   6945
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9330
   HelpContextID   =   220
   Icon            =   "FrmSearchItemSerial.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   RightToLeft     =   -1  'True
   ScaleHeight     =   6945
   ScaleWidth      =   9330
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
      Height          =   6945
      Left            =   0
      TabIndex        =   10
      TabStop         =   0   'False
      Top             =   0
      Width           =   9330
      _cx             =   16457
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
      BackColor       =   -2147483633
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
      GridRows        =   4
      GridCols        =   4
      Frame           =   3
      FrameStyle      =   0
      FrameWidth      =   1
      FrameColor      =   -2147483628
      FrameShadow     =   -2147483632
      FloodStyle      =   1
      _GridInfo       =   $"FrmSearchItemSerial.frx":0CCA
      AccessibleName  =   ""
      AccessibleDescription=   ""
      AccessibleValue =   ""
      AccessibleRole  =   9
      Begin C1SizerLibCtl.C1Elastic Ele 
         Height          =   2340
         Index           =   2
         Left            =   30
         TabIndex        =   13
         TabStop         =   0   'False
         Top             =   720
         Width           =   9270
         _cx             =   16351
         _cy             =   4128
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
         Begin C1SizerLibCtl.C1Elastic Ele 
            Height          =   825
            Index           =   3
            Left            =   5160
            TabIndex        =   23
            TabStop         =   0   'False
            Top             =   30
            Width           =   3975
            _cx             =   7011
            _cy             =   1455
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   8.25
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
            ForeColor       =   192
            FloodColor      =   6553600
            ForeColorDisabled=   -2147483631
            Caption         =   "≈Œ — ‰Ê⁄ «·»ÕÀ"
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
            Begin VB.OptionButton Opt 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "»ÕÀ ⁄‰ ”Ì—Ì«· „⁄Ì‰"
               Height          =   255
               Index           =   1
               Left            =   750
               RightToLeft     =   -1  'True
               TabIndex        =   25
               Top             =   240
               Value           =   -1  'True
               Width           =   3075
            End
            Begin VB.OptionButton Opt 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "»ÕÀ ⁄‰ ﬂ· «—ﬁ«„ «·”Ì—Ì«· «·Œ«’… »’‰›"
               Height          =   255
               Index           =   0
               Left            =   750
               RightToLeft     =   -1  'True
               TabIndex        =   24
               Top             =   510
               Width           =   3075
            End
         End
         Begin VB.TextBox XPTxtCode 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00C0FFFF&
            Height          =   375
            Left            =   4230
            RightToLeft     =   -1  'True
            TabIndex        =   1
            Top             =   900
            Width           =   3930
         End
         Begin VB.TextBox TxtItemCode 
            Alignment       =   1  'Right Justify
            Height          =   345
            Left            =   6210
            RightToLeft     =   -1  'True
            TabIndex        =   3
            Top             =   1575
            Width           =   1950
         End
         Begin VB.CheckBox Chk 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "»ÕÀ Ã“∆Ï"
            Height          =   255
            Left            =   7710
            RightToLeft     =   -1  'True
            TabIndex        =   2
            Top             =   1275
            Width           =   1440
         End
         Begin ImpulseButton.ISButton CmdUpdate 
            Height          =   345
            Left            =   2100
            TabIndex        =   14
            Top             =   915
            Width           =   1995
            _ExtentX        =   3519
            _ExtentY        =   609
            ButtonStyle     =   1
            ButtonPositionImage=   1
            Caption         =   " ⁄œÌ· Â–« «·”Ì—Ì«·"
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
            ButtonImage     =   "FrmSearchItemSerial.frx":0D4D
            ColorButton     =   14871017
            DrawFocusRectangle=   0   'False
         End
         Begin MSDataListLib.DataCombo DcboItemName 
            Height          =   315
            Left            =   1410
            TabIndex        =   4
            Top             =   1965
            Width           =   6750
            _ExtentX        =   11906
            _ExtentY        =   556
            _Version        =   393216
            Text            =   ""
            RightToLeft     =   -1  'True
         End
         Begin ImpulseButton.ISButton CmdItemSearch 
            Height          =   345
            Left            =   855
            TabIndex        =   15
            Top             =   1935
            Width           =   480
            _ExtentX        =   847
            _ExtentY        =   609
            ButtonStyle     =   1
            ButtonPositionImage=   1
            Caption         =   "..."
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   178
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ButtonImage     =   "FrmSearchItemSerial.frx":10E7
            DrawFocusRectangle=   0   'False
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "«·”Ì—Ì«·"
            Height          =   315
            Index           =   0
            Left            =   8055
            RightToLeft     =   -1  'True
            TabIndex        =   0
            Top             =   960
            Width           =   1035
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "ﬂÊœ «·’‰›"
            Height          =   315
            Index           =   4
            Left            =   8055
            RightToLeft     =   -1  'True
            TabIndex        =   19
            Top             =   1665
            Width           =   1035
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "«”„ «·’‰›"
            Height          =   315
            Index           =   3
            Left            =   8055
            RightToLeft     =   -1  'True
            TabIndex        =   18
            Top             =   1965
            Width           =   1035
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "«ﬂ » ﬂÊœ «·’‰› À„ ≈÷€ÿ ≈‰ —"
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
            Index           =   2
            Left            =   3285
            RightToLeft     =   -1  'True
            TabIndex        =   17
            Top             =   1635
            Width           =   2880
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "≈Œ Ì«— «·»ÕÀ «·Ã“∆Ï ÌﬁÊ„ »«·»ÕÀ ⁄‰ √ﬁ—» ”Ì—Ì«· „‘«»Â… ·Â–« «·—ﬁ„"
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
            Index           =   1
            Left            =   1755
            RightToLeft     =   -1  'True
            TabIndex        =   16
            Top             =   1305
            Width           =   5910
         End
      End
      Begin C1SizerLibCtl.C1Elastic Ele 
         Height          =   840
         Index           =   1
         Left            =   30
         TabIndex        =   11
         TabStop         =   0   'False
         Top             =   6075
         Width           =   9270
         _cx             =   16351
         _cy             =   1482
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
         Begin ImpulseButton.ISButton CmdPrintReport 
            Height          =   360
            Left            =   6240
            TabIndex        =   20
            Top             =   0
            Width           =   2820
            _ExtentX        =   4974
            _ExtentY        =   635
            ButtonStyle     =   1
            ButtonPositionImage=   1
            Caption         =   " ﬁ—Ì— ⁄„·Ì«  «·’Ì«‰…"
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
            ButtonImage     =   "FrmSearchItemSerial.frx":1681
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
            Height          =   360
            Left            =   1065
            TabIndex        =   8
            Top             =   360
            Width           =   1155
            _ExtentX        =   2037
            _ExtentY        =   635
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
            ButtonImage     =   "FrmSearchItemSerial.frx":1A1B
            ColorButton     =   14871017
            ColorHighlight  =   16777215
            ColorHoverText  =   16711680
            ColorShadow     =   4210752
            ColorOutline    =   0
            DrawFocusRectangle=   0   'False
            RightToLeft     =   -1  'True
            ColorToggledHoverText=   16711680
            ColorTextShadow =   4210752
         End
         Begin ImpulseButton.ISButton Cmd 
            Height          =   360
            Index           =   0
            Left            =   3255
            TabIndex        =   6
            Top             =   360
            Width           =   885
            _ExtentX        =   1561
            _ExtentY        =   635
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
            BackStyle       =   0
            ColorButton     =   14871017
            ColorHighlight  =   16777215
            ColorHoverText  =   16711680
            ColorShadow     =   4210752
            ColorOutline    =   0
            DrawFocusRectangle=   0   'False
            DisabledImageExtraction=   0
            RightToLeft     =   -1  'True
            ColorToggledHoverText=   16711680
            LowerToggledContent=   0   'False
            ColorTextShadow =   4210752
         End
         Begin ImpulseButton.ISButton Cmd 
            Height          =   360
            Index           =   1
            Left            =   2340
            TabIndex        =   7
            Top             =   360
            Width           =   825
            _ExtentX        =   1455
            _ExtentY        =   635
            ButtonStyle     =   1
            ButtonPositionImage=   1
            Caption         =   "„”Õ"
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
            ColorButton     =   14871017
            ColorHighlight  =   16777215
            ColorHoverText  =   16711680
            ColorShadow     =   4210752
            ColorOutline    =   0
            DrawFocusRectangle=   0   'False
            RightToLeft     =   -1  'True
            ColorToggledHoverText=   16711680
            LowerToggledContent=   0   'False
            ColorTextShadow =   4210752
         End
         Begin ImpulseButton.ISButton Cmd 
            Cancel          =   -1  'True
            Height          =   360
            Index           =   2
            Left            =   120
            TabIndex        =   9
            Top             =   360
            Width           =   885
            _ExtentX        =   1561
            _ExtentY        =   635
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
            BackStyle       =   0
            ColorButton     =   14871017
            ColorHighlight  =   16777215
            ColorHoverText  =   16711680
            ColorShadow     =   4210752
            ColorOutline    =   0
            DrawFocusRectangle=   0   'False
            RightToLeft     =   -1  'True
            ColorToggledHoverText=   16711680
            LowerToggledContent=   0   'False
            ColorTextShadow =   4210752
         End
         Begin VB.Label LblPlace 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Height          =   255
            Left            =   6030
            RightToLeft     =   -1  'True
            TabIndex        =   22
            Top             =   60
            Width           =   3150
         End
         Begin VB.Label LblRemark 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "«÷€ÿ ⁄·Ï √Ì ⁄„·Ì… ÷€ÿ… „“œÊÃ… ·Ì „ ⁄—÷ »Ì«‰« Â« »’Ê—… „›’·…"
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
            Height          =   255
            Left            =   -120
            RightToLeft     =   -1  'True
            TabIndex        =   21
            Top             =   30
            Width           =   6030
         End
      End
      Begin C1SizerLibCtl.C1Elastic Ele 
         Height          =   675
         Index           =   0
         Left            =   30
         TabIndex        =   12
         TabStop         =   0   'False
         Top             =   30
         Width           =   9270
         _cx             =   16351
         _cy             =   1191
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial (Arabic)"
            Size            =   18
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Enabled         =   -1  'True
         Appearance      =   0
         MousePointer    =   0
         Version         =   801
         BackColor       =   16777215
         ForeColor       =   4210688
         FloodColor      =   6553600
         ForeColorDisabled=   -2147483631
         Picture         =   "FrmSearchItemSerial.frx":1DB5
         Caption         =   "«·⁄„·Ì«  «· Ì  „  ⁄·Ï ﬁÿ⁄… „⁄Ì‰…"
         Align           =   0
         AutoSizeChildren=   0
         BorderWidth     =   2
         ChildSpacing    =   1
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
         PicturePos      =   1
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
      Begin VSFlex8UCtl.VSFlexGrid FG 
         Height          =   2985
         Left            =   30
         TabIndex        =   5
         Top             =   3075
         Width           =   9270
         _cx             =   16351
         _cy             =   5265
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
         Rows            =   15
         Cols            =   9
         FixedRows       =   1
         FixedCols       =   0
         RowHeightMin    =   300
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   -1  'True
         FormatString    =   $"FrmSearchItemSerial.frx":2A8F
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
         WallPaperAlignment=   4
         AccessibleName  =   ""
         AccessibleDescription=   ""
         AccessibleValue =   ""
         AccessibleRole  =   24
      End
   End
End
Attribute VB_Name = "FrmSerialData"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rs As New ADODB.Recordset
Dim cDcboSearch As clsDCboSearch

Public Sub Cmd_Click(Index As Integer)

    Dim Msg As String
    Dim StrSQL As String
    Dim RsTemp As ADODB.Recordset
    Dim cProgress As ClsProgress

    Dim i As Integer

    On Error GoTo ErrTrap

    Select Case Index

        Case 0

            If rs.State = adStateOpen Then
                rs.Close
            End If

            If Me.Opt(1).value = True Then
                If XPTxtCode.text = "" Then
                    Msg = "«ﬂ » —ﬁ„ «·”Ì—Ì«· «·„—«œ «·»ÕÀ ⁄‰Â...!!! "
                    MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
                    XPTxtCode.SetFocus
                    Exit Sub
                End If

            ElseIf Me.Opt(0).value = True Then

                If val(Me.DcboItemName.BoundText) = 0 Then
                    Msg = "ÌÃ» ≈Œ Ì«— «·’‰› «·„—«œ «·»ÕÀ ⁄‰Â...!!! "
                    MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
                    XPTxtCode.SetFocus
                    Exit Sub
                End If
            End If

            If SystemOptions.SysDataBaseType = AccessDataBase Then
                StrSQL = "Select * From SearchSerialData "
            
                If Me.Chk.value = vbUnchecked Then
                    StrSQL = StrSQL + " Where ItemSerial='" & Trim(XPTxtCode.text) & "'"
                Else
                    StrSQL = StrSQL + " Where ItemSerial like '%" & Trim(XPTxtCode.text) & "%'"
                End If

                If Me.DcboItemName.BoundText <> "" Then
                    StrSQL = StrSQL + " and ItemID=" & Me.DcboItemName.BoundText & ""
                End If

                StrSQL = StrSQL + " Order by Transaction_ID "
            ElseIf SystemOptions.SysDataBaseType = SQLServerDataBase Then

                If Me.Opt(1).value = True Then
                    StrSQL = " Select ItemID,ItemName,ItemCode,Count(ItemID) as  CountID From"
                    StrSQL = StrSQL + "("
                    StrSQL = StrSQL + " SELECT SearchSerialData.* FROM dbo.SearchSerialData() SearchSerialData"

                    If Me.Chk.value = vbUnchecked Then
                        StrSQL = StrSQL + " Where ItemSerial='" & Trim(XPTxtCode.text) & "'"
                    Else
                        StrSQL = StrSQL + " Where ItemSerial like '%" & Trim(XPTxtCode.text) & "%'"
                    End If

                    If Me.DcboItemName.BoundText <> "" Then
                        StrSQL = StrSQL + " and ItemID=" & Me.DcboItemName.BoundText & ""
                    End If

                    StrSQL = StrSQL + " )XTable"
                    StrSQL = StrSQL + " Group By ItemID,ItemName,ItemCode"
                    StrSQL = StrSQL + " Order By ItemCode"
            
                    Set RsTemp = New ADODB.Recordset
                    RsTemp.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText + adAsyncExecute
                    Set cProgress = New ClsProgress
                    cProgress.ProgressType = Waiting
                    cProgress.StartProgress

                    Do While RsTemp.State = adStateExecuting
                        DoEvents
                    Loop

                    cProgress.StopProgess
                    Set cProgress = Nothing

                    If Not (RsTemp.BOF Or RsTemp.EOF) Then
                        If RsTemp.RecordCount > 1 Then
                            Load FrmChooseItem

                            With FrmChooseItem.FG
                                .Rows = .FixedRows + RsTemp.RecordCount
                                RsTemp.MoveFirst

                                For i = .FixedRows To RsTemp.RecordCount
                                    .TextMatrix(i, .ColIndex("Ser")) = i
                                    .TextMatrix(i, .ColIndex("ItemID")) = IIf(IsNull(RsTemp("ItemID").value), "", RsTemp("ItemID").value)
                                    .TextMatrix(i, .ColIndex("ItemName")) = IIf(IsNull(RsTemp("ItemName").value), "", RsTemp("ItemName").value)
                                    .TextMatrix(i, .ColIndex("ItemCode")) = IIf(IsNull(RsTemp("ItemCode").value), "", RsTemp("ItemCode").value)
                                    .TextMatrix(i, .ColIndex("CountID")) = IIf(IsNull(RsTemp("CountID").value), "", RsTemp("CountID").value)
                                    RsTemp.MoveNext
                                Next i

                                .AutoSize 0, .Cols - 1, False
                                FrmChooseItem.lbl(2).Caption = .Aggregate(flexSTCount, .FixedRows, .ColIndex("ItemID"), .Rows - 1, .ColIndex("ItemID"))
                            End With

                            FrmChooseItem.show vbModal

                            If FrmChooseItem.UserCanceld = False Then
                                Me.DcboItemName.BoundText = FrmChooseItem.LngChooseItemID
                                Unload FrmChooseItem
                            Else
                                Unload FrmChooseItem
                                Exit Sub
                            End If
                        End If
                    End If

                    StrSQL = "SELECT SearchSerialData.* FROM dbo.SearchSerialData() SearchSerialData "

                    If Me.Chk.value = vbUnchecked Then
                        StrSQL = StrSQL + " Where ItemSerial='" & Trim(XPTxtCode.text) & "'"
                    Else
                        StrSQL = StrSQL + " Where ItemSerial like '%" & Trim(XPTxtCode.text) & "%'"
                    End If

                    If Me.DcboItemName.BoundText <> "" Then
                        StrSQL = StrSQL + " and ItemID=" & Me.DcboItemName.BoundText & ""
                    End If

                    If SystemOptions.SysDataBaseType = SQLServerDataBase Then
                        StrSQL = StrSQL + " Order By  MainOPerationID "
                    End If

                ElseIf Me.Opt(0).value = True Then
                    StrSQL = "SELECT SearchSerialData.* FROM dbo.SearchSerialData() SearchSerialData "

                    If Me.DcboItemName.BoundText <> "" Then
                        StrSQL = StrSQL + " Where ItemID=" & Me.DcboItemName.BoundText & ""
                    End If

                    If SystemOptions.SysDataBaseType = SQLServerDataBase Then
                        StrSQL = StrSQL + " Order By  MainOPerationID "
                    End If
                End If
            End If

            rs.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText + adAsyncExecute
            Set cProgress = New ClsProgress
            cProgress.ProgressType = Waiting
            cProgress.StartProgress

            Do While rs.State = adStateExecuting
                DoEvents
            Loop

            cProgress.StopProgess
            Set cProgress = Nothing
            Retrive

        Case 1
            XPTxtCode.text = ""
            Me.DcboItemName.BoundText = ""
            Me.TxtItemCode.text = ""
            LblPlace.Caption = ""
            FG.Clear flexClearScrollable, flexClearEverything

        Case 2
            Unload Me
    End Select

    Exit Sub
ErrTrap:

    If Err.Number = -2147217900 Then
        If Not cProgress Is Nothing Then
            cProgress.StopProgess
            Set cProgress = Nothing
        End If

        Msg = Msg + "·ﬁœ  „ «œŒ«· ﬁÌ„ €Ì— ’«·Õ… " & Chr(13)
        Msg = Msg + " √ﬂœ „‰ œﬁ… „⁄«ÌÌ— «·»ÕÀ Ê√⁄œ «·„Õ«Ê·…"
        MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
        Exit Sub
    End If

End Sub

Private Sub CmdHelp_Click()
    SystemOptions.SysHelp.HHTopicID = Me.HelpContextID
    SystemOptions.SysHelp.HHDisplayTopicID Me.hwnd
End Sub

Private Sub CmdItemSearch_Click()
    Load FrmItemSearch
    FrmItemSearch.RetrunType = 1
    Set FrmItemSearch.DcboItems = Me.DcboItemName
    FrmItemSearch.show vbModal
End Sub

Private Sub CmdPrintReport_Click()
    On Error GoTo ErrTrap
    Dim StrSQL As String
    Dim Msg As String
    Dim PrinReport As ClsRepoerts
    Set PrinReport = New ClsRepoerts

    If XPTxtCode.text = "" Then
        Msg = "ÌÃ»  ÕœÌœ «·”Ì—Ì«· «·–Ì  —€» ›Ì «· ⁄—› ⁄·Ï «·⁄„·Ì«  «· Ì  „  ⁄·ÌÂ"
        MsgBox Msg, vbOKOnly + vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
        XPTxtCode.SetFocus
        Exit Sub
    End If

    StrSQL = "select * From SearchSerialData where Transaction_Type='14' and ItemSerial='" & Trim(XPTxtCode.text) & "'"

    If Me.DcboItemName.BoundText <> "" Then
        StrSQL = StrSQL + " and ItemID=" & Me.DcboItemName.BoundText & ""
    End If

    PrinReport.SerialMaintenance StrSQL
    Exit Sub
ErrTrap:
End Sub

Private Sub CmdUpdate_Click()
    Dim StrNewSerial As String
    Dim StrSQL As String
    Dim Msg As String
    Dim StrTransIDs As String
    Dim rs As ADODB.Recordset
    Dim BolBegine As Boolean
    Dim StrMSQL As String
    On Error GoTo ErrTrap
    StrTransIDs = GetTransIDs

    If Trim(StrTransIDs) = "" Then
        Msg = "·«»œ «‰  ﬁÊ„ »⁄„·Ì… «·»ÕÀ Õ Ï Ì „ «· ⁄œÌ· ..!!!"
        MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
        Exit Sub
    End If

    If Me.DcboItemName.BoundText = "" Then
        Msg = "ﬁ»· ≈Ã—«¡ ⁄„·Ì…  ⁄œÌ· «·”Ì—Ì«· ·«»œ „‰  ÕœÌœ «·’‰› ...!!"
        MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
        DcboItemName.SetFocus
        SendKeys "{F4}"
        Exit Sub
    End If

    StrNewSerial = Trim(InputBox("√œŒ· «·”Ì—Ì«· «·ÃœÌœ..", App.title, XPTxtCode.text))

    If Trim(StrNewSerial) = "" Then Exit Sub

    If Trim(StrNewSerial) = Trim(Me.XPTxtCode.text) Then
        Msg = "»—Ã«¡ ≈œŒ«· ”Ì—Ì«· ÃœÌœ.. €Ì— Â–« «·”Ì—Ì«· «·„œŒ·"
        MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
        Exit Sub
    End If

    StrSQL = "Update Transaction_Details Set ItemSerial='" & Trim(StrNewSerial) & "'"

    If StrTransIDs <> "" Then
        StrSQL = StrSQL + " Where Transaction_Details.Transaction_ID IN(" & StrTransIDs & ")"
    End If

    StrSQL = StrSQL + " And Item_ID=" & Me.DcboItemName.BoundText & ""
    StrSQL = StrSQL + " And Transaction_Details.ItemSerial='" & Trim(Me.XPTxtCode.text) & "'"

    StrMSQL = "Update TblMainteneceDetails Set ItemSerial='" & Trim(StrNewSerial) & "'"

    If StrTransIDs <> "" Then
        StrMSQL = StrMSQL + " Where TblMainteneceDetails.MaintananceID IN(" & StrTransIDs & ")"
    End If

    StrMSQL = StrMSQL + " And ItemID=" & Me.DcboItemName.BoundText & ""
    StrMSQL = StrMSQL + " And TblMainteneceDetails.ItemSerial='" & Trim(Me.XPTxtCode.text) & "'"

    Cn.BeginTrans
    BolBegine = True
    Cn.Execute StrSQL
    Cn.Execute StrMSQL
    Cn.CommitTrans
    BolBegine = False
    Msg = " „  ⁄„·Ì… «· ⁄œÌ· »‰Ã«Õ"
    MsgBox Msg, vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
    Exit Sub
ErrTrap:

    If BolBegine = True Then
        Cn.RollbackTrans
        BolBegine = False
    End If

    Msg = "›‘·  ⁄„·Ì… «· ⁄œÌ·.!!"
    MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
End Sub

Private Sub DcboItemName_Change()
    Dim StrItemCode As String

    If Me.DcboItemName.BoundText <> "" Then
        StrItemCode = GetItemCode(val(Me.DcboItemName.BoundText))

        If StrItemCode <> Trim(Me.TxtItemCode.text) Then
            Me.TxtItemCode.text = StrItemCode
        End If
    End If

End Sub

Private Sub DcboItemName_Click(Area As Integer)
    DcboItemName_Change
End Sub

Private Sub DcboItemName_DblClick(Area As Integer)

    If Me.DcboItemName.BoundText <> "" Then
        Load FrmSelectData
        FrmSelectData.DcboItemName.BoundText = Me.DcboItemName.BoundText
        FrmSelectData.show
    End If

End Sub

Private Sub Fg_DblClick()
    On Error GoTo ErrTrap
    Dim TransType As Integer
    Dim TransID As Long
    Dim LngCusID As Long
    Dim StrTemp As String

    If FG.Col <= 0 Then Exit Sub
    If FG.Row <= 0 Then Exit Sub
    If FG.Col = FG.ColIndex("Client") Then
        LngCusID = val(FG.TextMatrix(FG.Row, FG.ColIndex("CusID")))

        If LngCusID <> 0 Then
            ShowCusBalDailog LngCusID, 0
        End If

    ElseIf FG.Col = FG.ColIndex("SerialNumber") Then
        Me.Opt(1).value = True
        Opt_Click 1
        StrTemp = Trim$(FG.TextMatrix(FG.Row, FG.ColIndex("SerialNumber")))

        If StrTemp <> "" Then
            Me.XPTxtCode.text = StrTemp

            DoEvents
            Cmd_Click 0
        End If

    Else
        TransType = IIf(FG.Cell(flexcpData, FG.Row, FG.ColIndex("TransType")) = "", "", FG.Cell(flexcpData, FG.Row, FG.ColIndex("TransType")))
    
        If FG.TextMatrix(FG.Row, FG.ColIndex("TranseNum")) = "" Then Exit Sub
        TransID = FG.TextMatrix(FG.Row, FG.ColIndex("TranseNum"))

        Select Case TransType

            Case 1

                With FrmBillBuy
                    .show
                    .Retrive (TransID)
                End With

            Case 2

                With frmsalebill
                    .show
                    .Retrive (TransID)
                End With

            Case 3

                With FrmOpeningBalance
                    .show
                    .Retrive (TransID)
                End With

            Case 5

                With FrmReturnpurchases
                    .show
                    .Retrive (TransID)
                End With

            Case 8

                With FrmDestruction
                    .show
                    .Retrive (TransID)
                End With

            Case 9

                With FrmReturnSalling
                    .show
                    .Retrive (TransID)
                End With

            Case 10

                With FrmMoving
                    .show
                    .Retrive (TransID)
                End With

            Case 14

                'With FrmMaintenence
                '    .show
                '    .Retrive (TransID)
                'End With

        End Select

    End If

    Exit Sub
ErrTrap:
End Sub

Private Sub Form_Activate()
    ShowDynamicHelp Me.HelpContextID
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, _
                         Shift As Integer)
    'If KeyCode = vbKeyReturn Then
    '    If Trim(Me.TxtItemCode.text) <> "" Then
    '        If Me.ActiveControl Is Me.TxtItemCode Then
    '        Cmd_Click 0
    '    End If
    'End If
End Sub

Private Sub Form_Load()
    Dim RsTemp As New ADODB.Recordset
    Dim StrSQL As String
    Dim Dcombos As ClsDataCombos
    Dim BG As New ClsBackGroundPic
    On Error GoTo ErrTrap

    If SystemOptions.UserInterface = EnglishInterface Then
        SetInterface Me
        ChangeLang
    End If

    Set Cmd(0).ButtonImage = mdifrmmain.ImgLstTree.ListImages("Search").Picture
    Set Cmd(1).ButtonImage = mdifrmmain.ImgLstTree.ListImages("Clear").Picture
    Set Cmd(2).ButtonImage = mdifrmmain.ImgLstTree.ListImages("Exit").Picture

    Set Dcombos = New ClsDataCombos
    Dcombos.GetItemsNames DcboItemName
    Set cDcboSearch = New clsDCboSearch
    Set cDcboSearch.Client = DcboItemName
    FG.WallPaper = BG.SearchWallpaper

    With Me.FG
        .AutoSize 0, .Cols - 1, False
    End With

    Me.Width = 9450
    Me.Height = 7455
    CenterForm Me

    FormPostion Me, GetPostion
    Me.Opt(1).value = True
    Opt_Click 1
    Exit Sub
ErrTrap:
End Sub

Private Sub Form_Unload(Cancel As Integer)
    On Error GoTo ErrTrap

    FormPostion Me, SavePostion

    If rs.State = adStateOpen Then
        rs.Close
        Set rs = Nothing
    End If

    Set cDcboSearch = Nothing
    Exit Sub
ErrTrap:
End Sub

Private Sub Retrive()
    Dim Num As Integer
    Dim RsTemp As New ADODB.Recordset
    Dim StrSQL As String
    Dim Msg As String
    Dim StrTemp As String
    Dim StrItemsNames As String

    On Error GoTo ErrTrap
    FG.Clear flexClearScrollable, flexClearEverything

    If Not (rs.EOF Or rs.BOF) Then
        rs.MoveFirst
        FG.Rows = rs.RecordCount + 1
        LblRemark.Visible = True
    
        Me.TxtItemCode.text = IIf(IsNull(rs("ItemCode").value), "", rs("ItemCode").value)

        If Me.DcboItemName.BoundText <> rs("ItemID").value Then
            Me.DcboItemName.BoundText = rs("ItemID").value
        End If

        For Num = 1 To rs.RecordCount

            With FG
                '            If Opt(2).Value = True Then
                .TextMatrix(Num, .ColIndex("NumIndex")) = Num
                .TextMatrix(Num, .ColIndex("Transaction_Serial")) = IIf(IsNull(rs("Transaction_Serial").value), "", (rs("Transaction_Serial").value))
                .TextMatrix(Num, .ColIndex("SerialNumber")) = IIf(IsNull(rs("ItemSerial").value), "", (rs("ItemSerial").value))
                .TextMatrix(Num, .ColIndex("CusID")) = IIf(IsNull(rs("CusID").value), "", (rs("CusID").value))
              
                .TextMatrix(Num, .ColIndex("TranseNum")) = IIf(IsNull(rs("Transaction_ID").value), "", (rs("Transaction_ID").value))
                
                .Cell(flexcpData, Num, .ColIndex("TransType")) = IIf(IsNull(rs("Transaction_Type").value), "", (rs("Transaction_Type").value))
                
                If SystemOptions.UserInterface = ArabicInterface Then
                    .TextMatrix(Num, .ColIndex("TransType")) = IIf(IsNull(rs("TransactionTypeName").value), "", (rs("TransactionTypeName").value))
                
                    .TextMatrix(Num, .ColIndex("ItemPlace")) = IIf(IsNull(rs("StoreName").value), "", (rs("StoreName").value))
                    .TextMatrix(Num, .ColIndex("Client")) = IIf(IsNull(rs("CusName").value), "", (rs("CusName").value))
                Else
                    Dim TransactionEnglishName As String
                    Dim CusNamee As String
                    Dim storenamee As String
                    GetTransactionsEData val(.TextMatrix(Num, .ColIndex("TranseNum"))), TransactionEnglishName, CusNamee, storenamee
                    .TextMatrix(Num, .ColIndex("TransType")) = TransactionEnglishName
                
                    .TextMatrix(Num, .ColIndex("ItemPlace")) = storenamee
                    .TextMatrix(Num, .ColIndex("Client")) = CusNamee
                End If
           
                .TextMatrix(Num, .ColIndex("TransDate")) = IIf(IsNull(rs("Transaction_Date").value), "", Format((rs("Transaction_Date").value), "yyyy/m/d"))
                
            End With

            rs.MoveNext

            DoEvents
        Next Num

        ' ÕœÌœ „ﬂ«‰ «·ﬁÿ⁄Â
        If SystemOptions.SysDataBaseType = AccessDataBase Then
            StrSQL = "select * From QryGardComplete where ItemSerial='" & XPTxtCode.text & "'"

            If Me.DcboItemName.BoundText <> "" Then
                StrSQL = StrSQL + " and ItemID=" & Me.DcboItemName.BoundText & ""
            End If

        ElseIf SystemOptions.SysDataBaseType = SQLServerDataBase Then
            StrSQL = "Select * From dbo.QryGardComplete(0)QryGardComplete where ItemSerial='" & XPTxtCode.text & "'"

            If Me.DcboItemName.BoundText <> "" Then
                StrSQL = StrSQL + " and ItemID=" & Me.DcboItemName.BoundText & ""
            End If
        End If

        RsTemp.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

        If Not (RsTemp.BOF Or RsTemp.EOF) Then
            If RsTemp("QTY").value > 0 Then
                LblPlace.Caption = "„ÊÃÊœ… ›Ì «·„Œ“‰ " & RsTemp("StoreName").value
            Else
                LblPlace.Caption = "€Ì— „ÊÃÊœ… ›Ì «·„Œ“‰/«·„Œ«“‰"
            End If

        Else
            LblPlace.Caption = "€Ì— „ÊÃÊœ… ›Ì «·„Œ“‰/«·„Œ«“‰"
        End If

        RsTemp.Close
        Set RsTemp = Nothing
        FG.AutoSize 0, FG.Cols - 1, False
    Else
        FG.Clear flexClearScrollable, flexClearEverything
        FG.Rows = 1
        LblPlace.Caption = ""
        LblRemark.Visible = False
        Msg = "·«  ÊÃœ √Ì »Ì«‰«  ⁄‰ Â–Â «·ﬁÿ⁄… "
        MsgBox Msg, vbOKOnly + vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
        Exit Sub
    End If

    Exit Sub
ErrTrap:
End Sub

Private Sub Opt_Click(Index As Integer)
    Me.lbl(0).Enabled = Me.Opt(1).value
    Me.XPTxtCode.Enabled = Me.Opt(1).value
    Me.cmdUpdate.Enabled = Me.Opt(1).value
    Me.Chk.Enabled = Me.Opt(1).value
    Me.lbl(1).Enabled = Me.Opt(1).value
End Sub

Private Sub TxtItemCode_KeyDown(KeyCode As Integer, _
                                Shift As Integer)
    Dim Msg As String
    Dim Lngid As Long

    If KeyCode = vbKeyReturn Then
        If Trim(Me.TxtItemCode.text) = "" Then Exit Sub
        Lngid = GetItemID(Trim(Me.TxtItemCode.text))

        If Lngid <> 0 Then
            DcboItemName.BoundText = Lngid
        Else
            DcboItemName.BoundText = ""
            Msg = "·«ÌÊÃœ ’‰› „”Ã· »Â–« «·ﬂÊœ..!"
            MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
        End If
    End If

End Sub

Private Function GetTransIDs() As String
    Dim StrTemp As String
    Dim i As Integer

    With Me.FG

        For i = FG.FixedRows To FG.Rows - 1

            If Trim(.TextMatrix(i, .ColIndex("TranseNum"))) <> "" Then
                StrTemp = StrTemp & Trim(.TextMatrix(i, .ColIndex("TranseNum"))) & ","
            End If

        Next i

    End With

    If StrTemp <> "" Then
        StrTemp = Mid(StrTemp, 1, Len(StrTemp) - 1)
    End If

    GetTransIDs = StrTemp
End Function

Private Sub XPTxtCode_KeyPress(KeyAscii As Integer)

    If KeyAscii = vbKeyReturn Then
        KeyAscii = 0
        Cmd_Click (0)
    End If

End Sub

Private Sub ChangeLang()
    Me.Caption = "Search For Item Serail"
    Me.Ele(0).Caption = "Operations On Spare Parts"
    Me.Ele(3).Caption = "Choose your Search Type"
    Opt(1).Caption = "Search for an Item Serial"
    Opt(0).Caption = "Search for an All Item Serials"
    lbl(0).Caption = "Serial"
    cmdUpdate.Caption = "Update this Serial"
    Chk.Caption = "Partial Search"
    lbl(1).Caption = "Choose the (Partial Search)to Found All the Similar Serials"
    lbl(2).Caption = "Write Item Code then Press Enter"
    lbl(4).Caption = "Item Code"
    lbl(3).Caption = "Item Name"
    LblRemark.Caption = "Double Click On any Transaction to Display"
    CmdPrintReport.Caption = "Report"

    With Me.FG
        .TextMatrix(0, .ColIndex("NumIndex")) = "S"
        .TextMatrix(0, .ColIndex("TranseNum")) = "Transaction ID"
        .TextMatrix(0, .ColIndex("Transaction_Serial")) = "Transaction Serial"
        .TextMatrix(0, .ColIndex("TransDate")) = "Transaction Date"
        .TextMatrix(0, .ColIndex("TransType")) = "Transaction Type"
        .TextMatrix(0, .ColIndex("SerialNumber")) = "Serial Number"
        .TextMatrix(0, .ColIndex("Client")) = "Customer&&Supplier Name"
        .TextMatrix(0, .ColIndex("ItemPlace")) = "Store Name"
    End With

    Cmd(0).Caption = "&Search"
    Cmd(1).Caption = "&Clear"
    Cmd(2).Caption = "E&xit"
    CmdHelp.Caption = "&Help"

End Sub

