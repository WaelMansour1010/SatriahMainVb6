VERSION 5.00
Object = "{0AFE7BE0-11B7-4A3E-978D-D4501E9A57FE}#1.0#0"; "c1sizer.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{FE5DCFAD-BC1D-11D2-94CF-004005455FAA}#1.4#0"; "ImpulseButton.ocx"
Begin VB.Form FrmSchooleFile 
   BackColor       =   &H00E2E9E9&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "„·› «·„œ«—” "
   ClientHeight    =   9420
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   12495
   Icon            =   "frmSchooleFile.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   RightToLeft     =   -1  'True
   ScaleHeight     =   9420
   ScaleWidth      =   12495
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   8
      Charset         =   178
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Begin C1SizerLibCtl.C1Elastic C1Elastic1 
      Height          =   9420
      Left            =   0
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   0
      Width           =   12495
      _cx             =   22040
      _cy             =   16616
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
      Begin C1SizerLibCtl.C1Elastic C1Elastic6 
         Height          =   492
         Left            =   120
         TabIndex        =   41
         TabStop         =   0   'False
         Top             =   8076
         Width           =   5772
         _cx             =   10186
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
         Begin VB.Label XPTxtCurrent 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Height          =   312
            Left            =   2940
            RightToLeft     =   -1  'True
            TabIndex        =   45
            Top             =   120
            Width           =   828
         End
         Begin VB.Label XPTxtCount 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Height          =   312
            Left            =   120
            RightToLeft     =   -1  'True
            TabIndex        =   44
            Top             =   120
            Width           =   660
         End
         Begin VB.Label Lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   " «·”Ã· «·Õ«·Ì:"
            Height          =   312
            Index           =   2
            Left            =   3816
            RightToLeft     =   -1  'True
            TabIndex        =   43
            Top             =   120
            Width           =   1104
         End
         Begin VB.Label Lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   " ⁄œœ «·”Ã·« :"
            Height          =   312
            Index           =   4
            Left            =   816
            RightToLeft     =   -1  'True
            TabIndex        =   42
            Top             =   120
            Width           =   1104
         End
      End
      Begin C1SizerLibCtl.C1Elastic C1Elastic4 
         Height          =   2670
         Left            =   6915
         TabIndex        =   33
         TabStop         =   0   'False
         Top             =   3120
         Width           =   5460
         _cx             =   9631
         _cy             =   4710
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
         Begin VB.TextBox txtMotName 
            Alignment       =   1  'Right Justify
            Height          =   300
            Left            =   150
            MaxLength       =   50
            RightToLeft     =   -1  'True
            TabIndex        =   60
            Top             =   2340
            Width           =   3900
         End
         Begin VB.TextBox txtTel 
            Alignment       =   1  'Right Justify
            Height          =   300
            Left            =   120
            MaxLength       =   50
            RightToLeft     =   -1  'True
            TabIndex        =   10
            Top             =   1332
            Width           =   3924
         End
         Begin VB.ComboBox cbSchooleType 
            Height          =   288
            Left            =   120
            RightToLeft     =   -1  'True
            TabIndex        =   9
            Top             =   960
            Width           =   3924
         End
         Begin VB.TextBox txtSupervisorName 
            Alignment       =   1  'Right Justify
            Height          =   300
            Left            =   120
            MaxLength       =   50
            RightToLeft     =   -1  'True
            TabIndex        =   8
            Top             =   -12
            Visible         =   0   'False
            Width           =   1620
         End
         Begin VB.TextBox txtSupervisorCode 
            Alignment       =   1  'Right Justify
            Height          =   300
            Left            =   2904
            MaxLength       =   50
            RightToLeft     =   -1  'True
            TabIndex        =   7
            Top             =   -12
            Visible         =   0   'False
            Width           =   1140
         End
         Begin VB.TextBox txtManager 
            Alignment       =   1  'Right Justify
            Height          =   300
            Left            =   120
            MaxLength       =   50
            RightToLeft     =   -1  'True
            TabIndex        =   6
            Top             =   252
            Width           =   3924
         End
         Begin MSDataListLib.DataCombo dcCity 
            Height          =   288
            Left            =   120
            TabIndex        =   11
            Top             =   1644
            Width           =   3924
            _ExtentX        =   6932
            _ExtentY        =   556
            _Version        =   393216
            BackColor       =   16777215
            Text            =   ""
            RightToLeft     =   -1  'True
         End
         Begin MSDataListLib.DataCombo dcManagerialID 
            Height          =   288
            Left            =   120
            TabIndex        =   12
            Top             =   2052
            Width           =   3924
            _ExtentX        =   6932
            _ExtentY        =   556
            _Version        =   393216
            BackColor       =   16777215
            Text            =   ""
            RightToLeft     =   -1  'True
         End
         Begin MSDataListLib.DataCombo DcEmployee 
            Height          =   288
            Left            =   120
            TabIndex        =   57
            Top             =   600
            Width           =   3912
            _ExtentX        =   6906
            _ExtentY        =   556
            _Version        =   393216
            Enabled         =   0   'False
            BackColor       =   16777215
            Text            =   ""
            RightToLeft     =   -1  'True
         End
         Begin VB.Label Label8 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "«·„ ⁄Âœ"
            ForeColor       =   &H00000000&
            Height          =   390
            Left            =   4080
            TabIndex        =   61
            Top             =   2340
            Width           =   1245
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "«·„‰ÿﬁ… «· ⁄·Ì„Ì…  "
            ForeColor       =   &H00000000&
            Height          =   360
            Left            =   4200
            TabIndex        =   40
            Top             =   2055
            Width           =   1245
         End
         Begin VB.Label Label2 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "«·„Õ«›Ÿ…"
            ForeColor       =   &H00000000&
            Height          =   288
            Left            =   4080
            TabIndex        =   39
            Top             =   1644
            Width           =   1248
         End
         Begin VB.Label Lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "«· ·Ì›Ê‰"
            Height          =   180
            Index           =   10
            Left            =   4164
            RightToLeft     =   -1  'True
            TabIndex        =   38
            Top             =   1332
            Width           =   1164
         End
         Begin VB.Label Lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "‰Ê⁄ «·„œ—”…"
            Height          =   324
            Index           =   1
            Left            =   4164
            RightToLeft     =   -1  'True
            TabIndex        =   37
            Top             =   960
            Width           =   1164
         End
         Begin VB.Label Label6 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "«”„ «·„‘—›"
            ForeColor       =   &H00000000&
            Height          =   384
            Left            =   4152
            TabIndex        =   36
            Top             =   588
            Width           =   1128
         End
         Begin VB.Label Label4 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "ﬂÊœ «·„‘—›"
            ForeColor       =   &H00000000&
            Height          =   384
            Left            =   4080
            TabIndex        =   35
            Top             =   -12
            Visible         =   0   'False
            Width           =   1248
         End
         Begin VB.Label Label3 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "„œÌ— «·„œ—”…"
            ForeColor       =   &H00000000&
            Height          =   468
            Left            =   4080
            TabIndex        =   34
            Top             =   252
            Width           =   1248
         End
      End
      Begin C1SizerLibCtl.C1Elastic C1Elastic2 
         Height          =   2550
         Left            =   6915
         TabIndex        =   24
         TabStop         =   0   'False
         Top             =   5865
         Width           =   5460
         _cx             =   9631
         _cy             =   4498
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
         Begin VB.TextBox txtEducationType 
            Alignment       =   1  'Right Justify
            Height          =   420
            Left            =   180
            MaxLength       =   50
            RightToLeft     =   -1  'True
            TabIndex        =   58
            Top             =   2130
            Width           =   3924
         End
         Begin VB.ComboBox Peroid 
            Height          =   315
            ItemData        =   "frmSchooleFile.frx":038A
            Left            =   120
            List            =   "frmSchooleFile.frx":038C
            RightToLeft     =   -1  'True
            TabIndex        =   16
            Top             =   1785
            Width           =   3948
         End
         Begin VB.ComboBox cbSexType 
            Height          =   315
            ItemData        =   "frmSchooleFile.frx":038E
            Left            =   120
            List            =   "frmSchooleFile.frx":0390
            RightToLeft     =   -1  'True
            TabIndex        =   13
            Top             =   375
            Width           =   3948
         End
         Begin VB.TextBox txtMoney 
            Alignment       =   1  'Right Justify
            Height          =   420
            Left            =   120
            MaxLength       =   50
            RightToLeft     =   -1  'True
            TabIndex        =   15
            Top             =   1245
            Width           =   3924
         End
         Begin VB.TextBox txtCount 
            Alignment       =   1  'Right Justify
            Height          =   420
            Left            =   120
            MaxLength       =   50
            RightToLeft     =   -1  'True
            TabIndex        =   14
            Top             =   825
            Width           =   3924
         End
         Begin VB.Label Lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "‰Ê⁄ «· ⁄·Ì„"
            Height          =   315
            Index           =   11
            Left            =   4200
            RightToLeft     =   -1  'True
            TabIndex        =   59
            Top             =   2100
            Width           =   1170
         End
         Begin VB.Label Label7 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "«·„—Õ·…"
            ForeColor       =   &H00000000&
            Height          =   510
            Left            =   4050
            TabIndex        =   47
            Top             =   1785
            Width           =   1245
         End
         Begin VB.Label Label5 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Ã‰” «·ÿ·»Â"
            ForeColor       =   &H00000000&
            Height          =   360
            Left            =   4050
            TabIndex        =   27
            Top             =   375
            Width           =   1245
         End
         Begin VB.Label Lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "⁄œœ «·ÿ·«»"
            Height          =   315
            Index           =   5
            Left            =   4170
            RightToLeft     =   -1  'True
            TabIndex        =   26
            Top             =   825
            Width           =   1170
         End
         Begin VB.Label Lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "„Œ’’ «·ÿ«·»"
            Height          =   360
            Index           =   6
            Left            =   4170
            RightToLeft     =   -1  'True
            TabIndex        =   25
            Top             =   1245
            Width           =   1170
         End
      End
      Begin C1SizerLibCtl.C1Elastic EleHeader 
         Height          =   672
         Left            =   0
         TabIndex        =   18
         TabStop         =   0   'False
         Top             =   0
         Width           =   12552
         _cx             =   22146
         _cy             =   1191
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial (Arabic)"
            Size            =   22.5
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
         Caption         =   "„·› «·„œ«—”     "
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
         PicturePos      =   4
         CaptionStyle    =   1
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
            Left            =   2250
            RightToLeft     =   -1  'True
            TabIndex        =   19
            Top             =   180
            Visible         =   0   'False
            Width           =   855
         End
         Begin ImpulseButton.ISButton XPBtnMove 
            Height          =   345
            Index           =   0
            Left            =   1155
            TabIndex        =   20
            Top             =   120
            Width           =   495
            _ExtentX        =   873
            _ExtentY        =   609
            ButtonStyle     =   1
            ButtonPositionImage=   4
            Caption         =   ""
            BackColor       =   -2147483634
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   178
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ButtonImage     =   "frmSchooleFile.frx":0392
            ColorButton     =   -2147483634
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
            Left            =   90
            TabIndex        =   21
            Top             =   120
            Width           =   495
            _ExtentX        =   873
            _ExtentY        =   609
            ButtonStyle     =   1
            ButtonPositionImage=   4
            Caption         =   ""
            BackColor       =   -2147483634
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   178
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ButtonImage     =   "frmSchooleFile.frx":072C
            ColorButton     =   -2147483634
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
            TabIndex        =   22
            Top             =   120
            Width           =   495
            _ExtentX        =   873
            _ExtentY        =   609
            ButtonStyle     =   1
            ButtonPositionImage=   4
            Caption         =   ""
            BackColor       =   -2147483634
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   178
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ButtonImage     =   "frmSchooleFile.frx":0AC6
            ColorButton     =   -2147483634
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
            Left            =   615
            TabIndex        =   23
            Top             =   120
            Width           =   495
            _ExtentX        =   873
            _ExtentY        =   609
            ButtonStyle     =   1
            ButtonPositionImage=   4
            Caption         =   ""
            BackColor       =   -2147483634
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   178
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ButtonImage     =   "frmSchooleFile.frx":0E60
            ColorButton     =   -2147483634
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
      Begin C1SizerLibCtl.C1Elastic C1Elastic3 
         Height          =   2220
         Left            =   5040
         TabIndex        =   28
         TabStop         =   0   'False
         Top             =   840
         Width           =   7380
         _cx             =   13018
         _cy             =   3916
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
         Begin VB.TextBox txtLastContract 
            Alignment       =   1  'Right Justify
            Height          =   315
            Left            =   120
            MaxLength       =   50
            RightToLeft     =   -1  'True
            TabIndex        =   5
            Top             =   45
            Width           =   3570
         End
         Begin VB.TextBox txtMinisterNo 
            Alignment       =   1  'Right Justify
            Height          =   315
            Left            =   2040
            MaxLength       =   50
            RightToLeft     =   -1  'True
            TabIndex        =   2
            Top             =   450
            Width           =   3924
         End
         Begin VB.TextBox txtNameE 
            Alignment       =   1  'Right Justify
            Height          =   555
            Left            =   120
            MaxLength       =   255
            MultiLine       =   -1  'True
            RightToLeft     =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   4
            Top             =   1605
            Width           =   5850
         End
         Begin VB.TextBox txtName 
            Alignment       =   1  'Right Justify
            Height          =   675
            Left            =   120
            MaxLength       =   255
            MultiLine       =   -1  'True
            RightToLeft     =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   3
            Top             =   840
            Width           =   5850
         End
         Begin VB.TextBox txtID 
            Alignment       =   1  'Right Justify
            Height          =   315
            Left            =   4830
            Locked          =   -1  'True
            MaxLength       =   10
            RightToLeft     =   -1  'True
            TabIndex        =   1
            Top             =   120
            Width           =   1140
         End
         Begin VB.Label Lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "—ﬁ„ «Œ— ⁄ﬁœ"
            Height          =   315
            Index           =   9
            Left            =   3570
            RightToLeft     =   -1  'True
            TabIndex        =   46
            Top             =   165
            Width           =   1170
         End
         Begin VB.Label Lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "„”·”·"
            Height          =   315
            Index           =   8
            Left            =   5970
            RightToLeft     =   -1  'True
            TabIndex        =   32
            Top             =   120
            Width           =   1170
         End
         Begin VB.Label Lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "«·«”„ «‰Ã·Ì“Ì"
            Height          =   315
            Index           =   7
            Left            =   6090
            RightToLeft     =   -1  'True
            TabIndex        =   31
            Top             =   1605
            Width           =   1170
         End
         Begin VB.Label Lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "—ﬁ„ Ê“«—Ï"
            Height          =   375
            Index           =   0
            Left            =   6045
            RightToLeft     =   -1  'True
            TabIndex        =   30
            Top             =   450
            Width           =   1200
         End
         Begin VB.Label Lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "«·«”„ ⁄—»Ì"
            Height          =   315
            Index           =   3
            Left            =   6090
            RightToLeft     =   -1  'True
            TabIndex        =   29
            Top             =   810
            Width           =   1170
         End
      End
      Begin C1SizerLibCtl.C1Elastic C1Elastic7 
         Height          =   696
         Left            =   120
         TabIndex        =   17
         TabStop         =   0   'False
         Top             =   8640
         Width           =   12204
         _cx             =   21537
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
         Begin ImpulseButton.ISButton Cmd 
            Height          =   456
            Index           =   0
            Left            =   10668
            TabIndex        =   48
            Top             =   120
            Width           =   1308
            _ExtentX        =   2302
            _ExtentY        =   794
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
            ButtonImage     =   "frmSchooleFile.frx":11FA
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
            Height          =   456
            Index           =   1
            Left            =   9276
            TabIndex        =   49
            Top             =   120
            Width           =   1344
            _ExtentX        =   2381
            _ExtentY        =   794
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
            ButtonImage     =   "frmSchooleFile.frx":7A5C
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
            Height          =   456
            Index           =   2
            Left            =   7920
            TabIndex        =   50
            Top             =   120
            Width           =   1332
            _ExtentX        =   2355
            _ExtentY        =   794
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
            ButtonImage     =   "frmSchooleFile.frx":E2BE
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
            Height          =   456
            Index           =   3
            Left            =   6588
            TabIndex        =   51
            Top             =   120
            Width           =   1272
            _ExtentX        =   2249
            _ExtentY        =   794
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
            ButtonImage     =   "frmSchooleFile.frx":14B20
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
            Height          =   456
            Index           =   4
            Left            =   5232
            TabIndex        =   52
            Top             =   120
            Width           =   1284
            _ExtentX        =   2275
            _ExtentY        =   794
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
            ButtonImage     =   "frmSchooleFile.frx":1B382
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
            Height          =   456
            Index           =   6
            Left            =   1452
            TabIndex        =   53
            Top             =   120
            Width           =   1212
            _ExtentX        =   2143
            _ExtentY        =   794
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
            ButtonImage     =   "frmSchooleFile.frx":21BE4
            ColorButton     =   14871017
            ColorHighlight  =   16777215
            ColorHoverText  =   16711680
            ColorShadow     =   -2147483637
            ColorOutline    =   0
            DrawFocusRectangle=   0   'False
            ColorToggledHoverText=   16711680
            ColorTextShadow =   -2147483637
         End
         Begin ImpulseButton.ISButton CmdAttach 
            Height          =   456
            Left            =   120
            TabIndex        =   54
            Top             =   120
            Width           =   1284
            _ExtentX        =   2275
            _ExtentY        =   794
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
            ButtonImage     =   "frmSchooleFile.frx":4B806
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
            Height          =   456
            Index           =   7
            Left            =   4020
            TabIndex        =   55
            Top             =   120
            Width           =   1188
            _ExtentX        =   2090
            _ExtentY        =   794
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
            ButtonImage     =   "frmSchooleFile.frx":52068
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
            Height          =   456
            Index           =   9
            Left            =   2688
            TabIndex        =   56
            Top             =   120
            Width           =   1332
            _ExtentX        =   2355
            _ExtentY        =   794
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
            ButtonImage     =   "frmSchooleFile.frx":588CA
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
      Begin VB.Image Image2 
         Height          =   6720
         Left            =   0
         Picture         =   "frmSchooleFile.frx":5F12C
         Stretch         =   -1  'True
         Top             =   690
         Width           =   4935
      End
   End
End
Attribute VB_Name = "FrmSchooleFile"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rs As ADODB.Recordset
Dim Rs_Temp As ADODB.Recordset
Dim TTP As clstooltip

Private Sub Cmd_Click(Index As Integer)
'    On Error GoTo ErrTrap
    Select Case Index
        Case 0

            If DoPremis(Do_New, Me.Name, True) = False Then
                Exit Sub
            End If

            TxtModFlg.Text = "N"
            clear_all Me
            txtID.Text = CStr(new_id("TblSchooleFile", "ID", "", True))
            txtName.SetFocus
'
        Case 1

            If DoPremis(Do_Edit, Me.Name, True) = False Then
                Exit Sub
            End If

            TxtModFlg.Text = "E"

        Case 2

            SaveData

        Case 3
            Undo

        Case 4

            If DoPremis(Do_Delete, Me.Name, True) = False Then
                Exit Sub
            End If

            Del_Company

        Case 5

        Case 6
            Unload Me
         Case 7
         'print_report2
         Case 9
         Unload FrmSearch_BasicData
         FrmSearch_BasicData.SendForm = "schoolfile"
         FrmSearch_BasicData.show
         
    End Select

    Exit Sub
ErrTrap:
End Sub

Private Sub CmdHelp_Click()
    SystemOptions.SysHelp.HHTopicID = Me.HelpContextID
    SystemOptions.SysHelp.HHDisplayTopicID Me.hwnd
End Sub

Private Sub dcCity_Change()
If Me.TxtModFlg = "R" Then Exit Sub
Dim str As String
Set Rs_Temp = New ADODB.Recordset
Set dcManagerialID.RowSource = Rs_Temp
If SystemOptions.UserInterface = ArabicInterface Then
    str = " Select ID , Name   from TblManagerialArea  where cityid = " & val(dcCity.BoundText)
Else
    str = " Select ID , NameE   from TblManagerialArea  where cityid = " & val(dcCity.BoundText)
End If
    fill_combo dcManagerialID, str
dcManagerialID.Refresh


End Sub

Private Sub Form_Activate()
'    txtid.SetFocus
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, _
                         Shift As Integer)
    On Error GoTo ErrTrap

    If KeyCode = vbKeyReturn Then
        If Me.TxtModFlg.Text = "R" Then
            Cmd_Click (0)
        Else
            SendKeys "{TAB}"
        End If
    End If

    If Me.TxtModFlg.Text = "R" Then
        If KeyCode = vbKeyDown Or KeyCode = vbKeyEnd Then
            XPBtnMove_Click (2)
        ElseIf KeyCode = vbKeyUp Or KeyCode = vbKeyHome Then
            XPBtnMove_Click (1)
        ElseIf KeyCode = vbKeyRight Or KeyCode = vbKeyPageDown Then
            XPBtnMove_Click (3)
        ElseIf KeyCode = vbKeyLeft Or KeyCode = vbKeyPageUp Then
            XPBtnMove_Click (0)
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

    If Shift = 2 Then
        If KeyCode = vbKeyX Then
            If Cmd(6).Enabled = False Then Exit Sub
            Cmd_Click (6)
        End If
    End If

    Exit Sub
ErrTrap:
End Sub

Private Sub Form_Load()
 '   On Error GoTo ErrTrap
    Dim Dcombos As ClsDataCombos

    Set Dcombos = New ClsDataCombos
    
    Dcombos.getCountriesGovernments Me.dcCity
    
    Dim StrSQL As String
    If SystemOptions.UserInterface = ArabicInterface Then
    StrSQL = "Select  ID , Name  from TblManagerialArea "
    Else
    StrSQL = "Select  ID , Namee  from TblManagerialArea "
    End If
   fill_combo dcManagerialID, StrSQL
   
 With cbSchooleType
 
 If SystemOptions.UserInterface = EnglishInterface Then
 .Clear
 .AddItem ("Governmental")
 .AddItem ("Domestic School")
.AddItem ("International")
 Else
  .Clear
 .AddItem ("ÕﬂÊ„Ï")
 .AddItem ("√Â·Ï")
 .AddItem ("«‰ —‰«‘Ê‰«·")
 End If
 End With
 
 
 With cbSexType
 If SystemOptions.UserInterface = EnglishInterface Then
 .Clear
 .AddItem ("Male")
 .AddItem ("Female")
 Else
  .Clear
 .AddItem ("»‰Ì‰")
 .AddItem ("»‰« ")
 End If
 End With


 With Peroid
 If SystemOptions.UserInterface = EnglishInterface Then
 .Clear
 .AddItem ("Basic")
 .AddItem ("intermediate")
 .AddItem ("Secondery")
 .AddItem ("Compound")
 Else
  .Clear
 
    
 .AddItem ("«» œ«∆Ì")
 .AddItem ("„ Ê”ÿ")
 .AddItem ("À«‰ÊÌ")
 .AddItem ("„Ã„⁄")
 End If
 End With
 
 



    If SystemOptions.UserInterface = EnglishInterface Then
        SetInterface Me
        ChangeLang
    End If

    LogTextA = "   «·œŒÊ· «·Ì ‘«‘… " & "  „·› «·„œ«—”  "
    LogTexte = " Open Window " & "  Boxes Data "
    AddToLogFile CInt(user_id), 0, Date, Time, LogTextA, LogTexte, Me.Name, "O", "", ""

    

    Set Cmd(0).ButtonImage = mdifrmmain.ImgLstTree.ListImages("New").Picture
    Set Cmd(1).ButtonImage = mdifrmmain.ImgLstTree.ListImages("Edit").Picture
    Set Cmd(2).ButtonImage = mdifrmmain.ImgLstTree.ListImages("save").Picture
    Set Cmd(3).ButtonImage = mdifrmmain.ImgLstTree.ListImages("Undo").Picture
    Set Cmd(4).ButtonImage = mdifrmmain.ImgLstTree.ListImages("Del").Picture
    Set Cmd(6).ButtonImage = mdifrmmain.ImgLstTree.ListImages("Exit").Picture
 '   Set CmdHelp.ButtonImage = mdifrmmain.ImgLstTree.ListImages("Help").Picture
    Resize_Form Me
    AddTip
    Set rs = New ADODB.Recordset
    
    StrSQL = ""
     If SystemOptions.usertype <> UserAdminAll Then
      
StrSQL = "SELECT  *  From TblSchooleFile    "
  Else
 StrSQL = "SELECT  *  From TblSchooleFile"
    End If
  rs.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText
      
      
     Dim str As String
     If SystemOptions.UserInterface = ArabicInterface Then
            str = "   select  emp_id , emp_name  from tblEmployee  "
    Else
            str = "   select  emp_id , emp_nameE  from tblEmployee  "
    End If
    fill_combo DcEmployee, str
        

        
    Me.TxtModFlg.Text = "R"
    XPBtnMove_Click 2

    If OPEN_NEW_SCREEN = True Then
        Cmd_Click (0)
    End If

    Exit Sub

ErrTrap:
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, _
                             UnloadMode As Integer)
    Dim IntResult As String
    Dim StrMSG As String
    On Error GoTo ErrTrap

    If Me.TxtModFlg.Text <> "R" Then

        Select Case Me.TxtModFlg.Text

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

Private Sub ChangeLang()
    Dim XPic As IPictureDisp

    Set XPic = Me.XPBtnMove(1).ButtonImage
    Set Me.XPBtnMove(1).ButtonImage = Me.XPBtnMove(2).ButtonImage
    Set Me.XPBtnMove(2).ButtonImage = XPic

    Set XPic = Me.XPBtnMove(0).ButtonImage
    Set Me.XPBtnMove(0).ButtonImage = Me.XPBtnMove(3).ButtonImage
    Set Me.XPBtnMove(3).ButtonImage = XPic
    
   
    Lbl(7).Caption = " Name En"
    Lbl(3).Caption = " Name Ar"
    Lbl(8).Caption = "Process No"
    Lbl(0).Caption = "Minister No."
    Label3.Caption = "School Manager"
    Label1.Caption = "Managerial Area"
    Label2.Caption = "City"
    Lbl(5).Caption = "Student Count"
    Lbl(6).Caption = "Custom"
    Lbl(1).Caption = "School Type"
    Lbl(10).Caption = "Telephone"
    Label4.Caption = "Supervisor Code"
    Label6.Caption = "Supervisor"
    Label5.Caption = "Student Gender"
    Me.Caption = "School Data"
    EleHeader.Caption = Me.Caption
    
    Lbl(2).Caption = "Current Record"
    Lbl(4).Caption = "NO. Recordes"

    Me.Cmd(0).Caption = "New"
    Me.Cmd(1).Caption = "Edit"
    Me.Cmd(2).Caption = "Save"
    Me.Cmd(3).Caption = "Undo"
    Me.Cmd(4).Caption = "Delete"
    'Me.Cmd(5).Caption = "Search"
    Me.Cmd(6).Caption = "Exit"
    Me.Cmd(7).Caption = "Print"
   CmdAttach.Caption = "Attachment"

Lbl(9).Caption = "Last Contract"



End Sub

Private Sub Form_Unload(Cancel As Integer)
    On Error GoTo ErrTrap
    LogTextA = "     «·Œ—ÊÃ „‰ ‘«‘… " & "  »Ì«‰«  «·„œ—”…   "
    LogTexte = " Exit Window " & "  Boxes Data "
    AddToLogFile CInt(user_id), 0, Date, Time, LogTextA, LogTexte, Me.Name, "O", "", ""

    If rs.State = adStateOpen Then
        If Not (rs.EOF Or rs.BOF) Then
            If rs.EditMode <> adEditNone Then
                rs.CancelUpdate
            End If
        End If

        rs.Close
    End If

    Set rs = Nothing
    Set TTP = Nothing
    Exit Sub
ErrTrap:
End Sub


Private Sub txtCount_Change()
If Not IsNumeric(txtCount.Text) Then
txtCount.Text = ""
End If
End Sub

Private Sub TxtModFlg_Change()
    On Error GoTo ErrTrap

    Select Case Me.TxtModFlg.Text

        Case "R"

            If SystemOptions.UserInterface = ArabicInterface Then
                Me.Caption = "»Ì«‰«  «·„œ—”…"
            Else
                Me.Caption = "School  Data"
            End If

            Me.Cmd(2).Enabled = False
            Me.Cmd(3).Enabled = False
        
            Me.Cmd(0).Enabled = True
            Me.Cmd(1).Enabled = True
            Me.Cmd(4).Enabled = True
        Me.Cmd(9).Enabled = True
        
            Me.XPBtnMove(0).Enabled = True
            Me.XPBtnMove(1).Enabled = True
            Me.XPBtnMove(2).Enabled = True
            Me.XPBtnMove(3).Enabled = True
        
            Me.txtID.locked = True
            Me.txtName.locked = True
        '    Me.XPMTxtRemark.locked = True

            If rs.RecordCount < 1 Then
                Me.XPBtnMove(0).Enabled = False
                Me.XPBtnMove(1).Enabled = False
                Me.XPBtnMove(2).Enabled = False
                Me.XPBtnMove(3).Enabled = False
                Me.Cmd(1).Enabled = False
                Me.Cmd(4).Enabled = False
            End If
            
            C1Elastic2.Enabled = False
            C1Elastic3.Enabled = False
            C1Elastic4.Enabled = False
            
        Case "N"

            If SystemOptions.UserInterface = ArabicInterface Then
                Me.Caption = "»Ì«‰«  «·„œ—”… ( ÃœÌœ )"
            Else
                Me.Caption = "ÚSchool Data(New)"
            End If
        
            Me.Cmd(2).Enabled = True
            Me.Cmd(3).Enabled = True

            If SystemOptions.UserInterface = ArabicInterface Then
                Me.Caption = "»Ì«‰«  «·„œ—”…( ÃœÌœ )"
            Else
                Me.Caption = "ÚSchool Data(New)"
            End If
        
            Me.Cmd(0).Enabled = False
            Me.Cmd(1).Enabled = False
            Me.Cmd(4).Enabled = False
        Me.Cmd(9).Enabled = False
            '       Me.XPBtnMove(0).Enabled = False
            '       Me.XPBtnMove(1).Enabled = False
            '       Me.XPBtnMove(2).Enabled = False
            '       Me.XPBtnMove(3).Enabled = False
        
            Me.txtID.locked = True
            Me.txtName.locked = False
            'Me.XPMTxtRemark.locked = False
            
               C1Elastic2.Enabled = True
            C1Elastic3.Enabled = True
            C1Elastic4.Enabled = True
        Case "E"

            If SystemOptions.UserInterface = ArabicInterface Then
                Me.Caption = "»Ì«‰«   «·„œ—”… (  ⁄œÌ· )"
            Else
                Me.Caption = "School Data(Edit)"
            End If
        
            Me.Cmd(2).Enabled = True
            Me.Cmd(3).Enabled = True
        
            Me.Cmd(0).Enabled = False
            Me.Cmd(1).Enabled = False
            Me.Cmd(4).Enabled = False
            Me.Cmd(9).Enabled = False
            
            Me.XPBtnMove(0).Enabled = False
            Me.XPBtnMove(1).Enabled = False
            Me.XPBtnMove(2).Enabled = False
            Me.XPBtnMove(3).Enabled = False
        
            Me.txtID.locked = True
            Me.txtName.locked = False
            'Me.XPMTxtRemark.locked = False
            
            C1Elastic2.Enabled = True
            C1Elastic3.Enabled = True
            C1Elastic4.Enabled = True
    End Select

    Exit Sub
ErrTrap:
End Sub
Function print_report2(Optional NoteSerial As String)
    
     
    Dim MySQL As String
    Dim RsData As New ADODB.Recordset
    Dim xApp As New CRAXDRT.Application
    Dim xReport As CRAXDRT.Report
    Dim CViewer As ClsReportViewer
    Dim StrReportTitle As String
    Dim StrFileName As String
    Dim Msg As String
MySQL = "SELECT     dbo.TblBoxesData.BoxID, dbo.TblBoxesData.BoxName, dbo.TblBoxesData.Comments, dbo.TblBoxesData.Type, dbo.TblBoxesData.Account_Code, "
MySQL = MySQL & "                      dbo.TblBoxesData.empid, dbo.TblEmployee.Emp_Code, dbo.TblEmployee.Emp_Name, dbo.TblEmployee.Emp_Name1, dbo.TblEmployee.Emp_Name2,"
MySQL = MySQL & "                      dbo.TblEmployee.Emp_Name3, dbo.TblEmployee.Emp_Name4, dbo.TblEmployee.Fullcode, dbo.TblEmployee.Emp_Namee4, dbo.TblEmployee.Emp_Namee3,"
MySQL = MySQL & "                      dbo.TblEmployee.Emp_Namee2, dbo.TblEmployee.Emp_Namee1, dbo.TblEmployee.Emp_Namee, dbo.TblBoxesData.BranchId, dbo.TblBranchesData.branch_name,"
MySQL = MySQL & "                      dbo.TblBranchesData.branch_namee, dbo.TblBoxesData.BoxNameE, dbo.TblBoxesData.Account_Code1, dbo.TblBoxesData.OpenBalanceDate,"
MySQL = MySQL & "                      dbo.TblBoxesData.OpenBalanceType, dbo.TblBoxesData.OpenBalance, dbo.TblBoxesData.boxValue, dbo.TblBoxesData.Account_Code2, dbo.TblBoxesData.BTtype,"
MySQL = MySQL & "                      dbo.TblBoxesData.Driverid , dbo.TblBoxesData.opening_balance_voucher_id, dbo.TblBoxesData.ChequeBox, dbo.TblBoxesData.ParentAccount"
MySQL = MySQL & " FROM         dbo.TblBoxesData LEFT OUTER JOIN"
MySQL = MySQL & "                      dbo.TblBranchesData ON dbo.TblBoxesData.BranchId = dbo.TblBranchesData.branch_id LEFT OUTER JOIN"
MySQL = MySQL & "                      dbo.TblEmployee ON dbo.TblBoxesData.empid = dbo.TblEmployee.Emp_ID"
'MySQL = MySQL & " Where (dbo.TblBoxesData.BoxID =" & val(XPTxtBoxID.text) & ")"


 If SystemOptions.UserInterface = ArabicInterface Then
          StrFileName = App.path & "\REPORTS\REPORTS NEW\" & "RepBoxesData.rpt"
     Else
        StrFileName = App.path & "\REPORTS\REPORTS NEW\" & "RepBoxesData.rpt"
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
        ' xReport.ParameterFields(2).AddCurrentValue RPTComment_Arabic
        StrReportTitle = "" '& StrAccountName
        'If Me.DTPickerAccFrom.value <> Empty Or Me.DTPickerAccFrom.value <> Null Then
        '    StrReportTitle = StrReportTitle + " »œ«Ì… „‰ " & Format(Me.DTPickerAccFrom.value, "yyyy/M/d") & ""
        'End If
        'If Me.DTPickerAccTo.value <> Empty Or Me.DTPickerAccTo.value <> Null Then
        '    StrReportTitle = StrReportTitle + " ≈·Ï " & Format(Me.DTPickerAccTo.value, "yyyy/M/d") & " "
        'End If
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
       ' xReport.ParameterFields(4).AddCurrentValue WriteNo(Format(val(TxtAdvanceValue.text), "0.00"), 0, True, ".")
'        xReport.ParameterFields(6).AddCurrentValue val(lbl(23).Caption)
      '   xReport.ParameterFields(7).AddCurrentValue DBIssueDate.value
   ' xReport.ParameterFields(8).AddCurrentValue IIf(IsNumeric(Fg.TextMatrix(Me.Fg.FixedRows, Fg.ColIndex("PartValue"))), val(Fg.TextMatrix(Me.Fg.FixedRows, Fg.ColIndex("PartValue"))), 0)
' xReport.ParameterFields(9).AddCurrentValue val(lbl(22).Caption)
'  xReport.ParameterFields(10).AddCurrentValue val(TxtDiscount.text)
 '  xReport.ParameterFields(11).AddCurrentValue txtDiscountDES.text
   
'    xReport.ParameterFields(5).AddCurrentValue ToHijriDate(RsData("notedate").value)
    xReport.reporttitle = StrReportTitle
    xReport.EnableParameterPrompting = False
    xReport.ApplicationName = App.title
    xReport.ReportAuthor = App.title
    Set CViewer = New ClsReportViewer
    CViewer.FireReport xReport, WindowTarget, "", , , , StrFileName

    RsData.Close
    Set RsData = Nothing
    Screen.MousePointer = vbDefault


 
  
 
End Function
Public Sub Retrive(Optional Lngid As Long = 0)

    'On Error GoTo ErrTrap
    If rs.RecordCount < 1 Then
        XPTxtCurrent.Caption = 0
        XPTxtCount.Caption = 0
        Exit Sub
    End If
    
     If rs.EOF Or rs.BOF Then
        Exit Sub
    Else
        If Lngid <> 0 Then
            rs.Find "ID =" & Lngid, , adSearchForward, adBookmarkFirst

            If rs.EOF Or rs.BOF Then
                Exit Sub
            End If
        End If
    End If
   
    txtID.Text = IIf(IsNull(rs("ID").value), "", (rs("ID").value))
    txtMinisterNo.Text = IIf(IsNull(rs("MinisterNo").value), "", rs("MinisterNo").value)
    txtName.Text = IIf(IsNull(rs("Name").value), "", Trim(rs("Name").value))
    txtNameE.Text = IIf(IsNull(rs("NameE").value), "", Trim(rs("NameE").value))
    txtEducationType.Text = IIf(IsNull(rs("EducationType").value), "", Trim(rs("EducationType").value))
    
    txtMotName.Text = IIf(IsNull(rs("MotName").value), "", Trim(rs("MotName").value))
    
    
    
'    XPMTxtRemark.text = IIf(IsNull(rs("Comments").value), "", Trim(rs("Comments").value))
    txtManager.Text = IIf(IsNull(rs("ManagerID").value), "", rs("ManagerID").value)
    dcManagerialID.BoundText = IIf(IsNull(rs("ManagerialID")), Null, Trim(rs("ManagerialID")))
    dcCity.BoundText = IIf(IsNull(rs("CityID")), Null, Trim(rs("CityID")))
    txtCount.Text = IIf(IsNull(rs("StudentCount").value), "", Trim(rs("StudentCount").value))
    txtMoney.Text = IIf(IsNull(rs("custom").value), "", Trim(rs("custom").value))
    cbSchooleType.ListIndex = IIf(IsNull(rs("SchooleType").value), -1, rs("SchooleType").value)
    txtTel.Text = IIf(IsNull(rs("telephone").value), "", Trim(rs("telephone").value))
            
    cbSexType.ListIndex = IIf(IsNull(rs("SexType").value), -1, rs("SexType").value)
    txtSupervisorCode.Text = IIf(IsNull(rs("SupervisorCode").value), "", rs("SupervisorCode").value)
    txtSupervisorName.Text = IIf(IsNull(rs("SupervisorName").value), "", rs("SupervisorName").value)
    txtLastContract.Text = IIf(IsNull(rs("LastCode").value), "", rs("LastCode").value)
    Peroid.ListIndex = IIf(IsNull(rs("phase").value), -1, rs("phase").value)
    DcEmployee.BoundText = IIf(IsNull(rs("Emp_ID").value), "", rs("Emp_ID").value)
    
    
    XPTxtCurrent.Caption = rs.AbsolutePosition
    XPTxtCount.Caption = rs.RecordCount
    Exit Sub
ErrTrap:
End Sub


Private Sub txtMoney_Change()
If Not IsNumeric(txtCount.Text) Then
txtMoney.Text = ""
End If
End Sub

Private Sub TxtName_GotFocus()
    SwitchKeyboardLang LANG_ARABIC
End Sub


Private Sub txtNameE_GotFocus()
    SwitchKeyboardLang LANG_ENGLISH
End Sub

Private Sub XPBtnMove_Click(Index As Integer)

    'On Error GoTo ErrTrap
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

Function CuurentLogdata(Optional Currentmode As String)
 

End Function
 
Private Sub SaveData()
    Dim Msg As String
    Dim StrSQL As String
    Dim RsTemp As New ADODB.Recordset
    Dim RsTempM As New ADODB.Recordset
    Dim BeginTrans As Boolean
   ' On Error GoTo ErrTrap

    If Me.TxtModFlg.Text <> "R" Then
    
        If Trim(dcManagerialID.BoundText) = "" Then
            If SystemOptions.UserInterface = EnglishInterface Then
                Msg = "Specify Managerial Area"
            Else
                Msg = "Õœœ «·„‰ÿﬁ… «· ⁄·Ì„Ì… "
            End If

            MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
            dcManagerialID.SetFocus
   '         SendKeys "{F4}"
            Screen.MousePointer = vbDefault
            Exit Sub
        End If
    
        If Trim(txtMinisterNo.Text) = "" Then
            If SystemOptions.UserInterface = EnglishInterface Then
                Msg = "Enter Ministery No."
            Else
                Msg = "«œŒ· «·—ﬁ„ «·Ê“«—Ï «Ê·« . "
            End If

            MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
            txtMinisterNo.SetFocus
            Screen.MousePointer = vbDefault
            Exit Sub
        End If
    
    
        If txtName.Text = "" Then
            MsgBox "„‰ ›÷·ﬂ √œŒ· «”„ «·„œ—”…", vbOKOnly + vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
            txtName.SetFocus
            Exit Sub
        End If

        Select Case Me.TxtModFlg.Text

            Case "N"
                StrSQL = " select * From  TblSchooleFile where  Name='" & Trim(txtName.Text) & "'"
                Set RsTemp = New ADODB.Recordset
                RsTemp.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText
                If RsTemp.RecordCount > 0 Then
                    Msg = "Â‰«ﬂ „œ—”…  „”Ã· „”»ﬁ« »Â–« «·«”„" & CHR(13)
                    Msg = Msg + "»—Ã«¡ «· √ﬂœ „‰ «·«”„ «·’ÕÌÕ " & CHR(13)
                    Msg = Msg + "√Ê  €ÌÌ— √Ê  „ÌÌ“ «”„ „œ—”…"
                    MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
                    txtName.SetFocus
                    Exit Sub
                End If
                RsTemp.Close
                StrSQL = " select * From  TblSchooleFile where  ministerNo ='" & Trim(txtMinisterNo.Text) & "'"
                Set RsTemp = New ADODB.Recordset
                RsTemp.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText
                If RsTemp.RecordCount > 0 Then
                    Msg = "Â‰«ﬂ „œ—”…  „”Ã· „”»ﬁ« »Â–« «·—ﬁ„ «·Ê“«—Ï " & CHR(13)
                    Msg = Msg + "»—Ã«¡ «· √ﬂœ „‰ «·—ﬁ„ «·’ÕÌÕ " & CHR(13)
                 
                    MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
                    txtName.SetFocus
                    Exit Sub
                End If

            Case "E"
                StrSQL = "select * From  TblSchooleFile  where Name='" & Trim(txtName.Text) & "'"
                RsTemp.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText
                If RsTemp.RecordCount > 0 Then
                    If RsTemp("ID").value <> val(txtID.Text) Then
                        Msg = "Â‰«ﬂ „œ—”…  „”Ã· „”»ﬁ« »Â–« «·«”„" & CHR(13)
                        Msg = Msg + "»—Ã«¡ «· √ﬂœ „‰ «·«”„ «·’ÕÌÕ " & CHR(13)
                        Msg = Msg + "√Ê  €ÌÌ— √Ê  „ÌÌ“ «”„ „œ—”…"
                        MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
                        txtName.SetFocus
                        Exit Sub
                    End If
                End If
                RsTemp.Close
                StrSQL = "select * From  TblSchooleFile  where ministerNo ='" & Trim(txtMinisterNo.Text) & "'"
                RsTemp.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText
                If RsTemp.RecordCount > 0 Then
                    If RsTemp("ID").value <> val(txtID.Text) Then
                        Msg = "Â‰«ﬂ „œ—”…  „”Ã· „”»ﬁ« »Â–«  «·—ﬁ„ «·Ê“«—Ï" & CHR(13)
                        Msg = Msg + "»—Ã«¡ «· √ﬂœ „‰ «·—ﬁ„ «·’ÕÌÕ " & CHR(13)
                        MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
                        txtName.SetFocus
                        Exit Sub
                    End If
                End If
        End Select

        Select Case Me.TxtModFlg.Text

           Case "N"
                rs.AddNew
                txtID.Text = CStr(new_id("TblSchooleFile", "ID", "", True))
           End Select

        Cn.BeginTrans
        BeginTrans = True
        rs("ID").value = val(txtID.Text)
        rs("MinisterNo").value = IIf(txtMinisterNo.Text = "", Null, txtMinisterNo.Text)
        rs("Name").value = IIf(txtName.Text = "", Null, Trim(txtName.Text))
        rs("Namee").value = IIf(txtNameE.Text = "", Null, Trim(txtNameE.Text))
        
        rs("MotName").value = IIf(txtMotName.Text = "", Null, Trim(txtMotName.Text))
        
        rs("ManagerID").value = IIf(txtManager = "", Null, Trim(txtManager.Text))
        rs("ManagerialID").value = IIf(dcManagerialID.BoundText = "", Null, val(dcManagerialID.BoundText))
        rs("CityID").value = IIf(dcCity.BoundText = "", Null, val(dcCity.BoundText))
        rs("StudentCount").value = IIf(Not IsNumeric(txtCount.Text), Null, val(txtCount.Text))
        rs("Custom").value = IIf(Not IsNumeric(txtMoney.Text), 0, val(txtMoney.Text))
        rs("SchooleType").value = IIf(cbSchooleType.ListIndex = -1, Null, cbSchooleType.ListIndex)
        rs("Telephone").value = txtTel.Text
        rs("phase").value = IIf(Peroid.ListIndex = -1, Null, Peroid.ListIndex)
        rs("EducationType").value = IIf(txtEducationType.Text = "", Null, Trim(txtEducationType.Text))
  
    
    rs("SexType").value = IIf(cbSexType.ListIndex = -1, Null, cbSexType.ListIndex)
    rs("SupervisorCode").value = IIf(txtSupervisorCode.Text = "", "", txtSupervisorCode.Text)
    rs("SupervisorName").value = IIf(txtSupervisorName.Text = "", "", txtSupervisorName.Text)
    rs("LastCode").value = IIf(txtLastContract.Text = "", "", txtLastContract.Text)
    
      rs("Emp_ID").value = IIf(DcEmployee.BoundText = "", Null, DcEmployee.BoundText)
    
        rs.update
    
        Dim StrDes As String

     

        Cn.CommitTrans
        BeginTrans = False
        XPTxtCurrent.Caption = rs.AbsolutePosition
        XPTxtCount.Caption = rs.RecordCount
        'CuurentLogdata

        Select Case Me.TxtModFlg.Text

            Case "N"

                If SystemOptions.UserInterface = ArabicInterface Then
                    Msg = "  „ Õ›Ÿ »Ì«‰«  Â–Â «·„œ—”… " & CHR(13)
                    Msg = Msg + "Â·  —€» ›Ì ≈÷«›… »Ì«‰«  √Œ—Ì"
                Else
                    Msg = "Saved" & CHR(13)
                    Msg = Msg + "Do you want enter another One"
                End If

                If MsgBox(Msg, vbYesNo + vbQuestion + vbMsgBoxRight + vbMsgBoxRtlReading + vbDefaultButton2, App.title) = vbYes Then
                    Cmd_Click (0)
                    Exit Sub
                End If
            
            Case "E"
        
                If SystemOptions.UserInterface = ArabicInterface Then
                    MsgBox " „ Õ›Ÿ Â–Â «· ⁄œÌ·« ", vbInformation + vbOKOnly + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
                Else
                    MsgBox "Saved", vbInformation + vbOKOnly + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
                End If

        End Select

        TxtModFlg.Text = "R"
    End If

    Exit Sub
ErrTrap:

    If BeginTrans = True Then
        BeginTrans = False
        Cn.RollbackTrans
    End If

    If Err.Number = -2147217900 Then
        Msg = "·« Ì„ﬂ‰ Õ›Ÿ Â–Â «·»Ì«‰«  " & CHR(13)
        Msg = Msg + "·ﬁœ  „ «œŒ«· ﬁÌ„ €Ì— ’«·Õ… " & CHR(13)
        Msg = Msg + " √ﬂœ „‰ œﬁ… «·»Ì«‰«  Ê√⁄œ «·„Õ«Ê·…"
        MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
        Exit Sub
    End If

    Msg = "⁄›Ê«...ÕœÀ Œÿ√ „« √À‰«¡ Õ›Ÿ Â–Â «·»Ì«‰«  " & CHR(13)
    MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
End Sub

Private Sub Undo()
    On Error GoTo ErrTrap

    Select Case TxtModFlg.Text

        Case "N"
            clear_all Me
            Me.TxtModFlg.Text = "R"
            XPBtnMove_Click (1)

        Case "E"
            rs.Find " ID='" & val(txtID.Text) & "'", , adSearchForward, adBookmarkFirst

            If rs.EOF Or rs.BOF Then
                Me.TxtModFlg.Text = "R"
                Exit Sub
            End If

            Retrive
            Me.TxtModFlg.Text = "R"
    End Select

    Exit Sub
ErrTrap:
End Sub

Private Sub Del_Company()
  
        Dim Msg As String
    Dim StrSQL As String
    Dim RsTemp As New ADODB.Recordset
    Dim StrAccountCode As String
    Dim StrAccountCode1 As String
    Dim ParentAccount As String
    '
 
    On Error GoTo ErrTrap
            
        If txtID.Text <> "" Then
        If SystemOptions.UserInterface = ArabicInterface Then
        Msg = "”Ì „ Õ–› »Ì«‰«  «·„œ—”… —ﬁ„ " & CHR(13)
        Msg = Msg + (txtID.Text) & CHR(13)
        Msg = Msg + " Â·  —€» ›Ì Õ–› Â–Â «·»Ì«‰« ø"
        Else
        Msg = "Delete School File ? " & CHR(13)
        Msg = Msg + (txtID.Text) & CHR(13)
        Msg = Msg + "  Are you sure you want to delete ?"
        End If
        
        
        If MsgBox(Msg, vbYesNo + vbQuestion + vbMsgBoxRight + vbMsgBoxRtlReading, App.title) = vbYes Then
            If Not rs.RecordCount < 1 Then
                                
                 Dim mm As String
                 Set Rs_Temp = New ADODB.Recordset
                 mm = " select * from TblVehicleAllocation_Details   where schoolfileid =  " & val(txtID.Text)
                 Rs_Temp.Open mm, Cn, adOpenStatic, adLockOptimistic, adCmdText
                 If Rs_Temp.RecordCount > 0 Then
                            MsgBox ("·« Ì„ﬂ‰ Õ–› «·„⁄œÂ/«·”Ì«—… · ﬂ«„· «·»Ì«‰« ")
                            Exit Sub
                 End If
                           
                Set Rs_Temp = New ADODB.Recordset
                 mm = " select * from TblVehicleAllocation   where schoolfileid =  " & val(txtID.Text)
                 Rs_Temp.Open mm, Cn, adOpenStatic, adLockOptimistic, adCmdText
                 If Rs_Temp.RecordCount > 0 Then
                            MsgBox ("·« Ì„ﬂ‰ Õ–› «·„⁄œÂ/«·”Ì«—… · ﬂ«„· «·»Ì«‰« ")
                            Exit Sub
                 End If
                           
                Set Rs_Temp = New ADODB.Recordset
                 mm = " select * from TblSupervisorAllocation_details   where schoolfileid =  " & val(txtID.Text)
                 Rs_Temp.Open mm, Cn, adOpenStatic, adLockOptimistic, adCmdText
                 If Rs_Temp.RecordCount > 0 Then
                            MsgBox ("·« Ì„ﬂ‰ Õ–› «·„⁄œÂ/«·”Ì«—… · ﬂ«„· «·»Ì«‰« ")
                            Exit Sub
                 End If
                 
                           
                StrSQL = "delete From TblSchooleFile where  ID =" & val(txtID.Text)
                Cn.Execute StrSQL, , adExecuteNoRecords
                rs.MoveFirst
                    
                   StrSQL = "SELECT  *  From TblSchooleFile "
                   rs.Close
                   rs.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText
                   
                   
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
        If SystemOptions.UserInterface = ArabicInterface Then
        Msg = "Â–Â «·⁄„·Ì… €Ì— „ «Õ… ÕÌÀ √‰Â ·«ÌÊÃœ √Ì ”Ã·« "
        Else
         Msg = "this process Not Aailable"
        End If
        MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
        TxtModFlg_Change
        Exit Sub
    End If

    TxtModFlg_Change
    Exit Sub
ErrTrap:
    'If Err.Number = -2147217887 Then
    Msg = "·«Ì„ﬂ‰ Õ–› Â–« «·”Ã· · ﬂ«„· «·»Ì«‰«  " & CHR(13) & " ÊÃœ »Ì«‰«  „— »ÿ… »Â–« «·„œ—”… "
    Msg = Msg & CHR(13) & Err.description
    MsgBox Msg, vbMsgBoxRight + vbMsgBoxRtlReading + vbExclamation, App.title
    rs.CancelUpdate
    'End If

End Sub

Function DeleteOpeningBalance()


End Function

Private Sub AddTip()
    Dim Wrap As String
    On Error GoTo ErrTrap
    Set TTP = New clstooltip
    Wrap = CHR(13) + CHR(10)

    With TTP
        .Create Me.hwnd, "»Ì«‰«  «·„œ—”…", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
        .AddControl Cmd(0), "ÃœÌœ ..." & Wrap & "·«÷«›… »Ì«‰«  „œ—”… ÃœÌœ…" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
    End With

    With TTP
        .Create Me.hwnd, "»Ì«‰«  «·„œ—”…", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
        .AddControl Cmd(1), " ⁄œÌ· ..." & Wrap & "· ⁄œÌ· »Ì«‰«  «·„œ—”…" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
    End With

    With TTP
        .Create Me.hwnd, "»Ì«‰«  «·„œ—”…", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
        .AddControl Cmd(2), "Õ›Ÿ ..." & Wrap & "·Õ›Ÿ »Ì«‰«  «·„œ—”… «·ÃœÌœ" & Wrap & "·Õ›Ÿ «· ⁄œÌ·« " & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
    End With

    With TTP
        .Create Me.hwnd, "»Ì«‰«  «·„œ—”…", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
        .AddControl Cmd(3), " —«Ã⁄ ..." & Wrap & "·· —«Ã⁄ ⁄‰ ⁄„·Ì… «·«÷«›…" & Wrap & "··· —«Ã⁄ ⁄‰ ⁄„·Ì… «· ⁄œÌ·" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
    End With

    With TTP
        .Create Me.hwnd, "»Ì«‰«  «·„œ—”…", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
        .AddControl Cmd(4), "Õ–› ..." & Wrap & "·Õ–› »Ì«‰«  Â–« «·„œ—”…" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
    End With

    With TTP
        .Create Me.hwnd, "»Ì«‰«  «·„œ—”…", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
        .AddControl Cmd(5), "»ÕÀ ..." & Wrap & "···»ÕÀ ⁄‰ Œ“‰…" & Wrap & "Ì‰ÿ»ﬁ ⁄·ÌÂ ‘—Êÿ „⁄Ì‰…" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
    End With

    With TTP
        .Create Me.hwnd, "»Ì«‰«  «·„œ—”…", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
        .AddControl Cmd(6), "Œ—ÊÃ ..." & Wrap & "·«€·«ﬁ Â–Â «·‰«›–…" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
    End With

    With TTP
        .Create Me.hwnd, "»Ì«‰«  «·„œ—”…", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
        .AddControl XPBtnMove(1), "«·√Ê· ..." & Wrap & "··«‰ ﬁ«· «·Ï √Ê· ”Ã·" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
    End With

    With TTP
        .Create Me.hwnd, "»Ì«‰«  «·„œ—”…", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
        .AddControl XPBtnMove(0), "«·”«»ﬁ ..." & Wrap & "··«‰ ﬁ«· «·Ï «·”Ã· «·”«»ﬁ" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
    End With

    With TTP
        .Create Me.hwnd, "»Ì«‰«  «·„œ—”…", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
        .AddControl XPBtnMove(3), "«· «·Ì ..." & Wrap & "··«‰ ﬁ«· «·Ï «·”Ã· «· «·Ì" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
    End With

    With TTP
        .Create Me.hwnd, "»Ì«‰«  «·„œ—”…", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
        .AddControl XPBtnMove(2), "«·√ŒÌ— ..." & Wrap & "··«‰ ﬁ«· «·Ï ¬Œ— ”Ã·" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
    End With

    With TTP
        .Create Me.hwnd, "»Ì«‰«  «·„œ—”…", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
       ' .AddControl CmdHelp, "„”«⁄œ… ..." & Wrap & "·· ⁄—› ⁄·Ï ÊŸÌ›… Â–Â «·‰«›–…" & Wrap & "ÊﬂÌ›Ì… «· ⁄«„· „⁄Â«" & Wrap & "≈÷€ÿ Â‰«" & Wrap, True
    End With

    Exit Sub
ErrTrap:
End Sub

Private Sub XPTxtBoxName_GotFocus()
    SwitchKeyboardLang LANG_ARABIC
End Sub

Private Sub XPTxtBoxNamee_GotFocus()
    SwitchKeyboardLang LANG_ENGLISH
End Sub
