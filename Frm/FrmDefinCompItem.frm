VERSION 5.00
Object = "{C115893A-A3BF-43AF-B28D-69DB846077F3}#1.0#0"; "vsflex8u.ocx"
Object = "{0AFE7BE0-11B7-4A3E-978D-D4501E9A57FE}#1.0#0"; "c1sizer.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{FE5DCFAD-BC1D-11D2-94CF-004005455FAA}#1.4#0"; "ImpulseButton.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form FrmDefinCompItem 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "‘«‘…  ⁄—Ì› „ﬂÊ‰«  «·«’‰«›/«· Ã„Ì⁄"
   ClientHeight    =   9750
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   16545
   Icon            =   "FrmDefinCompItem.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   9750
   ScaleWidth      =   16545
   Begin C1SizerLibCtl.C1Elastic ELe 
      Height          =   9750
      Index           =   7
      Left            =   0
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   0
      Width           =   16545
      _cx             =   29184
      _cy             =   17198
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
      Begin VB.TextBox TxtAttachedItemCode3 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   2400
         RightToLeft     =   -1  'True
         TabIndex        =   254
         Top             =   420
         Visible         =   0   'False
         Width           =   1725
      End
      Begin VB.TextBox txtPeriod 
         Alignment       =   1  'Right Justify
         Height          =   360
         Left            =   630
         TabIndex        =   244
         Top             =   3135
         Width           =   1425
      End
      Begin VB.TextBox txtRemark 
         Alignment       =   1  'Right Justify
         Height          =   360
         Left            =   8490
         MultiLine       =   -1  'True
         RightToLeft     =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   238
         Top             =   3165
         Width           =   5625
      End
      Begin VB.TextBox XPTxtDiscountVal 
         Alignment       =   1  'Right Justify
         Height          =   360
         Left            =   3450
         TabIndex        =   224
         Top             =   3135
         Width           =   1425
      End
      Begin VB.ComboBox XPCboDiscountType 
         Height          =   315
         Left            =   5970
         Style           =   2  'Dropdown List
         TabIndex        =   222
         Top             =   3165
         Width           =   1245
      End
      Begin C1SizerLibCtl.C1Elastic ELe 
         Height          =   705
         Index           =   6
         Left            =   0
         TabIndex        =   9
         TabStop         =   0   'False
         Top             =   0
         Width           =   16530
         _cx             =   29157
         _cy             =   1244
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
         Caption         =   "‘«‘…  ⁄—Ì› „ﬂÊ‰«  «·«’‰«›/«· Ã„Ì⁄  "
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
         Begin VB.TextBox txtPassword 
            Height          =   315
            IMEMode         =   3  'DISABLE
            Left            =   4290
            PasswordChar    =   "*"
            TabIndex        =   325
            Top             =   390
            Width           =   735
         End
         Begin VB.CheckBox chkIsBranch 
            Caption         =   "»«·›—⁄"
            Height          =   225
            Left            =   5070
            TabIndex        =   324
            Top             =   390
            Width           =   945
         End
         Begin VB.CommandButton cmdReSave 
            Caption         =   "÷»ÿ «·Õ—ﬂ« "
            Height          =   285
            Left            =   8910
            TabIndex        =   321
            Top             =   390
            Visible         =   0   'False
            Width           =   1125
         End
         Begin VB.TextBox TXTTransactionID5 
            Alignment       =   1  'Right Justify
            Height          =   345
            Left            =   4710
            Locked          =   -1  'True
            RightToLeft     =   -1  'True
            TabIndex        =   267
            Top             =   0
            Visible         =   0   'False
            Width           =   1260
         End
         Begin VB.TextBox TXTTransactionID4 
            Alignment       =   1  'Right Justify
            Height          =   345
            Left            =   11940
            Locked          =   -1  'True
            RightToLeft     =   -1  'True
            TabIndex        =   263
            Top             =   1830
            Visible         =   0   'False
            Width           =   1260
         End
         Begin VB.TextBox txtNoteid3 
            Alignment       =   1  'Right Justify
            Height          =   375
            Left            =   3960
            RightToLeft     =   -1  'True
            TabIndex        =   194
            Top             =   0
            Visible         =   0   'False
            Width           =   975
         End
         Begin VB.TextBox TXTTransactionID3 
            Alignment       =   1  'Right Justify
            Height          =   345
            Left            =   4350
            Locked          =   -1  'True
            RightToLeft     =   -1  'True
            TabIndex        =   192
            Top             =   0
            Visible         =   0   'False
            Width           =   1260
         End
         Begin VB.TextBox Text1 
            Alignment       =   1  'Right Justify
            Height          =   375
            Left            =   3840
            RightToLeft     =   -1  'True
            TabIndex        =   10
            Top             =   240
            Visible         =   0   'False
            Width           =   975
         End
         Begin ImpulseButton.ISButton XPBtnMove 
            Height          =   345
            Index           =   3
            Left            =   1005
            TabIndex        =   11
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
            ButtonImage     =   "FrmDefinCompItem.frx":000C
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
            TabIndex        =   12
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
            ButtonImage     =   "FrmDefinCompItem.frx":03A6
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
            TabIndex        =   13
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
            ButtonImage     =   "FrmDefinCompItem.frx":0740
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
            Index           =   0
            Left            =   1860
            TabIndex        =   14
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
            ButtonImage     =   "FrmDefinCompItem.frx":0ADA
            ColorHighlight  =   4194304
            ColorHoverText  =   16777215
            ColorShadow     =   -2147483631
            ColorOutline    =   -2147483631
            DrawFocusRectangle=   0   'False
            DisabledImageStyle=   1
            ColorToggledHoverText=   16777215
            ColorTextShadow =   16777215
         End
         Begin MSComCtl2.DTPicker txtFromDateReSave 
            Height          =   315
            Left            =   7500
            TabIndex        =   322
            Top             =   300
            Visible         =   0   'False
            Width           =   1380
            _ExtentX        =   2434
            _ExtentY        =   556
            _Version        =   393216
            Format          =   141033473
            CurrentDate     =   38784
         End
         Begin MSComCtl2.DTPicker txtToDateReSave 
            Height          =   315
            Left            =   5940
            TabIndex        =   323
            Top             =   330
            Visible         =   0   'False
            Width           =   1560
            _ExtentX        =   2752
            _ExtentY        =   556
            _Version        =   393216
            Format          =   141033473
            CurrentDate     =   38784
         End
      End
      Begin C1SizerLibCtl.C1Elastic ELe 
         Height          =   375
         Index           =   3
         Left            =   135
         TabIndex        =   15
         TabStop         =   0   'False
         Top             =   8895
         Width           =   16215
         _cx             =   28601
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
            Height          =   390
            Left            =   18405
            Locked          =   -1  'True
            RightToLeft     =   -1  'True
            TabIndex        =   16
            TabStop         =   0   'False
            Top             =   -600
            Visible         =   0   'False
            Width           =   1515
         End
         Begin MSDataListLib.DataCombo DCboUserName 
            Height          =   315
            Left            =   13905
            TabIndex        =   17
            Top             =   60
            Width           =   1395
            _ExtentX        =   2461
            _ExtentY        =   556
            _Version        =   393216
            Enabled         =   0   'False
            Text            =   ""
            RightToLeft     =   -1  'True
         End
         Begin ImpulseButton.ISButton ISButton1 
            Height          =   420
            Index           =   0
            Left            =   12240
            TabIndex        =   169
            Top             =   -30
            Width           =   1515
            _ExtentX        =   2672
            _ExtentY        =   741
            ButtonStyle     =   1
            ButtonPositionImage=   1
            Caption         =   "ÿ»«⁄… «·›« Ê—…"
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
            ButtonImage     =   "FrmDefinCompItem.frx":0E74
            ColorButton     =   14871017
            ColorHoverText  =   16777215
            DrawFocusRectangle=   0   'False
            ColorToggledHoverText=   16777215
         End
         Begin ImpulseButton.ISButton ISButton1 
            Height          =   420
            Index           =   1
            Left            =   7185
            TabIndex        =   183
            Top             =   -60
            Width           =   1695
            _ExtentX        =   2990
            _ExtentY        =   741
            ButtonStyle     =   1
            ButtonPositionImage=   1
            Caption         =   "ÿ»«⁄… ⁄—÷ «·”⁄—"
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
            ButtonImage     =   "FrmDefinCompItem.frx":120E
            ColorButton     =   14871017
            ColorHoverText  =   16777215
            DrawFocusRectangle=   0   'False
            ColorToggledHoverText=   16777215
         End
         Begin ImpulseButton.ISButton ISButton1 
            Height          =   420
            Index           =   2
            Left            =   8895
            TabIndex        =   240
            Top             =   -60
            Width           =   1620
            _ExtentX        =   2858
            _ExtentY        =   741
            ButtonStyle     =   1
            ButtonPositionImage=   1
            Caption         =   "ÿ»«⁄… ”‰œ «·«” ·«„"
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
            ButtonImage     =   "FrmDefinCompItem.frx":15A8
            ColorButton     =   14871017
            ColorHoverText  =   16777215
            DrawFocusRectangle=   0   'False
            ColorToggledHoverText=   16777215
         End
         Begin ImpulseButton.ISButton ISButton1 
            Height          =   420
            Index           =   3
            Left            =   4920
            TabIndex        =   280
            Top             =   -60
            Width           =   2235
            _ExtentX        =   3942
            _ExtentY        =   741
            ButtonStyle     =   1
            ButtonPositionImage=   1
            Caption         =   "ÿ»«⁄… «·„Ê«œ «·Œ«„ «·„ﬁœ—…"
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
            ButtonImage     =   "FrmDefinCompItem.frx":1942
            ColorButton     =   14871017
            ColorHoverText  =   16777215
            DrawFocusRectangle=   0   'False
            ColorToggledHoverText=   16777215
         End
         Begin ImpulseButton.ISButton ISButton1 
            Height          =   420
            Index           =   4
            Left            =   10680
            TabIndex        =   286
            Top             =   -30
            Width           =   1455
            _ExtentX        =   2566
            _ExtentY        =   741
            ButtonStyle     =   1
            ButtonPositionImage=   1
            Caption         =   "«·›« Ê—… „Œ ’—…"
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
            ButtonImage     =   "FrmDefinCompItem.frx":1CDC
            ColorButton     =   14871017
            ColorHoverText  =   16777215
            DrawFocusRectangle=   0   'False
            ColorToggledHoverText=   16777215
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Õ—— »Ê«”ÿ… : "
            Height          =   330
            Index           =   1
            Left            =   15015
            RightToLeft     =   -1  'True
            TabIndex        =   23
            Top             =   90
            Width           =   1245
         End
         Begin VB.Label XPTxtCount 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Height          =   285
            Left            =   330
            RightToLeft     =   -1  'True
            TabIndex        =   22
            Top             =   90
            Width           =   765
         End
         Begin VB.Label XPTxtCurrent 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Height          =   315
            Left            =   2730
            RightToLeft     =   -1  'True
            TabIndex        =   21
            Top             =   90
            Width           =   705
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "⁄œœ «·”Ã·« :"
            Height          =   285
            Index           =   2
            Left            =   1200
            RightToLeft     =   -1  'True
            TabIndex        =   20
            Top             =   90
            Width           =   1320
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "«·”Ã· «·Õ«·Ì:"
            Height          =   285
            Index           =   0
            Left            =   3645
            RightToLeft     =   -1  'True
            TabIndex        =   19
            Top             =   105
            Width           =   1170
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "≈Ã„«·Ì «·„Ê«œ «·Œ«„"
            Height          =   270
            Index           =   3
            Left            =   20730
            RightToLeft     =   -1  'True
            TabIndex        =   18
            Top             =   -60
            Visible         =   0   'False
            Width           =   2700
         End
      End
      Begin C1SizerLibCtl.C1Elastic ELe 
         Height          =   5400
         Index           =   5
         Left            =   0
         TabIndex        =   24
         TabStop         =   0   'False
         Top             =   3510
         Width           =   16575
         _cx             =   29236
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
         Appearance      =   4
         MousePointer    =   0
         Version         =   801
         BackColor       =   14871017
         ForeColor       =   -2147483630
         FloodColor      =   6553600
         ForeColorDisabled=   -2147483631
         Caption         =   "≈’œ«— ≈–‰ ‘Õ‰"
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
         Begin VB.CommandButton Command2 
            Caption         =   "Command2"
            Height          =   315
            Left            =   1860
            TabIndex        =   326
            Top             =   570
            Width           =   885
         End
         Begin VB.TextBox txtQty5 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   3420
            RightToLeft     =   -1  'True
            TabIndex        =   292
            Text            =   "1"
            Top             =   450
            Width           =   975
         End
         Begin VB.TextBox txtTotalWithVat 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   3345
            Locked          =   -1  'True
            RightToLeft     =   -1  'True
            TabIndex        =   186
            Top             =   735
            Width           =   1065
         End
         Begin VB.TextBox TxtVAt2 
            Alignment       =   1  'Right Justify
            Height          =   330
            Left            =   6390
            Locked          =   -1  'True
            RightToLeft     =   -1  'True
            TabIndex        =   185
            Top             =   720
            Width           =   855
         End
         Begin VB.TextBox TxtVATValue 
            Alignment       =   1  'Right Justify
            Height          =   390
            Left            =   3540
            RightToLeft     =   -1  'True
            TabIndex        =   184
            Top             =   -480
            Visible         =   0   'False
            Width           =   960
         End
         Begin VB.TextBox txtNet 
            Alignment       =   1  'Right Justify
            Height          =   300
            Left            =   8520
            Locked          =   -1  'True
            RightToLeft     =   -1  'True
            TabIndex        =   172
            Top             =   735
            Width           =   1110
         End
         Begin VB.TextBox txtTotalDisc 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00C0C0FF&
            Height          =   285
            Left            =   10605
            Locked          =   -1  'True
            RightToLeft     =   -1  'True
            TabIndex        =   170
            Top             =   750
            Width           =   1125
         End
         Begin VB.TextBox txtTotalAdd 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00C0FFC0&
            Height          =   285
            Left            =   13455
            Locked          =   -1  'True
            RightToLeft     =   -1  'True
            TabIndex        =   165
            Top             =   750
            Width           =   1155
         End
         Begin C1SizerLibCtl.C1Elastic ELe 
            Height          =   810
            Index           =   8
            Left            =   180
            TabIndex        =   25
            TabStop         =   0   'False
            Top             =   4530
            Width           =   16290
            _cx             =   28734
            _cy             =   1429
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
            Begin VB.CommandButton cmdTransfer 
               Caption         =   " ÕÊÌ·"
               Height          =   360
               Left            =   1500
               RightToLeft     =   -1  'True
               TabIndex        =   305
               Top             =   60
               Width           =   1410
            End
            Begin VB.CommandButton cmdCancel 
               Caption         =   "«·€«¡ «· ÕÊÌ· Ê«·›« Ê—…"
               Height          =   360
               Left            =   0
               RightToLeft     =   -1  'True
               TabIndex        =   304
               Top             =   60
               Width           =   1410
            End
            Begin VB.CommandButton cmdfrmRec 
               Caption         =   "ﬁ»÷ œ›⁄… „"
               Height          =   345
               Left            =   0
               RightToLeft     =   -1  'True
               TabIndex        =   303
               Top             =   435
               Width           =   1410
            End
            Begin VB.CommandButton CMDSHOWISSUE2 
               Caption         =   "⁄—÷ ”‰œ ’—› «·›« Ê—…"
               Height          =   345
               Left            =   6690
               RightToLeft     =   -1  'True
               TabIndex        =   269
               Top             =   435
               Width           =   1830
            End
            Begin VB.TextBox TxtNoteSerial15 
               Alignment       =   1  'Right Justify
               Height          =   345
               Left            =   5070
               RightToLeft     =   -1  'True
               TabIndex        =   268
               Top             =   435
               Width           =   1620
            End
            Begin VB.CommandButton cmdCancel2 
               Caption         =   "«·€«¡ «„— «·«‰ «Ã"
               Height          =   345
               Left            =   1440
               RightToLeft     =   -1  'True
               TabIndex        =   264
               Top             =   435
               Visible         =   0   'False
               Width           =   1500
            End
            Begin VB.TextBox TxtNoteSerial14 
               Alignment       =   1  'Right Justify
               Height          =   345
               Left            =   3120
               RightToLeft     =   -1  'True
               TabIndex        =   262
               Top             =   435
               Visible         =   0   'False
               Width           =   1890
            End
            Begin VB.CommandButton cmdCreateProduction 
               Caption         =   "«‰‘«¡ «„— «‰ «Ã"
               Enabled         =   0   'False
               Height          =   360
               Left            =   3120
               RightToLeft     =   -1  'True
               TabIndex        =   261
               Top             =   60
               Width           =   1890
            End
            Begin VB.CommandButton cmdCreateSales 
               Caption         =   "⁄—÷ «·›« Ê—…"
               Height          =   360
               Left            =   6690
               RightToLeft     =   -1  'True
               TabIndex        =   195
               Top             =   60
               Width           =   1830
            End
            Begin VB.TextBox TxtNoteSerial13 
               Alignment       =   1  'Right Justify
               Height          =   360
               Left            =   5070
               RightToLeft     =   -1  'True
               TabIndex        =   193
               Top             =   60
               Width           =   1620
            End
            Begin VB.CommandButton CMDSHOWecive 
               Caption         =   "⁄—÷  ”‰œ «” ·«„ „‰ Ã  «„"
               Height          =   345
               Left            =   8505
               RightToLeft     =   -1  'True
               TabIndex        =   154
               Top             =   435
               Width           =   1515
            End
            Begin VB.CommandButton CMDSHOWISSUE 
               Caption         =   "⁄—÷  ”‰œ ’—› „Ê«œ Œ«„"
               Height          =   360
               Left            =   8520
               RightToLeft     =   -1  'True
               TabIndex        =   153
               Top             =   60
               Width           =   1515
            End
            Begin VB.TextBox TxtNoteSerial11 
               Alignment       =   1  'Right Justify
               Height          =   360
               Left            =   10005
               RightToLeft     =   -1  'True
               TabIndex        =   152
               Top             =   60
               Width           =   1500
            End
            Begin VB.TextBox TxtNoteSerial12 
               Alignment       =   1  'Right Justify
               Height          =   345
               Left            =   10005
               RightToLeft     =   -1  'True
               TabIndex        =   151
               Top             =   435
               Width           =   1500
            End
            Begin VB.TextBox TXTTransactionID2 
               Alignment       =   1  'Right Justify
               Height          =   705
               Left            =   16380
               RightToLeft     =   -1  'True
               TabIndex        =   148
               Top             =   675
               Visible         =   0   'False
               Width           =   2460
            End
            Begin VB.TextBox TXTTransactionID1 
               Alignment       =   1  'Right Justify
               Height          =   540
               Left            =   16380
               RightToLeft     =   -1  'True
               TabIndex        =   147
               Top             =   150
               Visible         =   0   'False
               Width           =   2460
            End
            Begin VB.CheckBox Selct 
               Alignment       =   1  'Right Justify
               Caption         =   "Ì „ ⁄„· ’—› „Ê«œ Œ«„"
               Height          =   360
               Index           =   1
               Left            =   14670
               RightToLeft     =   -1  'True
               TabIndex        =   27
               Top             =   60
               Width           =   1560
            End
            Begin VB.CheckBox Selct 
               Alignment       =   1  'Right Justify
               Caption         =   "Ì „ ⁄„· «” ·«„ „‰ Ã  «„"
               Height          =   345
               Index           =   2
               Left            =   14325
               RightToLeft     =   -1  'True
               TabIndex        =   26
               Top             =   435
               Width           =   1905
            End
            Begin MSDataListLib.DataCombo DCboStore2Name 
               Height          =   315
               Left            =   11520
               TabIndex        =   28
               Top             =   150
               Width           =   1995
               _ExtentX        =   3519
               _ExtentY        =   556
               _Version        =   393216
               Text            =   ""
               RightToLeft     =   -1  'True
            End
            Begin MSDataListLib.DataCombo DCboStore3Name 
               Height          =   315
               Left            =   11520
               TabIndex        =   29
               Top             =   450
               Width           =   1995
               _ExtentX        =   3519
               _ExtentY        =   556
               _Version        =   393216
               Text            =   ""
               RightToLeft     =   -1  'True
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               Caption         =   "Õœœ «·„Œ“‰"
               Height          =   360
               Index           =   47
               Left            =   13320
               RightToLeft     =   -1  'True
               TabIndex        =   31
               Top             =   60
               Width           =   1005
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               Caption         =   "Õœœ «·„Œ“‰"
               Height          =   345
               Index           =   48
               Left            =   13320
               RightToLeft     =   -1  'True
               TabIndex        =   30
               Top             =   435
               Width           =   1005
            End
         End
         Begin C1SizerLibCtl.C1Elastic ELe 
            Height          =   540
            Index           =   9
            Left            =   120
            TabIndex        =   140
            TabStop         =   0   'False
            Top             =   -150
            Width           =   16260
            _cx             =   28681
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
            Begin VB.CheckBox chkIsAdd 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "«÷«›« "
               ForeColor       =   &H00FF0000&
               Height          =   300
               Left            =   15330
               RightToLeft     =   -1  'True
               TabIndex        =   199
               Top             =   255
               Width           =   855
            End
            Begin VB.TextBox TxtAttachedItemCode2 
               Alignment       =   1  'Right Justify
               Height          =   300
               Left            =   540
               TabIndex        =   4
               Top             =   150
               Visible         =   0   'False
               Width           =   1395
            End
            Begin VB.TextBox txtQty 
               Alignment       =   1  'Right Justify
               Height          =   285
               Left            =   3300
               RightToLeft     =   -1  'True
               TabIndex        =   7
               Text            =   "1"
               Top             =   300
               Width           =   975
            End
            Begin MSDataListLib.DataCombo DcbUnit2 
               Height          =   315
               Left            =   6045
               TabIndex        =   6
               Top             =   240
               Width           =   1080
               _ExtentX        =   1905
               _ExtentY        =   556
               _Version        =   393216
               Text            =   ""
               RightToLeft     =   -1  'True
            End
            Begin MSDataListLib.DataCombo DcboItemID2 
               Height          =   315
               Left            =   8340
               TabIndex        =   5
               Top             =   240
               Width           =   2775
               _ExtentX        =   4895
               _ExtentY        =   556
               _Version        =   393216
               Text            =   ""
               RightToLeft     =   -1  'True
            End
            Begin MSDataListLib.DataCombo XPCboGroup2 
               Height          =   315
               Left            =   12345
               TabIndex        =   200
               Top             =   240
               Width           =   1650
               _ExtentX        =   2910
               _ExtentY        =   556
               _Version        =   393216
               Text            =   ""
               RightToLeft     =   -1  'True
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackColor       =   &H00E2E9E9&
               Caption         =   " «·„Ã„Ê⁄…"
               Height          =   330
               Index           =   11
               Left            =   14010
               TabIndex        =   201
               Top             =   285
               Width           =   855
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "ﬂÊœ «·’‰›"
               Height          =   300
               Index           =   39
               Left            =   1935
               RightToLeft     =   -1  'True
               TabIndex        =   145
               Top             =   180
               Visible         =   0   'False
               Width           =   960
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   " Ã·«Ìœ—"
               Height          =   285
               Index           =   38
               Left            =   10995
               RightToLeft     =   -1  'True
               TabIndex        =   144
               Top             =   345
               Width           =   765
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "«·ÊÕœÂ"
               Height          =   300
               Index           =   33
               Left            =   6975
               RightToLeft     =   -1  'True
               TabIndex        =   143
               Top             =   240
               Width           =   675
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "«·ﬂ„Ì…"
               Height          =   255
               Index           =   32
               Left            =   4200
               RightToLeft     =   -1  'True
               TabIndex        =   142
               Top             =   345
               Width           =   690
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "«·„ﬂÊ‰« "
               BeginProperty Font 
                  Name            =   "Times New Roman"
                  Size            =   14.25
                  Charset         =   178
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00FF0000&
               Height          =   330
               Index           =   31
               Left            =   12480
               RightToLeft     =   -1  'True
               TabIndex        =   141
               Top             =   1140
               Width           =   3705
            End
         End
         Begin C1SizerLibCtl.C1Tab TabMain 
            Height          =   3480
            Left            =   30
            TabIndex        =   174
            Top             =   990
            Width           =   16485
            _cx             =   29078
            _cy             =   6138
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
            ForeColor       =   -2147483630
            FrontTabColor   =   14871017
            BackTabColor    =   12648447
            TabOutlineColor =   -2147483632
            FrontTabForeColor=   16711680
            Caption         =   "«·„Ê«œ «·Œ«„ «·„Õ–Ê›…|«·„Ê«œ «·Œ«„ |»Ì«‰« |Õ—ﬂ«  «·’—› «· «»⁄… "
            Align           =   0
            CurrTab         =   2
            FirstTab        =   0
            Style           =   3
            Position        =   1
            AutoSwitch      =   -1  'True
            AutoScroll      =   -1  'True
            TabPreview      =   -1  'True
            ShowFocusRect   =   -1  'True
            TabsPerPage     =   4
            BorderWidth     =   0
            BoldCurrent     =   0   'False
            DogEars         =   -1  'True
            MultiRow        =   -1  'True
            MultiRowOffset  =   200
            CaptionStyle    =   0
            TabHeight       =   0
            TabCaptionPos   =   4
            TabPicturePos   =   0
            CaptionEmpty    =   ""
            Separators      =   0   'False
            AccessibleName  =   ""
            AccessibleDescription=   ""
            AccessibleValue =   ""
            AccessibleRole  =   37
            Begin C1SizerLibCtl.C1Elastic ELe 
               Height          =   3105
               Index           =   10
               Left            =   -17340
               TabIndex        =   175
               TabStop         =   0   'False
               Top             =   45
               Width           =   16395
               _cx             =   28919
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
               Begin VSFlex8UCtl.VSFlexGrid FgItems 
                  Height          =   3105
                  Index           =   0
                  Left            =   22905
                  TabIndex        =   176
                  Top             =   855
                  Width           =   16155
                  _cx             =   28496
                  _cy             =   5477
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
                  RowHeightMin    =   300
                  RowHeightMax    =   0
                  ColWidthMin     =   0
                  ColWidthMax     =   0
                  ExtendLastCol   =   -1  'True
                  FormatString    =   $"FrmDefinCompItem.frx":2076
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
               Begin VSFlex8UCtl.VSFlexGrid FGDeleted 
                  Height          =   3000
                  Left            =   60
                  TabIndex        =   180
                  Top             =   75
                  Width           =   16275
                  _cx             =   28707
                  _cy             =   5292
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
                  Cols            =   40
                  FixedRows       =   1
                  FixedCols       =   1
                  RowHeightMin    =   300
                  RowHeightMax    =   0
                  ColWidthMin     =   0
                  ColWidthMax     =   0
                  ExtendLastCol   =   -1  'True
                  FormatString    =   $"FrmDefinCompItem.frx":2136
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
            End
            Begin C1SizerLibCtl.C1Elastic ELe 
               Height          =   3105
               Index           =   11
               Left            =   -17040
               TabIndex        =   177
               TabStop         =   0   'False
               Top             =   45
               Width           =   16395
               _cx             =   28919
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
               Begin VB.Frame Frame4 
                  BorderStyle     =   0  'None
                  Height          =   645
                  Left            =   1920
                  TabIndex        =   306
                  Top             =   30
                  Width           =   5145
                  Begin VB.TextBox txtwidtj2 
                     Alignment       =   1  'Right Justify
                     Height          =   270
                     Left            =   4215
                     RightToLeft     =   -1  'True
                     TabIndex        =   313
                     Text            =   "1"
                     Top             =   330
                     Width           =   735
                  End
                  Begin VB.TextBox txthight2 
                     Alignment       =   1  'Right Justify
                     Height          =   270
                     Left            =   3510
                     RightToLeft     =   -1  'True
                     TabIndex        =   312
                     Text            =   "1"
                     Top             =   330
                     Width           =   690
                  End
                  Begin VB.TextBox txtLength2 
                     Alignment       =   1  'Right Justify
                     Height          =   270
                     Left            =   2820
                     RightToLeft     =   -1  'True
                     TabIndex        =   311
                     Text            =   "1"
                     Top             =   330
                     Width           =   690
                  End
                  Begin VB.TextBox txtDiameter2 
                     Alignment       =   1  'Right Justify
                     Height          =   270
                     Left            =   0
                     RightToLeft     =   -1  'True
                     TabIndex        =   310
                     Text            =   "1"
                     Top             =   330
                     Width           =   720
                  End
                  Begin VB.TextBox txtthickness2 
                     Alignment       =   1  'Right Justify
                     Height          =   270
                     Left            =   2100
                     RightToLeft     =   -1  'True
                     TabIndex        =   309
                     Text            =   "1"
                     Top             =   330
                     Width           =   720
                  End
                  Begin VB.TextBox txtDO2 
                     Alignment       =   1  'Right Justify
                     Height          =   270
                     Left            =   1410
                     RightToLeft     =   -1  'True
                     TabIndex        =   308
                     Text            =   "1"
                     Top             =   330
                     Width           =   690
                  End
                  Begin VB.TextBox txtDI2 
                     Alignment       =   1  'Right Justify
                     Height          =   270
                     Left            =   690
                     RightToLeft     =   -1  'True
                     TabIndex        =   307
                     Text            =   "1"
                     Top             =   330
                     Width           =   720
                  End
                  Begin VB.Label lbl 
                     Alignment       =   1  'Right Justify
                     BackColor       =   &H00E2E9E9&
                     BackStyle       =   0  'Transparent
                     Caption         =   "«·⁄—÷ W"
                     Height          =   300
                     Index           =   102
                     Left            =   4200
                     RightToLeft     =   -1  'True
                     TabIndex        =   320
                     Top             =   0
                     Width           =   675
                  End
                  Begin VB.Label lbl 
                     Alignment       =   1  'Right Justify
                     BackColor       =   &H00E2E9E9&
                     BackStyle       =   0  'Transparent
                     Caption         =   "«— ›«⁄L "
                     Height          =   300
                     Index           =   101
                     Left            =   3285
                     RightToLeft     =   -1  'True
                     TabIndex        =   319
                     Top             =   30
                     Width           =   840
                  End
                  Begin VB.Label lbl 
                     Alignment       =   1  'Right Justify
                     BackColor       =   &H00E2E9E9&
                     BackStyle       =   0  'Transparent
                     Caption         =   "«·⁄„ﬁ"
                     Height          =   300
                     Index           =   100
                     Left            =   2895
                     RightToLeft     =   -1  'True
                     TabIndex        =   318
                     Top             =   30
                     Width           =   420
                  End
                  Begin VB.Label lbl 
                     Alignment       =   1  'Right Justify
                     BackColor       =   &H00E2E9E9&
                     BackStyle       =   0  'Transparent
                     Caption         =   "«·ﬁÿ—"
                     Height          =   300
                     Index           =   97
                     Left            =   75
                     RightToLeft     =   -1  'True
                     TabIndex        =   317
                     Top             =   30
                     Width           =   450
                  End
                  Begin VB.Label lbl 
                     Alignment       =   1  'Right Justify
                     BackColor       =   &H00E2E9E9&
                     BackStyle       =   0  'Transparent
                     Caption         =   "«·”„ﬂ"
                     Height          =   300
                     Index           =   96
                     Left            =   2145
                     RightToLeft     =   -1  'True
                     TabIndex        =   316
                     Top             =   30
                     Width           =   450
                  End
                  Begin VB.Label lbl 
                     Alignment       =   1  'Right Justify
                     BackColor       =   &H00E2E9E9&
                     BackStyle       =   0  'Transparent
                     Caption         =   "DO"
                     Height          =   300
                     Index           =   95
                     Left            =   1455
                     RightToLeft     =   -1  'True
                     TabIndex        =   315
                     Top             =   30
                     Width           =   450
                  End
                  Begin VB.Label lbl 
                     Alignment       =   1  'Right Justify
                     BackColor       =   &H00E2E9E9&
                     BackStyle       =   0  'Transparent
                     Caption         =   "DI"
                     Height          =   300
                     Index           =   94
                     Left            =   765
                     RightToLeft     =   -1  'True
                     TabIndex        =   314
                     Top             =   30
                     Width           =   450
                  End
               End
               Begin VB.CheckBox chkSelectAll 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   " ÕœÌœ «·ﬂ·"
                  ForeColor       =   &H00FF0000&
                  Height          =   195
                  Index           =   1
                  Left            =   15105
                  RightToLeft     =   -1  'True
                  TabIndex        =   266
                  Top             =   2820
                  Width           =   1080
               End
               Begin VB.TextBox txtItemCode 
                  Alignment       =   1  'Right Justify
                  Height          =   285
                  Left            =   14160
                  RightToLeft     =   -1  'True
                  TabIndex        =   234
                  Top             =   120
                  Width           =   1260
               End
               Begin VB.TextBox txtQty3 
                  Alignment       =   1  'Right Justify
                  Height          =   300
                  Left            =   7110
                  RightToLeft     =   -1  'True
                  TabIndex        =   229
                  Text            =   "1"
                  Top             =   30
                  Width           =   960
               End
               Begin VSFlex8UCtl.VSFlexGrid FgItems 
                  Height          =   3000
                  Index           =   1
                  Left            =   22905
                  TabIndex        =   178
                  Top             =   855
                  Width           =   16155
                  _cx             =   28496
                  _cy             =   5292
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
                  RowHeightMin    =   300
                  RowHeightMax    =   0
                  ColWidthMin     =   0
                  ColWidthMax     =   0
                  ExtendLastCol   =   -1  'True
                  FormatString    =   $"FrmDefinCompItem.frx":272E
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
               Begin VSFlex8UCtl.VSFlexGrid FG 
                  Height          =   1980
                  Left            =   -60
                  TabIndex        =   179
                  Top             =   750
                  Width           =   16395
                  _cx             =   28919
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
                  Cols            =   63
                  FixedRows       =   1
                  FixedCols       =   1
                  RowHeightMin    =   300
                  RowHeightMax    =   0
                  ColWidthMin     =   0
                  ColWidthMax     =   0
                  ExtendLastCol   =   -1  'True
                  FormatString    =   $"FrmDefinCompItem.frx":27EE
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
               Begin ImpulseButton.ISButton Cmd 
                  Height          =   375
                  Index           =   8
                  Left            =   0
                  TabIndex        =   205
                  Top             =   2700
                  Width           =   780
                  _ExtentX        =   1376
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
                  ButtonImage     =   "FrmDefinCompItem.frx":310B
                  ColorButton     =   14871017
                  DrawFocusRectangle=   0   'False
               End
               Begin MSDataListLib.DataCombo DcboItemID3 
                  Height          =   315
                  Left            =   10410
                  TabIndex        =   227
                  Top             =   90
                  Width           =   2760
                  _ExtentX        =   4868
                  _ExtentY        =   556
                  _Version        =   393216
                  Text            =   ""
                  RightToLeft     =   -1  'True
               End
               Begin MSDataListLib.DataCombo DcboItemID4 
                  Height          =   315
                  Left            =   10410
                  TabIndex        =   231
                  Top             =   420
                  Width           =   2745
                  _ExtentX        =   4842
                  _ExtentY        =   556
                  _Version        =   393216
                  Enabled         =   0   'False
                  Text            =   ""
                  RightToLeft     =   -1  'True
               End
               Begin ImpulseButton.ISButton cmdAdd2 
                  Height          =   495
                  Left            =   150
                  TabIndex        =   233
                  Top             =   90
                  Width           =   1110
                  _ExtentX        =   1958
                  _ExtentY        =   873
                  ButtonStyle     =   1
                  ButtonPositionImage=   2
                  Caption         =   "≈÷«›… „Ê«œ Œ«„"
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
                  ButtonImage     =   "FrmDefinCompItem.frx":36A5
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
               Begin MSDataListLib.DataCombo DcbUnit3 
                  Height          =   315
                  Left            =   8700
                  TabIndex        =   236
                  Top             =   60
                  Width           =   1080
                  _ExtentX        =   1905
                  _ExtentY        =   556
                  _Version        =   393216
                  Text            =   ""
                  RightToLeft     =   -1  'True
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "«·ÊÕœÂ"
                  Height          =   345
                  Index           =   73
                  Left            =   9780
                  RightToLeft     =   -1  'True
                  TabIndex        =   237
                  Top             =   60
                  Width           =   615
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  BackStyle       =   0  'Transparent
                  Caption         =   "ﬂÊœ «·’‰›"
                  ForeColor       =   &H00000000&
                  Height          =   240
                  Index           =   72
                  Left            =   15120
                  RightToLeft     =   -1  'True
                  TabIndex        =   235
                  Top             =   150
                  Width           =   1200
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "Ì »⁄ ·’‰›"
                  Enabled         =   0   'False
                  Height          =   390
                  Index           =   71
                  Left            =   13170
                  RightToLeft     =   -1  'True
                  TabIndex        =   232
                  Top             =   420
                  Width           =   975
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "«·ﬂ„Ì…"
                  Height          =   345
                  Index           =   70
                  Left            =   8040
                  RightToLeft     =   -1  'True
                  TabIndex        =   230
                  Top             =   60
                  Width           =   660
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "«·’‰› «·Œ«„"
                  Height          =   270
                  Index           =   69
                  Left            =   13050
                  RightToLeft     =   -1  'True
                  TabIndex        =   228
                  Top             =   120
                  Width           =   1095
               End
            End
            Begin C1SizerLibCtl.C1Elastic ELe 
               Height          =   3105
               Index           =   12
               Left            =   45
               TabIndex        =   202
               TabStop         =   0   'False
               Top             =   45
               Width           =   16395
               _cx             =   28919
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
               Begin VB.CheckBox chkSelectAll 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   " ÕœÌœ «·ﬂ·"
                  ForeColor       =   &H00FF0000&
                  Height          =   300
                  Index           =   0
                  Left            =   15195
                  RightToLeft     =   -1  'True
                  TabIndex        =   265
                  Top             =   2790
                  Width           =   1080
               End
               Begin VB.TextBox txtTotalWithVat2 
                  Alignment       =   1  'Right Justify
                  Height          =   405
                  Left            =   1695
                  Locked          =   -1  'True
                  RightToLeft     =   -1  'True
                  TabIndex        =   217
                  Top             =   2475
                  Width           =   1080
               End
               Begin VB.TextBox TxtVAt22 
                  Alignment       =   1  'Right Justify
                  Height          =   405
                  Left            =   5130
                  Locked          =   -1  'True
                  RightToLeft     =   -1  'True
                  TabIndex        =   215
                  Top             =   2475
                  Width           =   840
               End
               Begin VB.TextBox txtNet2 
                  Alignment       =   1  'Right Justify
                  Height          =   345
                  Left            =   7110
                  Locked          =   -1  'True
                  RightToLeft     =   -1  'True
                  TabIndex        =   213
                  Top             =   2490
                  Width           =   1095
               End
               Begin VB.TextBox txtTotalDisc2 
                  Alignment       =   1  'Right Justify
                  Height          =   345
                  Left            =   9105
                  Locked          =   -1  'True
                  RightToLeft     =   -1  'True
                  TabIndex        =   211
                  Top             =   2490
                  Width           =   1140
               End
               Begin VB.TextBox txtTotalAdd2 
                  Alignment       =   1  'Right Justify
                  Height          =   345
                  Left            =   11820
                  Locked          =   -1  'True
                  RightToLeft     =   -1  'True
                  TabIndex        =   209
                  Top             =   2490
                  Width           =   1080
               End
               Begin VB.TextBox txtTotal2 
                  Alignment       =   1  'Right Justify
                  Height          =   405
                  Left            =   14460
                  RightToLeft     =   -1  'True
                  TabIndex        =   207
                  Text            =   "1"
                  Top             =   2475
                  Width           =   795
               End
               Begin VSFlex8UCtl.VSFlexGrid FgItems 
                  Height          =   2775
                  Index           =   2
                  Left            =   22905
                  TabIndex        =   203
                  Top             =   945
                  Width           =   16155
                  _cx             =   28496
                  _cy             =   4895
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
                  RowHeightMin    =   300
                  RowHeightMax    =   0
                  ColWidthMin     =   0
                  ColWidthMax     =   0
                  ExtendLastCol   =   -1  'True
                  FormatString    =   $"FrmDefinCompItem.frx":3A3F
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
               Begin VSFlex8UCtl.VSFlexGrid FG2 
                  Height          =   2355
                  Left            =   0
                  TabIndex        =   204
                  Top             =   120
                  Width           =   16395
                  _cx             =   28919
                  _cy             =   4154
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
                  Cols            =   52
                  FixedRows       =   1
                  FixedCols       =   1
                  RowHeightMin    =   300
                  RowHeightMax    =   0
                  ColWidthMin     =   0
                  ColWidthMax     =   0
                  ExtendLastCol   =   -1  'True
                  FormatString    =   $"FrmDefinCompItem.frx":3AFF
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
               Begin ImpulseButton.ISButton Cmd 
                  Height          =   270
                  Index           =   11
                  Left            =   60
                  TabIndex        =   206
                  Top             =   2775
                  Width           =   780
                  _ExtentX        =   1376
                  _ExtentY        =   476
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
                  ButtonImage     =   "FrmDefinCompItem.frx":42AB
                  ColorButton     =   14871017
                  DrawFocusRectangle=   0   'False
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "«·’«›Ï »⁄œ «·ﬁÌ„… «·„÷«›…"
                  Height          =   375
                  Index           =   63
                  Left            =   2835
                  RightToLeft     =   -1  'True
                  TabIndex        =   218
                  Top             =   2490
                  Width           =   2115
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "«·ﬁÌ„… «·„÷«›…"
                  Height          =   375
                  Index           =   62
                  Left            =   6030
                  RightToLeft     =   -1  'True
                  TabIndex        =   216
                  Top             =   2490
                  Width           =   1080
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "«·’«›Ì"
                  Height          =   390
                  Index           =   61
                  Left            =   8445
                  RightToLeft     =   -1  'True
                  TabIndex        =   214
                  Top             =   2475
                  Width           =   600
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "«Ã„«·Ì «·Œ’Ê„« "
                  Height          =   345
                  Index           =   60
                  Left            =   10245
                  RightToLeft     =   -1  'True
                  TabIndex        =   212
                  Top             =   2490
                  Width           =   1515
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "«Ã„«·Ì «·«÷«›« "
                  Height          =   315
                  Index           =   44
                  Left            =   13080
                  RightToLeft     =   -1  'True
                  TabIndex        =   210
                  Top             =   2475
                  Width           =   1380
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "«·«Ã„«·Ì"
                  Height          =   405
                  Index           =   25
                  Left            =   15315
                  RightToLeft     =   -1  'True
                  TabIndex        =   208
                  Top             =   2475
                  Width           =   660
               End
            End
            Begin C1SizerLibCtl.C1Elastic ELe 
               Height          =   3105
               Index           =   13
               Left            =   17130
               TabIndex        =   282
               TabStop         =   0   'False
               Top             =   45
               Width           =   16395
               _cx             =   28919
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
               Begin VSFlex8UCtl.VSFlexGrid FgItems 
                  Height          =   3105
                  Index           =   3
                  Left            =   22905
                  TabIndex        =   283
                  Top             =   855
                  Width           =   16155
                  _cx             =   28496
                  _cy             =   5477
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
                  RowHeightMin    =   300
                  RowHeightMax    =   0
                  ColWidthMin     =   0
                  ColWidthMax     =   0
                  ExtendLastCol   =   -1  'True
                  FormatString    =   $"FrmDefinCompItem.frx":4845
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
               Begin VSFlex8UCtl.VSFlexGrid FG3 
                  Height          =   3000
                  Left            =   60
                  TabIndex        =   284
                  Top             =   60
                  Width           =   16275
                  _cx             =   28707
                  _cy             =   5292
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
                  FormatString    =   $"FrmDefinCompItem.frx":4905
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
            End
         End
         Begin ImpulseButton.ISButton cmdAdd_ 
            Height          =   600
            Left            =   150
            TabIndex        =   221
            Top             =   390
            Width           =   990
            _ExtentX        =   1746
            _ExtentY        =   1058
            ButtonStyle     =   1
            ButtonPositionImage=   2
            Caption         =   "≈÷«›…"
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
            ButtonImage     =   "FrmDefinCompItem.frx":4AFF
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
         Begin MSDataListLib.DataCombo DcbUnit5 
            Height          =   315
            Left            =   6165
            TabIndex        =   293
            Top             =   390
            Width           =   1080
            _ExtentX        =   1905
            _ExtentY        =   556
            _Version        =   393216
            Text            =   ""
            RightToLeft     =   -1  'True
         End
         Begin MSDataListLib.DataCombo DcboItemID5 
            Height          =   315
            Left            =   8460
            TabIndex        =   294
            Top             =   390
            Width           =   2775
            _ExtentX        =   4895
            _ExtentY        =   556
            _Version        =   393216
            Text            =   ""
            RightToLeft     =   -1  'True
         End
         Begin MSDataListLib.DataCombo XPCboGroup5 
            Height          =   315
            Left            =   12465
            TabIndex        =   295
            Top             =   390
            Width           =   1650
            _ExtentX        =   2910
            _ExtentY        =   556
            _Version        =   393216
            Text            =   ""
            RightToLeft     =   -1  'True
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "«·ﬂ„Ì…"
            Height          =   255
            Index           =   92
            Left            =   4320
            RightToLeft     =   -1  'True
            TabIndex        =   299
            Top             =   495
            Width           =   690
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "«·ÊÕœÂ"
            Height          =   300
            Index           =   91
            Left            =   7095
            RightToLeft     =   -1  'True
            TabIndex        =   298
            Top             =   390
            Width           =   675
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "  À»Ì "
            Height          =   285
            Index           =   90
            Left            =   11115
            RightToLeft     =   -1  'True
            TabIndex        =   297
            Top             =   495
            Width           =   765
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackColor       =   &H00E2E9E9&
            Caption         =   " «·„Ã„Ê⁄…"
            Height          =   330
            Index           =   89
            Left            =   14130
            TabIndex        =   296
            Top             =   435
            Width           =   855
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "«·’«›Ï »⁄œ «·ﬁÌ„… «·„÷«›…"
            Height          =   300
            Index           =   99
            Left            =   3885
            RightToLeft     =   -1  'True
            TabIndex        =   188
            Top             =   750
            Width           =   2160
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "«·ﬁÌ„… «·„÷«›…"
            Height          =   300
            Index           =   98
            Left            =   7185
            RightToLeft     =   -1  'True
            TabIndex        =   187
            Top             =   780
            Width           =   1125
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "«·’«›Ì"
            Height          =   330
            Index           =   56
            Left            =   9360
            RightToLeft     =   -1  'True
            TabIndex        =   173
            Top             =   795
            Width           =   1125
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "«Ã„«·Ì «·Œ’Ê„« "
            Height          =   285
            Index           =   55
            Left            =   11715
            RightToLeft     =   -1  'True
            TabIndex        =   171
            Top             =   780
            Width           =   1575
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "«Ã„«·Ì «·«÷«›« "
            Height          =   315
            Index           =   53
            Left            =   14565
            RightToLeft     =   -1  'True
            TabIndex        =   166
            Top             =   765
            Width           =   1485
         End
         Begin VB.Label LblItemsCount 
            Alignment       =   2  'Center
            BackColor       =   &H00404040&
            ForeColor       =   &H0000FFFF&
            Height          =   450
            Left            =   30
            RightToLeft     =   -1  'True
            TabIndex        =   32
            Top             =   3420
            Width           =   495
         End
      End
      Begin C1SizerLibCtl.C1Elastic ELe 
         Height          =   2385
         Index           =   0
         Left            =   -240
         TabIndex        =   33
         TabStop         =   0   'False
         Top             =   750
         Width           =   16770
         _cx             =   29580
         _cy             =   4207
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
         Begin VB.TextBox txtOrderID 
            Alignment       =   1  'Right Justify
            Height          =   375
            Left            =   0
            RightToLeft     =   -1  'True
            TabIndex        =   302
            Top             =   0
            Visible         =   0   'False
            Width           =   675
         End
         Begin VB.TextBox TXT_order_no 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   13740
            RightToLeft     =   -1  'True
            TabIndex        =   300
            Top             =   780
            Width           =   1725
         End
         Begin VB.TextBox TxtSearchCode 
            Alignment       =   1  'Right Justify
            Height          =   315
            Left            =   10500
            RightToLeft     =   -1  'True
            TabIndex        =   285
            Top             =   495
            Width           =   1335
         End
         Begin VB.ComboBox CboPayMentType 
            Height          =   315
            Left            =   660
            RightToLeft     =   -1  'True
            Style           =   2  'Dropdown List
            TabIndex        =   270
            Top             =   1035
            Width           =   1560
         End
         Begin VB.TextBox txtCustomerName 
            Alignment       =   1  'Right Justify
            Height          =   360
            Left            =   1560
            TabIndex        =   246
            Top             =   360
            Width           =   2220
         End
         Begin VB.CommandButton cmdAddCustomer 
            Caption         =   "«÷«›… ⁄„Ì· ÃœÌœ"
            Height          =   360
            Left            =   210
            RightToLeft     =   -1  'True
            TabIndex        =   243
            Top             =   360
            Width           =   1335
         End
         Begin VB.TextBox TxtPhone 
            Alignment       =   1  'Right Justify
            Height          =   360
            Left            =   4710
            TabIndex        =   241
            Top             =   360
            Width           =   2220
         End
         Begin VB.TextBox TxtEmployeeID 
            Alignment       =   1  'Right Justify
            Height          =   330
            Left            =   5595
            TabIndex        =   189
            Top             =   735
            Width           =   1305
         End
         Begin VB.TextBox txtFile 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFFF&
            Height          =   315
            Left            =   0
            RightToLeft     =   -1  'True
            TabIndex        =   158
            Top             =   -150
            Visible         =   0   'False
            Width           =   1500
         End
         Begin VB.TextBox TxtMaxName 
            Alignment       =   1  'Right Justify
            Height          =   360
            Left            =   7605
            MultiLine       =   -1  'True
            RightToLeft     =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   155
            Top             =   1005
            Width           =   4935
         End
         Begin VB.TextBox TxtSearchCode2 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   10890
            RightToLeft     =   -1  'True
            TabIndex        =   131
            Top             =   585
            Visible         =   0   'False
            Width           =   1290
         End
         Begin VB.TextBox TxtMaxNo 
            Alignment       =   1  'Right Justify
            Height          =   315
            Left            =   13740
            RightToLeft     =   -1  'True
            TabIndex        =   0
            Top             =   1065
            Width           =   1725
         End
         Begin VB.TextBox TxtTransSerial 
            Alignment       =   1  'Right Justify
            Height          =   405
            Left            =   14415
            Locked          =   -1  'True
            RightToLeft     =   -1  'True
            TabIndex        =   97
            Top             =   60
            Width           =   1065
         End
         Begin VB.TextBox TxtManualNo1 
            Alignment       =   1  'Right Justify
            Height          =   300
            Left            =   9270
            RightToLeft     =   -1  'True
            TabIndex        =   96
            Top             =   -345
            Width           =   1905
         End
         Begin VB.Frame Frame1 
            Height          =   2025
            Left            =   18360
            RightToLeft     =   -1  'True
            TabIndex        =   68
            Top             =   495
            Width           =   15615
            Begin VB.TextBox TxtProductionPlanno 
               Alignment       =   1  'Right Justify
               Height          =   330
               Left            =   120
               RightToLeft     =   -1  'True
               TabIndex        =   73
               Top             =   120
               Width           =   1425
            End
            Begin VB.ComboBox CboPayMentTypess 
               Height          =   315
               Index           =   0
               Left            =   13680
               RightToLeft     =   -1  'True
               Style           =   2  'Dropdown List
               TabIndex        =   72
               Top             =   600
               Width           =   2145
            End
            Begin VB.TextBox TxtShipmentArae 
               Alignment       =   1  'Right Justify
               Height          =   285
               Left            =   14280
               RightToLeft     =   -1  'True
               TabIndex        =   71
               Top             =   600
               Width           =   3735
            End
            Begin VB.CheckBox chkshipped 
               Alignment       =   1  'Right Justify
               Caption         =   " „ «·‘Õ‰"
               Height          =   195
               Left            =   120
               RightToLeft     =   -1  'True
               TabIndex        =   70
               Top             =   -2760
               Width           =   975
            End
            Begin VB.TextBox TxtWorkHour 
               Alignment       =   1  'Right Justify
               Height          =   330
               Left            =   5040
               RightToLeft     =   -1  'True
               TabIndex        =   69
               Top             =   1680
               Visible         =   0   'False
               Width           =   2145
            End
            Begin MSDataListLib.DataCombo Dccurrency 
               Height          =   315
               Left            =   15000
               TabIndex        =   74
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
               Left            =   13800
               TabIndex        =   75
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
               Left            =   13800
               TabIndex        =   76
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
               Left            =   14760
               TabIndex        =   77
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
               Left            =   13800
               TabIndex        =   78
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
               TabIndex        =   79
               Top             =   2040
               Width           =   1905
               _ExtentX        =   3360
               _ExtentY        =   556
               _Version        =   393216
               ListField       =   "6"
               Text            =   ""
               RightToLeft     =   -1  'True
            End
            Begin MSComCtl2.DTPicker startDate 
               Height          =   315
               Left            =   1800
               TabIndex        =   80
               Top             =   840
               Width           =   2175
               _ExtentX        =   3836
               _ExtentY        =   556
               _Version        =   393216
               Format          =   140902401
               CurrentDate     =   38784
            End
            Begin MSComCtl2.DTPicker EndDate 
               Height          =   315
               Left            =   1800
               TabIndex        =   81
               Top             =   1200
               Width           =   2175
               _ExtentX        =   3836
               _ExtentY        =   556
               _Version        =   393216
               Format          =   140902401
               CurrentDate     =   38784
            End
            Begin MSComCtl2.DTPicker startTime 
               Height          =   285
               Left            =   120
               TabIndex        =   82
               Top             =   840
               Visible         =   0   'False
               Width           =   1575
               _ExtentX        =   2778
               _ExtentY        =   503
               _Version        =   393216
               CustomFormat    =   "'Time: 'hh:mm tt"
               Format          =   140902403
               UpDown          =   -1  'True
               CurrentDate     =   39240
            End
            Begin MSComCtl2.DTPicker EndTime 
               Height          =   285
               Left            =   120
               TabIndex        =   83
               Top             =   1200
               Visible         =   0   'False
               Width           =   1575
               _ExtentX        =   2778
               _ExtentY        =   503
               _Version        =   393216
               CustomFormat    =   "'Time: 'hh:mm tt"
               Format          =   140902403
               UpDown          =   -1  'True
               CurrentDate     =   39240
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "Œÿ… ≈‰ «Ã"
               ForeColor       =   &H00000000&
               Height          =   285
               Index           =   45
               Left            =   1560
               RightToLeft     =   -1  'True
               TabIndex        =   95
               Top             =   120
               Width           =   975
            End
            Begin VB.Label Label1 
               Alignment       =   1  'Right Justify
               Caption         =   "ÃÂ… «· ”·Ì„"
               ForeColor       =   &H00000000&
               Height          =   375
               Index           =   0
               Left            =   13920
               RightToLeft     =   -1  'True
               TabIndex        =   94
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
               TabIndex        =   93
               Top             =   2040
               Width           =   975
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               Caption         =   "«· ’‰Ì›"
               Height          =   285
               Index           =   16
               Left            =   13680
               RightToLeft     =   -1  'True
               TabIndex        =   92
               Top             =   240
               Width           =   1215
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               Caption         =   "ÿ—Ìﬁ… «·œ›⁄"
               ForeColor       =   &H00000000&
               Height          =   285
               Index           =   15
               Left            =   13800
               RightToLeft     =   -1  'True
               TabIndex        =   91
               Top             =   600
               Width           =   975
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               Caption         =   "ÿ—Ìﬁ… «·‘Õ‰"
               ForeColor       =   &H00000000&
               Height          =   285
               Index           =   14
               Left            =   13560
               RightToLeft     =   -1  'True
               TabIndex        =   90
               Top             =   240
               Width           =   1215
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               Caption         =   "«·»·œ"
               Height          =   285
               Index           =   13
               Left            =   14880
               RightToLeft     =   -1  'True
               TabIndex        =   89
               Top             =   960
               Width           =   1335
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               Caption         =   "«·⁄„·Â"
               Height          =   285
               Index           =   12
               Left            =   13680
               RightToLeft     =   -1  'True
               TabIndex        =   88
               Top             =   600
               Width           =   1335
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               BackStyle       =   0  'Transparent
               Caption         =   " «—ÌŒ  »œ«Ì… «·«‰ «Ã"
               ForeColor       =   &H00000000&
               Height          =   195
               Index           =   28
               Left            =   4080
               RightToLeft     =   -1  'True
               TabIndex        =   87
               Top             =   840
               Width           =   1530
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               BackStyle       =   0  'Transparent
               Caption         =   "„Œ“‰  «·«‰ «Ã «· «„"
               ForeColor       =   &H00000000&
               Height          =   240
               Index           =   34
               Left            =   3960
               RightToLeft     =   -1  'True
               TabIndex        =   86
               Top             =   480
               Width           =   1665
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               BackStyle       =   0  'Transparent
               Caption         =   " «—ÌŒ ‰Â«Ì… «·«‰ «Ã"
               ForeColor       =   &H00000000&
               Height          =   195
               Index           =   35
               Left            =   4080
               RightToLeft     =   -1  'True
               TabIndex        =   85
               Top             =   1200
               Width           =   1530
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "«Ã„«·Ì ”«⁄«  «·«” Â·«ﬂ ··Œÿ"
               ForeColor       =   &H00000000&
               Height          =   435
               Index           =   37
               Left            =   7560
               RightToLeft     =   -1  'True
               TabIndex        =   84
               Top             =   1680
               Visible         =   0   'False
               Width           =   1050
            End
         End
         Begin VB.TextBox TxtNoteSerial1 
            Alignment       =   1  'Right Justify
            Height          =   390
            Left            =   16785
            RightToLeft     =   -1  'True
            TabIndex        =   67
            Top             =   -210
            Visible         =   0   'False
            Width           =   2340
         End
         Begin VB.TextBox TXTNoteID 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   5700
            RightToLeft     =   -1  'True
            TabIndex        =   66
            Text            =   "Text4"
            Top             =   -1035
            Visible         =   0   'False
            Width           =   1035
         End
         Begin VB.Frame Frame3 
            Height          =   1935
            Left            =   17175
            RightToLeft     =   -1  'True
            TabIndex        =   53
            Top             =   1830
            Width           =   8655
            Begin VB.TextBox Text5 
               Alignment       =   1  'Right Justify
               Height          =   285
               Left            =   4080
               RightToLeft     =   -1  'True
               TabIndex        =   54
               Top             =   600
               Width           =   2295
            End
            Begin MSComCtl2.DTPicker DTPicker2 
               Height          =   315
               Left            =   120
               TabIndex        =   55
               Top             =   600
               Width           =   2100
               _ExtentX        =   3704
               _ExtentY        =   556
               _Version        =   393216
               Format          =   140902401
               CurrentDate     =   38784
            End
            Begin MSComCtl2.DTPicker DTPicker3 
               Height          =   315
               Left            =   4800
               TabIndex        =   56
               Top             =   960
               Width           =   1620
               _ExtentX        =   2858
               _ExtentY        =   556
               _Version        =   393216
               Format          =   140902401
               CurrentDate     =   38784
            End
            Begin MSComCtl2.DTPicker DTPicker4 
               Height          =   315
               Left            =   120
               TabIndex        =   57
               Top             =   960
               Width           =   2100
               _ExtentX        =   3704
               _ExtentY        =   556
               _Version        =   393216
               Format          =   140902401
               CurrentDate     =   38784
            End
            Begin MSComCtl2.DTPicker DTPicker5 
               Height          =   315
               Left            =   4800
               TabIndex        =   58
               Top             =   1320
               Width           =   1620
               _ExtentX        =   2858
               _ExtentY        =   556
               _Version        =   393216
               Format          =   140902401
               CurrentDate     =   38784
            End
            Begin MSComCtl2.DTPicker DTPicker6 
               Height          =   315
               Left            =   120
               TabIndex        =   59
               Top             =   1320
               Width           =   2100
               _ExtentX        =   3704
               _ExtentY        =   556
               _Version        =   393216
               Format          =   140902401
               CurrentDate     =   38784
            End
            Begin VB.Label Label8 
               Alignment       =   1  'Right Justify
               Caption         =   " «—ÌŒ «·Ê’Ê· «·„ Êﬁ⁄"
               Height          =   255
               Left            =   2280
               RightToLeft     =   -1  'True
               TabIndex        =   65
               Top             =   1440
               Width           =   1575
            End
            Begin VB.Label Label7 
               Alignment       =   1  'Right Justify
               Caption         =   " «—ÌŒ «· √ŒÌ—"
               Height          =   255
               Left            =   6480
               RightToLeft     =   -1  'True
               TabIndex        =   64
               Top             =   1440
               Width           =   1215
            End
            Begin VB.Label Label6 
               Alignment       =   1  'Right Justify
               Caption         =   "«· «—ÌŒ «·›⁄·Ì"
               Height          =   375
               Left            =   2640
               RightToLeft     =   -1  'True
               TabIndex        =   63
               Top             =   1200
               Width           =   1215
            End
            Begin VB.Label Label5 
               Alignment       =   1  'Right Justify
               Caption         =   "«· «—ÌŒ «·„ Êﬁ⁄"
               Height          =   375
               Left            =   6480
               RightToLeft     =   -1  'True
               TabIndex        =   62
               Top             =   1080
               Width           =   1215
            End
            Begin VB.Label Label3 
               Alignment       =   1  'Right Justify
               Caption         =   "«· «—ÌŒ"
               Height          =   375
               Left            =   2520
               RightToLeft     =   -1  'True
               TabIndex        =   61
               Top             =   720
               Width           =   1335
            End
            Begin VB.Label Label2 
               Alignment       =   1  'Right Justify
               Caption         =   "«·—ﬁ„"
               Height          =   375
               Left            =   6720
               RightToLeft     =   -1  'True
               TabIndex        =   60
               Top             =   720
               Width           =   975
            End
         End
         Begin VB.Frame Frame2 
            Height          =   1935
            Left            =   16785
            RightToLeft     =   -1  'True
            TabIndex        =   40
            Top             =   1830
            Width           =   6270
            Begin VB.TextBox Text7 
               Alignment       =   1  'Right Justify
               Height          =   285
               Left            =   5400
               RightToLeft     =   -1  'True
               TabIndex        =   43
               Top             =   600
               Width           =   3855
            End
            Begin VB.TextBox Text3 
               Alignment       =   1  'Right Justify
               Height          =   285
               Left            =   2640
               RightToLeft     =   -1  'True
               TabIndex        =   42
               Top             =   1320
               Width           =   1455
            End
            Begin VB.TextBox Text2 
               Alignment       =   1  'Right Justify
               Height          =   285
               Left            =   240
               RightToLeft     =   -1  'True
               TabIndex        =   41
               Top             =   960
               Width           =   1335
            End
            Begin MSComCtl2.DTPicker DTPicker1 
               Height          =   315
               Left            =   240
               TabIndex        =   44
               Top             =   1320
               Width           =   1320
               _ExtentX        =   2328
               _ExtentY        =   556
               _Version        =   393216
               Format          =   140902401
               CurrentDate     =   38784
            End
            Begin MSDataListLib.DataCombo DataCombo9 
               Height          =   315
               Left            =   1920
               TabIndex        =   45
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
               TabIndex        =   46
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
               TabIndex        =   52
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
               TabIndex        =   51
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
               TabIndex        =   50
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
               TabIndex        =   49
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
               TabIndex        =   48
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
               TabIndex        =   47
               Top             =   240
               Width           =   1095
            End
         End
         Begin VB.ComboBox CboPriceType 
            Height          =   315
            Left            =   16875
            RightToLeft     =   -1  'True
            Style           =   2  'Dropdown List
            TabIndex        =   39
            Top             =   345
            Width           =   2460
         End
         Begin VB.TextBox TxtModFlg 
            Alignment       =   1  'Right Justify
            Height          =   405
            Left            =   3180
            RightToLeft     =   -1  'True
            TabIndex        =   38
            Top             =   -465
            Visible         =   0   'False
            Width           =   750
         End
         Begin VB.TextBox TxtFillData 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   2115
            RightToLeft     =   -1  'True
            TabIndex        =   37
            Top             =   -495
            Visible         =   0   'False
            Width           =   990
         End
         Begin VB.TextBox XPTxtBillID 
            Alignment       =   1  'Right Justify
            Height          =   315
            Left            =   30
            Locked          =   -1  'True
            RightToLeft     =   -1  'True
            TabIndex        =   36
            Top             =   -495
            Visible         =   0   'False
            Width           =   2070
         End
         Begin VB.TextBox txtShipmentPrice 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   12030
            RightToLeft     =   -1  'True
            TabIndex        =   35
            Top             =   -360
            Visible         =   0   'False
            Width           =   2415
         End
         Begin VB.CheckBox Selct 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "Ì „  Œ’Ì’ «·„ﬂÊ‰«  »‘ﬂ· ›⁄·Ì"
            Height          =   390
            Index           =   0
            Left            =   525
            RightToLeft     =   -1  'True
            TabIndex        =   34
            Top             =   15
            Width           =   3195
         End
         Begin MSDataListLib.DataCombo DCboStoreName1 
            Height          =   315
            Index           =   0
            Left            =   10830
            TabIndex        =   98
            Top             =   2655
            Visible         =   0   'False
            Width           =   2685
            _ExtentX        =   4736
            _ExtentY        =   556
            _Version        =   393216
            ListField       =   "7"
            Text            =   ""
            RightToLeft     =   -1  'True
         End
         Begin MSComCtl2.DTPicker XPDtbBill 
            Height          =   405
            Left            =   12465
            TabIndex        =   99
            Top             =   60
            Width           =   1335
            _ExtentX        =   2355
            _ExtentY        =   714
            _Version        =   393216
            Format          =   142802945
            CurrentDate     =   38784
         End
         Begin ImpulseButton.ISButton XPBtnNewClients 
            Height          =   405
            Left            =   6885
            TabIndex        =   100
            TabStop         =   0   'False
            Top             =   2430
            Width           =   75
            _ExtentX        =   132
            _ExtentY        =   714
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
            ButtonImage     =   "FrmDefinCompItem.frx":4E99
            ColorButton     =   14871017
            ColorHighlight  =   16777215
            ColorShadow     =   -2147483631
            ColorOutline    =   -2147483631
            DrawFocusRectangle=   0   'False
         End
         Begin ImpulseButton.ISButton CmdTemplate 
            Height          =   510
            Left            =   3810
            TabIndex        =   101
            Top             =   -1605
            Visible         =   0   'False
            Width           =   1935
            _ExtentX        =   3413
            _ExtentY        =   900
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
         Begin C1SizerLibCtl.C1Elastic ELe 
            Height          =   645
            Index           =   4
            Left            =   5820
            TabIndex        =   102
            TabStop         =   0   'False
            Top             =   -2010
            Width           =   4185
            _cx             =   7382
            _cy             =   1138
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
               TabIndex        =   104
               Top             =   210
               Width           =   1815
            End
            Begin VB.TextBox XPTxtTaxValue 
               Alignment       =   1  'Right Justify
               Height          =   390
               Left            =   30
               RightToLeft     =   -1  'True
               TabIndex        =   103
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
               TabIndex        =   105
               Top             =   285
               Width           =   720
            End
         End
         Begin MSDataListLib.DataCombo DataCombo1 
            Height          =   315
            Left            =   16875
            TabIndex        =   106
            Top             =   750
            Width           =   2460
            _ExtentX        =   4339
            _ExtentY        =   556
            _Version        =   393216
            ListField       =   "6"
            Text            =   ""
            RightToLeft     =   -1  'True
         End
         Begin ImpulseButton.ISButton CmdConvert 
            Height          =   315
            Left            =   12405
            TabIndex        =   107
            Top             =   3705
            Visible         =   0   'False
            Width           =   2385
            _ExtentX        =   4207
            _ExtentY        =   556
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
         Begin MSDataListLib.DataCombo Dcbranch 
            Height          =   315
            Left            =   10125
            TabIndex        =   108
            Top             =   60
            Width           =   1875
            _ExtentX        =   3307
            _ExtentY        =   556
            _Version        =   393216
            Text            =   ""
            RightToLeft     =   -1  'True
         End
         Begin MSDataListLib.DataCombo DCboStoreName 
            Height          =   315
            Left            =   12495
            TabIndex        =   109
            Top             =   495
            Width           =   2970
            _ExtentX        =   5239
            _ExtentY        =   556
            _Version        =   393216
            Text            =   ""
            RightToLeft     =   -1  'True
         End
         Begin MSDataListLib.DataCombo DBCboClientName 
            Height          =   315
            Left            =   7650
            TabIndex        =   1
            Top             =   495
            Width           =   2880
            _ExtentX        =   5080
            _ExtentY        =   556
            _Version        =   393216
            ListField       =   "6"
            Text            =   ""
            RightToLeft     =   -1  'True
         End
         Begin C1SizerLibCtl.C1Elastic ELe 
            Height          =   960
            Index           =   2
            Left            =   270
            TabIndex        =   135
            TabStop         =   0   'False
            Top             =   1410
            Width           =   16425
            _cx             =   28972
            _cy             =   1693
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
            Begin VB.TextBox txtDI 
               Alignment       =   1  'Right Justify
               Height          =   270
               Left            =   3630
               RightToLeft     =   -1  'True
               TabIndex        =   290
               Text            =   "1"
               Top             =   270
               Width           =   720
            End
            Begin VB.TextBox txtDO 
               Alignment       =   1  'Right Justify
               Height          =   270
               Left            =   4350
               RightToLeft     =   -1  'True
               TabIndex        =   288
               Text            =   "1"
               Top             =   270
               Width           =   690
            End
            Begin VB.CommandButton cmdRecalc 
               Caption         =   "÷»ÿ «· ﬂ«·Ì›"
               Height          =   285
               Left            =   420
               TabIndex        =   281
               Top             =   525
               Visible         =   0   'False
               Width           =   1215
            End
            Begin VB.TextBox TxtMaxNo2 
               Alignment       =   1  'Right Justify
               Height          =   345
               Left            =   2970
               RightToLeft     =   -1  'True
               TabIndex        =   279
               Top             =   525
               Width           =   1725
            End
            Begin VB.TextBox txtthickness 
               Alignment       =   1  'Right Justify
               Height          =   270
               Left            =   5040
               RightToLeft     =   -1  'True
               TabIndex        =   276
               Text            =   "1"
               Top             =   270
               Width           =   720
            End
            Begin VB.TextBox txtDiameter 
               Alignment       =   1  'Right Justify
               Height          =   270
               Left            =   2940
               RightToLeft     =   -1  'True
               TabIndex        =   274
               Text            =   "1"
               Top             =   270
               Width           =   720
            End
            Begin VB.CommandButton Command1 
               Caption         =   "⁄—÷ «·’Ê—…"
               Height          =   330
               Left            =   9660
               RightToLeft     =   -1  'True
               TabIndex        =   271
               Top             =   -60
               Width           =   1215
            End
            Begin VB.TextBox txtLength 
               Alignment       =   1  'Right Justify
               Height          =   270
               Left            =   5760
               RightToLeft     =   -1  'True
               TabIndex        =   259
               Text            =   "1"
               Top             =   270
               Width           =   690
            End
            Begin VB.TextBox txtItemCodeBuiltin 
               Alignment       =   1  'Right Justify
               Height          =   285
               Left            =   10890
               RightToLeft     =   -1  'True
               TabIndex        =   253
               Top             =   615
               Visible         =   0   'False
               Width           =   1725
            End
            Begin VB.TextBox txtTotal 
               Alignment       =   1  'Right Justify
               Height          =   270
               Left            =   465
               RightToLeft     =   -1  'True
               TabIndex        =   181
               Text            =   "1"
               Top             =   270
               Width           =   1035
            End
            Begin VB.TextBox txtPrice 
               Alignment       =   1  'Right Justify
               Height          =   270
               Left            =   1500
               RightToLeft     =   -1  'True
               TabIndex        =   167
               Text            =   "1"
               Top             =   270
               Width           =   915
            End
            Begin VB.TextBox txthight 
               Alignment       =   1  'Right Justify
               Height          =   270
               Left            =   6450
               RightToLeft     =   -1  'True
               TabIndex        =   163
               Text            =   "1"
               Top             =   270
               Width           =   690
            End
            Begin VB.TextBox txtwidtj 
               Alignment       =   1  'Right Justify
               Height          =   270
               Left            =   7155
               RightToLeft     =   -1  'True
               TabIndex        =   161
               Text            =   "1"
               Top             =   270
               Width           =   735
            End
            Begin VB.TextBox txtQty1 
               Alignment       =   1  'Right Justify
               Height          =   270
               Left            =   2385
               RightToLeft     =   -1  'True
               TabIndex        =   149
               Text            =   "1"
               Top             =   270
               Width           =   585
            End
            Begin VB.TextBox TxtAttachedItemCode 
               Alignment       =   1  'Right Justify
               Height          =   300
               Left            =   12150
               TabIndex        =   2
               Top             =   300
               Width           =   1500
            End
            Begin MSDataListLib.DataCombo DcbUnit 
               Height          =   315
               Left            =   7905
               TabIndex        =   3
               Top             =   270
               Width           =   945
               _ExtentX        =   1667
               _ExtentY        =   556
               _Version        =   393216
               Text            =   ""
               RightToLeft     =   -1  'True
            End
            Begin ImpulseButton.ISButton cmdAdd 
               Default         =   -1  'True
               Height          =   1200
               Left            =   -90
               TabIndex        =   146
               Top             =   -390
               Visible         =   0   'False
               Width           =   690
               _ExtentX        =   1217
               _ExtentY        =   2117
               ButtonStyle     =   1
               ButtonPositionImage=   2
               Caption         =   "≈÷«›…"
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
               ButtonImage     =   "FrmDefinCompItem.frx":5233
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
            Begin MSDataListLib.DataCombo DcboItemID1 
               Height          =   315
               Left            =   9330
               TabIndex        =   156
               Top             =   285
               Width           =   2775
               _ExtentX        =   4895
               _ExtentY        =   556
               _Version        =   393216
               Text            =   ""
               RightToLeft     =   -1  'True
            End
            Begin MSDataListLib.DataCombo XPCboGroup 
               Height          =   315
               Left            =   14280
               TabIndex        =   219
               Top             =   285
               Width           =   1155
               _ExtentX        =   2037
               _ExtentY        =   556
               _Version        =   393216
               Text            =   ""
               RightToLeft     =   -1  'True
            End
            Begin MSDataListLib.DataCombo DcboBuiltinItemID 
               Height          =   315
               Left            =   6690
               TabIndex        =   248
               Top             =   615
               Visible         =   0   'False
               Width           =   3315
               _ExtentX        =   5847
               _ExtentY        =   556
               _Version        =   393216
               Text            =   ""
               RightToLeft     =   -1  'True
            End
            Begin MSDataListLib.DataCombo XPCboGroupBuiltin 
               Height          =   315
               Left            =   13770
               TabIndex        =   249
               Top             =   585
               Visible         =   0   'False
               Width           =   1665
               _ExtentX        =   2937
               _ExtentY        =   556
               _Version        =   393216
               Text            =   ""
               RightToLeft     =   -1  'True
            End
            Begin MSDataListLib.DataCombo cmbSpecification 
               Bindings        =   "FrmDefinCompItem.frx":55CD
               Height          =   315
               Left            =   10920
               TabIndex        =   272
               Top             =   -15
               Width           =   4485
               _ExtentX        =   7911
               _ExtentY        =   556
               _Version        =   393216
               Style           =   2
               BackColor       =   16777215
               ListField       =   "account_name"
               BoundColumn     =   "code"
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
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "DI"
               Height          =   300
               Index           =   88
               Left            =   3705
               RightToLeft     =   -1  'True
               TabIndex        =   291
               Top             =   -30
               Width           =   450
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "DO"
               Height          =   300
               Index           =   87
               Left            =   4395
               RightToLeft     =   -1  'True
               TabIndex        =   289
               Top             =   -30
               Width           =   450
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "»‰«¡« ⁄·Ï „ﬂ” —ﬁ„"
               Height          =   255
               Index           =   86
               Left            =   4650
               RightToLeft     =   -1  'True
               TabIndex        =   278
               Top             =   615
               Width           =   1905
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "«·”„ﬂ"
               Height          =   300
               Index           =   85
               Left            =   5085
               RightToLeft     =   -1  'True
               TabIndex        =   277
               Top             =   -30
               Width           =   450
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "«·ﬁÿ—"
               Height          =   300
               Index           =   83
               Left            =   3015
               RightToLeft     =   -1  'True
               TabIndex        =   275
               Top             =   -30
               Width           =   450
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "«·›—“"
               Height          =   315
               Index           =   64
               Left            =   15090
               RightToLeft     =   -1  'True
               TabIndex        =   273
               Top             =   -45
               Width           =   1215
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "«·⁄„ﬁ"
               Height          =   300
               Index           =   82
               Left            =   5835
               RightToLeft     =   -1  'True
               TabIndex        =   260
               Top             =   -30
               Width           =   420
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "ﬂÊœ  «·’‰›"
               Height          =   255
               Index           =   79
               Left            =   12210
               RightToLeft     =   -1  'True
               TabIndex        =   252
               Top             =   645
               Visible         =   0   'False
               Width           =   1455
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   " «·’‰›"
               Height          =   255
               Index           =   78
               Left            =   9975
               RightToLeft     =   -1  'True
               TabIndex        =   251
               Top             =   645
               Visible         =   0   'False
               Width           =   765
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "«·„Ã„Ê⁄…"
               Height          =   540
               Index           =   77
               Left            =   15585
               TabIndex        =   250
               Top             =   585
               Visible         =   0   'False
               Width           =   765
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "«·„Ã„Ê⁄…"
               Height          =   330
               Index           =   51
               Left            =   15585
               TabIndex        =   220
               Top             =   270
               Width           =   765
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "«·«Ã„«·Ì"
               Height          =   300
               Index           =   57
               Left            =   675
               RightToLeft     =   -1  'True
               TabIndex        =   182
               Top             =   -30
               Width           =   615
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "«·”⁄—"
               Height          =   300
               Index           =   52
               Left            =   1740
               RightToLeft     =   -1  'True
               TabIndex        =   168
               Top             =   -30
               Width           =   435
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "«— ›«⁄L "
               Height          =   300
               Index           =   49
               Left            =   6225
               RightToLeft     =   -1  'True
               TabIndex        =   164
               Top             =   -30
               Width           =   840
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "«·⁄—÷ W"
               Height          =   300
               Index           =   41
               Left            =   7155
               RightToLeft     =   -1  'True
               TabIndex        =   162
               Top             =   -30
               Width           =   675
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "«·ﬂ„Ì…"
               Height          =   300
               Index           =   40
               Left            =   2475
               RightToLeft     =   -1  'True
               TabIndex        =   150
               Top             =   -30
               Width           =   435
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "«·„ﬂÊ‰« "
               BeginProperty Font 
                  Name            =   "Times New Roman"
                  Size            =   14.25
                  Charset         =   178
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00FF0000&
               Height          =   765
               Index           =   65
               Left            =   12645
               RightToLeft     =   -1  'True
               TabIndex        =   139
               Top             =   1980
               Width           =   3705
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "«·ÊÕœÂ"
               Height          =   285
               Index           =   27
               Left            =   8715
               RightToLeft     =   -1  'True
               TabIndex        =   138
               Top             =   270
               Width           =   615
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   " «·’‰›"
               Height          =   240
               Index           =   26
               Left            =   13440
               RightToLeft     =   -1  'True
               TabIndex        =   137
               Top             =   300
               Width           =   765
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "ﬂÊœ «·’‰›"
               Height          =   75
               Index           =   17
               Left            =   12405
               RightToLeft     =   -1  'True
               TabIndex        =   136
               Top             =   -525
               Visible         =   0   'False
               Width           =   765
            End
         End
         Begin ImpulseButton.ISButton ISButton3 
            Height          =   345
            Left            =   3780
            TabIndex        =   159
            Top             =   0
            Width           =   1275
            _ExtentX        =   2249
            _ExtentY        =   609
            Caption         =   "«” Ì—«œ «·„·›"
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
            ButtonImage     =   "FrmDefinCompItem.frx":55E2
            ColorButton     =   14871017
            DrawFocusRectangle=   0   'False
            DisabledImageExtraction=   0
            LowerToggledContent=   0   'False
         End
         Begin ImpulseButton.ISButton ISButton4 
            Height          =   345
            Left            =   5085
            TabIndex        =   160
            ToolTipText     =   "«÷«›… «·»Ì«‰«  «·Ï «·œ« «"
            Top             =   0
            Width           =   1200
            _ExtentX        =   2117
            _ExtentY        =   609
            Caption         =   "Õœœ «·„”«—"
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
            ButtonImage     =   "FrmDefinCompItem.frx":BE44
            ColorButton     =   14871017
            DrawFocusRectangle=   0   'False
            DisabledImageExtraction=   0
            LowerToggledContent=   0   'False
         End
         Begin MSComDlg.CommonDialog CD1 
            Left            =   0
            Top             =   0
            _ExtentX        =   847
            _ExtentY        =   847
            _Version        =   393216
         End
         Begin MSDataListLib.DataCombo DcboEmp 
            Height          =   315
            Left            =   645
            TabIndex        =   190
            Top             =   735
            Width           =   4860
            _ExtentX        =   8573
            _ExtentY        =   556
            _Version        =   393216
            ListField       =   "7"
            Text            =   ""
            RightToLeft     =   -1  'True
         End
         Begin MSDataListLib.DataCombo DcboBox 
            Height          =   315
            Left            =   3360
            TabIndex        =   196
            Top             =   1095
            Width           =   3510
            _ExtentX        =   6191
            _ExtentY        =   556
            _Version        =   393216
            Text            =   ""
            RightToLeft     =   -1  'True
         End
         Begin MSComCtl2.DTPicker XPDtRecDate 
            Height          =   330
            Left            =   8220
            TabIndex        =   255
            Top             =   30
            Width           =   1335
            _ExtentX        =   2355
            _ExtentY        =   582
            _Version        =   393216
            Format          =   139132929
            CurrentDate     =   38784
         End
         Begin MSComCtl2.DTPicker txtRecTime 
            Height          =   315
            Left            =   6300
            TabIndex        =   257
            Top             =   0
            Width           =   1245
            _ExtentX        =   2196
            _ExtentY        =   556
            _Version        =   393216
            Format          =   139132930
            CurrentDate     =   41640
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "⁄—÷ ”⁄—"
            Height          =   375
            Index           =   93
            Left            =   15600
            RightToLeft     =   -1  'True
            TabIndex        =   301
            Top             =   780
            Width           =   1050
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "Êﬁ  «· ”·Ì„"
            ForeColor       =   &H00000000&
            Height          =   405
            Index           =   81
            Left            =   7440
            RightToLeft     =   -1  'True
            TabIndex        =   258
            Top             =   0
            Width           =   660
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   " «—ÌŒ «· ”·Ì„"
            ForeColor       =   &H00000000&
            Height          =   420
            Index           =   80
            Left            =   9390
            RightToLeft     =   -1  'True
            TabIndex        =   256
            Top             =   30
            Width           =   660
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "«”„ «·⁄„Ì·"
            Height          =   330
            Index           =   76
            Left            =   3900
            TabIndex        =   247
            Top             =   390
            Width           =   735
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   " ·Ì›Ê‰"
            Height          =   330
            Index           =   84
            Left            =   6900
            TabIndex        =   242
            Top             =   420
            Width           =   645
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "ÿ—Ìﬁ… «·œ›⁄"
            Height          =   330
            Index           =   54
            Left            =   2010
            RightToLeft     =   -1  'True
            TabIndex        =   198
            Top             =   1065
            Width           =   1140
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "«·’‰œÊﬁ"
            Height          =   330
            Index           =   59
            Left            =   6900
            TabIndex        =   197
            Top             =   1125
            Width           =   675
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "«·»«∆⁄"
            Height          =   225
            Index           =   58
            Left            =   6330
            TabIndex        =   191
            Top             =   780
            Width           =   1185
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            BackStyle       =   0  'Transparent
            Caption         =   "«·⁄„Ì·"
            ForeColor       =   &H00000000&
            Height          =   285
            Index           =   42
            Left            =   11190
            RightToLeft     =   -1  'True
            TabIndex        =   134
            Top             =   525
            Width           =   1170
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "ﬂÊœ «·„ﬂ”"
            ForeColor       =   &H00000000&
            Height          =   240
            Index           =   29
            Left            =   15555
            RightToLeft     =   -1  'True
            TabIndex        =   133
            Top             =   1125
            Width           =   1065
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "«”„ «·„ﬂ”"
            Height          =   270
            Index           =   30
            Left            =   12375
            RightToLeft     =   -1  'True
            TabIndex        =   132
            Top             =   1035
            Width           =   1125
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            BackStyle       =   0  'Transparent
            Caption         =   "—ﬁ„ ÌœÊÌ"
            ForeColor       =   &H00000000&
            Height          =   285
            Index           =   46
            Left            =   11355
            RightToLeft     =   -1  'True
            TabIndex        =   119
            Top             =   -345
            Width           =   885
         End
         Begin VB.Shape Shape2 
            BorderWidth     =   2
            Height          =   810
            Left            =   16785
            Top             =   345
            Visible         =   0   'False
            Width           =   4440
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "„·«ÕŸ… Â«„…:-"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   178
               Weight          =   700
               Underline       =   -1  'True
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000080&
            Height          =   345
            Index           =   43
            Left            =   17310
            RightToLeft     =   -1  'True
            TabIndex        =   118
            Top             =   345
            Visible         =   0   'False
            Width           =   1395
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "«·›—⁄"
            ForeColor       =   &H00000000&
            Height          =   285
            Index           =   36
            Left            =   11700
            RightToLeft     =   -1  'True
            TabIndex        =   117
            Top             =   60
            Width           =   660
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "„—ﬂ“ «· ﬂ·›…"
            Height          =   345
            Index           =   10
            Left            =   18030
            RightToLeft     =   -1  'True
            TabIndex        =   116
            Top             =   750
            Width           =   1185
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "‰Ê⁄ «·«„—"
            Height          =   345
            Index           =   9
            Left            =   18150
            RightToLeft     =   -1  'True
            TabIndex        =   115
            Top             =   345
            Width           =   945
         End
         Begin VB.Label lbl 
            BackColor       =   &H00E2E9E9&
            BackStyle       =   0  'Transparent
            Caption         =   "—ﬁ„ «·”‰œ"
            ForeColor       =   &H00000000&
            Height          =   315
            Index           =   5
            Left            =   15900
            RightToLeft     =   -1  'True
            TabIndex        =   114
            Top             =   60
            Width           =   795
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "«· «—ÌŒ"
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   6
            Left            =   13560
            RightToLeft     =   -1  'True
            TabIndex        =   113
            Top             =   60
            Width           =   735
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "«·⁄„Ì· / «·„Ê—œ"
            ForeColor       =   &H00000000&
            Height          =   330
            Index           =   7
            Left            =   16860
            RightToLeft     =   -1  'True
            TabIndex        =   112
            Top             =   1035
            Width           =   1170
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "«”„ «·„Œ“‰"
            Height          =   285
            Index           =   8
            Left            =   13710
            RightToLeft     =   -1  'True
            TabIndex        =   111
            Top             =   3540
            Visible         =   0   'False
            Width           =   1065
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "«·„Œ“‰"
            Height          =   375
            Index           =   50
            Left            =   15975
            RightToLeft     =   -1  'True
            TabIndex        =   110
            Top             =   465
            Width           =   600
         End
      End
      Begin C1SizerLibCtl.C1Elastic ELe 
         Height          =   1530
         Index           =   1
         Left            =   -270
         TabIndex        =   120
         TabStop         =   0   'False
         Top             =   9195
         Width           =   17685
         _cx             =   31194
         _cy             =   2699
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
         Begin VB.CheckBox chkHiddLogo 
            Caption         =   "«Œ›«¡ «··ÊÃÊ"
            Height          =   225
            Left            =   14400
            TabIndex        =   287
            Top             =   240
            Width           =   2175
         End
         Begin ImpulseButton.ISButton Cmd 
            Height          =   390
            Index           =   0
            Left            =   13050
            TabIndex        =   121
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
            Left            =   11850
            TabIndex        =   122
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
            Left            =   10335
            TabIndex        =   123
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
            Left            =   9210
            TabIndex        =   124
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
            Left            =   7800
            TabIndex        =   125
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
            Left            =   6765
            TabIndex        =   126
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
            Left            =   1680
            TabIndex        =   127
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
            Left            =   5280
            TabIndex        =   128
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
            Left            =   3945
            TabIndex        =   129
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
            Left            =   360
            TabIndex        =   130
            Top             =   120
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
         Begin ImpulseButton.ISButton Cmd 
            Height          =   390
            Index           =   10
            Left            =   2760
            TabIndex        =   157
            Top             =   90
            Width           =   930
            _ExtentX        =   1640
            _ExtentY        =   688
            ButtonStyle     =   1
            ButtonPositionImage=   1
            Caption         =   "‰”Œ… „„«À·…"
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
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "„œ… «·⁄—÷"
         Height          =   270
         Index           =   75
         Left            =   2130
         TabIndex        =   245
         Top             =   3210
         Width           =   750
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H00E2E9E9&
         Caption         =   " «·»Ì«‰"
         Height          =   195
         Index           =   74
         Left            =   14415
         TabIndex        =   239
         Top             =   3270
         Width           =   420
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "ﬁÌ„…"
         Height          =   270
         Index           =   68
         Left            =   4800
         TabIndex        =   226
         Top             =   3240
         Width           =   1080
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
         ForeColor       =   &H000000FF&
         Height          =   270
         Index           =   67
         Left            =   3150
         TabIndex        =   225
         Top             =   3240
         Width           =   345
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "‰Ê⁄ «·Œ’„"
         Height          =   270
         Index           =   66
         Left            =   7140
         TabIndex        =   223
         Top             =   3165
         Width           =   1080
      End
   End
End
Attribute VB_Name = "FrmDefinCompItem"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rs As ADODB.Recordset
Dim TTP As clstooltip
Dim CostTOTAL As Double
Dim StrTempAccountCode As String
Dim Account_Code_dynamic As String
Dim CostAccount As String
Dim StoreAccount As String
Dim mQtyTotal As Double
Dim mTotalSecond As Double
Dim cSearchDcbo(4)   As clsDCboSearch
Dim Dcombos As New ClsDataCombos
Dim mNoteSerial14 As String
Dim mTransactionID4 As String

Dim mNoteSerial15 As String
Dim mTransactionID5 As String

Dim PercetageVat As Double
Dim mNewId As Long
Dim mIdDisplay As Long
Dim Msg As String
Dim rsDummy As New ADODB.Recordset
Dim BranchID As Double
Dim StoreId As Double
Public mCustId As Long
Dim mBranchIDReSave As Integer
Dim mIsFinishSave As Boolean
Dim IsSaveWithOutMsg As Boolean
Dim mIsStart As Boolean


Private Sub Command2_Click()
FixRowsLine True
End Sub

Private Sub txtPassword_Change()
If Trim(txtPassword) = "Salim2020" Then
    cmdReSave.Visible = True
    txtFromDateReSave.Visible = True
    txtToDateReSave.Visible = True
    chkIsBranch.Visible = True
Else
    cmdReSave.Visible = False
    txtFromDateReSave.Visible = False
    txtToDateReSave.Visible = False
   chkIsBranch.Visible = False
End If
txtFromDateReSave.value = Date
txtToDateReSave.value = Date
End Sub


Private Sub cmdReSave_Click()
'    Dim s As String
'    Dim rsDummy As ADODB.Recordset
'    mBranchIDReSave = 0
'    If chkIsBranch.value = vbChecked Then
'        mBranchIDReSave = dcBranch.BoundText
'    Else
'        mBranchIDReSave = 0
'    End If
'
'    XPBtnMove_Click (2)
'    DoEvents
' Dim i As Double
'        For i = 1 To rs.RecordCount
'
'            IsSaveWithOutMsg = True
'
'            Cmd_Click (1)
'            DoEvents
'            DoEvents
'            DoEvents
'
'NewGrid.updateProfit
'       NewGrid.Calculate 1, , , True
'           DoEvents
'            DoEvents
'            DoEvents
'
'            SaveData True
'     DoEvents
'            DoEvents
'            DoEvents
'
'
'        XPBtnMove_Click (0)
'
'    Next i
'    IsSaveWithOutMsg = False
'    MsgBox " „ «·Õ›Ÿ"


   Dim s As String
    Dim rsDummy As ADODB.Recordset
    Dim mBranchID As Integer
    
    
    
    XPBtnMove_Click (2)
    DoEvents
    
    
'    XPBtnMove_Click (1)
    DoEvents
    Dim i As Double
    
       Dim StrSQL As String
       StrSQL = "SELECT * FROM TblDefComItem WHERE "
       
        StrSQL = StrSQL & "   ( RecordDate >= " & SQLDate(txtFromDateReSave.value, True) & " and "
        StrSQL = StrSQL & "   RecordDate <=   " & SQLDate(txtToDateReSave.value, True) & " )"

    
        If chkIsBranch.value = vbChecked <> 0 Then
            StrSQL = StrSQL & "  and BranchID =   " & val(Me.dcBranch.BoundText)
        
          
           
            Me.dcBranch.Enabled = True
      
      
        End If
    
            StrSQL = StrSQL & " Order by RecordDate,Id"
          
                
            Set rsDummy = New ADODB.Recordset
            rsDummy.Open StrSQL, Cn, adOpenKeyset, adLockOptimistic, adCmdText


    
    Do While Not rsDummy.EOF
            IsSaveWithOutMsg = True
       
                If val(rsDummy!ID & "") = 25777 Then
                    IsSaveWithOutMsg = True
                End If
                Retrive val(rsDummy!ID & ""), True
       
                   DoEvents
                 DoEvents
                 TxtModFlg.Text = "E"
                 Me.DCboUserName.BoundText = user_id
                
            
            DoEvents
            DoEvents

            


            SaveData True
         
 
        DoEvents
        
       '   Cmd_Click (0)
          
        rsDummy.MoveNext
        DoEvents
    Loop
                 
 
  
    IsSaveWithOutMsg = False
    MsgBox " „ «·Õ›Ÿ"

End Sub

Private Sub chkSelectAll_Click(Index As Integer)
Dim i As Long
If Index = 1 Then
    For i = 1 To FG.Rows - 1
        FG.TextMatrix(i, FG.ColIndex("Select")) = IIf(chkSelectAll(Index), 1, 0)
    Next
Else
    For i = 1 To FG2.Rows - 1
        FG2.TextMatrix(i, FG2.ColIndex("Select")) = IIf(chkSelectAll(Index), 1, 0)
    Next
End If
End Sub

Private Sub cmdCancel2_Click()
Dim StrSQL As String
'
    StrSQL = "Delete Transactions Where Transaction_ID In (Select TransactionID4 From TblDefComItemData Where IDDefCIT = " & val(TxtTransSerial) & ")"
    Cn.Execute StrSQL, , adExecuteNoRecords
    StrSQL = "Delete Transaction_Details Where Transaction_ID In (Select TransactionID4 From TblDefComItemData Where IDDefCIT = " & val(TxtTransSerial) & ")"
    Cn.Execute StrSQL, , adExecuteNoRecords
'
'
'
'
    StrSQL = "UPDATE TblDefComItem SET    NoteSerial14='' ,TransactionID4 = 0 WHERE ID  =" & val(TxtTransSerial)
               Cn.Execute StrSQL


    StrSQL = "UPDATE TblDefComItemData SET    NoteSerial14='' ,TransactionID4 = 0 WHERE IDDefCIT  =" & val(TxtTransSerial)
               Cn.Execute StrSQL
      StrSQL = "SELECT * FROM TblDefComItem "
StrSQL = StrSQL & "  WHERE      BranchId in(" & Current_branchSql & ")"
    StrSQL = StrSQL + " Order By ID"
    Set rs = New ADODB.Recordset
    rs.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText

If SystemOptions.UserInterface = ArabicInterface Then
    MsgBox " „ «·€«¡ «„— «·«‰ «Ã"
Else
    MsgBox "Production order canceled"
End If
cmdCreateProduction.Enabled = True
cmdCancel2.Enabled = False
Retrive val(TxtTransSerial)
End Sub

Private Sub cmdCreateProduction_Click()

    If Not SystemOptions.IsMultiItemsInCompItem Then
        If val(TxtNoteSerial14) <> 0 Then
            'TXTTransactionID4 = val(fg2.TextMatrix(i, fg2.ColIndex("TransactionID4")))
            FrmProductionOrder.show
            FrmProductionOrder.XPBtnMove_Click (2)
            FrmProductionOrder.Retrive val(TXTTransactionID4.Text)
            Exit Sub
        End If
    
    
    Else
        If val(TxtNoteSerial14) <> 0 Then
            
            FrmProductionOrder.show
            FrmProductionOrder.XPBtnMove_Click (2)
            FrmProductionOrder.Retrive val(TXTTransactionID4.Text)
            Exit Sub
        End If
    End If
    If DBCboClientName.BoundText = "" Then
        If SystemOptions.UserInterface = ArabicInterface Then
            MsgBox ("·« Ì„ﬂ‰ «‰‘«¡ «„— «·«‰ «Ã »œÊ‰ «œŒ«· «·⁄„Ì·")
        Else
             MsgBox ("Can not create Production order without inserting client")
        
        End If
        DBCboClientName.SetFocus
        Exit Sub
    End If
                    
    If DCboStore2Name.BoundText = "" Then
        If SystemOptions.UserInterface = ArabicInterface Then
            MsgBox ("·« Ì„ﬂ‰ «‰‘«¡ «„— «·«‰ «Ã »œÊ‰ «œŒ«· «·„Œ“‰")
           
       Else
            MsgBox ("Can not create Production order without inserting buffer")
       End If
        DCboStore2Name.SetFocus
        Exit Sub
    End If
                    
                    
    Dim Transaction_ID As Long
Dim Transaction_serial As String
    Dim RsTemp As New ADODB.Recordset
    Dim RsTest As New ADODB.Recordset
    Dim RowNum As Long
    Dim Transaction_Date As Date
    Transaction_Date = XPDtbBill
Transaction_Type = 26
    StrSQL = "Delete Transactions Where Transaction_ID In (Select TransactionID4 From TblDefComItemData Where IDDefCIT = " & val(TxtTransSerial) & ")"
    Cn.Execute StrSQL, , adExecuteNoRecords
    StrSQL = "Delete Transaction_Details Where Transaction_ID In (Select TransactionID4 From TblDefComItemData Where IDDefCIT = " & val(TxtTransSerial) & ")"
    Cn.Execute StrSQL, , adExecuteNoRecords
    
 
'        For RowNum = 1 To FG.Rows - 1
'
'             If FG.RowHidden(RowNum) Or CBool(FG.ValueMatrix(RowNum, FG.ColIndex("IsDeleted"))) = True Then GoTo NextRow333
'
'            If FG.TextMatrix(RowNum, FG.ColIndex("ItemID")) <> "" Then
'
'
'                 If SystemOptions.SysAllowStockNegative = False Then
'
'
'                    StrSQL = "Select * From TblItems where ItemID=" & val(FG.TextMatrix(RowNum, FG.ColIndex("ItemID")))
'                    Set RsTemp = New ADODB.Recordset
'                    RsTemp.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText
'
'                    If Not RsTemp.EOF Then
'
'                        If DCboStore2Name.BoundText <> "" Then
'                            Set RsTest = New ADODB.Recordset
'                            Set RsTest = GetItemQuantityStock(val(FG.TextMatrix(RowNum, FG.ColIndex("ItemID"))), val(Me.DCboStore2Name.BoundText), , , True, , Trim(FG.TextMatrix(RowNum, FG.ColIndex("Serial"))))
'
'                            If RsTest.EOF Or RsTemp.BOF Then
'                                Msg = "«·ﬁÿ⁄… –«  «·”Ì—Ì«· : "
'                                Msg = Msg + " «·’‰› : " & Trim(FG.Cell(flexcpTextDisplay, RowNum, FG.ColIndex("ItemName"))) & CHR(13) & "Ê«·„ÊÃÊœ ›Ï «·”ÿ— —ﬁ„  " & RowNum
'                                Msg = Msg + " €Ì— „ÊÃÊœ… ›Ì «·„Œ“‰ «·„Õœœ" & CHR(13)
'                                Msg = Msg + "Ê»«· «·Ï ·„ Ì „ «‰‘«¡ «–‰ «·’—›"
'
'                                MsgBox Msg
'                                Exit Sub
'                            End If
'                        End If
'                    End If
'                End If
'            End If
'NextRow333:
'        Next
    
Dim NoteSerial1 As String

            Dim Current_case As Integer, s As String, mBoxID As Long
        Dim rsDummy As New ADODB.Recordset
        Dim rsDummy2 As New ADODB.Recordset
        s = "Select GroupID From TblDefComItemData"
        s = s & " Where (IDDefCIT =" & val(TxtTransSerial.Text) & ") "
        s = s & " And  ItemId In (Select ItemId2 From TblDefComItemDet Det Where IsNull(Det.IsDeleted,0) <> 1 and Det.ItemID <> Det.ItemId2 "
        s = s & " and Det.IDDefCIT =" & val(TxtTransSerial.Text) & ") "
 
        s = s & " GROUP BY GroupID"
        rsDummy.Open s, Cn, adOpenStatic, adLockReadOnly
        Do While Not rsDummy.EOF
        
'                       TXTTransactionID4 = rsDummy!TransactionID4 & ""
'                TxtNoteSerial14 = rsDummy!NoteSerial14 & ""
'                StrSqlDel = "delete From Transactions where Transaction_ID=" & val(Me.TXTTransactionID4.Text) 'Val(rs("Transaction_ID").value)
'                Cn.Execute StrSqlDel, , adExecuteNoRecords
'
'
'                StrSqlDel = "delete From Transaction_Details where Transaction_ID=" & val(Me.TXTTransactionID4.Text) 'Val(rs("Transaction_ID").value)
'                Cn.Execute StrSqlDel, , adExecuteNoRecords
'
'                StrSqlDel = "delete From Notes where NoteSerial1=" & val(Me.TxtNoteSerial14.Text)  'Val(rs("Transaction_ID").value)
'                Cn.Execute StrSqlDel, , adExecuteNoRecords
'                StrSQL = "Delete From DOUBLE_ENTREY_VOUCHERS Where Transaction_ID=" & val(Me.TXTTransactionID4.Text)
'                Cn.Execute StrSQL, , adExecuteNoRecords
'
'                Cn.Execute "Delete from TransactionValueAdded where Transaction_ID=" & val(Me.TXTTransactionID4.Text) & ""
                TXTTransactionID4 = 0
                TxtNoteSerial14 = 0
              '  StrSQL = "delete From Notes where noteid=" & val(TxtNoteSerial14.Text)
               ' Cn.Execute StrSQL, , adExecuteNoRecords
        
        '            CurrentVoucherNo = GetVoucherGLNO(val(Text1.Text), CurrentVoucherSerialNo)
        '             If Trim(CurrentVoucherNo) <> "" And DateChanged <> True Then
        '            TxtNoteSerialV = CurrentVoucherNo '—ﬁ„ «·ﬁÌœ
        '            TxtNoteSerial1V = Trim(CurrentVoucherSerialNo)
        '             End If
        '
        '
        '
        '        DeleteTransactiomsVoucher val(Text1.Text)
                       Dim BranchID  As Double, StoreId As Double, StoreId2 As Double
                      
                BranchID = val(dcBranch.BoundText)
                  
                TxtNoteSerial14.Text = Voucher_coding(val(BranchID), XPDtbBill.value, 49, 0, , 26, , val(DCboStore2Name.BoundText))
                      
        
                StoreId = val(DCboStore2Name.BoundText)
                StoreId2 = val(DCboStore3Name.BoundText)
               ' If SystemOptions.IsMultiItemsInCompItem Then
       
                    
        '            If DcboEmp.BoundText = "" Then
        '                MsgBox ("·« Ì„ﬂ‰ «‰‘«¡ «·›« Ê—… »œÊ‰ «œŒ«· «·„‰œÊ»")
        '                DcboEmp.SetFocus
        '                Exit Sub
        '            End If
                            
                            
                           CostTOTAL = 0
'Check
StoreId = val(DCboStore2Name.BoundText)
  
    If DCboStore2Name.Text = "" Then
        If SystemOptions.UserInterface = ArabicInterface Then
            Msg = "ÌÃ»  ÕœÌœ «·„Œ“‰"
        Else
            Msg = "Select Inventory First"
        End If

        MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
      If DCboStore2Name.Enabled = True Then
        DCboStore2Name.SetFocus
      SendKeys "{F4}"
        End If
       cmd(2).Enabled = True
        Screen.MousePointer = vbDefault
      '  Cmd(2).Enabled = True
        Exit Sub
    End If
    



 

 NoteSerial1 = TxtNoteSerial14
 

Transaction_ID = CStr(new_id("Transactions", "Transaction_ID", "", True))
 
  
   
    

       

NoteSerial = Notes_coding(val(BranchID), Transaction_Date)
     If NoteSerial = "" Then
        If NoteSerial = "error" Then
            MsgBox " ·« Ì„ﬂ‰ «÷«›… ﬁÌÊœ ·Â–… «·⁄„·Ì… ·«‰ﬂ  ⁄œÌ  «·Õœ «·«ﬁ’Ì ··ﬁÌÊœ ﬂ„« Õœœ  ›Ì  —ﬁÌ„ «·”‰œ«  ": Exit Sub
        ElseIf NoteSerial = "" Then
                MsgBox " ·«Ì„ﬂ‰ «‰‘«¡ «·ﬁÌœ ·Â–« «·„” ‰œ ·«‰ﬂ Õœœ   —ﬁÌ„ ﬁÌÊœ ÌœÊÌ  ": Exit Sub
             
        End If
    End If
              

    StoreAccount = get_store_Account(CInt(StoreId), "Account_Code")

 
        TXTTransactionID4.Text = Transaction_ID
        TxtNoteSerial14.Text = NoteSerial1
        Transaction_serial = NoteSerial1
     Dim rsOut As New ADODB.Recordset
             
            Set rsOut = New ADODB.Recordset
            s = "Select BoxID From TblBoxesData Where Empid = " & val(Me.DcboEmp.BoundText)



            rsOut.Open s, Cn, adOpenStatic, adLockReadOnly
            If Not rsOut.EOF Then
                BoxID = val(rsOut!BoxID & "")
            End If
            mBoxID = val(DcboBox.BoundText)
 sql = "INSERT INTO  Transactions (  "
sql = sql & " Transaction_ID ,"
sql = sql & " BranchID ,"
sql = sql & " NoteSerial ,"
sql = sql & " NoteSerial1 ,"
sql = sql & " boxId ,"
sql = sql & " Transaction_serial ,"
sql = sql & " Transaction_Date ,"
sql = sql & " Transaction_Type ,"
sql = sql & " BillBasedOn ,"
sql = sql & " UserID ,"
sql = sql & " Trans_DiscountType ,"
sql = sql & " CusID ,"
sql = sql & " StoreId ,"
sql = sql & " StoreId1 ,"

sql = sql & " PaymentType ,"
sql = sql & " Emp_id ,"
sql = sql & " Transaction_NetValue ,"
sql = sql & " Vat, netvalue, PayedValue, "
sql = sql & " Currency_rate, Currency_id,sumVatLine,DueDate,"
 sql = sql & " TransactionComment,MIxCode,MixID,CBoBasedON,OrderType,order_no )"
 
    
 sql = sql & " VALUES("
sql = sql & " " & Transaction_ID & " ,"
sql = sql & " " & BranchID & " ,"
sql = sql & "'" & NoteSerial & "' ,"
sql = sql & "'" & NoteSerial1 & "' ,"
sql = sql & " " & val(BoxID) & " ,"
sql = sql & "'" & Transaction_serial & "',"
sql = sql & " " & SQLDate(Transaction_Date, True) & " ,"
sql = sql & " " & 26 & " ,"
sql = sql & " 3 ,"
sql = sql & " " & user_id & " ,"
sql = sql & " 0 ,"
sql = sql & " " & val(DBCboClientName.BoundText) & " ,"
sql = sql & " " & StoreId2 & " ,"
sql = sql & " " & StoreId & " ,"
sql = sql & " " & CboPayMentType.ListIndex & " ,"
sql = sql & " " & val(Emp_id) & " ,"
sql = sql & " " & val(txtTotalWithVat2) & " ,"
sql = sql & " " & val(TxtVAt22) & " ,"
sql = sql & " " & val(txtNet2) & " ,"
sql = sql & " " & val(txtNet2) & " ,"
sql = sql & " " & 1 & " ,"
sql = sql & " " & 1 & " ,0,"
sql = sql & " " & SQLDate(Transaction_Date, True) & " ,"
sql = sql & "'" & TransactionComment & "',"
sql = sql & "" & val(TxtMaxNo) & "," & val(TxtMaxNo) & ",3,0," & val(TxtTransSerial) & ")"

 
Cn.Execute sql
        
            s = "Select * From TblDefComItemData"
            s = s & " Where (IDDefCIT =" & val(TxtTransSerial.Text) & ") "
            s = s & " And  IsNull(GroupID,0) =  " & val(rsDummy!GroupID & "")
            s = s & " And  ItemId In (Select ItemId2 From TblDefComItemDet Det Where IsNull(Det.IsDeleted,0) <> 1 and Det.ItemID <> Det.ItemId2 "
            s = s & " and Det.IDDefCIT =" & val(TxtTransSerial.Text) & ") "
            
            Set rsDummy2 = New ADODB.Recordset
            rsDummy2.Open s, Cn, adOpenKeyset, adLockOptimistic
            Do While Not rsDummy2.EOF

  
                    CreateProduction BranchID, 0, XPDtbBill.value, 26, 0, val(user_id), 0, DBCboClientName.BoundText, StoreId, CboPayMentType.ListIndex, val(DcboEmp.BoundText), "«„— «‰ «Ã", val(rsDummy2!ID & ""), val(TXTTransactionID4)
               ' End If
                 
                StrSQL = "UPDATE TblDefComItemData SET  TransactionID4=" & val(TXTTransactionID4) & ",  NoteSerial14='" & TxtNoteSerial14 & "' WHERE ID  =" & val(TxtTransSerial)
                Cn.Execute StrSQL
                rsDummy2!TransactionID4 = val(TXTTransactionID4)
                rsDummy2!NoteSerial14 = Trim(TxtNoteSerial14)
                rsDummy2.update
                rsDummy2.MoveNext
            Loop
            rsDummy.MoveNext
            
        Loop
            
        
         StrSQL = "SELECT * FROM TblDefComItem "
StrSQL = StrSQL & "  WHERE      BranchId in(" & Current_branchSql & ")"
    StrSQL = StrSQL + " Order By ID"
    Set rs = New ADODB.Recordset
    rs.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText
    



Retrive val(TxtTransSerial)
End Sub

Private Sub cmdfrmRec_Click()
FrmCashing.show
FrmCashing.Cmd_Click 0
FrmCashing.Option2 = True
FrmCashing.DBCboClientName.BoundText = DBCboClientName.BoundText
FrmCashing.XPTxtVal = txtNet2
FrmCashing.txtTradingContractID = TxtTransSerial.Text
FrmCashing.lbl(95).Caption = "—ﬁ„ «·ÿ·»ÌÂ"
'FrmCashing.TxtVAt2 = TxtVAt22
'FrmCashing.txtTotal = txtTotalWithVat2
FrmCashing.DcboBox.BoundText = DcboBox.BoundText
'FrmCashing.Option2 = True

End Sub

Private Sub CMDSHOWISSUE2_Click()
 FrmOut.XPBtnMove_Click (2)
    
 FrmOut.Retrive val(TXTTransactionID5.Text)
End Sub

Private Sub cmdRecalc_Click()
    
FixRowsLine
Exit Sub
    Dim i As Long
    Dim Rs3 As New ADODB.Recordset
    Dim mItemNo As Long
    Dim mItemNo2 As Long
    Dim mLineID As Long
    Dim mTableID As Long
    Dim j As Long
    Dim MySQL As String
    Dim s As String
    With FG
          
        For i = 1 To .Rows - 1
          
            If i = 28 Then
                i = i
            End If
            mItemNo = val(.TextMatrix(i, .ColIndex("ItemID")))
            mItemNo2 = val(.TextMatrix(i, .ColIndex("ItemID2")))
            mTableID = val(.TextMatrix(i, .ColIndex("TableID")))
            mLineID = val(.TextMatrix(i, .ColIndex("LineID")))
            If mItemNo2 = 656 Then
                mItemNo2 = 656
            End If
            If i = 20 Then
            i = 20
            'i = 40
            End If
            If Trim(TxtMaxNo2) <> "" Then GoTo FromMix
Default:
            MySQL = " SELECT          TT.ItemID MainItemID , TT.ItemID, ForUnit, MethodCalc, TblItemsParts.lowering, TblItemsParts.increase, dbo.TblItemsParts.UnitID, dbo.TblItemsParts.isReplaced, dbo.TblItemsParts.PartItemPrice, dbo.TblItemsParts.PartItemQty, dbo.TblItemsParts.PartItemID, "
           
            MySQL = MySQL + "      dbo.TblItemsParts.ItemID, dbo.TblItemsParts.TableID, dbo.TblUnites.UnitName, dbo.TblUnites.UnitNamee, dbo.TblItems.ItemCode, dbo.TblItems.ItemName,"
            MySQL = MySQL + "      dbo.TblItems.ItemNamee , dbo.TblItems.fullcode"
            MySQL = MySQL + "  FROM         dbo.TblItemsParts INNER JOIN"
            MySQL = MySQL + "      dbo.TblUnites ON dbo.TblItemsParts.Unitid = dbo.TblUnites.UnitID RIGHT OUTER JOIN"
            MySQL = MySQL + "      dbo.TblItems ON dbo.TblItemsParts.PartItemID = dbo.TblItems.ItemID"
            MySQL = MySQL + "                 RIGHT OUTER JOIN dbo.TblItems TT"
            MySQL = MySQL + "                  ON  dbo.TblItemsParts.ItemID = TT.ItemID"
            MySQL = MySQL + " Where dbo.TblItemsParts.ItemID = " & mItemNo2 & " And TblItemsParts.PartItemID = " & mItemNo & ""
            If mTableID <> 0 Then
                MySQL = MySQL + " And TblItemsParts.TableID = " & mTableID
            End If
            MySQL = MySQL + " ORDER BY dbo.TblItemsParts.TableID"
            
            
                    
            Set Rs3 = New ADODB.Recordset
            Rs3.Open MySQL, Cn, adOpenStatic, adLockReadOnly
            If Not Rs3.EOF Then

                .TextMatrix(i, .ColIndex("FlgX")) = IIf(IsNull(Rs3("PartItemQty").value), 0, Rs3("PartItemQty").value)
              ' .TextMatrix(i, .ColIndex("Qty")) = 1
                .TextMatrix(i, .ColIndex("ForUnit")) = IIf(IsNull(Rs3("ForUnit").value), 0, Rs3("ForUnit").value)
                .TextMatrix(i, .ColIndex("MethodCalc")) = IIf(IsNull(Rs3("MethodCalc").value), 0, Rs3("MethodCalc").value)
                .TextMatrix(i, .ColIndex("lowering")) = IIf(IsNull(Rs3("lowering").value), 0, Rs3("lowering").value)
                .TextMatrix(i, .ColIndex("Increase")) = IIf(IsNull(Rs3("increase").value), 0, Rs3("increase").value)
                .TextMatrix(i, .ColIndex("PartItemQty")) = IIf(IsNull(Rs3("PartItemQty").value), 0, Rs3("PartItemQty").value)
            Else
            
                If Trim(TxtMaxNo2) <> "" Then
FromMix:
                    'Exit Sub
                    'If fg2.Rows > 1 And FG.Rows > 1 Then Exit Sub
                        s = " SELECT TblItems.ItemName,TblItems.FullCode itemcode, tu.UnitName,TblDefComItemData.Qty,TblDefComItemData.Qty FlgX,LineID = 1, ItemId2 =" & val(mItemNo2) & ",ItemName2 =N'" & Trim(DcboItemID1.Text) & "',"
                        s = s & " TblDefComItemData.cost,TblDefComItemData.Price,TblDefComItemData.Total,TblDefComItemData.UnitId,TblDefComItemData.ItemID"
                        
                        s = s & " FROM  TblDefComItemData INNER JOIN TblDefComItem ON TblDefComItem.ID = TblDefComItemData.IDDefCIT"
                        
                        s = s & " INNER JOIN TblItems ON TblItems.ItemID = TblDefComItemData.ItemID"
                        s = s & " INNER JOIN TblUnites AS tu"
                        
                        s = s & " ON  tu.UnitID= TblDefComItemData.UnitID"
                        
                        s = s & " Where TblDefComItem.MaxNo = N'" & Trim(TxtMaxNo2) & "'"
                        s = s & " and  TblDefComItemData.ItemID = " & val(mItemNo)
                        
                       '  s = s & " and  TblDefComItemData.LineId = " & val(mLineID)
                        
                        
                         Set Rs3 = New ADODB.Recordset
                        
                         Rs3.Open s, Cn, adOpenStatic, adLockReadOnly
                         If Not Rs3.EOF Then

                              .TextMatrix(i, .ColIndex("FlgX")) = IIf(IsNull(Rs3("FlgX").value), 0, Rs3("FlgX").value)
                              If val(.TextMatrix(i, .ColIndex("Qty"))) = 0 Then
                                .TextMatrix(i, .ColIndex("Qty")) = IIf(IsNull(Rs3("FlgX").value), 0, Rs3("FlgX").value)
                              End If
                            ' .TextMatrix(i, .ColIndex("Qty")) = 1
                              '.TextMatrix(i, .ColIndex("ForUnit")) = IIf(IsNull(Rs3("ForUnit").value), 0, Rs3("ForUnit").value)
                              '.TextMatrix(i, .ColIndex("MethodCalc")) = IIf(IsNull(Rs3("MethodCalc").value), 0, Rs3("MethodCalc").value)
                              '.TextMatrix(i, .ColIndex("lowering")) = IIf(IsNull(Rs3("lowering").value), 0, Rs3("lowering").value)
                              '.TextMatrix(i, .ColIndex("Increase")) = IIf(IsNull(Rs3("increase").value), 0, Rs3("increase").value)
                              .TextMatrix(i, .ColIndex("PartItemQty")) = IIf(IsNull(Rs3("Qty").value), 0, Rs3("Qty").value)
                        End If
                        If val(.TextMatrix(i, .ColIndex("Qty"))) = 0 Then
                            .TextMatrix(i, .ColIndex("Qty")) = 1
                            .TextMatrix(i, .ColIndex("FlgX")) = 1
                           ' GoTo Default
                        End If
                       '------------
                        
                        
                        
                End If
            End If
NextRow:
        Next
    End With
   
    For i = 1 To FG2.Rows - 1
        If SystemOptions.IsMultiItemsInCompItem Then
            ReLineGrid i, True, , True
            
        Else
       ' ReLineGrid i,  True
    End If
Next
End Sub

Private Sub DcboBuiltinItemID_Click(Area As Integer)
    Me.TxtAttachedItemCode3.Text = GetItemCode(val(Me.DcboBuiltinItemID.BoundText))
End Sub

Private Sub DcboItemID3_KeyDown(KeyCode As Integer, Shift As Integer)
 If KeyCode = vbKeyF3 Then
        FrmItemSearch.RetrunType = 2026
        FrmItemSearch.show vbModal
        
    End If
End Sub

Private Sub DcboItemID5_Change()
On Error Resume Next
If val(DcboItemID5.BoundText) = 0 Then Exit Sub
'If val(txtQty1) = 0 Then txtQty1 = 1
If val(DcboItemID5.BoundText) = val(val(DcboItemID1.BoundText)) Then DcboItemID5.Text = "": Exit Sub
Dim UnitID As Long
Dim UnitName As String
    Dim Dcombos As ClsDataCombos
 Set Dcombos = New ClsDataCombos
 
    'Me.TxtAttachedItemCode2.Text = GetItemCode(val(Me.DcboItemID5.BoundText))
    Dcombos.GetItemsUnitsDetai DcbUnit2, val(DcboItemID5.BoundText)
    GetDefaultItemUnit val(Me.DcboItemID5.BoundText), UnitID, UnitName
    DcbUnit5.Text = UnitName
    DcbUnit5.BoundText = UnitID
    
    'Me.TxtAttachedItemCode2.Text = GetItemCode(val(Me.DcboItemID2.BoundText))
Dcombos.GetItemsUnitsDetai DcbUnit5, val(DcboItemID5.BoundText)
If Me.TxtModFlg.Text <> "R" Then
   GetDefaultItemUnit val(Me.DcboItemID5.BoundText), UnitID, UnitName
    DcbUnit5.Text = UnitName
    DcbUnit5.BoundText = UnitID
    
    Dim l As Long

    
   ' FillGrid2
  End If
Dim widthPrice  As Double


End Sub

Private Sub FG2_CellButtonClick(ByVal Row As Long, ByVal Col As Long)
Select Case FG2.ColKey(Col)
Case "NoteSerial14"
            TXTTransactionID4 = val(FG2.TextMatrix(Row, FG2.ColIndex("TransactionID4")))
            FrmProductionOrder.show
         FrmProductionOrder.XPBtnMove_Click (2)
        FrmProductionOrder.Retrive val(TXTTransactionID4.Text)
 End Select
End Sub

Private Sub FG3_CellButtonClick(ByVal Row As Long, ByVal Col As Long)
    Select Case FG3.ColKey(Col)
    
    Case "NoteSerial1"
        Dim mTrans As Long
        mTrans = val(FG3.TextMatrix(Row, FG3.ColIndex("Transaction_ID")))
        
        On Error Resume Next
        FrmOutProductionOrder.Retrive mTrans
    Case "NoteSerial"
        ShowGL_cc val(FG3.TextMatrix(Row, FG3.ColIndex("NoteSerial"))), , 200, val(FG3.TextMatrix(Row, FG3.ColIndex("NoteID")))
    End Select
    
 
End Sub

Private Sub ISButton2_Click()
    On Error Resume Next
          If DoPremis(Do_Attach, Me.Name, True) = False Then
                Exit Sub
            End If
            
ShowAttachments TxtNoteSerial1, "15012018001"


End Sub

Private Sub Txt_order_no_Change()
    Dim StrSQL As String
    Dim rs2 As ADODB.Recordset
    Dim Transaction_Type As Integer
        Transaction_Type = 42
    If TXT_order_no.Text = "" Then Exit Sub


   'Transaction_ID = get_transactionData("order_no", Txt_order_no.text, "Transaction_ID", Transaction_Type)
    If Me.TxtModFlg <> "R" And Me.TxtModFlg <> "" Then
        RetriveOrder Me.TXT_order_no, Transaction_Type
    End If

End Sub

Private Sub txtItemCodeBuiltin_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        If TxtAttachedItemCode3.Text = "" Then
            Me.DcboBuiltinItemID.BoundText = ""
        Else
            Me.DcboBuiltinItemID.BoundText = GetItemID(Trim$(Me.TxtAttachedItemCode3.Text))
        End If
    End If

End Sub

Private Sub txtLength_Change()
If Not SystemOptions.IsMultiItemsInCompItem Then
    If FG2.Rows > 1 Then
        FG2.TextMatrix(1, FG2.ColIndex("Qty")) = txtQty1
        FG2.TextMatrix(1, FG2.ColIndex("Price")) = TxtPrice
        FG2.TextMatrix(1, FG2.ColIndex("widtj")) = txtwidtj
        FG2.TextMatrix(1, FG2.ColIndex("hight")) = txthight
        FG2.TextMatrix(1, FG2.ColIndex("Length")) = txtLength
        
    End If
End If
callarea
End Sub

Private Sub TxtMaxNo2_Change()
    If Me.TxtModFlg = "N" Or Me.TxtModFlg = "E" Then

                
    End If
        
End Sub

Private Sub TxtPhone_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyReturn Then
    GetCustomerNamebyPhone (TxtPhone.Text)
End If
End Sub

Private Sub CboPayMentType_Change()
Exit Sub
    If CboPayMentType.ListIndex = 0 Then '‰ﬁœÌ
        DcboBox.Enabled = True
    Else
        DcboBox.BoundText = ""
        DcboBox.Enabled = False
    End If
End Sub

Private Sub CboPayMentType_Click()



    CboPayMentType_Change
 
End Sub

Private Sub chkIsAdd_Click()
    If chkIsAdd Then
        XPCboGroup2.Enabled = True
        XPCboGroup5.Enabled = True
    Else
        XPCboGroup2.Enabled = False
        XPCboGroup2.BoundText = ""
        
        Dcombos.GetItemsNames Me.DcboItemID2
        
        XPCboGroup5.Enabled = False
        XPCboGroup5.BoundText = ""
        
        Dcombos.GetItemsNames Me.DcboItemID5
    End If
End Sub



Private Sub cmbSpecification_Change()
    Dim Dcombos As New ClsDataCombos
    Dim mIndex As Integer
    If Trim(cmbSpecification.BoundText) <> "" Then
        mIndex = myRound(cmbSpecification.BoundText)
        Dcombos.GetItemSGroupsupdate Me.XPCboGroup, , " SpecificationID =  " & val(cmbSpecification.BoundText)
        'Dcombos.GetItemsNamesupdate Me.DcboItemID2, , , , , mIndex
    Else
        Dcombos.GetItemSGroupsupdate Me.XPCboGroup
        'Dcombos.GetItemsNamesupdate Me.DcboItemID2
    End If

End Sub
Private Sub cmbSpecification_KeyUp(KeyCode As Integer, _
                             Shift As Integer)

    If KeyCode = vbKeyF5 Then
        Dim Dcombos As ClsDataCombos
        
        Set Dcombos = New ClsDataCombos
        Dcombos.GetItemSGroupsupdate XPCboGroup, , " SpecificationID =  " & val(cmbSpecification.BoundText)
        
    End If

End Sub
Private Sub Command1_Click()
            On Error Resume Next
ShowAttachments TxtAttachedItemCode, "0701201407"


End Sub

Private Sub cmdAddCustomer_Click()
    
If SystemOptions.DontShowMoreDetailsCompItem Then
    
    FrmCustemers.show
    FrmCustemers.Retrive val(DBCboClientName.BoundText), Me.Name
    FrmCustemers.FormNamee = Me.Name
    Dim Dcombos As New ClsDataCombos
   ' Dcombos.GetCustomersSuppliers 1, Me.DBCboClientName, True
    If DBCboClientName.Text = "" Then
   '     DBCboClientName.BoundText = mCustId
    End If
    Exit Sub
End If
           
Dim CUSTID As Double
Dim mCode As String

If SystemOptions.UserInterface = ArabicInterface Then
    If Trim(txtCustomerName) = "" Then MsgBox "«œŒ· «”„ «·⁄„Ì·": Exit Sub
    If Trim(TxtPhone) = "" Then MsgBox "«œŒ· —ﬁ„ «·Â« ›/«·ÃÊ«·  ": Exit Sub
Else
    If Trim(txtCustomerName) = "" Then MsgBox "Enter the customer name": Exit Sub
    If Trim(TxtPhone) = "" Then MsgBox "Enter your phone / mobile number  ": Exit Sub

End If

Dim s As String
Dim rsDummy As New ADODB.Recordset

s = "Select * from TblCustemers WHere 1=1 "
If Trim(TxtPhone) <> "" Then
    s = s & " And Cus_mobile = N'" & Trim(TxtPhone) & "' "
End If
If Trim(txtCustomerName) <> "" Then
    'If Trim(TxtPhone) <> "" Then
    '    s = s & " Or CusName = '" & Trim(txtCustomerName.Text) & "'"
    'Else
    '    s = s & " and CusName = '" & Trim(txtCustomerName.Text) & "'"
    'End If
End If
rsDummy.Open s, Cn, adOpenStatic
If Not rsDummy.EOF Then
    TxtSearchCode.Text = rsDummy!Fullcode & ""
    TxtSearchCode2.Text = rsDummy!Fullcode & ""
    DBCboClientName.BoundText = val(rsDummy!CusID & "")
   
    txtCustomerName.backcolor = vbGreen
    TxtPhone.backcolor = vbGreen
    Exit Sub
Else
    txtCustomerName.backcolor = vbWhite
    TxtPhone.backcolor = vbWhite
End If

    createCustomer txtCustomerName.Text, txtCustomerName.Text, val(dcBranch.BoundText), CUSTID, TxtPhone.Text, mCode
    TxtSearchCode.Text = mCode
    TxtSearchCode2.Text = mCode
    Dcombos.GetCustomersSuppliers 1, Me.DBCboClientName, True
    DBCboClientName.BoundText = CUSTID
    If SystemOptions.UserInterface = ArabicInterface Then
        MsgBox " „ «÷«›… «·⁄„Ì·"
    Else
        MsgBox "Customer added"
    End If
    'txtCustomerName = ""

End Sub

Private Sub DcboBox_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyF5 Then
ReloadCombos
End If
End Sub

Private Sub DcboItemID1_LostFocus()
DcboItemID1_Validate False

End Sub

Private Sub DcboItemID1_Validate(Cancel As Boolean)
On Error Resume Next
'If Not IsNumeric(DcboItemID1.BoundText) Then cboItemID1.Text = "": DcboItemID1.BoundText = 0
If Not IsNumeric(DcboItemID1.BoundText) Then DcboItemID1.Text = "": DcboItemID1.BoundText = 0
If val(DcboItemID1.Tag) = val(DcboItemID1.BoundText) And val(DcboItemID1.BoundText) <> 0 Then Exit Sub
If val(DcboItemID1.BoundText) = 0 Then Exit Sub
If val(txtQty1) = 0 Then txtQty1 = 1
Dim UnitID As Long
Dim UnitName As String

    Me.TxtAttachedItemCode.Text = GetItemCode(val(Me.DcboItemID1.BoundText))
Dcombos.GetItemsUnitsDetai DcbUnit, val(DcboItemID1.BoundText)
If Me.TxtModFlg.Text <> "R" Or IsSaveWithOutMsg Then
  GetDefaultItemUnit val(Me.DcboItemID1.BoundText), UnitID, UnitName
    DcbUnit.Text = UnitName
    DcbUnit.BoundText = UnitID
    
    If Not SystemOptions.IsMultiItemsInCompItem Then
        FG2.Rows = 1
        FG.Rows = 1
        If Not FillGrid Then Exit Sub
        FillGridItemType val(DcboItemID1.BoundText), DcboItemID1.Text, Trim$(TxtAttachedItemCode.Text), 1, val(DcbUnit.BoundText), DcbUnit.Text, val(txtQty1), val(TxtPrice), val(XPCboGroup.BoundText), XPCboGroup.Text
        
    Else
       
        Dim widthPrice  As Double
        If Not IsSaveWithOutMsg Then
            TxtPrice = GetItemPriceByWitdth(val(DcboItemID1.BoundText), val(txtwidtj), val(txthight))
            
            If val(TxtPrice) = 0 Then
                TxtPrice = GetItemPrice(val(Me.DcboItemID1.BoundText), , val(UnitID))
                
            End If
            TxtPrice = val(TxtPrice) + GetItemAddPrice(val(DcboItemID1.BoundText))
        End If
        txtTotal = val(TxtPrice) * val(txtQty1)
        
    End If
     
    'fillgrid
    
    
  End If

DcboItemID1.Tag = DcboItemID1.BoundText
End Sub

Private Sub DcboItemID3_Change()
On Error Resume Next
If val(DcboItemID3.BoundText) = 0 Then Exit Sub
'If val(txtQty1) = 0 Then txtQty1 = 1
Dim UnitID As Long
Dim UnitName As String
    Dim Dcombos As ClsDataCombos
 Set Dcombos = New ClsDataCombos
 
    Me.txtItemCode.Text = GetItemCode(val(Me.DcboItemID3.BoundText))
    Dcombos.GetItemsUnitsDetai DcbUnit3, val(DcboItemID3.BoundText)
    GetDefaultItemUnit val(Me.DcboItemID2.BoundText), UnitID, UnitName
    DcbUnit2.Text = UnitName
    DcbUnit2.BoundText = UnitID
    
    
    GetDefaultItemUnit val(Me.DcboItemID5.BoundText), UnitID, UnitName
    DcbUnit5.Text = UnitName
    DcbUnit5.BoundText = UnitID
    
    Me.txtItemCode.Text = GetItemCode(val(Me.DcboItemID3.BoundText))
Dcombos.GetItemsUnitsDetai DcbUnit2, val(DcboItemID2.BoundText)
Dcombos.GetItemsUnitsDetai DcbUnit5, val(DcboItemID5.BoundText)
If Me.TxtModFlg.Text <> "R" Then
   GetDefaultItemUnit val(Me.DcboItemID3.BoundText), UnitID, UnitName
    DcbUnit3.Text = UnitName
    DcbUnit3.BoundText = UnitID
    
    Dim l As Long

    
   ' FillGrid2
  End If
Dim widthPrice  As Double

'txtPrice = GetItemPriceByWitdth(val(DcboIte
End Sub

Private Sub DCboStore3Name_Change()
    If Me.TxtModFlg.Text = "N" Or Me.TxtModFlg.Text = "E" Then
        DCboStoreName.BoundText = DCboStore3Name.BoundText
    End If
End Sub

Private Sub DCboStore3Name_Click(Area As Integer)
    DCboStore3Name_Change
End Sub

Private Sub FG_CellButtonClick(ByVal Row As Long, ByVal Col As Long)
 Dim mItemFullCode As String
 Select Case FG.ColKey(Col)
    
    Case "ShowAttatch"
        mItemFullCode = Trim(FG.TextMatrix(Row, FG.ColIndex("FullCode")))
            
            On Error Resume Next
        ShowAttachments mItemFullCode, "0701201407"
    End Select
End Sub

Private Sub fg_Click()
    mIdDisplay = val(FG.TextMatrix(FG.Row, FG.ColIndex("LineID")))
    If mIdDisplay = 0 Then mIdDisplay = mNewId
End Sub
Private Sub FixRowsLine(Optional ByVal IsReste As Boolean = False)
    Dim i As Long
    Dim j As Long
    Dim mLineNo As Long
    Dim mParentLineNo As Long
    Dim mItemId2 As Long
    Dim mItemId As Long
    Dim LngRow2 As Long
Dim m As Long
  '   Dim LngRow  As Long
   Dim mLineCor As Long
    
'    For i = 1 To Fg.Rows - 1
'        Fg.TextMatrix(i, Fg.ColIndex("LineID")) = ""
'    Next
' '  If val(TxtMaxNo2) = 0 Then Exit Sub
'
'         For j = 1 To fg2.Rows - 1
'
'            mItemId = val(fg2.TextMatrix(j, fg2.ColIndex("ItemId")))
'            mParentLineNo = val(fg2.TextMatrix(j, fg2.ColIndex("LineID")))
'            LngRow2 = fg2.FindRow(mParentLineNo, fg2.FixedRows, fg2.ColIndex("LineID"), False, True)
'            'If j > LngRow2 Then
'                For m = 1 To Fg.Rows - 1
'                    mLineNo = val(Fg.TextMatrix(m, Fg.ColIndex("LineID")))
'                    mItemId2 = val(Fg.TextMatrix(m, Fg.ColIndex("ItemID2")))
'                    If mLineNo = 0 Then
'                        If mItemId2 = mItemId Then
'                            Fg.TextMatrix(m, Fg.ColIndex("LineID")) = mParentLineNo
'                        End If
'
'                    End If
''                    mLineNo = val(Fg.TextMatrix(m, Fg.ColIndex("LineID")))
''                    If val(Fg.TextMatrix(m, Fg.ColIndex("ItemId2"))) = mItemId And mParentLineNo = mLineNo Then
''                        Fg.TextMatrix(m, Fg.ColIndex("LineID")) = mParentLineNo + 1
''                    End If
'                Next
'              '  mParentLineNo = mParentLineNo + 1
'              '  fg2.TextMatrix(j, fg2.ColIndex("LineID")) = mParentLineNo
'          '  End If
'        Next
'
'
     
'     For j = 1 To fg2.Rows - 1
'
'            mItemId = val(fg2.TextMatrix(j, fg2.ColIndex("ItemId")))
'            mParentLineNo = val(fg2.TextMatrix(j, fg2.ColIndex("LineID")))
'            LngRow2 = fg2.FindRow(mParentLineNo, fg2.FixedRows, fg2.ColIndex("LineID"), False, True)
'            If j > LngRow2 Then
'                For m = 1 To FG.Rows - 1
'                    mLineNo = val(FG.TextMatrix(m, FG.ColIndex("LineID")))
'                    If val(FG.TextMatrix(m, FG.ColIndex("ItemId2"))) = mItemId And mParentLineNo = mLineNo Then
'                 '       FG.TextMatrix(m, FG.ColIndex("LineID")) = mParentLineNo + 1
'                    End If
'                Next
'                mParentLineNo = mParentLineNo + 1
'              ' FG2.TextMatrix(j, FG2.ColIndex("LineID")) = mParentLineNo
'            End If
'        Next
'
'  '   Dim LngRow  As Long
'
'    With FG
'        For i = 1 To FG.Rows - 1
'            mItemId2 = val(.TextMatrix(i, .ColIndex("ItemId2")))
'            mLineNo = val(.TextMatrix(i, .ColIndex("LineID")))
'            For j = 1 To fg2.Rows - 1
'
'                mItemId = val(fg2.TextMatrix(j, fg2.ColIndex("ItemId")))
'                mParentLineNo = val(fg2.TextMatrix(j, fg2.ColIndex("LineID")))
'
'
'                If mItemId = mItemId2 Then
'                    LngRow = fg2.FindRow(mLineNo, fg2.FixedRows, fg2.ColIndex("LineID"), False, True)
'                    If j = 6 Then
'                    j = j
'                    End If
'                    If LngRow = -1 Then
'                        LngRow = LngRow
'                      '  FG2.TextMatrix(j, FG2.ColIndex("LineID")) = j
'                       ' .TextMatrix(i, .ColIndex("LineID")) = j
'                    ElseIf mItemId = mItemId2 And mParentLineNo <> LngRow Then
'                       '  FG2.TextMatrix(j, FG2.ColIndex("LineID")) = j
'                        '.TextMatrix(i, .ColIndex("LineID")) = j
'                    End If
'
'                  '  FG2.FindRow "ItemId2 = " & mItemId2 & " and LineID = " & mItemId
'
'                End If
'
'            Next
'        Next
'    End With
'If Not IsSaveWithOutMsg Then Exit Sub
Dim s As String
Dim rsDummy As New ADODB.Recordset
Dim rsDummy2 As New ADODB.Recordset
Dim mrsDummy2 As Boolean
s = s & "  SELECT *"
s = s & " From TblDefComItemData"
s = s & " WHERE  LineID NOT IN (SELECT TblDefComItemDet.LineID"
s = s & "                       From TblDefComItemDet"
s = s & "                       Where TblDefComItemDet.IDDefCIT = TblDefComItemData.IDDefCIT"
s = s & "                              AND TblDefComItemDet.ItemID2 = TblDefComItemData.ItemID)"
s = s & "        AND IDDefCIT = " & val(TxtTransSerial)
rsDummy.Open s, Cn, adOpenStatic, adLockReadOnly
'
If rsDummy.EOF Then

    s = " Select LineID From TblDefComItemData"
    s = s & " WHERE  LineID NOT IN (SELECT TblDefComItemDet.LineID"
    s = s & "                       From TblDefComItemDet"
    s = s & "                       Where TblDefComItemDet.IDDefCIT = TblDefComItemData.IDDefCIT)"
    s = s & "                              "
    s = s & "        AND IDDefCIT = " & val(TxtTransSerial)
    Set rsDummy2 = New ADODB.Recordset
    rsDummy2.Open s, Cn, adOpenStatic, adLockReadOnly
    If rsDummy2.EOF Then
        mrsDummy2 = True
    End If
End If

Dim Rs3 As New ADODB.Recordset
If Not rsDummy.EOF Or mrsDummy2 Then
     Dim jj As Long
     Dim mRow As Long
     Dim mLineID As Long
     mRow = 0
     Dim mQty As Double
     Dim mCount As Double
     FG.ColHidden(FG.ColIndex("LineID")) = False
'     If FG.TextMatrix(1, FG.ColIndex("LineId")) = 1 Then Exit Sub
    If IsReste Or IsSaveWithOutMsg Then
        FG.Rows = 1
    
        For i = 1 To FG2.Rows - 1
            mItemId = val(FG2.TextMatrix(i, FG2.ColIndex("ItemId")))
            mLineID = val(FG2.TextMatrix(i, FG2.ColIndex("LineID")))
            mQty = val(FG2.TextMatrix(i, FG2.ColIndex("Qty")))
            DcboItemID1.Tag = ""
            DcboItemID1.BoundText = mItemId
            txtQty1 = mQty
            TxtPrice = val(FG2.TextMatrix(i, FG2.ColIndex("Price")))
            txthight = val(FG2.TextMatrix(i, FG2.ColIndex("hight")))
            txtwidtj = val(FG2.TextMatrix(i, FG2.ColIndex("widtj")))
            'txtlow = val(FG2.TextMatrix(i, FG2.ColIndex("lowering")))
            txtLength = val(FG2.TextMatrix(i, FG2.ColIndex("Length")))
            txtthickness = val(FG2.TextMatrix(i, FG2.ColIndex("thickness")))
            txtDO = val(FG2.TextMatrix(i, FG2.ColIndex("DO")))
            txtDI = val(FG2.TextMatrix(i, FG2.ColIndex("DI")))
            txtDiameter = val(FG2.TextMatrix(i, FG2.ColIndex("Diameter")))
            
            s = " SELECT        TblItemsParts.ItemID,dbo.TblItemsParts.PartItemID "
            s = s & " From TblItemsParts "
            s = s & " Where dbo.TblItemsParts.ItemID = " & mItemId
            Set Rs3 = New ADODB.Recordset
            Rs3.Open s, Cn, adOpenStatic, adLockOptimistic, adCmdText
            'If Rs3.EOF Then GoTo NextRow
'            Rs3.MoveLast
'            mCount = val(Rs3.RecordCount & "") + mRow
'            If i = 6 Then
'                i = i
'            End If
'            'Do While Not Rs3.EOF
'                If mRow = 0 Then mRow = 1
'                For jj = mRow To mCount - 1
'                    If val(FG.Rows) > (mCount - 1) Then
'                    If val(FG.TextMatrix(jj, FG.ColIndex("ItemId2"))) = mItemId Then
'                        FG.TextMatrix(jj, FG.ColIndex("LineId")) = mLineID
'                    End If
'                    Else
'                    Exit Sub
'                    End If
'                Next
'
'                If mRow = 1 Then
'                    mRow = mCount
'                Else
'                    mRow = mCount
'                End If
'            '    Rs3.MoveNext
'            'Loop
'
'            If SystemOptions.IsMultiItemsInCompItem Then
'                  If IsSaveWithOutMsg Then
'                    ReLineGrid i, True, , True
'                End If
'
'              End If
         '   cmdRecalc_Click
            DcboItemID1_Validate False
            AddNewFgRow val(DcboItemID2.BoundText), "ItemID2", "ItemName2", i
NextRow:
        Next
        
      
       

    End If
        FG.ColHidden(FG.ColIndex("LineID")) = False
   ' End If
End If
   Exit Sub
'    With Fg
'
'        For i = 1 To Fg.Rows - 1
'            mItemId2 = val(.TextMatrix(i, .ColIndex("ItemId2")))
'            mLineNo = val(.TextMatrix(i, .ColIndex("LineID")))
'            For j = 1 To fg2.Rows - 1
'
'                mItemID = val(fg2.TextMatrix(j, fg2.ColIndex("ItemId")))
'                mParentLineNo = val(fg2.TextMatrix(j, fg2.ColIndex("LineID")))
'
'                If mItemID = mItemId2 Then
'                    LngRow = fg2.FindRow(mLineNo, fg2.FixedRows, fg2.ColIndex("LineID"), False, True)
'                    If LngRow = -1 Then
'                        LngRow = LngRow
'                        fg2.TextMatrix(j, fg2.ColIndex("LineID")) = j
'                        mLineCor = 0
'                        .TextMatrix(i, .ColIndex("LineID")) = j
'                    ElseIf mItemID = mItemId2 And mParentLineNo <> LngRow Then
'                         fg2.TextMatrix(j, fg2.ColIndex("LineID")) = j
'                         mLineCor = j
'                         GoTo NextRow
'                        '.TextMatrix(i, .ColIndex("LineID")) = j
'                    End If
'
'                  '  FG2.FindRow "ItemId2 = " & mItemId2 & " and LineID = " & mItemId
'
'                End If
'
'            Next
'NextRow:
'           If mLineCor <> 0 Then
'            .TextMatrix(i, .ColIndex("LineID")) = mLineCor
'            End If
'           mLineCor = 0
'        Next
'    End With
'

    Exit Sub
        Dim LngRow  As Long
        With FG2
    
        For i = 1 To .Rows - 1
            mItemId2 = val(.TextMatrix(i, .ColIndex("ItemId")))
            mLineNo = val(.TextMatrix(i, .ColIndex("LineID")))
            For j = 1 To FG.Rows - 1
                
                mItemId = val(FG.TextMatrix(j, FG.ColIndex("ItemId2")))
                mParentLineNo = val(FG.TextMatrix(j, FG.ColIndex("LineID")))
                If val(FG.TextMatrix(j, FG.ColIndex("IsSer"))) = 0 Then
                    If mItemId = mItemId2 Then
                        LngRow = FG2.FindRow(mParentLineNo, FG2.FixedRows, FG2.ColIndex("LineID"), False, True)
                        If LngRow = -1 Then
                            LngRow = LngRow
                            FG2.TextMatrix(i, FG2.ColIndex("LineID")) = i
                            mLineCor = 0
                            FG.TextMatrix(j, FG.ColIndex("LineID")) = i
                            FG.TextMatrix(j, FG.ColIndex("IsSer")) = i
                        ElseIf mItemId = mItemId2 And mParentLineNo <> LngRow Then
                             FG.TextMatrix(j, FG.ColIndex("LineID")) = i
                             FG.TextMatrix(j, FG.ColIndex("IsSer")) = i
                             mLineCor = j
                             
                            ' GoTo NextRow2
                            '.TextMatrix(i, .ColIndex("LineID")) = j
                        End If
                        FG.TextMatrix(j, FG.ColIndex("IsSer")) = FG.TextMatrix(j, FG.ColIndex("LineID"))
                      '  FG2.FindRow "ItemId2 = " & mItemId2 & " and LineID = " & mItemId
                        
                    End If
                End If
                
            Next
NextRow2:
'           If mLineCor <> 0 Then
'            .TextMatrix(i, .ColIndex("LineID")) = mLineCor
'            End If
           mLineCor = 0
           
        Next
    End With
End Sub
Private Sub FG2_AfterEdit(ByVal Row As Long, ByVal Col As Long)
    
    mIdDisplay = val(FG2.TextMatrix(Row, FG2.ColIndex("LineID")))
    If mIdDisplay = 0 Then mIdDisplay = mNewId
    With FG2
        Select Case .ColKey(Col)
        Case "Qty", "Price"
            CalcGrid2
        Case "widtj"
            mPrice = GetItemPriceByWitdth(val(.TextMatrix(Row, .ColIndex("ItemId"))), val(.TextMatrix(Row, .ColIndex("widtj"))))
            If val(mPrice) = 0 Then
                mPrice = GetItemPrice(val(.TextMatrix(Row, .ColIndex("ItemId"))), , val(.TextMatrix(Row, .ColIndex("UnitId"))))
            End If
   
            .TextMatrix(Row, .ColIndex("Price")) = mPrice + GetItemAddPrice(val(DcboItemID1.BoundText))
            CalcGrid2
        Case "Trans_DiscountType"
            CalcGrid2
        Case "Trans_Discount"
            CalcGrid2
            
        Case "Trans_DiscountPercent"
           CalcGrid2
            
        
        End Select
        ReLineGrid Row
    End With
End Sub




Private Sub Fg2_Click()
    If FG2.Row <> 0 Then
        DcboItemID1.BoundText = val(FG2.TextMatrix(FG2.Row, FG2.ColIndex("ItemId")))
        DcboItemID4.BoundText = val(FG2.TextMatrix(FG2.Row, FG2.ColIndex("ItemId")))
        If val(FG2.TextMatrix(FG2.Row, FG2.ColIndex("LineID"))) = 0 Then
            DcboItemID4.Tag = FG2.Row
        Else
            DcboItemID4.Tag = val(FG2.TextMatrix(FG2.Row, FG2.ColIndex("LineID")))
        End If
        TxtAttachedItemCode = (FG2.TextMatrix(FG2.Row, FG2.ColIndex("ItemCode")))
        DcbUnit.BoundText = val(FG2.TextMatrix(FG2.Row, FG2.ColIndex("UnitId")))
        XPCboGroup.BoundText = val(FG2.TextMatrix(FG2.Row, FG2.ColIndex("GroupID")))
        txtQty1 = val(FG2.TextMatrix(FG2.Row, FG2.ColIndex("Qty")))
        TxtPrice = val(FG2.TextMatrix(FG2.Row, FG2.ColIndex("Price")))
        txtwidtj = val(FG2.TextMatrix(FG2.Row, FG2.ColIndex("widtj")))
        txthight = val(FG2.TextMatrix(FG2.Row, FG2.ColIndex("hight")))
        txtLength = val(FG2.TextMatrix(FG2.Row, FG2.ColIndex("Length")))
        txtTotalDisc = val(FG2.TextMatrix(FG2.Row, FG2.ColIndex("TotalDisc")))
        txtTotalAdd = val(FG2.TextMatrix(FG2.Row, FG2.ColIndex("TotalAdd")))
        txtNet = val(FG2.TextMatrix(FG2.Row, FG2.ColIndex("Net")))
        TxtVAt2 = val(FG2.TextMatrix(FG2.Row, FG2.ColIndex("Vat2")))
        txtTotalWithVat = val(FG2.TextMatrix(FG2.Row, FG2.ColIndex("TotalWithVat")))
        mIdDisplay = val(FG2.TextMatrix(FG2.Row, FG2.ColIndex("LineID")))
        If mIdDisplay = 0 Then mIdDisplay = mNewId
    End If
End Sub

Private Sub FG2_StartEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    If Not SystemOptions.IsMultiItemsInCompItem Then Cancel = True: Exit Sub
    If FG2.ColKey(Col) = "NoteSerial14" Then
        FG2.EditMaxLength = 10
        Exit Sub
    End If
    If Me.TxtModFlg.Text = "R" Then Cancel = True
    Select Case FG2.ColKey(Col)
     
        Case "Qty", "Price", "widtj", "hight", "Trans_DiscountType", "Trans_Discount", "Trans_DiscountPercent", "Select", "NoteSerial14"
            FG2.EditMaxLength = 10
        Case "Remark"
                FG2.EditMaxLength = 100
            Case "AreaL"
                FG2.EditMaxLength = 200
            Case "Select"
            
        Case Else
            Cancel = True
    End Select
End Sub

Private Sub FGDeleted_CellButtonClick(ByVal Row As Long, ByVal Col As Long)
    Select Case FGDeleted.ColKey(Col)
    Case "Redo"
        mRow = val(FGDeleted.TextMatrix(Row, FGDeleted.ColIndex("Row2")))
        
        FG.RowHidden(mRow) = False
        FGDeleted.RemoveItem Row
        FG.TextMatrix(mRow, FG.ColIndex("IsDeleted")) = 0
        CalcGrid2
       
    End Select
End Sub

Private Sub cmdAdd2_Click()
  Dim isFound As Boolean
        
        If DcboItemID4.Text = "" Then DcboItemID4.BoundText = DcboItemID1.BoundText
        Dim MethodCalc As Double
         Dim PartItemQty As Double
          Dim ForUnit As Double
        Dim rsDummy3 As New ADODB.Recordset
        If Me.DcboItemID3.Text <> "" And DcboItemID1.Text <> "" Then
            DcboItemID4.BoundText = DcboItemID1.BoundText
        End If
          '  If SystemOptions.IsMultiItemsInCompItem Then
                        
'                For i = 1 To FG.Rows - 1
'                    'If FG.ValueMatrix(i, FG.ColIndex("isReplaced")) And mNewId = val(FG.TextMatrix(i, FG.ColIndex("LineID"))) Then
'                    '    DeleteRow i, True
'                        LngNewRow = i
'                    '    isFound = False
'                    '    Exit For
'
'                        FG.TextMatrix(LngNewRow, FG.ColIndex("ItemID")) = Me.DcboItemID3.BoundText
'                        FG.TextMatrix(LngNewRow, FG.ColIndex("itemcode")) = Trim$(Me.TxtItemCode.Text)
'                        FG.TextMatrix(LngNewRow, FG.ColIndex("Fullcode")) = Trim$(Me.TxtItemCode.Text)
'
'
'                        FG.TextMatrix(LngNewRow, FG.ColIndex("itemname")) = Me.DcboItemID3.Text
'                        FG.TextMatrix(LngNewRow, FG.ColIndex("OldPrice")) = val(FG.TextMatrix(LngNewRow, FG.ColIndex("Total")))
'                        FG.TextMatrix(LngNewRow, FG.ColIndex("IsAdd")) = 0
'                        mPrice = GetItemPrice(FG.TextMatrix(LngNewRow, FG.ColIndex("ItemID")), , val(FG.TextMatrix(i, FG.ColIndex("UnitID"))))
'                       FG.TextMatrix(LngNewRow, FG.ColIndex("Price")) = mPrice
'                       FG.TextMatrix(LngNewRow, FG.ColIndex("Total")) = mPrice * val(FG.TextMatrix(i, FG.ColIndex("Qty")))
'
'                       ' mLineID = val(FG.TextMatrix(LngNewRow, FG.ColIndex("LineID")))
'                        GoTo NextStep
'
'                    'End If
'                Next
'            End If
'           If Not isFound Then
'                LngNewRow = ModFgLib.SetFgForNewRow(FG, FG.ColIndex("ItemID"))
'            End If
    If Me.DcboItemID3.Text <> "" Then
           Dim mmID As Long
           If val(DcboItemID4.Tag) = 0 Then
                mmID = FG2.FindRow(val(Me.DcboItemID4.BoundText), FG2.FixedRows, FG2.ColIndex("ItemID"), False, True)
            Else
                mmID = val(DcboItemID4.Tag)
            End If
            
            If mmID < 0 Then
                mmID = 1
'            Else
'                mmID = mNewId
            End If
          
             LngNewRow = ModFgLib.SetFgForNewRow(FG, FG.ColIndex("ItemID"))
          With FG
            .TextMatrix(LngNewRow, .ColIndex("LineID")) = mmID
            .TextMatrix(LngNewRow, .ColIndex("ItemID")) = Me.DcboItemID3.BoundText
            .TextMatrix(LngNewRow, .ColIndex("itemcode")) = Trim$(Me.txtItemCode.Text)
            .TextMatrix(LngNewRow, .ColIndex("Fullcode")) = Trim$(Me.txtItemCode.Text)
            .TextMatrix(LngNewRow, .ColIndex("itemname")) = Me.DcboItemID3.Text
            .TextMatrix(LngNewRow, .ColIndex("UnitID")) = Me.DcbUnit3.BoundText
            .TextMatrix(LngNewRow, .ColIndex("unitname")) = Me.DcbUnit3.Text

            
            If SystemOptions.AllowChangManualQtyMix = True Then
                .TextMatrix(LngNewRow, .ColIndex("Qty")) = val(Me.txtQty3.Text)
            Else
                .TextMatrix(LngNewRow, .ColIndex("Qty")) = val(Me.txtQty3.Text) * IIf(val(txtQty1.Text) = 0, 1, val(txtQty1.Text))
            End If
            .TextMatrix(LngNewRow, .ColIndex("IsRow")) = 1
            .TextMatrix(LngNewRow, .ColIndex("widtj")) = Me.txtwidtj2.Text
            .TextMatrix(LngNewRow, .ColIndex("hight")) = Me.txthight2.Text
            .TextMatrix(LngNewRow, .ColIndex("Length")) = Me.txtLength2.Text
            .TextMatrix(LngNewRow, .ColIndex("thickness")) = Me.txtthickness2.Text
            .TextMatrix(LngNewRow, .ColIndex("DO")) = Me.txtDO2.Text
            .TextMatrix(LngNewRow, .ColIndex("DI")) = Me.txtDI2.Text
            .TextMatrix(LngNewRow, .ColIndex("Diameter")) = Me.txtDiameter2.Text
            



            .TextMatrix(LngNewRow, .ColIndex("FlgX")) = val(Me.txtQty3.Text)
            .TextMatrix(LngNewRow, .ColIndex("TepQty")) = .TextMatrix(LngNewRow, .ColIndex("Qty"))
            .TextMatrix(LngNewRow, .ColIndex("IsAdd")) = 0
           
            mPrice = GetItemPrice(Me.DcboItemID3.BoundText, val(.TextMatrix(LngNewRow, .ColIndex("Qty"))), val(Me.DcbUnit3.BoundText))
            .TextMatrix(LngNewRow, .ColIndex("Price")) = mPrice
            .TextMatrix(LngNewRow, .ColIndex("Total")) = mPrice * val(.TextMatrix(LngNewRow, .ColIndex("Qty")))
            

                .TextMatrix(LngNewRow, .ColIndex("ItemID2")) = Me.DcboItemID4.BoundText
                .TextMatrix(LngNewRow, .ColIndex("itemcode2")) = Trim$(Me.TxtAttachedItemCode.Text)
                .TextMatrix(LngNewRow, .ColIndex("ItemName2")) = Me.DcboItemID4.Text
                If SystemOptions.IsGeometricProportions Then
                    StrSQL = " SELECT IsNull(MethodCalc,99) MethodCalc,IsNull(PartItemQty,99) PartItemQty,IsNull(ForUnit ,99)  ForUnit  FROM TblItemsUnits"
                    StrSQL = StrSQL & " WHERE ItemID =" & val(.TextMatrix(LngNewRow, .ColIndex("ItemID")))
                    StrSQL = StrSQL & " AND UnitID =" & val(.TextMatrix(LngNewRow, .ColIndex("UnitID")))
                    rsDummy3.Open StrSQL, Cn, adOpenKeyset, adLockReadOnly
                    If Not rsDummy3.EOF Then
                        MethodCalc = IIf(val(rsDummy3!MethodCalc & "") <> 99, val(rsDummy3!MethodCalc & ""), 0)
                        PartItemQty = IIf(val(rsDummy3!PartItemQty & "") <> 99, val(rsDummy3!PartItemQty & ""), 0)
                        ForUnit = IIf(val(rsDummy3!ForUnit & "") <> 99, val(rsDummy3!ForUnit & ""), 0)
                    End If
                End If
                .TextMatrix(LngNewRow, .ColIndex("ForUnit")) = ForUnit
                .TextMatrix(LngNewRow, .ColIndex("MethodCalc")) = MethodCalc
                .TextMatrix(LngNewRow, .ColIndex("PartItemQty")) = PartItemQty
          ' End If
           ' .TextMatrix(LngNewRow, .ColIndex("ItemPrice")) = val(Me.TxtItemPrice(0).text)
            .AutoSize 0, .Cols - 1, False
        End With
    End If
   'Else
        With FG
'        .TextMatrix(LngNewRow, .ColIndex("IsAdd")) = 1
'        .TextMatrix(LngNewRow, .ColIndex("ItemID2")) = Me.DcboItemID1.BoundText
'        .TextMatrix(LngNewRow, .ColIndex("itemcode2")) = Trim$(Me.TxtAttachedItemCode.Text)
'        .TextMatrix(LngNewRow, .ColIndex("ItemName2")) = Me.DcboItemID1.Text
        End With
      ' FillGrid2
    
NextStep:
    

'
'If chkIsAdd Then
'    FillGridItemType DcboItemID2.BoundText, DcboItemID2.Text, Trim$(TxtAttachedItemCode2.Text), 2, DcbUnit2.BoundText, DcbUnit2.Text, val(txtQty), 0, XPCboGroup2.BoundText, XPCboGroup2.Text
'End If

    'Me.lbl(21).Caption = ModFgLib.GetItemsInFg(FG, FG.ColIndex("ItemID"))
For i = 1 To FG2.Rows - 1
    
    ReLineGrid i, True, , , SystemOptions.IsGeometricProportions
    
Next

Me.txtItemCode.Text = ""
    Me.DcboItemID3.BoundText = ""
    
    'Me.TxtAttachedItemCode2.SetFocus
End Sub

Private Sub lbl_MouseMove(Index As Integer, _
                          Button As Integer, _
                          Shift As Integer, _
                          X As Single, _
                          Y As Single)

    If val(lbl(Index).Caption) <> 0 Then
        lbl(Index).ToolTipText = WriteNo(lbl(Index).Caption, 0, True)
    End If
    'ff

End Sub


Function ReloadCombos()
LoadCombosData
'LoadDataCombos
End Function
Private Sub LoadCombosData()
    Dim StrSQL As String
    
    Dcombos.GetSalesRepData DcboEmp
    
    Dcombos.GetUsers Me.DCboUserName
    Dcombos.GetBoxes Me.DcboBox
    Dcombos.GetBranches Me.dcBranch
    'Dcombos.get«hay Me.DCGroupI
    
    Dcombos.GetCustomersSuppliers 1, Me.DBCboClientName, , , 1
    Dcombos.GetStores Me.DCboStoreName
    
    Set cSearchDcbo(0) = New clsDCboSearch
    Set cSearchDcbo(0).Client = Me.DBCboClientName
    
 '   cSearchDcbo(0).SetBuddyText TxtSearchCode

    Set cSearchDcbo(1) = New clsDCboSearch
    Set cSearchDcbo(1).Client = Me.DCboStoreName
   ' cSearchDcbo(1).SetBuddyText Me.TxtStoreID

    Set cSearchDcbo(3) = New clsDCboSearch
    Set cSearchDcbo(3).Client = DcboEmp
 
End Sub




Private Sub cmdCreateSales_Click()
If TxtNoteSerial13 <> "" Then
    
    
    frmsalebill.show
    frmsalebill.XPBtnMove_Click (2)
    frmsalebill.Retrive val(TXTTransactionID3.Text)
End If

        
End Sub

Private Sub DBCboClientName_Change()
    If Me.TxtModFlg.Text = "N" Or Me.TxtModFlg.Text = "E" Then
            Dim DefaultSalesPersonId As Integer
         '    Me.DcboEmp.BoundText = ""
            Dim mFull As String
            GetCustomersDetail val(DBCboClientName.BoundText), DefaultSalesPersonId, mFull
            TxtSearchCode.Text = mFull
            TxtSearchCode2.Text = mFull
            If Not DefaultSalesPersonId = 0 Then

 '               Me.DcboEmp.BoundText = DefaultSalesPersonId
            End If
            GetCustomerNamebyPhone , , DBCboClientName.BoundText
            
        End If
End Sub

Private Sub DcboEmp_Change()
Dim StoreId As Integer
 If Me.TxtModFlg = "N" Or Me.TxtModFlg = "E" Then
         If val(Me.DcboEmp.BoundText) = 0 Then Exit Sub
           Me.TxtEmployeeID.Text = get_EMPLOYEE_Data(val(Me.DcboEmp.BoundText), "Fullcode")
        'DCEmP.text = DCEmP.text
'End If
' If Me.TxtModFlg = "N" Or Me.TxtModFlg = "E" Then
'  StoreId = get_StoreBYSalesPerson(val(Me.DcboEmp.BoundText))
' If StoreId <> 0 Then
' DCboStoreName.BoundText = StoreId
' End If
 
 End If

End Sub

Private Sub DcboEmp_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyF5 Then
ReloadCombos
End If
End Sub



Private Sub Tbar_ButtonClick(ByVal Button As MSComctlLib.Button)

End Sub

Private Sub TxtEmployeeID_Change()

If Me.TxtModFlg = "N" Or Me.TxtModFlg = "E" Then
    DcboEmp.BoundText = GeTEmpIDByEmpCode(TxtEmployeeID.Text, True)
End If

End Sub

Private Sub TxtEmployeeID_KeyDown(KeyCode As Integer, Shift As Integer)
   Dim EmpID As Integer

    If KeyCode = vbKeyReturn Then
        GetEmployeeIDFromCode TxtEmployeeID.Text, EmpID
        DcboEmp.BoundText = EmpID
    End If

End Sub

 

Private Sub C1Tab1_Click()

End Sub

Private Sub cmdCancel_Click()
    
If SystemOptions.CanTransferItemDef = False Then
        If SystemOptions.UserInterface = ArabicInterface Then
            Msg = "·« Ì„ﬂ‰ ⁄„· «· ÕÊÌ· ·«  „ ·ﬂ ’·«ÕÌ… ·–·ﬂ"
        Else
            Msg = "The conversion cannot be made. You do not have permission to do this"
        End If
        MsgBox Msg, vbOKOnly + vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
    Exit Sub
End If
    Dim s As String
    s = "Delete TblProductLineDistribution Where IDDefCIT = " & val(TxtTransSerial)
    Cn.Execute s
'    s = "Delete TblProductLineDistributionDet Where IDDefCIT = " & val(TxtTransSerial)
'    Cn.Execute s
    
    StrSqlDel = "delete From Transactions where Transaction_ID=" & val(Me.TXTTransactionID3.Text) 'Val(rs("Transaction_ID").value)
    Cn.Execute StrSqlDel, , adExecuteNoRecords
    
    
    StrSqlDel = "delete From Transaction_Details where Transaction_ID=" & val(Me.TXTTransactionID3.Text) 'Val(rs("Transaction_ID").value)
    Cn.Execute StrSqlDel, , adExecuteNoRecords
    
    
    DeleteTransactiomsVoucher2 val(TXTTransactionID5.Text)
    
    StrSqlDel = "delete From Notes where NoteID=" & val(Me.txtNoteid3.Text)  'Val(rs("Transaction_ID").value)
    Cn.Execute StrSqlDel, , adExecuteNoRecords
    StrSQL = "Delete From DOUBLE_ENTREY_VOUCHERS Where Transaction_ID=" & val(Me.TXTTransactionID3.Text)
    Cn.Execute StrSQL, , adExecuteNoRecords
    
    Cn.Execute "Delete from TransactionValueAdded where Transaction_ID=" & val(Me.TXTTransactionID3.Text) & ""

      
    StrSQL = "delete From Notes where noteid=" & val(txtNoteid3.Text)
    Cn.Execute StrSQL, , adExecuteNoRecords

    DeleteTransactiomsVoucher2 val(TXTTransactionID3.Text)
    DeleteTransactiomsVoucher2 val(TXTTransactionID4.Text)
    DeleteTransactiomsVoucher2 val(TXTTransactionID5.Text)
     StrSQL = "Update TblDefComItemDet Set QtyOut = 0 Where IDDefCIT =  " & val(TxtTransSerial)
    Cn.Execute StrSQL, , adExecuteNoRecords
    If SystemOptions.DontCreateOut Then
        Dim mNoteId As Long
        Dim mTransID As Long
        For i = 1 To FG3.Rows - 1
            mNoteId = val(FG3.TextMatrix(i, FG3.ColIndex("NoteID")))
            mTransID = val(FG3.TextMatrix(i, FG3.ColIndex("Transaction_ID")))
            DeleteTransactiomsVoucher2 (mTransID)
        Next
    End If
    
    
    StrSQL = "UPDATE TblDefComItem SET    NoteSerial13='0' ,TransactionID3 = 0,TransactionID4=0,TransactionID5 = 0,NoteSerial15 = 0,NoteSerial14 = 0, Allocated=0,AlloPay = 0 ,AlloRecep = 0 WHERE ID  =" & val(TxtTransSerial)
        Cn.Execute StrSQL

'   StrSQL = "UPDATE TblDefComItem SET Allocated=0,AlloPay = 0 ,AlloRecep = 0,TransactionID3 = 0, NoteSerial13 = 0, TransactionID1=" & val(0) & ",  NoteSerial11='" & 0 & "' WHERE ID  =" & val(TxtTransSerial)
'         Cn.Execute StrSQL
'            StrSQL = "UPDATE TblDefComItem SET  TransactionID2=" & val(0) & ",  NoteSerial12='" & 0 & "' WHERE ID  =" & val(TxtTransSerial)
'         Cn.Execute StrSQL
'
    
    'TXTTransactionID3.Text = rs!TransactionID3 & ""
    'TxtNoteSerial13.Text = rs!NoteSerial13 & ""
  TXTTransactionID3 = ""
    TxtNoteSerial13 = ""
    TXTTransactionID3 = ""
    TXTTransactionID5 = ""
    TxtNoteSerial15 = ""
TxtNoteSerial13 = ""
TxtNoteSerial15 = ""

TxtNoteSerial14 = ""

txtNoteid3 = ""
'TxtNoteSerial1 = ""

  StrSQL = "SELECT * FROM TblDefComItem "
StrSQL = StrSQL & "  WHERE      BranchId in(" & Current_branchSql & ")"
    StrSQL = StrSQL + " Order By ID"
    Set rs = New ADODB.Recordset
    rs.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText
    
If SystemOptions.UserInterface = ArabicInterface Then
    MsgBox " „ «·€«¡ «· ÕÊÌ·"
Else
    MsgBox "Conversion canceled"
End If
cmdTransfer.Enabled = True
cmdCancel.Enabled = False
Retrive val(TxtTransSerial)
    
End Sub

Private Sub cmdTransfer_Click()
   
    
If SystemOptions.CanTransferItemDef = False Then
        If SystemOptions.UserInterface = ArabicInterface Then
            Msg = "·« Ì„ﬂ‰ ⁄„· «· ÕÊÌ· ·«  „ ·ﬂ ’·«ÕÌ… ·–·ﬂ"
        Else
            Msg = "The conversion cannot be made. You do not have permission to do this"
        End If
        MsgBox Msg, vbOKOnly + vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
    Exit Sub
End If

Dim BeginTrans As Boolean
    If Trim(DCboStoreName.Text) = "" Then



        If SystemOptions.UserInterface = ArabicInterface Then
            Msg = "·« Ì„ﬂ‰ ⁄„· «· ÕÊÌ· ﬁ»· «œŒ«· „Œ“‰ «·»Ì⁄"
        Else
            Msg = "·The transfer can be made before entering the store"
        End If
        MsgBox Msg, vbOKOnly + vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
        DCboStoreName.SetFocus
        SendKeys "{F4}"
        
        Exit Sub
    End If
    
     
   If Trim(DCboStore2Name.Text) = "" Then
         MsgBox "·„ Ì „  ÕœÌœ Õ”«»  ··„Œ“Ê‰ «·”·⁄Ì ·Â–« «·„Œ“‰ ›Ì Â–« «·›—⁄    ", vbCritical
         Exit Sub
   End If
  
    If Trim(DCboStoreName.Text) = "" Then

        If SystemOptions.UserInterface = ArabicInterface Then
        
            Msg = "·« Ì„ﬂ‰ ⁄„· «· ÕÊÌ· ﬁ»· «œŒ«· „Œ“‰ «·’—›"
        Else
            Msg = "The conversion can not be done before the exchange store is inserted"
        End If
        MsgBox Msg, vbOKOnly + vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
        Exit Sub
    End If

    Dim RsTemp As New ADODB.Recordset
    Dim RsTest As New ADODB.Recordset
    Dim RowNum As Long
 
        For RowNum = 1 To FG.Rows - 1
             
             If FG.RowHidden(RowNum) Or CBool(FG.ValueMatrix(RowNum, FG.ColIndex("IsDeleted"))) = True Then GoTo NextRow333

            If FG.TextMatrix(RowNum, FG.ColIndex("ItemID")) <> "" Then
                
                
                 If SystemOptions.SysAllowStockNegative = False Then
                        
                            
                    StrSQL = "Select * From TblItems where ItemID=" & val(FG.TextMatrix(RowNum, FG.ColIndex("ItemID")))
                    Set RsTemp = New ADODB.Recordset
                    RsTemp.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

                    If Not RsTemp.EOF Then
                        
                        If DCboStoreName.BoundText <> "" Then
                            Set RsTest = New ADODB.Recordset
                            Set RsTest = GetItemQuantityStock(val(FG.TextMatrix(RowNum, FG.ColIndex("ItemID"))), val(Me.DCboStoreName.BoundText), , , True, , Trim(FG.TextMatrix(RowNum, FG.ColIndex("Serial"))))

                            If RsTest.EOF Or RsTemp.BOF Then
                                Msg = "«·ﬁÿ⁄… –«  «·”Ì—Ì«· : "
                                Msg = Msg + " «·’‰› : " & Trim(FG.Cell(flexcpTextDisplay, RowNum, FG.ColIndex("ItemName"))) & CHR(13) & "Ê«·„ÊÃÊœ ›Ï «·”ÿ— —ﬁ„  " & RowNum
                                Msg = Msg + " €Ì— „ÊÃÊœ… ›Ì «·„Œ“‰ «·„Õœœ" & CHR(13)
                                Msg = Msg + "Ê»«· «·Ï ·„ Ì „ «‰‘«¡ «–‰ «·’—›"
                
                                MsgBox Msg
                                Exit Sub
                            End If
                        End If
                    End If
                End If
            End If
NextRow333:
        Next
    
      Dim rsOut As New ADODB.Recordset
            Dim Current_case As Integer, s As String, mBoxID As Long
            Set rsOut = New ADODB.Recordset
            s = "Select BoxID From TblBoxesData Where Empid = " & val(Me.DcboEmp.BoundText)
            rsOut.Open s, Cn, adOpenStatic, adLockReadOnly
            If Not rsOut.EOF Then
                mBoxID = val(rsOut!BoxID & "")
            End If
            mBoxID = val(DcboBox.BoundText)
            
            StrTempAccountCode = GetMyAccountCode("TblBoxesData", "BoxID", mBoxID)
            If SystemOptions.IsMultiItemsInCompItem Then
                If CboPayMentType.ListIndex = 0 Then
                    If StrTempAccountCode = "" Then
                        If SystemOptions.UserInterface = ArabicInterface Then
                            MsgBox "·« Ì„ﬂ‰ «‰‘«¡ «·›« Ê—… ·⁄œ„ ÊÃÊœ Œ“Ì‰… ··„‰œÊ»"
                        Else
                            MsgBox "The invoice can not be created because there is no safe for the delegate"
                        End If
                        Exit Sub
                    End If
                End If
            End If
        
        
  
       Dim BranchID  As Double, StoreId As Double
              
        BranchID = val(dcBranch.BoundText)
        StoreId = val(DCboStoreName.BoundText)
        If SystemOptions.IsMultiItemsInCompItem Then
            If DBCboClientName.BoundText = "" Then
                If SystemOptions.UserInterface = ArabicInterface Then
                    MsgBox ("·« Ì„ﬂ‰ «‰‘«¡ «·›« Ê—… »œÊ‰ «œŒ«· «·⁄„Ì·")
                Else
                     MsgBox ("The invoice can not be created without entering the customer")
                    End If
                     
                DBCboClientName.SetFocus
                Exit Sub
            End If
            
            If DcboEmp.BoundText = "" Then
                If SystemOptions.UserInterface = ArabicInterface Then
                    MsgBox ("·« Ì„ﬂ‰ «‰‘«¡ «·›« Ê—… »œÊ‰ «œŒ«· «·„‰œÊ»")
                Else
                    MsgBox ("The invoice can not be created without the introduction of a salesman")
                End If
                
                DcboEmp.SetFocus
                Exit Sub
            End If
                    
                    
            Cn.BeginTrans
            BeginTrans = True
        
            
            StrSqlDel = "delete From Transactions where Transaction_ID=" & val(Me.TXTTransactionID3.Text) 'Val(rs("Transaction_ID").value)
            Cn.Execute StrSqlDel, , adExecuteNoRecords
            
            
            StrSqlDel = "delete From Transaction_Details where Transaction_ID=" & val(Me.TXTTransactionID3.Text) 'Val(rs("Transaction_ID").value)
            Cn.Execute StrSqlDel, , adExecuteNoRecords
            
            StrSqlDel = "delete From Notes where NoteID=" & val(Me.txtNoteid3.Text)  'Val(rs("Transaction_ID").value)
            Cn.Execute StrSqlDel, , adExecuteNoRecords
            StrSQL = "Delete From DOUBLE_ENTREY_VOUCHERS Where Transaction_ID=" & val(Me.TXTTransactionID3.Text)
            Cn.Execute StrSQL, , adExecuteNoRecords
            
            Cn.Execute "Delete from TransactionValueAdded where Transaction_ID=" & val(Me.TXTTransactionID3.Text) & ""
    
              

        
          If TxtNoteSerial13.Text = "" Then
                TxtNoteSerial13.Text = Voucher_coding(val(val(dcBranch.BoundText)), XPDtbBill.value, 7, 170, , 21, , val(DCboStoreName.BoundText))
          End If
                    
            If SystemOptions.TransferNotInvItemDef = False Then
                CreateSalesTrans BranchID, 0, XPDtbBill.value, 21, 0, val(user_id), 0, DBCboClientName.BoundText, StoreId, CboPayMentType.ListIndex, DcboEmp.BoundText, "›« Ê—… „»Ì⁄«  »‰«¡« ⁄·Ï ”‰œ  Ã„Ì⁄ —ﬁ„ " & TxtTransSerial & "   " & TxtMaxName.Text
            End If
        End If
         
        StrSQL = "UPDATE TblDefComItem SET  TransactionID3=" & val(TXTTransactionID3) & ",  NoteSerial13='" & TxtNoteSerial13 & "' WHERE ID  =" & val(TxtTransSerial)
        Cn.Execute StrSQL
    
  
    Dim mQty2 As Double
    For i = 1 To FG2.Rows - 1
        mQtyTotal = 0
        mQty2 = val(FG2.TextMatrix(i, FG2.ColIndex("Qty")))
        SaveItemsProduction True, mQty2, i
'        If mQtyTotal <> val(mQty2) Then
'            mTotalSecond = Abs(val(mQty2) - mQtyTotal)
'            SaveItemsProduction False, mQty2, i
'        End If
    Next
    
    Selct(1).value = vbChecked
    '   Dim BranchID As Double
    '    Dim StoreId As Double
        
    DeleteTransactiomsVoucher2 val(TXTTransactionID5.Text)
        
        
      '  If Selct(1).value = vbChecked Then
    BranchID = val(dcBranch.BoundText)
    StoreId = val(DCboStoreName.BoundText)
      If Not SystemOptions.DontCreateOut2 Then
        createVoucher BranchID, 0, XPDtbBill.value, 19, 0, val(user_id), 0, val(DBCboClientName.BoundText), StoreId, 0, 0, "”‰œ  ’—› »‰«¡ ⁄·Ì ›« Ê—… „»Ì⁄«  —ﬁ„ : " & TxtNoteSerial13 & " »‰«¡« ⁄·Ï ”‰œ  Ã„Ì⁄ —ﬁ„ " & TxtTransSerial, 1
    End If
'End If


    BranchID = val(dcBranch.BoundText)
    StoreId = val(DCboStoreName.BoundText)

    StrSQL = "UPDATE TblDefComItem SET Allocated=1,AlloPay = 1 ,AlloRecep = 1,  TransactionID1=" & val(TXTTransactionID1) & ",  NoteSerial11='" & TxtNoteSerial11 & "'"
    StrSQL = StrSQL & " ,TransactionID2=" & val(TXTTransactionID2) & ",  NoteSerial12='" & TxtNoteSerial12 & "'"
    StrSQL = StrSQL & " Where ID = " & val(TxtTransSerial)
    Cn.Execute StrSQL
    StrSQL = "UPDATE TblDefComItem SET  TransactionID2=" & val(TXTTransactionID2) & ",  NoteSerial12='" & TxtNoteSerial12 & "' WHERE ID  =" & val(TxtTransSerial)
    Cn.Execute StrSQL
         
    rs.Resync
            
    Cn.CommitTrans
    cmdTransfer.Enabled = False
    cmdCancel.Enabled = True
  
End Sub

Private Sub SaveItemsProduction(ByVal IsFirst As Boolean, ByVal mQty22 As Double, ByVal mRow As Long)
    
    Dim s As String, s2 As String, mCount As Long, mQty As Double, mPart As Double, mMod As Double, i As Integer, mAvgQty As Double, mIsSecond As Boolean
    Dim mItemNo As Long, mUnitNo As Long, mGroupID As Long, mLineID As Long
    mItemNo = val(FG2.TextMatrix(mRow, FG2.ColIndex("ItemID")))
    mUnitNo = val(FG2.TextMatrix(mRow, FG2.ColIndex("UnitID")))
    mGroupID = val(FG2.TextMatrix(mRow, FG2.ColIndex("GroupID")))
    mLineID = val(FG2.TextMatrix(mRow, FG2.ColIndex("LineID")))
   
    
    Dim RsData As New ADODB.Recordset
    Dim RsDataLine As New ADODB.Recordset
    
        s = "SELECT Count(*) CC FROM TblProductLine Where IsBasicLine = 1"
        RsData.Open s, Cn, adOpenStatic, adLockReadOnly, adCmdText
        If Not RsData.EOF Then
            mCount = val(RsData!CC & "")
        End If
        If mCount = 0 Then
            
           ' MsgBox "·« ÌÊÃœ ŒÿÊÿ «‰ «Ã „⁄—›… "
            Exit Sub
        End If
        s = "SELECT * FROM TblProductLineDistribution Where "
        s = s & "  ItemNameID = " & val(mItemNo)
        s = s & " and UnitID = " & val(mUnitNo)
        s = s & " and LineID = " & val(mLineID)
        s = s & " and IDDefCIT = " & TxtTransSerial
        Set RsData = New ADODB.Recordset
        Cn.CommandTimeout = 10000
        RsData.Open s, Cn, adOpenKeyset, adLockOptimistic, adCmdText
        
        Dim mValuePart As Double
    
        
        i = 0
    
        
        s = " SELECT * FROM ("
        s = s & "  SELECT SUM(Qty) Qty,ItemNameID,UnitID,T2.ID FROM TblProductLineDistribution T"
        s = s & " RIGHT Outer JOIN TblProductLine T2 ON T2.id =T.ProductLineID"

        s = s & " and UnitId = " & val(mUnitNo)
        s = s & " and ItemNameID = " & val(mItemNo)
        s = s & " and LineID = " & val(mLineID)
        s = s & " Where IsBasicLine = 1"
        s = s & " Group BY ItemNameID,UnitID,T2.ID"
        s = s & " ) T "
        s = s & " Order BY T.Qty DESC "
        

        
        s = "Select ProductLineId as ID from TblItemProductLine Where ItemID = " & val(mItemNo)
        Dim isFirstTime As Boolean
'        RsDataLine.Close
        RsDataLine.Open s, Cn, adOpenStatic, adLockReadOnly, adCmdText
        If RsDataLine.EOF Then
            RsDataLine.Close
            isFirstTime = True
            s = "SELECT *,Qty = 0 FROM TblProductLine Where IsBasicLine = 1"
            RsDataLine.Open s, Cn, adOpenStatic, adLockReadOnly, adCmdText
        End If
        Dim mId As Long, mDec As Double, mQtyVal As Long, mQtyCompare As Double, mQtyNew As Double, Total As Double
        Do While Not RsDataLine.EOF
            i = i + 1
            
            
                    RsData.AddNew
                    RsData!ItemNameID = val(mItemNo)
                    RsData!UnitID = val(mUnitNo)
                    RsData!GroupID = val(mGroupID)
                    RsData!LineID = val(mLineID)
                    RsData!IDDefCIT = val(TxtTransSerial)
                    RsData!ProductLineID = RsDataLine!ID
                    RsData!SalesID = val(TxtNoteSerial13)
                    RsData!Qty1 = val(mQty22)
                    RsData!Qty = val(mQty22)
                   ' mQtyVal = RoundDown(Abs(mPart + mAvgQty - val(RsDataLine!Qty & "")))
                    'mQtyVal = Round(Abs(mPart + mAvgQty - val(RsDataLine!Qty & "")))
        
        
                    'RsData!Qty = mQtyNew
                    'If i <> mCount Then
                       ' RsData!Qty = mQtyNew
'                    Else
'                        If isFirstTime Then
'                            RsData!Qty = (mPart - ((mPart * mCount) - val(mQty22)))
'                        Else
'                            RsData!Qty = mQtyNew
'                        End If
'                    End If
                    mId = CStr(new_id("TblProductLineDistribution", "ID", "", True))
                    
                    RsData!ID = mId
                    RsData.update
            

ExitLoop:

            RsDataLine.MoveNext
        Loop
   
End Sub



Private Sub SaveItemsProduction2(ByVal IsFirst As Boolean)
    
    Dim s As String, s2 As String, mCount As Long, mQty As Double, mPart As Double, mMod As Double, i As Integer, mAvgQty As Double, mIsSecond As Boolean
    
    
    
    Dim RsData As New ADODB.Recordset
    Dim RsDataLine As New ADODB.Recordset
    
        s = "SELECT Count(*) CC FROM TblProductLine "
        RsData.Open s, Cn, adOpenStatic, adLockReadOnly, adCmdText
        If Not RsData.EOF Then
            mCount = val(RsData!CC & "")
        End If
        If mCount = 0 Then
            
            MsgBox "·« ÌÊÃœ ŒÿÊÿ «‰ «Ã „⁄—›… "
            Exit Sub
        End If
        s = "SELECT * FROM TblProductLineDistribution Where "
        s = s & "  ItemNameID = " & val(DcboItemID1.BoundText)
        s = s & " and UnitID = " & val(DcbUnit.BoundText)
        s = s & " and IDDefCIT = " & TxtTransSerial
        Set RsData = New ADODB.Recordset
        Cn.CommandTimeout = 10000
        RsData.Open s, Cn, adOpenKeyset, adLockOptimistic, adCmdText
        
        Dim mValuePart As Double
    
        
        i = 0
    
        
        s = " SELECT * FROM ("
        s = s & "  SELECT SUM(Qty) Qty,ItemNameID,UnitID,T2.ID FROM TblProductLineDistribution T"
        s = s & " RIGHT Outer JOIN TblProductLine T2 ON T2.id =T.ProductLineID"

        s = s & " and UnitId = " & val(DcbUnit.BoundText)
        s = s & " and ItemNameID = " & val(DcboItemID1.BoundText)
        s = s & " Group BY ItemNameID,UnitID,T2.ID"
        s = s & " ) T "
        s = s & " Order BY T.Qty DESC "

        Dim isFirstTime As Boolean
'        RsDataLine.Close
        RsDataLine.Open s, Cn, adOpenStatic, adLockReadOnly, adCmdText
        If RsDataLine.EOF Then
            RsDataLine.Close
            isFirstTime = True
            s = "SELECT *,Qty = 0 FROM TblProductLine "
            RsDataLine.Open s, Cn, adOpenStatic, adLockReadOnly, adCmdText
        End If
        Dim mId As Long, mDec As Double, mQtyVal As Long, mQtyCompare As Double, mQtyNew As Double, Total As Double
        Do While Not RsDataLine.EOF
            i = i + 1
            If i = 1 Then
                mQtyCompare = val(RsDataLine!Qty & "")
            End If
            If IsFirst Then
                If isFirstTime Then
                    mPart = Round(val(mQty22) / mCount)
                    mQtyNew = mPart
                Else
                    mQtyNew = mQtyCompare - val(RsDataLine!Qty & "")
                End If
                
                    If i = mCount Then
                        If isFirstTime And Not IsFirst Then
                            mQtyNew = (mPart - ((mPart * mCount) - val(mQty22)))
                        End If
                    End If
'                    If i = mCount Then
'                        If Not isFirstTime And IsFirst Then
'                            mQtyNew = val(txtQty1) - mQtyTotal
'                        End If
'                    End If
                    
                        
                mQtyTotal = mQtyTotal + mQtyNew
                If mQtyTotal > val(mQty22) Then
                    
                    mQtyTotal = mQtyTotal - mQtyNew
                    mQtyNew = 0
                End If
            End If
         '   If (mQtyTotal > val(mQty22) And Not isFirstTime) And Not IsFirst Then GoTo ExitLoop
                If IsFirst Then
                    RsData.AddNew
                    RsData!ItemNameID = val(DcboItemID1.BoundText)
                    RsData!UnitID = val(DcbUnit.BoundText)
                    RsData!GroupID = val(XPCboGroup.BoundText)
                    RsData!IDDefCIT = val(TxtTransSerial)
                    RsData!ProductLineID = RsDataLine!ID
                    RsData!Qty1 = val(mQty22)
                    mQtyVal = RoundDown(Abs(mPart + mAvgQty - val(RsDataLine!Qty & "")))
                    mQtyVal = Round(Abs(mPart + mAvgQty - val(RsDataLine!Qty & "")))
        
        
                    'RsData!Qty = mQtyNew
                    If i <> mCount Then
                        RsData!Qty = mQtyNew
                    Else
                        If isFirstTime Then
                            RsData!Qty = (mPart - ((mPart * mCount) - val(mQty22)))
                        Else
                            RsData!Qty = mQtyNew
                        End If
                    End If
                    mId = CStr(new_id("TblProductLineDistribution", "ID", "", True))
                    
                    RsData!ID = mId
                    RsData.update
                Else
                    RsData.Close
                    
                    s = "SELECT * FROM TblProductLineDistribution Where "
                    s = s & "  ItemNameID = " & val(DcboItemID1.BoundText)
                    s = s & " and UnitID = " & val(DcbUnit.BoundText)
                    s = s & " and IDDefCIT = " & val(TxtTransSerial)
                    s = s & " and ProductLineID = " & val(RsDataLine!ID)
                    
                    Set RsData = New ADODB.Recordset
                    Cn.CommandTimeout = 10000
                    RsData.Open s, Cn, adOpenKeyset, adLockOptimistic, adCmdText
                    
                    mPart = Round(mTotalSecond / mCount)
                    Total = Total + mPart
                    If Total <= mTotalSecond Then
                    
                        If i <> mCount Then
                            RsData!Qty = RsData!Qty + mPart
                        Else
                            RsData!Qty = Abs(RsData!Qty + (mPart - ((mPart * mCount) - mTotalSecond)))
                        End If
                    End If
                    
                    RsData.update
                    
                End If
ExitLoop:

            RsDataLine.MoveNext
        Loop
   
End Sub



Public Function RoundDown(DblValue As Double) As Double

    On Error GoTo Err
    
    Dim myDec As Long
    myDec = InStr(1, CStr(DblValue), ".", vbTextCompare)
    If myDec > 0 Then
        RoundDown = CDbl(Left(CStr(DblValue), myDec))
    Else
        RoundDown = DblValue
    End If
    Exit Function
    
Err:
        Resume Next

End Function
Public Function roundUp(DblValue As Double) As Double



Dim myDec As Long
myDec = InStr(1, CStr(DblValue), ".", vbTextCompare)
If myDec > 0 Then
    roundUp = CDbl(Left(CStr(DblValue), myDec)) + 1
Else
    roundUp = DblValue
End If

Exit Function



End Function


 
Private Sub ISButton1_Click(Index As Integer)
 
If Index <= 2 Or Index = 4 Then
    Dim StrSQL As String
    Dim StrWhere As String


   
    Dim MySQL As String
    Dim RsData As New ADODB.Recordset
    Dim xApp As New CRAXDRT.Application
    Dim xReport As CRAXDRT.Report
    Dim CViewer As ClsReportViewer
    Dim StrReportTitle As String
    Dim StrFileName As String
 '   Dim Msg As String

Dim Balance As String, balanceString As String
If SystemOptions.ShowBalanceCustInv Then
Dim mAccount As String

    mAccount = ModAccounts.GetMyAccountCode("TblCustemers", "CusID", val(Me.DBCboClientName.BoundText), "Account_code2")
    WriteCustomerBalPublic mAccount, Balance, balanceString, , , , , , XPDtbBill.value
    
End If


    MySQL = " SELECT    TblDefComItem.PaymentType,Grou.GroupName,TblDefComItem.PaymentType ,TblDefComItem.id Transaction_ID, TblDefComItem.Qty1,  dbo.TblDefComItemDet.ItemID, dbo.TblItems.ItemCode, dbo.TblItems.ItemName, "
    
    MySQL = " SELECT TblDefComItem.ID,Grou.GroupName, TblDefComItem.Period TransactionID2,      TblDefComItem.PaymentType,TblDefComItem.Vat2,TblDefComItem.TotalWithVat,       TblDefComItem.id    Transaction_ID, TblDefComItem.Qty1,"
    MySQL = MySQL & "        dbo.TblItems.ItemCode,       dbo.TblItems.ItemName,       dbo.TblItems.ItemNamee,dbo.TblDefComItem.RecordDate,"
    MySQL = MySQL & "        dbo.TblDefComItem.CusID,dbo.TblCustemers.CusName,dbo.TblCustemers.CusNamee,TblCustemers.VATNO,TblCustemers.ZipCode ,'" & DcboEmp.Text & "' as ResponsibleContact,dbo.TblDefComItem.BranchID,dbo.TblBranchesData.branch_name,"
    MySQL = MySQL & "       dbo.TblBranchesData.branch_nameE,dbo.TblDefComItem.ItemNameID,TblDefComItem.widtj,"
    MySQL = MySQL & "      TblDefComItem.hight,TblDefComItem.Price           ,dbo.TblDefComItem.TotalAdd,"
    MySQL = MySQL & "       TblDefComItem.TotalDisc,TblDefComItem.Net,"
    MySQL = MySQL & "       tu.UnitName UnitName2"
    
    MySQL = MySQL & " From dbo.TblItems"
         
    MySQL = MySQL & "       RIGHT OUTER JOIN dbo.TblBranchesData"
         
    MySQL = MySQL & "       RIGHT OUTER JOIN dbo.TblDefComItem"
         
    MySQL = MySQL & "            ON  dbo.TblBranchesData.branch_id = dbo.TblDefComItem.BranchID"
    MySQL = MySQL & "       LEFT OUTER JOIN dbo.TblCustemers"
    MySQL = MySQL & "            ON  dbo.TblDefComItem.CusID = dbo.TblCustemers.CusID"
         
    MySQL = MySQL & "            ON  dbo.TblItems.ItemID = dbo.TblDefComItem.ItemNameID"
    
    MySQL = MySQL & "       LEFT OUTER JOIN TblUnites  AS tu"
    MySQL = MySQL & "            ON  tu.UnitID = dbo.TblDefComItem.UnitID"
    MySQL = MySQL & "       LEFT OUTER JOIN Groups     AS Grou"
    MySQL = MySQL & "            ON  Grou.GroupID = dbo.TblDefComItem.GroupID "
    MySQL = MySQL & "  Where (dbo.TblDefComItem.id = " & val(Me.TxtTransSerial.Text) & ")"

MySQL = " SELECT TblDefComItem.ID,"
MySQL = MySQL & "         Grou.GroupName,"
MySQL = MySQL & "         TblDefComItem.PaymentType,"
MySQL = MySQL & "         tdcid.Vat2,"
MySQL = MySQL & "         tdcid.TotalWithVat,"
MySQL = MySQL & "         TblDefComItem.id              Transaction_ID,"
MySQL = MySQL & "         tdcid.Qty Qty1,TblDefComItem.Period TransactionID2,"
MySQL = MySQL & "         dbo.TblItems.ItemCode,"
MySQL = MySQL & "         dbo.TblItems.ItemName,"
MySQL = MySQL & "         Item5.ItemName  BuiltInItemName  ,"
MySQL = MySQL & "         dbo.TblItems.ItemNamee,Item2.ItemName ItemName2,"
MySQL = MySQL & "         dbo.TblDefComItem.RecordDate,"
MySQL = MySQL & "         dbo.TblDefComItem.CusID,"
MySQL = MySQL & "         dbo.TblCustemers.FullCode,TblCustemers.Address,TblCustemers.Mobile1,TblCustemers.VATNO,TblCustemers.ZipCode , '" & DcboEmp.Text & "' as ResponsibleContact,"
MySQL = MySQL & "         dbo.TblCustemers.CusName,"
MySQL = MySQL & "         dbo.TblCustemers.CusNamee,"
MySQL = MySQL & "         dbo.TblDefComItem.BranchID,"
MySQL = MySQL & "         dbo.TblBranchesData.branch_name,"
MySQL = MySQL & "         dbo.TblBranchesData.branch_nameE,"
MySQL = MySQL & "         tdcid.ItemID,"
MySQL = MySQL & "         tdcid.widtj,"
MySQL = MySQL & "         tdcid.hight,"
MySQL = MySQL & "         tdcid.Price,"
MySQL = MySQL & "         tdcid.TotalAdd,tdcid.Remark as MaxName,"
MySQL = MySQL & "         tdcid.TotalDisc,"
MySQL = MySQL & "         tdcid.Net,tdcid.Vat2 Vat22,"
MySQL = MySQL & "         tu.UnitName UnitName2"
MySQL = MySQL & "                      ,BalnceCust = " & val(Balance) & ",tdcid.AreaL"
MySQL = MySQL & "  From dbo.TblItems"
MySQL = MySQL & "         RIGHT OUTER JOIN dbo.TblBranchesData"
MySQL = MySQL & "         RIGHT OUTER JOIN dbo.TblDefComItem"
MySQL = MySQL & "              ON  dbo.TblBranchesData.branch_id = dbo.TblDefComItem.BranchID"
MySQL = MySQL & "         LEFT OUTER JOIN dbo.TblDefComItemData AS tdcid"
MySQL = MySQL & "         ON             tdcid.IDDefCIT = TblDefComItem.ID"
MySQL = MySQL & "         LEFT OUTER JOIN dbo.TblCustemers"
MySQL = MySQL & "              ON  dbo.TblDefComItem.CusID = dbo.TblCustemers.CusID"
MySQL = MySQL & "              ON  dbo.TblItems.ItemID = tdcid.ItemID"
MySQL = MySQL & "         LEFT OUTER JOIN TblUnites  AS tu"
MySQL = MySQL & "              ON  tu.UnitID = dbo.TblDefComItem.UnitID"
MySQL = MySQL & "         LEFT OUTER JOIN Groups     AS Grou"
MySQL = MySQL & "              ON  Grou.GroupID = tdcid.GroupID"
MySQL = MySQL & "         LEFT OUTER JOIN TblItems     AS Item2"
MySQL = MySQL & "              ON  Item2.ItemID = tdcid.ItemId2"

MySQL = MySQL & "         LEFT OUTER JOIN Groups     AS Grou5"
MySQL = MySQL & "              ON  Grou5.GroupID = tdcid.GroupIDBuiltin"
MySQL = MySQL & "         LEFT OUTER JOIN TblItems     AS Item5"
MySQL = MySQL & "              ON  Item5.ItemID = tdcid.BuiltinItemID"


MySQL = MySQL & "  Where (dbo.TblDefComItem.id = " & val(Me.TxtTransSerial.Text) & ")"




'StrWhere = StrWhere & " order by TblStuFingerprint.StudID"
  StrSQL = MySQL & StrWhere
  print_report2 StrSQL, Index
ElseIf Index = 3 Then
            
     Dim mItemsCodes As String
     Dim i As Integer
     mItemsCodes = val(DcboItemID1.BoundText)
     With FG2
     For i = 1 To FG2.Rows - 1
        If val(.TextMatrix(i, .ColIndex("ItemID"))) <> 0 Then
            mItemsCodes = mItemsCodes & "," & val(.TextMatrix(i, .ColIndex("ItemID")))
        End If
        
        
     Next
     End With
            MySQL = " SELECT   TT.ItemName MainItemName,       TT.ItemID MainItemID , TT.ItemID, ForUnit, MethodCalc, TblItemsParts.lowering, TblItemsParts.increase, dbo.TblItemsParts.UnitID, dbo.TblItemsParts.isReplaced, dbo.TblItemsParts.PartItemPrice, dbo.TblItemsParts.PartItemQty, dbo.TblItemsParts.PartItemID, "
    MySQL = MySQL + " Price = dbo.GetItemPrice(dbo.TblItemsParts.PartItemID,dbo.TblUnites.UnitID," & IIf(SystemOptions.AllowLastPrice, 1, 0) & "),"
    MySQL = MySQL + "      dbo.TblItemsParts.ItemID, dbo.TblItemsParts.TableID, dbo.TblUnites.UnitName, dbo.TblUnites.UnitNamee, dbo.TblItems.ItemCode, dbo.TblItems.ItemName,"
    MySQL = MySQL + "      dbo.TblItems.ItemNamee , dbo.TblItems.fullcode"
    MySQL = MySQL + "  FROM         dbo.TblItemsParts INNER JOIN"
    MySQL = MySQL + "      dbo.TblUnites ON dbo.TblItemsParts.Unitid = dbo.TblUnites.UnitID RIGHT OUTER JOIN"
    MySQL = MySQL + "      dbo.TblItems ON dbo.TblItemsParts.PartItemID = dbo.TblItems.ItemID"
    MySQL = MySQL + "                 RIGHT OUTER JOIN dbo.TblItems TT"
    MySQL = MySQL + "                  ON  dbo.TblItemsParts.ItemID = TT.ItemID"
    MySQL = MySQL + " Where (dbo.TblItemsParts.ItemID In ( " & mItemsCodes & "))    "
    MySQL = MySQL + " ORDER BY dbo.TblItemsParts.TableID"
    '       Rs3.Close
        
          print_report3 MySQL, Index
ElseIf Index = 10 And val(TxtNoteSerial13) <> 0 Then

 
 Dim SaleReport As ClsSaleReport


     If DoPremis(Do_Print, Me.Name, True) = False Then
                Exit Sub
            End If

            If Me.TXTTransactionID3.Text = "" Then
                If SystemOptions.UserInterface = ArabicInterface Then
                    Msg = "·« ÊÃœ ›Ê« Ì— ·Ì „ ÿ»«⁄ Â«"
                Else
                    Msg = "There are no invoices to print"
                End If
                MsgBox Msg, vbOKOnly + vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
                Exit Sub
            End If

            AskOption = GetSetting(StrAppRegPath, "View_Type", "ShowMe", False)

          If AskOption = False Then
'             FrmSallReportOptions.show vbModal
'
'              If FrmSallReportOptions.UserCanceled = True Then
'                   Unload FrmSallReportOptions
'
'             Exit Sub
'               End If
'
'            Unload FrmSallReportOptions
            End If
        updateCopyNo "Transactions", "CopyNO", "Transaction_ID", val(Me.TXTTransactionID3.Text)
        
        If TXTTransactionID3.Text <> "" Then
            Set SaleReport = New ClsSaleReport
            SaleReport.ShowSallingDataDetailed TXTTransactionID3.Text, 18, , , Round(val(txtTotalWithVat2), SystemOptions.Count_ACCOUNT_digit), TxtSearchCode.Text, , , , , , XPDtbBill.value, , , , , , , , , val(DcCurrency.BoundText), , , val(Me.dcBranch.BoundText)
        
            '  If MDIFrmMain.MnuInvPrintReceipt.Checked = True Then
            '      SaleReport.PrintInvoiceReceipt Val(XPTxtBillID.text), P_Target
            '  End If
        End If
        rs.Resync adAffectCurrent
            
End If
Exit Sub


End Sub
Function print_report2(Optional NoteSerial As String, Optional Ind As Integer = 0)
     
   
     
    'Dim MySQL As String
    Dim RsData As New ADODB.Recordset
    Dim xApp As New CRAXDRT.Application
    Dim xReport As CRAXDRT.Report
    Dim CViewer As ClsReportViewer
    Dim StrReportTitle As String
    Dim StrFileName As String
    Dim Msg As String
    
     Set RsData = New ADODB.Recordset
    RsData.Open NoteSerial, Cn, adOpenStatic, adLockReadOnly, adCmdText
    If Ind = 4 Then
        StrFileName = App.path & "\REPORTS\REPORTS NEW\" & "EnSalesInvoiceCompItem2.rpt"
        GoTo PrintFile
    End If
 
    If Ind = 9 Then
        StrFileName = App.path & "\REPORTS\REPORTS NEW\" & "CompItemNew.rpt"
        GoTo PrintFile
    End If
 
    
   If Ind <> 2 Then
        If SystemOptions.UserInterface = ArabicInterface Then
            Msg = "”Ì „ ÿ»«⁄… «· ﬁ—Ì— »«··€… «·«‰Ã·Ì“Ì…  " & CHR(13)
            Msg = Msg + "«÷€ÿ ‰⁄„ ··„Ê«›ﬁ… «Ê ·« ··ÿ»«⁄… »«·⁄—»Ì…"
        Else
            Msg = "The report will be printed in English  " & CHR(13)
            Msg = Msg + "Click Yes to approve or not to print in Arabic"
        End If

        If MsgBox(Msg, vbYesNo + vbQuestion + vbMsgBoxRight + vbMsgBoxRtlReading, App.title) = vbNo Then
        
            StrFileName = App.path & "\REPORTS\REPORTS NEW\" & "ArabSalesInvoiceCompItem.rpt"
        Else
            StrFileName = App.path & "\REPORTS\REPORTS NEW\" & "EnSalesInvoiceCompItem.rpt"
        End If
    ElseIf Ind = 2 Then
        If SystemOptions.UserInterface = ArabicInterface Then
            StrFileName = App.path & "\REPORTS\REPORTS NEW\" & "ArabSalesInvoiceCompItemRec.rpt"
        Else
            StrFileName = App.path & "\REPORTS\REPORTS NEW\" & "ArabSalesInvoiceCompItemRec.rpt"
        End If
    
    
    End If
PrintFile:
  
    If Dir(StrFileName) = "" Then
        'GetMsgs 139, vbExclamation
        Screen.MousePointer = vbDefault
        Exit Function
    End If

    Set RsData = New ADODB.Recordset
        Cn.CommandTimeout = 10000
    RsData.Open NoteSerial, Cn, adOpenStatic, adLockReadOnly, adCmdText

    If RsData.BOF Or RsData.EOF Then
        If SystemOptions.UserInterface = ArabicInterface Then
        Msg = "·« ÊÃœ »Ì«‰«  ··⁄—÷"
        Else
        Msg = "No Data"
        End If
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
        StrReportTitle = "" '& StrAccountName
    Else
        xReport.ParameterFields(1).AddCurrentValue cCompanyInfo.ArabCompanyName ' RPTCompany_Name_Eng
        StrReportTitle = ""
    End If
    
     Dim i As Long
     Dim mTitle As String
     Dim mTitleNo As String
     Dim mTitle2 As String
     Select Case Ind
     Case 0
        
        mTitle = "›« Ê—… «·„»Ì⁄« "
        mTitleNo = "—ﬁ„ «·›« Ê—…"
        mTitle2 = "Invoice"
     Case 1
        mTitle = "⁄—÷ «·«”⁄«—"
        mTitleNo = "—ﬁ„ «·⁄—÷"
        mTitle2 = "Quotaion"
     Case 2
        mTitle = "”‰œ «·«” ·«„ ·⁄„Ì·"
        mTitleNo = "—ﬁ„ ”‰œ «·«” ·«„"
        mTitle2 = "Recive Voucher"
        
     End Select
    For i = 1 To xReport.FormulaFields.count
        Select Case xReport.FormulaFields.Item(i).Name
        Case "{@Title}"
            
            xReport.FormulaFields.Item(i).Text = "'" & mTitle & "'"
 Case "{@Title2}"
            
            xReport.FormulaFields.Item(i).Text = "'" & mTitle2 & "'"
        Case "{@TitleOrderNo}"
            
            xReport.FormulaFields.Item(i).Text = "'" & mTitleNo & "'"
        Case "{@OrderNo}"
            If Ind = 0 And Trim(TxtNoteSerial13) <> "" Then
                xReport.FormulaFields.Item(i).Text = "'" & TxtNoteSerial13 & "'"
               ElseIf Ind = 2 And Trim(TxtNoteSerial12) <> "" Then
                xReport.FormulaFields.Item(i).Text = "'" & TxtNoteSerial12 & "'"
            End If
        Case "{@HideSection}"
            
            xReport.FormulaFields.Item(i).Text = IIf(chkHiddLogo.value = vbChecked, True, False)
        
        End Select
    Next i
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

Function print_report3(Optional NoteSerial As String, Optional Ind As Integer = 0)
     
   
     
    'Dim MySQL As String
    Dim RsData As New ADODB.Recordset
    Dim xApp As New CRAXDRT.Application
    Dim xReport As CRAXDRT.Report
    Dim CViewer As ClsReportViewer
    Dim StrReportTitle As String
    Dim StrFileName As String
    Dim Msg As String
    
     Set RsData = New ADODB.Recordset
    RsData.Open NoteSerial, Cn, adOpenStatic, adLockReadOnly, adCmdText
  


        
    StrFileName = App.path & "\REPORTS\REPORTS NEW\ItemsRows.rpt"
 
  
  
    If Dir(StrFileName) = "" Then
        'GetMsgs 139, vbExclamation
        Screen.MousePointer = vbDefault
        Exit Function
    End If

    Set RsData = New ADODB.Recordset
        Cn.CommandTimeout = 10000
    RsData.Open NoteSerial, Cn, adOpenStatic, adLockReadOnly, adCmdText

    If RsData.BOF Or RsData.EOF Then
        If SystemOptions.UserInterface = ArabicInterface Then
        Msg = "·« ÊÃœ »Ì«‰«  ··⁄—÷"
        Else
        Msg = "No Data"
        End If
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
        StrReportTitle = "" '& StrAccountName
    Else
        xReport.ParameterFields(1).AddCurrentValue cCompanyInfo.ArabCompanyName ' RPTCompany_Name_Eng
        StrReportTitle = ""
    End If
    
     Dim i As Long
     Dim mTitle As String
     Dim mTitleNo As String
     Dim mTitle2 As String
     Select Case Ind
     Case 0
        mTitle = "«·„Ê«œ «·Œ«„ «·„ﬁœ—…"
        mTitleNo = "—ﬁ„ «·›« Ê—…"
        mTitle2 = "Invoice"
     Case 1
        mTitle = "⁄—÷ «·«”⁄«—"
        mTitleNo = "—ﬁ„ «·⁄—÷"
        mTitle2 = "Quotaion"
     Case 2
        mTitle = "”‰œ «·«” ·«„ ·⁄„Ì·"
        mTitleNo = "—ﬁ„ ”‰œ «·«” ·«„"
        mTitle2 = "Recive Voucher"
        
     End Select
    For i = 1 To xReport.FormulaFields.count
        Select Case xReport.FormulaFields.Item(i).Name
        Case "{@Title}"
            
            xReport.FormulaFields.Item(i).Text = "'" & mTitle & "'"
 Case "{@Title2}"
            
            xReport.FormulaFields.Item(i).Text = "'" & mTitle2 & "'"
        Case "{@TitleOrderNo}"
            
            xReport.FormulaFields.Item(i).Text = "'" & mTitleNo & "'"
        Case "{@OrderNo}"
            If Ind = 0 And Trim(TxtNoteSerial13) <> "" Then
                xReport.FormulaFields.Item(i).Text = "'" & TxtNoteSerial13 & "'"
               ElseIf Ind = 2 And Trim(TxtNoteSerial12) <> "" Then
                xReport.FormulaFields.Item(i).Text = "'" & TxtNoteSerial12 & "'"
            End If
        
        End Select
    Next i
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


Private Sub TxtItemCode_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        If txtItemCode.Text = "" Then
            Me.DcboItemID3.BoundText = ""
        Else
            Me.DcboItemID3.BoundText = GetItemID(Trim$(Me.txtItemCode.Text))
        End If
    End If

End Sub

Private Sub txtPrice_Validate(Cancel As Boolean)
If Me.TxtModFlg.Text <> "R" Then
    txtTotal = val(TxtPrice) * val(txtQty1)
    'ReLineGrid
End If
    If Not SystemOptions.IsMultiItemsInCompItem Then
        If FG2.Rows > 1 Then
            FG2.TextMatrix(1, FG2.ColIndex("Qty")) = txtQty1
            FG2.TextMatrix(1, FG2.ColIndex("Price")) = TxtPrice
            FG2.TextMatrix(1, FG2.ColIndex("widtj")) = txtwidtj
            FG2.TextMatrix(1, FG2.ColIndex("hight")) = txthight
            FG2.TextMatrix(1, FG2.ColIndex("Length")) = txtLength
            
        End If
    End If
    CalcTotalNet
End Sub

Private Sub txtTotalAdd_Change()
'CalcTotalNet
txtNet = val(txtTotal) + val(txtTotalAdd) - val(txtTotalDisc)
CalCulteVAT 3

End Sub



Private Sub XPCboGroup2_Change()
    Dim Dcombos As New ClsDataCombos
    Dim mIndex As Integer
    If Trim(XPCboGroup2.BoundText) <> "" Then
        mIndex = myRound(XPCboGroup2.BoundText)
        Dcombos.GetItemsNamesupdate Me.DcboItemID2, , , , , mIndex
        'Dcombos.GetItemsNamesupdate Me.DcboItemID2, , , , , mIndex
    Else
        Dcombos.GetItemsNamesupdate Me.DcboItemID2
        'Dcombos.GetItemsNamesupdate Me.DcboItemID2
    End If

End Sub


Private Function myRound(ByVal mNumber As Variant, _
                        Optional NoOfDecimalDigits As Integer) As Double
    Dim X As Double

    If IsNumeric(Trim(mNumber)) Then X = CDbl(Trim(mNumber)) Else X = val(Trim(mNumber))
    '-------------------------
    If X = 0 Then myRound = 0 Else myRound = Round(X + 1E-17, IIf(NoOfDecimalDigits = 0, 2, NoOfDecimalDigits))
End Function

Private Sub XPCboGroup_Click(Area As Integer)
    On Error Resume Next
    Dim OverHead As Double
    
     GetGroupData val(XPCboGroup.BoundText), , , , , "groups", , , OverHead
 
End Sub

Private Sub XPCboGroup_Change()
    Dim Dcombos As New ClsDataCombos
    Dim mIndex As Integer
    If Trim(XPCboGroup.BoundText) <> "" Then
        mIndex = myRound(XPCboGroup.BoundText)
        Dcombos.GetItemsNamesupdate Me.DcboItemID1, , , , , mIndex
        'Dcombos.GetItemsNamesupdate Me.DcboItemID2, , , , , mIndex
    Else
        Dcombos.GetItemsNamesupdate Me.DcboItemID1
        'Dcombos.GetItemsNamesupdate Me.DcboItemID2
    End If

End Sub
Private Sub XPCboGroup_KeyUp(KeyCode As Integer, _
                             Shift As Integer)

    If KeyCode = vbKeyF5 Then
        Dim Dcombos As ClsDataCombos
        
        Set Dcombos = New ClsDataCombos
        Dcombos.GetItemSGroups Me.XPCboGroup, False
        
    End If

End Sub


Private Sub XPCboGroup2_Click(Area As Integer)
    On Error Resume Next
    Dim OverHead As Double
    
     GetGroupData val(XPCboGroup2.BoundText), , , , , "groups", , , OverHead
 
End Sub

Private Sub XPCboGroup5_Change()
    Dim Dcombos As New ClsDataCombos
    Dim mIndex As Integer
    If Trim(XPCboGroup5.BoundText) <> "" Then
        mIndex = myRound(XPCboGroup5.BoundText)
        Dcombos.GetItemsNamesupdate Me.DcboItemID5, , , , , mIndex
        'Dcombos.GetItemsNamesupdate Me.DcboItemID2, , , , , mIndex
    Else
        Dcombos.GetItemsNamesupdate Me.DcboItemID5
        'Dcombos.GetItemsNamesupdate Me.DcboItemID2
    End If
End Sub

Private Sub XPCboGroup5_Click(Area As Integer)
    On Error Resume Next
    Dim OverHead As Double
    
     GetGroupData val(XPCboGroup5.BoundText), , , , , "groups", , , OverHead
 
End Sub
Private Sub XPCboGroup2_KeyUp(KeyCode As Integer, _
                             Shift As Integer)

    If KeyCode = vbKeyF5 Then
        Dim Dcombos As ClsDataCombos
        
        Set Dcombos = New ClsDataCombos
        Dcombos.GetItemSGroups Me.XPCboGroup2, False
        
    End If

End Sub

Private Sub XPCboGroup5_KeyUp(KeyCode As Integer, _
                             Shift As Integer)

    If KeyCode = vbKeyF5 Then
        Dim Dcombos As ClsDataCombos
        
        Set Dcombos = New ClsDataCombos
        Dcombos.GetItemSGroups Me.XPCboGroup5, False
        
    End If

End Sub





Private Sub XPCboGroupBuiltin_Change()
    Dim Dcombos As New ClsDataCombos
    Dim mIndex As Integer
    If Trim(XPCboGroupBuiltin.BoundText) <> "" Then
        mIndex = myRound(XPCboGroupBuiltin.BoundText)
        Dcombos.GetItemsNamesupdate Me.DcboBuiltinItemID, , , , , mIndex
        'Dcombos.GetItemsNamesupdate Me.DcboItemID2, , , , , mIndex
    Else
        Dcombos.GetItemsNamesupdate Me.DcboBuiltinItemID
        'Dcombos.GetItemsNamesupdate Me.DcboItemID2
    End If

End Sub
Private Sub XPCboGroupBuiltin_KeyUp(KeyCode As Integer, _
                             Shift As Integer)

    If KeyCode = vbKeyF5 Then
        Dim Dcombos As ClsDataCombos
        
        Set Dcombos = New ClsDataCombos
        Dcombos.GetItemSGroups Me.XPCboGroupBuiltin, False
        
    End If

End Sub





 
Function CREATE_VOUCHER_GE(Transaction_ID As Long, TxtNoteSerialV As String, TxtNoteSerial1V As String, general_noteid As Long, BranchID As Integer, StoreId As Double, Transaction_Date As Date, BoxID As Double)
'Exit Function
Dim LngDevID As Long
Dim LngDevNO As Integer
Dim BillTOTAL  As Double
 Dim StrTempDes As String
    LngDevID = new_id("DOUBLE_ENTREY_VOUCHERS", "Double_Entry_Vouchers_ID", "")
    '«·ÿ—› «·œ«∆‰
     
    my_branch = BranchID
LngDevNO = 1
    StrTempDes = "”‰œ «” ·«„ »‰«¡« ⁄·Ï ”‰œ  Ã„Ì⁄ —ﬁ„" & TxtTransSerial

 
Account_Code_dynamic = get_account_code_branch(37, val(dcBranch.BoundText))
  ' StrTempAccountCode = GetMyAccountCode("TblBoxesData", "BoxID", CLng(BoxID))  '????????
  StrTempAccountCode = get_account_code_branch(0, val(dcBranch.BoundText))
   StrTempAccountCode = get_store_Account(CInt(DCboStore3Name.BoundText), "Account_Code")
Dim mCostTotal As Double
If FG2.Rows = 1 Then Exit Function
mCostTotal = FG2.Aggregate(flexSTSum, FG2.FixedRows, FG2.ColIndex("TotalCost"), FG2.Rows - 1, FG2.ColIndex("TotalCost"))
BillTOTAL = mCostTotal
If BillTOTAL > 0 Then
            If ModAccounts.AddNewDev(LngDevID, LngDevNO, StrTempAccountCode, BillTOTAL, 0, StrTempDes, general_noteid, , , , Transaction_Date, val(user_id), Transaction_ID, , , , , , , , , , , , , , , , , BranchID) = False Then
                GoTo ErrTrap
            End If
        LngDevNO = LngDevNO + 1
        
                    If ModAccounts.AddNewDev(LngDevID, LngDevNO, Account_Code_dynamic, BillTOTAL, 1, StrTempDes, general_noteid, , , , Transaction_Date, val(user_id), Transaction_ID, , , , , , , , , , , , , , , , , BranchID) = False Then
                GoTo ErrTrap
            End If
            
End If

     
   ' Dim StrSQL  As String
   ' StrSQL = "UPDATE Transactions SET NOTS=" & Transaction_ID & " WHERE Transaction_ID=" & Transaction_ID
   ' Cn.Execute StrSQL
ErrTrap:
End Function

Function CREATE_VOUCHER_GE1(Transaction_ID As Long, TxtNoteSerialV As String, TxtNoteSerial1V As String, general_noteid As Long, BranchID As Integer, StoreId As Double, Transaction_Date As Date, BoxID As Double, Optional invoice As Integer = 0)
    Dim LngDevID As Long
    Dim LngDevNO  As Integer
    Dim StrTempAccountCode As String
    Dim StrTempDes As String
    Dim Line1 As Double
    Dim Line2 As Double
    Dim SngTemp As Double
    Dim OtherInformation As New ClsGLOther
    Dim DebitAccount  As String
    Dim TxtBillComment As String
    Dim CreditAccount  As String
    Dim i As Long
    LngDevID = new_id("DOUBLE_ENTREY_VOUCHERS", "Double_Entry_Vouchers_ID", "")
    '----------------
    Dim Account_Code_dynamic As String
    'SngTemp = NewGrid.GetItemsCostTotal * RSTransDetails("quantity").value / Cnt
    SngTemp = CostTOTAL
 
    If SngTemp > 0 Then
        '1 work with branch
        '2 work with inventory
        '3 work with groups
OtherInformation.NextAccount_Code = get_store_Account(val(StoreId), "Account_Code")
        If detect_inventory_work_type = 1 Then
            Account_Code_dynamic = get_account_code_branch(1, val(dcBranch.BoundText))
        
            If Account_Code_dynamic = "NO branch" Then
                MsgBox "·„ Ì „ «‰‘«¡ «·›—⁄", vbCritical
                GoTo ErrTrap
            Else

                If Account_Code_dynamic = "NO account" Then
                    MsgBox "·„ Ì „  ÕœÌœ Õ”«»   ﬂ·›… «·„»Ì⁄«  ›Ì «·›—⁄ ·Â–… «·⁄„·Ì…", vbCritical
                    GoTo ErrTrap
         
                End If
            End If

            Dim UseCustomerAcc As Integer

    
                StrTempAccountCode = Account_Code_dynamic ' ﬂ·›… «·„»Ì⁄«  1
   

            DebitAccount = StrTempAccountCode
    
            'StrTempAccountCode = "a3a2" ' ﬂ·›… «·„»Ì⁄« 
            If SystemOptions.UserInterface = ArabicInterface Then
                StrTempDes = "  √–‰ ’—›  —ﬁ„     " & Me.TxtNoteSerial1.Text & "  " & TxtBillComment & " „‰ ”‰œ  Ã„Ì⁄ —ﬁ„ " & TxtTransSerial
            Else
                StrTempDes = "Issue Voucher No.  " & Me.TxtNoteSerial1.Text & "  " & TxtBillComment & " „‰ ”‰œ  Ã„Ì⁄ —ﬁ„ " & TxtTransSerial
            End If

            Line1 = setfoxy_Line
            LngDevNO = LngDevNO + 1

            If ModAccounts.AddNewDev(LngDevID, LngDevNO, StrTempAccountCode, SngTemp, 0, StrTempDes, general_noteid, , , , Me.XPDtbBill.value, Me.DCboUserName.BoundText, Transaction_ID, , , , , , , , Line1, , , , , , , , , val(Me.dcBranch.BoundText), , , , , , , , , , , , , , , , , , , , , , , , , , , OtherInformation) = False Then
                GoTo ErrTrap
            End If
    
    
    
            '«·„Œ“Ê‰ ›Ì «·›—⁄
            Account_Code_dynamic = get_account_code_branch(0, val(dcBranch.BoundText))
        
            If Account_Code_dynamic = "NO branch" Then
                If SystemOptions.UserInterface = ArabicInterface Then
                    MsgBox "·„ Ì „ «‰‘«¡ «·›—⁄", vbCritical
                Else
                    MsgBox "The branch was not created", vbCritical
                End If
                GoTo ErrTrap
            Else

                If Account_Code_dynamic = "NO account" Then
                     If SystemOptions.UserInterface = ArabicInterface Then
                    MsgBox "·„ Ì „  ÕœÌœ Õ”«»   ﬂ·›… «·„Œ“Ê‰ ›Ì «·›—⁄ ·Â–… «·⁄„·Ì…", vbCritical
                Else
                                    MsgBox "The inventory cost calculation in the branch is not specified for this process", vbCritical
                End If
                    GoTo ErrTrap
         
                End If
            End If
        
           
                StrTempAccountCode = Account_Code_dynamic '«·„Œ“Ê‰ 0 ›Ì «·›—⁄
          

            CreditAccount = StrTempAccountCode
    
            If SystemOptions.UserInterface = ArabicInterface Then
                StrTempDes = "√–‰ ’—›  —ﬁ„ " & Me.TxtNoteSerial1.Text & "  " & TxtBillComment & " „‰ ”‰œ  Ã„Ì⁄ —ﬁ„ " & TxtTransSerial
            Else
                StrTempDes = "Issue Voucher No. " & Me.TxtNoteSerial1.Text & "  " & TxtBillComment & " „‰ ”‰œ  Ã„Ì⁄ —ﬁ„ " & TxtTransSerial
            End If
    
            LngDevNO = LngDevNO + 1
            Line2 = setfoxy_Line

            If ModAccounts.AddNewDev(LngDevID, LngDevNO, StrTempAccountCode, SngTemp, 1, StrTempDes, general_noteid, , , , Me.XPDtbBill.value, Me.DCboUserName.BoundText, Transaction_ID, , , , , , , , Line2, , , , , , , , , val(Me.dcBranch.BoundText), , , , , , , , , , , , , , , , , , , , , , , , , , , OtherInformation) = False Then
                GoTo ErrTrap
            End If
    
        ElseIf detect_inventory_work_type = 2 Then
            
     'salimhere
     If invoice = 0 Then '« «Ã
     Account_Code_dynamic = get_account_code_branch(37, CInt(BranchID))
        Else
        
        Account_Code_dynamic = get_account_code_branch(1, val(dcBranch.BoundText))  '„»Ì⁄« 
        End If
            If Account_Code_dynamic = "NO branch" Then
                If SystemOptions.UserInterface = ArabicInterface Then
                    MsgBox "·„ Ì „ «‰‘«¡ «·›—⁄", vbCritical
                Else
                    MsgBox "The branch was not created", vbCritical
                End If
                GoTo ErrTrap
            Else

                If Account_Code_dynamic = "NO account" Then
                    If SystemOptions.UserInterface = ArabicInterface Then
                        MsgBox "·„ Ì „  ÕœÌœ Õ”«»   ﬂ·›… «·«‰ «Ã ›Ì «·›—⁄ ·Â–… «·⁄„·Ì…", vbCritical
                    Else
                         MsgBox "The production cost calculation is not determined in the section for this process", vbCritical
                    End If
                    GoTo ErrTrap
         
                End If
            End If

           
            StrTempAccountCode = Account_Code_dynamic ' ﬂ·›… «·„»Ì⁄«  1
          
            DebitAccount = StrTempAccountCode
            
            Line1 = setfoxy_Line

            'StrTempAccountCode = "a3a2" ' ﬂ·›… «·„»Ì⁄« 
            If SystemOptions.UserInterface = ArabicInterface Then
                StrTempDes = "√–‰ ’—›  —ﬁ„ " & Me.TxtNoteSerial11.Text & " „‰ ”‰œ  Ã„Ì⁄ —ﬁ„ " & TxtTransSerial
            Else
                StrTempDes = "Issue Voucher No. " & Me.TxtNoteSerial1.Text & " „‰ ”‰œ  Ã„Ì⁄ —ﬁ„ " & TxtTransSerial
            End If
    
            LngDevNO = LngDevNO + 1
       Dim project_id As Integer
'        project_id = IIf(Me.DcbProject.BoundText = "", 0, Me.DcbProject.BoundText)
             If ModAccounts.AddNewDev(LngDevID, LngDevNO, StrTempAccountCode, SngTemp, 0, StrTempDes, general_noteid, , , , Me.XPDtbBill.value, Me.DCboUserName.BoundText, Transaction_ID, , , , , , , , Line1, , , , , , , , , val(Me.dcBranch.BoundText), , , , , , , , , , , , , , , , , , , , , , , , , , , OtherInformation) = False Then
                GoTo ErrTrap
            End If

            '«·„Œ“Ê‰ «·”·⁄Ì ⁄·Ï „” ÊÏ «·„Œ“‰
            SngTemp = CostTOTAL

            
            Account_Code_dynamic = get_store_Account(val(StoreId), "Account_Code")
            
        
            If Account_Code_dynamic = "" Then
                If SystemOptions.UserInterface = ArabicInterface Then
                    MsgBox "·„ Ì „  ÕœÌœ Õ”«»  ··„Œ“Ê‰ «·”·⁄Ì ·Â–« «·„Œ“‰ ›Ì Â–« «·›—⁄    ", vbCritical
                Else
                    MsgBox "No inventory account for this store has been specified in this section  ", vbCritical
                End If
                
                GoTo ErrTrap
            End If
    
            StrTempAccountCode = Account_Code_dynamic  '„Õ“Ê‰ «·”·⁄Ì ··„Œ“‰
            CreditAccount = StrTempAccountCode
OtherInformation.NextAccount_Code = DebitAccount
            ' StrTempAccountCode = "a1a2a5" '„Õ“Ê‰ «·»÷«⁄…
            If SystemOptions.UserInterface = ArabicInterface Then
                StrTempDes = "√–‰ ’—›  —ﬁ„ " & Me.TxtNoteSerial1.Text & "  " & TxtBillComment & " „‰ ”‰œ  Ã„Ì⁄ —ﬁ„ " & TxtTransSerial
            Else
                StrTempDes = "Issue Voucher No. " & Me.TxtNoteSerial1.Text & "  " & TxtBillComment & " „‰ ”‰œ  Ã„Ì⁄ —ﬁ„ " & TxtTransSerial
            End If

            Line2 = setfoxy_Line
         
            LngDevNO = LngDevNO + 1

            If ModAccounts.AddNewDev(LngDevID, LngDevNO, StrTempAccountCode, SngTemp, 1, StrTempDes, general_noteid, , , , Me.XPDtbBill.value, Me.DCboUserName.BoundText, Transaction_ID, , , , , , , , Line2, , , , , , , , , val(Me.dcBranch.BoundText), , , , , , , , , , , , , , , , , , , , , , , , , , , OtherInformation) = False Then
                GoTo ErrTrap
            End If

        ElseIf detect_inventory_work_type = 3 Then
            Dim groupAccount As String
             
            Dim line_value As Single

            With FG

                For i = 1 To FG.Rows - 1

                    If FG.TextMatrix(i, FG.ColIndex("itemcode")) <> "" Then
                        If FG.RowHidden(i) Or CBool(FG.ValueMatrix(i, FG.ColIndex("IsDeleted"))) = True Then GoTo NextRow2
                        ' groupAccount = get_item_group_account(FG.TextMatrix(i, FG.ColIndex("Code")), DCboStoreName.BoundText, 2)
                        groupAccount = get_item_group_account_in_branch(FG.TextMatrix(i, FG.ColIndex("itemcode")), val(val(dcBranch.BoundText)), 1)

                        If groupAccount = "Error" Then
                            If SystemOptions.UserInterface = ArabicInterface Then
                                MsgBox "«·’‰› ›Ì «·”ÿ— —ﬁ„" & i & " €Ì— „Õœœ —ﬁ„ Õ”«»   ﬂ·›… ·„Ã„Ê⁄ …"
                            Else
                                MsgBox "Item in line no " & i & "Group Name Account Not Defined"
                            End If

                            GoTo ErrTrap
                        End If

                        line_value = FG.TextMatrix(i, FG.ColIndex("Price")) * FG.TextMatrix(i, FG.ColIndex("Qty"))
    
                        If SystemOptions.UserInterface = ArabicInterface Then
                            StrTempDes = "√–‰ ’—›  —ﬁ„ " & Me.TxtNoteSerial1.Text
                        Else
                            StrTempDes = "Issue Voucher No. " & Me.TxtNoteSerial1.Text
                        End If
    
                        LngDevNO = LngDevNO + 1

                        If ModAccounts.AddNewDev(LngDevID, LngDevNO, groupAccount, line_value, 0, StrTempDes, general_noteid, , , , Me.XPDtbBill.value, Me.DCboUserName.BoundText, Transaction_ID, , , , , , , , , , , , , , , , , val(Me.dcBranch.BoundText)) = False Then
                            GoTo ErrTrap
                        End If
    
                    End If
NextRow2:
                Next i

            End With

            With FG

                For i = 1 To FG.Rows - 1

                    If FG.TextMatrix(i, FG.ColIndex("itemcode")) <> "" Then
                        If FG.RowHidden(i) Or CBool(FG.ValueMatrix(i, FG.ColIndex("IsDeleted"))) = True Then GoTo NextRow
                        ' groupAccount = get_item_group_account(FG.TextMatrix(i, FG.ColIndex("Code")), DCboStoreName.BoundText, 2)
                        groupAccount = get_item_group_account_inventory(FG.TextMatrix(i, FG.ColIndex("itemcode")), DCboStore2Name.BoundText, 0)

                        If groupAccount = "Error" Then
                            If SystemOptions.UserInterface = ArabicInterface Then
                                MsgBox "«·’‰› ›Ì «·”ÿ— —ﬁ„" & i & " €Ì— „Õœœ —ﬁ„ Õ”«»  «·„Œ“Ê‰ «·”·⁄Ì ··„Œ“‰ «·„Õœœ   ·„Ã„Ê⁄ …"
                            Else
                                MsgBox "Item in line no " & i & "Group Name Account Not Defined"
                            End If

                            GoTo ErrTrap
                        End If

                        line_value = FG.TextMatrix(i, FG.ColIndex("Price")) * FG.TextMatrix(i, FG.ColIndex("Qty"))
    
                        If SystemOptions.UserInterface = ArabicInterface Then
                            StrTempDes = "√–‰ ’—›  —ﬁ„ " & Me.TxtNoteSerial1.Text & "  " & TxtBillComment & " „‰ ”‰œ  Ã„Ì⁄ —ﬁ„ " & TxtTransSerial
                        Else
                            StrTempDes = "Issue Voucher No. " & Me.TxtNoteSerial1.Text & "  " & TxtBillComment & " „‰ ”‰œ  Ã„Ì⁄ —ﬁ„ " & TxtTransSerial
                        End If

                        LngDevNO = LngDevNO + 1

                        If ModAccounts.AddNewDev(LngDevID, LngDevNO, groupAccount, line_value, 1, StrTempDes, general_noteid, , , , Me.XPDtbBill.value, Me.DCboUserName.BoundText, Transaction_ID, , , , , , , , , , , , , , , , , val(Me.dcBranch.BoundText), , , , , , , , , , , , , , , , , , , , , , , , , , , OtherInformation) = False Then
                            GoTo ErrTrap
                        End If
    
                    End If
NextRow:
                Next i

            End With

        End If

        '----------------
        'LngDevID = LngDevID + 1
        'LngDevNO = 0
    End If
   ' ute StrSQL
ErrTrap:
End Function



Private Sub createVoucher1(BranchID As Double, _
BoxID As Double, _
Transaction_Date As Date, _
Transaction_Type As Double, _
CBoBasedON As Double, _
UserID As Double, _
Trans_DiscountType As Double, _
CusID As Double, _
StoreId As Double, _
PaymentType As Double, _
Emp_id As Double, _
TransactionComment As String)
Dim sql As String
Dim Msg As String
Dim NoteID As Long
Dim Transaction_ID As Long
Dim Transaction_ID1 As Long
Dim Transaction_serial As String
Dim NoteSerial As String
Dim NoteSerial1 As String
Dim StrSQL As String
Dim s As String
Dim mSaveAgin  As Boolean
Dim mCostTotal  As Double
Dim mItemNo As Long
 Dim RSNoteID As New ADODB.Recordset
 Dim rsDummyItem As New ADODB.Recordset
 Dim costPrice As Double
'BillTOTAL = 0
'CostTOTAL = 0
'Check
  'NoteSerial1 = Voucher_coding(val(BranchID), Transaction_Date, 10, 180, , 27)
   If Not IsSaveWithOutMsg Then
SaveAgin:
    NoteSerial1 = Voucher_coding(val(BranchID), Transaction_Date, 19, 250, , 28) '’—› «” ·«„  Œ«„
        If NoteSerial1 = "" Then
                 If NoteSerial1 = "error" Then
                     MsgBox " ·« Ì„ﬂ‰ «÷«›… ”‰œ   „Ê«œ Œ«„ ··«‰ «Ã  ÃœÌœ ·«‰ﬂ  ⁄œÌ  «·Õœ «·–Ì ﬁ„  » ÕœÌœ… „‰ «·”‰œ«   ": Exit Sub
                 ElseIf NoteSerial1 = "" Then
                         MsgBox " ·«»œ „‰ ﬂ «»… —ﬁ„ «·”‰œ ÌœÊÌ« ﬂ„« Õœœ   ": Exit Sub
        
                 End If
        End If

 
NoteSerial = Notes_coding(val(BranchID), Transaction_Date)
 If NoteSerial = "" Then
            If NoteSerial = "error" Then
                MsgBox " ·« Ì„ﬂ‰ «÷«›… ﬁÌÊœ ·Â–… «·⁄„·Ì… ·«‰ﬂ  ⁄œÌ  «·Õœ «·«ﬁ’Ì ··ﬁÌÊœ ﬂ„« Õœœ  ›Ì  —ﬁÌ„ «·”‰œ«  ": Exit Sub
            ElseIf NoteSerial = "" Then
                    MsgBox " ·«Ì„ﬂ‰ «‰‘«¡ «·ﬁÌœ ·Â–« «·„” ‰œ ·«‰ﬂ Õœœ   —ﬁÌ„ ﬁÌÊœ ÌœÊÌ  ": Exit Sub
                 
            End If
End If
           
 
              
  
  
 
           CostAccount = get_account_code_branch(37, CInt(BranchID))
        
            If CostAccount = "NO branch" Or CostAccount = "NO account" Then
                If SystemOptions.UserInterface = ArabicInterface Then
                    MsgBox "·„ Ì „ —»ÿ  ﬂ·›… «·«‰ «Ã „Ê«œ  ", vbCritical
                Else
                    MsgBox "Sales Not Created", vbCritical
                End If

             Exit Sub
              End If
              
              

    StoreAccount = get_store_Account(CInt(StoreId), "Account_Code")
      If StoreAccount = "" Then
                If SystemOptions.UserInterface = ArabicInterface Then
                    MsgBox "·„ Ì „  ÕœÌœ Õ”«»  ··„Œ“Ê‰ «·”·⁄Ì ·Â–« «·„Œ“‰ ›Ì Â–« «·›—⁄    ", vbCritical
                Else
                    MsgBox "No inventory account for this store has been specified in this section   ", vbCritical
                End If
           Exit Sub
            End If


 'end Check
        If TxtNoteSerial12 = "" Then
 NoteSerial1 = Voucher_coding(val(BranchID), Transaction_Date, 19, 250, , 28)
 End If
Transaction_ID = CStr(new_id("Transactions", "Transaction_ID", "", True))
 
        TXTTransactionID2.Text = Transaction_ID
               TxtNoteSerial12.Text = NoteSerial1
Dim mCust As Long
Dim rsDummyChkCust As New ADODB.Recordset
sql = "Select * from TblCustemers Where CusId = " & CusID

rsDummyChkCust.Open sql, Cn, adOpenStatic, adLockReadOnly
If rsDummyChkCust.EOF Then
    sql = "Select Top 1 CusId from TblCustemers "
    rsDummyChkCust.Close
    rsDummyChkCust.Open sql, Cn, adOpenStatic, adLockReadOnly
    CusID = val(rsDummyChkCust!CusID & "")
End If
               
 sql = "INSERT INTO  Transactions (  "
sql = sql & " Transaction_ID ,"
sql = sql & " BranchID ,"
sql = sql & " NoteSerial ,"
sql = sql & " NoteSerial1 ,"
sql = sql & " boxId ,"
sql = sql & " Transaction_serial ,"
sql = sql & " Transaction_Date ,"
sql = sql & " Transaction_Type ,"
sql = sql & " BillBasedOn ,"
sql = sql & " UserID ,"
sql = sql & " Trans_DiscountType ,"
sql = sql & " CusID ,"
sql = sql & " StoreId ,"
sql = sql & " PaymentType ,"
sql = sql & " Emp_id ,InvoiceOrderNo,"
 sql = sql & " TransactionComment )"
 
 sql = sql & " VALUES("
sql = sql & " " & Transaction_ID & " ,"
sql = sql & " " & BranchID & " ,"
sql = sql & "'" & NoteSerial & "' ,"
sql = sql & "'" & NoteSerial1 & "' ,"
sql = sql & " " & BoxID & " ,"
sql = sql & "'" & Transaction_serial & "',"
sql = sql & " " & SQLDate(Transaction_Date, True) & " ,"
sql = sql & " " & Transaction_Type & " ,"
sql = sql & " 0 ,"
sql = sql & " " & user_id & " ,"
sql = sql & " 0 ,"
sql = sql & " " & CusID & " ,"
sql = sql & " " & StoreId & " ,"
sql = sql & " 0 ,"
sql = sql & " " & Emp_id & " ," & val(TxtTransSerial) & ","
 sql = sql & "'" & TransactionComment & "')"
 

         Cn.Execute sql
Else
    Transaction_ID = val(TXTTransactionID2.Text)
    NoteSerial1 = TxtNoteSerial12.Text
    Cn.Execute "Delete Transaction_Details Where Transaction_ID = " & Transaction_ID
    
    sql = "Select NoteId from Transactions Where Transaction_ID = " & Transaction_ID
   
    RSNoteID.Open sql, Cn, adOpenKeyset, adLockOptimistic, adCmdText
    If Not RSNoteID.EOF Then
        NoteID = val(RSNoteID!NoteID)
    End If
    
    
        s = "SELECT * FROM Transactions AS t WHERE t.Transaction_ID =  " & Transaction_ID
        Dim rsTest2 As New ADODB.Recordset
        rsTest2.Open s, Cn, adOpenStatic, adLockReadOnly
        If rsTest2.EOF Then
            mSaveAgin = True
            GoTo SaveAgin
        End If
End If

If Transaction_ID = 0 Then Exit Sub
 
        Dim RSTransDetails As New ADODB.Recordset

StrSQL = "SELECT     dbo.Transaction_Details.* from dbo.Transaction_Details Where (Transaction_ID = -1)"
   RSTransDetails.Open StrSQL, Cn, adOpenKeyset, adLockOptimistic, adCmdText
   
   Dim RowNum As Integer
        For RowNum = 1 To FG2.Rows - 1
                
                          
                s = "Select ItemId from tblItems Where ItemId = " & val(FG2.TextMatrix(RowNum, FG2.ColIndex("ItemID"))) & " and ItemType <> 1"
                
                Set rsDummyItem = New ADODB.Recordset
                rsDummyItem.Open s, Cn, adOpenStatic, adLockReadOnly
                If rsDummyItem.EOF Then
                    GoTo NextRow
                End If

            'If DcboItemID1.BoundText <> "" Then
                RSTransDetails.AddNew
                RSTransDetails("Transaction_ID").value = Transaction_ID
             
                RSTransDetails("ColorID").value = 1
                RSTransDetails("ItemSize").value = 1
                RSTransDetails("ClassId").value = 1
                
                RSTransDetails("Item_ID").value = val(FG2.TextMatrix(RowNum, FG2.ColIndex("ItemID")))
                RSTransDetails("UnitID").value = val(FG2.TextMatrix(RowNum, FG2.ColIndex("UnitID")))
                RSTransDetails("SHOWQTY").value = val(FG2.TextMatrix(RowNum, FG2.ColIndex("Qty")))
'                If Not SystemOptions.IsMultiItemsInCompItem Then
'                    costPrice = ModItemCostPrice.GetCostItemPrice(CLng(val(DcboItemID1.BoundText)), 0, "", , SystemOptions.SysMainStockCostMethod, , , XPDtbBill, , val(DcbUnit.BoundText))
'                Else
                    costPrice = val(FG2.TextMatrix(RowNum, FG2.ColIndex("cost")))
                'End If
               ' CostTOTAL = costPrice * val(fg2.TextMatrix(RowNum, fg2.ColIndex("Qty")))
                
                If val(FG2.TextMatrix(RowNum, FG2.ColIndex("ItemID"))) = 810 Or val(FG2.TextMatrix(RowNum, FG2.ColIndex("ItemID"))) = 643 Then
                 
                mItemNo = mItemNo
            End If
            
             FG2.TextMatrix(RowNum, FG2.ColIndex("TotalCost")) = val(FG2.TextMatrix(RowNum, FG2.ColIndex("Cost"))) * val(FG2.TextMatrix(RowNum, FG2.ColIndex("Qty")))
               'RSTransDetails("showPrice").value = Round(costPrice / IIf(val(fg2.TextMatrix(RowNum, fg2.ColIndex("Qty"))) <> 0, val(fg2.TextMatrix(RowNum, fg2.ColIndex("Qty"))), 1), 3)
              
               
                          '«·ÊÕœ« 
           
            Dim RsUnitData As ADODB.Recordset
            Dim LngCurItemID As Long
            Dim LngUnitID As Long
            Dim DblQty As Double
        
            LngCurItemID = val(RSTransDetails("Item_ID").value & "")
            LngUnitID = val(RSTransDetails("UnitID").value & "")
            If LngUnitID = 0 Then
                GetDefaultItemUnit val(LngCurItemID), LngUnitID
            End If
            
            DblQty = val(RSTransDetails("SHOWQTY").value & "")
           
            RSTransDetails("ShowQty").value = DblQty
     RSTransDetails("showPrice").value = Round(costPrice, 3) '/ DblQty
            StrSQL = "Select * From TblItemsUnits Where ItemID=" & LngCurItemID
            StrSQL = StrSQL + " AND UnitID=" & LngUnitID
            Set RsUnitData = New ADODB.Recordset
            RsUnitData.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText


            If Not (rs.BOF Or rs.EOF) Then
                RSTransDetails("QtyBySmalltUnit").value = RsUnitData("UnitFactor").value
                RSTransDetails("Quantity").value = RSTransDetails("QtyBySmalltUnit").value * RSTransDetails("showqty").value
             '   RSTransDetails("OpeningSalesQty").value = RSTransDetails("Quantity").value
             '   RSTransDetails("OpeningSalesValue").value = IIf((Fg.TextMatrix(RowNum, Fg.ColIndex("Valu")) = ""), Null, val(Fg.TextMatrix(RowNum, Fg.ColIndex("Valu"))))
                If costPrice < 0.09 Then
                    RSTransDetails("Price").value = Round((costPrice / IIf(val(DblQty) <> 0, DblQty, 1)) / RSTransDetails("QtyBySmalltUnit").value, 9)
                Else
                    RSTransDetails("Price").value = Round((costPrice / IIf(val(DblQty) <> 0, DblQty, 1)) / RSTransDetails("QtyBySmalltUnit").value, 3)
                End If
            
            End If
         '   RSTransDetails("CostPrice").value = IIf((FG.TextMatrix(RowNum, FG.ColIndex("Price")) = ""), Null, val(FG.TextMatrix(RowNum, FG.ColIndex("Price"))))
            
              
               
           '    BillTOTAL = BillTOTAL + (RSTransDetails("Price").value * RSTransDetails("SHOWQTY").value)
          
                
           '     RSTransDetails("CostPrice").value = IIf((Fg.TextMatrix(RowNum, Fg.ColIndex("CostPrice")) = ""), Null, val(Fg.TextMatrix(RowNum, Fg.ColIndex("CostPrice"))))
               
        
     'RSTransDetails("SavedItemType").value = 0
            
                RSTransDetails.update
                             UpdateTransactionsCost CStr(Transaction_ID)
            'End If
NextRow:
       Next RowNum
'Exit Sub
 If Not IsSaveWithOutMsg Or mSaveAgin Then
    NoteSerial = Notes_coding(val(BranchID), Transaction_Date)
End If

'Cn.Execute "Delete Notes where NoteId = "

If Not IsSaveWithOutMsg Or mSaveAgin Then
CreateNotesEntry:
    CreateNotes NoteID, Transaction_Date, CInt(BranchID), 250, 0, NoteSerial, NoteSerial1, "Transactions", "Transaction_ID", Transaction_ID, " »‰«¡« ⁄·Ï ”‰œ  Ã„Ì⁄ —ﬁ„ " & TxtTransSerial, ToHijriDate(Transaction_Date)
    CREATE_VOUCHER_GE TXTTransactionID2, TxtNoteSerial12, "", NoteID, val(dcBranch.BoundText), Store_ID, Transaction_Date, 0
Else
    If FG2.Rows = 1 Then Exit Sub
    s = "Select * from DOUBLE_ENTREY_VOUCHERS where Notes_Id = " & NoteID
    Dim rsDummyNotes As New ADODB.Recordset
    rsDummyNotes.Open s, Cn, adOpenStatic, adLockReadOnly
    If rsDummyNotes.EOF Then
        GoTo CreateNotesEntry
    Else
        mCostTotal = FG2.Aggregate(flexSTSum, FG2.FixedRows, FG2.ColIndex("TotalCost"), FG2.Rows - 1, FG2.ColIndex("TotalCost"))
        Cn.Execute "Update Notes Set Note_Value = " & mCostTotal & " Where NoteId = " & NoteID
        Cn.Execute "Update DOUBLE_ENTREY_VOUCHERS  Set Value = " & mCostTotal & " Where Notes_ID = " & NoteID
     
    End If
    
End If



'***********************
         StrSQL = "UPDATE Transactions SET NOTS=" & val(TxtTransSerial) & ",Transaction_Type = " & Transaction_Type & "  WHERE Transaction_ID=" & Transaction_ID
         Cn.Execute StrSQL
'***********************

   
       
 
        'StrSQL = "UPDATE Transactions SET NOTS=" & Transaction_ID & " WHERE Transaction_ID=" & val(Me.XPTxtBillID.text)
        'Cn.Execute StrSQL
  'MsgBox " „   «·‰ﬁ·"
  
'******************************************************issueVoucher








     
 
    '
 
ErrTrap:

End Sub
Private Sub createVoucher(BranchID As Double, _
BoxID As Double, _
Transaction_Date As Date, _
Transaction_Type As Double, _
CBoBasedON As Double, _
UserID As Double, _
Trans_DiscountType As Double, _
CusID As Double, _
StoreId As Double, _
PaymentType As Double, _
Emp_id As Double, _
TransactionComment As String, Optional invoice As Integer = 0)
Dim sql As String
Dim Msg As String
Dim NoteID As Long
Dim Transaction_ID As Long
Dim Transaction_ID1 As Long
Dim Transaction_serial As String
Dim NoteSerial As String
Dim NoteSerial1 As String
Dim RSNoteID As New ADODB.Recordset
Dim mSaveAgin  As Boolean
Dim s As String
Dim StrSQL  As String
Dim RowNum  As Long
Dim costPrice  As Double
mSaveAgin = False
'BillTOTAL = 0
CostTOTAL = 0
'Check
  'NoteSerial1 = Voucher_coding(val(BranchID), Transaction_Date, 10, 180, , 27)
    If Not IsSaveWithOutMsg Then
SaveAgin:
            If Transaction_Type = 27 Then
            NoteSerial1 = Voucher_coding(val(BranchID), Transaction_Date, 18, 240, , CInt(Transaction_Type), , CDbl(StoreId))               '’—› „Ê«œ Œ«„
            Else
            NoteSerial1 = Voucher_coding(val(BranchID), Transaction_Date, 7, 170, , CInt(Transaction_Type))    '’—› „Ê«œ Œ«„
            End If
                
            If NoteSerial1 = "" Then
                 If NoteSerial1 = "error" Then
                    If SystemOptions.UserInterface = ArabicInterface Then
                        MsgBox " ·« Ì„ﬂ‰ «÷«›… ”‰œ   „Ê«œ Œ«„ ··«‰ «Ã  ÃœÌœ ·«‰ﬂ  ⁄œÌ  «·Õœ «·–Ì ﬁ„  » ÕœÌœ… „‰ «·”‰œ«   ": Exit Sub
                    Else
                        MsgBox " You can not add a raw material bond to a new production because you have exceeded the limit on which you have selected the bonds ": Exit Sub
                    End If
            
                 ElseIf NoteSerial1 = "" Then
                         MsgBox " ·«»œ „‰ ﬂ «»… —ﬁ„ «·”‰œ ÌœÊÌ« ﬂ„« Õœœ   ": Exit Sub
            
                 End If
            End If
            
            NoteSerial = Notes_coding(val(BranchID), Transaction_Date)
            If NoteSerial = "" Then
            If NoteSerial = "error" Then
                MsgBox " ·« Ì„ﬂ‰ «÷«›… ﬁÌÊœ ·Â–… «·⁄„·Ì… ·«‰ﬂ  ⁄œÌ  «·Õœ «·«ﬁ’Ì ··ﬁÌÊœ ﬂ„« Õœœ  ›Ì  —ﬁÌ„ «·”‰œ«  ": Exit Sub
            ElseIf NoteSerial = "" Then
                    MsgBox " ·«Ì„ﬂ‰ «‰‘«¡ «·ﬁÌœ ·Â–« «·„” ‰œ ·«‰ﬂ Õœœ   —ﬁÌ„ ﬁÌÊœ ÌœÊÌ  ": Exit Sub
                 
            End If
            End If
            
            
            If Trim(StoreId) = 0 Then
            MsgBox "·„ Ì „  ÕœÌœ Õ”«»  ··„Œ“Ê‰ «·”·⁄Ì ·Â–« «·„Œ“‰ ›Ì Â–« «·›—⁄    ", vbCritical
            End If
            
            
            
            'CostAccount = get_account_code_branch(137, CInt(BranchID))
            
            
            
            If Transaction_Type = 27 Then
            CostAccount = get_account_code_branch(37, CInt(BranchID))
            Else
            CostAccount = get_account_code_branch(1, CInt(BranchID))
            End If
                        
                        
            If CostAccount = "NO branch" Or CostAccount = "NO account" Then
                If SystemOptions.UserInterface = ArabicInterface Then
                    MsgBox "·„ Ì „ —»ÿ  ﬂ·›…   «·„»Ì⁄«   ", vbCritical
                Else
                    MsgBox "Sales Not Created", vbCritical
                End If
            
             Exit Sub
              End If
              
              
            
            StoreAccount = get_store_Account(CInt(StoreId), "Account_Code")
            If StoreAccount = "" Then
                If SystemOptions.UserInterface = ArabicInterface Then
                    MsgBox "·„ Ì „  ÕœÌœ Õ”«»  ··„Œ“Ê‰ «·”·⁄Ì ·Â–« «·„Œ“‰ ›Ì Â–« «·›—⁄    ", vbCritical
                Else
                    MsgBox "No inventory account for this store has been specified in this section", vbCritical
                End If
            Exit Sub
            End If
        End If
        
          Dim RsUnitData As ADODB.Recordset
            Dim LngCurItemID As Long
            Dim LngUnitID As Long
            Dim DblQty As Double

 'end Check

If Not IsSaveWithOutMsg Or mSaveAgin Then
    Transaction_ID = CStr(new_id("Transactions", "Transaction_ID", "", True))
    Transaction_serial = NoteSerial1
        If Transaction_Type <> 19 Then
            TXTTransactionID1.Text = Transaction_ID
            TxtNoteSerial11.Text = NoteSerial1
        Else
            TXTTransactionID5.Text = Transaction_ID
            TxtNoteSerial15.Text = NoteSerial1
    
        End If
            
    Dim mCust As Long
    Dim rsDummyChkCust As New ADODB.Recordset
    sql = "Select * from TblCustemers Where CusId = " & CusID
    
    rsDummyChkCust.Open sql, Cn, adOpenStatic, adLockReadOnly
    If rsDummyChkCust.EOF Then
        sql = "Select Top 1 CusId from TblCustemers "
        rsDummyChkCust.Close
        rsDummyChkCust.Open sql, Cn, adOpenStatic, adLockReadOnly
        CusID = val(rsDummyChkCust!CusID & "")
    End If
            
     sql = "INSERT INTO  Transactions (  "
    sql = sql & " Transaction_ID ,"
    sql = sql & " BranchID ,"
    sql = sql & " NoteSerial ,"
    sql = sql & " NoteSerial1 ,"
    sql = sql & " boxId ,"
    sql = sql & " Transaction_serial ,"
    sql = sql & " Transaction_Date ,"
    sql = sql & " Transaction_Type ,"
    sql = sql & " BillBasedOn ,"
    sql = sql & " UserID ,"
    sql = sql & " Trans_DiscountType ,"
    sql = sql & " CusID ,"
    sql = sql & " StoreId ,"
    sql = sql & " PaymentType ,"
    sql = sql & " Emp_id ,InvoiceOrderNo,"
     sql = sql & " TransactionComment )"
     
     sql = sql & " VALUES("
    sql = sql & " " & Transaction_ID & " ,"
    sql = sql & " " & BranchID & " ,"
    sql = sql & "'" & NoteSerial & "' ,"
    sql = sql & "'" & NoteSerial1 & "' ,"
    sql = sql & " " & BoxID & " ,"
    sql = sql & "'" & Transaction_serial & "',"
    sql = sql & " " & SQLDate(Transaction_Date, True) & " ,"
    sql = sql & " " & Transaction_Type & " ,"
    sql = sql & " 2 ,"
    sql = sql & " " & user_id & " ,"
    sql = sql & " 0 ,"
    sql = sql & " " & CusID & " ,"
    sql = sql & " " & StoreId & " ,"
    sql = sql & " 0 ,"
    sql = sql & " " & Emp_id & " ," & val(TxtTransSerial) & ","
     sql = sql & "'" & TransactionComment & "')"
     
    
             Cn.Execute sql
Else
        If Transaction_Type <> 19 Then
            Transaction_ID = val(TXTTransactionID1.Text)
            NoteSerial1 = TxtNoteSerial11.Text
        Else
            Transaction_ID = val(TXTTransactionID5.Text)
            NoteSerial1 = TxtNoteSerial15.Text
    
        End If
        
        s = "SELECT * FROM Transactions AS t WHERE t.Transaction_ID =  " & Transaction_ID
        Dim rsTest2 As New ADODB.Recordset
        rsTest2.Open s, Cn, adOpenStatic, adLockReadOnly
        If rsTest2.EOF Then
            mSaveAgin = True
            GoTo SaveAgin
        End If
        
        Cn.Execute "Delete Transaction_Details Where Transaction_ID = " & Transaction_ID
        sql = "Select NoteId from Transactions Where Transaction_ID = " & Transaction_ID
        
        RSNoteID.Open sql, Cn, adOpenKeyset, adLockOptimistic, adCmdText
        If Not RSNoteID.EOF Then
            NoteID = val(RSNoteID!NoteID & "")
        End If
End If
If Transaction_ID = 0 Then
    mSaveAgin = True
    GoTo SaveAgin
End If

Dim mTotal As Double
mTotal = 0
 Dim rsDummyItem As New ADODB.Recordset
        Dim RSTransDetails As New ADODB.Recordset
     
StrSQL = "SELECT     dbo.Transaction_Details.* from dbo.Transaction_Details Where (Transaction_ID = -1)"

   RSTransDetails.Open StrSQL, Cn, adOpenKeyset, adLockOptimistic, adCmdText
   
    If Transaction_Type = 19 Then
    
        For RowNum = 1 To FG2.Rows - 1
             
             

            If FG2.TextMatrix(RowNum, FG2.ColIndex("ItemID")) <> "" Then
                
                
                s = "Select ItemId from tblItems Where ItemId = " & val(FG2.TextMatrix(RowNum, FG2.ColIndex("ItemID"))) & " and ItemType <> 1"
                
                Set rsDummyItem = New ADODB.Recordset
                rsDummyItem.Open s, Cn, adOpenStatic, adLockReadOnly
                If rsDummyItem.EOF Then
                    GoTo NextRow2
                End If
                
                RSTransDetails.AddNew
                RSTransDetails("Transaction_ID").value = Transaction_ID
             
                RSTransDetails("ColorID").value = 1
                RSTransDetails("ItemSize").value = 1
                RSTransDetails("ClassId").value = 1
        RSTransDetails("Item_ID").value = IIf((FG2.TextMatrix(RowNum, FG2.ColIndex("ItemID")) = ""), Null, val(FG2.TextMatrix(RowNum, FG2.ColIndex("ItemID"))))
                RSTransDetails("UnitID").value = IIf((FG2.TextMatrix(RowNum, FG2.ColIndex("unitid")) = ""), Null, val(FG2.TextMatrix(RowNum, FG2.ColIndex("unitid"))))
               RSTransDetails("SHOWQTY").value = IIf((FG2.TextMatrix(RowNum, FG2.ColIndex("Qty")) = ""), Null, val(FG2.TextMatrix(RowNum, FG2.ColIndex("Qty"))))
               RSTransDetails("showPrice").value = IIf((FG2.TextMatrix(RowNum, FG2.ColIndex("Price")) = ""), Null, val(FG2.TextMatrix(RowNum, FG2.ColIndex("Price"))))
              
              

        
            LngCurItemID = val(FG2.TextMatrix(RowNum, FG2.ColIndex("ItemID")))
            
            LngUnitID = val(FG2.TextMatrix(RowNum, FG2.ColIndex("UnitID"))) 'val(Fg.Cell(flexcpData, RowNum, Fg.ColIndex("UnitID")))
            If LngUnitID = 0 Then
                GetDefaultItemUnit val(LngCurItemID), LngUnitID
            End If
            
            DblQty = val(FG2.TextMatrix(RowNum, FG2.ColIndex("Qty")))
            costPrice = val(FG2.TextMatrix(RowNum, FG2.ColIndex("cost")))
       '     costPrice = ModItemCostPrice.GetCostItemPrice(CLng(LngCurItemID), 0, "", , SystemOptions.SysMainStockCostMethod, DblQty, , XPDtbBill, , LngUnitID)
  ' costPrice = ModItemCostPrice.GetCostItemPrice(CLng(LngCurItemID), 0, "", , SystemOptions.SysMainStockCostMethod, DblQty, , XPDtbBill, , LngUnitID)
 'costPrice = 20
  ' CostTOTAL = CostTOTAL + costPrice * DblQty
  
            ' FG2.TextMatrix(RowNum, FG2.ColIndex("cost")) = costPrice
                  
          'RSTransDetails("ShowPrice").value = costPrice
          If costPrice < 0.09 Then
                RSTransDetails("showPrice").value = Round(costPrice / IIf(val(FG2.TextMatrix(RowNum, FG2.ColIndex("Qty"))) <> 0, val(FG2.TextMatrix(RowNum, FG2.ColIndex("Qty"))), 1), 7)
            Else
                RSTransDetails("showPrice").value = Round(costPrice / IIf(val(FG2.TextMatrix(RowNum, FG2.ColIndex("Qty"))) <> 0, val(FG2.TextMatrix(RowNum, FG2.ColIndex("Qty"))), 1), 6)
            
          End If
          RSTransDetails("showPrice").value = costPrice
        '  RSTransDetails("showPrice").value = Round(costPrice / IIf(val(fg2.TextMatrix(RowNum, fg2.ColIndex("Qty"))) <> 0, val(fg2.TextMatrix(RowNum, fg2.ColIndex("Qty"))), 1), 6)
         RSTransDetails("ShowQty").value = DblQty
                    
          

            StrSQL = "Select * From TblItemsUnits Where ItemID=" & LngCurItemID
            StrSQL = StrSQL + " AND UnitID=" & LngUnitID
            Set RsUnitData = New ADODB.Recordset
            RsUnitData.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText
        'fg2.TextMatrix(RowNum, fg2.ColIndex("Price")) = 0

            If Not (rs.BOF Or rs.EOF) And Not RsUnitData.EOF Then
 
                RSTransDetails("QtyBySmalltUnit").value = IIf(IsNull(RsUnitData("UnitFactor").value), 1, RsUnitData("UnitFactor").value)
                RSTransDetails("Quantity").value = RSTransDetails("QtyBySmalltUnit").value * RSTransDetails("showqty").value
                  RSTransDetails("Price").value = Round(val(IIf((FG2.TextMatrix(RowNum, FG2.ColIndex("cost")) = ""), 0, val(FG2.TextMatrix(RowNum, FG2.ColIndex("cost"))))) / RSTransDetails("QtyBySmalltUnit").value, 3)
            
            End If
            RSTransDetails("CostPrice").value = IIf((FG2.TextMatrix(RowNum, FG2.ColIndex("Price")) = ""), Null, val(FG2.TextMatrix(RowNum, FG2.ColIndex("Price"))))
                     If costPrice < 0.09 Then
                            CostTOTAL = CostTOTAL + (val(Round(val(RSTransDetails("showPrice").value) / RSTransDetails("QtyBySmalltUnit").value, 10)) * DblQty)
                    Else
                        CostTOTAL = CostTOTAL + (val(Round(val(RSTransDetails("showPrice").value) / RSTransDetails("QtyBySmalltUnit").value, 3)) * DblQty)
                    End If
            
                RSTransDetails.update
            End If
NextRow2:
        Next RowNum
    
    Else
        For RowNum = 1 To FG.Rows - 1
              If FG.RowHidden(RowNum) Or CBool(FG.ValueMatrix(RowNum, FG.ColIndex("IsDeleted"))) = True Then
                RowNum = RowNum
              End If
             If FG.RowHidden(RowNum) Or CBool(FG.ValueMatrix(RowNum, FG.ColIndex("IsDeleted"))) = True Then GoTo NextRow

            If FG.TextMatrix(RowNum, FG.ColIndex("ItemID")) <> "" Then
                
                
                
                s = "Select ItemId from tblItems Where ItemId = " & val(FG.TextMatrix(RowNum, FG.ColIndex("ItemID"))) & " and ItemType <> 1"
                
                Set rsDummyItem = New ADODB.Recordset
                rsDummyItem.Open s, Cn, adOpenStatic, adLockReadOnly
                If rsDummyItem.EOF Then
                    GoTo NextRow
                End If
                
                RSTransDetails.AddNew
                RSTransDetails("Transaction_ID").value = Transaction_ID
             
                RSTransDetails("ColorID").value = 1
                RSTransDetails("ItemSize").value = 1
                RSTransDetails("ClassId").value = 1
        RSTransDetails("Item_ID").value = IIf((FG.TextMatrix(RowNum, FG.ColIndex("ItemID")) = ""), Null, val(FG.TextMatrix(RowNum, FG.ColIndex("ItemID"))))
                RSTransDetails("UnitID").value = IIf((FG.TextMatrix(RowNum, FG.ColIndex("unitid")) = ""), Null, val(FG.TextMatrix(RowNum, FG.ColIndex("unitid"))))
               RSTransDetails("SHOWQTY").value = IIf((FG.TextMatrix(RowNum, FG.ColIndex("Qty")) = ""), Null, val(FG.TextMatrix(RowNum, FG.ColIndex("Qty"))))
              ' RSTransDetails("showPrice").value = IIf((fg.TextMatrix(RowNum, fg.ColIndex("Price")) = ""), Null, val(fg.TextMatrix(RowNum, fg.ColIndex("Price"))))
              
              
                          '«·ÊÕœ« 
           
   Dim mIsFromMix As Boolean
'             costPrice = GetCostFromMix2(RowNum)
'
'             If costPrice = 0 Then
'                costPrice = ModItemCostPrice.GetCostItemPrice(CLng(LngCurItemID), 0, , , SystemOptions.SysMainStockCostMethod, , , XPDtbBill, val(Me.TXTTransactionID2.Text), LngUnitID, val(Me.DCboStore2Name.BoundText))
'                mIsFromMix = False
'            Else
'                mIsFromMix = True
'             '   getItemCostData XPDtbBill.value, CLng(LngCurItemID), val(DCboStore2Name.BoundText), val(Me.TXTTransactionID2.Text), OldQty, OldCost, NewQty, NewCost
'             End If
             'FG.TextMatrix(RowNum, FG.ColIndex("cost")) = costPrice
        
            LngCurItemID = val(FG.TextMatrix(RowNum, FG.ColIndex("ItemID")))
            LngUnitID = val(FG.TextMatrix(RowNum, FG.ColIndex("UnitID"))) 'val(Fg.Cell(flexcpData, RowNum, Fg.ColIndex("UnitID")))
            If LngUnitID = 0 Then
                GetDefaultItemUnit val(LngCurItemID), LngUnitID
            End If
            
            DblQty = val(FG.TextMatrix(RowNum, FG.ColIndex("Qty")))
            costPrice = val(FG.TextMatrix(RowNum, FG.ColIndex("cost")))
       '     costPrice = ModItemCostPrice.GetCostItemPrice(CLng(LngCurItemID), 0, "", , SystemOptions.SysMainStockCostMethod, DblQty, , XPDtbBill, , LngUnitID)
  ' costPrice = ModItemCostPrice.GetCostItemPrice(CLng(LngCurItemID), 0, "", , SystemOptions.SysMainStockCostMethod, DblQty, , XPDtbBill, , LngUnitID)
 'costPrice = 20
   CostTOTAL = CostTOTAL + costPrice * DblQty
  mTotal = costPrice + mTotal
  
        If mIsFromMix Then
            
        Else
            costPrice = costPrice '* DblQty
        End If
            ' FG.TextMatrix(RowNum, FG.ColIndex("cost")) = costPrice
                  
          RSTransDetails("ShowPrice").value = costPrice
          
         RSTransDetails("ShowQty").value = DblQty
         RSTransDetails("Quantity").value = DblQty
         
                    
          

            StrSQL = "Select * From TblItemsUnits Where ItemID=" & LngCurItemID
            StrSQL = StrSQL + " AND UnitID=" & LngUnitID
            Set RsUnitData = New ADODB.Recordset
            RsUnitData.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText
       ' fg.TextMatrix(RowNum, fg.ColIndex("Price")) = 0

            If Not (rs.BOF Or rs.EOF) And Not RsUnitData.EOF Then
 
                RSTransDetails("QtyBySmalltUnit").value = DblQty ' IIf(IsNull(RsUnitData("UnitFactor").value), 1, RsUnitData("UnitFactor").value)
                RSTransDetails("Quantity").value = DblQty ' RSTransDetails("QtyBySmalltUnit").value * RSTransDetails("showqty").value
                  If val(RSTransDetails("QtyBySmalltUnit").value & "") <> 0 Then
                    If costPrice < 0.09 Then
                        RSTransDetails("Price").value = Round(val(IIf((FG.TextMatrix(RowNum, FG.ColIndex("cost")) = ""), 0, val(FG.TextMatrix(RowNum, FG.ColIndex("cost"))))) / RSTransDetails("QtyBySmalltUnit").value, 9)
                    Else
                        RSTransDetails("Price").value = Round(val(IIf((FG.TextMatrix(RowNum, FG.ColIndex("cost")) = ""), 0, val(FG.TextMatrix(RowNum, FG.ColIndex("cost"))))) / RSTransDetails("QtyBySmalltUnit").value, 3)
                    End If
                End If
            RSTransDetails("Price").value = costPrice
            End If
            RSTransDetails("CostPrice").value = costPrice
            
 
            
                RSTransDetails.update
            End If
NextRow:
        Next RowNum
    End If
             UpdateTransactionsCost CStr(Transaction_ID)
             
'Exit Sub
 
If Not IsSaveWithOutMsg Or mSaveAgin Then
CreateNotesEntry:
    NoteSerial = Notes_coding(val(BranchID), Transaction_Date)
     If Transaction_Type = 27 Then
        CreateNotes NoteID, Transaction_Date, CInt(BranchID), 240, mTotal, NoteSerial, NoteSerial1, "Transactions", "Transaction_ID", Transaction_ID, " »‰«¡« ⁄·Ï ”‰œ  Ã„Ì⁄ —ﬁ„ " & TxtTransSerial, ToHijriDate(Transaction_Date)
    Else
        CreateNotes NoteID, Transaction_Date, CInt(BranchID), 180, mTotal, NoteSerial, NoteSerial1, "Transactions", "Transaction_ID", Transaction_ID, " »‰«¡« ⁄·Ï ›« Ê—… „»Ì⁄«  —ﬁ„ " & TxtNoteSerial13, ToHijriDate(Transaction_Date)
    End If
Else
  If FG2.Rows = 1 Then Exit Sub
    s = "Select * from DOUBLE_ENTREY_VOUCHERS where Notes_Id = " & NoteID
    Dim rsDummyNotes As New ADODB.Recordset
    rsDummyNotes.Open s, Cn, adOpenStatic, adLockReadOnly
    If rsDummyNotes.EOF Then
        mSaveAgin = True
        GoTo CreateNotesEntry
    Else

        Cn.Execute "Update Notes Set Note_Value = " & CostTOTAL & " Where NoteId = " & NoteID
        Cn.Execute "Update DOUBLE_ENTREY_VOUCHERS  Set Value = " & CostTOTAL & " Where Notes_ID = " & NoteID
    End If
    
End If
'TxtNoteSerial11
'***********************
         If Transaction_Type = 19 Then
            StrSQL = "UPDATE TblDefComItem SET  TransactionID5=" & val(Transaction_ID) & ",  NoteSerial15='" & NoteSerial1 & "' WHERE ID  =" & val(TxtTransSerial)
            Cn.Execute StrSQL
            
            StrSQL = "UPDATE Transactions SET  Nots=" & val(TXTTransactionID3) & ",BillBasedOn =2,nots2 = '" & Trim(TxtNoteSerial13.Text) & "',Closed = 1   WHERE Transaction_ID  =" & val(TXTTransactionID5)
            Cn.Execute StrSQL
            
        Else
            StrSQL = "UPDATE TblDefComItem SET  TransactionID1=" & val(Transaction_ID) & ",  NoteSerial11='" & NoteSerial1 & "' WHERE ID  =" & val(TxtTransSerial)
            
            Cn.Execute StrSQL
            TxtNoteSerial1 = NoteSerial1
        End If
'***********************
If Not IsSaveWithOutMsg Or mSaveAgin Then
  CREATE_VOUCHER_GE1 Transaction_ID, NoteSerial1, "", NoteID, val(dcBranch.BoundText), StoreId, Transaction_Date, 0, invoice
End If
 
        'StrSQL = "UPDATE Transactions SET NOTS=" & Transaction_ID & " WHERE Transaction_ID=" & val(Me.XPTxtBillID.text)
        'Cn.Execute StrSQL
  'MsgBox " „   «·‰ﬁ·"
  
'******************************************************issueVoucher








     
 
    '
 
ErrTrap:

End Sub

 
 

   
  

 

Private Sub Undo()
    On Error GoTo ErrTrap
    
    
        Dim i As Long, m As Long
        
        With FG
            For i = 1 To .Rows - 1
                If .RowHidden(i) And CBool(FG.ValueMatrix(.Row, FG.ColIndex("IsDeleted"))) = False Then
                    .RowHidden(i) = False

                End If
            Next
        End With
    
    Select Case TxtModFlg.Text

        Case "N"
            clear_all Me
            Me.TxtModFlg.Text = "R"
            XPBtnMove_Click (1)

        Case "E"
            rs.Find "id='" & val(TxtTransSerial.Text) & "'", , adSearchForward, adBookmarkFirst

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
Private Sub Del_Trans()
    Dim Msg As String
    Dim StrSQL As String
Dim StrSQL1 As String
Dim sql As String
Dim i As Integer
    'On Error GoTo ErrTrap

    If Me.CboPayMentType.ListIndex = 0 Or Me.CboPayMentType.ListIndex = 1 Then

        '›« Ê—… ‰ﬁœÌ…
        If CheckBoxAccount(val(Me.DcboBox.BoundText), val(txtTotalWithVat), XPDtbBill.value, False) = False Then
            If SystemOptions.UserInterface = ArabicInterface Then
                Msg = "·‰ Ì„ﬂ‰ «·”„«Õ »Õ–› Â–« «·⁄„·Ì…..!!!"
                Msg = Msg & CHR(13) & "ÕÌÀ «‰Â« ”Ê› Ì‰ Ã ⁄‰Â« Œÿ« ›Ï Õ”«»«  «·Œ“‰…"
            Else
                Msg = "You will not be allowed to delete this process .. !!!"
                Msg = Msg & CHR(13) & "Where it will result in a line in the treasury accounts"
            
            End If
            MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
            Exit Sub
        End If
    End If
    

    If TxtTransSerial.Text <> "" Then
        If SystemOptions.UserInterface = EnglishInterface Then
            Msg = "The process data will be deleted" & CHR(13)
            Msg = Msg + " Do you want to delete this data?"
        Else
            Msg = "”Ì „ Õ–› »Ì«‰«  «·⁄„·Ì… —ﬁ„ " & CHR(13)
            Msg = Msg + " Â·  —€» ›Ì Õ–› Â–Â «·»Ì«‰« ø"
        
        End If
        If MsgBox(Msg, vbYesNo + vbQuestion + vbMsgBoxRight + vbMsgBoxRtlReading, App.title) = vbYes Then
            If Not rs.RecordCount < 1 Then
           DeleteTransactiomsVoucher2 val(TXTTransactionID1.Text)
        DeleteTransactiomsVoucher2 val(TXTTransactionID2.Text)
          DeleteTransactiomsVoucher2 val(TXTTransactionID3.Text)
            DeleteTransactiomsVoucher2 val(TXTTransactionID4.Text)
              DeleteTransactiomsVoucher2 val(TXTTransactionID5.Text)
       
       rs.delete
               
               
                    
                    StrSQL = "Delete Transactions Where Transaction_ID In (Select TransactionID4 From TblDefComItemData Where IDDefCIT = " & val(TxtTransSerial) & ")"
                    Cn.Execute StrSQL, , adExecuteNoRecords
                    StrSQL = "Delete Transaction_Details Where Transaction_ID In (Select TransactionID4 From TblDefComItemData Where IDDefCIT = " & val(TxtTransSerial) & ")"
                    Cn.Execute StrSQL, , adExecuteNoRecords
                    
              
                
                   StrSQL = "Delete TblDefComItem where ID=" & val(Me.TxtTransSerial.Text)
        Cn.Execute StrSQL, , adExecuteNoRecords
   
        StrSQL = "Delete From TblDefComItemDet Where IDDefCIT=" & val(TxtTransSerial.Text)
        Cn.Execute StrSQL, , adExecuteNoRecords

        StrSQL = "Delete From TblDefComItemData Where IDDefCIT=" & val(TxtTransSerial.Text)
        Cn.Execute StrSQL, , adExecuteNoRecords




'                StrSQL = "Delete From DOUBLE_ENTREY_VOUCHERS Where AdvanceID=" & val(Me.XPTxtID.text)
'                Cn.Execute StrSQL, , adExecuteNoRecords
                rs.MoveFirst
   

              
                    clear_all Me
                      '  ListGroupSelected.Clear
   ' ListStoreSelected.Clear

                   FG.Clear flexClearScrollable, flexClearEverything
                   FG.Rows = 1
                    TxtModFlg_Change
                    XPTxtCurrent.Caption = 0
                    XPTxtCount.Caption = 0
                Else
                 
                End If
           ' End If
        End If
   Retrive
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
    Msg = "⁄›Ê«...ÕœÀ Œÿ√ „« √À‰«¡ Õ–› Â–Â «·»Ì«‰«  " & CHR(13)
    MsgBox Msg, vbMsgBoxRight + vbMsgBoxRtlReading + vbExclamation, App.title
    rs.CancelUpdate
End Sub
Sub DeleteRow(Optional ByVal mRow As Long = 0, Optional ByVal FromIsAdd As Boolean = False)
Dim j As Long
Dim mItemNo  As Long



        With Me.FG
            If mRow = 0 Then mRow = .Row
            If mRow <= 0 Then Exit Sub
            Dim mIsAdd As Boolean, mValue As Double
            
            mIsAdd = CBool(FG.ValueMatrix(mRow, FG.ColIndex("IsAdd")))
            If (Not mIsAdd And FromIsAdd) Or Not mIsAdd Then
                mValue = val(FG.TextMatrix(mRow, FG.ColIndex("Total")))
                FG.TextMatrix(mRow, FG.ColIndex("IsDeleted")) = True
                mItemNo = val(FG.TextMatrix(mRow, FG.ColIndex("ItemId2")))
                txtTotalDisc = val(txtTotalDisc) + mValue
                
                .RowHidden(mRow) = True
                FillDelGrid
             Else
                .RemoveItem mRow
            End If
            
        End With
        CalcTotalNet mItemNo
      '  ReLineGrid
        CalcGrid2

End Sub


Sub DeleteRow2()
Dim i As Long
Dim j As Long
        For j = 1 To FG2.Rows - 1
        
            If j > FG2.Rows - 1 Then
                Exit Sub
            End If
            mItemNo = val(FG2.TextMatrix(j, FG2.ColIndex("ItemID")))
            mLineNo = val(FG2.TextMatrix(j, FG2.ColIndex("LineID")))
            If CBool(FG2.ValueMatrix(j, FG2.ColIndex("Select"))) Then
                For i = 1 To FG.Rows - 1
                    If i <= FG.Rows - 1 Then
                        If val(FG.TextMatrix(i, FG.ColIndex("ItemID2"))) = val(FG2.TextMatrix(j, FG2.ColIndex("ItemID"))) And val(FG.TextMatrix(i, FG.ColIndex("LineID"))) = val(FG2.TextMatrix(j, FG2.ColIndex("LineID"))) Then
                           FG.RemoveItem i
                           i = i - 1
                           
                        End If
                    End If
                Next
            
            
               ' For i = 1 To fg2.Rows - 1
                    If j <= FG2.Rows - 1 Then
                        If val(FG2.TextMatrix(j, FG2.ColIndex("ItemId"))) = mItemNo And val(FG2.TextMatrix(j, FG2.ColIndex("LineID"))) = mLineNo Then
                           FG2.RemoveItem j
                           
                           j = j - 1
                        End If
                    End If
                
            CalcTotalNet
         '   ReLineGrid
            CalcGrid2
        End If
        
Next j
End Sub



Private Sub FillDelGrid()
        Dim i As Long, m As Long
        FGDeleted.Rows = 1
        With FG
            For i = 1 To .Rows - 1
                If .RowHidden(i) Then
                    m = m + 1
                    FGDeleted.AddItem m
                     
                    FGDeleted.TextMatrix(m, FGDeleted.ColIndex("Row2")) = i
                    FGDeleted.TextMatrix(m, FGDeleted.ColIndex("FlgX")) = .TextMatrix(i, .ColIndex("FlgX"))
                    FGDeleted.TextMatrix(m, FGDeleted.ColIndex("Ser")) = .TextMatrix(i, .ColIndex("Ser"))
                    FGDeleted.TextMatrix(m, FGDeleted.ColIndex("ItemID")) = .TextMatrix(i, .ColIndex("ItemID"))
                    FGDeleted.TextMatrix(m, FGDeleted.ColIndex("itemcode")) = .TextMatrix(i, .ColIndex("itemcode"))
                    FGDeleted.TextMatrix(m, FGDeleted.ColIndex("UnitID")) = .TextMatrix(i, .ColIndex("UnitID"))
                    FGDeleted.TextMatrix(m, FGDeleted.ColIndex("itemname")) = .TextMatrix(i, .ColIndex("itemname"))
                    FGDeleted.TextMatrix(m, FGDeleted.ColIndex("unitname")) = .TextMatrix(i, .ColIndex("unitname"))
                    FGDeleted.TextMatrix(m, FGDeleted.ColIndex("Price")) = .TextMatrix(i, .ColIndex("Price"))
                    FGDeleted.TextMatrix(m, FGDeleted.ColIndex("Total")) = .TextMatrix(i, .ColIndex("Total"))
                End If
            Next
        End With
         
End Sub
Private Sub Cmd_Click(Index As Integer)
    'Dim intDef As Integer
            Dim s As String
             Dim intDef As Integer
            Dim RsData As New ADODB.Recordset
' On Error GoTo ErrTrap
Dim j As Long
    Select Case Index
Case 10
TxtTransSerial.Text = ""
TxtModFlg.Text = "N"
  FG.Rows = FG.Rows + 1
            FG.Enabled = True
            cmd(1).Enabled = True

        Case 0
            If DoPremis(Do_New, Me.Name, True) = False Then
                Exit Sub
            End If

            clear_all Me
            mNewId = 0
            mIdDisplay = 0
            TxtModFlg.Text = "N"
            cmdTransfer.Enabled = True
            cmdCancel.Enabled = False
            Me.DCboUserName.BoundText = user_id
 
            FG.Clear flexClearScrollable, flexClearEverything
             
            FG.Rows = 1
            FG.Enabled = True
            Selct_Click (0)
            Selct_Click (1)
            Selct_Click (2)
            Dim RsOptions As ADODB.Recordset
            
            FG.Enabled = True
          '  FG.Rows = 2
            XPDtbBill.value = Date
            XPDtRecDate.value = DateAdd("d", 3, XPDtbBill.value)
            FG.Editable = flexEDKbdMouse
            
            Me.DCboUserName.BoundText = user_id
            intDef = val(GetSetting(StrAppRegPath, "DefaultOptions", "DefaultClient", 2))
            DBCboClientName.BoundText = intDef
            intDef = val(GetSetting(StrAppRegPath, "DefaultOptions", "DefaultSaleStore", 1))
            DCboStoreName.BoundText = intDef
            DcboItemID1.Tag = ""
            Set RsOptions = New ADODB.Recordset
            RsOptions.Open "tbloptions", Cn, adOpenStatic, adLockReadOnly, adCmdTable

            If Not (RsOptions.BOF Or RsOptions.EOF) Then
                Me.DcboBox.BoundText = IIf(IsNull(RsOptions("SalesBoxID").value), "", RsOptions("SalesBoxID").value)
            End If

            dcBranch.BoundText = Current_branch

            Dim dstore As Integer
            Dim dBox As Integer
            Dim usertype As Integer
            Dim EmpID As Integer
            Dim userbranchid As Integer
            Dim CUSTID As Integer
            Dim dStore2 As Integer
            'GetBranchData branch_id, dstore, dBox
                cmdCreateProduction.Enabled = False
                cmdCancel2.Visible = False
           GetUserData user_id, usertype, userbranchid, dstore, dBox, , EmpID, , CUSTID, dStore2
     'intDef
            CboPayMentType.ListIndex = 0
            DBCboClientName.BoundText = CUSTID
           
          
          Me.dcBranch.BoundText = userbranchid
          Me.DCboStoreName.BoundText = dstore
          Me.DcboBox.BoundText = dBox
          Me.DcboEmp.BoundText = EmpID
          Me.DCboStore2Name.BoundText = dStore2
          DCboStore3Name.BoundText = dstore
                    
            s = "Select StoreID,StoreID1,StoreID2,StoreID3 from tblUsers Where UserID = " & user_id
            Set rsDummy = New ADODB.Recordset
            rsDummy.Open s, Cn, adOpenKeyset, adLockReadOnly, adCmdText
            If Not rsDummy.EOF Then
                DCboStore2Name.BoundText = val(rsDummy!StoreId2 & "")
                DCboStore3Name.BoundText = val(rsDummy!StoreID3 & "")
                DCboStoreName.BoundText = val(rsDummy!StoreId1 & "")
            End If
  

             Selct(0).value = vbChecked
           Selct(1).value = vbChecked
          

 
            If Current_branch = 0 Then
                'branch_id = my_branch
                Me.dcBranch.BoundText = Current_branch
            End If
 
             If Not SystemOptions.UserInterface = ArabicInterface Then

                cmdCreateProduction.Caption = "«‰‘«¡ «„— «‰ «Ã"
            Else
                cmdCreateProduction.Caption = "Create a production order"

            End If

            cmdCreateProduction.Enabled = False
            If SystemOptions.PaymentMethLaterCompItem = True Then
                CboPayMentType.Enabled = False
                CboPayMentType.ListIndex = 1
            End If

            'cmdAddCustomer.Caption = ""
'            TxtNoteSerialV = ""
'            TxtNoteSerial1V = ""
'            If SystemOptions.DefaultIsCreditSales = False Then
'                CboPayMentType.ListIndex = 0
'             Else
'                CboPayMentType.ListIndex = 1
'             End If
'            CboPayMentType_Click

        Case 1

            If DoPremis(Do_Edit, Me.Name, True) = False Then
                Exit Sub
            End If
            
           Selct(0).value = vbChecked
           Selct(1).value = vbChecked
            s = "Select UserId From  TblProductLineDistribution Where IDDefCIT = " & val(TxtTransSerial) & " and  IsNull(UserId,'') <> '' "
            Set RsData = New ADODB.Recordset
            RsData.Open s, Cn, adOpenStatic, adLockReadOnly, adCmdText
            If Not RsData.EOF Or val(TXTTransactionID3) <> 0 Then
                RsData.Close
                
                 If SystemOptions.UserInterface = ArabicInterface Then
                    MsgBox "·« Ì„ﬂ‰  ⁄œÌ· «·Õ—ﬂ… ‰Ÿ—« ·ÊÃÊœ  ÕÊÌ· "
                Else
                    MsgBox "The Transaction can not be modified because of a conversion "
                End If
                
                Exit Sub
            End If
'            If val(TXTTransactionID4) <> 0 Then
'                MsgBox "·« Ì„ﬂ‰  ⁄œÌ· «·Õ—ﬂ… ‰Ÿ—« ·ÊÃÊœ «„— «‰ «Ã "
'                Exit Sub
'            End If
         
            mNewId = 0
            mIdDisplay = 0
            TxtModFlg.Text = "E"
            Me.DCboUserName.BoundText = user_id
           ' FG.Rows = FG.Rows
           cmdRecalc_Click
            FG.Enabled = True
            cmdCreateProduction.Enabled = False
        
       Case 2
      ' If TxtMaxNo.Text = "" Then
      ' If SystemOptions.UserInterface = ArabicInterface Then
      ' MsgBox "Ì—ÃÏ «œŒ«· ﬂÊœ «·„ﬂ”"
      ' Else
      ' MsgBox "Please enter code"
      ' End If
      ' TxtMaxNo.SetFocus
      ' Exit Sub
      ' End If
If SystemOptions.RawMaterMix = True Then
If val(DCboStore2Name.BoundText) = 0 Then
If SystemOptions.UserInterface = ArabicInterface Then
MsgBox "Ì—ÃÏ «Œ Ì«— „Œ“‰ «·’—›"
Else
MsgBox "Please select store"
End If
DCboStore2Name.SetFocus
Exit Sub
End If
End If
            SaveData

        Case 3
Undo
mNewId = 0
mIdDisplay = 0

        Case 4

            If DoPremis(Do_Delete, Me.Name, True) = False Then
                Exit Sub
            End If
        

            s = "Select UserId From  TblProductLineDistribution Where IDDefCIT = " & val(TxtTransSerial) & " and  IsNull(UserId,'') <> '' "
            Set RsData = New ADODB.Recordset
            RsData.Open s, Cn, adOpenStatic, adLockReadOnly, adCmdText
            If Not RsData.EOF Or val(TXTTransactionID3) <> 0 Then
                RsData.Close
                If SystemOptions.UserInterface = ArabicInterface Then
                    MsgBox "·« Ì„ﬂ‰ «·€«¡ «·Õ—ﬂ… ‰Ÿ—« ·ÊÃÊœ  ÕÊÌ· "
                Else
                    MsgBox "You can not cancel the animation because there is a conversion "
                End If
                Exit Sub
            Else
                mNewId = 0
                mIdDisplay = 0
                cmdCancel_Click
'                s = "Delete TblProductLineDistribution Where IDDefCIT = " & val(TxtTransSerial)
'                Cn.Execute s
            End If
            
                If val(TXTTransactionID4) <> 0 Then
                    If SystemOptions.UserInterface = ArabicInterface Then
                        MsgBox "·« Ì„ﬂ‰  ⁄œÌ· «·Õ—ﬂ… ‰Ÿ—« ·ÊÃÊœ «„— «‰ «Ã "
                    Else
                        MsgBox "The motion can not be adjusted due to a production order "
                    End If
                    Exit Sub
                End If
        Del_Trans

        Case 5

            If DoPremis(Do_Search, Me.Name, True) = False Then
                Exit Sub
            End If

          Load FrmSearchDevComItem
            FrmSearchDevComItem.show
    
  '  Retrive 104523
        Case 7

            If DoPremis(Do_Print, Me.Name, True) = False Then
                Exit Sub
            End If

            print_report

            '        PrintReport1 (Txt_order_no.text)
        Case 6
          Unload Me

       Case 8
            
            For j = 1 To FG.Rows - 1
                If CBool(FG.ValueMatrix(j, FG.ColIndex("Select"))) Then
                    DeleteRow j
                End If
            Next
       Case 11
        DeleteRow2
      
    End Select

    Exit Sub
ErrTrap:
End Sub




 

Private Sub cmdAdd__Click()
If val(DcboItemID2.BoundText) = val(val(DcboItemID1.BoundText)) Then DcboItemID2.Text = "": Exit Sub
If val(DcboItemID5.BoundText) = val(val(DcboItemID1.BoundText)) Then DcboItemID5.Text = "": Exit Sub

AddNewFgRow val(DcboItemID2.BoundText), "ItemID2", "ItemName2"
If TXT_order_no = "" Then
    AddNewFgRow val(DcboItemID5.BoundText), "ItemID5", "ItemName5"
End If

CalcGrid2
End Sub
Private Sub CalcGrid2(Optional ByVal isRetrive As Boolean = False)
    Dim i As Long
    Dim isProductOrder As Boolean
    txtTotal2 = 0
    txtTotalAdd2 = 0
    txtTotalDisc2 = 0
    txtNet2 = 0
    TxtVAt22 = 0
    txtTotalWithVat2 = 0
    isProductOrder = False
    For i = 1 To FG2.Rows - 1
        
        'If Me.TxtModFlg.Text <> "R" Then
            FG2.TextMatrix(i, FG2.ColIndex("Total")) = val(FG2.TextMatrix(i, FG2.ColIndex("Qty"))) * val(FG2.TextMatrix(i, FG2.ColIndex("Price")))
        'End If
            If Trim(FG2.TextMatrix(i, FG2.ColIndex("NoteSerial14"))) <> "" And Not (isProductOrder) And isRetrive Then
                isProductOrder = True
                cmdCancel2.Visible = True
                cmdCancel2.Enabled = True
               ' cmdCreateProduction.Enabled = False
                
            End If
         FG2.TextMatrix(i, FG2.ColIndex("TotalAdd")) = 0
         FG2.TextMatrix(i, FG2.ColIndex("TotalDisc")) = 0
         FG2.TextMatrix(i, FG2.ColIndex("Net")) = 0
         Dim j As Long
         For j = 1 To FG.Rows - 1
            If val(FG.TextMatrix(j, FG.ColIndex("ItemID2"))) = val(FG2.TextMatrix(i, FG2.ColIndex("ItemID"))) And val(FG.TextMatrix(j, FG.ColIndex("LineID"))) = val(FG2.TextMatrix(i, FG2.ColIndex("LineID"))) Then
                If (FG.ValueMatrix(j, FG.ColIndex("IsAdd"))) Then
                    FG2.TextMatrix(i, FG2.ColIndex("TotalAdd")) = val(FG2.TextMatrix(i, FG2.ColIndex("TotalAdd"))) + val(FG.TextMatrix(j, FG.ColIndex("Total")))
                End If
               ' FG2.TextMatrix(i, FG2.ColIndex("cost")) = (val(FG2.TextMatrix(i, FG2.ColIndex("cost"))) * val(FG2.TextMatrix(i, FG2.ColIndex("Qty")))) + val(FG.TextMatrix(j, FG.ColIndex("cost")))
                If (FG.ValueMatrix(j, FG.ColIndex("IsDeleted"))) Or val((FG.ValueMatrix(j, FG.ColIndex("OldPrice")))) <> 0 Then
                    If val((FG.ValueMatrix(j, FG.ColIndex("OldPrice")))) <> 0 Then
                        FG2.TextMatrix(i, FG2.ColIndex("TotalDisc")) = val(FG2.TextMatrix(i, FG2.ColIndex("TotalDisc"))) + val((FG.ValueMatrix(j, FG.ColIndex("OldPrice"))))
                    Else
                        FG2.TextMatrix(i, FG2.ColIndex("TotalDisc")) = val(FG2.TextMatrix(i, FG2.ColIndex("TotalDisc"))) + val(FG.TextMatrix(j, FG.ColIndex("Total")))
                    End If
                End If
            End If
         Next
        If isProductOrder And isRetrive Then
            isProductOrder = True
            cmdCancel2.Visible = True
            cmdCancel2.Enabled = True
            'cmdCreateProduction.Enabled = False
        ElseIf Not isProductOrder And isRetrive Then
            
            cmdCancel2.Enabled = False
            cmdCreateProduction.Enabled = True
            
        End If
        CalcDisc i
        
        FG2.TextMatrix(i, FG2.ColIndex("Net")) = val(FG2.TextMatrix(i, FG2.ColIndex("Total"))) + val(FG2.TextMatrix(i, FG2.ColIndex("TotalAdd"))) - val(FG2.TextMatrix(i, FG2.ColIndex("TotalDisc")))
        CalCulteVAT 3, i
        
        
    
    
        txtTotal2 = val(txtTotal2) + val(FG2.TextMatrix(i, FG2.ColIndex("Total")))
        txtTotalAdd2 = val(txtTotalAdd2) + val(FG2.TextMatrix(i, FG2.ColIndex("TotalAdd")))
        txtTotalDisc2 = val(txtTotalDisc2) + val(FG2.TextMatrix(i, FG2.ColIndex("TotalDisc")))
        txtNet2 = val(txtNet2) + val(FG2.TextMatrix(i, FG2.ColIndex("Net")))
        TxtVAt22 = val(TxtVAt22) + val(FG2.TextMatrix(i, FG2.ColIndex("VAt2")))
        txtTotalWithVat2 = val(txtTotalWithVat2) + val(FG2.TextMatrix(i, FG2.ColIndex("TotalWithVat")))
        
    Next
End Sub
Private Sub GetCostFromMix(ByVal mRow As Long)
    If Trim(TxtMaxNo2) = "" Then Exit Sub
    Dim mItemNo As Long
    Dim mUnitId As Integer
    mItemNo = val(FG2.TextMatrix(mRow, FG2.ColIndex("ItemID")))
    mUnitId = val(FG2.TextMatrix(mRow, FG2.ColIndex("UnitID")))
    
    Dim s As String
    Dim rsDummy As New ADODB.Recordset
    
    s = " SELECT TblDefComItemData.cost,TblDefComItemData.Price FROM TblDefComItemData"
    s = s & " Inner Join"
    s = s & " TblDefComItem"
    s = s & " ON TblDefComItem.ID = TblDefComItemData.IDDefCIT"
    s = s & " Where MaxNo = N'" & Trim(TxtMaxNo2) & "'"
    s = s & " AND TblDefComItemData.ItemID = " & mItemNo
    s = s & " AND TblDefComItemData.UnitID =" & mUnitId
    rsDummy.Open s, Cn, adOpenStatic, adLockReadOnly
    If Not rsDummy.EOF Then
       ' GetCostFromMix = val(rsDummy!cost & "")
        If val(rsDummy!cost & "") <> 0 Then
            FG2.TextMatrix(mRow, FG2.ColIndex("cost")) = val(rsDummy!cost & "")
        End If
        If val(rsDummy!Price & "") <> 0 Then
            FG2.TextMatrix(mRow, FG2.ColIndex("Price")) = val(rsDummy!Price & "")
        End If
        
    End If
    


End Sub

Private Function GetCostFromMix2(ByVal mRow As Long) As Double
    If Trim(TxtMaxNo2) = "" Then Exit Function
    Dim mItemNo As Long
    Dim mUnitId As Integer
    Dim mItemNo2 As Long
    mItemNo = val(FG.TextMatrix(mRow, FG.ColIndex("ItemID")))
    mItemNo2 = val(FG.TextMatrix(mRow, FG.ColIndex("ItemID2")))
    mUnitId = val(FG.TextMatrix(mRow, FG.ColIndex("UnitID")))
    
    Dim s As String
    Dim rsDummy As New ADODB.Recordset
    
        s = " SELECT tdcid.cost,tdcid.Price"
    s = s & " FROM   TblDefComItemDet AS tdcid"
    s = s & "        INNER JOIN TblDefComItemData"
    s = s & "             ON  tdcid.IDDefCIT = TblDefComItemData.IDDefCIT"
    s = s & "             AND tdcid.ItemID2 = TblDefComItemData.ItemID"
    s = s & "        RIGHT OUTER JOIN TblDefComItem"
    s = s & "             ON  TblDefComItemData.IDDefCIT = TblDefComItem.ID"
    s = s & " Where MaxNo = N'" & Trim(TxtMaxNo2) & "'"
    s = s & " AND tdcid.itemId= " & mItemNo
    s = s & " AND tdcid.UnitID =" & mUnitId
    s = s & " AND tdcid.ItemID2 =" & mItemNo2
    
    s = " SELECT TblDefComItemData.cost,TblDefComItemData.Price FROM TblDefComItemData"
    s = s & " Inner Join"
    s = s & " TblDefComItem"
    s = s & " ON TblDefComItem.ID = TblDefComItemData.IDDefCIT"
    s = s & " Where MaxNo = N'" & Trim(TxtMaxNo2) & "'"
    s = s & " AND TblDefComItemData.ItemID = " & mItemNo
    s = s & " AND TblDefComItemData.UnitID =" & mUnitId
    
    rsDummy.Open s, Cn, adOpenStatic, adLockReadOnly
    If Not rsDummy.EOF Then
        GetCostFromMix2 = val(rsDummy!cost & "")
        If val(rsDummy!cost & "") <> 0 Then
            FG.TextMatrix(mRow, FG.ColIndex("cost")) = val(rsDummy!cost & "")
            GetCostFromMix2 = val(rsDummy!cost & "")
        End If
        If val(rsDummy!Price & "") <> 0 Then
            FG.TextMatrix(mRow, FG.ColIndex("Price")) = val(rsDummy!Price & "")
        End If
        
    End If
    


End Function
Private Sub cmdAdd_Click()
If val(DcboItemID1.BoundText) = 0 Then
If SystemOptions.UserInterface = ArabicInterface Then
MsgBox "Ì—ÃÏ «Œ Ì«— «·’‰› "
Else
MsgBox "Please Select Item"
End If
DcboItemID1.SetFocus
Exit Sub
End If
 If val(txtQty1.Text) = 0 Then
 If SystemOptions.UserInterface = ArabicInterface Then
 MsgBox "Ì—ÃÏ «œŒ«· «·ﬂ„Ì…"
 Else
 MsgBox "Please Eneter Qty"
 End If
 txtQty1.SetFocus
 Exit Sub
 End If
 FillGridItemType val(DcboItemID1.BoundText), DcboItemID1.Text, Trim$(TxtAttachedItemCode.Text), 1, val(DcbUnit.BoundText), DcbUnit.Text, val(txtQty1), val(TxtPrice), val(XPCboGroup.BoundText), XPCboGroup.Text
 If SystemOptions.IsMultiItemsInCompItem Then
    DcboItemID1.BoundText = ""
    XPCboGroup2.BoundText = ""
    XPCboGroup5.BoundText = ""
 End If

End Sub

Private Sub CMDSHOWecive_Click()
  
  FrmInpoutWorkOrder.Retrive val(TXTTransactionID2.Text)
End Sub

Private Sub CMDSHOWISSUE_Click()
 'FrmOut.Retrive val(TXTTransactionID1.Text)
 FrmOutProductionOrder.Retrive val(TXTTransactionID1.Text)
 
End Sub

'Private Function GetMaxLineNo() As Long
'    Dim i As Long
'    Dim mLine As Long, mMaxNo As Long
'
'    With FG
'    For i = 1 To .Rows - 1
'
'    Next
'
'End Function

Private Function GetFieldName(ByVal mTable As String) As String
If mTable = "TblItems" Then
    If SystemOptions.UserInterface = ArabicInterface Then
        GetFieldName = " TblItems.ItemName"
    Else
        GetFieldName = " TblItems.ItemNamee"
    End If
ElseIf mTable = "TblUnites" Then
    If SystemOptions.UserInterface = ArabicInterface Then
        GetFieldName = " TblUnites.UnitName"
    Else
        GetFieldName = " TblUnites.UnitNamee"
    End If
End If
End Function
Private Function FillGrid(Optional ByVal mLineID As Long = 0) As Boolean
 
 Dim LngNewRow As Long
    If val(Me.DcboItemID2.BoundText) = 0 And chkIsAdd.value = vbChecked Then
        If SystemOptions.UserInterface = ArabicInterface Then
            Msg = "ÌÃ»  ÕœÌœ «Ú”„ «·’‰› ...!!!"
        Else
            Msg = "Must specify the name of the product ... !!!"
        End If
        MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
        Me.DcboItemID2.SetFocus
        FillGrid = False
        Exit Function
    End If


' If val(txtQty.Text) = 0 And chkIsAdd.value = vbChecked Then
'    If SystemOptions.UserInterface = ArabicInterface Then
'        MsgBox "Ì—ÃÏ «œŒ«· «·ﬂ„Ì…"
'    Else
'        MsgBox "Please Eneter Qty"
'    End If
'    txtQty.SetFocus
'    fillgrid = False
'    Exit Function
' End If


    If val(Me.DcbUnit2.BoundText) = 0 And chkIsAdd.value = vbChecked Then
        
        If SystemOptions.UserInterface = ArabicInterface Then
            Msg = "ÌÃ»  ÕœÌœ ÊÕœ…  «·’‰› ...!!!"
        Else
                    Msg = "Must select the unit of the item ... !!!"
        End If
        MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
        Me.DcbUnit2.SetFocus
        FillGrid = False
        Exit Function
    End If
 
 'FG.Clear flexClearScrollable, flexClearEverything
          '  FG.Rows = 1
          
' For l = 1 To FG.Rows - 1
'        If l > FG.Rows - 1 Then Exit For
'        If FG.TextMatrix(l, FG.ColIndex("ItemID2")) = DcboItemID1.BoundText Or FG.TextMatrix(l, FG.ColIndex("ItemID")) = "" Then
'            FG.RemoveItem l
'            l = l - 1
'        End If
'    Next
Dim StrSQL As String
Dim i As Integer
Dim k As Integer
Dim Rs3 As ADODB.Recordset
Set Rs3 = New ADODB.Recordset

 


    



    



                    StrSQL = " SELECT       ForUnit   ,MethodCalc,   TblItemsParts.lowering ,TblItemsParts.increase, dbo.TblItemsParts.Unitid, dbo.TblItemsParts.isReplaced, dbo.TblItemsParts.PartItemPrice, dbo.TblItemsParts.PartItemQty, dbo.TblItemsParts.PartItemID, "
                StrSQL = StrSQL + " Price = dbo.GetItemPrice(dbo.TblItemsParts.PartItemID,dbo.TblUnites.UnitID," & IIf(SystemOptions.AllowLastPrice, 1, 0) & "),"
                StrSQL = StrSQL + "      dbo.TblItemsParts.ItemID, dbo.TblItemsParts.TableID, dbo.TblUnites.UnitName, dbo.TblUnites.UnitNamee, dbo.TblItems.ItemCode, dbo.TblItems.ItemName,"
                StrSQL = StrSQL + "      dbo.TblItems.ItemNamee , dbo.TblItems.fullcode"
                StrSQL = StrSQL + "  FROM         dbo.TblItemsParts INNER JOIN"
                StrSQL = StrSQL + "      dbo.TblUnites ON dbo.TblItemsParts.Unitid = dbo.TblUnites.UnitID RIGHT OUTER JOIN"
                StrSQL = StrSQL + "      dbo.TblItems ON dbo.TblItemsParts.PartItemID = dbo.TblItems.ItemID"
                StrSQL = StrSQL + " Where (dbo.TblItemsParts.ItemID = " & val(DcboItemID1.BoundText) & ")"
                StrSQL = StrSQL + " ORDER BY dbo.TblItemsParts.TableID"
'       Rs3.Close
        Rs3.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText
        If SystemOptions.IsMultiItemsInCompItem Then
            If FG2.Rows <= 2 And FG2.Rows > 1 Then
                If Trim(FG2.TextMatrix(1, FG2.ColIndex("itemcode"))) = "" Then
                    mNewId = 1
                
               Else
                    mNewId = FG2.Rows - 1
                End If
            Else
                'if fg2.Rows
                 mNewId = FG2.Rows - 1
            End If
            
        Else
        
            mNewId = 1
        End If
        If Rs3.EOF Then
            If SystemOptions.IsMultiItemsInCompItem Then
                     StrSQL = " SELECT       1 ForUnit   ,1 MethodCalc,  0 lowering ,0 increase, " & val(DcbUnit.BoundText) & "  Unitid, 0 isReplaced, " & TxtPrice & " PartItemPrice, " & val(txtQty1) & " PartItemQty, dbo.TblItems.ItemID as PartItemID, "
                If val(TxtPrice) = 0 Then
                    StrSQL = StrSQL + " Price = dbo.GetItemPrice( dbo.TblItems.ItemID," & val(DcbUnit.BoundText) & " ," & IIf(SystemOptions.AllowLastPrice, 1, 0) & "),"
                Else
                    StrSQL = StrSQL + " Price = " & TxtPrice & ","
                End If
                StrSQL = StrSQL + "      dbo.TblItems.ItemID, 0 TableID, N'" & DcbUnit.Text & "' UnitName, '' UnitNamee, dbo.TblItems.ItemCode, dbo.TblItems.ItemName,"
                StrSQL = StrSQL + "      dbo.TblItems.ItemNamee , dbo.TblItems.fullcode"
                StrSQL = StrSQL + "  FROM     dbo.TblItems  "
                
                StrSQL = StrSQL + " Where (dbo.TblItems.ItemID =        " & val(DcboItemID1.BoundText) & ")"
                
                Set Rs3 = New ADODB.Recordset

                Rs3.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText
            End If
        Else
            
            Rs3.MoveFirst
        End If
        
        If Not Rs3.EOF Then
        With FG
'        If .Rows = 0 Then
'        .Rows = .Rows + 1
'        End If
        LngNewRow = ModFgLib.SetFgForNewRow(FG, FG.ColIndex("ItemID"))
        'k = .Rows
        .Rows = LngNewRow + Rs3.RecordCount
         
        Dim ForUnit As Double
        Dim MethodCalc As Double
        Dim lowering  As Double
        Dim Totallowering  As Double
        Dim increase  As Double
        Dim Qty As Double
        k = LngNewRow
        
        
        For i = k To .Rows - 1
        .TextMatrix(i, .ColIndex("FlgX")) = IIf(IsNull(Rs3("PartItemQty").value), 0, Rs3("PartItemQty").value)
        .TextMatrix(i, .ColIndex("Ser")) = i
        .TextMatrix(i, .ColIndex("isReplaced")) = IIf(IsNull(Rs3("isReplaced").value), "", Rs3("isReplaced").value)
        .TextMatrix(i, .ColIndex("ItemID")) = IIf(IsNull(Rs3("PartItemID").value), 0, Rs3("PartItemID").value)
        .TextMatrix(i, .ColIndex("itemcode")) = IIf(IsNull(Rs3("Fullcode").value), "", Rs3("Fullcode").value)
        .TextMatrix(i, .ColIndex("Fullcode")) = IIf(IsNull(Rs3("Fullcode").value), "", (Rs3("Fullcode").value))


        .TextMatrix(i, .ColIndex("UnitID")) = IIf(IsNull(Rs3("Unitid").value), 0, Rs3("Unitid").value)
         If mLineID = 0 Then
            .TextMatrix(i, .ColIndex("LineID")) = mNewId
        Else
            .TextMatrix(i, .ColIndex("LineID")) = mLineID
        End If

        .TextMatrix(i, .ColIndex("ItemID2")) = DcboItemID1.BoundText
        .TextMatrix(i, .ColIndex("ItemName2")) = DcboItemID1.Text
        .TextMatrix(i, .ColIndex("ItemCode2")) = Trim$(TxtAttachedItemCode.Text)
       
        
        
        .TextMatrix(i, .ColIndex("ForUnit")) = IIf(IsNull(Rs3("ForUnit").value), 0, Rs3("ForUnit").value)
        .TextMatrix(i, .ColIndex("MethodCalc")) = IIf(IsNull(Rs3("MethodCalc").value), 0, Rs3("MethodCalc").value)
        
        
        .TextMatrix(i, .ColIndex("lowering")) = IIf(IsNull(Rs3("lowering").value), 0, Rs3("lowering").value)
        .TextMatrix(i, .ColIndex("Increase")) = IIf(IsNull(Rs3("increase").value), 0, Rs3("increase").value)
        .TextMatrix(i, .ColIndex("PartItemQty")) = IIf(IsNull(Rs3("PartItemQty").value), 0, Rs3("PartItemQty").value)
        
          

                  ForUnit = IIf(IsNull(Rs3("ForUnit").value), 1, Rs3("ForUnit").value)
          MethodCalc = IIf(IsNull(Rs3("MethodCalc").value), 1, Rs3("MethodCalc").value)
          lowering = IIf(IsNull(Rs3("lowering").value), 0, Rs3("lowering").value)
          Totallowering = Totallowering + lowering
          increase = IIf(IsNull(Rs3("increase").value), 0, Rs3("increase").value)
          Qty = IIf(IsNull(Rs3("PartItemQty").value), 0, Rs3("PartItemQty").value)
        
          If ForUnit = 0 Then ForUnit = 1
        If MethodCalc = 1 Then 'ﬂ„Ì…
            
        ElseIf MethodCalc = 2 Then '⁄—÷
          Qty = ((val(txtwidtj) / ForUnit) * Qty) - lowering
        
        ElseIf MethodCalc = 3 Then 'ÿÊ·
        Qty = ((val(txthight) / ForUnit) * Qty) - lowering
        
        
        
         ElseIf MethodCalc = 4 Then 'ÿÊ·+⁄—÷
         Qty = ((val(txtwidtj) + val(txthight)) / ForUnit * Qty) - lowering
          ElseIf MethodCalc = 5 Then 'ÿÊ·*⁄—÷
                   Qty = ((val(txtwidtj) * val(txthight)) / ForUnit * Qty) - lowering
                   
        
                 ElseIf MethodCalc = 6 Then ' «·ÿÊ· ·ﬂ· ⁄—÷
                    
                  Qty = ((val(txthight) / ForUnit) - lowering) * Qty * val(txtwidtj)
                     Dim ff As Double
                    
                      .TextMatrix(i, .ColIndex("FlgX")) = Round(Qty, 2)
                        .TextMatrix(i, .ColIndex("Qty")) = Round(val(.TextMatrix(i, .ColIndex("FlgX"))) * val(txtQty1), 2)
                   
                     ElseIf MethodCalc = 7 Then ' «·⁄—÷ ·ﬂ· ÿÊ·
                    
                     Qty = ((val(txtwidtj) / ForUnit * Qty) * val(txthight)) - lowering   ' ((val(mwidtj) +  / ForUnit * Qty) - lowering
                      .TextMatrix(i, .ColIndex("FlgX")) = Round(Qty, 2)
                        .TextMatrix(i, .ColIndex("Qty")) = Round(val(.TextMatrix(i, .ColIndex("FlgX"))) * val(txtQty1), 2)

                     ElseIf MethodCalc = 8 Then '  * «·«— ›«⁄ *«·⁄—÷ * ÿÊ·
                     Qty = (((val(txtwidtj) * val(txthight) * val(txtLength))) / ForUnit * Qty) - lowering
                        
                        .TextMatrix(i, .ColIndex("FlgX")) = Round(Qty, 2)
                        .TextMatrix(i, .ColIndex("Qty")) = Round(val(.TextMatrix(i, .ColIndex("FlgX"))) * val(txtQty1), 2)
                        
                        
            ElseIf MethodCalc = 9 Then 'ÿÊ·+⁄—÷
                    Qty = (val(txthight) * 3.14 * ((val(txtDiameter) / 2) ^ 2) / ForUnit * Qty) - lowering
                        
            ElseIf MethodCalc = 10 Then 'ÿÊ·+⁄—÷
                    Qty = (((val(txtwidtj) * val(txthight) * val(txtthickness))) / ForUnit * Qty) - lowering
            ElseIf MethodCalc = 11 Then 'ÿÊ·+⁄—÷
                    Qty = (val(txthight) * 3.14 * ((val(txtDO) - val(txtDI))) / ForUnit * Qty) - lowering
                        
                End If
        If MethodCalc <> 1 Then
          .TextMatrix(i, .ColIndex("FlgX")) = Round(Qty, 2)
        .TextMatrix(i, .ColIndex("Qty")) = Round(val(.TextMatrix(i, .ColIndex("FlgX"))) * val(txtQty1.Text), 2)
        Else
        .TextMatrix(i, .ColIndex("FlgX")) = Qty
        .TextMatrix(i, .ColIndex("Qty")) = val(.TextMatrix(i, .ColIndex("FlgX"))) * val(txtQty1.Text)
        
        End If
        
        If SystemOptions.UserInterface = ArabicInterface Then
        .TextMatrix(i, .ColIndex("itemname")) = IIf(IsNull(Rs3("ItemName").value), "", Rs3("ItemName").value)
        .TextMatrix(i, .ColIndex("unitname")) = IIf(IsNull(Rs3("unitname").value), "", Rs3("unitname").value)
        Else
        .TextMatrix(i, .ColIndex("unitname")) = IIf(IsNull(Rs3("UnitNamee").value), "", Rs3("UnitNamee").value)
        .TextMatrix(i, .ColIndex("itemname")) = IIf(IsNull(Rs3("ItemNamee").value), "", Rs3("ItemNamee").value)
        End If
'         mPrice = GetItemPrice(.TextMatrix(i, .ColIndex("ItemID")), , val(.TextMatrix(i, .ColIndex("UnitID"))))
        .TextMatrix(i, .ColIndex("Price")) = Rs3!Price & ""
        .TextMatrix(i, .ColIndex("Total")) = val(Rs3!Price & "") * val(.TextMatrix(i, .ColIndex("Qty")))
'        .TextMatrix(i, .ColIndex("Total")) = mPrice * val(.TextMatrix(i, .ColIndex("Qty")))
        
            If val(.TextMatrix(i, .ColIndex("ItemID2"))) = val(.TextMatrix(i, .ColIndex("ItemID"))) Then
              .TextMatrix(i, .ColIndex("FlgX")) = IIf(IsNull(Rs3("PartItemQty").value), 0, Rs3("PartItemQty").value)
                    .TextMatrix(i, .ColIndex("Qty")) = IIf(IsNull(Rs3("PartItemQty").value), 0, Rs3("PartItemQty").value)
                       .TextMatrix(i, .ColIndex("Total")) = val(.TextMatrix(i, .ColIndex("Price"))) * val(.TextMatrix(i, .ColIndex("Qty")))
          End If
        
        Rs3.MoveNext
        .TextMatrix(i, .ColIndex("TepQty")) = .TextMatrix(i, .ColIndex("FlgX"))
      '  CalcTotal i
        If val(.TextMatrix(i, .ColIndex("Qty"))) = 0 Then
        .TextMatrix(i, .ColIndex("TepQty")) = .TextMatrix(i, .ColIndex("FlgX"))
        End If
   
    
                        If val(.TextMatrix(i, .ColIndex("FlgX"))) <> 0 Then
                            If val(.TextMatrix(i, .ColIndex("ItemID"))) = val(.TextMatrix(i, .ColIndex("ItemID2"))) Then
                                .TextMatrix(i, .ColIndex("Qty")) = val(txtQty1.Text)
                            Else
                                .TextMatrix(i, .ColIndex("Qty")) = val(.TextMatrix(i, .ColIndex("FlgX"))) * IIf(val(txtQty1.Text) = 0, 1, val(txtQty1.Text))
                            End If
                            .TextMatrix(i, .ColIndex("Total")) = val(.TextMatrix(i, .ColIndex("Price"))) * val(.TextMatrix(i, .ColIndex("Qty")))
                            
                        End If
                        If val(.TextMatrix(i, .ColIndex("ItemID"))) = val(.TextMatrix(i, .ColIndex("ItemID2"))) Then FillGrid = True: Exit Function
                    
        Next i
        If SystemOptions.IsMultiItemsInCompItem Then
            ReLineGrid
        End If
        End With
        End If
        FillGrid = True
End Function


Sub FillGrid2(Optional ByVal mItemIDGG As Long = 0, Optional ByVal mFild As String = "ItemID2", Optional ByVal mFildName As String = "ItemName2")
' FG.Clear flexClearScrollable, flexClearEverything
'            FG.Rows = 1
    If mItemIDGG = 0 Then mItemIDGG = DcboItemID2.BoundText
    For l = 1 To FG.Rows - 1
        If l > FG.Rows - 1 Then Exit For
        If FG.TextMatrix(l, FG.ColIndex(mFild)) = mItemIDGG Or FG.TextMatrix(l, FG.ColIndex("ItemID")) = "" Then
            FG.RemoveItem l
            l = l - 1
        End If
       
        
    Next

Dim StrSQL As String
Dim i As Integer
Dim k As Integer
Dim Rs3 As ADODB.Recordset
Set Rs3 = New ADODB.Recordset
                StrSQL = " SELECT       ForUnit   ,MethodCalc,   TblItemsParts.lowering ,TblItemsParts.increase, dbo.TblItemsParts.Unitid, dbo.TblItemsParts.PartItemPrice, dbo.TblItemsParts.PartItemQty, dbo.TblItemsParts.PartItemID, "
                StrSQL = StrSQL + "      dbo.TblItemsParts.ItemID, dbo.TblItemsParts.TableID, dbo.TblUnites.UnitName, dbo.TblUnites.UnitNamee, dbo.TblItems.ItemCode, dbo.TblItems.ItemName,"
                StrSQL = StrSQL + "      dbo.TblItems.ItemNamee , dbo.TblItems.fullcode"
                StrSQL = StrSQL + "  FROM         dbo.TblItemsParts INNER JOIN"
                StrSQL = StrSQL + "      dbo.TblUnites ON dbo.TblItemsParts.Unitid = dbo.TblUnites.UnitID RIGHT OUTER JOIN"
                StrSQL = StrSQL + "      dbo.TblItems ON dbo.TblItemsParts.PartItemID = dbo.TblItems.ItemID"
                StrSQL = StrSQL + " Where (dbo.TblItemsParts.ItemID = " & val(mItemIDGG) & ")"
                StrSQL = StrSQL + " ORDER BY dbo.TblItemsParts.TableID"
        Rs3.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText
        If Rs3.RecordCount > 0 Then
        Rs3.MoveFirst
        With FG
        If .Rows = 0 Then
        .Rows = .Rows + 1
        End If
        k = .Rows
        If .TextMatrix(k - 1, .ColIndex("ItemID")) = "" Then
            .Rows = .Rows - 1
            k = k - 1
        End If

        .Rows = .Rows + Rs3.RecordCount
        Dim ForUnit As Double
        Dim MethodCalc As Double
        Dim lowering  As Double
        Dim increase As Double
        
        Dim Qty As Double
        
        
        For i = k To .Rows - 1
        .TextMatrix(i, .ColIndex("FlgX")) = IIf(IsNull(Rs3("PartItemQty").value), 0, Rs3("PartItemQty").value)
        .TextMatrix(i, .ColIndex("Ser")) = i
        .TextMatrix(i, .ColIndex("ItemID")) = IIf(IsNull(Rs3("PartItemID").value), 0, Rs3("PartItemID").value)
        .TextMatrix(i, .ColIndex("itemcode")) = IIf(IsNull(Rs3("Fullcode").value), "", Rs3("Fullcode").value)
        .TextMatrix(i, .ColIndex("Fullcode")) = IIf(IsNull(Rs3("Fullcode").value), "", (Rs3("Fullcode").value))


        .TextMatrix(i, .ColIndex("UnitID")) = IIf(IsNull(Rs3("Unitid").value), 0, Rs3("Unitid").value)
        
        .TextMatrix(i, .ColIndex("ItemCode2")) = Trim$(TxtAttachedItemCode2.Text)
        .TextMatrix(i, .ColIndex(mFild)) = mItemIDGG
        .TextMatrix(i, .ColIndex(mFildName)) = IIf(mFildName = "ItemName2", DcboItemID2.Text, DcboItemID5.Text)
        
        
        
        .TextMatrix(i, .ColIndex("ForUnit")) = IIf(IsNull(Rs3("ForUnit").value), 0, Rs3("ForUnit").value)
        .TextMatrix(i, .ColIndex("MethodCalc")) = IIf(IsNull(Rs3("MethodCalc").value), 0, Rs3("MethodCalc").value)
        .TextMatrix(i, .ColIndex("lowering")) = IIf(IsNull(Rs3("lowering").value), 0, Rs3("lowering").value)
        .TextMatrix(i, .ColIndex("Increase")) = IIf(IsNull(Rs3("increase").value), 0, Rs3("increase").value)
        
        .TextMatrix(i, .ColIndex("PartItemQty")) = IIf(IsNull(Rs3("PartItemQty").value), 0, Rs3("PartItemQty").value)
        
          

          ForUnit = IIf(IsNull(Rs3("ForUnit").value), 1, Rs3("ForUnit").value)
          MethodCalc = IIf(IsNull(Rs3("MethodCalc").value), 1, Rs3("MethodCalc").value)
          lowering = IIf(IsNull(Rs3("lowering").value), 0, Rs3("lowering").value)
          increase = IIf(IsNull(Rs3("increase").value), 0, Rs3("increase").value)
          
          Qty = IIf(IsNull(Rs3("PartItemQty").value), 0, Rs3("PartItemQty").value)
        If MethodCalc = 1 Then 'ﬂ„Ì…
            
        ElseIf MethodCalc = 2 Then '⁄—÷
          Qty = ((val(txtwidtj) / ForUnit) * Qty) - lowering
        
        ElseIf MethodCalc = 3 Then 'ÿÊ·
        Qty = ((val(txthight) / ForUnit) * Qty) - lowering
        
        
        
         ElseIf MethodCalc = 4 Then 'ÿÊ·+⁄—÷
         Qty = ((val(txtwidtj) + val(txthight)) / ForUnit * Qty) - lowering
          ElseIf MethodCalc = 5 Then 'ÿÊ·*⁄—÷
                   Qty = ((val(txtwidtj) * val(txthight)) / ForUnit * Qty) - lowering
                   
        End If
        If MethodCalc <> 1 Then
          .TextMatrix(i, .ColIndex("FlgX")) = Round(Qty, 2)
        .TextMatrix(i, .ColIndex("Qty")) = Round(val(.TextMatrix(i, .ColIndex("FlgX"))) * val(txtQty1.Text), 2)
        Else
        .TextMatrix(i, .ColIndex("FlgX")) = Qty
        .TextMatrix(i, .ColIndex("Qty")) = val(.TextMatrix(i, .ColIndex("FlgX"))) * val(txtQty1.Text)
        
        End If
        
        If SystemOptions.UserInterface = ArabicInterface Then
        .TextMatrix(i, .ColIndex("itemname")) = IIf(IsNull(Rs3("ItemName").value), "", Rs3("ItemName").value)
        .TextMatrix(i, .ColIndex("unitname")) = IIf(IsNull(Rs3("unitname").value), "", Rs3("unitname").value)
        Else
        .TextMatrix(i, .ColIndex("unitname")) = IIf(IsNull(Rs3("UnitNamee").value), "", Rs3("UnitNamee").value)
        .TextMatrix(i, .ColIndex("itemname")) = IIf(IsNull(Rs3("ItemNamee").value), "", Rs3("ItemNamee").value)
        End If
         mPrice = GetItemPrice(.TextMatrix(i, .ColIndex("ItemID")), , val(.TextMatrix(i, .ColIndex("UnitID"))))
        .TextMatrix(i, .ColIndex("Price")) = mPrice
        .TextMatrix(i, .ColIndex("Total")) = mPrice * val(.TextMatrix(i, .ColIndex("Qty")))
        
        Rs3.MoveNext
        .TextMatrix(i, .ColIndex("TepQty")) = .TextMatrix(i, .ColIndex("FlgX"))
        CalcTotal i
        ReLineGrid
        Next i
        End With
        End If
End Sub
Private Sub DBCboClientName_Click(Area As Integer)
 If val(DBCboClientName.BoundText) = 0 Then Exit Sub

    Dim EmpCode  As String
 
    GetTblCustemersCode , , DBCboClientName.BoundText, EmpCode
    Me.TxtSearchCode.Text = EmpCode
    Me.TxtSearchCode2.Text = EmpCode
End Sub

Private Sub DBCboClientName_KeyUp(KeyCode As Integer, _
                                  Shift As Integer)

    If KeyCode = vbKeyF3 Then
        FrmCustemerSearch.SearchType = 2020
        FrmCustemerSearch.show vbModal
    End If
    
    If KeyCode = vbKeyF5 Then
        Dim Dcombos As ClsDataCombos
        
        Set Dcombos = New ClsDataCombos
        Dcombos.GetCustomersSuppliers 0, Me.DBCboClientName, True
 
    End If
        
End Sub



Private Sub DcboItemID1_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF3 Then
        FrmItemSearch.RetrunType = 2020
        FrmItemSearch.show vbModal
        
    End If
End Sub


Private Sub CalcTotalNet(Optional ByVal mItemId As Long = 0)
    Dim i As Integer
    Dim IntCounter As Integer
    'If mItemId <> 0 Then

       
        With FG
           'If SystemOptions.IsMultiItemsInCompItem Then
                txtTotal = val(TxtPrice) * val(txtQty1)
                txtTotalAdd = 0
                txtNet = 0
                If mItemId = 0 Then mItemId = val(DcboItemID1.BoundText)
        
                For i = .FixedRows To .Rows - 1
                    
                    If val(.TextMatrix(i, .ColIndex("itemId2"))) = mItemId Then
                        If (.ValueMatrix(i, .ColIndex("IsAdd"))) Then
                            txtTotalAdd = val(txtTotalAdd) + val(.TextMatrix(i, .ColIndex("Total")))
                        End If
                        If val(.ValueMatrix(i, .ColIndex("OldPrice"))) <> 0 Then
                            txtTotalDisc = val(txtTotalDisc) + val(.TextMatrix(i, .ColIndex("OldPrice")))
                        End If
                    End If
         
                Next i
                txtNet = val(txtTotal) + val(txtTotalAdd) - val(txtTotalDisc)
           'End If
            CalCulteVAT 3
    End With

End Sub




Private Sub ReLineGrid(Optional ByVal mRow As Long = 0, Optional ByVal mIsEdit As Boolean = False, Optional ByVal isFromDeleteRow As Boolean = False, Optional ByVal isFromCalcRow As Boolean = False, Optional ByVal isFromGeometric As Boolean = False)
    Dim i As Integer
    Dim IntCounter As Integer
    Dim MethodCalc As Double
    Dim mQtyGrid As Double
    Dim mItemNo As Long, mUnitNo As Long, mLineNo As Long
    Dim mwidtj  As Double, mhight As Double
    Dim mAdd As Boolean
    Dim mLength As Double
    Dim mthickness As Double
    Dim mDI As Double
    Dim mDiameter As Double
    Dim mDO As Double
    Dim Totallowering As Double
    Dim lowering As Double
    Dim ForUnit  As Long
    Dim increase  As Double
    Dim Qty  As Double
    If mRow <> 0 And mRow <= FG2.Rows - 1 Then
       ' mQtyGrid = val(FG2.TextMatrix(mRow, FG2.ColIndex("Qty")))
        mLineNo = val(FG2.TextMatrix(mRow, FG2.ColIndex("LineID")))
        mUnitNo = val(FG2.TextMatrix(mRow, FG2.ColIndex("UnitID")))
        mItemNo = val(FG2.TextMatrix(mRow, FG2.ColIndex("ItemID")))
        mwidtj = val(FG2.TextMatrix(mRow, FG2.ColIndex("widtj")))
        mhight = val(FG2.TextMatrix(mRow, FG2.ColIndex("hight")))
        mLength = val(FG2.TextMatrix(mRow, FG2.ColIndex("Length")))
        mQtyGrid = val(FG2.TextMatrix(mRow, FG2.ColIndex("Qty")))
       
    Else
        mLineNo = 0
        mItemNo = val(DcboItemID1.BoundText)
        mUnitNo = val(DcbUnit.BoundText)
        mQtyGrid = val(txtQty1)
        mwidtj = val(txtwidtj)
        mhight = val(txthight)
        mLength = val(txtLength)
        mDiameter = val(txtDiameter)
        mDO = val(txtDO)
        mDI = val(txtDI)
        mthickness = val(txtthickness)
    End If
    
    
    
    If isFromGeometric Then
         mwidtj = val(txtwidtj2)
        mhight = val(txthight2)
        mLength = val(txtLength2)
        mDiameter = val(txtDiameter2)
        mDO = val(txtDO2)
        mDI = val(txtDI2)
        mthickness = val(txtthickness2)
    End If
    txtTotalAdd = 0
    With FG

        For i = .FixedRows To .Rows - 1
        
        
            If isFromGeometric Then
                mwidtj = IIf(val(.TextMatrix(i, .ColIndex("widtj"))) = 0, val(txtwidtj), val(.TextMatrix(i, .ColIndex("widtj"))))
                mhight = IIf(val(.TextMatrix(i, .ColIndex("hight"))) = 0, val(txthight), val(.TextMatrix(i, .ColIndex("hight"))))
                mLength = IIf(val(.TextMatrix(i, .ColIndex("Length"))) = 0, val(txtLength), val(.TextMatrix(i, .ColIndex("Length"))))
                mDiameter = IIf(val(.TextMatrix(i, .ColIndex("Diameter"))) = 0, val(txtDiameter), val(.TextMatrix(i, .ColIndex("Diameter"))))
                mDO = IIf(val(.TextMatrix(i, .ColIndex("DO"))) = 0, val(txtDO), val(.TextMatrix(i, .ColIndex("DO"))))
                mDI = IIf(val(.TextMatrix(i, .ColIndex("DI"))) = 0, val(txtDI), val(.TextMatrix(i, .ColIndex("DI"))))
                mthickness = IIf(val(.TextMatrix(i, .ColIndex("thickness"))) = 0, val(txtthickness), val(.TextMatrix(i, .ColIndex("thickness"))))
            End If
            
            
            If mItemNo = val(.TextMatrix(i, .ColIndex("ItemID2"))) And (mLineNo = val(.TextMatrix(i, .ColIndex("LineID")))) And mRow <> 0 Then
                lowering = val(.TextMatrix(i, .ColIndex("lowering")))
                Totallowering = Totallowering + lowering
            End If
         
            If i = 12 Then
                i = i
            End If
            If mIsEdit Then
                If val(.TextMatrix(i, .ColIndex("ItemID"))) = val(.TextMatrix(i, .ColIndex("ItemID2"))) Then
                    .TextMatrix(i, .ColIndex("FlgX")) = mQtyGrid
                    .TextMatrix(i, .ColIndex("Qty")) = mQtyGrid
                       .TextMatrix(i, .ColIndex("Total")) = val(.TextMatrix(i, .ColIndex("Price"))) * val(.TextMatrix(i, .ColIndex("Qty")))
                        GoTo NextRow
                End If
             
             End If
            
                    
            If mItemNo = val(.TextMatrix(i, .ColIndex("ItemID2"))) And (mLineNo = val(.TextMatrix(i, .ColIndex("LineID"))) Or mLineNo = 0) Then
               ' if .TextMatrix(i, .ColIndex("IsAdd")) <> "" then
               
                    
                    IntCounter = IntCounter + 1
      
                    .TextMatrix(i, .ColIndex("Ser")) = IntCounter
                     mAdd = CBool(.ValueMatrix(i, .ColIndex("IsAdd")))
                    If mAdd Then GoTo NextRow
            
                If .TextMatrix(i, .ColIndex("itemcode")) <> "" Then
                  
                    If SystemOptions.AllowChangManualQtyMix = False Then
                        If val(.TextMatrix(i, .ColIndex("FlgX"))) <> 0 Then
                            .TextMatrix(i, .ColIndex("Qty")) = val(.TextMatrix(i, .ColIndex("FlgX"))) * IIf(mQtyGrid = 0, 1, mQtyGrid)
                            
                        End If
                        
                    End If
                    If mIsEdit And Not isFromCalcRow Then
                        If val(.TextMatrix(i, .ColIndex("FlgX"))) <> 0 Then
                            If val(.TextMatrix(i, .ColIndex("ItemID"))) = val(.TextMatrix(i, .ColIndex("ItemID2"))) Then
                                .TextMatrix(i, .ColIndex("Qty")) = mQtyGrid
                            Else
                                .TextMatrix(i, .ColIndex("Qty")) = val(.TextMatrix(i, .ColIndex("FlgX"))) * IIf(mQtyGrid = 0, 1, mQtyGrid)
                            End If
                            .TextMatrix(i, .ColIndex("Total")) = val(.TextMatrix(i, .ColIndex("Price"))) * val(.TextMatrix(i, .ColIndex("Qty")))
                            GoTo NextRow
                        End If
                        
                    End If
                      
                        ForUnit = .TextMatrix(i, .ColIndex("ForUnit"))
                        MethodCalc = val(.TextMatrix(i, .ColIndex("MethodCalc")))
                        lowering = val(.TextMatrix(i, .ColIndex("lowering")))
                      
                        increase = val(.TextMatrix(i, .ColIndex("increase")))
                        Qty = val(.TextMatrix(i, .ColIndex("PartItemQty")))
                        If val(ForUnit) = 0 Then ForUnit = 1
                        If Trim(ForUnit) = "" And MethodCalc = 0 And MethodCalc = 0 And lowering = 0 Then Exit Sub
                     
                    If MethodCalc = 1 Then 'ﬂ„Ì…
                        
                    ElseIf MethodCalc = 2 Then '⁄—÷
                      Qty = ((val(mwidtj) / ForUnit) * Qty) - lowering
                        .TextMatrix(i, .ColIndex("FlgX")) = Round(Qty, 2)
                        .TextMatrix(i, .ColIndex("Qty")) = Round(val(.TextMatrix(i, .ColIndex("FlgX"))) * val(mQtyGrid), 2)
                    ElseIf MethodCalc = 3 Then 'ÿÊ·
                    Qty = ((val(mhight) / ForUnit) * Qty) - lowering
                    
                        .TextMatrix(i, .ColIndex("FlgX")) = Round(Qty, 2)
                        .TextMatrix(i, .ColIndex("Qty")) = Round(val(.TextMatrix(i, .ColIndex("FlgX"))) * val(mQtyGrid), 2)
                    
                     ElseIf MethodCalc = 4 Then 'ÿÊ·+⁄—÷
                     Qty = ((val(mwidtj) + val(mhight)) / ForUnit * Qty) - lowering
                      .TextMatrix(i, .ColIndex("FlgX")) = Round(Qty, 2)
                        .TextMatrix(i, .ColIndex("Qty")) = Round(val(.TextMatrix(i, .ColIndex("FlgX"))) * val(mQtyGrid), 2)
                      ElseIf MethodCalc = 5 Then 'ÿÊ·*⁄—÷
                               Qty = ((val(mwidtj) * val(mhight)) / ForUnit * Qty) - lowering
                                .TextMatrix(i, .ColIndex("FlgX")) = Round(Qty, 2)
                            .TextMatrix(i, .ColIndex("Qty")) = Round(val(.TextMatrix(i, .ColIndex("FlgX"))) * val(mQtyGrid), 2)
                   ElseIf MethodCalc = 6 Then ' «·ÿÊ· ·ﬂ· ⁄—÷
                        Qty = ((val(txthight) / ForUnit) - lowering) * Qty * val(mwidtj)
                  
                      .TextMatrix(i, .ColIndex("FlgX")) = Round(Qty, 2)
                        .TextMatrix(i, .ColIndex("Qty")) = Round(val(.TextMatrix(i, .ColIndex("FlgX"))) * val(mQtyGrid), 2)
                     ElseIf MethodCalc = 7 Then ' «·⁄—÷ ·ﬂ· ÿÊ·
                    
                     Qty = ((val(mwidtj) / ForUnit * Qty) * val(mhight)) - lowering   ' ((val(mwidtj) +  / ForUnit * Qty) - lowering
                      .TextMatrix(i, .ColIndex("FlgX")) = Round(Qty, 2)
                        .TextMatrix(i, .ColIndex("Qty")) = Round(val(.TextMatrix(i, .ColIndex("FlgX"))) * val(mQtyGrid), 2)
                     
                     ElseIf MethodCalc = 8 Then ' «·«— ›«⁄ * «·⁄—÷ * ÿÊ·
                    Qty = (((val(mwidtj) * val(mhight) * val(mLength))) / ForUnit * Qty) - lowering
                     
                      .TextMatrix(i, .ColIndex("FlgX")) = Round(Qty, 2)
                        .TextMatrix(i, .ColIndex("Qty")) = Round(val(.TextMatrix(i, .ColIndex("FlgX"))) * val(mQtyGrid), 2)
                                           
                    ElseIf MethodCalc = 9 Then 'ÿÊ·+⁄—÷
                            Qty = (val(mhight) * 3.14 * ((val(mDiameter) / 2) ^ 2) / ForUnit * Qty) - lowering
                                
                    ElseIf MethodCalc = 10 Then 'ÿÊ·+⁄—÷
                            Qty = (((val(mwidtj) * val(mhight) * val(mthickness))) / ForUnit * Qty) - lowering
                    ElseIf MethodCalc = 11 Then 'ÿÊ·+⁄—÷
                        Qty = (val(mhight) * 3.14 * ((val(mDO) - val(mDI))) / ForUnit * Qty) - lowering
                                   
                      End If
                    If MethodCalc <> 1 Then
                        If val(.TextMatrix(i, .ColIndex("Qty"))) = 0 Then
                          .TextMatrix(i, .ColIndex("FlgX")) = Round(Qty, 2)
                        .TextMatrix(i, .ColIndex("Qty")) = Round(val(.TextMatrix(i, .ColIndex("FlgX"))) * val(mQtyGrid), 2)
                        End If
                            
                    
                    
                    Else
                    .TextMatrix(i, .ColIndex("FlgX")) = Qty
                    .TextMatrix(i, .ColIndex("Qty")) = val(.TextMatrix(i, .ColIndex("FlgX"))) * val(mQtyGrid)
                            
                             .TextMatrix(i, .ColIndex("Total")) = val(.TextMatrix(i, .ColIndex("Price"))) * val(.TextMatrix(i, .ColIndex("Qty")))
                          
                        End If
                End If
NextRow:
                  If (.ValueMatrix(i, .ColIndex("IsAdd"))) Then
                        txtTotalAdd = val(txtTotalAdd) + val(.TextMatrix(i, .ColIndex("Total")))
                  End If
            Else
            
                i = i
                '.TextMatrix(i, .ColIndex("Qty")) = Round(val(.TextMatrix(i, .ColIndex("FlgX"))) * val(mQtyGrid), 2)
            End If
        Next i
        If mRow <> 0 Then
            
'            If val(FG2.TextMatrix(mRow, FG2.ColIndex("BuiltinItemID"))) <> 0 Then
'                FG2.TextMatrix(mRow, FG2.ColIndex("lowering")) = Totallowering
'            End If
        End If
        
        
        txtNet = val(txtTotal) + val(txtTotalAdd) - val(txtTotalDisc)
    End With
    If SystemOptions.IsMultiItemsInCompItem Then
        FG.Select 1, FG.ColIndex("LineID")
    End If
FG.Sort = flexSortGenericAscending

End Sub

Private Sub AddNewFgRow(Optional ByVal mItemIDGG As Long = 0, Optional ByVal mFild As String = "ItemID2", Optional ByVal mFildName As String = "ItemName2", Optional ByVal mLineID As Long = 0)

Dim i As Long

' FG.Clear flexClearScrollable, flexClearEverything
'            FG.Rows = 1
    Dim mItemNameG As String
    Dim mUnitNameG As String
    Dim mUnitIDG As Long
    Dim mQtyG As Double
    If mItemIDGG = 0 Then mItemIDGG = val(DcboItemID2.BoundText)
    If mFildName = "ItemName2" Then
        mItemNameG = DcboItemID2.Text
        mUnitIDG = val(DcbUnit2.BoundText)
        mUnitNameG = DcbUnit2.Text
        mQtyG = val(txtQty)
    Else
        mItemNameG = DcboItemID5.Text
        mUnitIDG = val(DcbUnit5.BoundText)
        mUnitNameG = DcbUnit5.Text
        mQtyG = val(txtQty5)
    End If
    Dim Msg As String
    Dim LngFindRow As Long
    Dim LngNewRow As Long
    Dim mUnitName As String
                Dim mUnitId As Long
Dim mPrice As Double

    'With Me.FG
      '  LngFindRow = .FindRow(val(Me.DcboItems.BoundText), .FixedRows, .ColIndex("ItemID"), False, True)

      '  If LngFindRow <> -1 Then
      '      Msg = "Â–« «·’‰› „ÊÃÊœ ›⁄·« ...!!!"
      '      MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
      '      .SetFocus
      '      Exit Sub
      '  End If

    'End With
    If mFildName = "ItemName2" Then
    If SystemOptions.IsMultiItemsInCompItem Then
        If FG2.Rows > 1 And FG.Rows > 1 And Trim(TxtMaxNo2) <> "" Then Exit Sub
        
       
        FillGridItemType val(DcboItemID1.BoundText), DcboItemID1.Text, Trim$(TxtAttachedItemCode.Text), 1, val(DcbUnit.BoundText), DcbUnit.Text, val(txtQty1), val(TxtPrice), val(XPCboGroup.BoundText), XPCboGroup.Text, mLineID
      
                Dim s As String
        
        If Trim(TxtMaxNo2) <> "" And val(DcboItemID1.BoundText) <> 0 Then
            If FG2.Rows > 1 And FG.Rows > 1 Then Exit Sub
            s = " SELECT TblItems.ItemName,TblItems.FullCode itemcode, tu.UnitName,TblDefComItemData.Qty,TblDefComItemData.Qty FlgX,LineID = 1, ItemId2 =" & val(DcboItemID1.BoundText) & ",ItemName2 =N'" & Trim(DcboItemID1.Text) & "',"
            s = s & " TblDefComItemData.cost,TblDefComItemData.Price,TblDefComItemData.Total,TblDefComItemData.UnitId,TblDefComItemData.ItemID"
            
            s = s & " FROM  TblDefComItemData INNER JOIN TblDefComItem ON TblDefComItem.ID = TblDefComItemData.IDDefCIT"
            
            s = s & " INNER JOIN TblItems ON TblItems.ItemID = TblDefComItemData.ItemID"
            s = s & " INNER JOIN TblUnites AS tu"
            s = s & " ON  tu.UnitID= TblDefComItemData.UnitID"
            s = s & " Where TblDefComItem.MaxNo = N'" & Trim(TxtMaxNo2) & "'"
            loadgrid s, FG, True, False
            Exit Sub
        Else
            If TXT_order_no = "" Then
                If Not FillGrid(mLineID) Then Exit Sub
            End If
        End If
        
    End If
   
     End If
        Dim isFound As Boolean
        If mItemNameG <> "" Then
            
            If SystemOptions.IsMultiItemsInCompItem Then
                        
               
                For i = 1 To FG.Rows - 1
                    If FG.ValueMatrix(i, FG.ColIndex("isReplaced")) And mNewId = val(FG.TextMatrix(i, FG.ColIndex("LineID"))) Then
                    '    DeleteRow i, True
                        LngNewRow = i
                    '    isFound = False
                    '    Exit For
                        
                        FG.TextMatrix(LngNewRow, FG.ColIndex("ItemID")) = mItemIDGG
                        FG.TextMatrix(LngNewRow, FG.ColIndex("itemcode")) = Trim$(Me.TxtAttachedItemCode2.Text)
                        FG.TextMatrix(LngNewRow, FG.ColIndex("Fullcode")) = Trim$(Me.TxtAttachedItemCode2.Text)


                        FG.TextMatrix(LngNewRow, FG.ColIndex("itemname")) = mItemNameG
                        FG.TextMatrix(LngNewRow, FG.ColIndex("OldPrice")) = val(FG.TextMatrix(LngNewRow, FG.ColIndex("Total")))
                        FG.TextMatrix(LngNewRow, FG.ColIndex("IsAdd")) = 1
                        mPrice = GetItemPrice(FG.TextMatrix(LngNewRow, FG.ColIndex("ItemID")), , val(FG.TextMatrix(i, FG.ColIndex("UnitID"))))
                       FG.TextMatrix(LngNewRow, FG.ColIndex("Price")) = mPrice
                       FG.TextMatrix(LngNewRow, FG.ColIndex("Total")) = mPrice * val(FG.TextMatrix(i, FG.ColIndex("Qty")))
                                           
                       ' mLineID = val(FG.TextMatrix(LngNewRow, FG.ColIndex("LineID")))
                        GoTo NextStep
                    
                    End If
                Next
            End If
           If Not isFound Then
                LngNewRow = ModFgLib.SetFgForNewRow(FG, FG.ColIndex("ItemID"))
            End If
          With FG
            .TextMatrix(LngNewRow, .ColIndex("LineID")) = mNewId
            .TextMatrix(LngNewRow, .ColIndex("ItemID")) = mItemIDGG
            .TextMatrix(LngNewRow, .ColIndex("itemcode")) = Trim$(Me.TxtAttachedItemCode2.Text)
            .TextMatrix(LngNewRow, .ColIndex("Fullcode")) = Trim$(Me.TxtAttachedItemCode2.Text)
            .TextMatrix(LngNewRow, .ColIndex("itemname")) = mItemNameG
            .TextMatrix(LngNewRow, .ColIndex("UnitID")) = mUnitIDG
            .TextMatrix(LngNewRow, .ColIndex("unitname")) = mUnitNameG
            If SystemOptions.AllowChangManualQtyMix = True Then
                .TextMatrix(LngNewRow, .ColIndex("Qty")) = mQtyG
            Else
                .TextMatrix(LngNewRow, .ColIndex("Qty")) = mQtyG * IIf(val(txtQty1.Text) = 0, 1, val(txtQty1.Text))
            End If
            .TextMatrix(LngNewRow, .ColIndex("FlgX")) = mQtyG
            .TextMatrix(LngNewRow, .ColIndex("TepQty")) = .TextMatrix(LngNewRow, .ColIndex("Qty"))
            .TextMatrix(LngNewRow, .ColIndex("IsAdd")) = 1
           
            mPrice = GetItemPrice(mItemIDGG, val(.TextMatrix(LngNewRow, .ColIndex("Qty"))), mUnitIDG)
            .TextMatrix(LngNewRow, .ColIndex("Price")) = mPrice
            .TextMatrix(LngNewRow, .ColIndex("Total")) = mPrice * val(.TextMatrix(LngNewRow, .ColIndex("Qty")))
            

                .TextMatrix(LngNewRow, .ColIndex("ItemID2")) = Me.DcboItemID1.BoundText
                .TextMatrix(LngNewRow, .ColIndex("itemcode2")) = Trim$(Me.TxtAttachedItemCode.Text)
                .TextMatrix(LngNewRow, .ColIndex("ItemName2")) = Me.DcboItemID1.Text
                
            
          ' End If
           ' .TextMatrix(LngNewRow, .ColIndex("ItemPrice")) = val(Me.TxtItemPrice(0).text)
            .AutoSize 0, .Cols - 1, False
        End With
    End If
   'Else
        With FG
'        .TextMatrix(LngNewRow, .ColIndex("IsAdd")) = 1
'        .TextMatrix(LngNewRow, .ColIndex("ItemID2")) = Me.DcboItemID1.BoundText
'        .TextMatrix(LngNewRow, .ColIndex("itemcode2")) = Trim$(Me.TxtAttachedItemCode.Text)
'        .TextMatrix(LngNewRow, .ColIndex("ItemName2")) = Me.DcboItemID1.Text
        End With
      ' FillGrid2
    
NextStep:
    
    If chkIsAdd.value = vbChecked Then
        For i = 1 To FG2.Rows - 1
            If val(FG.TextMatrix(LngNewRow, FG.ColIndex("ItemID2"))) = val(FG2.TextMatrix(i, FG2.ColIndex("ItemID"))) And val(FG.TextMatrix(LngNewRow, FG.ColIndex("LineID"))) = val(FG2.TextMatrix(i, FG2.ColIndex("LineID"))) Then
                txtTotalAdd = val(FG.TextMatrix(LngNewRow, FG.ColIndex("Total")))
                FG2.TextMatrix(i, FG2.ColIndex(mFild)) = val(FG.TextMatrix(LngNewRow, FG.ColIndex("ItemID")))
                FG2.TextMatrix(i, FG2.ColIndex("CountItem2")) = txtQty
                FG2.TextMatrix(i, FG2.ColIndex("CountItem5")) = txtQty5
               ' FG2.TextMatrix(i, FG2.ColIndex("LineID")) = val(FG.TextMatrix(LngNewRow, FG.ColIndex("LineID")))
                FG2.TextMatrix(i, FG2.ColIndex(mFildName)) = Trim(FG.TextMatrix(LngNewRow, FG.ColIndex("itemname")))
                FG2.TextMatrix(i, FG2.ColIndex("TotalAdd")) = val(FG.TextMatrix(LngNewRow, FG.ColIndex("Total")))
                FG2.TextMatrix(i, FG2.ColIndex("TotalDisc")) = val(FG2.TextMatrix(i, FG2.ColIndex("TotalDisc"))) + val(FG.TextMatrix(LngNewRow, FG.ColIndex("OldPrice")))
            End If
        Next
    End If
'
'If chkIsAdd Then
'    FillGridItemType DcboItemID2.BoundText, DcboItemID2.Text, Trim$(TxtAttachedItemCode2.Text), 2, DcbUnit2.BoundText, DcbUnit2.Text, val(txtQty), 0, XPCboGroup2.BoundText, XPCboGroup2.Text
'End If

    'Me.lbl(21).Caption = ModFgLib.GetItemsInFg(FG, FG.ColIndex("ItemID"))


If IsSaveWithOutMsg Or mLineID <> 0 Then
    If SystemOptions.IsMultiItemsInCompItem Then
        'ReLineGrid mLineID
    End If
Else
    For i = 1 To FG2.Rows - 1
        If SystemOptions.IsMultiItemsInCompItem Then
            ReLineGrid i
        Else
           ' ReLineGrid i,  True
        End If
    Next

End If
 If SystemOptions.IsMultiItemsInCompItem Then
        If DcboBuiltinItemID.Text <> "" Then
                LngNewRow = ModFgLib.SetFgForNewRow(FG, FG.ColIndex("ItemID"))
                FG.TextMatrix(LngNewRow, FG.ColIndex("ItemID")) = Me.DcboBuiltinItemID.BoundText
                FG.TextMatrix(LngNewRow, FG.ColIndex("itemcode")) = Trim$(Me.TxtAttachedItemCode3.Text)
                FG.TextMatrix(LngNewRow, FG.ColIndex("Fullcode")) = Trim$(Me.TxtAttachedItemCode3.Text)
    
    
                FG.TextMatrix(LngNewRow, FG.ColIndex("itemname")) = Me.DcboBuiltinItemID.Text
                FG.TextMatrix(LngNewRow, FG.ColIndex("OldPrice")) = val(FG.TextMatrix(LngNewRow, FG.ColIndex("Total")))
                FG.TextMatrix(LngNewRow, FG.ColIndex("IsAdd")) = 0
                mPrice = GetItemPrice(FG.TextMatrix(LngNewRow, FG.ColIndex("ItemID")), , val(FG.TextMatrix(i, FG.ColIndex("UnitID"))))
                GetDefaultItemUnit val(FG.TextMatrix(LngNewRow, FG.ColIndex("ItemID"))), mUnitId, mUnitName
                FG.TextMatrix(LngNewRow, FG.ColIndex("UnitID")) = mUnitId
                
                FG.TextMatrix(LngNewRow, FG.ColIndex("unitname")) = mUnitName
                If SystemOptions.AllowChangManualQtyMix = True Then
                    FG.TextMatrix(LngNewRow, FG.ColIndex("Qty")) = val(txthight) + 20
                Else
                    'FG.TextMatrix(LngNewRow, FG.ColIndex("Qty")) = 1 * IIf(val(txtQty1.Text) = 0, 1, val(txtQty1.Text))
                    FG.TextMatrix(LngNewRow, FG.ColIndex("Qty")) = val(txthight) + 20
                    
                End If
                
            
                FG.TextMatrix(LngNewRow, FG.ColIndex("FlgX")) = 1
              '  FG.TextMatrix(LngNewRow, FG.ColIndex("Qty")) = 1
                FG.TextMatrix(LngNewRow, FG.ColIndex("Price")) = mPrice
                FG.TextMatrix(LngNewRow, FG.ColIndex("Total")) = mPrice * val(FG.TextMatrix(i, FG.ColIndex("Qty")))
                FG.TextMatrix(LngNewRow, FG.ColIndex("LineID")) = mNewId
                
                FG.TextMatrix(LngNewRow, FG.ColIndex("ItemID2")) = Me.DcboItemID1.BoundText
                FG.TextMatrix(LngNewRow, FG.ColIndex("itemcode2")) = Trim$(Me.TxtAttachedItemCode.Text)
                FG.TextMatrix(LngNewRow, FG.ColIndex("ItemName2")) = Me.DcboItemID1.Text
            End If
        End If
Me.TxtAttachedItemCode2.Text = ""
   ' Me.DcboItemID2.BoundText = ""
    
    'Me.TxtAttachedItemCode2.SetFocus
End Sub
Private Sub FillGridItemType(ByVal mItemNo As Long, ByVal mItemName As String, ByVal mItemCode As String, mType As Integer, ByVal mUnitNo As Long, ByVal mUnitName As String, ByVal mQty As Double, ByVal mPrice As Double, ByVal mGroupID As Long, ByVal mGroupName As String, Optional ByVal mLineID As Long = 0)
    If IsSaveWithOutMsg Or mLineID <> 0 Then Exit Sub
    Dim i As Long
    Dim k As Long
    Dim LngNewRow As Long
      For i = 1 To FG2.Rows - 1
        If i > FG2.Rows - 1 Then Exit For
'        If val(FG2.TextMatrix(l, FG2.ColIndex("ItemType"))) = mType And mType = 1 Then
'            FG2.RemoveItem l
'            Exit For
'        End If
'        If val(FG2.TextMatrix(l, FG2.ColIndex("ItemID"))) = mItemNo And val(FG2.TextMatrix(l, FG2.ColIndex("ItemType"))) = mType Then
        If val(FG2.TextMatrix(i, FG2.ColIndex("ItemID"))) = 0 Then
            FG2.RemoveItem i
            i = i - 1
        End If
    Next
    Dim s As String
    Dim rsDummy As New ADODB.Recordset
    s = "Select increase,lowering From TblItems Where   ItemID = " & val(mItemNo)
    rsDummy.Open s, Cn, adOpenStatic, adLockReadOnly
    Dim mlowering As Double, mIncrease As Double
    If Not rsDummy.EOF Then
        mlowering = val(rsDummy!lowering & "")
        mIncrease = val(rsDummy!increase & "")
    End If

    If FG2.Rows = 1 Then FG2.Rows = 2 Else FG2.Rows = FG2.Rows + 1
        
        k = FG2.Rows
        If FG2.TextMatrix(k - 1, FG2.ColIndex("ItemID")) = "" Then
            'FG2.Rows = FG2.Rows - 1
           ' k = k - 1
        End If
        If FG2.Rows <= 1 Then
            FG2.Rows = FG2.Rows + 1
        End If
        LngNewRow = FG2.Rows - 1
        mNewId = LngNewRow
        If mNewId = 0 Then mNewId = LngNewRow
        With Me.FG2
            .TextMatrix(LngNewRow, .ColIndex("LineID")) = mNewId
            .TextMatrix(LngNewRow, .ColIndex("ItemID")) = mItemNo
            .TextMatrix(LngNewRow, .ColIndex("itemcode")) = mItemCode
            '.TextMatrix(LngNewRow, .ColIndex("FullCode")) = mItemCode
            .TextMatrix(LngNewRow, .ColIndex("itemname")) = mItemName
            .TextMatrix(LngNewRow, .ColIndex("UnitID")) = mUnitNo
            .TextMatrix(LngNewRow, .ColIndex("unitname")) = mUnitName
             
             .TextMatrix(LngNewRow, .ColIndex("GroupIDBuiltin")) = XPCboGroupBuiltin.BoundText
              .TextMatrix(LngNewRow, .ColIndex("GroupBuiltinName")) = XPCboGroupBuiltin.Text
               .TextMatrix(LngNewRow, .ColIndex("BuiltinItemID")) = DcboBuiltinItemID.BoundText
                .TextMatrix(LngNewRow, .ColIndex("BuiltinItemName")) = DcboBuiltinItemID.Text
                .TextMatrix(LngNewRow, .ColIndex("lowering")) = mlowering
                .TextMatrix(LngNewRow, .ColIndex("Increase")) = mIncrease
                .TextMatrix(LngNewRow, .ColIndex("NoteSerial14")) = mNoteSerial14
                .TextMatrix(LngNewRow, .ColIndex("TransactionID4")) = mTransactionID4
                .TextMatrix(LngNewRow, .ColIndex("NoteSerial15")) = mNoteSerial15
                .TextMatrix(LngNewRow, .ColIndex("TransactionID5")) = mTransactionID5
                
       
            .TextMatrix(LngNewRow, .ColIndex("Remark")) = txtRemark
       If SystemOptions.AllowChangManualQtyMix = True Then
            .TextMatrix(LngNewRow, .ColIndex("Qty")) = mQty
       Else
            .TextMatrix(LngNewRow, .ColIndex("Qty")) = mQty ' val(mQty) * IIf(val(txtQty1.Text) = 0, 1, val(txtQty1.Text))
       End If
'        .TextMatrix(LngNewRow, .ColIndex("FlgX")) = val(Me.TxtQty.Text)
'        .TextMatrix(LngNewRow, .ColIndex("TepQty")) = .TextMatrix(LngNewRow, .ColIndex("Qty"))
'        .TextMatrix(LngNewRow, .ColIndex("IsAdd")) = 1
     '   Dim mPrice As Double
        
        If mType = 2 Then
            mPrice = GetItemPrice(mItemNo, val(.TextMatrix(LngNewRow, .ColIndex("Qty"))), mUnitNo)
            .TextMatrix(LngNewRow, .ColIndex("Price")) = mPrice
        '    GetCostFromMix LngNewRow
            
            
            .TextMatrix(LngNewRow, .ColIndex("Total")) = mPrice * val(.TextMatrix(LngNewRow, .ColIndex("Qty")))
           ' .TextMatrix(LngNewRow, .ColIndex("TotalAdd")) = mPrice
        Else
             .TextMatrix(LngNewRow, .ColIndex("Price")) = mPrice
         '    GetCostFromMix LngNewRow
             .TextMatrix(LngNewRow, .ColIndex("widtj")) = txtwidtj
             .TextMatrix(LngNewRow, .ColIndex("hight")) = txthight
             .TextMatrix(LngNewRow, .ColIndex("Length")) = txtLength
             
             .TextMatrix(LngNewRow, .ColIndex("Length")) = txtLength
            .TextMatrix(LngNewRow, .ColIndex("thickness")) = txtthickness
            .TextMatrix(LngNewRow, .ColIndex("DO")) = txtDO
            .TextMatrix(LngNewRow, .ColIndex("DI")) = txtDI
            .TextMatrix(LngNewRow, .ColIndex("Diameter")) = txtDiameter
             .TextMatrix(LngNewRow, .ColIndex("Total")) = txtTotal
             
             .TextMatrix(LngNewRow, .ColIndex("Trans_DiscountType")) = XPCboDiscountType.ListIndex + 1
             .TextMatrix(LngNewRow, .ColIndex("Trans_Discount")) = XPTxtDiscountVal
             .TextMatrix(LngNewRow, .ColIndex("TotalDisc")) = txtTotalDisc
             .TextMatrix(LngNewRow, .ColIndex("TotalAdd")) = txtTotalAdd
             .TextMatrix(LngNewRow, .ColIndex("Net")) = txtNet
             .TextMatrix(LngNewRow, .ColIndex("Vat2")) = TxtVAt2
             .TextMatrix(LngNewRow, .ColIndex("TotalWithVat")) = txtTotalWithVat
        
            
        End If
        .TextMatrix(LngNewRow, .ColIndex("ItemType")) = mType
        
        .TextMatrix(LngNewRow, .ColIndex("GroupName")) = mGroupName
        .TextMatrix(LngNewRow, .ColIndex("GroupID")) = mGroupID
        
        
       ' .TextMatrix(LngNewRow, .ColIndex("ItemPrice")) = val(Me.TxtItemPrice(0).text)
        .AutoSize 0, .Cols - 1, False
    End With
    
    
txtRemark = ""

End Sub
Private Sub DcboItemID22_Change()
 Dim UnitID As Long
 Dim UnitName As String
'On Error Resume Next
    Dim Dcombos As ClsDataCombos
   Set Dcombos = New ClsDataCombos
    Me.TxtAttachedItemCode2.Text = GetItemCode(val(Me.DcboItemID2.BoundText))
    Dcombos.GetItemsUnitsDetai DcbUnit2, val(DcboItemID2.BoundText)
      GetDefaultItemUnit val(Me.DcboItemID2.BoundText), UnitID, UnitName
    DcbUnit2.Text = UnitName
    DcbUnit2.BoundText = UnitID
'     Me.TxtAttachedItemCode2.Text = GetItemCode(val(Me.DcboItemID2.BoundText))
End Sub



Private Sub txtTotalAdd_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        If TxtAttachedItemCode2.Text = "" Then
            Me.DcboItemID2.BoundText = ""
        Else
            Me.DcboItemID2.BoundText = GetItemID(Trim$(Me.TxtAttachedItemCode2.Text))
        End If
    End If
    
End Sub

Private Sub DcboItemID2_KeyUp(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyF3 Then
   Load FrmItemSearch
        FrmItemSearch.RetrunType = 1302
        FrmItemSearch.show vbModal
   End If
   
   
End Sub

Private Sub Dcbranch_Change()
        TxtNoteSerial11.Text = ""
              TxtNoteSerial12.Text = ""
End Sub


Sub FillAuto()
Dim code As String
Dim Name As String
Dim Unit As String
Dim Qty As Double
Dim ItemName As String
Dim ItemID As Double
Dim UnitID As Double
On Error Resume Next
Dim astrSplit2tems2() As String
If Me.TxtModFlg.Text <> "R" Then
     
    If SystemOptions.UserInterface = ArabicInterface Then
        If txtFile.Text = "" Then MsgBox "Õœœ «·„·› «Ê·«": Exit Sub
    Else
        If txtFile.Text = "" Then MsgBox "Select file first": Exit Sub
    End If
    Dim ExcelObj As Object
    Dim ExcelBook As Object
    Dim ExcelSheet As Object
    Dim i As Long

    Set ExcelObj = CreateObject("Excel.Application")
    Set ExcelSheet = CreateObject("Excel.Sheet")
    ExcelObj.Workbooks.Open txtFile.Text   ' App.Path & "\TrialBalance.xls"
DoEvents
    Set ExcelBook = ExcelObj.Workbooks(1)
    Set ExcelSheet = ExcelBook.Worksheets(1)
 Dim Counter As Integer
 Counter = 0
 Dim MixCode As String
 Dim MixName As String
 Dim ProductCode As String
 Dim UnitName As String
 Dim k As Integer
     With ExcelSheet
    i = 2
    k = 1
    Do Until .cells(i, k) & "" = "" Or i > 2
    MixCode = .cells(i, k)
    MixName = .cells(i, k + 1)
    ProductCode = .cells(i, k + 2)
    UnitName = .cells(i, k + 4)
    Qty = .cells(i, k + 5)
    TxtMaxNo.Text = MixCode
    TxtMaxName.Text = MixName
    TxtAttachedItemCode.Text = ProductCode
    Me.DcboItemID1.BoundText = GetItemID(Trim$(Me.TxtAttachedItemCode.Text))
    GetUnitID UnitName, UnitID
    DcbUnit.BoundText = UnitID
    txtQty1.Text = val(Qty)
        i = i + 1
    Loop

    End With
 FG.Clear flexClearScrollable, flexClearEverything
   FG.Rows = 1

 ''//////////////
     With ExcelSheet
    i = 6
    Do Until .cells(i, 1) & "" = ""
    code = .cells(i, 1)
    Unit = .cells(i, 3)
     Qty = .cells(i, 4)
  With FG
  
.Rows = .Rows + 1

GetItemsInformation code, ItemID, ItemName
GetUnitID Unit, UnitID
        .TextMatrix(.Rows - 1, .ColIndex("itemcode")) = code
        .TextMatrix(.Rows - 1, .ColIndex("FullCode")) = code
        .TextMatrix(.Rows - 1, .ColIndex("itemname")) = ItemName
        .TextMatrix(.Rows - 1, .ColIndex("ItemID")) = ItemID
        .TextMatrix(.Rows - 1, .ColIndex("FlgX")) = val(Qty)
        .TextMatrix(.Rows - 1, .ColIndex("Qty")) = val(txtQty1.Text) * val(.TextMatrix(.Rows - 1, .ColIndex("Qty")))
        .TextMatrix(.Rows - 1, .ColIndex("UnitID")) = UnitID
        .TextMatrix(.Rows - 1, .ColIndex("unitname")) = Unit
        
 End With

 If .cells(i, 1) & "" = "" Then Exit Sub
        i = i + 1
    Loop

    End With

Grid.SetFocus
       ExcelObj.Workbooks.Close

    Set ExcelSheet = Nothing
    Set ExcelBook = Nothing
    Set ExcelObj = Nothing
 End If
End Sub

Private Sub Dcbranch_GotFocus()
Dcbranch_Change

End Sub

Private Sub ISButton3_Click()
FillAuto
ReLineGrid
End Sub

Private Sub ISButton4_Click()
      FG.Clear flexClearScrollable, flexClearEverything
      FG.Rows = 1
CD1.ShowOpen
txtFile.Text = CD1.filename
TxtMaxNo.Text = ""
TxtMaxName.Text = ""
DcboItemID1.BoundText = 0
End Sub

Private Sub TxtAttachedItemCode_KeyDown(KeyCode As Integer, _
                                        Shift As Integer)
Dim UnitID As Long
Dim UnitName As String
    If KeyCode = vbKeyReturn Then
        If TxtAttachedItemCode.Text = "" Then
            Me.DcboItemID1.BoundText = ""
        Else
            Me.DcboItemID1.BoundText = GetItemID(Trim$(Me.TxtAttachedItemCode.Text))
            GetDefaultItemUnit val(Me.DcboItemID1.BoundText), UnitID, UnitName
            DcbUnit.Text = UnitName
            DcbUnit.BoundText = UnitID
        End If
    End If

End Sub
Private Sub DcboItemID1_Click(Area As Integer)
    
    DcboItemID1_Validate False
    
End Sub


Private Sub DcboItemID2_Change()
On Error Resume Next
If val(DcboItemID2.BoundText) = 0 Then Exit Sub
'If val(txtQty1) = 0 Then txtQty1 = 1
If val(DcboItemID2.BoundText) = val(val(DcboItemID1.BoundText)) Then DcboItemID2.Text = "": Exit Sub
Dim UnitID As Long
Dim UnitName As String
    Dim Dcombos As ClsDataCombos
 Set Dcombos = New ClsDataCombos
 
    Me.TxtAttachedItemCode2.Text = GetItemCode(val(Me.DcboItemID2.BoundText))
    Dcombos.GetItemsUnitsDetai DcbUnit2, val(DcboItemID2.BoundText)
    GetDefaultItemUnit val(Me.DcboItemID2.BoundText), UnitID, UnitName
    DcbUnit2.Text = UnitName
    DcbUnit2.BoundText = UnitID
    
    Me.TxtAttachedItemCode2.Text = GetItemCode(val(Me.DcboItemID2.BoundText))
Dcombos.GetItemsUnitsDetai DcbUnit2, val(DcboItemID2.BoundText)
If Me.TxtModFlg.Text <> "R" Then
   GetDefaultItemUnit val(Me.DcboItemID2.BoundText), UnitID, UnitName
    DcbUnit2.Text = UnitName
    DcbUnit2.BoundText = UnitID
    
    Dim l As Long

    
   ' FillGrid2
  End If
Dim widthPrice  As Double

'txtPrice = GetItemPriceByWitdth(val(DcboItemID1.BoundText), val(txtwidtj))
'If val(txtPrice) = 0 Then
'    txtPrice = GetItemPrice(val(Me.DcboItemID1.BoundText), , val(UnitID))
'    txtTotal = val(txtPrice) * val(txtQty1)
'End If
End Sub

Private Sub DcboItemID2_Click(Area As Integer)
    DcboItemID2_Change
End Sub

Private Sub Fg_AfterEdit(ByVal Row As Long, ByVal Col As Long)
    Dim Msg As String
    Dim rs As New ADODB.Recordset
    Dim StrSQL As String
    Dim ClsAcc As New ClsAccounts
    Dim LngRow As Long
    Dim k As Integer
Dim StrComboList As String
 Dim StrAccountCode As String
        mIdDisplay = val(FG.TextMatrix(Row, FG.ColIndex("LineID")))
        If mIdDisplay = 0 Then mIdDisplay = mNewId
    With FG

        Select Case .ColKey(Col)
              Case "name1"
 StrAccountCode = .ComboData
                             LngRow = .FindRow(StrAccountCode, .FixedRows, .ColIndex("SpecID1"), False, True)
                .TextMatrix(Row, .ColIndex("SpecID1")) = StrAccountCode
             Case "name2"
 StrAccountCode = .ComboData
                             LngRow = .FindRow(StrAccountCode, .FixedRows, .ColIndex("SpecID2"), False, True)
                .TextMatrix(Row, .ColIndex("SpecID2")) = StrAccountCode
                          Case "name3"
 StrAccountCode = .ComboData
                             LngRow = .FindRow(StrAccountCode, .FixedRows, .ColIndex("SpecID3"), False, True)
                .TextMatrix(Row, .ColIndex("SpecID3")) = StrAccountCode
                    Case "name4"
 StrAccountCode = .ComboData
                             LngRow = .FindRow(StrAccountCode, .FixedRows, .ColIndex("SpecID4"), False, True)
                .TextMatrix(Row, .ColIndex("SpecID4")) = StrAccountCode
                          Case "unitname"
 StrAccountCode = .ComboData
                             LngRow = .FindRow(StrAccountCode, .FixedRows, .ColIndex("UnitID"), False, True)
                .TextMatrix(Row, .ColIndex("UnitID")) = StrAccountCode
                CalcTotal Row
                Case "Qty"
                    .TextMatrix(Row, .ColIndex("FlgX")) = Round(val(.TextMatrix(Row, .ColIndex("Qty"))) / val(txtQty1), 2)
                    .TextMatrix(Row, .ColIndex("TepQty")) = val(.TextMatrix(Row, .ColIndex("Qty")))
                    .TextMatrix(Row, .ColIndex("Total")) = val(.TextMatrix(Row, .ColIndex("Qty"))) * val(.TextMatrix(Row, .ColIndex("Price")))
                    
                Case "Price"
                    .TextMatrix(Row, .ColIndex("Total")) = val(.TextMatrix(Row, .ColIndex("Qty"))) * val(.TextMatrix(Row, .ColIndex("Price")))
                    
   Case "itemname"
                StrAccountCode = .ComboData
                LngRow = .FindRow(StrAccountCode, .FixedRows, .ColIndex("ItemID"), False, True)
               .TextMatrix(Row, .ColIndex("ItemID")) = StrAccountCode
               .TextMatrix(Row, .ColIndex("ItemCode")) = GetItemCode(val(.TextMatrix(Row, .ColIndex("ItemID"))))
               If CheckItemParts(Row) = True Then
               .RemoveItem Row
               End If
             Case "itemcode"
             Set rs = New ADODB.Recordset
             StrSQL = " SELECT        TOP (100) PERCENT ItemID, ItemName, ItemNamee, Fullcode"
             StrSQL = StrSQL & "            From dbo.TblItems"
             StrSQL = StrSQL & "          WHERE        (Fullcode = N'" & .TextMatrix(Row, .ColIndex("ItemCode")) & "')"
             rs.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText
             If rs.RecordCount > 0 Then
             .TextMatrix(Row, .ColIndex("ItemID")) = IIf(IsNull(rs("ItemID").value), "", rs("ItemID").value)
             If SystemOptions.UserInterface = ArabicInterface Then
             .TextMatrix(Row, .ColIndex("itemname")) = IIf(IsNull(rs("ItemName").value), "", rs("ItemName").value)
             Else
             .TextMatrix(Row, .ColIndex("itemname")) = IIf(IsNull(rs("ItemNamee").value), "", rs("ItemNamee").value)
             End If
             Else
             .TextMatrix(Row, .ColIndex("ItemID")) = 0
              .TextMatrix(Row, .ColIndex("itemname")) = ""
             End If
                    If CheckItemParts(Row) = True Then
               .RemoveItem Row
               End If
              Case "unitname"
                StrAccountCode = .ComboData
                LngRow = .FindRow(StrAccountCode, .FixedRows, .ColIndex("UnitId"), False, True)
               .TextMatrix(Row, .ColIndex("UnitId")) = StrAccountCode
                    
               Case "Qty"
                    
                Case "FlgX"
                .TextMatrix(Row, .ColIndex("Total")) = val(.TextMatrix(Row, .ColIndex("Price"))) * val(.TextMatrix(Row, .ColIndex("Qty")))
                End Select
           '     CalcTotal Row
                If .ColKey(Col) <> "Qty" And .ColKey(Col) <> "Qty" And .ColKey(Col) <> "Select" Then
                    ReLineGrid val(FG.TextMatrix(Row, FG.ColIndex("LineID"))), True
                End If
                
                End With
  
    CalcGrid2
End Sub
Function CheckItemParts(Optional Row As Long) As Boolean
Dim i As Integer
With FG
CheckItemParts = False
For i = 1 To .Rows - 1
If val(.TextMatrix(i, .ColIndex("ItemID"))) = val(.TextMatrix(Row, .ColIndex("ItemID"))) And i <> Row Then
CheckItemParts = True
Exit Function
End If
Next i
End With
End Function
Private Sub CalcTotal(ByVal LngNewRow As Long)
        Dim mPrice As Double, mItemNo As Long, mUnitNo As Long
        With FG
        mItemNo = val(.TextMatrix(LngNewRow, .ColIndex("ItemID")))
        mUnitNo = val(.TextMatrix(LngNewRow, .ColIndex("UnitID")))
        mPrice = GetItemPrice(mItemNo, , mUnitNo)
        .TextMatrix(LngNewRow, .ColIndex("Price")) = mPrice
        .TextMatrix(LngNewRow, .ColIndex("Total")) = mPrice * val(.TextMatrix(LngNewRow, .ColIndex("Qty")))
        End With

End Sub

Private Function GetItemPriceByWitdth(Item_ID As Long, Width As Double, Optional ByVal mLenth As Double = 0) As Double
   'Dim StrSQL  As String
   Dim mIsPriceIsLenthW As Boolean
   Dim StrSQL As String
     Dim rs As ADODB.Recordset
    Set rs = New ADODB.Recordset
     StrSQL = "SELECT     IsNull(IsPriceIsPerview,0) IsPriceIsPerview ,IsNull(IsPriceIsLenthW,0) IsPriceIsLenthW "
StrSQL = StrSQL & " From dbo.TblItems"
StrSQL = StrSQL & "  Where (IsNull(IsPriceIsPerview,0) =1 Or IsNull(IsPriceIsLenthW,0) =1) and ItemID = " & Item_ID
rs.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText
If rs.EOF Then
    GetItemPriceByWitdth = 0
    Exit Function
End If

    
If (rs!IsPriceIsLenthW & "") Then
    mIsPriceIsLenthW = True
    Width = (Width * (mLenth + 20))
End If

 StrSQL = "SELECT     dbo.Fn_GetPriceItem(" & Item_ID & ", " & Width & ") AS WidthPrice  "
StrSQL = StrSQL & " From dbo.TblItems"
StrSQL = StrSQL & "  Where (IsNull(IsPriceIsPerview,0) =1 Or IsNull(IsPriceIsLenthW,0) =1) and ItemID = " & Item_ID
Set rs = New ADODB.Recordset
 
  
  rs.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText
   
    If (rs.RecordCount) > 0 Then
       If Not mIsPriceIsLenthW Then
            GetItemPriceByWitdth = IIf(IsNull(rs("WidthPrice").value), 0, val((rs("WidthPrice").value) & "")) * Width / 100
            ' «·Ê÷⁄ «·ﬁœÌ„ ﬂ‰« »‰ﬁ”„ ⁄·Ï 100
            'GetItemPriceByWitdth = IIf(IsNull(rs("WidthPrice").value), 0, val((rs("WidthPrice").value) & "")) * Width / 100
        Else
            GetItemPriceByWitdth = IIf(IsNull(rs("WidthPrice").value), 0, val((rs("WidthPrice").value) & "")) * Width / 10000
        End If
 
Else
GetItemPriceByWitdth = 0
    End If
End Function

Private Function GetItemAddPrice(Item_ID As Long) As Double
   'Dim StrSQL  As String
     Dim rs As ADODB.Recordset
    Set rs = New ADODB.Recordset
    Dim StrSQL  As String

    StrSQL = " SELECT PartItemID,PartItemQty,Unitid "
    StrSQL = StrSQL & " FROM TblItemsParts WHERE ItemId =" & Item_ID & "  AND ISNULL(IsAddToPrice,0) = 1"
 
  
    rs.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText
    GetItemAddPrice = 0
    Do While Not rs.EOF
       GetItemAddPrice = GetItemAddPrice + GetItemPrice(val(rs!PartItemID), , val(rs!UnitID))
       rs.MoveNext
    Loop


End Function

Private Sub FG_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
With FG
Select Case .ColKey(Col)
Case "Amout1"
.ComboList = ""
Case "Amout2"
.ComboList = ""
Case "Amout3"
.ComboList = ""
Case "Amout4"
.ComboList = ""
Case "itemcode"
.ComboList = ""
Case "itemname"
.ComboList = ""
Case "unitname"
.ComboList = ""
Case "cost"
.ComboList = ""
Case "FlgX"
.ComboList = ""
Case "Qty"

If SystemOptions.AllowChangManualQtyMix = True Then
.ComboList = ""

Else
Cancel = True
End If
Case "ShowAttatch", "lowering", "Increase"
    .EditMaxLength = 10
Case "IsAdd"
    Cancel = True
End Select
End With
End Sub

Private Sub fg_StartEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    
  Dim rs As New ADODB.Recordset
    Dim StrSQL  As String
    Dim StrAccountType As String
    Dim StrComboList As String
    Dim Msg As String
Dim StrAccountCode As String
Dim LngRow As Double
    With FG

        Select Case .ColKey(Col)
Case "name1"
  StrSQL = "select * from TblSpecification"
                rs.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

                If SystemOptions.UserInterface = ArabicInterface Then
                
                    StrComboList = FG.BuildComboList(rs, "name", "id")
                Else
                    StrComboList = FG.BuildComboList(rs, "namee", "id")
                End If
       
                If StrComboList <> "" Then
                    StrComboList = "|" & StrComboList
                End If

                .ComboList = StrComboList
    Case "name2"
      StrSQL = "select * from TblSpecification"
                rs.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

                If SystemOptions.UserInterface = ArabicInterface Then
                
                    StrComboList = FG.BuildComboList(rs, "name", "id")
                Else
                    StrComboList = FG.BuildComboList(rs, "namee", "id")
                End If
       
                If StrComboList <> "" Then
                    StrComboList = "|" & StrComboList
                End If

                .ComboList = StrComboList
           Case "name3"
  StrSQL = "select * from TblSpecification"
                rs.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

                If SystemOptions.UserInterface = ArabicInterface Then
                
                    StrComboList = FG.BuildComboList(rs, "name", "id")
                Else
                    StrComboList = FG.BuildComboList(rs, "namee", "id")
                End If
       
                If StrComboList <> "" Then
                    StrComboList = "|" & StrComboList
                End If

                .ComboList = StrComboList
                
   Case "name4"
  StrSQL = "select * from TblSpecification"
                rs.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

                If SystemOptions.UserInterface = ArabicInterface Then
                
                    StrComboList = FG.BuildComboList(rs, "name", "id")
                Else
                    StrComboList = FG.BuildComboList(rs, "namee", "id")
                End If
       
                If StrComboList <> "" Then
                    StrComboList = "|" & StrComboList
                End If

                .ComboList = StrComboList
          Case "unitname"
  'StrSQL = "select * from TblUnites"
                
            If SystemOptions.UserInterface = ArabicInterface Then
                StrSQL = "SELECT dbo.TblUnites.UnitID, dbo.TblUnites.UnitName"
            Else
                StrSQL = "SELECT dbo.TblUnites.UnitID, dbo.TblUnites.UnitNamee"
            End If
            StrSQL = StrSQL + "   FROM  dbo.TblItemsUnits LEFT OUTER JOIN"
             StrSQL = StrSQL + "  dbo.TblUnites ON dbo.TblItemsUnits.UnitID = dbo.TblUnites.UnitID"
            StrSQL = StrSQL + " Where dbo.TblItemsUnits.ItemID=" & val(.TextMatrix(Row, .ColIndex("ItemID"))) & " "
    
    
                rs.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

                If SystemOptions.UserInterface = ArabicInterface Then
                
                    StrComboList = FG.BuildComboList(rs, "UnitName", "UnitID")
                Else
                    StrComboList = FG.BuildComboList(rs, "UnitNamee", "UnitID")
                End If
       
                If StrComboList <> "" Then
                    StrComboList = "|" & StrComboList
                End If

                .ComboList = StrComboList
 Case "itemname"
     StrSQL = " SELECT     ItemID, ItemName, ItemNamee"
     StrSQL = StrSQL & "  From dbo.TblItems"
     Set rs = New ADODB.Recordset
                rs.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText
        If SystemOptions.UserInterface = ArabicInterface Then
                    StrComboList = .BuildComboList(rs, "ItemName", "ItemID")
                Else
                    StrComboList = .BuildComboList(rs, "ItemNamee", "ItemID")
                End If
       
                If StrComboList <> "" Then
                    StrComboList = "|" & StrComboList
                End If
                 .ComboList = StrComboList
                    
                
        End Select

    End With


End Sub



Private Sub Form_Activate()
    'XPTxtBillID.SetFocus
End Sub

Private Sub Selct_Click(Index As Integer)
Select Case Index
Case 0
 'If Selct(0).value = vbChecked Then
 'Selct(1).Enabled = True
 ' Selct(2).Enabled = True
 ' Else
 ' Selct(1).Enabled = False
 ' Selct(2).Enabled = False
 ' End If
  Case 1
  
 ' If Selct(1).value = vbChecked Then
 ' DCboStore2Name.Enabled = True
 ' Else
 ' DCboStore2Name.Enabled = False
 ' TxtNoteSerial11.Text = ""
 ' End If
  
  
    Case 2
'  If Selct(2).value = vbChecked Then
'  DCboStore3Name.Enabled = True
'  Else
'  DCboStore3Name.Enabled = False
'  TxtNoteSerial12.Text = ""
'  End If
  End Select
End Sub



Private Sub TxtAttachedItemCode2_KeyPress(KeyAscii As Integer)
 If KeyAscii = vbKeyReturn Then
        If TxtAttachedItemCode2.Text = "" Then
            Me.DcboItemID2.BoundText = ""
        Else
            Me.DcboItemID2.BoundText = GetItemID(Trim$(Me.TxtAttachedItemCode2.Text))
        End If
    End If


End Sub

Private Sub TxtAttachedItemCode2_KeyUp(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyF3 Then
   Load FrmItemSearch
        FrmItemSearch.RetrunType = 1302
        FrmItemSearch.show vbModal
   End If
   
End Sub

Private Sub txthight_Change()

Dim widthPrice  As Double
TxtPrice = 0
TxtPrice = GetItemPriceByWitdth(val(DcboItemID1.BoundText), val(txtwidtj), val(txthight))
If val(TxtPrice) = 0 Then
    TxtPrice = GetItemPrice(val(Me.DcboItemID1.BoundText), , val(DcbUnit.BoundText))
End If
TxtPrice = val(TxtPrice) + GetItemAddPrice(val(DcboItemID1.BoundText))
If Me.TxtModFlg.Text <> "R" Then
    txtTotal = val(TxtPrice) * val(txtQty1)
    'ReLineGrid
End If
If Not SystemOptions.IsMultiItemsInCompItem Then
    If FG2.Rows > 1 Then
        FG2.TextMatrix(1, FG2.ColIndex("Qty")) = txtQty1
        FG2.TextMatrix(1, FG2.ColIndex("Price")) = TxtPrice
        FG2.TextMatrix(1, FG2.ColIndex("widtj")) = txtwidtj
        FG2.TextMatrix(1, FG2.ColIndex("hight")) = txthight
        FG2.TextMatrix(1, FG2.ColIndex("Length")) = txtLength
        
    End If
End If
CalcTotalNet



callarea
End Sub

Private Sub txtQty1_Change()
If Me.TxtModFlg.Text <> "R" Then
    txtTotal = val(TxtPrice) * val(txtQty1)
    If Not SystemOptions.IsMultiItemsInCompItem Then
        If FG2.Rows > 1 Then
            FG2.TextMatrix(1, FG2.ColIndex("Qty")) = txtQty1
            FG2.TextMatrix(1, FG2.ColIndex("Price")) = TxtPrice
            FG2.TextMatrix(1, FG2.ColIndex("widtj")) = txtwidtj
            FG2.TextMatrix(1, FG2.ColIndex("hight")) = txthight
            FG2.TextMatrix(1, FG2.ColIndex("Length")) = txtLength
            
        End If
        ReLineGrid , True
    End If
    If val(DcboItemID4.Tag) <> 0 Then
      '  ReLineGrid val(DcboItemID4.Tag), True
    End If
End If
CalcTotalNet
End Sub

Private Sub TxtSearchCode_KeyPress(KeyAscii As Integer)
 Dim EmpID As Integer

    If KeyAscii = vbKeyReturn Then
        'GetTblCustemersCode TxtSearchCode.Text, EmpID
        'DBCboClientName.BoundText = EmpID
        GetCustomerNamebyPhone , , , TxtSearchCode.Text
    End If
End Sub

Private Sub txtwidtj_Change()
'callarea
Dim widthPrice  As Double
TxtPrice = 0
TxtPrice = GetItemPriceByWitdth(val(DcboItemID1.BoundText), val(txtwidtj), val(txthight))
If val(TxtPrice) = 0 Then
    TxtPrice = GetItemPrice(val(Me.DcboItemID1.BoundText), , val(DcbUnit.BoundText))
End If
TxtPrice = val(TxtPrice) + GetItemAddPrice(val(DcboItemID1.BoundText))
If Me.TxtModFlg.Text <> "R" Then
    txtTotal = val(TxtPrice) * val(txtQty1)
    'ReLineGrid
End If
If Not SystemOptions.IsMultiItemsInCompItem Then
    If FG2.Rows > 1 Then
        FG2.TextMatrix(1, FG2.ColIndex("Qty")) = txtQty1
        FG2.TextMatrix(1, FG2.ColIndex("Price")) = TxtPrice
        FG2.TextMatrix(1, FG2.ColIndex("widtj")) = txtwidtj
        FG2.TextMatrix(1, FG2.ColIndex("hight")) = txthight
        FG2.TextMatrix(1, FG2.ColIndex("Length")) = txtLength
        
    End If
End If
CalcTotalNet



End Sub
Function callarea()
DcboItemID1_Validate False
'txtQty1 = val(Me.txtwidtj) * val(Me.txthight)
End Function

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

Private Sub Form_KeyDown(KeyCode As Integer, _
                         Shift As Integer)

    'On Error GoTo ErrTrap
    If KeyCode = vbKeyReturn Then
        If Me.TxtModFlg.Text = "R" Then
            '        Cmd_Click (0)
        Else
            '        SendKeys "{TAB}"
        End If
    End If

    If KeyCode = vbKeyF12 Then
        If cmd(0).Enabled = False Then Exit Sub
        Cmd_Click (0)
    End If

    If KeyCode = vbKeyF11 Then
        If cmd(1).Enabled = False Then Exit Sub
        Cmd_Click (1)
    End If

    If KeyCode = vbKeyF10 Then
        If cmd(2).Enabled = False Then Exit Sub
        Cmd_Click (2)
    End If

    If KeyCode = vbKeyF9 Then
        If cmd(3).Enabled = False Then Exit Sub
        Cmd_Click (3)
    End If

    If KeyCode = vbKeyF8 Then
        If cmd(4).Enabled = False Then Exit Sub
        Cmd_Click (4)
    End If

    If KeyCode = vbKeyF3 Then
        If cmd(5).Enabled = False Then Exit Sub
        Cmd_Click (5)
    End If

    If KeyCode = vbKeyF6 Then
        If cmd(7).Enabled = False Then Exit Sub
        Cmd_Click (7)
    End If

    If KeyCode = vbKeyF2 Then
        If TxtModFlg.Text = "N" Or TxtModFlg.Text = "E" Then
        
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
  '  Dim ShowTax As Boolean
   ' Dim Dcombos As ClsDataCombos
     If SystemOptions.AllowChangManualQtyMix = True Then
        FG.ColHidden(FG.ColIndex("FlgX")) = True
     End If
        Selct_Click (0)
        Selct_Click (1)
        Selct_Click (2)
    AddTip
    XPDtbBill.value = Date
    Set Dcombos = New ClsDataCombos

   'Me.GetItemsNames , -1, -1, 1, , storename1
    

  Dim s As String

    If Not SystemOptions.DontCreateOut2 Then
        cmdCreateSales.Visible = False
    Else
        cmdCreateSales.Visible = True
    End If

  RegisterLogInOut Me.Name, ScreenNameArabic, ScreenNameEnglish, "1"
    If SystemOptions.IsGeometricProportions Then
        Frame4.Visible = True
    Else
        Frame4.Visible = False
    End If

    Set cmd(0).ButtonImage = mdifrmmain.ImgLstTree.ListImages("New").Picture
    Set cmd(1).ButtonImage = mdifrmmain.ImgLstTree.ListImages("Edit").Picture
    Set cmd(2).ButtonImage = mdifrmmain.ImgLstTree.ListImages("save").Picture
    Set cmd(3).ButtonImage = mdifrmmain.ImgLstTree.ListImages("Undo").Picture
    Set cmd(4).ButtonImage = mdifrmmain.ImgLstTree.ListImages("Del").Picture
    Set cmd(5).ButtonImage = mdifrmmain.ImgLstTree.ListImages("Search").Picture
    Set cmd(6).ButtonImage = mdifrmmain.ImgLstTree.ListImages("Exit").Picture
    Set cmd(7).ButtonImage = mdifrmmain.ImgLstTree.ListImages("Print").Picture
    Set CmdHelp.ButtonImage = mdifrmmain.ImgLstTree.ListImages("Help").Picture
    If SystemOptions.UserInterface = ArabicInterface Then
        FG2.ColComboList(FG2.ColIndex("Trans_DiscountType")) = "#1;·«ÌÊÃœ Œ’„|#2;Œ’„ »ﬁÌ„…|#3;Œ’„ »‰”»…|#4;„Ã«‰Ì"
    ElseIf SystemOptions.UserInterface = EnglishInterface Then
        FG2.ColComboList(FG2.ColIndex("Trans_DiscountType")) = "#1;NO Discount|#2;Value Discount|#3;Percent Discount|#4;Free"
    End If
    Dcombos.GetItemsNames Me.DcboItemID1, -1, -1, 1
    Dcombos.GetItemsUnits DcbUnit
    Dcombos.GetItemsNames Me.DcboItemID2
    Dcombos.GetItemsNames Me.DcboBuiltinItemID
    Dcombos.GetItemsNames Me.DcboItemID3
    Dcombos.GetItemsNames Me.DcboItemID4
    Dcombos.GetItemsUnits DcbUnit2
    Dcombos.GetItemsUnits DcbUnit3
    Dcombos.GetStores Me.DCboStore2Name
    Dcombos.GetStores Me.DCboStoreName
    Dcombos.GetStores Me.DCboStore3Name
    Dcombos.GetUsers Me.DCboUserName
    Dcombos.GetSalesRepData Me.DcboEmp
    Dcombos.GetItemSpecifications Me.cmbSpecification
    
    
    If SystemOptions.UserInterface = ArabicInterface Then
        StrSQL = "SELECT * From groups where IsProducer = 1 "
        fill_combo XPCboGroup, StrSQL
    
        
        
        
        StrSQL = "SELECT * From groups "
        fill_combo XPCboGroupBuiltin, StrSQL
    
        
        StrSQL = "SELECT * From groups where IsAdditions = 1 "
        fill_combo XPCboGroup2, StrSQL
        fill_combo XPCboGroup5, StrSQL
    Else
        StrSQL = "SELECT GroupID, GroupNamee From groups where IsProducer = 1 "
        fill_combo XPCboGroup, StrSQL
    
        
        
        
        StrSQL = "SELECT GroupID, GroupNamee  From groups "
        fill_combo XPCboGroupBuiltin, StrSQL
    
        
        StrSQL = "SELECT GroupID, GroupNamee  From groups where IsAdditions = 1 "
        fill_combo XPCboGroup2, StrSQL
        fill_combo XPCboGroup5, StrSQL
    End If
  If SystemOptions.UserInterface = EnglishInterface Then
        SetInterface Me
        ChangeLang
    End If
 

    
    If SystemOptions.UserInterface = ArabicInterface Then
        FG2.ColComboList(FG2.ColIndex("ItemType")) = "#1;’‰› „‰ Ã |#2;’‰› «÷«›« |"
    Else
        FG2.ColComboList(FG2.ColIndex("ItemType")) = "#2;Production |#1;Add|"
    End If
    Dcombos.GetCustomersSuppliers 1, Me.DBCboClientName, True
    Dcombos.GetBranches Me.dcBranch
    Dcombos.GetStores Me.DCboStoreName
        LoadCombosData
        


    StrSQL = "SELECT * FROM TblDefComItem "
StrSQL = StrSQL & "  WHERE      BranchId in(" & Current_branchSql & ")"
    StrSQL = StrSQL + " Order By ID"
    Set rs = New ADODB.Recordset
    rs.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText


  s = "Select StoreID,StoreID1,StoreID2,StoreID3 from tblUsers Where UserID = " & user_id
  Set rsDummy = New ADODB.Recordset

rsDummy.Open s, Cn, adOpenKeyset, adLockReadOnly, adCmdText
If Not rsDummy.EOF Then
    DCboStore2Name.BoundText = val(rsDummy!StoreId2 & "")
    DCboStore3Name.BoundText = val(rsDummy!StoreID3 & "")
    DCboStoreName.BoundText = val(rsDummy!StoreId1 & "")
End If
 'Wael
' SystemOptions.IsMultiItemsInCompItem = True
  TabMain.TabVisible(3) = False
If SystemOptions.IsMultiItemsInCompItem Then
 '   TabMain.TabIndex = 1
   ' TabMain.TabVisible(0) = True
   XPCboGroupBuiltin.Visible = True
   lbl(77).Visible = True
   lbl(79).Visible = True
   txtItemCodeBuiltin.Visible = True
   lbl(78).Visible = True
   DcboBuiltinItemID.Visible = True
   lbl(82).Visible = True
   txtLength.Visible = True
   CboPayMentType.Visible = True
   lbl(54).Visible = True
    lbl(76).Visible = True
   txtCustomerName.Visible = True
    lbl(75).Visible = True
    txtPeriod.Visible = True
    cmdfrmRec.Visible = True
    TabMain.CurrTab = 2
   TxtPhone.Visible = True
  cmdAddCustomer.Visible = True
    lbl(84).Visible = True
        lbl(87).Visible = True
        lbl(88).Visible = True
    
    If SystemOptions.UserInterface = ArabicInterface Then
        lbl(64).Caption = "«·‰Ê⁄"
        lbl(51).Caption = "«·„ÊœÌ·"
        lbl(26).Caption = "«·’‰›"
    Else
        lbl(64).Caption = "Type"
        lbl(51).Caption = "Model"
        lbl(26).Caption = "Item"
    
    End If
    ISButton4.Visible = False
    ISButton3.Visible = False
    Selct(0).Visible = False
    
    DcboEmp.Visible = True
    TxtEmployeeID.Visible = True
    lbl(58).Visible = True
    DcboBox.Visible = True
    lbl(59).Visible = True
    lbl(54).Visible = True
'   CboPayMentType.Visible = False
  '  TabMain.CurrTab = 0
  '  TabMain.CurrTab = 1
  TxtNoteSerial14.Visible = False
  If SystemOptions.HideCost = True Then
    TabMain.CurrTab = 2
    'TabMain.TabVisible(0) = True
    TabMain.TabVisible(1) = False
    TabMain.TabVisible(0) = False
'
    TabMain.FirstTab = 2
  End If
Else
    lbl(76).Visible = False
   txtCustomerName.Visible = False
   CboPayMentType.Visible = False
   lbl(54).Visible = False
    lbl(75).Visible = False
    txtPeriod.Visible = False
    cmdfrmRec.Visible = False
    
    lbl(84).Visible = False
    cmdAddCustomer.Visible = False
    TxtPhone.Visible = False
    txtRemark.Visible = False
    lbl(74).Visible = False
    TabMain.CurrTab = 1
    'TabMain.TabVisible(0) = True
    TabMain.TabVisible(2) = False
'
    TabMain.FirstTab = 1
    
    txtwidtj.Visible = False
    lbl(42).Visible = False
    TxtSearchCode.Visible = False
    DBCboClientName.Visible = False
    lbl(49).Visible = False
    txthight.Visible = False
    lbl(41).Visible = False
    TxtEmployeeID.Visible = False
    DcboEmp.Visible = False
    lbl(58).Visible = False
    DcboBox.Visible = False
    lbl(59).Visible = False
    lbl(54).Visible = False
    CboPayMentType.Visible = False
    lbl(66).Visible = False
    XPCboDiscountType.Visible = False
    lbl(68).Visible = False
    XPTxtDiscountVal.Visible = False
     lbl(53).Visible = False
     lbl(55).Visible = False
     lbl(56).Visible = False
     lbl(52).Visible = False
     TxtPrice.Visible = False
     txtTotalAdd.Visible = False
     txtTotalDisc.Visible = False
     txtNet.Visible = False
     lbl(57).Visible = False
     txtTotal.Visible = False
     lbl(98).Visible = False
      lbl(99).Visible = False
      lbl(72).Visible = False
    lbl(73).Visible = False
    lbl(70).Visible = False
    lbl(71).Visible = False
      lbl(69).Visible = False
      txtTotalWithVat.Visible = False
     TxtVAt2.Visible = False
     txtItemCode.Visible = False
    DcboItemID4.Visible = False
    txtQty3.Visible = False
    DcbUnit3.Visible = False
    DcboItemID3.Visible = False
    txtItemCode.Visible = False
    cmdAdd2.Visible = False
    TxtNoteSerial14.Visible = True
End If
If SystemOptions.DontCreateOut Then
    TabMain.TabVisible(3) = True
End If

   If SystemOptions.DontShowMoreDetailsCompItem Then
        lbl(82).Visible = False
        lbl(83).Visible = False
        lbl(85).Visible = False
        lbl(29).Visible = False
        lbl(81).Visible = False
        lbl(30).Visible = False
        lbl(77).Visible = False
        lbl(79).Visible = False
        lbl(78).Visible = False
        lbl(81).Visible = False
        lbl(86).Visible = False
        lbl(76).Visible = False
        lbl(84).Visible = False
        lbl(87).Visible = False
        lbl(88).Visible = False
        
        'lbl(54).Visible = False
        ' lbl(58).Visible = False
       ' TxtEmployeeID.Visible = False
       ' DcboEmp.Visible = False
        
        txtLength.Visible = False
        txtthickness.Visible = False
        txtDI.Visible = False
        txtDO.Visible = False
        txtDiameter.Visible = False
        txtDiameter.Visible = False
        TxtMaxNo.Visible = False
        TxtMaxName.Visible = False
        TxtPhone.Visible = False
        txtCustomerName.Visible = False
        TxtMaxNo2.Visible = False
        txtRecTime.Visible = False
        TxtMaxNo.Visible = False
        TxtMaxNo.Visible = False
        TxtMaxNo.Visible = False
        txtLength.Visible = False
         FG2.ColHidden(FG2.ColIndex("GroupBuiltinName")) = True
         FG2.ColHidden(FG2.ColIndex("BuiltinItemName")) = True
         FG2.ColHidden(FG2.ColIndex("cost")) = True
         FG2.ColHidden(FG2.ColIndex("ItemName2")) = True
         FG2.ColHidden(FG2.ColIndex("ItemName5")) = True
         
         FG2.ColHidden(FG2.ColIndex("CountItem2")) = True
         FG2.ColHidden(FG2.ColIndex("CountItem5")) = True
         
         FG2.ColHidden(FG2.ColIndex("lowering")) = True
         FG2.ColHidden(FG2.ColIndex("Length")) = True
         FG2.ColHidden(FG2.ColIndex("thickness")) = True
         FG2.ColHidden(FG2.ColIndex("Diameter")) = True
         FG2.ColHidden(FG2.ColIndex("DO")) = True
         FG2.ColHidden(FG2.ColIndex("DI")) = True
         
         
                
         FG2.ColHidden(FG2.ColIndex("lowering")) = True
         FG2.ColHidden(FG2.ColIndex("Length")) = True
         FG2.ColHidden(FG2.ColIndex("thickness")) = True
         FG2.ColHidden(FG2.ColIndex("Diameter")) = True
         FG2.ColHidden(FG2.ColIndex("DO")) = True
         FG2.ColHidden(FG2.ColIndex("DI")) = True
         
         
         
         If SystemOptions.IsGeometricProportions Then
            FG.ColHidden(FG.ColIndex("Length")) = True
            FG.ColHidden(FG.ColIndex("thickness")) = True
            FG.ColHidden(FG.ColIndex("widtj")) = True
            FG.ColHidden(FG.ColIndex("hight")) = True
            FG.ColHidden(FG.ColIndex("Diameter")) = True
            FG.ColHidden(FG.ColIndex("DO")) = True
            FG.ColHidden(FG.ColIndex("DI")) = True
         
         End If
         
         FG2.ColHidden(FG2.ColIndex("Trans_DiscountType")) = True
FG2.ColHidden(FG2.ColIndex("unitname")) = True
'fg2.ColHidden(fg2.ColIndex("hight")) = True
'fg2.ColHidden(fg2.ColIndex("widtj")) = True
FG2.ColHidden(FG2.ColIndex("lowering")) = True
FG2.ColHidden(FG2.ColIndex("Length")) = True
FG2.ColHidden(FG2.ColIndex("Trans_DiscountPercent")) = True
FG2.ColHidden(FG2.ColIndex("TotalAdd")) = True
FG2.ColHidden(FG2.ColIndex("Net")) = True
FG2.ColHidden(FG2.ColIndex("Vat2")) = True
'fg2.ColHidden(fg2.ColIndex("TotalWithVat")) = True
FG2.TextMatrix(0, FG2.ColIndex("TotalWithVat")) = "«·’«›Ì »⁄œ «·Œ’„"
FG2.ColHidden(FG2.ColIndex("NoteSerial14")) = True
    lbl(5).Caption = "—ﬁ„ ⁄—÷ «·”⁄—"
         XPCboGroupBuiltin.Visible = False
         txtItemCodeBuiltin.Visible = False
         DcboBuiltinItemID.Visible = False
         'CboPayMentType.Visible = False
         cmdAddCustomer.Caption = "› Õ „·› «·⁄„·«¡"
                  
   End If
    If SystemOptions.UserInterface = ArabicInterface Then

        With CboPayMentType
             .Clear
             .AddItem "‰ﬁœ«"
             .AddItem "¬Ã·"
         End With
       With XPCboDiscountType
            .Clear
            .AddItem "·«ÌÊÃœ Œ’„"
            .AddItem "Œ’„ »ﬁÌ„…"
            .AddItem "Œ’„ »‰”»…"
            .AddItem "„Ã«‰Ï"
        End With
         
    Else
         With CboPayMentType
            .Clear
            'AddItem "Cash"
            
            .AddItem "Cash"
            .AddItem "Credit"
        End With
        
      With XPCboDiscountType
            .Clear
            .AddItem "NO Discount"
            .AddItem "Value Discount"
            .AddItem "Precetage Discount"
            .AddItem "Free"
        End With
    End If
Retrive

    Me.TxtModFlg.Text = "R"

    If OPEN_NEW_SCREEN = True Then
        Cmd_Click (0)
    End If
  
    Exit Sub
ErrTrap:
End Sub

Private Sub Form_Unload(Cancel As Integer)
    On Error GoTo ErrTrap
    Dim i As Integer
    RegisterLogInOut Me.Name, ScreenNameArabic, ScreenNameEnglish, "1"

    If rs.State = adStateOpen Then
        If Not (rs.EOF Or rs.BOF) Then
            If rs.EditMode <> adEditNone Then
                rs.CancelUpdate
            End If
        End If

        rs.Close
    End If

    Set Dcombos = Nothing

    For i = LBound(cSearchDcbo) To UBound(cSearchDcbo)
        Set cSearchDcbo(i) = Nothing
    Next i

  '  For i = LBound(cSearchDcbo) To UBound(cSearchDcbo)
  '      Set cSearchDcbo(i) = Nothing
  '  Next i

    Set rs = Nothing
    Set TTP = Nothing
 '   NewGrid.Class_Terminate
  '  Set NewGrid = Nothing
    'Set SaleReport = Nothing
    Exit Sub
ErrTrap:
End Sub



Private Sub TxtModFlg_Change()
     On Error GoTo ErrTrap

    Select Case Me.TxtModFlg.Text

        Case "R"

            If SystemOptions.UserInterface = ArabicInterface Then
                Me.Caption = "   ⁄—Ì› „ﬂÊ‰«  «·«’‰«›/«· Ã„Ì⁄     "
            Else
                Me.Caption = "Definition of items / components assembly"
            End If

           ' Frame4.Enabled = False
'            Ele(11).Enabled = False
   
            Me.cmd(2).Enabled = False
            Me.cmd(3).Enabled = False
            Me.cmd(0).Enabled = True
            Me.cmd(1).Enabled = True
            Me.cmd(4).Enabled = True
            Me.cmd(5).Enabled = True
            Me.cmd(7).Enabled = True
            Me.XPBtnMove(0).Enabled = True
            Me.XPBtnMove(1).Enabled = True
            Me.XPBtnMove(2).Enabled = True
            Me.XPBtnMove(3).Enabled = True
            XPBtnNewClients.Enabled = False
            CboPayMentType.locked = True
            Me.XPDtbBill.Enabled = False
            Me.DBCboClientName.locked = True
            Me.DCboStoreName.locked = True
            FG.Editable = flexEDNone
            Me.DcboEmp.Enabled = True
            
            CmdConvert.Enabled = True
            ' CmdConvert.Visible = True
            CmdTemplate.Visible = False

            If rs.RecordCount < 1 Then
                Me.XPBtnMove(0).Enabled = False
                Me.XPBtnMove(1).Enabled = False
                Me.XPBtnMove(2).Enabled = False
                Me.XPBtnMove(3).Enabled = False
                Me.cmd(1).Enabled = False
                Me.cmd(4).Enabled = False
                Me.cmd(5).Enabled = False
                Me.cmd(7).Enabled = False
                CmdConvert.Enabled = False
            End If

            Ele(2).Enabled = False
            FG2.Enabled = True
        Case "N"

            If SystemOptions.UserInterface = ArabicInterface Then
                Me.Caption = "  ⁄—Ì› „ﬂÊ‰«  «·«’‰«›/«· Ã„Ì⁄    "
            Else
                Me.Caption = "Definition of items / components assembly"
            End If
   
        '    Frame4.Enabled = True
'            Ele(11).Enabled = True
         
            Me.cmd(2).Enabled = True
            Me.cmd(3).Enabled = True
            Me.cmd(0).Enabled = False
            Me.cmd(1).Enabled = False
            Me.cmd(4).Enabled = False
            Me.cmd(5).Enabled = False
            Me.cmd(7).Enabled = False
           CboPayMentType.locked = False
            CboPayMentType.ListIndex = 0
            ' Me.XPBtnMove(0).Enabled = False
            ' Me.XPBtnMove(1).Enabled = False
            ' Me.XPBtnMove(2).Enabled = False
            ' Me.XPBtnMove(3).Enabled = False
            XPBtnNewClients.Enabled = True
            Me.XPDtbBill.Enabled = True
            
            Me.DBCboClientName.locked = False
            Me.DCboStoreName.locked = False
            DcboEmp.Enabled = True
        
            '   CmdConvert.Visible = False
            CmdTemplate.Enabled = True
            CmdTemplate.Visible = True
            Ele(2).Enabled = True
           ' CboItemCase.ListIndex = 0

        Case "E"

           If SystemOptions.UserInterface = ArabicInterface Then
                Me.Caption = "  ⁄—Ì› „ﬂÊ‰«  «·«’‰«›/«· Ã„Ì⁄    "
            Else
                Me.Caption = "Definition of items / components assembly"
            End If

            Me.cmd(2).Enabled = True
            Me.cmd(3).Enabled = True
         '   Frame4.Enabled = True'
            'Ele(11).Enabled = True
   
            Me.cmd(0).Enabled = False
            Me.cmd(1).Enabled = False
            Me.cmd(4).Enabled = False
            Me.cmd(5).Enabled = False
            Me.cmd(7).Enabled = False
        
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
            CboPayMentType.locked = False
'             If CboPayMentType.ListIndex = 0 Then
'                 CboPayMentType_Change
'             End If
            ' CmdConvert.Visible = False
            CmdTemplate.Visible = False
            Ele(2).Enabled = True
            FG2.Enabled = True
    End Select

    Exit Sub
ErrTrap:
End Sub

Public Sub Retrive(Optional Lngid As Long = 0, Optional ByVal IsNotFixed As Boolean = True)
    Dim RsDetails As New ADODB.Recordset
    Dim StrSQL As String
    Dim RsNotes As New ADODB.Recordset
    Dim RsTemp As ADODB.Recordset
    Dim LngCurItemID As Long
    Dim LngUnitID As Long
    Dim DblQty As Double
      Dim ContactTime As Date
    Dim Num As Long


    'On Error GoTo ErrTrap
    If rs.RecordCount < 1 Then
        XPTxtCurrent.Caption = 0
        XPTxtCount.Caption = 0
        Exit Sub
    End If

    If rs.EOF Or rs.BOF Then
        Exit Sub
    End If

    If Lngid <> 0 Then
        rs.Find "ID=" & Lngid, , adSearchForward, adBookmarkFirst

        If rs.BOF Or rs.EOF Then
            Exit Sub
        End If
    End If
    
 txtCustomerName.backcolor = vbWhite
mNewId = 0
mIdDisplay = 0
Me.TxtModFlg.Text = "R"
    TxtTransSerial.Text = IIf(IsNull(rs("ID").value), "", val(rs("ID").value))
 XPDtbBill.value = IIf(IsNull(rs("RecordDate").value), Date, rs("RecordDate").value)
 XPDtRecDate.value = IIf(IsNull(rs("RecDate").value), Date, rs("RecDate").value)
 
DcboItemID4.Tag = ""

   dcBranch.BoundText = IIf(IsNull(rs("BranchID").value), "", rs("BranchID").value)
  Me.DCboStoreName.BoundText = IIf(IsNull(rs("StoreID").value), "", rs("StoreID").value)
   Me.DCboStore2Name.BoundText = IIf(IsNull(rs("StoreID2").value), "", rs("StoreID2").value)
     DCboStore3Name.BoundText = IIf(IsNull(rs("StoreID3").value), "", rs("StoreID3").value)
       DcboItemID1.BoundText = IIf(IsNull(rs("ItemNameID").value), "", rs("ItemNameID").value)
        DcbUnit.BoundText = IIf(IsNull(rs("UnitID").value), "", rs("UnitID").value)
        txtQty1.Text = IIf(IsNull(rs("Qty1").value), 0, rs("Qty1").value)
         TXTTransactionID1.Text = IIf(IsNull(rs("TransactionID1").value), "", rs("TransactionID1").value)
         TXTTransactionID2.Text = IIf(IsNull(rs("TransactionID2").value), "", rs("TransactionID2").value)
        
         TXTTransactionID5.Text = IIf(IsNull(rs("TransactionID5").value), "", rs("TransactionID5").value)
         
         TxtNoteSerial13.Text = IIf(IsNull(rs("NoteSerial13").value), "", rs("NoteSerial13").value)
         TxtNoteSerial15.Text = IIf(IsNull(rs("NoteSerial15").value), "", rs("NoteSerial15").value)
         TxtNoteSerial11.Text = IIf(IsNull(rs("NoteSerial11").value), "", rs("NoteSerial11").value)
         TxtNoteSerial12.Text = IIf(IsNull(rs("NoteSerial12").value), "", rs("NoteSerial12").value)
        txtNoteid3.Text = IIf(IsNull(rs("Noteid3").value), "", rs("Noteid3").value)

        TXT_order_no.Text = IIf(IsNull(rs("order_no").value), "", rs("order_no").value)
        txtOrderID.Text = IIf(IsNull(rs("OrderID").value), 0, rs("OrderID").value)
        

         DcboItemID1.Tag = ""
    CboPayMentType.ListIndex = IIf(IsNull(rs("PaymentType").value), 0, rs("PaymentType").value)

    
       TxtAttachedItemCode.Text = IIf(IsNull(rs("ItemCode").value), "", (rs("ItemCode").value))
     Me.TxtMaxNo.Text = IIf(IsNull(rs("MaxNo").value), "", (rs("MaxNo").value))
     Me.TxtMaxNo2.Text = IIf(IsNull(rs("MaxNo2").value), "", (rs("MaxNo2").value))
   
    Me.txtPeriod.Text = IIf(IsNull(rs("Period").value), "", (rs("Period").value))
     
       If Not IsNull(rs("RecTime").value) Then
      ContactTime = FormatDateTime(rs("RecTime").value, vbShortTime)
      
   
    End If
     Me.txtRecTime.value = ContactTime
    
    txtQty1.Text = IIf(IsNull(rs("Qty1").value), "", (rs("Qty1").value))
    txtwidtj.Text = IIf(IsNull(rs("widtj").value), "", (rs("widtj").value))
    txthight.Text = IIf(IsNull(rs("hight").value), "", (rs("hight").value))
    txtLength.Text = IIf(IsNull(rs("Length").value), "", (rs("Length").value))
    
        Me.TxtMaxName.Text = IIf(IsNull(rs("MaxName").value), "", (rs("MaxName").value))
    Me.DBCboClientName.BoundText = IIf(IsNull(rs("CusID").value), "", rs("CusID").value)
    Me.DCboUserName.BoundText = IIf(IsNull(rs("UserID").value), "", rs("UserID").value)
    
    If Not IsNull(rs("GroupID")) Then
        XPCboGroup.BoundText = rs("GroupID").value
    Else
        XPCboGroup.BoundText = ""
    End If
    
    

    XPCboGroupBuiltin.BoundText = IIf(IsNull(rs("GroupIDBuiltin").value), "", rs("GroupIDBuiltin").value)
    DcboBuiltinItemID.BoundText = IIf(IsNull(rs("BuiltinItemID").value), "", rs("BuiltinItemID").value)
    
    If IsNull(rs("BoxID").value) Then
        Me.DcboBox.BoundText = ""
    Else
        Me.DcboBox.BoundText = IIf(IsNull(rs("BoxID").value), "", rs("BoxID").value)
    End If
        
    TxtVAt2.Text = IIf(IsNull(rs("Vat2").value), "", rs("Vat2").value)
    txtTotalWithVat.Text = IIf(IsNull(rs("TotalWithVat").value), "", rs("TotalWithVat").value)
    
    TxtPrice.Text = IIf(IsNull(rs("Price").value), "", rs("Price").value)
    txtTotalAdd.Text = IIf(IsNull(rs("TotalAdd").value), "", rs("TotalAdd").value)
    txtTotalDisc.Text = IIf(IsNull(rs("TotalDisc").value), "", rs("TotalDisc").value)
    txtNet.Text = IIf(IsNull(rs("Net").value), "", rs("Net").value)
    Me.DcboEmp.BoundText = IIf(IsNull(rs("Emp_ID").value), "", rs("Emp_ID").value)
    
    TXTTransactionID3.Text = rs!TransactionID3 & ""
    TxtNoteSerial13.Text = rs!NoteSerial13 & ""
    txtNoteid3.Text = rs!Noteid3 & ""
   TXTTransactionID4.Text = rs!TransactionID4 & ""
    TxtNoteSerial14.Text = rs!NoteSerial14 & ""
    cmdCancel2.Visible = True
   If val(TXTTransactionID4) <> 0 Then
       'If Not SystemOptions.IsMultiItemsInCompItem Then
            cmdCancel2.Enabled = True
            cmdCreateProduction.Enabled = True
            If Not SystemOptions.UserInterface = EnglishInterface Then
                cmdCreateProduction.Caption = "⁄—÷ «„— «·«‰ «Ã"
                
            Else
                cmdCreateProduction.Caption = "Open the production order"
            End If
       ' End If
    Else
        If Not SystemOptions.UserInterface = EnglishInterface Then
            cmdCreateProduction.Caption = "«‰‘«¡ «„— «·«‰ «Ã"
        Else
            cmdCreateProduction.Caption = "Create a product order"
        End If
        cmdCancel2.Enabled = False
    End If
   If rs("Allocated").value = True Then
   Selct(0).value = vbChecked
   Selct(1).Enabled = True
   Selct(2).Enabled = True
   Else
   Selct(0).value = vbChecked
    Selct(1).Enabled = False
   Selct(2).Enabled = False
   End If
    If rs("AlloPay").value = True Then
    Selct(1).value = vbChecked
    Else
     Selct(1).value = vbChecked
    End If
       If rs("AlloRecep").value = True Then
       Selct(2).value = vbChecked
    Else
      Selct(2).value = vbChecked
    End If
    Selct(1).Enabled = True
    Selct(2).Enabled = True
    If TxtNoteSerial1 <> "" Then
    Selct(1).value = vbChecked
    End If
    If TxtNoteSerial12 <> "" Then
    Selct(2).value = vbChecked
    End If
    FG.Clear flexClearScrollable, flexClearEverything
    FG.Rows = 2
    FG.Clear flexClearScrollable, flexClearEverything
    FG.Refresh
    FGDeleted.Clear flexClearScrollable, flexClearEverything
    FGDeleted.Rows = 1

    StrSQL = "SELECT    DISTINCT T2.ItemName ItemName2,T2.ItemNamee ItemNamee2, ItemCode2,ItemID2, TblDefComItemDet.OldPrice,   TblDefComItemDet.lowering lowering2,TblDefComItemDet.increase increase2,dbo.TblDefComItemDet.ID,TblDefComItemDet.IsDeleted, dbo.TblDefComItemDet.IDDefCIT,dbo.TblDefComItemDet.IsAdd,dbo.TblDefComItemDet.Price,dbo.TblDefComItemDet.Total, dbo.TblDefComItemDet.ItemID, dbo.TblItems.ItemCode, dbo.TblItems.ItemName, "
    StrSQL = StrSQL & "                   dbo.TblItems.Fullcode ,dbo.TblItems.ItemNamee, dbo.TblDefComItemDet.UnitID, dbo.TblUnites.UnitName, dbo.TblUnites.UnitNamee, dbo.TblDefComItemDet.SpecID1,"
    StrSQL = StrSQL & "                  dbo.TblSpecification.Name AS Name1, dbo.TblSpecification.Namee AS Namee1, dbo.TblDefComItemDet.SpecID2, TblSpecification_1.Name AS Name2,"
    StrSQL = StrSQL & "                  TblSpecification_1.Namee AS Namee2, dbo.TblDefComItemDet.SpecID3, TblSpecification_2.Name AS Name3, TblSpecification_2.Namee AS Namee3,"
    StrSQL = StrSQL & "                  dbo.TblDefComItemDet.SpecID4, TblSpecification_3.Name AS Name4, TblSpecification_3.Namee AS Namee4, dbo.TblDefComItemDet.Amout1,"
    StrSQL = StrSQL & "                 dbo.TblDefComItemDet.Amout2 ,TblDefComItemDet.LineID ,dbo.TblDefComItemDet.Amout3, dbo.TblDefComItemDet.Amout4, dbo.TblDefComItemDet.Qty, dbo.TblDefComItemDet.cost ,dbo.TblDefComItemDet.FlgX ,dbo.TblDefComItemDet.TepQty,"
    StrSQL = StrSQL & "                 IsNull(TblDefComItemDet.IsRow,0) IsRow,TblDefComItemDet.widtj,TblDefComItemDet.hight,TblDefComItemDet.Length,TblDefComItemDet.thickness,TblDefComItemDet.DO,TblDefComItemDet.DI,TblDefComItemDet.Diameter,"
    StrSQL = StrSQL & "                 dbo.TblItemsParts.PartItemQty ,TblItemsParts.TableID, ForUnit,    TblItemsParts.lowering,TblItemsParts.increase,TblItemsParts.MethodCalc"
    StrSQL = StrSQL & " FROM         dbo.TblDefComItemDet LEFT OUTER JOIN"
    StrSQL = StrSQL & "                   dbo.TblSpecification TblSpecification_3 ON dbo.TblDefComItemDet.SpecID4 = TblSpecification_3.ID LEFT OUTER JOIN"
    StrSQL = StrSQL & "                 dbo.TblSpecification TblSpecification_2 ON dbo.TblDefComItemDet.SpecID3 = TblSpecification_2.ID LEFT OUTER JOIN"
    StrSQL = StrSQL & "                  dbo.TblSpecification TblSpecification_1 ON dbo.TblDefComItemDet.SpecID2 = TblSpecification_1.ID LEFT OUTER JOIN"
    StrSQL = StrSQL & "                 dbo.TblSpecification ON dbo.TblDefComItemDet.SpecID1 = dbo.TblSpecification.ID LEFT OUTER JOIN"
    StrSQL = StrSQL & "                  dbo.TblUnites ON dbo.TblDefComItemDet.UnitID = dbo.TblUnites.UnitID LEFT OUTER JOIN"
    StrSQL = StrSQL & "                  dbo.TblItems ON dbo.TblDefComItemDet.ItemID = dbo.TblItems.ItemID"
    
    StrSQL = StrSQL & "                  Left OUTER JOIN "
  
        
    


    StrSQL = StrSQL & "      dbo.TblItems T2 ON dbo.TblDefComItemDet.ItemID2 = T2.ItemID"
    StrSQL = StrSQL & "      LEFT OUTER JOIN dbo.TblItemsParts"
    StrSQL = StrSQL & "                  ON  dbo.TblItemsParts.ItemID = TblDefComItemDet.ItemID2"
    StrSQL = StrSQL & "                       and   TblItemsParts.PartItemID = TblDefComItemDet.ItemID"
    StrSQL = StrSQL & "                       and   TblItemsParts.UnitID = TblDefComItemDet.UnitID"
    StrSQL = StrSQL & " Where (dbo.TblDefComItemDet.IDDefCIT =" & val(TxtTransSerial.Text) & ")"
    
    StrSQL = StrSQL & " Order By TblDefComItemDet.ItemID2,TblDefComItemDet.LineID,TblDefComItemDet.ID"

 



    RsDetails.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText

Dim mm  As Long
Dim mTableID As String
Dim mUnitId As Long
Dim mUnitName As String
Dim rsDummy3 As New ADODB.Recordset
mTableID = "(0,0"
    If Not (RsDetails.EOF Or RsDetails.BOF) Then
        FG.Rows = RsDetails.RecordCount + 1
        
                  Dim rsDummy2 As New ADODB.Recordset
          Dim PartItemQty As Double, ForUnit As Double, lowering As Double, increase As Double, MethodCalc As Double
        For Num = 1 To RsDetails.RecordCount
        
'
'            StrSQL = "select    dbo.TblItemsParts.PartItemQty ,TableID, ForUnit,    TblItemsParts.lowering,TblItemsParts.increase,MethodCalc from TblItemsParts"
'            StrSQL = StrSQL & " Where dbo.TblItemsParts.ItemID = " & val(RsDetails!ItemID2 & "")
'            StrSQL = StrSQL & " and PartItemID = " & val(RsDetails!ItemID & "")
'            StrSQL = StrSQL & " and UnitID = " & val(RsDetails!UnitID & "")
'         '    StrSQL = StrSQL & " and TableID Not In  " & mTableID & ")"
'            Set rsDummy2 = New ADODB.Recordset
'            rsDummy2.Open StrSQL, Cn, adOpenStatic, adLockReadOnly
'
       
            If val(RsDetails!TableID & "") <> 0 Then
          '  If Not rsDummy2.EOF Then
              ' If val(RsDetails!PartItemQty & "") = 0 Then
                    PartItemQty = val(RsDetails!PartItemQty & "")
             '   Else
             '   PartItemQty = val(RsDetails!PartItemQty & "")
             '   End If
            '    If val(RsDetails!ForUnit & "") = 0 Then
                    ForUnit = val(RsDetails!ForUnit & "")
              '  Else
              '      ForUnit = val(RsDetails!ForUnit & "")
              '  End If
                FG.TextMatrix(Num, FG.ColIndex("TableID")) = val(RsDetails!TableID & "")
                    If mTableID = "" Then
               mTableID = "(" & FG.TextMatrix(Num, FG.ColIndex("TableID"))
            Else
                mTableID = mTableID & "," & FG.TextMatrix(Num, FG.ColIndex("TableID"))
            End If
                If val(RsDetails!lowering2 & "") = 0 Then
                    lowering = val(RsDetails!lowering & "")
                Else
                    lowering = val(RsDetails!lowering & "")
                End If

                If val(RsDetails!increase2 & "") = 0 Then
                    increase = val(RsDetails!increase & "")
                Else
                    increase = val(RsDetails!increase2 & "")
                End If
                If CBool(RsDetails!IsRow & "") Then
                    
                    StrSQL = " SELECT IsNull(MethodCalc,99) MethodCalc,IsNull(PartItemQty,99) PartItemQty,IsNull(ForUnit ,99)  ForUnit  FROM TblItemsUnits"
                    StrSQL = StrSQL & " WHERE ItemID =" & val(RsDetails!ItemID & "")
                    StrSQL = StrSQL & " AND UnitID =" & val(RsDetails!UnitID & "")
                    rsDummy3.Open StrSQL, Cn, adOpenKeyset, adLockReadOnly
                    If Not rsDummy3.EOF Then
                        MethodCalc = IIf(val(rsDummy3!MethodCalc & "") <> 99, val(rsDummy3!MethodCalc & ""), val(rsDummy2!MethodCalc & ""))
                        PartItemQty = IIf(val(rsDummy3!PartItemQty & "") <> 99, val(rsDummy3!PartItemQty & ""), val(rsDummy2!PartItemQty & ""))
                        ForUnit = IIf(val(rsDummy3!ForUnit & "") <> 99, val(rsDummy3!ForUnit & ""), val(rsDummy2!ForUnit & ""))
                    End If
                Else
               ' If val(RsDetails!MethodCalc & "") = 0 Then
                    MethodCalc = val(RsDetails!MethodCalc & "")
                End If
                
                If MethodCalc = 0 Then MethodCalc = 1
              '  Else
               ' MethodCalc = val(RsDetails!MethodCalc & "")
               ' End If

               
                

            End If
            If val(FG.TextMatrix(Num, FG.ColIndex("TableID"))) = 0 Then
                Num = Num
            End If
           '   rsDummy2.Close
            FG.TextMatrix(Num, FG.ColIndex("Ser")) = Num
            FG.TextMatrix(Num, FG.ColIndex("FlgX")) = IIf(IsNull(RsDetails("FlgX").value), "", Trim(RsDetails("FlgX").value))
            FG.TextMatrix(Num, FG.ColIndex("SpecID4")) = IIf(IsNull(RsDetails("SpecID4").value), "", Trim(RsDetails("SpecID4").value))
            FG.TextMatrix(Num, FG.ColIndex("SpecID3")) = IIf(IsNull(RsDetails("SpecID3").value), "", (RsDetails("SpecID3").value))
            FG.TextMatrix(Num, FG.ColIndex("SpecID2")) = IIf(IsNull(RsDetails("SpecID2").value), "", (RsDetails("SpecID2").value))
            FG.TextMatrix(Num, FG.ColIndex("Fullcode")) = IIf(IsNull(RsDetails("Fullcode").value), "", (RsDetails("Fullcode").value))
        
            FG.TextMatrix(Num, FG.ColIndex("widtj")) = RsDetails("widtj").value & ""
            FG.TextMatrix(Num, FG.ColIndex("hight")) = RsDetails("hight").value & ""
            FG.TextMatrix(Num, FG.ColIndex("Length")) = RsDetails("Length").value & ""
            FG.TextMatrix(Num, FG.ColIndex("thickness")) = RsDetails("thickness").value & ""
            FG.TextMatrix(Num, FG.ColIndex("DO")) = RsDetails("DO").value & ""
            FG.TextMatrix(Num, FG.ColIndex("DI")) = RsDetails("DI").value & ""
            FG.TextMatrix(Num, FG.ColIndex("IsRow")) = IIf(IsNull(RsDetails("IsRow").value), 0, (RsDetails("IsRow").value))
          
            
                        
    
        
        
            FG.TextMatrix(Num, FG.ColIndex("SpecID1")) = IIf(IsNull(RsDetails("SpecID1").value), "", (RsDetails("SpecID1").value))
            FG.TextMatrix(Num, FG.ColIndex("ItemID")) = IIf(IsNull(RsDetails("ItemID").value), "", (RsDetails("ItemID").value))
            
            FG.TextMatrix(Num, FG.ColIndex("ItemID2")) = IIf(IsNull(RsDetails("ItemID2").value), "", (RsDetails("ItemID2").value))
            FG.TextMatrix(Num, FG.ColIndex("ItemCode2")) = IIf(IsNull(RsDetails("ItemCode2").value), "", (RsDetails("ItemCode2").value))
           If SystemOptions.UserInterface = ArabicInterface Then
                FG.TextMatrix(Num, FG.ColIndex("ItemName2")) = IIf(IsNull(RsDetails("ItemName2").value), "", (RsDetails("ItemName2").value))
            Else
                FG.TextMatrix(Num, FG.ColIndex("ItemName2")) = IIf(IsNull(RsDetails("ItemNamee2").value), "", (RsDetails("ItemNamee2").value))
            End If
            FG.TextMatrix(Num, FG.ColIndex("LineID")) = IIf(IsNull(RsDetails("LineID").value), "", (RsDetails("LineID").value))
            FG.TextMatrix(Num, FG.ColIndex("UnitID")) = IIf(IsNull(RsDetails("UnitID").value), "", (RsDetails("UnitID").value))
            
          
            FG.TextMatrix(Num, FG.ColIndex("PartItemQty")) = PartItemQty
            FG.TextMatrix(Num, FG.ColIndex("ForUnit")) = ForUnit
            FG.TextMatrix(Num, FG.ColIndex("MethodCalc")) = MethodCalc
            FG.TextMatrix(Num, FG.ColIndex("lowering")) = lowering
            FG.TextMatrix(Num, FG.ColIndex("Increase")) = increase

            
            FG.TextMatrix(Num, FG.ColIndex("itemcode")) = IIf(IsNull(RsDetails("ItemCode").value), "", (RsDetails("ItemCode").value))
            FG.TextMatrix(Num, FG.ColIndex("cost")) = IIf(IsNull(RsDetails("cost").value), "", (RsDetails("cost").value))
            FG.TextMatrix(Num, FG.ColIndex("Qty")) = IIf(IsNull(RsDetails("Qty").value), "", (RsDetails("Qty").value))
            FG.TextMatrix(Num, FG.ColIndex("TepQty")) = IIf(IsNull(RsDetails("TepQty").value), val(FG.TextMatrix(Num, FG.ColIndex("Qty"))), Trim(RsDetails("TepQty").value))
           If SystemOptions.UserInterface = EnglishInterface Then
               ' FG.Cell(flexcpData, Num, FG.ColIndex("itemname")) = IIf(IsNull(RsDetails("ItemNamee").value), "", (RsDetails("ItemNamee").value))
            FG.TextMatrix(Num, FG.ColIndex("unitname")) = IIf(IsNull(RsDetails("UnitNamee").value), "", (RsDetails("UnitNamee").value))
            FG.TextMatrix(Num, FG.ColIndex("name1")) = IIf(IsNull(RsDetails("Namee1").value), "", (RsDetails("Namee1").value))
            FG.TextMatrix(Num, FG.ColIndex("name2")) = IIf(IsNull(RsDetails("Namee2").value), "", (RsDetails("Namee2").value))
            FG.TextMatrix(Num, FG.ColIndex("name3")) = IIf(IsNull(RsDetails("Namee3").value), "", (RsDetails("Namee3").value))
            FG.TextMatrix(Num, FG.ColIndex("name4")) = IIf(IsNull(RsDetails("Namee4").value), "", (RsDetails("Namee4").value))
            FG.TextMatrix(Num, FG.ColIndex("itemname")) = IIf(IsNull(RsDetails("ItemNamee").value), "", (RsDetails("ItemNamee").value))

       Else
            'FG.Cell(flexcpData, Num, FG.ColIndex("itemname")) = IIf(IsNull(RsDetails("ItemName").value), "", (RsDetails("ItemName").value))
            FG.TextMatrix(Num, FG.ColIndex("unitname")) = IIf(IsNull(RsDetails("UnitName")), "", (RsDetails("UnitName").value))
            FG.TextMatrix(Num, FG.ColIndex("name1")) = IIf(IsNull(RsDetails("name1").value), "", (RsDetails("name1").value))
            FG.TextMatrix(Num, FG.ColIndex("name2")) = IIf(IsNull(RsDetails("name2").value), "", (RsDetails("name2").value))
            FG.TextMatrix(Num, FG.ColIndex("name3")) = IIf(IsNull(RsDetails("name3").value), "", (RsDetails("name3").value))
            FG.TextMatrix(Num, FG.ColIndex("name4")) = IIf(IsNull(RsDetails("name4").value), "", (RsDetails("name4").value))
         FG.TextMatrix(Num, FG.ColIndex("itemname")) = IIf(IsNull(RsDetails("ItemName").value), "", (RsDetails("ItemName").value))

       
    End If
            If val(FG.TextMatrix(Num, FG.ColIndex("UnitID"))) = 0 Then
                FG.TextMatrix(Num, FG.ColIndex("UnitID")) = GetDefaultItemUnit(val(FG.TextMatrix(Num, FG.ColIndex("ItemID"))), mUnitId, mUnitName)
                FG.TextMatrix(Num, FG.ColIndex("UnitID")) = mUnitId
                FG.TextMatrix(Num, FG.ColIndex("unitname")) = mUnitName
            End If
            FG.TextMatrix(Num, FG.ColIndex("IsDeleted")) = IIf(IsNull(RsDetails("IsDeleted").value), 0, IIf((RsDetails("IsDeleted").value), -1, 0))
            FG.TextMatrix(Num, FG.ColIndex("IsAdd")) = IIf(IsNull(RsDetails("IsAdd").value), 0, (RsDetails("IsAdd").value))
           
           If val(RsDetails!Price & "") = 0 And IsSaveWithOutMsg Then
                CalcTotal Num
            Else
               FG.TextMatrix(Num, FG.ColIndex("Price")) = IIf(IsNull(RsDetails("Price").value), "", (RsDetails("Price").value))
                FG.TextMatrix(Num, FG.ColIndex("Total")) = IIf(IsNull(RsDetails("Total").value), "", (RsDetails("Total").value))
            End If
            FG.TextMatrix(Num, FG.ColIndex("OldPrice")) = IIf(IsNull(RsDetails("OldPrice").value), "", (RsDetails("OldPrice").value))
            
           
            
            FG.TextMatrix(Num, FG.ColIndex("Amout1")) = IIf(IsNull(RsDetails("Amout1").value), "", (RsDetails("Amout1").value))
            FG.TextMatrix(Num, FG.ColIndex("Amout2")) = IIf(IsNull(RsDetails("Amout2").value), "", (RsDetails("Amout2").value))
            FG.TextMatrix(Num, FG.ColIndex("Amout3")) = IIf(IsNull(RsDetails("Amout3").value), "", (RsDetails("Amout3").value))
            FG.TextMatrix(Num, FG.ColIndex("Amout4")) = IIf(IsNull(RsDetails("Amout4").value), "", (RsDetails("Amout4").value))

            
            If IIf(IsNull(RsDetails("IsDeleted").value), False, (RsDetails("IsDeleted").value)) Then
                FG.RowHidden(Num) = True
                'mmmm = (RsDetails("ItemID").value)
                
            Else
                FG.RowHidden(Num) = False
            End If
                    
            
            RsDetails.MoveNext
           

        Next Num
        FillDelGrid
        FG.AutoSize 0, FG.Cols - 1, False
    End If
    

    CalcTotalNet
    
        XPTxtCurrent.Caption = rs.AbsolutePosition
    XPTxtCount.Caption = rs.RecordCount
   
    Dim s As String
    s = "Select UserId From  TblProductLineDistribution Where IDDefCIT = " & val(TxtTransSerial) & " "
    Dim RsData As ADODB.Recordset
    Set RsData = New ADODB.Recordset
    RsData.Open s, Cn, adOpenStatic, adLockReadOnly, adCmdText
    If Not RsData.EOF Or val(TxtNoteSerial13) <> 0 Then
        cmdTransfer.Enabled = False
        cmdCancel.Enabled = True
    Else
        cmdTransfer.Enabled = True
        cmdCancel.Enabled = False
    End If
    

        s = " Select T.*,"
    If SystemOptions.UserInterface = ArabicInterface Then
        s = s & "     ti.ItemName,t3.ItemName BuiltinItemName,g3.GroupName GroupBuiltinName  "
        s = s & "     ,tu.UnitName,G.GroupName,ti2.ItemName ItemName2,ti5.ItemName ItemName5"
    Else
        s = s & "     ti.ItemNamee ItemName,t3.ItemNamee BuiltinItemName,g3.GroupNamee GroupBuiltinName  "
        s = s & "     ,tu.UnitNamee UnitName ,G.GroupNamee GroupName,ti2.ItemNamee ItemName2,ti5.ItemNamee ItemName5"
    End If
    s = s & "  from TblDefComItemData T"
    s = s & " LEFT OUTER JOIN Groups AS g ON g.GroupID =T.GroupID"
    s = s & " LEFT OUTER JOIN Groups AS g3 ON g3.GroupID =T.GroupIDBuiltin"
    s = s & " LEFT OUTER JOIN TblItems AS t3  ON t3.ItemID =T.BuiltinItemID"
    
    s = s & " LEFT OUTER JOIN TblItems AS ti  ON ti.ItemID =T.ItemID"
    s = s & " LEFT OUTER JOIN TblItems AS ti2  ON ti2.ItemID =T.ItemID2"
    s = s & " LEFT OUTER JOIN TblItems AS ti5  ON ti5.ItemID =T.ItemID5"
    s = s & " LEFT OUTER JOIN TblUnites AS tu  ON tu.UnitId =T.UnitId"
    s = s & " Where (T.IDDefCIT =" & val(TxtTransSerial.Text) & ")"
    s = s & " Order By T.ID"
    loadgrid s, FG2, True, False
    
    If SystemOptions.DontCreateOut Then
        s = "    SELECT t.Transaction_ID,"
        s = s & "          t.NoteSerial1,"
        s = s & "          t.NoteSerial,"
        s = s & "              t.NoteId,"
        s = s & "              t.Transaction_Date,"
        s = s & "              td.ShowQty,"
        s = s & "              td.showPrice,td.RemarksLine,"
        s = s & "              Total = td.ShowQty * td.showPrice,"
        s = s & "              ti.ItemName"
        s = s & "       FROM   Transactions         AS t"
        s = s & "              INNER JOIN Transaction_Details AS td"
        s = s & "                   ON  td.Transaction_ID = t.Transaction_ID"
        s = s & "              INNER JOIN TblItems  AS ti"
        s = s & "                   ON  ti.ItemID = td.Item_ID"
        s = s & "       Where t.Transaction_Type = 27"
        s = s & "              AND ISNULL(t.IDDefCIT, 0) = " & val(TxtTransSerial)
        loadgrid s, FG3, True, False
    End If
    
    
    If IsSaveWithOutMsg Then
        If IsNotFixed Then
            FixRowsLine
        End If
    End If
    
    CalcGrid2 True
mIsFinishSave = True
  '  cmdCreateProduction.Enabled = True
    Exit Sub
ErrTrap:
    Screen.MousePointer = vbDefault
End Sub



Private Sub AddTip()
    Dim Wrap As String
    On Error GoTo ErrTrap
    Wrap = CHR(13) + CHR(10)
    Set TTP = New clstooltip

    With TTP
        .Create Me.hwnd, "   ‘«‘…  ⁄—Ì› „ﬂÊ‰«  «·«’‰«›/«· Ã„Ì⁄       ", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
        .AddControl cmd(0), "ÃœÌœ ..." & Wrap & "·«÷«›…  „ﬂÊ‰«  «·«’‰«› ÃœÌœÂ" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
    End With

    With TTP
        .Create Me.hwnd, " ‘«‘…  ⁄—Ì› „ﬂÊ‰«  «·«’‰«›/«· Ã„Ì⁄     ", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
        .AddControl cmd(7), "ÿ»«⁄… ..." & Wrap & "·⁄—÷  ﬁ—Ì— „ﬂÊ‰«  «·«’‰«› «·Õ«·Ì… " & Wrap & " Ì„ﬂ‰ ÿ»«⁄ Â ⁄‰ ÿ—Ìﬁ «·ÿ«»⁄…", True
    End With

    With TTP
        .Create Me.hwnd, " ‘«‘…  ⁄—Ì› „ﬂÊ‰«  «·«’‰«›/«· Ã„Ì⁄ ", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
        .AddControl cmd(1), " ⁄œÌ· ..." & Wrap & "· ⁄œÌ· „ﬂÊ‰«  «·«’‰«›      «·Õ«·Ì" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
    End With

    With TTP
        .Create Me.hwnd, "   ‘«‘…  ⁄—Ì› „ﬂÊ‰«  «·«’‰«›/«· Ã„Ì⁄   ", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
        .AddControl cmd(2), "Õ›Ÿ ..." & Wrap & "·Õ›Ÿ »Ì«‰«  „ﬂÊ‰«  «·«’‰«›         «·ÃœÌœ" & Wrap & "·Õ›Ÿ «· ⁄œÌ·« " & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
    End With

    With TTP
        .Create Me.hwnd, "   ‘«‘…  ⁄—Ì› „ﬂÊ‰«  «·«’‰«›/«· Ã„Ì⁄  ", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
        .AddControl cmd(3), " —«Ã⁄ ..." & Wrap & "·· —«Ã⁄ ⁄‰ ⁄„·Ì… «·≈÷«›…" & Wrap & "··· —«Ã⁄ ⁄‰ ⁄„·Ì… «· ⁄œÌ·" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
    End With

    With TTP
        .Create Me.hwnd, "  ‘«‘…  ⁄—Ì› „ﬂÊ‰«  «·«’‰«›/«· Ã„Ì⁄    ", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
        .AddControl cmd(4), "Õ–› ..." & Wrap & "·Õ–› »Ì«‰«  «·⁄—÷ «·Õ«·Ì" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
    End With

    With TTP
        .Create Me.hwnd, "  ‘«‘…  ⁄—Ì› „ﬂÊ‰«  «·«’‰«›/«· Ã„Ì⁄      ", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
        .AddControl cmd(5), "»ÕÀ ..." & Wrap & "···»ÕÀ ⁄‰  „ﬂÊ‰«  «·«’‰«›   " & Wrap & "Ì‰ÿ»ﬁ ⁄·ÌÂ« ‘—Êÿ „⁄Ì‰…" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
    End With

    With TTP
        .Create Me.hwnd, "  ‘«‘…  ⁄—Ì› „ﬂÊ‰«  «·«’‰«›/«· Ã„Ì⁄ ", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
        .AddControl cmd(6), "Œ—ÊÃ ..." & Wrap & "·«€·«ﬁ Â–Â «·‰«›–…" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
    End With

    With TTP
        .Create Me.hwnd, "   ‘«‘…  ⁄—Ì› „ﬂÊ‰«  «·«’‰«›/«· Ã„Ì⁄   ", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
        .AddControl XPBtnNewClients, "≈÷«›…  „ﬂÊ‰«  «·«’‰«›     ÃœÌœ ..." & Wrap & "· ”ÃÌ· „ﬂÊ‰«  «·«’‰«›     ÃœÌœ" & Wrap & " «÷€ÿ Â‰«", True
    End With

    With TTP
        .Create Me.hwnd, "    ‘«‘…  ⁄—Ì› „ﬂÊ‰«  «·«’‰«›/«· Ã„Ì⁄   ", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
        .AddControl XPBtnMove(1), "«·√Ê· ..." & Wrap & "··«‰ ﬁ«· «·Ï √Ê· ”Ã·" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
    End With

    With TTP
        .Create Me.hwnd, "  ‘«‘…  ⁄—Ì› „ﬂÊ‰«  «·«’‰«›/«· Ã„Ì⁄    ", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
        .AddControl XPBtnMove(0), "«·”«»ﬁ ..." & Wrap & "··«‰ ﬁ«· «·Ï «·”Ã· «·”«»ﬁ" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
    End With

    With TTP
        .Create Me.hwnd, "    ‘«‘…  ⁄—Ì› „ﬂÊ‰«  «·«’‰«›/«· Ã„Ì⁄  ", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
        .AddControl XPBtnMove(3), "«· «·Ì ..." & Wrap & "··«‰ ﬁ«· «·Ï «·”Ã· «· «·Ì" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
    End With

    With TTP
        .Create Me.hwnd, "    ‘«‘…  ⁄—Ì› „ﬂÊ‰«  «·«’‰«›/«· Ã„Ì⁄  ", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
        .AddControl XPBtnMove(2), "«·√ŒÌ— ..." & Wrap & "··«‰ ﬁ«· «·Ï ¬Œ— ”Ã·" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
    End With

    With TTP
        .Create Me.hwnd, "    ‘«‘…  ⁄—Ì› „ﬂÊ‰«  «·«’‰«›/«· Ã„Ì⁄ ", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
        .AddControl CmdHelp, "„”«⁄œ… ..." & Wrap & "·· ⁄—› ⁄·Ï ÊŸÌ›… Â–Â «·‰«›–…" & Wrap & "ÊﬂÌ›Ì… «· ⁄«„· „⁄Â«" & Wrap & "≈÷€ÿ Â‰«" & Wrap, True
    End With

    Exit Sub
ErrTrap:
End Sub

 


Function print_report(Optional NoteSerial As String, Optional indexe As Integer)
    
     
    Dim MySQL As String
    Dim RsData As New ADODB.Recordset
    Dim xApp As New CRAXDRT.Application
    Dim xReport As CRAXDRT.Report
    Dim CViewer As ClsReportViewer
    Dim StrReportTitle As String
    Dim StrFileName As String
    Dim Msg As String
'MySQL = " SELECT     dbo.TblDefComItemDet.ID, dbo.TblDefComItemDet.IDDefCIT, dbo.TblDefComItemDet.ItemID, dbo.TblItems.ItemCode, dbo.TblItems.ItemName, "
'MySQL = MySQL & "                      dbo.TblItems.ItemNamee, dbo.TblDefComItemDet.UnitID, dbo.TblUnites.UnitName, dbo.TblUnites.UnitNamee, dbo.TblDefComItemDet.SpecID1,"
'MySQL = MySQL & "                      TblSpecification_3.Name AS Name1, TblSpecification_3.Namee AS Namee1, dbo.TblDefComItemDet.SpecID2, TblSpecification_1.Name AS Name2,"
'MySQL = MySQL & "                      TblSpecification_1.Namee AS Namee2, dbo.TblDefComItemDet.SpecID3, TblSpecification_2.Name AS Name3, TblSpecification_2.Namee AS Namee3,"
'MySQL = MySQL & "                      dbo.TblDefComItemDet.SpecID4, TblSpecification_3.Name AS Name4, TblSpecification_3.Namee AS Namee4, dbo.TblDefComItemDet.Amout1,"
'MySQL = MySQL & "                      dbo.TblDefComItemDet.Amout2, dbo.TblDefComItemDet.Amout3, dbo.TblDefComItemDet.Amout4, dbo.TblDefComItemDet.Qty, dbo.TblDefComItemDet.cost,"
'MySQL = MySQL & "                      dbo.TblDefComItem.RecordDate, dbo.TblDefComItem.StoreID, TblStore_2.StoreName, TblStore_2.StoreNamee, dbo.TblDefComItem.StoreID2,"
'MySQL = MySQL & "                      TblStore_1.StoreName AS StoreNam2, TblStore_1.StoreNamee AS StoreNamee3, dbo.TblDefComItem.StoreID3, TblStore_2.StoreName AS StoreName3,"
'MySQL = MySQL & "                      TblStore_2.StoreNamee AS StoreNamee4, dbo.TblDefComItem.CusID, dbo.TblCustemers.CusName, dbo.TblCustemers.CusNamee, dbo.TblDefComItem.MaxNo,"
'MySQL = MySQL & "                      dbo.TblDefComItem.MaxName, dbo.TblDefComItem.Allocated, dbo.TblDefComItem.AlloPay, dbo.TblDefComItem.AlloRecep, dbo.TblDefComItem.ID AS IDMain,"
'MySQL = MySQL & "                      dbo.TblDefComItem.BranchID , dbo.TblBranchesData.branch_name, dbo.TblBranchesData.branch_nameE"
'MySQL = MySQL & " FROM         dbo.TblItems RIGHT OUTER JOIN"
'MySQL = MySQL & "                      dbo.TblSpecification TblSpecification_3 RIGHT OUTER JOIN"
'MySQL = MySQL & "                      dbo.TblBranchesData RIGHT OUTER JOIN"
'MySQL = MySQL & "                      dbo.TblDefComItemDet RIGHT OUTER JOIN"
'MySQL = MySQL & "                      dbo.TblDefComItem ON dbo.TblDefComItemDet.IDDefCIT = dbo.TblDefComItem.ID ON"
'MySQL = MySQL & "                      dbo.TblBranchesData.branch_id = dbo.TblDefComItem.BranchID LEFT OUTER JOIN"
'MySQL = MySQL & "                      dbo.TblCustemers ON dbo.TblDefComItem.CusID = dbo.TblCustemers.CusID LEFT OUTER JOIN"
'MySQL = MySQL & "                      dbo.TblStore TblStore_2 ON dbo.TblDefComItem.StoreID3 = TblStore_2.StoreID LEFT OUTER JOIN"
'MySQL = MySQL & "                      dbo.TblStore TblStore_3 ON dbo.TblDefComItem.StoreID = TblStore_3.StoreID LEFT OUTER JOIN"
'MySQL = MySQL & "                      dbo.TblStore TblStore_1 ON dbo.TblDefComItem.StoreID2 = TblStore_1.StoreID ON TblSpecification_3.ID = dbo.TblDefComItemDet.SpecID4 LEFT OUTER JOIN"
'MySQL = MySQL & "                      dbo.TblSpecification TblSpecification_2 ON dbo.TblDefComItemDet.SpecID3 = TblSpecification_2.ID LEFT OUTER JOIN"
'MySQL = MySQL & "                      dbo.TblSpecification TblSpecification_1 ON dbo.TblDefComItemDet.SpecID2 = TblSpecification_1.ID LEFT OUTER JOIN"
'MySQL = MySQL & "                      dbo.TblSpecification TblSpecification_4 ON dbo.TblDefComItemDet.SpecID1 = TblSpecification_4.ID LEFT OUTER JOIN"
'MySQL = MySQL & "                      dbo.TblUnites ON dbo.TblDefComItemDet.UnitID = dbo.TblUnites.UnitID ON dbo.TblItems.ItemID = dbo.TblDefComItemDet.ItemID"
'MySQL = MySQL & "  Where (dbo.TblDefComItem.id = " & val(Me.TxtTransSerial.Text) & ")"
'
'        If SystemOptions.UserInterface = ArabicInterface Then
'            StrFileName = App.path & "\REPORTS\REPORTS NEW\" & "RepDefCompItem.rpt"
'        Else
'            StrFileName = App.path & "\REPORTS\REPORTS NEW\" & "RepDefCompIteme.rpt"
'        End If

        ''''''


    Dim StrSQL As String
    Dim StrWhere As String


   
    
    
 '   Dim Msg As String




    

MySQL = " SELECT TblDefComItem.ID,"
MySQL = MySQL & "         Grou.GroupName,"
MySQL = MySQL & "         TblDefComItem.PaymentType,"
MySQL = MySQL & "         tdcid.Vat2,"
MySQL = MySQL & "         tdcid.TotalWithVat,"
MySQL = MySQL & "         TblDefComItem.id              Transaction_ID,"
MySQL = MySQL & "         tdcid.Qty Qty1,TblDefComItem.Period TransactionID2,"
MySQL = MySQL & "         dbo.TblItems.ItemCode,"
MySQL = MySQL & "         dbo.TblItems.ItemName,"
MySQL = MySQL & "         Item5.ItemName  BuiltInItemName  ,"
MySQL = MySQL & "         dbo.TblItems.ItemNamee,Item2.ItemName ItemName2,"
MySQL = MySQL & "         dbo.TblDefComItem.RecordDate,"
MySQL = MySQL & "         dbo.TblDefComItem.CusID,"
MySQL = MySQL & "         dbo.TblCustemers.FullCode,TblCustemers.Address,TblCustemers.Mobile1,TblCustemers.VATNO,TblCustemers.ZipCode , '" & DcboEmp.Text & "' as ResponsibleContact,"
MySQL = MySQL & "         dbo.TblCustemers.CusName,"
MySQL = MySQL & "         dbo.TblCustemers.CusNamee,"
MySQL = MySQL & "         dbo.TblDefComItem.BranchID,"
MySQL = MySQL & "         dbo.TblBranchesData.branch_name,"
MySQL = MySQL & "         dbo.TblBranchesData.branch_nameE,"
MySQL = MySQL & "         tdcid.ItemID,"
MySQL = MySQL & "         tdcid.widtj,"
MySQL = MySQL & "         tdcid.hight,"
MySQL = MySQL & "         tdcid.Price,"
MySQL = MySQL & "         tdcid.TotalAdd,tdcid.Remark as MaxName,"
MySQL = MySQL & "         tdcid.TotalDisc,"
MySQL = MySQL & "         tdcid.Net,tdcid.Vat2 Vat22,"
MySQL = MySQL & "         tu.UnitName UnitName2"
MySQL = MySQL & "                      ,BalnceCust = " & val(Balance) & ",tdcid.AreaL"
MySQL = MySQL & "  From dbo.TblItems"
MySQL = MySQL & "         RIGHT OUTER JOIN dbo.TblBranchesData"
MySQL = MySQL & "         RIGHT OUTER JOIN dbo.TblDefComItem"
MySQL = MySQL & "              ON  dbo.TblBranchesData.branch_id = dbo.TblDefComItem.BranchID"
MySQL = MySQL & "         LEFT OUTER JOIN dbo.TblDefComItemData AS tdcid"
MySQL = MySQL & "         ON             tdcid.IDDefCIT = TblDefComItem.ID"
MySQL = MySQL & "         LEFT OUTER JOIN dbo.TblCustemers"
MySQL = MySQL & "              ON  dbo.TblDefComItem.CusID = dbo.TblCustemers.CusID"
MySQL = MySQL & "              ON  dbo.TblItems.ItemID = tdcid.ItemID"
MySQL = MySQL & "         LEFT OUTER JOIN TblUnites  AS tu"
MySQL = MySQL & "              ON  tu.UnitID = dbo.TblDefComItem.UnitID"
MySQL = MySQL & "         LEFT OUTER JOIN Groups     AS Grou"
MySQL = MySQL & "              ON  Grou.GroupID = tdcid.GroupID"
MySQL = MySQL & "         LEFT OUTER JOIN TblItems     AS Item2"
MySQL = MySQL & "              ON  Item2.ItemID = tdcid.ItemId2"

MySQL = MySQL & "         LEFT OUTER JOIN Groups     AS Grou5"
MySQL = MySQL & "              ON  Grou5.GroupID = tdcid.GroupIDBuiltin"
MySQL = MySQL & "         LEFT OUTER JOIN TblItems     AS Item5"
MySQL = MySQL & "              ON  Item5.ItemID = tdcid.BuiltinItemID"


MySQL = MySQL & "  Where (dbo.TblDefComItem.id = " & val(Me.TxtTransSerial.Text) & ")"




'StrWhere = StrWhere & " order by TblStuFingerprint.StudID"
  StrSQL = MySQL & StrWhere
  print_report2 StrSQL, 9
            

Exit Function

    If Dir(StrFileName) = "" Then
        'GetMsgs 139, vbExclamation
        Screen.MousePointer = vbDefault
        Exit Function
    End If

    Set RsData = New ADODB.Recordset
    RsData.Open MySQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

    If RsData.BOF Or RsData.EOF Then
        'GetMsgs 138, vbExclamation
        Msg = "?CE??I E?C?CE ?????"
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
    Dim oorderdate As Date
    Dim CBoBasedON As Integer
    Dim PONo As String

     
    
    If SystemOptions.UserInterface = ArabicInterface Then
        xReport.ParameterFields(1).AddCurrentValue cCompanyInfo.ArabCompanyName 'RPTCompany_Name_Arabic
        ' xReport.ParameterFields(2).AddCurrentValue RPTComment_Arabic
        StrReportTitle = "" '& StrAccountName
        'If Me.DTPickerAccFrom.value <> Empty Or Me.DTPickerAccFrom.value <> Null Then
        '    StrReportTitle = StrReportTitle + " EIC?E ?? " & Format(Me.DTPickerAccFrom.value, "yyyy/M/d") & ""
        'End If
        'If Me.DTPickerAccTo.value <> Empty Or Me.DTPickerAccTo.value <> Null Then
        '    StrReportTitle = StrReportTitle + " ??? " & Format(Me.DTPickerAccTo.value, "yyyy/M/d") & " "
        'End If
    Else
 
        xReport.ParameterFields(1).AddCurrentValue cCompanyInfo.ArabCompanyName ' RPTCompany_Name_Eng
        'xReport.ParameterFields(2).AddCurrentValue RPTComment_Eng
        xReport.ParameterFields(4).AddCurrentValue get_branch_name(val(val(dcBranch.BoundText)))
        StrReportTitle = ""
        'If Me.DTPickerAccFrom.value <> Empty Or Me.DTPickerAccFrom.value <> Null Then
        '    StrReportTitle = StrReportTitle + " From Date " & (Me.DTPickerAccFrom.value) & ""
        'End If
        'If Me.DTPickerAccTo.value <> Empty Or Me.DTPickerAccTo.value <> Null Then
        '    StrReportTitle = StrReportTitle + " To Date :  " & (Me.DTPickerAccTo.value) & ""
        'End If
    End If

    xReport.ParameterFields(3).AddCurrentValue user_name
      '  xReport.ParameterFields(4).AddCurrentValue WriteNo(Format(val(TxtAdvanceValue.text), "0.00"), 0, True, ".")
       ' xReport.ParameterFields(6).AddCurrentValue val(lbl(23).Caption)
       '  xReport.ParameterFields(7).AddCurrentValue DBIssueDate.value
'    xReport.ParameterFields(8).AddCurrentValue IIf(IsNumeric(fg.TextMatrix(Me.fg.FixedRows, fg.ColIndex("PartValue"))), val(fg.TextMatrix(Me.fg.FixedRows, fg.ColIndex("PartValue"))), 0)
' xReport.ParameterFields(9).AddCurrentValue val(lbl(22).Caption)
 ' xReport.ParameterFields(10).AddCurrentValue val(TxtDiscount.text)
  ' xReport.ParameterFields(12).AddCurrentValueval (lbTotalMente.Caption)
  
   
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




Private Sub SaveData(Optional ByVal IsSaveWithOutMsg As Boolean = False)
    Dim Msg As String
    Dim RowNum As Integer
    Dim RSTransDetails As ADODB.Recordset
    'Dim RsNotes As ADODB.Recordset
    Dim RsTemp  As New ADODB.Recordset
    Dim RsTest As New ADODB.Recordset
    Dim RsRepeat As ADODB.Recordset
    Dim StrSQL As String
    Dim StrSqlDel As String
    Dim mCostPrice  As Double
   Dim BeginTrans As Boolean
    Dim i As Long
   If Not IsSaveWithOutMsg Then
    If FG2.Rows > 1 Then
        If val(FG2.TextMatrix(1, FG2.ColIndex("ItemID"))) = 0 Then
            If SystemOptions.UserInterface = ArabicInterface Then
                Msg = "·« Ì„ﬂ‰ «·Õ›Ÿ »œÊ‰ «’‰«› „‰ ÃÂ...!!!"
            Else
                Msg = "Cannot save without items... !!!"
            End If
             MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
            Exit Sub
        End If
    Else
            If SystemOptions.UserInterface = ArabicInterface Then
                Msg = "·« Ì„ﬂ‰ «·Õ›Ÿ »œÊ‰ «’‰«›...!!!"
            Else
                Msg = "Can not duplicate Max code Please choose another one... !!!"
            End If
             MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
            Exit Sub
    End If
    
     If Trim(TxtMaxNo) <> "" Then
        StrSQL = "Select MaxNo from TblDefComItem Where MaxNo = N'" & Trim(TxtMaxNo) & "' And Id <> " & val(TxtTransSerial.Text)
        Dim rsDummyMax As New ADODB.Recordset
        rsDummyMax.Open StrSQL, Cn, adOpenStatic, adLockReadOnly
        If Not rsDummyMax.EOF Then
             If SystemOptions.UserInterface = ArabicInterface Then
                Msg = "·« Ì„ﬂ‰  ﬂ—«— ﬂÊœ «·„ﬂ” „‰ ›÷·ﬂ «Œ — „ﬂ” ¬Œ—...!!!"
            Else
                Msg = "Can not duplicate Max code Please choose another one... !!!"
            End If
            MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
            TxtMaxNo.SetFocus
            Exit Sub
        End If
        
    End If
    '    FixRowsLine
        If FG.Rows > 1 Then
            FG.Select 1, FG.ColIndex("LineID")
        End If
        FG.Sort = flexSortGenericAscending
        For RowNum = 1 To FG.Rows - 1

             If FG.RowHidden(RowNum) Or CBool(FG.ValueMatrix(RowNum, FG.ColIndex("IsDeleted"))) = True Then GoTo NextRow33

            If FG.TextMatrix(RowNum, FG.ColIndex("ItemID")) <> "" Then
                
                
                 If SystemOptions.SysAllowStockNegative = False And Selct(1).value = vbChecked Then
                        
                            
                    StrSQL = "Select * From TblItems where ItemID=" & val(FG.TextMatrix(RowNum, FG.ColIndex("ItemID")))
                    Set RsTemp = New ADODB.Recordset
                    RsTemp.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

                    If Not RsTemp.EOF Then
                        
                        If DCboStore2Name.BoundText <> "" Then
                            Set RsTest = New ADODB.Recordset
                            Set RsTest = GetItemQuantityStock(val(FG.TextMatrix(RowNum, FG.ColIndex("ItemID"))), val(Me.DCboStore2Name.BoundText), , , True)

                            If RsTest.EOF Or RsTemp.BOF Then
                                If SystemOptions.UserInterface = ArabicInterface Then
                                    Msg = "«·ﬁÿ⁄… –«  «·”Ì—Ì«· : "
                                    Msg = Msg + " «·’‰› : " & Trim(FG.Cell(flexcpTextDisplay, RowNum, FG.ColIndex("ItemName"))) & CHR(13) & "Ê«·„ÊÃÊœ ›Ï «·”ÿ— —ﬁ„  " & RowNum
                                    Msg = Msg + " €Ì— „ÊÃÊœ… ›Ì «·„Œ“‰ «·„Õœœ" & CHR(13)
                                    Msg = Msg + "Ê»«· «·Ï ·„ Ì „ «‰‘«¡ «–‰ «·’—›"
                                Else
                                    Msg = "The cathode "
                                    Msg = Msg + " Item : " & Trim(FG.Cell(flexcpTextDisplay, RowNum, FG.ColIndex("ItemName"))) & CHR(13) & "Located in the line " & RowNum
                                    Msg = Msg + " Not in the specified store" & CHR(13)
                                    Msg = Msg + "Consequently, no exchange permit was created"
                                End If
                                MsgBox Msg
                                Exit Sub
                            End If
                        End If
                    End If
                End If
            End If
NextRow33:
            
        Next
   
   '

          Dim RsUnitData As ADODB.Recordset
            Dim LngCurItemID As Long
            Dim LngUnitID As Long
            Dim DblQty As Double
    'On Error GoTo ErrTrap
   ' Screen.MousePointer = vbArrowHourglass
        If SystemOptions.IsMultiItemsInCompItem Then

            If CboPayMentType.ListIndex <> 2 And (val(DBCboClientName.BoundText) = 1 Or val(DBCboClientName.BoundText) = 2) And Me.TxtModFlg.Text <> "R" Then
              '  CboPayMentType.locked = True
                        CboPayMentType.ListIndex = 0
                        If SystemOptions.UserInterface = EnglishInterface Then
                            Msg = "You can not select a cash customer with a forward payment"
                        Else
                            Msg = "·« Ì„ﬂ‰ «Œ Ì«— ⁄„Ì· ‰ﬁœÌ „⁄ «·”œ«œ «·¬Ã·"
                        End If
             End If
             
                If Trim(dcBranch.BoundText) = "" Then
                    If SystemOptions.UserInterface = EnglishInterface Then
                        Msg = "Specify Departement"
                    Else
                        Msg = "Õœœ «·›—⁄ «Ê·« "
                    End If
                  cmd(2).Enabled = True
                    MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
                    dcBranch.SetFocus
                    SendKeys "{F4}"
                    Screen.MousePointer = vbDefault
                    Exit Sub
                End If
    
         
    
                If CboPayMentType.ListIndex = 0 Or CboPayMentType.ListIndex = 2 Then
                    If Not SystemOptions.IsMultiItemsInCompItem Then
                       If val(txtTotalWithVat.Text) < 0 Then
                           If SystemOptions.UserInterface = EnglishInterface Then
                               Msg = "Enter Correct Payed Value"
                           Else
                               Msg = "  «·ﬁÌ„… €Ì— ’ÕÌÕ…"
                           End If
                    
                           MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
                          cmd(2).Enabled = True
                           Exit Sub
                       End If
                    End If
                End If
    
                If (CboPayMentType.ListIndex = 1 Or CboPayMentType.ListIndex = 2) Then
                    'XPTxtValue(1).Text = LblTotal.Caption
                End If
                
                
                If val(Me.TxtVAt2.Text) > 0 Then
                    If GetValueAddedAccount(XPDtbBill.value, , , 1, 21) = False Then
                        If SystemOptions.UserInterface = ArabicInterface Then
                            MsgBox "·„ Ì „  ÕœÌœ Õ”«» «·ﬁÌ„… «·„÷«›…"
                        Else
                            MsgBox "Value added account not specified"
                        End If
                        cmd(2).Enabled = True
                        Exit Sub
                    End If
                End If
                
        End If
        Dim RsNotesGeneral As ADODB.Recordset
            Set RsNotesGeneral = New ADODB.Recordset
            'RsNotesGeneral.Open "[Notes]", Cn, adOpenStatic, adLockOptimistic, adCmdTable
            StrSQL = "SELECT     dbo.Notes.* from dbo.Notes Where (NoteID = -1)"
   RsNotesGeneral.Open StrSQL, Cn, adOpenKeyset, adLockOptimistic, adCmdText
 
            '    my_branch = Me.Dcbranch.BoundText
      
            '    If Me.TxtModFlg.text = "E" Then
             
            '     TxtInvID
            '     End If
        
   
End If

    If Me.TxtModFlg.Text <> "R" Then
        
'    If val(Me.DcboItemID1.BoundText) = 0 Then
'        Msg = "ÌÃ»  ÕœÌœ «Ú”„ «·’‰›  «·„‰ Ã...!!!"
'        MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
'        Me.DcboItemID1.SetFocus
'        Exit Sub
'    End If
If Selct(1).value = vbChecked Then

    If val(Me.DCboStore2Name.BoundText) = 0 Then
        
        If SystemOptions.UserInterface = ArabicInterface Then
            Msg = "ÌÃ»  ÕœÌœ «Ú”„ „Œ“‰  «·„Ê«œ «·Œ«„...!!!"
        Else
            Msg = "You must specify the name of the raw material store ... !!!"
        End If
        MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
        Me.DCboStore2Name.SetFocus
        Exit Sub
    End If
    
End If

If Not IsSaveWithOutMsg Then
            If SystemOptions.IsMultiItemsInCompItem Then
                    If CboPayMentType.ListIndex = 0 Then
                            If Me.DcboBox.BoundText = "" Then
                                         
                                If SystemOptions.UserInterface = ArabicInterface Then
                                    MsgBox "ÌÃ»  ÕœÌœ «”„ «·Œ“‰…...!!!", vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
                                Else
                                    MsgBox "Must Specify Box...!!!", vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
                                        
                                End If
                                cmd(2).Enabled = True
                                DcboBox.SetFocus
                                SendKeys "{F4}"
                                        
                                Screen.MousePointer = vbDefault
                              '  Cmd(2).Enabled = True
                                Exit Sub
                            End If
                    End If
                   If Trim(DcboEmp.BoundText) = "" Then
                        If SystemOptions.UserInterface = ArabicInterface Then
                            Msg = "ÌÃ»  ÕœÌœ «”„ «·»«∆⁄/«·„‰œÊ»..!!!"
                        Else
                            Msg = "Must Specify SalesPerson/Saller..!!!"
                        End If
                cmd(2).Enabled = True
                        MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
                        DcboEmp.SetFocus
                        SendKeys "{F4}"
                        Screen.MousePointer = vbDefault
                       ' Cmd(2).Enabled = True
                        Exit Sub
                    End If
                  If CboPayMentType.ListIndex = -1 Then
                        If SystemOptions.UserInterface = ArabicInterface Then
                            Msg = "ÌÃ»  ÕœÌœ ÿ—Ìﬁ… «·œ›⁄"
                        Else
                            Msg = "Specify Payment Method"
                        End If
                
                        MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
                      '  CboPayMentType.SetFocus
                        SendKeys "{F4}"
                        Screen.MousePointer = vbDefault
                        Exit Sub
                    End If
            End If
            If Selct(2).value = vbChecked Then
            
                If val(Me.DCboStore3Name.BoundText) = 0 Then
                    If SystemOptions.UserInterface = ArabicInterface Then
                        Msg = "ÌÃ»  ÕœÌœ «Ú”„ „Œ“‰     «·„‰ Ã «· «„ ...!!!"
                    Else
                        Msg = "You must specify the name of the complete product store ... !!!"
                    End If
                    MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
                    Me.DCboStore3Name.SetFocus
                    Exit Sub
                End If
                
            End If
            
            
            If val(txtQty1.Text) = 0 Then
                If SystemOptions.UserInterface = ArabicInterface Then
                    Msg = "ÌÃ»  ÕœÌœ ﬂ„Ì…  «·’‰›  «·„‰ Ã...!!!"
                Else
                    Msg = "Must specify the quantity of product category ... !!!"
                End If
                    MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
                    Me.txtQty1.SetFocus
                    Exit Sub
                    
            End If
            
            
            '    If val(Me.DcbUnit.BoundText) = 0 Then
            '        Msg = "ÌÃ»  ÕœÌœ ÊÕœ…  «·’‰›  «·„‰ Ã...!!!"
            '        MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
            '        Me.DcbUnit.SetFocus
            '        Exit Sub
            '    End If
                Dim TxtCashCustomerName As TextBox
                If SystemOptions.IsMultiItemsInCompItem Then
                    If Me.CboPayMentType.ListIndex = 1 Then
                            If val(Me.txtTotalWithVat.Text) > 0 Then
                                If CheckCusCredit(val(Me.DBCboClientName.BoundText), val(Me.txtTotalWithVat.Text), 0) = False Then
                                    Screen.MousePointer = vbDefault
                                    Exit Sub
                                End If
                            End If
                        End If
                        
                    If val(CboPayMentType.ListIndex) <> 1 Then
                        If SystemOptions.CashCustomerNameMustenter = True And val(DBCboClientName.BoundText) = 2 Then
                            If TxtCashCustomerName = "" Then
                                If SystemOptions.UserInterface = ArabicInterface Then
                                    MsgBox "Ì—ÃÏ «œŒ«· «”„ «·⁄„Ì·"
                                Else
                                    MsgBox "Please Enter Customer"
                                End If
                            'TxtCashCustomerName.SetFocus
                                cmd(2).Enabled = True
                                Exit Sub
                                cmd(2).Enabled = True
                            End If
                
                            cmd(2).Enabled = True
                            Exit Sub
                        End If
                    
                    End If
                'End If
                   
                   If val(Me.DBCboClientName.BoundText) <> 1 Or val(Me.DBCboClientName.BoundText) <> 2 Then
                        If Me.CboPayMentType.ListIndex = 1 Then
                            If val(Me.txtTotalWithVat.Text) > 0 Then
                            
                            Dim MsgRe2 As String
                            
                                         If CheckCusCredit(val(Me.DBCboClientName.BoundText), val(Me.txtTotalWithVat), 0, 0) = False Then
                                                   Screen.MousePointer = vbDefault
                                                   cmd(2).Enabled = True
                                              Exit Sub
                                            End If
                                   
                                            
                                            
                                            
                            End If
                        End If
                    End If
            End If
            
            
            If Selct(2).value = vbChecked Then
            
                If val(Me.DCboStore3Name.BoundText) = 0 Then
                    If SystemOptions.UserInterface = ArabicInterface Then
                        Msg = "ÌÃ»  ÕœÌœ «Ú”„ „Œ“‰     «·„‰ Ã «· «„ ...!!!"
                    Else
                        Msg = "You must specify the name of the complete product store ... !!!"
                    End If
                    MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
                    Me.DCboStore3Name.SetFocus
                    Exit Sub
                End If
                
            End If
            
            
            If val(txtQty1.Text) = 0 Then
                If SystemOptions.UserInterface = ArabicInterface Then
                     Msg = "ÌÃ»  ÕœÌœ ﬂ„Ì…  «·’‰›  «·„‰ Ã...!!!"
                Else
                     Msg = "Must specify the quantity of product category ... !!!"
                End If
                    MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
                    Me.txtQty1.SetFocus
                    Exit Sub
                    
            End If
            
                If Not SystemOptions.IsMultiItemsInCompItem Then
                    If val(Me.DcbUnit.BoundText) = 0 Then
                        If SystemOptions.UserInterface = ArabicInterface Then
                            Msg = "ÌÃ»  ÕœÌœ ÊÕœ…  «·’‰›  «·„‰ Ã...!!!"
                        Else
                            Msg = "The product unit must be specified"
                        End If
                        MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
                        Me.DcbUnit.SetFocus
                        Exit Sub
                    End If
                End If
                If SystemOptions.IsMultiItemsInCompItem Then
                    If Me.CboPayMentType.ListIndex = 1 Then
                            If val(Me.txtTotalWithVat.Text) > 0 Then
                                If CheckCusCredit(val(Me.DBCboClientName.BoundText), val(Me.txtTotalWithVat.Text), 0) = False Then
                                    Screen.MousePointer = vbDefault
                                    Exit Sub
                                End If
                            End If
                        End If
                        
                    If val(CboPayMentType.ListIndex) <> 1 Then
                        If SystemOptions.CashCustomerNameMustenter = True And val(DBCboClientName.BoundText) = 2 Then
                            If TxtCashCustomerName.Text = "" Then
                                If SystemOptions.UserInterface = ArabicInterface Then
                                    MsgBox "Ì—ÃÏ «œŒ«· «”„ «·⁄„Ì·"
                                Else
                                    MsgBox "Please Enter Customer"
                                End If
                            'TxtCashCustomerName.SetFocus
                                cmd(2).Enabled = True
                                Exit Sub
                                cmd(2).Enabled = True
                            End If
                
                            cmd(2).Enabled = True
                            Exit Sub
                        End If
                    
                    End If
                'End If
                   
                   If val(Me.DBCboClientName.BoundText) <> 1 Or val(Me.DBCboClientName.BoundText) <> 2 Then
                        If Me.CboPayMentType.ListIndex = 1 Then
                            If val(Me.txtTotalWithVat.Text) > 0 Then
                            
                           ' Dim MsgRe2 As String
                            
                                         If CheckCusCredit(val(Me.DBCboClientName.BoundText), val(Me.txtTotalWithVat), 0, 0) = False Then
                                                   Screen.MousePointer = vbDefault
                                                   cmd(2).Enabled = True
                                              Exit Sub
                                            End If
                                   
                                            
                                            
                                            
                            End If
                        End If
                    End If
                End If
            
            
                        If TXT_order_no = "" Then
                        For i = 1 To FG2.Rows - 1
                        
                            FG2.TextMatrix(i, FG2.ColIndex("TransactionID4")) = ""
                            FG2.TextMatrix(i, FG2.ColIndex("NoteSerial14")) = ""
                 
                     
                        Next
            End If
    End If
        Cn.BeginTrans
        BeginTrans = True

        If Not IsSaveWithOutMsg Then
            
            DeleteTransactiomsVoucher2 val(TXTTransactionID1.Text)
            DeleteTransactiomsVoucher2 val(TXTTransactionID2.Text)
            DeleteTransactiomsVoucher2 val(TXTTransactionID3.Text)
            DeleteTransactiomsVoucher2 val(TXTTransactionID4.Text)
            DeleteTransactiomsVoucher2 val(TXTTransactionID5.Text)
    
            
      
            
            Cn.Execute "Delete from TransactionValueAdded where Transaction_ID=" & val(Me.TXTTransactionID3.Text) & ""
    
       
    
            StrSQL = "delete From Notes where noteid=" & val(TXTNoteID.Text)
            Cn.Execute StrSQL, , adExecuteNoRecords
        
                
            If TXT_order_no = "" Then
                      StrSQL = "Delete Transactions Where Transaction_ID In (Select TransactionID4 From TblDefComItemData Where IDDefCIT = " & val(TxtTransSerial) & ")"
                            Cn.Execute StrSQL, , adExecuteNoRecords
                            StrSQL = "Delete Transaction_Details Where Transaction_ID In (Select TransactionID4 From TblDefComItemData Where IDDefCIT = " & val(TxtTransSerial) & ")"
                            Cn.Execute StrSQL, , adExecuteNoRecords
                           TXTTransactionID4 = ""
                           TxtNoteSerial14 = ""
                           
                            StrSQL = "UPDATE TblDefComItem SET  TransactionID4=" & val(TXTTransactionID4) & ",  NoteSerial14='" & TxtNoteSerial14 & "' WHERE ID  =" & val(TxtTransSerial)
                        Cn.Execute StrSQL
                           StrSQL = "UPDATE TblDefComItemData SET    NoteSerial14='' ,TransactionID4 = 0 WHERE IDDefCIT  =" & val(TxtTransSerial)
                    Cn.Execute StrSQL
                        End If
                        
                
                
                 
                Dim s As String
                s = "Delete TblProductLineDistribution Where IDDefCIT = " & val(TxtTransSerial)
                Cn.Execute s
        End If
        If Me.TxtModFlg.Text = "N" Then
            rs.AddNew
            TxtTransSerial.Text = CStr(new_id("TblDefComItem", "ID", "", True))
            
                
        End If
         If TxtMaxName = "" Then
            TxtMaxName = DBCboClientName.Text & "-" & TxtTransSerial
        End If
  If TxtMaxNo.Text = "" Then
  TxtMaxNo.Text = TxtTransSerial.Text
  End If
'  rs("ID").value = val(TxtTransSerial.text)
          rs("RecordDate").value = XPDtbBill.value
          rs("BranchID").value = val(Me.dcBranch.BoundText)
        rs("StoreID").value = val(Me.DCboStoreName.BoundText)
    rs("RecDate").value = XPDtRecDate.value
    '************
      rs("StoreID2").value = IIf(Me.DCboStore2Name.BoundText = "", Null, (Me.DCboStore2Name.BoundText))
      rs("StoreID3").value = IIf(Me.DCboStore3Name.BoundText = "", Null, (Me.DCboStore3Name.BoundText))
      rs("ItemNameID").value = IIf(Me.DcboItemID1.BoundText = "", Null, (Me.DcboItemID1.BoundText))
      rs("UnitID").value = IIf(Me.DcbUnit.BoundText = "", Null, (Me.DcbUnit.BoundText))
      rs("CusID").value = IIf(Me.DBCboClientName.BoundText = "", Null, (Me.DBCboClientName.BoundText))
      rs("MaxNo").value = IIf(TxtMaxNo.Text = "", "", (TxtMaxNo.Text))
      rs("MaxNo2").value = IIf(TxtMaxNo2.Text = "", "", (TxtMaxNo2.Text))
      rs("MaxName").value = IIf(TxtMaxName.Text = "", "", (TxtMaxName.Text))
      rs("ItemCode").value = IIf(TxtAttachedItemCode.Text = "", "", (TxtAttachedItemCode.Text))
      rs("UserID").value = user_id
    rs("order_no").value = IIf(TXT_order_no.Text = "", "", (TXT_order_no.Text))
    rs("OrderID").value = IIf(txtOrderID.Text = "", 0, (txtOrderID.Text))
              
      rs("Qty1").value = val(txtQty1.Text)
      rs("hight").value = val(txthight.Text)
      rs("widtj").value = val(txtwidtj.Text)
      rs("Length").value = val(txtLength.Text)
        rs("Period").value = val(txtPeriod.Text)
    rs("RecTime").value = FormatDateTime(Me.txtRecTime.value, vbShortTime)
      rs("TransactionID1").value = val(TXTTransactionID1.Text)
            rs("TransactionID2").value = val(TXTTransactionID2.Text)
             
            rs("NoteSerial11").value = (TxtNoteSerial11.Text)
            rs("NoteSerial12").value = (TxtNoteSerial12.Text)
        If CboPayMentType.ListIndex = -1 Then
            rs("PaymentType").value = 0
        Else
            rs("PaymentType").value = val(CboPayMentType.ListIndex)
        End If
       

        rs("Vat2").value = val(TxtVAt2.Text)
        
        rs("TotalWithVat").value = val(txtTotalWithVat.Text)
        
        rs("Price").value = val(TxtPrice.Text)
        rs("TotalAdd").value = val(txtTotalAdd.Text)
        rs("TotalDisc").value = val(txtTotalDisc.Text)
        
        
        
        rs("Net").value = val(txtNet.Text)
        If XPCboGroup.BoundText = "" Then
            rs("GroupID").value = Null
        Else
            rs("GroupID").value = val(XPCboGroup.BoundText)
        End If
        
       rs("BuiltinItemID").value = IIf(DcboBuiltinItemID.BoundText = "", Null, DcboBuiltinItemID.BoundText)
       rs("GroupIDBuiltin").value = IIf(XPCboGroupBuiltin.BoundText = "", Null, XPCboGroupBuiltin.BoundText)
        
        rs("Emp_ID").value = IIf(DcboEmp.BoundText = "", Null, DcboEmp.BoundText)
        
    


    If CboPayMentType.ListIndex = 0 Or CboPayMentType.ListIndex = 2 Then
        rs("BoxID").value = IIf(DcboBox.BoundText = "", Null, val(DcboBox.BoundText))
    Else
        rs("BoxID").value = Null
      
    End If
If Me.Selct(0).value = vbChecked Then
rs("Allocated").value = 1
Else
rs("Allocated").value = 0
End If
If Me.Selct(1).value = vbChecked Then
rs("AlloPay").value = 1
Else
rs("AlloPay").value = 0
End If
If Me.Selct(2).value = vbChecked Then
rs("AlloRecep").value = 1
Else
rs("AlloRecep").value = 0
End If
       rs.update
        
If Me.TxtModFlg.Text = "N" Then
          
            TxtTransSerial.Text = IIf(IsNull(rs("id").value), 0, rs("id").value)
            
            
            
                
        End If


        If Me.TxtModFlg.Text = "E" Then
            StrSqlDel = "delete From TblDefComItemDet where IDDefCIT=" & val(TxtTransSerial.Text) & ""
            Cn.Execute StrSqlDel, , adExecuteNoRecords
            StrSqlDel = "delete From TblDefComItemData where IDDefCIT=" & val(TxtTransSerial.Text) & ""
            Cn.Execute StrSqlDel, , adExecuteNoRecords
            
            
        End If


   StrSQL = "SELECT * FROM TblDefComItemDet where 1=-1 "
    Dim mLineNo As Long
    CostTOTAL = 0
    Set RSTransDetails = New ADODB.Recordset
    RSTransDetails.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText
        For RowNum = 1 To FG.Rows - 1
           mCostPrice = 0
            If FG.TextMatrix(RowNum, FG.ColIndex("itemname")) <> "" Then
                mLineNo = val(FG.TextMatrix(RowNum, FG.ColIndex("LineID")))
                RSTransDetails.AddNew
                RSTransDetails("IDDefCIT").value = val(TxtTransSerial.Text)
                RSTransDetails("TepQty").value = IIf((FG.TextMatrix(RowNum, FG.ColIndex("TepQty")) = ""), Null, val(FG.TextMatrix(RowNum, FG.ColIndex("TepQty"))))
                RSTransDetails("FlgX").value = IIf((FG.TextMatrix(RowNum, FG.ColIndex("FlgX")) = ""), Null, val(FG.TextMatrix(RowNum, FG.ColIndex("FlgX"))))
                RSTransDetails("ItemID").value = IIf((FG.TextMatrix(RowNum, FG.ColIndex("ItemID")) = ""), Null, val(FG.TextMatrix(RowNum, FG.ColIndex("ItemID"))))
                RSTransDetails("ItemID2").value = IIf((FG.TextMatrix(RowNum, FG.ColIndex("ItemID2")) = ""), Null, val(FG.TextMatrix(RowNum, FG.ColIndex("ItemID2"))))
                RSTransDetails("ItemCode2").value = IIf((FG.TextMatrix(RowNum, FG.ColIndex("ItemCode2")) = ""), Null, val(FG.TextMatrix(RowNum, FG.ColIndex("ItemCode2"))))
                RSTransDetails("UnitID").value = IIf((FG.TextMatrix(RowNum, FG.ColIndex("UnitID")) = ""), Null, val(FG.TextMatrix(RowNum, FG.ColIndex("UnitID"))))
                RSTransDetails("LineID").value = IIf((FG.TextMatrix(RowNum, FG.ColIndex("LineID")) = ""), Null, val(FG.TextMatrix(RowNum, FG.ColIndex("LineID"))))
                RSTransDetails("SpecID1").value = IIf((FG.TextMatrix(RowNum, FG.ColIndex("SpecID1")) = ""), Null, val(FG.TextMatrix(RowNum, FG.ColIndex("SpecID1"))))
                RSTransDetails("SpecID2").value = IIf((FG.TextMatrix(RowNum, FG.ColIndex("SpecID2")) = ""), Null, val(FG.TextMatrix(RowNum, FG.ColIndex("SpecID2"))))
                RSTransDetails("SpecID3").value = IIf((FG.TextMatrix(RowNum, FG.ColIndex("SpecID3")) = ""), Null, val(FG.TextMatrix(RowNum, FG.ColIndex("SpecID3"))))
                RSTransDetails("TableID").value = IIf((FG.TextMatrix(RowNum, FG.ColIndex("TableID")) = ""), Null, val(FG.TextMatrix(RowNum, FG.ColIndex("TableID"))))

                
'                RSTransDetails("ForUnit").value = IIf((fg.TextMatrix(RowNum, fg.ColIndex("ForUnit")) = ""), 0, val(fg.TextMatrix(RowNum, fg.ColIndex("ForUnit"))))
'                RSTransDetails("MethodCalc").value = IIf((fg.TextMatrix(RowNum, fg.ColIndex("MethodCalc")) = ""), 0, val(fg.TextMatrix(RowNum, fg.ColIndex("MethodCalc"))))
'                RSTransDetails("PartItemQty").value = IIf((fg.TextMatrix(RowNum, fg.ColIndex("PartItemQty")) = ""), 0, val(fg.TextMatrix(RowNum, fg.ColIndex("PartItemQty"))))
            
                RSTransDetails("SpecID4").value = IIf((FG.TextMatrix(RowNum, FG.ColIndex("SpecID4")) = ""), Null, val(FG.TextMatrix(RowNum, FG.ColIndex("SpecID4"))))
                RSTransDetails("Amout1").value = IIf((FG.TextMatrix(RowNum, FG.ColIndex("Amout1")) = ""), 0, val(FG.TextMatrix(RowNum, FG.ColIndex("Amout1"))))
                RSTransDetails("Amout2").value = IIf((FG.TextMatrix(RowNum, FG.ColIndex("Amout2")) = ""), 0, val(FG.TextMatrix(RowNum, FG.ColIndex("Amout2"))))
                RSTransDetails("Amout3").value = IIf((FG.TextMatrix(RowNum, FG.ColIndex("Amout3")) = ""), 0, val(FG.TextMatrix(RowNum, FG.ColIndex("Amout3"))))
                RSTransDetails("increase").value = IIf((FG.TextMatrix(RowNum, FG.ColIndex("Increase")) = ""), 0, val(FG.TextMatrix(RowNum, FG.ColIndex("Increase"))))
                RSTransDetails("lowering").value = IIf((FG.TextMatrix(RowNum, FG.ColIndex("lowering")) = ""), 0, val(FG.TextMatrix(RowNum, FG.ColIndex("lowering"))))
                
                RSTransDetails("widtj").value = val(FG.TextMatrix(RowNum, FG.ColIndex("widtj")))
                RSTransDetails("hight").value = val(FG.TextMatrix(RowNum, FG.ColIndex("hight")))
                RSTransDetails("Length").value = val(FG.TextMatrix(RowNum, FG.ColIndex("Length")))
                RSTransDetails("thickness").value = val(FG.TextMatrix(RowNum, FG.ColIndex("thickness")))
                RSTransDetails("DO").value = val(FG.TextMatrix(RowNum, FG.ColIndex("DO")))
                RSTransDetails("DI").value = val(FG.TextMatrix(RowNum, FG.ColIndex("DI")))
                RSTransDetails("Diameter").value = val(FG.TextMatrix(RowNum, FG.ColIndex("Diameter")))
                RSTransDetails("IsRow").value = FG.ValueMatrix(RowNum, FG.ColIndex("IsRow"))
            
                        
   


            
                RSTransDetails("Amout4").value = IIf(FG.Cell(flexcpData, RowNum, FG.ColIndex("Amout4")) = "", 0, (FG.Cell(flexcpData, RowNum, FG.ColIndex("Amout4"))))
                RSTransDetails("Qty").value = IIf((FG.TextMatrix(RowNum, FG.ColIndex("Qty")) = ""), Null, val(FG.TextMatrix(RowNum, FG.ColIndex("Qty"))))
                
     
         
             LngCurItemID = val(FG.TextMatrix(RowNum, FG.ColIndex("ItemID")))
             LngUnitID = val(FG.TextMatrix(RowNum, FG.ColIndex("UnitID"))) 'val(Fg.Cell(flexcpData, RowNum, Fg.ColIndex("UnitID")))
                If LngUnitID = 0 Then
                GetDefaultItemUnit val(LngCurItemID), LngUnitID
            End If
            
             DblQty = val(FG.TextMatrix(RowNum, FG.ColIndex("Qty")))
           
              
            LngCurItemID = val(FG.TextMatrix(RowNum, FG.ColIndex("ItemID")))
             
             
                  'If LngCurItemID = 810 Or val(fg2.TextMatrix(mLineNo, fg2.ColIndex("ItemID"))) = 643 Then
                 'If LngCurItemID = 810 And val(FG.TextMatrix(RowNum, FG.ColIndex("ItemID2"))) = 634 Then
              
                                 
             Dim OldQty As Double, OldCost As Double, NewQty As Double, NewCost As Double
             Dim mIsFromMix As Boolean
          
             
             'costPrice = 20
               If val(TXT_order_no) <> 0 And SystemOptions.CostByProduction Then
            For i = 1 To FG2.Rows - 1
            
                    s = "SELECT T2.* "
                    s = s & " from  Transactions AS t "
                    s = s & " Inner Join Transaction_Details T2 On T2.Transaction_ID = t.Transaction_ID"
                    If val(FG2.TextMatrix(i, FG2.ColIndex("TransactionID4"))) <> 0 Then
                        s = s & " WHERE t.Transaction_Type = 26 and t.Transaction_ID =  " & val(FG2.TextMatrix(i, FG2.ColIndex("TransactionID4")))
                    ElseIf val(FG2.TextMatrix(i, FG2.ColIndex("TransactionID5"))) <> 0 Then
                        s = s & " WHERE t.Transaction_Type = 10 and t.Transaction_ID =  " & val(FG2.TextMatrix(i, FG2.ColIndex("TransactionID5")))
                    End If
                    s = s & " and  T2.Item_ID = " & val(FG2.TextMatrix(i, FG2.ColIndex("ItemID")))
                    s = s & " and T2.UnitId= " & val(FG2.TextMatrix(i, FG2.ColIndex("UnitID")))
                    Set rsDummy = New ADODB.Recordset
    
    '
                    rsDummy.Open s, Cn, adOpenStatic, adLockReadOnly, adCmdText
                    If rsDummy.EOF Then
                        mCostPrice = 0
                    Else
                        mCostPrice = val(rsDummy!ShowPrice & "")
                    End If
    
   
    
                If mCostPrice <> 0 Then
                    FG2.TextMatrix(i, FG2.ColIndex("cost")) = mCostPrice
                
                End If
            Next
        End If
        If val(FG.TextMatrix(RowNum, FG.ColIndex("LineId"))) = 12 Then
            FG.TextMatrix(RowNum, FG.ColIndex("LineId")) = 12
        End If
            If mCostPrice = 0 Then
                mCostPrice = GetCostFromMix2(RowNum)
                 
                 If mCostPrice = 0 Then
                    mCostPrice = ModItemCostPrice.GetCostItemPrice(CLng(LngCurItemID), 0, , , SystemOptions.SysMainStockCostMethod, , , XPDtbBill, , LngUnitID, val(Me.DCboStore2Name.BoundText))
                    mIsFromMix = False
                Else
                    mIsFromMix = True
                 '   getItemCostData XPDtbBill.value, CLng(LngCurItemID), val(DCboStore2Name.BoundText), val(Me.TXTTransactionID2.Text), OldQty, OldCost, NewQty, NewCost
                 End If
                End If
    '             If costPrice = 0 Then
'                costPrice = ModItemCostPrice.GetCostItemPrice(CLng(LngCurItemID), , , , SystemOptions.SysMainStockCostMethod, , , Date, , LngUnitID)
'             End If
'
             
             If FG.RowHidden(RowNum) Or CBool(FG.ValueMatrix(RowNum, FG.ColIndex("IsDeleted"))) = True Then
                CostTOTAL = CostTOTAL
             Else
                'CostTOTAL = CostTOTAL + (costPrice) '* IIf(mIsFromMix, 1, DblQty)
                CostTOTAL = CostTOTAL + (mCostPrice * DblQty)
             '   CostTOTAL = CostTOTAL + (mCostPrice)
            End If
  
                If RowNum = 9 Then
                    RowNum = 9
                End If
                            
            'mCostPrice = ModItemCostPrice.GetCostItemPrice(CLng(LngCurItemID), 0, "", , SystemOptions.SysMainStockCostMethod, DblQty, , XPDtbBill, , LngUnitID)
            
                RSTransDetails("cost").value = mCostPrice
                
                FG.TextMatrix(RowNum, FG.ColIndex("cost")) = mCostPrice
                
               
                'IIf((fg.TextMatrix(RowNum, fg.ColIndex("cost")) = ""), 0, val(fg.TextMatrix(RowNum, fg.ColIndex("cost"))))
                RSTransDetails("IsAdd").value = FG.ValueMatrix(RowNum, FG.ColIndex("IsAdd"))
                RSTransDetails("IsDeleted").value = FG.ValueMatrix(RowNum, FG.ColIndex("IsDeleted"))
                RSTransDetails("Price").value = IIf((FG.TextMatrix(RowNum, FG.ColIndex("Price")) = ""), 0, val(FG.TextMatrix(RowNum, FG.ColIndex("Price"))))
                RSTransDetails("OldPrice").value = IIf((FG.TextMatrix(RowNum, FG.ColIndex("OldPrice")) = ""), 0, val(FG.TextMatrix(RowNum, FG.ColIndex("OldPrice"))))
                RSTransDetails("Total").value = IIf((FG.TextMatrix(RowNum, FG.ColIndex("Total")) = ""), 0, val(FG.TextMatrix(RowNum, FG.ColIndex("Total"))))
                
                If mLineNo = 2 Then
                        mLineNo = mLineNo
                End If
              '  mLineNo = val(FG.TextMatrix(RowNum, FG.ColIndex("LineID")))
                If SystemOptions.IsMultiItemsInCompItem Then
                    If RowNum + 1 < FG.Rows Then
                        If mLineNo <> val(FG.TextMatrix(RowNum + 1, FG.ColIndex("LineID"))) Then
                            If val(FG2.TextMatrix(mLineNo, FG2.ColIndex("Qty"))) <> 0 Then
                                FG2.TextMatrix(mLineNo, FG2.ColIndex("cost")) = CostTOTAL / val(FG2.TextMatrix(mLineNo, FG2.ColIndex("Qty")))
                            End If
                            CostTOTAL = 0
                        End If
                    Else
                        If mLineNo < FG2.Rows Then
                            If val(FG2.TextMatrix(mLineNo, FG2.ColIndex("Qty"))) <> 0 Then
                                FG2.TextMatrix(mLineNo, FG2.ColIndex("cost")) = CostTOTAL / val(FG2.TextMatrix(mLineNo, FG2.ColIndex("Qty")))
                            End If
                        End If
                    End If
                'Else
                    
                End If
                RSTransDetails.update
            End If

        Next RowNum
        If Not SystemOptions.IsMultiItemsInCompItem Then
            If FG2.Rows = 1 Then
                FillGridItemType val(DcboItemID1.BoundText), DcboItemID1.Text, Trim$(TxtAttachedItemCode.Text), 1, val(DcbUnit.BoundText), DcbUnit.Text, val(txtQty1), val(TxtPrice), val(XPCboGroup.BoundText), XPCboGroup.Text
                CalcGrid2
            End If
            FG2.TextMatrix(1, FG2.ColIndex("cost")) = CostTOTAL
        End If
        
        
            If val(TXT_order_no) <> 0 And SystemOptions.CostByProduction Then
                 For i = 1 To FG2.Rows - 1
                 
                         s = "SELECT T2.* "
                         s = s & " from  Transactions AS t "
                         s = s & " Inner Join Transaction_Details T2 On T2.Transaction_ID = t.Transaction_ID"
                         
                        If val(FG2.TextMatrix(i, FG2.ColIndex("TransactionID4"))) <> 0 Then
                            s = s & " WHERE t.Transaction_Type = 26 and t.Transaction_ID =  " & val(FG2.TextMatrix(i, FG2.ColIndex("TransactionID4")))
                        ElseIf val(FG2.TextMatrix(i, FG2.ColIndex("TransactionID5"))) <> 0 Then
                            s = s & " WHERE t.Transaction_Type = 10 and t.Transaction_ID =  " & val(FG2.TextMatrix(i, FG2.ColIndex("TransactionID5")))
                        End If
                    
                   
                         s = s & " and  T2.Item_ID = " & val(FG2.TextMatrix(i, FG2.ColIndex("ItemID")))
                         s = s & " and T2.UnitId= " & val(FG2.TextMatrix(i, FG2.ColIndex("UnitID")))
                         Set rsDummy = New ADODB.Recordset
         
         '
                         rsDummy.Open s, Cn, adOpenStatic, adLockReadOnly, adCmdText
                         If rsDummy.EOF Then
                             mCostPrice = 0
                         Else
                             mCostPrice = val(rsDummy!ShowPrice & "")
                         End If
         
                    If mCostPrice = 0 Then
                            mCostPrice = GetCostFromMix2(RowNum)
                 
                       If mCostPrice = 0 Then
                           mCostPrice = ModItemCostPrice.GetCostItemPrice(CLng(LngCurItemID), 0, , , SystemOptions.SysMainStockCostMethod, , , XPDtbBill, val(Me.TXTTransactionID2.Text), LngUnitID, val(Me.DCboStore2Name.BoundText))
                           mIsFromMix = False
                       Else
                           mIsFromMix = True
                    '   getItemCostData XPDtbBill.value, CLng(LngCurItemID), val(DCboStore2Name.BoundText), val(Me.TXTTransactionID2.Text), OldQty, OldCost, NewQty, NewCost
                       End If
                    End If
         
                     If mCostPrice <> 0 Then
                         FG2.TextMatrix(i, FG2.ColIndex("cost")) = mCostPrice
                     
                     End If
                 Next
            End If
            
    '             If costPrice = 0 Then
'                costPrice = ModItemCostPrice.GetCostItemPrice(CLng(LngCurItemID), , , , SystemOptions.SysMainStockCostMethod, , , Date, , LngUnitID)
'             End If
'
        
        
        CalcCostPercent
      
        s = "Select * from TblDefComItemData"
        saveGrid s, FG2, "itemcode", "", "IDDefCIT", val(TxtTransSerial)
        
        If Not IsSaveWithOutMsg Then
            
            DeleteTransactiomsVoucher2 val(TXTTransactionID1.Text)
            TXTTransactionID1.Text = ""
            DeleteTransactiomsVoucher2 val(TXTTransactionID2.Text)
            TXTTransactionID2.Text = ""
            DeleteTransactiomsVoucher2 val(TXTTransactionID5.Text)
            TXTTransactionID5.Text = ""
         
            DeleteTransactiomsVoucher2 val(TXTTransactionID4.Text)
            TXTTransactionID4.Text = ""
            TxtNoteSerial11 = ""
            TxtNoteSerial12 = ""
            TxtNoteSerial13 = ""
            TxtNoteSerial14 = ""
            TxtNoteSerial15 = ""
        End If
            'Selct(1).value = vbChecked
            If Selct(1).value = vbChecked Then
                BranchID = val(dcBranch.BoundText)
                 StoreId = val(DCboStore2Name.BoundText)
                If Not SystemOptions.DontCreateOut Then
                    createVoucher BranchID, 0, XPDtbBill.value, 27, 0, val(user_id), 0, 2, StoreId, 0, 0, "”‰œ  ’—› »‰«¡ ⁄·Ì  Ã„Ì⁄ —ﬁ„ " & TxtTransSerial
                
                End If
            End If
    
       'Selct(2).value = vbChecked
          '  If Selct(2).value = vbChecked Then
            
                BranchID = val(dcBranch.BoundText)
                StoreId = val(DCboStore3Name.BoundText)
                createVoucher1 BranchID, 0, XPDtbBill.value, 28, 0, val(user_id), 0, 2, StoreId, 0, 0, "”‰œ  «” ·«„  »‰«¡ ⁄·Ì  Ã„Ì⁄" & TxtTransSerial
             
         '   End If
                BranchID = val(dcBranch.BoundText)
                StoreId = val(DCboStore2Name.BoundText)
    
            
                StrSQL = "UPDATE TblDefComItem SET  TransactionID1=" & val(TXTTransactionID1) & ",  NoteSerial11='" & TxtNoteSerial11 & "' WHERE ID  =" & val(TxtTransSerial)
                Cn.Execute StrSQL
                StrSQL = "UPDATE TblDefComItem SET  TransactionID2=" & val(TXTTransactionID2) & ",  NoteSerial12='" & TxtNoteSerial12 & "' WHERE ID  =" & val(TxtTransSerial)
                Cn.Execute StrSQL
                If IsSaveWithOutMsg Then
                    If val(TXTTransactionID3) <> 0 Then
                         '   DeleteTransactiomsVoucher2 val(TXTTransactionID5.Text)
        
        
      '  If Selct(1).value = vbChecked Then
                            BranchID = val(dcBranch.BoundText)
                            StoreId = val(DCboStoreName.BoundText)
                        '    cmdTransfer_Click
                          If Not SystemOptions.DontCreateOut2 Then
                            If Trim(TxtNoteSerial13) <> "" Then
                            createVoucher BranchID, 0, XPDtbBill.value, 19, 0, val(user_id), 0, val(DBCboClientName.BoundText), StoreId, 0, 0, "”‰œ  ’—› »‰«¡ ⁄·Ì ›« Ê—… „»Ì⁄«  —ﬁ„ : " & TxtNoteSerial13 & " »‰«¡« ⁄·Ï ”‰œ  Ã„Ì⁄ —ﬁ„ " & TxtTransSerial, 1
                            End If
                         End If
                    End If
                End If
        
        Cn.CommitTrans
        BeginTrans = False
        XPTxtCurrent.Caption = rs.AbsolutePosition
        XPTxtCount.Caption = rs.RecordCount
         
            cmdCreateProduction.Enabled = True
            If Not SystemOptions.UserInterface = EnglishInterface Then
                cmdCreateProduction.Caption = "«‰‘«¡ «„— «‰ «Ã"
            Else
                cmdCreateProduction.Caption = "Create a production order"

            End If
            
            cmdCancel2.Visible = False
         rs.Resync
         
'***********************
        If IsSaveWithOutMsg Then Exit Sub

        '   CmdIssueVoucher_Click
    
        Select Case Me.TxtModFlg.Text

            Case "N"
    If SystemOptions.UserInterface = ArabicInterface Then
                    MsgBox " „ Õ›Ÿ  «·»Ì«‰«  ", vbInformation + vbOKOnly + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
                Else
                    MsgBox "Saved Data Successfully", vbInformation + vbOKOnly + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
                End If
      
            
            
            Case "E"

                If SystemOptions.UserInterface = ArabicInterface Then
                    MsgBox " „ Õ›Ÿ Â–Â «· ⁄œÌ·« ", vbInformation + vbOKOnly + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
                    
                Else
                    MsgBox "Saved Changes Successfully", vbInformation + vbOKOnly + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
                End If
                Retrive val(TxtTransSerial), False
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
            Msg = "·« Ì„ﬂ‰ Õ›Ÿ Â–Â «·»Ì«‰«  " & CHR(13)
            Msg = Msg + "·ﬁœ  „ «œŒ«· ﬁÌ„ €Ì— ’«·Õ… " & CHR(13)
            Msg = Msg + " √ﬂœ „‰ œﬁ… «·»Ì«‰«  Ê√⁄œ «·„Õ«Ê·…"
        Else
    
            Msg = "Cant Save Error"
        End If

        MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
        Exit Sub
    End If

    If SystemOptions.UserInterface = ArabicInterface Then
        Msg = "⁄›Ê«...ÕœÀ Œÿ√ „« √À‰«¡ Õ›Ÿ Â–Â «·»Ì«‰«  " & CHR(13)
    Else
        Msg = "Sorry... Error During Saving " & CHR(13)
    End If

    MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
End Sub

Private Sub XPTxtDiscountVal_KeyPress(KeyAscii As Integer)
    KeyAscii = KeyAscii_Num(KeyAscii, Me.XPTxtDiscountVal.Text, 0)
End Sub
Private Sub XPCboDiscountType_Change()
    XPCboDiscountType_Click
End Sub

Private Sub XPCboDiscountType_Click()
    On Error GoTo ErrTrap

    If XPCboDiscountType.ListIndex = 0 Or XPCboDiscountType.ListIndex = 3 Or XPCboDiscountType.ListIndex = -1 Then
    
        XPTxtDiscountVal.Enabled = False
        XPTxtDiscountVal.Text = ""
    Else
    
        XPTxtDiscountVal.Enabled = True
        XPTxtDiscountVal.Text = ""
    End If

    If Me.TxtModFlg.Text = "N" Or Me.TxtModFlg.Text = "E" Then
        If FG2.TextMatrix(1, FG2.ColIndex("itemcode")) <> "" Then
            CalcGrid2
        End If
    End If

    Me.lbl(67).Visible = (Me.XPCboDiscountType.ListIndex = 2)

    'Me.lbl(21).Visible = (Me.XPCboDiscountType.ListIndex = 2)
    If XPCboDiscountType.ListIndex = 0 Then
        lbl(68).Visible = False
        XPTxtDiscountVal.Visible = False
        lbl(68).Visible = False
    Else
        lbl(68).Visible = True
        XPTxtDiscountVal.Visible = True
        lbl(68).Visible = True
    End If

    Exit Sub
ErrTrap:
End Sub
Private Sub CalcDisc(ByVal RowNum As Long)
Dim DblDiscountTotal  As Double
DblDiscountTotal = 0
Dim DblRowTotal  As Double
Dim s As String
Dim rsDummy As New ADODB.Recordset
Dim mGroupID As Long
Dim mMaxPercent As Double
        Dim discountvalue As Double
        Dim DblPrice As Double
        With FG2
        
        mGroupID = val(.TextMatrix(RowNum, .ColIndex("GroupID")))
        s = "Select MaxPercentDiscount From groups Where GroupId =  " & mGroupID
        rsDummy.Open s, Cn, adOpenKeyset, adLockReadOnly, adCmdText
        If Not rsDummy.EOF Then
            mMaxPercent = val(rsDummy!MaxPercentDiscount & "")
        End If


        If .ColIndex("Trans_DiscountType") <> -1 Then

          '  For RowNum = 1 To .Rows - 1
                DblRowTotal = Round(val(.TextMatrix(RowNum, .ColIndex("Total"))), SystemOptions.SysDefCurrencyForamt)
                
                
          
                Select Case val(.TextMatrix(RowNum, .ColIndex("Trans_DiscountType")))

                                Case 0
                                   ' .TextMatrix(RowNum, .ColIndex("Valu")) = DblRowTotal
                                    
                                    .TextMatrix(RowNum, .ColIndex("Trans_Discount")) = 0
                                    DblDiscountTotal = DblDiscountTotal + 0
                                    discountvalue = 0
            
                                Case 1
                                    .TextMatrix(RowNum, .ColIndex("Total")) = DblRowTotal
                                    .TextMatrix(RowNum, .ColIndex("Trans_Discount")) = 0
                                    
                                    DblDiscountTotal = DblDiscountTotal + 0
                                    discountvalue = 0
            
                                Case 2
'                                    .TextMatrix(RowNum, .ColIndex("Valu")) = (DblRowTotal) - val(.TextMatrix(RowNum, .ColIndex("DiscountVal")))
                             '     DblRowTotal = .TextMatrix(RowNum, .ColIndex("Valu"))
                                    DblDiscountTotal = DblDiscountTotal + val(.TextMatrix(RowNum, .ColIndex("Trans_Discount")))
                                    discountvalue = val(.TextMatrix(RowNum, .ColIndex("Trans_Discount")))
            
                                Case 3
'                                    .TextMatrix(RowNum, .ColIndex("Valu")) = (DblRowTotal) * (1 - (val(.TextMatrix(RowNum, .ColIndex("DiscountVal"))) / 100))
                                    DblRowTotal = .TextMatrix(RowNum, .ColIndex("Total"))
                                    DblDiscountTotal = DblDiscountTotal + ((val(.TextMatrix(RowNum, .ColIndex("Trans_Discount"))) * val(.TextMatrix(RowNum, .ColIndex("Total")))) / 100)
                                    discountvalue = ((val(.TextMatrix(RowNum, .ColIndex("Trans_Discount"))) * DblRowTotal) / 100)
            
                                Case 4
                              
                                DblRowTotal = .TextMatrix(RowNum, .ColIndex("Total"))
                                      .TextMatrix(RowNum, .ColIndex("Total")) = 0
                                    DblDiscountTotal = DblDiscountTotal + DblRowTotal
                                    discountvalue = DblRowTotal
                            End Select
               
 If Not SystemOptions.AllowSkipDiscountGroup Then
    If mMaxPercent < (discountvalue / IIf(DblRowTotal <> 0, DblRowTotal, 1) * 100) Then
        discountvalue = 0
        .TextMatrix(RowNum, .ColIndex("Trans_Discount")) = 0
        .TextMatrix(RowNum, .ColIndex("Total")) = DblRowTotal
        
        If SystemOptions.UserInterface = ArabicInterface Then
            MsgBox "·« Ì„ﬂ‰ «·”„«Õ » ŒÿÏ ‰”»… «·Œ’„ «·„Õœœ… ÿ»ﬁ« ··„Ã„Ê⁄…"
        Else
            MsgBox "·You can not allow the specified discount rate to be skipped according to the group"
        End If
        Exit Sub
    End If
 End If
 .TextMatrix(RowNum, .ColIndex("TotalDisc")) = val(.TextMatrix(RowNum, .ColIndex("TotalDisc"))) + discountvalue
 'Next
 End If
 End With
End Sub
Private Sub XPChkTAX_Click()
    On Error GoTo ErrTrap

    If XPChkTAX.value = Checked Then
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



Private Sub Form_QueryUnload(Cancel As Integer, _
                             UnloadMode As Integer)
    Dim StrMSG As String
    Dim IntResult As String
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
   cmd(8).Caption = "Delete"
   cmd(10).Caption = "Same Copy"
    lbl(5).Caption = "No"
    lbl(32).Caption = "Qty"
    lbl(33).Caption = "Unit"
    lbl(38).Caption = "Item"
    lbl(39).Caption = "Code"
    lbl(40).Caption = "Qty"
    ISButton4.Caption = "Select Path"
    ISButton3.Caption = "Imports"
    CMDSHOWISSUE.Caption = "Show exchange into action"
    CMDSHOWecive.Caption = "Show work to receive"
        lbl(74).Caption = "Remark"
 ISButton1(4).Caption = "Short print"
   lbl(77).Caption = "Group"
   lbl(79).Caption = "Item Code"
      lbl(11).Caption = "Group"
      chkIsAdd.Caption = "Add"
      ISButton2.Caption = "Attachments"
      lbl(83).Caption = "Diameter"
   lbl(86).Caption = "Max NO"
   lbl(78).Caption = "Item"
  lbl(66).Caption = "Discount type"
  cmdTransfer.Caption = "Transfer"
        ISButton1(3).Caption = "Raw material printing"
    lbl(25).Caption = "Raw Materials"
  Me.Caption = "Definition of varieties / assemble components"
    lbl(6).Caption = "Date"
     lbl(36).Caption = "Branch"
        lbl(50).Caption = "Store"
        Selct(0).Caption = "Customize components"
        Selct(1).Caption = "Exchange into action"
        Selct(2).Caption = "Work to receive"
        lbl(47).Caption = "Select Store"
         lbl(48).Caption = "Select Store"
         lbl(42).Caption = "Customer"
       cmdAdd_.Caption = "Add"
       chkHiddLogo.Caption = "Hide the logo"
lbl(17).Caption = "ItemNo"
lbl(26).Caption = "ItemName"
lbl(29).Caption = "MaixNo"
lbl(30).Caption = "MaixName"
lbl(27).Caption = "Unit"
Ele(8).Caption = "Create Vouchers"
CmdAdd.Caption = "Add"
    Ele(6).Caption = Me.Caption
    
    lbl(1).Caption = "By"
    lbl(0).Caption = "Currenr rec."
    lbl(2).Caption = "Total rec."

    cmd(0).Caption = "New"
    cmd(1).Caption = "Edit"
    cmd(2).Caption = "Save"
    cmd(3).Caption = "Undo"
    cmd(4).Caption = "Delete"
    cmd(5).Caption = "Search"
    cmd(7).Caption = "Print"
    cmd(6).Caption = "Exit"
    CmdHelp.Caption = "Help"
lbl(65).Caption = "Components"
lbl(80).Caption = "Rec Date"
lbl(81).Caption = "Rec Time"
lbl(84).Caption = "Tel"
lbl(58).Caption = "Sales Man"
lbl(59).Caption = "Box"
lbl(51).Caption = "Model"
Command1.Caption = "Show Pic"
cmdAddCustomer.Caption = "Add New Customer"
lbl(76).Caption = "Customer Name"
lbl(54).Caption = "Payment Method"

chkSelectAll(0).Caption = "Select All"
chkSelectAll(1).Caption = "Select All"
lbl(25).Caption = "Total"
lbl(44).Caption = "Total Add"
lbl(60).Caption = "Total Disc"
lbl(61).Caption = "Net"
lbl(62).Caption = "Vat"
lbl(63).Caption = "Net Value"
cmd(11).Caption = " delete"

lbl(41).Caption = "widtj"
lbl(49).Caption = "hight"
lbl(82).Caption = "Length"
lbl(64).Caption = "Type"
lbl(52).Caption = "Price"
lbl(57).Caption = "Total"
lbl(75).Caption = "Period"
lbl(68).Caption = "Value"
lbl(71).Caption = "Remark"
lbl(99).Caption = "Net Value"

lbl(99).Caption = "Net Value"
lbl(98).Caption = "Vat "
lbl(56).Caption = "Net "

lbl(55).Caption = "Total Disc"
lbl(53).Caption = "Total Add"

lbl(72).Caption = "Item Code"
lbl(69).Caption = "Row Item"
lbl(73).Caption = "Unit"
lbl(70).Caption = "Qty"
lbl(85).Caption = "thickness"
cmdAdd2.Caption = "Add Rows"

cmdCancel.Caption = "Cancel conversion"
cmdfrmRec.Caption = "Catch a down payment"
cmdCreateSales.Caption = "Display invoice"
CMDSHOWISSUE2.Caption = "Display "
cmdCancel2.Caption = "Cancel production order"
cmdCreateProduction.Caption = "Create an output command"
 '   Me.XPTab301.TabCaption(0) = "Data "
 
 ISButton1(1).Caption = "Print Offer Price"
 ISButton1(2).Caption = "Print receipt receipt"
 ISButton1(0).Caption = "Print invoice"

TabMain.TabCaption(0) = "Deleted Items"
TabMain.TabCaption(1) = "Rows"
TabMain.TabCaption(2) = "Data"
    With Me.FG

  .TextMatrix(0, .ColIndex("Ser")) = " Serial "
  .TextMatrix(0, .ColIndex("itemcode")) = "itemNo "
        .TextMatrix(0, .ColIndex("itemname")) = "ItemName "
        .TextMatrix(0, .ColIndex("unitname")) = "Unitn"
        .TextMatrix(0, .ColIndex("cost")) = "Cost "
          .TextMatrix(0, .ColIndex("FlgX")) = "Basic Qty "
           .TextMatrix(0, .ColIndex("Qty")) = "Produc.Qty "
           .TextMatrix(0, .ColIndex("Total")) = "Total"
        .TextMatrix(0, .ColIndex("name1")) = "Specifications1 "
        .TextMatrix(0, .ColIndex("name2")) = "Specifications2"
        .TextMatrix(0, .ColIndex("name3")) = "Specifications3"
          .TextMatrix(0, .ColIndex("name4")) = " Specifications4 "
        .TextMatrix(0, .ColIndex("Amout1")) = "Amout1 "
        .TextMatrix(0, .ColIndex("Amout2")) = "Amout2"
        .TextMatrix(0, .ColIndex("Amout3")) = "Amout3 "
 .TextMatrix(0, .ColIndex("Amout4")) = "Amout4 "
 .TextMatrix(0, .ColIndex("Remarks")) = "Remarks "
 .TextMatrix(0, .ColIndex("ShowAttatch")) = "ShowAttatch"
 .TextMatrix(0, .ColIndex("Select")) = "Select"
 .TextMatrix(0, .ColIndex("IsAdd")) = "IsAdd"
 .TextMatrix(0, .ColIndex("OldPrice")) = "Old Price"
 .TextMatrix(0, .ColIndex("IsAdd")) = "IsAdd"
 .TextMatrix(0, .ColIndex("ItemName2")) = "Main Item"
 .TextMatrix(0, .ColIndex("Price")) = "Price "
 .TextMatrix(0, .ColIndex("lowering")) = "lowering"
 .TextMatrix(0, .ColIndex("Increase")) = "Increase"
 
 
    End With

With FG2

    .TextMatrix(0, .ColIndex("Select")) = "Select"
    .TextMatrix(0, .ColIndex("Ser")) = "Ser"
    .TextMatrix(0, .ColIndex("GroupName")) = "Group Name"
    .TextMatrix(0, .ColIndex("itemcode")) = "item code"
    .TextMatrix(0, .ColIndex("increase")) = "Increase"
    .TextMatrix(0, .ColIndex("thickness")) = "Thickness"
    
    .TextMatrix(0, .ColIndex("itemname")) = "ItemName "
    .TextMatrix(0, .ColIndex("unitname")) = "Unitn"
    .TextMatrix(0, .ColIndex("cost")) = "Cost "
    .TextMatrix(0, .ColIndex("Price")) = "Price "
    .TextMatrix(0, .ColIndex("Qty")) = "Produc.Qty "
    .TextMatrix(0, .ColIndex("Diameter")) = "Diameter "
    
    .TextMatrix(0, .ColIndex("ItemName2")) = "Features"
    .TextMatrix(0, .ColIndex("ItemName2")) = "Features2"
    
    .TextMatrix(0, .ColIndex("CountItem2")) = "Count"
    .TextMatrix(0, .ColIndex("CountItem5")) = "Count"
    
    
    .TextMatrix(0, .ColIndex("BuiltinItemName")) = "Builtin Item"
    .TextMatrix(0, .ColIndex("GroupBuiltinName")) = "Builtin Group"
    .TextMatrix(0, .ColIndex("lowering")) = "lowering"
    .TextMatrix(0, .ColIndex("widtj")) = "widtj"
    .TextMatrix(0, .ColIndex("hight")) = " hight "
    .TextMatrix(0, .ColIndex("Length")) = "Length "
    .TextMatrix(0, .ColIndex("Total")) = "Total"
    .TextMatrix(0, .ColIndex("Trans_DiscountType")) = "Discount Type "
    .TextMatrix(0, .ColIndex("Trans_Discount")) = "Trans_Discount "
    .TextMatrix(0, .ColIndex("TotalDisc")) = "Total Disc"
    .TextMatrix(0, .ColIndex("TotalAdd")) = "Total Add"
    .TextMatrix(0, .ColIndex("Net")) = "Net"
    .TextMatrix(0, .ColIndex("Vat2")) = "Vat2"
    .TextMatrix(0, .ColIndex("TotalWithVat")) = "Total With Vat"
    .TextMatrix(0, .ColIndex("Remark")) = "Remark"
    .TextMatrix(0, .ColIndex("NoteSerial14")) = "Product Order No"
    
End With
   With Me.FGDeleted

  .TextMatrix(0, .ColIndex("Ser")) = " Serial "
  .TextMatrix(0, .ColIndex("itemcode")) = "itemNo "
        .TextMatrix(0, .ColIndex("itemname")) = "ItemName "
        .TextMatrix(0, .ColIndex("unitname")) = "Unitn"
        .TextMatrix(0, .ColIndex("cost")) = "Cost "
          .TextMatrix(0, .ColIndex("FlgX")) = "Basic Qty "
           .TextMatrix(0, .ColIndex("Qty")) = "Produc.Qty "
           .TextMatrix(0, .ColIndex("Total")) = "Total"
           .TextMatrix(0, .ColIndex("Price")) = "Price "
        .TextMatrix(0, .ColIndex("name1")) = "Specifications1 "
        .TextMatrix(0, .ColIndex("name2")) = "Specifications2"
        .TextMatrix(0, .ColIndex("name3")) = "Specifications3"
          .TextMatrix(0, .ColIndex("name4")) = " Specifications4 "
        .TextMatrix(0, .ColIndex("Amout1")) = "Amout1 "
        .TextMatrix(0, .ColIndex("Amout2")) = "Amout2"
        .TextMatrix(0, .ColIndex("Amout3")) = "Amout3 "
 .TextMatrix(0, .ColIndex("Amout4")) = "Amout4 "
 .TextMatrix(0, .ColIndex("Remarks")) = "Remarks "
 
 .TextMatrix(0, .ColIndex("Remarks")) = "Remarks "
 .TextMatrix(0, .ColIndex("Redo")) = "Redo"
 
    End With


End Sub



Private Sub XPDtbBill_Change()
              TxtNoteSerial11.Text = ""
              TxtNoteSerial12.Text = ""
              If Me.TxtModFlg <> "R" Then
                XPDtRecDate.value = DateAdd("d", 3, XPDtbBill.value)
              End If
End Sub


Sub CalCulteVAT(Optional Ind As Integer = 0, Optional ByVal mRow As Long)
Dim AccountVATCreit As String
Dim Percetage As Double

Dim mVal As Double

    If Ind = 3 Then
        PercentgValueAddedAccount_Transec XPDtbBill.value, 21, 0, AccountVATCreit, Percetage
        PercetageVat = Percetage
       ' Percetage = 5
  If SystemOptions.PriceWithVAT = True Then
        TxtVAt2.Text = 0
        
         TxtVATValue.Text = 0
         TxtVAt2.Text = 0
         
         
         mVal = 0
         TxtVATValue.Text = 0
         txtTotalWithVat.Text = 0
         
         
  Else
        TxtVAt2.Text = val(Format((txtNet.Text), "###.00")) * Percetage / 100
        
         TxtVATValue.Text = val(Format((txtNet.Text), "###.00")) * Percetage / 100
         TxtVAt2.Text = TxtVATValue.Text
         
         
         mVal = val(Format((txtNet.Text), "###.00"))
         TxtVATValue.Text = val(Format((mVal), "###.00")) * Percetage / 100
         txtTotalWithVat.Text = Round(val(Format((mVal), "###.00")) + val(TxtVATValue.Text), 2)
         
   End If

         
         
         If mRow <> 0 Then
          If SystemOptions.PriceWithVAT = True Then
            FG2.TextMatrix(mRow, FG2.ColIndex("VAt2")) = 0
            FG2.TextMatrix(mRow, FG2.ColIndex("TotalWithVat")) = (val(Format((FG2.TextMatrix(mRow, FG2.ColIndex("Net"))), "###.00")) + val(FG2.TextMatrix(mRow, FG2.ColIndex("VAt2"))))
            Else
            FG2.TextMatrix(mRow, FG2.ColIndex("VAt2")) = val(FG2.TextMatrix(mRow, FG2.ColIndex("Net"))) * Percetage / 100
            FG2.TextMatrix(mRow, FG2.ColIndex("TotalWithVat")) = (val(Format((FG2.TextMatrix(mRow, FG2.ColIndex("Net"))), "###.00")) + val(FG2.TextMatrix(mRow, FG2.ColIndex("VAt2"))))
            End If
            
         End If
         
'         Exit Sub
    End If
    'XPDtbBill.value = 100
    'XPTxtVal = 100
     txtTotalWithVat.Text = (val(Format((mVal), "###.00")) + val(TxtVATValue.Text))
     TxtVAt2.Text = TxtVATValue.Text
    
    
'    For i = 1 To FG2.Rows - 1
'        If val(FG2.TextMatrix(i, FG2.ColIndex("ItemID"))) = val(DcboItemID1.BoundText) Then
'            FG2.TextMatrix(i, FG2.ColIndex("VAt2")) = TxtVAt2
'            FG2.TextMatrix(i, FG2.ColIndex("TotalWithVat")) = txtTotalWithVat
'            FG2.TextMatrix(i, FG2.ColIndex("Total")) = txtTotal
'            FG2.TextMatrix(i, FG2.ColIndex("Net")) = txtNet
'            FG2.TextMatrix(i, FG2.ColIndex("TotalAdd")) = txtTotalAdd
'            FG2.TextMatrix(i, FG2.ColIndex("TotalDisc")) = txtTotalDisc
'            FG2.TextMatrix(i, FG2.ColIndex("Qty")) = txtQty1
'            FG2.TextMatrix(i, FG2.ColIndex("Price")) = txtPrice
'
'
'        End If
'    Next
    
End Sub


Private Sub CreateProduction(BranchID As Double, _
BoxID As Double, _
Transaction_Date As Date, _
Transaction_Type As Double, _
CBoBasedON As Double, _
UserID As Double, _
Trans_DiscountType As Double, _
CusID As Double, _
StoreId As Double, _
PaymentType As Double, _
Emp_id As Double, _
TransactionComment As String, ByVal mmID As Long, Transaction_ID As Long)

Dim BolTemp As Boolean
Dim sql As String
Dim Msg As String
Dim NoteID As Long

Dim Transaction_ID1 As Long
Dim Transaction_serial As String
Dim NoteSerial As String
Dim NoteSerial1 As String
Dim StrSQL As String
Dim Percetage As Double
Dim AccountVATCreit As String
Dim mPrice As Double
Dim rsDummy As New ADODB.Recordset
' «·”⁄— Â‰« ÂÊ ’«›Ï «·”⁄— »⁄œ Œ’„ «·«÷«›Ï Ê«·Œ’Ê„« 
'
'PercentgValueAddedAccount_Transec XPDtbBill.value, 21, 0, AccountVATCreit, Percetage
'PercetageVat = Percetage

'BillTOTAL = 0




 
Dim RSTransDetails As New ADODB.Recordset
     
StrSQL = "SELECT     dbo.Transaction_Details.* from dbo.Transaction_Details Where (Transaction_ID = -1)"
RSTransDetails.Open StrSQL, Cn, adOpenKeyset, adLockOptimistic, adCmdText
   
StrSQL = "Select ItemID,UnitID,Sum(Qty) Qty,Sum(Cost) Price,Sum(Cost) Cost,Sum(TotalWithVat) TotalWithVat,Sum(PercentCost) PercentCost from TblDefComItemData Where ID = " & mmID
StrSQL = StrSQL & " And  ItemId In (Select ItemId2 From TblDefComItemDet Det Where IsNull(Det.IsDeleted,0) <> 1 and Det.ItemID <> Det.ItemId2 "
StrSQL = StrSQL & " and Det.IDDefCIT =" & val(TxtTransSerial.Text) & ") "
StrSQL = StrSQL & " Group By ItemID,UnitID "
rsDummy.Open StrSQL, Cn, adOpenKeyset, adLockOptimistic
    If Not rsDummy.EOF Then
    
        Dim mItemNo As Long, mUnitNo As Long, mQty As Long, mVAt2 As Double, mTotal As Double
        Dim mwidtj As Double, mhight As Double, mTotalAdd As Double, mTotalDisc As Double, mNet As Double, mTotalWithVat As Double, mLength As Double
        Dim mItemName2 As String, mCostPercent As Double
        Dim mRemark As String
        mItemNo = val(rsDummy!ItemID & "")
        If mItemNo = 0 Then GoTo NextRow
        
               
            mItemNo = val(rsDummy!ItemID & "")
           
            mUnitNo = val(rsDummy!UnitID & "")
            mQty = val(rsDummy!Qty & "")
            mPrice = val(rsDummy!Price & "")
'            mwidtj = val(rsDummy!widtj & "")
'            mhight = val(rsDummy!hight & "")
'            mLength = val(rsDummy!Length & "")
           ' mTotal = val(rsDummy!Total & "")
        '    mRemark = Trim(rsDummy!Remark & "")
        '    mTotalDisc = val(rsDummy!TotalDisc & "")
        '    mTotalAdd = val(rsDummy!TotalAdd & "")
        '    mNet = val(rsDummy!net & "")
        '    mVAt2 = val(rsDummy!Vat2 & "")
            mTotalWithVat = val(rsDummy!TotalWithVat & "")
            mPrice = (val(mTotal) + val(mTotalAdd)) / val(mQty)
            mCostPercent = val(rsDummy!PercentCost & "")
            
        RSTransDetails.AddNew
        RSTransDetails("Transaction_ID").value = Transaction_ID
        RSTransDetails("ColorID").value = 1
        RSTransDetails("ItemSize").value = 1
        RSTransDetails("ClassId").value = 1
        RSTransDetails("Item_ID").value = mItemNo
        RSTransDetails("UnitID").value = mUnitNo
        RSTransDetails("SHOWQTY").value = mQty
        RSTransDetails("PercentCost").value = mCostPercent
        RSTransDetails("showPrice").value = mPrice
        RSTransDetails("Lineexpenses").value = mPrice
        
        RSTransDetails("ItemDiscountType").value = 2
        
        If SystemOptions.TypicalProduction = False Then

            RSTransDetails("CostPrice").value = ModItemCostPrice.GetCostItemPrice(mItemNo, 0, , , SystemOptions.SysMainStockCostMethod, , , XPDtbBill.value, val(Me.Text1.Text), RSTransDetails("UnitID").value, StoreId)

            If RSTransDetails("CostPrice").value = 0 Then
                RSTransDetails("CostPrice").value = ModItemCostPrice.GetCostItemPrice(mItemNo, 0, , , LastPurPriceType, , , XPDtbBill.value, val(Me.Text1.Text), RSTransDetails("UnitID").value, val(Me.DCboStore2Name.BoundText))
                
            End If
              
        Else
            RSTransDetails("CostPrice").value = 0
        
        End If
                      
          
                      '«·ÊÕœ« 
       
        Dim RsUnitData As ADODB.Recordset
        Dim LngCurItemID As Long
        Dim LngUnitID As Long
        Dim DblQty As Double
    
        LngCurItemID = val(mItemNo)
        LngUnitID = val(mUnitNo)
        DblQty = val(mQty)

        StrSQL = "Select * From TblItemsUnits Where ItemID=" & LngCurItemID
        StrSQL = StrSQL + " AND UnitID=" & LngUnitID
        Set RsUnitData = New ADODB.Recordset
        RsUnitData.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

        If Not (RsUnitData.BOF Or RsUnitData.EOF) Then
            RSTransDetails("QtyBySmalltUnit").value = RsUnitData("UnitFactor").value
            RSTransDetails("Quantity").value = RSTransDetails("QtyBySmalltUnit").value * RSTransDetails("showqty").value
            RSTransDetails("OpeningSalesQty").value = RSTransDetails("Quantity").value
            RSTransDetails("OpeningSalesValue").value = RSTransDetails("CostPrice").value * val(mQty)
            RSTransDetails("Price").value = val(IIf((mPrice = 0), 0, val(mPrice))) / RSTransDetails("QtyBySmalltUnit").value
        
        End If

    
         UpdateTransactionsCost CStr(Transaction_ID)
         RSTransDetails.update
    
      '  Dim i As Integer
        'Dim sql As String
    End If
NextRow:


NoteSerial = Notes_coding(val(BranchID), Transaction_Date)





'***********************
         StrSQL = "UPDATE TblDefComItem SET  TransactionID4=" & val(TXTTransactionID4) & ",  NoteSerial14='" & TxtNoteSerial14 & "' WHERE ID  =" & val(TxtTransSerial)
         Cn.Execute StrSQL
'***********************
If Not SystemOptions.IsMultiItemsInCompItem Then
        cmdCancel2.Visible = True
        cmdCancel2.Enabled = True
        If Not SystemOptions.UserInterface = EnglishInterface Then

            cmdCreateProduction.Caption = "⁄—÷ «„— «·«‰ «Ã"
            MsgBox " „   «‰‘«¡ «„— «·«‰ «Ã"
        Else
             cmdCreateProduction.Caption = "Display product order"
             MsgBox "Production order was created"
        End If
        'StrSQL = "UPDATE Transactions SET NOTS=" & Transaction_ID & " WHERE Transaction_ID=" & val(Me.XPTxtBillID.text)
        'Cn.Execute StrSQL
    
    
End If
  
'******************************************************issueVoucher








     
 
    '
 
ErrTrap:



 

End Sub




Private Sub CreateSalesTrans(BranchID As Double, _
BoxID As Double, _
Transaction_Date As Date, _
Transaction_Type As Double, _
CBoBasedON As Double, _
UserID As Double, _
Trans_DiscountType As Double, _
CusID As Double, _
StoreId As Double, _
PaymentType As Double, _
Emp_id As Double, _
TransactionComment As String)

Dim BolTemp As Boolean
Dim sql As String
Dim Msg As String
Dim NoteID As Long
Dim Transaction_ID As Long
Dim Transaction_ID1 As Long
Dim Transaction_serial As String
Dim NoteSerial As String
Dim NoteSerial1 As String
Dim StrSQL As String
Dim Percetage As Double
Dim AccountVATCreit As String
Dim mPrice As Double

' «·”⁄— Â‰« ÂÊ ’«›Ï «·”⁄— »⁄œ Œ’„ «·«÷«›Ï Ê«·Œ’Ê„« 

PercentgValueAddedAccount_Transec XPDtbBill.value, 21, 0, AccountVATCreit, Percetage
PercetageVat = Percetage

'BillTOTAL = 0
CostTOTAL = 0
'Check
StoreId = val(DCboStoreName.BoundText)
  
    If DCboStoreName.Text = "" Then
        If SystemOptions.UserInterface = ArabicInterface Then
            Msg = "ÌÃ»  ÕœÌœ «·„Œ“‰"
        Else
            Msg = "Select Inventory First"
        End If

        MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
      If DCboStoreName.Enabled = True Then
        DCboStoreName.SetFocus
      SendKeys "{F4}"
        End If
       cmd(2).Enabled = True
        Screen.MousePointer = vbDefault
      '  Cmd(2).Enabled = True
        Exit Sub
    End If
    
 If Trim(DcboEmp.BoundText) = "" Then
        If SystemOptions.UserInterface = ArabicInterface Then
            Msg = "ÌÃ»  ÕœÌœ «”„ «·»«∆⁄/«·„‰œÊ»..!!!"
        Else
            Msg = "Must Specify SalesPerson/Saller..!!!"
        End If
cmd(2).Enabled = True
        MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
        DcboEmp.SetFocus
        SendKeys "{F4}"
        Screen.MousePointer = vbDefault
       ' Cmd(2).Enabled = True
        Exit Sub
    End If
    

 If TxtNoteSerial13 = "" Then
 NoteSerial1 = Voucher_coding(val(BranchID), Transaction_Date, 7, 170, , 21)
 TxtNoteSerial13 = NoteSerial1
 End If
Transaction_ID = CStr(new_id("Transactions", "Transaction_ID", "", True))
 
  
    NoteSerial1 = Voucher_coding(val(BranchID), Transaction_Date, 7, 170, , 21)  '„»Ì⁄« 
        If NoteSerial1 = "" Then
                 If NoteSerial1 = "error" Then
                     MsgBox " ·« Ì„ﬂ‰ «÷«›… ”‰œ   „»Ì⁄«   ÃœÌœ ·«‰ﬂ  ⁄œÌ  «·Õœ «·–Ì ﬁ„  » ÕœÌœ… „‰ «·”‰œ«   ": Exit Sub
                 ElseIf NoteSerial1 = "" Then
                         MsgBox " ·«»œ „‰ ﬂ «»… —ﬁ„ «·”‰œ ÌœÊÌ« ﬂ„« Õœœ   ": Exit Sub
        
                 End If
        End If

NoteSerial = Notes_coding(val(BranchID), Transaction_Date)
 If NoteSerial = "" Then
            If NoteSerial = "error" Then
                MsgBox " ·« Ì„ﬂ‰ «÷«›… ﬁÌÊœ ·Â–… «·⁄„·Ì… ·«‰ﬂ  ⁄œÌ  «·Õœ «·«ﬁ’Ì ··ﬁÌÊœ ﬂ„« Õœœ  ›Ì  —ﬁÌ„ «·”‰œ«  ": Exit Sub
            ElseIf NoteSerial = "" Then
                    MsgBox " ·«Ì„ﬂ‰ «‰‘«¡ «·ﬁÌœ ·Â–« «·„” ‰œ ·«‰ﬂ Õœœ   —ﬁÌ„ ﬁÌÊœ ÌœÊÌ  ": Exit Sub
                 
            End If
End If
           
              
  
   '«· √ﬂœ „‰ ⁄œ„  ﬂ—«— —ﬁ„ «·›« Ê—…
    If Voucher_coding(val(dcBranch.BoundText), XPDtbBill.value, 7, 170, , 21) = "" Then
        If Me.TxtModFlg.Text = "N" Then
    
            BolTemp = UniqueNoteSerial1(Trim(Me.TxtNoteSerial13.Text), 21, , val(dcBranch.BoundText))
        ElseIf Me.TxtModFlg.Text = "E" Then
        
            BolTemp = UniqueNoteSerial1(Trim(Me.TxtNoteSerial13.Text), 21, Transaction_ID, val(dcBranch.BoundText))
        End If
 
        If BolTemp = False Then
            If SystemOptions.UserInterface = ArabicInterface Then
                Msg = "—ﬁ„ «·›« Ê—… „”Ã· „”»ﬁ« ›Ï «·»—‰«„Ã.." & CHR(13)
                Msg = Msg & "Ê·«Ì„ﬂ‰  ﬂ—«— —ﬁ„ «·›« Ê—…"
            Else
                Msg = "This Bill No Already Exist" & CHR(13)
        
            End If
            cmd(2).Enabled = True
            MsgBox Msg, vbOKOnly + vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
            TxtNoteSerial13.SetFocus
            Screen.MousePointer = vbDefault
          '  Cmd(2).Enabled = True
            Exit Sub
        End If
     
    End If
      
  
'
'           CostAccount = get_account_code_branch(1, CInt(BranchID))
'
'            If CostAccount = "NO branch" Or CostAccount = "NO account" Then
'                If SystemOptions.UserInterface = ArabicInterface Then
'                    MsgBox "·„ Ì „ —»ÿ  ﬂ·›… «·«‰ «Ã „Ê«œ  ", vbCritical
'                Else
'                    MsgBox "Sales Not Created", vbCritical
'                End If
'
'             Exit Sub
'              End If
              
              
If SystemOptions.PaymentMethLaterCompItem Then
    TransactionComment = ""
End If
    StoreAccount = get_store_Account(CInt(StoreId), "Account_Code")
      If StoreAccount = "" Then
                If SystemOptions.UserInterface = ArabicInterface Then
                    MsgBox "·„ Ì „  ÕœÌœ Õ”«»  ··„Œ“Ê‰ «·”·⁄Ì ·Â–« «·„Œ“‰ ›Ì Â–« «·›—⁄    ", vbCritical
                Else
                    MsgBox "No inventory account for this store has been specified in this section", vbCritical
                End If
                Exit Sub
            End If



 'end Check
 
        TXTTransactionID3.Text = Transaction_ID
        TxtNoteSerial13.Text = NoteSerial1
     Dim rsOut As New ADODB.Recordset
            Dim Current_case As Integer, s As String, mBoxID As Long
            Set rsOut = New ADODB.Recordset
            s = "Select BoxID From TblBoxesData Where Empid = " & Me.DcboEmp.BoundText



            rsOut.Open s, Cn, adOpenStatic, adLockReadOnly
            If Not rsOut.EOF Then
                BoxID = val(rsOut!BoxID & "")
            End If
            mBoxID = val(DcboBox.BoundText)
 sql = "INSERT INTO  Transactions (  "
sql = sql & " Transaction_ID ,"
sql = sql & " BranchID ,"
sql = sql & " NoteSerial ,"
sql = sql & " NoteSerial1 ,"
sql = sql & " boxId ,"
sql = sql & " Transaction_serial ,"
sql = sql & " Transaction_Date ,"
sql = sql & " Transaction_Type ,"
sql = sql & " BillBasedOn ,"
sql = sql & " UserID ,"
sql = sql & " Trans_DiscountType ,"
sql = sql & " CusID ,"
sql = sql & " StoreId ,"
sql = sql & " PaymentType ,"
sql = sql & " Emp_id ,"
sql = sql & " Transaction_NetValue ,"
sql = sql & " Vat, netvalue, PayedValue, "
sql = sql & " Currency_rate, Currency_id,sumVatLine,DueDate,"
 sql = sql & " TransactionComment )"
 sql = sql & " VALUES("
sql = sql & " " & Transaction_ID & " ,"
sql = sql & " " & BranchID & " ,"
sql = sql & "'" & NoteSerial & "' ,"
sql = sql & "'" & NoteSerial1 & "' ,"
sql = sql & " " & BoxID & " ,"
sql = sql & "'" & Transaction_serial & "',"
sql = sql & " " & SQLDate(Transaction_Date, True) & " ,"
sql = sql & " " & Transaction_Type & " ,"
sql = sql & " 0 ,"
sql = sql & " " & user_id & " ,"
sql = sql & " 0 ,"
sql = sql & " " & CusID & " ,"
sql = sql & " " & StoreId & " ,"
sql = sql & " " & CboPayMentType.ListIndex & " ,"
sql = sql & " " & Emp_id & " ,"
sql = sql & " " & val(txtTotalWithVat2) & " ,"
sql = sql & " " & val(TxtVAt22) & " ,"
sql = sql & " " & val(txtNet2) & " ,"
sql = sql & " " & val(txtNet2) & " ,"
sql = sql & " " & 1 & " ,"
sql = sql & " " & 1 & " ,0,"
sql = sql & " " & SQLDate(Transaction_Date, True) & " ,"
sql = sql & "'" & TransactionComment & "')"
 
Cn.Execute sql
 



 
Dim RSTransDetails As New ADODB.Recordset
     
StrSQL = "SELECT     dbo.Transaction_Details.* from dbo.Transaction_Details Where (Transaction_ID = -1)"
RSTransDetails.Open StrSQL, Cn, adOpenKeyset, adLockOptimistic, adCmdText
   
For i = 1 To FG2.Rows - 1
    
    Dim mItemNo As Long, mUnitNo As Long, mQty As Double, mVAt2 As Double, mTotal As Double
    Dim mwidtj As Double, mhight As Double, mTotalAdd As Double, mTotalDisc As Double, mNet As Double, mTotalWithVat As Double, mLength As Double
    Dim mAreaL As String
    
    Dim mItemName2 As String
    Dim mCost As Double
    Dim mRemark As String
    mItemNo = val(FG2.TextMatrix(i, FG2.ColIndex("ItemID")))
    If mItemNo = 0 Then GoTo NextRow
    With FG2
           
        mItemNo = val(.TextMatrix(i, .ColIndex("ItemID")))
        mItemName2 = Trim(.TextMatrix(i, .ColIndex("ItemName2")))
        mUnitNo = val(.TextMatrix(i, .ColIndex("UnitID")))
        mQty = val(.TextMatrix(i, .ColIndex("Qty")))
        mPrice = val(.TextMatrix(i, .ColIndex("Price")))
        mCost = val(.TextMatrix(i, .ColIndex("Cost")))
        mwidtj = val(.TextMatrix(i, .ColIndex("widtj")))
        mhight = val(.TextMatrix(i, .ColIndex("hight")))
        mLength = val(.TextMatrix(i, .ColIndex("Length")))
        mTotal = val(.TextMatrix(i, .ColIndex("Total")))
        mRemark = Trim(.TextMatrix(i, .ColIndex("Remark")))
        mTotalDisc = val(.TextMatrix(i, .ColIndex("TotalDisc")))
        mTotalAdd = val(.TextMatrix(i, .ColIndex("TotalAdd")))
        mNet = val(.TextMatrix(i, .ColIndex("Net")))
        mAreaL = Trim(.TextMatrix(i, .ColIndex("AreaL")))
        mVAt2 = val(.TextMatrix(i, .ColIndex("Vat2")))
        mTotalWithVat = val(.TextMatrix(i, .ColIndex("TotalWithVat")))
        mPrice = (val(mTotal) + val(mTotalAdd)) / val(mQty)
        mAreaL = Trim(.TextMatrix(i, .ColIndex("AreaL")))
        
    End With
        
    RSTransDetails.AddNew
    RSTransDetails("Transaction_ID").value = Transaction_ID
    
    RSTransDetails("ColorID").value = 1
    RSTransDetails("ItemSize").value = 1
    RSTransDetails("ClassId").value = 1
    RSTransDetails("Item_ID").value = mItemNo
    RSTransDetails("UnitID").value = mUnitNo
    RSTransDetails("SHOWQTY").value = mQty
    RSTransDetails("showPrice").value = mPrice
    RSTransDetails("Vat").value = mVAt2
    RSTransDetails("AreaL").value = mAreaL
    
    If SystemOptions.PriceWithVAT = True Then
    Percetage = 0
    RSTransDetails("TypeVAT").value = 0
    
    RSTransDetails("Vatyo").value = 0
    Else
    RSTransDetails("TypeVAT").value = Percetage
    
    RSTransDetails("Vatyo").value = val(Percetage)
    End If
    RSTransDetails("Remarks").value = IIf(mRemark <> "", " " & mRemark, "")
    
    'FG.TextMatrix(Num, FG.ColIndex("Vat")) = IIf(IsNull(RsDetails("Vat")), "", (RsDetails("Vat").value))
                  
            'RSTransDetails("NoCount").value = IIf((Fg.TextMatrix(RowNum, Fg.ColIndex("NoCount")) = ""), Null, val(Fg.TextMatrix(RowNum, Fg.ColIndex("NoCount"))))
            RSTransDetails("Width").value = val(mwidtj)
            RSTransDetails("Height").value = val(mhight)
            RSTransDetails("Length").value = val(mLength)
            RSTransDetails("ItemDiscountType").value = 2
            RSTransDetails("ItemDiscount").value = val(mTotalDisc)
            
              RSTransDetails("CostPrice").value = mCost
              If mCost = 0 Then
                    If SystemOptions.TypicalProduction = False Then
          
                        RSTransDetails("CostPrice").value = ModItemCostPrice.GetCostItemPrice(mItemNo, 0, , , SystemOptions.SysMainStockCostMethod, , , XPDtbBill.value, val(Me.Text1.Text), RSTransDetails("UnitID").value, StoreId)
        
                        If RSTransDetails("CostPrice").value = 0 Then
                            RSTransDetails("CostPrice").value = ModItemCostPrice.GetCostItemPrice(mItemNo, 0, , , LastPurPriceType, , , XPDtbBill.value, val(Me.Text1.Text), RSTransDetails("UnitID").value, val(Me.DCboStoreName.BoundText))
                            
                        End If
                          
                    Else
                        RSTransDetails("CostPrice").value = 0
                    
                    End If
                End If
                  
                              '«·ÊÕœ« 
               
                Dim RsUnitData As ADODB.Recordset
                Dim LngCurItemID As Long
                Dim LngUnitID As Long
                Dim DblQty As Double
            
                LngCurItemID = val(mItemNo)
                LngUnitID = val(mUnitNo)
                DblQty = val(mQty)
    
                StrSQL = "Select * From TblItemsUnits Where ItemID=" & LngCurItemID
                StrSQL = StrSQL + " AND UnitID=" & LngUnitID
                Set RsUnitData = New ADODB.Recordset
                RsUnitData.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText
    
                If Not (RsUnitData.BOF Or RsUnitData.EOF) Then
                    RSTransDetails("QtyBySmalltUnit").value = RsUnitData("UnitFactor").value
                    RSTransDetails("Quantity").value = RSTransDetails("QtyBySmalltUnit").value * RSTransDetails("showqty").value
                    RSTransDetails("OpeningSalesQty").value = RSTransDetails("Quantity").value
                    RSTransDetails("OpeningSalesValue").value = RSTransDetails("CostPrice").value * val(txtQty1)
                    RSTransDetails("Price").value = val(IIf((mPrice = 0), 0, val(mPrice))) / RSTransDetails("QtyBySmalltUnit").value
                
                End If
    
            
                 UpdateTransactionsCost CStr(Transaction_ID)
                 RSTransDetails.update

  '  Dim i As Integer
    'Dim sql As String
    Dim rs2 As ADODB.Recordset
    Set rs2 = New ADODB.Recordset
    
    sql = "Select * from  TransactionValueAdded where 1=-1"
    rs2.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
    If val(LngCurItemID) <> 0 And SystemOptions.PriceWithVAT = False Then
        rs2.AddNew
        rs2("Transaction_ID").value = val(Transaction_ID)
        rs2("Transaction_Type").value = 21
        rs2("ItemID").value = LngCurItemID
        rs2("Vatyo").value = Percetage
        rs2("Vat").value = val(mVAt2)
        rs2("Valu").value = val(mTotal) + val(mTotalAdd)
        rs2("selectd").value = 1
    
    End If
    If SystemOptions.PriceWithVAT = False Then
    rs2.update
    End If
NextRow:
Next

NoteSerial = Notes_coding(val(BranchID), Transaction_Date)


CreateNotes NoteID, Transaction_Date, CInt(BranchID), 170, val(txtTotalWithVat2), NoteSerial, NoteSerial1, "Transactions", "Transaction_ID", Transaction_ID, " »‰«¡« ⁄·Ï ”‰œ  Ã„Ì⁄ —ﬁ„ " & TxtTransSerial, ToHijriDate(Transaction_Date)
txtNoteid3 = NoteID

'***********************
         StrSQL = "UPDATE TblDefComItem SET  Noteid3=" & val(txtNoteid3) & " , TransactionID3=" & val(TXTTransactionID3) & ",  NoteSerial13='" & TxtNoteSerial13 & "' WHERE ID  =" & val(TxtTransSerial)
         Cn.Execute StrSQL
'***********************
        Dim cnt As Double
        Dim usedaccount As Integer
        Dim ItemsGoodsTotalsnew As Variant
        cnt = val(txtQty1)
        PG IIf(IsNull(RSTransDetails("quantity").value), 0, RSTransDetails("quantity").value), cnt, usedaccount, ItemsGoodsTotalsnew
       
 
        'StrSQL = "UPDATE Transactions SET NOTS=" & Transaction_ID & " WHERE Transaction_ID=" & val(Me.XPTxtBillID.text)
        'Cn.Execute StrSQL
  
    If SystemOptions.UserInterface = ArabicInterface Then
        MsgBox " „ «‰‘«¡ ›« Ê—… „»Ì⁄« "
    Else
        MsgBox "Sales Invoice created"
    End If
  
'******************************************************issueVoucher








     
 
    '
 
ErrTrap:



 

End Sub



Sub PG(Optional Qty As Double, Optional cnt As Double, Optional usedaccount As Integer, Optional ItemsGoodsTotalsnew As Variant, Optional ItemsServiceTotalsnew As Variant)
Dim i As Integer
    Dim LngDevID As Long
    Dim LngDevNO  As Integer
    Dim StrTempAccountCode As String
    Dim StrTempDes As String
    Dim Account_Code_dynamic As String
    Dim SngTemp As Variant
    Dim TotalValue As Double
    On Error GoTo ErrTrap
    Dim TepAccount As String
    Dim OtherInformation As New ClsGLOther
    Dim general_noteid As Long
    Dim mBoxID As Long
    LngDevID = new_id("DOUBLE_ENTREY_VOUCHERS", "Double_Entry_Vouchers_ID", "")
    '----------------
    general_noteid = val(txtNoteid3)
    
    
    'SngTemp = NewGrid.GetItemsCostTotal * Qty / cnt


    Dim bankCommAccount As String
    Dim commision As Variant
   
    Dim Commisionvalue As Single
    Dim BankID As Long
    BankID = 0 ' GetPaymentTypeBank(val(Me.DCPaymentNet.BoundText))
    ' totalvalue = Val(Me.XPTxtValue(0).text) * Val(txt_Currency_rate.text)
   
    TotalValue = TotalValue + val(Me.TxtVAt22.Text)
    
    TotalValue = val(txtTotalWithVat2) '- val(txtTotalDisc)
   'TotalValue = Format((TotalValue), "#,###." & String(Abs(SystemOptions.Count_ACCOUNT_digit), "0"))

TotalValue = TotalValue
   Dim AdvancedAccount As String
   If SystemOptions.CustomerhavethreeAccounts = True Then
   AdvancedAccount = ModAccounts.GetMyAccountCode("TblCustemers", "CusID", val(Me.DBCboClientName.BoundText), "Account_code2")
   Else
   AdvancedAccount = ModAccounts.GetMyAccountCode("TblCustemers", "CusID", val(Me.DBCboClientName.BoundText), "Account_code")
   End If
   If AdvancedAccount = "" Then txtAdvPay = 0
   TepAccount = AdvancedAccount
  OtherInformation.NextAccount_Code = get_account_code_branch(2, val(dcBranch.BoundText))
  'OtherInformation.NextAccount_Code = get_account_code_branch(149, VAL(Dcbranch.BoundText ))
   Dim DebitAccountTemp As String
       'Dim AdvancedAccount As String
   If SystemOptions.CustomerhavethreeAccounts = True Then
        AdvancedAccount = ModAccounts.GetMyAccountCode("TblCustemers", "CusID", val(Me.DBCboClientName.BoundText), "Account_code2")
   Else
        AdvancedAccount = ModAccounts.GetMyAccountCode("TblCustemers", "CusID", val(Me.DBCboClientName.BoundText), "Account_code")
   End If
   
    If Me.CboPayMentType.ListIndex = 0 Then 'cash
            mBoxID = val(DcboBox.BoundText)

          '  mBoxID = 2
            StrTempAccountCode = GetMyAccountCode("TblBoxesData", "BoxID", mBoxID)   '«·„»Ì⁄« 
     Else
            StrTempAccountCode = AdvancedAccount
            StrTempAccountCode = GetMyAccountCode("TblCustemers", "CusID", val(Me.DBCboClientName.BoundText))
        
     End If
        Dim maxvalue As Double
       
    
        If SystemOptions.UserInterface = ArabicInterface Then
            StrTempDes = "›« Ê—… »Ì⁄ —ﬁ„ " & Me.TxtNoteSerial13.Text & " »‰«¡« ⁄·Ï ”‰œ  Ã„Ì⁄ —ﬁ„ " & TxtTransSerial
        Else
            StrTempDes = "Sales Invoice NO: " & Me.TxtNoteSerial13.Text & " »‰«¡« ⁄·Ï ”‰œ  Ã„Ì⁄ —ﬁ„ " & TxtTransSerial
        End If

        LngDevNO = LngDevNO + 1
    Dim ValuGird As Double
   Dim StrMSG As String
   OtherInformation.NextAccount_Code = get_account_code_branch(2, val(dcBranch.BoundText))
       'If val(CboPayMentType.ListIndex) = 0 Then
        If ModAccounts.AddNewDev(LngDevID, LngDevNO, StrTempAccountCode, TotalValue - val(txtAdvPay), 0, StrTempDes, general_noteid, , , , Me.XPDtbBill.value, Me.DCboUserName.BoundText, val(Transaction_ID), , , , , , , , , , , , , , , , , val(Me.dcBranch.BoundText), , , , , , , , , , , , , , , , , , , , , , , , , , , OtherInformation) = False Then
            GoTo ErrTrap
        End If
        TepAccount = StrTempAccountCode
        DebitAccountTemp = StrTempAccountCode
            LngDevNO = LngDevNO + 1
            
            
     

       'End If
        DebitAccountTemp = StrTempAccountCode
  






    '«·œ«∆‰ ›Ì Õ«·… «·«’‰«›

    '  ÕœÌœ ÿ—Ìﬁ… —»ÿ «·„Œ«“‰ Ê «·Õ”«»«  ÊÂÌ ⁄·Ï „” ÊÏ «·›—⁄ Ê —»ÿ ⁄·Ï „” ÊÏ «·„Ã„Ê⁄«  Ê«·›—⁄ «Ê «·„Ã„Ê⁄«  Ê «·„Œ«“‰

    '1 work with branch
    '2 work with inventory
    '3 work with groups
    SngTemp = val(txtTotalAdd2) + val(txtTotal2) - val(txtTotalDisc2)

    SngTemp = Round(SngTemp, SystemOptions.Count_ACCOUNT_digit)
'    TotalValue = Format((TotalValue), "#,###." & String(Abs(SystemOptions.Count_ACCOUNT_digit), "0"))
If SystemOptions.PriceWithVAT = True Then
SngTemp = SngTemp / 1.05
End If
    If SngTemp > 0 Then
        If detect_inventory_work_type = 1 Or detect_inventory_work_type = 2 Then
            Account_Code_dynamic = get_account_code_branch(2, val(dcBranch.BoundText))
        
            If Account_Code_dynamic = "NO branch" Then
                If SystemOptions.UserInterface = ArabicInterface Then
                    MsgBox "·„ Ì „ «‰‘«¡ «·›—⁄", vbCritical
                Else
                    MsgBox "Branch Not Created", vbCritical
                End If

                GoTo ErrTrap
            Else

                If Account_Code_dynamic = "NO account" Then
                    If SystemOptions.UserInterface = ArabicInterface Then
                        MsgBox "·„ Ì „  ÕœÌœ Õ”«»  «·„»Ì⁄«   ›Ì «·›—⁄ ·Â–… «·⁄„·Ì…", vbCritical
                    Else
                        MsgBox "Sales Account Not Defined in this Branch", vbCritical
                    End If

                    GoTo ErrTrap
         
                End If
            End If

    
                StrTempAccountCode = Account_Code_dynamic '«·„»Ì⁄« 
   

OtherInformation.NextAccount_Code = TepAccount
            '           StrTempAccountCode = Account_Code_dynamic '«·„»Ì⁄« 
            'StrTempAccountCode = "a4a1" '«·„»Ì⁄« 
            If SystemOptions.UserInterface = ArabicInterface Then
                StrTempDes = "›« Ê—… »Ì⁄ —ﬁ„ " & Me.TxtNoteSerial13.Text & " »‰«¡« ⁄·Ï ”‰œ  Ã„Ì⁄ —ﬁ„ " & TxtTransSerial
            Else
                StrTempDes = "Sales Invoice NO: " & Me.TxtNoteSerial13.Text & " »‰«¡« ⁄·Ï ”‰œ  Ã„Ì⁄ —ﬁ„ " & TxtTransSerial
            End If

            LngDevNO = LngDevNO + 1

            If ModAccounts.AddNewDev(LngDevID, LngDevNO, StrTempAccountCode, SngTemp, 1, StrTempDes, general_noteid, , , , Me.XPDtbBill.value, Me.DCboUserName.BoundText, val(Transaction_ID), , , , , , , , , , , , , , , , , val(Me.dcBranch.BoundText), , , , , , , , , , , , , , , , , , , , , , , , , , , OtherInformation) = False Then
                GoTo ErrTrap
            End If
            
            
  Dim value As Double
'  value = val(Me.txtTotalDisc)
'  If value > 0 Then
'        Account_Code_dynamic = get_account_code_branch(12, VAL(Dcbranch.BoundText ))
'
'        If Account_Code_dynamic = "NO branch" Then
'            If SystemOptions.UserInterface = ArabicInterface Then
'                MsgBox "·„ Ì „ «‰‘«¡ «·›—⁄", vbCritical
'            Else
'                MsgBox "Branch Not Created ", vbCritical
'            End If
'
'            GoTo ErrTrap
'        Else
'
'            If Account_Code_dynamic = "NO account" Then
'                If SystemOptions.UserInterface = ArabicInterface Then
'                    MsgBox "·„ Ì „  ÕœÌœ Õ”«»    «·Œ’„ «·„”„ÊÕ »Â   ›Ì «·›—⁄ ·Â–… «·⁄„·Ì…", vbCritical
'                Else
'                    MsgBox "Allowance Discount Not Deined in this Branch", vbCritical
'                End If
'
'                GoTo ErrTrap
'
'            End If
'        End If
'
'
'        If val(Me.txtTotalDisc) > 0 Then
'         StrTempAccountCode = Account_Code_dynamic
'                If SystemOptions.DiscountSalesCreateVchr = True Then
'                 LngDevNO = LngDevNO + 1
'                       '     If ModAccounts.AddNewDev(LngDevID, LngDevNO, StrTempAccountCode, val(Me.LblDiscountsTotal.Caption), 0, StrTempDes, , , , , Me.XPDtbBill.value, Me.DCboUserName.BoundText, val(Me.XPTxtBillID.text)) = False Then
'                                            If ModAccounts.AddNewDev(LngDevID, LngDevNO, StrTempAccountCode, value, 0, StrTempDes, general_noteid, , , , Me.XPDtbBill.value, Me.DCboUserName.BoundText, val(Transaction_ID), , , , , , , , , , , , , , , , , val(Me.Dcbranch.BoundText), , , , , , , , , , , , , , , , , , , , , , , , , , , OtherInformation) = False Then
'
'                                                GoTo ErrTrap
'                                            End If
'
'                                End If
'
'                End If
'
   ' End If



'Õ”«» «·«÷«›« 

    
        ElseIf detect_inventory_work_type = 3 Then
'
        End If

    End If
   



    '
    
If SystemOptions.PriceWithVAT = True Then
TxtVAt22.Text = (TotalValue / 1.05) * 0.05
End If
        If val(TxtVAt22.Text) > 0 Then
    Dim AccountVATCreit As String
 GetValueAddedAccount XPDtbBill.value, , AccountVATCreit, 1, 21


         If SystemOptions.UserInterface = ArabicInterface Then
                                StrTempDes = "  ﬁÌ„… „÷«›… »‰”»… " & PercetageVat & " %  " & "›« Ê—… »Ì⁄ —ﬁ„ " & Me.TxtNoteSerial13.Text & " »‰«¡« ⁄·Ï ”‰œ  Ã„Ì⁄ —ﬁ„ " & TxtTransSerial
                            Else
                                StrTempDes = "VAT Sales Invoice NO: " & Me.TxtNoteSerial13.Text
        End If
            
                            LngDevNO = LngDevNO + 1
        If ModAccounts.AddNewDev(LngDevID, LngDevNO, AccountVATCreit, val(TxtVAt22.Text), 1, StrTempDes, general_noteid, , , , Me.XPDtbBill.value, Me.DCboUserName.BoundText, val(Transaction_ID), , , , , , , , , , , , , , , , , val(Me.dcBranch.BoundText), , , , , , , , , , , , , , , , , , , , , , , , , , , OtherInformation) = False Then
            GoTo ErrTrap
        End If
        TxtVAt22.Text = 0
     End If
     ''/////////////
     Dim Account_Code_dynamic82 As String

     ''//////////
'     If SystemOptions.DealingWithPrepayAccount = True Then
'      If val(TxtVAt2.Text) > 0 Then
'
'             GetValueAddedAccount XPDtbBill.value, , AccountVATCreit, 1, 21
'         If SystemOptions.UserInterface = ArabicInterface Then
'                                StrTempDes = "  ﬁÌ„… „÷«›… " & "›« Ê—… »Ì⁄ —ﬁ„ " & Me.TxtNoteSerial13.Text
'                            Else
'                                StrTempDes = "VAT ""Sales Invoice NO: " & Me.TxtNoteSerial13.Text
'        End If
'
'                            LngDevNO = LngDevNO + 1
'        If ModAccounts.AddNewDev(LngDevID, LngDevNO, AccountVATCreit, val(TxtVAt2.Text), 0, StrTempDes, general_noteid, , , , Me.XPDtbBill.value, Me.DCboUserName.BoundText, val(Transaction_ID), , , , , , , , , , , , , , , , , val(Me.Dcbranch.BoundText), , , , , , , , , , , , , , , , , , , , , , , , , , , OtherInformation) = False Then
'            GoTo ErrTrap
'        End If
'                 If SystemOptions.UserInterface = ArabicInterface Then
'                                StrTempDes = "  Õ”«» «·⁄„Ì· " & "›« Ê—… »Ì⁄ —ﬁ„ " & Me.TxtNoteSerial13.Text
'                            Else
'                                StrTempDes = "Customer ""Sales Invoice NO: " & Me.TxtNoteSerial13.Text
'                 End If
'                  LngDevNO = LngDevNO + 1
'        AccountVATCreit = GetMyAccountCode("TblCustemers", "CusID", val(Me.DBCboClientName.BoundText))
'             If ModAccounts.AddNewDev(LngDevID, LngDevNO, AccountVATCreit, val(TxtVAt2.Text), 1, StrTempDes, general_noteid, , , , Me.XPDtbBill.value, Me.DCboUserName.BoundText, val(Transaction_ID), , , , , , , , , , , , , , , , , val(Me.Dcbranch.BoundText), , , , , , , , , , , , , , , , , , , , , , , , , , , OtherInformation) = False Then
'            GoTo ErrTrap
'        End If
'     End If
'     End If
   
xl:

'************************************************************************************
cmdTransfer.Enabled = False
cmdCancel.Enabled = True

ErrTrap:
End Sub


Public Sub loadgrid(ByVal Sqlstmt As String, _
                          ByRef tGrd As Control, _
                          Optional ResetRows As Boolean = True, _
                          Optional InsertRow As Boolean = False, _
                          Optional mReCreateColumns As Boolean = False)
    Dim tRs As New ADODB.Recordset
  
    Dim sCur  As Long
    Dim mWithMyFormat As Boolean
    tRs.Open Sqlstmt, Cn, adOpenStatic, adLockReadOnly, adCmdText
    Dim i As Long
    ' ******************************************
    If ResetRows Then tGrd.Rows = tGrd.FixedRows
    ' ******************************************
    If mReCreateColumns Then
        tGrd.Cols = 1
        tGrd.Cols = tRs.Fields.count + 1
        For i = 1 To tGrd.Cols - 1
            tGrd.ColKey(i) = tRs.Fields.Item(i - 1).Name
            tGrd.TextMatrix(0, i) = tRs.Fields.Item(i - 1).Name
        Next
    End If
    ' ******************************************
    ' ******************************************
    tGrd.Redraw = flexRDNone
    ' ******************************************
    
    Dim j As Long
    i = tGrd.Rows
    sCur = 0
    Do While Not tRs.EOF
        tGrd.AddItem i - tGrd.FixedRows + 1
        For j = 0 To tRs.Fields.count - 1
            If tGrd.ColIndex(tRs.Fields.Item(j).Name) <> -1 Then
                If tRs.Fields.Item(j).Type = adCurrency And mWithMyFormat Then
                    tGrd.TextMatrix(i, tGrd.ColIndex(tRs.Fields.Item(j).Name)) = (val(tRs.Fields.Item(j).value & ""))
                Else
                    tGrd.TextMatrix(i, tGrd.ColIndex(tRs.Fields.Item(j).Name)) = Trim(tRs.Fields.Item(j).value & "")
                End If
            End If
        Next
        i = i + 1
        sCur = sCur + 1

        tRs.MoveNext
    Loop
    tRs.Close
    Set tRs = Nothing

    If InsertRow Then tGrd.AddItem tGrd.Rows - tGrd.FixedRows + 1
    tGrd.Redraw = flexRDDirect
End Sub


Public Sub saveGrid(ByVal Sqlstmt As String, ByRef tGrd As vsFlexGrid, ByVal ChekPoint As String, ByVal Index As String, ParamArray FieldValue())
    On Error GoTo Err
    Dim tRs As New ADODB.Recordset
    Dim i As Long
    Dim k As Long
    Dim j As Long
    tRs.Open Sqlstmt, Cn, adOpenKeyset, adLockOptimistic, adCmdText
    ' *******************************************
    Dim II As Integer
    II = 0
    For i = tGrd.FixedRows To tGrd.Rows - 1
        If ChekPoint <> "" Then
            If Trim(tGrd.TextMatrix(i, tGrd.ColIndex(ChekPoint))) = "" Then GoTo NextStep
        End If
        '**********************
        tRs.AddNew
        II = II + 1
        If Index <> "" Then tRs(Index) = II
        For k = 0 To UBound(FieldValue) Step 2
            tRs.Fields.Item(FieldValue(k)).value = FieldValue(k + 1)
            'Debug.Print FieldValue(k) & " " & tRs.Fields.Item(FieldValue(k)).Value
        Next
        '*************************
        'Debug.Print "fields count " & tRs.Fields.count
        For j = 0 To tRs.Fields.count - 1

            If tGrd.ColIndex(tRs.Fields.Item(j).Name) <> -1 Then
                If tRs.Fields.Item(j).Type = adInteger Or tRs.Fields.Item(j).Type = adCurrency Or tRs.Fields.Item(j).Type = adBoolean Or tRs.Fields.Item(j).Type = adSmallInt Or tRs.Fields.Item(j).Type = adBigInt Or tRs.Fields.Item(j).Type = adTinyInt Or tRs.Fields.Item(j).Type = adUnsignedTinyInt Or tRs.Fields.Item(j).Type = adNumeric Or tRs.Fields.Item(j).Type = adDouble Or tRs.Fields.Item(j).Type = adDecimal Then
                    If tRs.Fields.Item(j).Type = adBoolean Then
                        tRs.Fields.Item(j).value = (UCase(tGrd.ValueMatrix(i, tGrd.ColIndex(tRs.Fields.Item(j).Name))) = "TRUE") Or (UCase(tGrd.TextMatrix(i, tGrd.ColIndex(tRs.Fields.Item(j).Name))) = "-1") Or (val(tGrd.ValueMatrix(i, tGrd.ColIndex(tRs.Fields.Item(j).Name))) = -1)
                    Else
'                        If tGrd.ColComboList(tGrd.ColIndex(tRS.Fields.Item(j).Name)) <> "" Then
'                            tRS.Fields.Item(j).Value = tGrd.ValueMatrix(i, tGrd.ColIndex(tRS.Fields.Item(j).Name))
'                        Else
                            'If Index <> "" And UCase(tRs.Fields.Item(j).Name) <> UCase(tRs(Index).Name) Then
                            tRs.Fields.Item(j).value = val(tGrd.TextMatrix(i, tGrd.ColIndex(tRs.Fields.Item(j).Name)))
                            'End If
'                        End If
                    End If
                Else
                    If tRs.Fields.Item(j).Type = adDBTimeStamp Or tRs.Fields.Item(j).Type = adDBTime Or tRs.Fields.Item(j).Type = adDBDate Then
                        If Not IsDate(tGrd.TextMatrix(i, tGrd.ColIndex(tRs.Fields.Item(j).Name))) Then
                            tRs.Fields.Item(j).value = Null
                        Else
                            tRs.Fields.Item(j).value = tGrd.TextMatrix(i, tGrd.ColIndex(tRs.Fields.Item(j).Name))
                        End If
                    Else
                        'If Index <> "" And UCase(tRs.Fields.Item(j).Name) <> UCase(tRs(Index).Name) Then
                        tRs.Fields.Item(j).value = Trim(tGrd.TextMatrix(i, tGrd.ColIndex(tRs.Fields.Item(j).Name) & ""))
                        'End If
                    End If
                End If
            End If
            'Debug.Print tRs.Fields.Item(j).Name & " = " & tRs.Fields.Item(j).Value
        Next
tRs.update
NextStep:
    Next
    tRs.Close
    Exit Sub
Err:
    If Err.Number = -2147217887 Then        ' one item is empty
        Resume Next
    End If
    '    Resume Next
End Sub




Public Sub GetCustomerNamebyPhone(Optional ByVal phone As String = "", Optional ByVal Name As String = "", Optional ByVal CUSTID As String = "", Optional ByVal SearchCode As String = "")
            If phone = "" And Name = "" And CUSTID = "" And SearchCode = "" Then Exit Sub
    Dim sql As String
    Dim rs As New ADODB.Recordset
    Dim Balance As Double

        If phone <> "" Then
            sql = "SELECT     Cus_mobile , CusName,CusID,Fullcode,cPaymentType,EmpId From dbo.TblCustemers  WHERE     (Cus_mobile = '" & phone & "')"
        ElseIf Name <> "" Then
            sql = "SELECT     Cus_mobile, CusName,CusID,Fullcode,cPaymentType,EmpId From dbo.TblCustemers  WHERE     (CusName = '" & Name & "')"
        ElseIf CUSTID <> "" Then
            sql = "SELECT     Cus_mobile, CusName,CusID,Fullcode,cPaymentType,EmpId From dbo.TblCustemers  WHERE     (CusID = " & val(CUSTID) & ")"
        ElseIf SearchCode <> "" Then
            sql = "SELECT     Cus_mobile, CusName,CusID,Fullcode,cPaymentType,EmpId From dbo.TblCustemers  WHERE     Fullcode ='" & SearchCode & "'"
        Else
        Exit Sub
        End If
  
    rs.Open sql, Cn, adOpenStatic, adLockReadOnly, adCmdText

    If rs.RecordCount > 0 Then

        TxtPhone = rs!Cus_mobile & ""
        TxtSearchCode2.Text = rs!Fullcode & ""
        TxtSearchCode.Text = rs!Fullcode & ""
        DBCboClientName.BoundText = val(rs!CusID & "")
        DcboEmp.BoundText = val(rs!EmpID & "")
        txtCustomerName.Text = IIf(IsNull(rs!CusName), "", rs!CusName)
        If SystemOptions.DontShowMoreDetailsCompItem Then
            CboPayMentType.ListIndex = IIf(IsNull(rs("cPaymentType").value), 0, rs("cPaymentType").value)
        End If
    Else
         TxtPhone = ""
         TxtSearchCode = ""
         DBCboClientName.BoundText = ""
          txtCustomerName.Text = ""
              If Me.TxtModFlg <> "R" Then

        If SystemOptions.UserInterface = ArabicInterface Then
            MsgBox "Â–« «·⁄„Ì· €Ì— „ÊÃÊœ", vbCritical
        Else
            MsgBox "This client does not exist", vbCritical
        End If
End If
    End If

    rs.Close

End Sub
 Private Sub CalcCostPercent()
    Dim i As Long
    Dim mCostPercent As Double
    Dim mCostTotal As Double
    mCostTotal = FG2.Aggregate(flexSTSum, FG2.FixedRows, FG2.ColIndex("Cost"), FG2.Rows - 1, FG2.ColIndex("Cost"))
    If mCostTotal <> 0 Then
        For i = 1 To FG2.Rows - 1
            FG2.TextMatrix(i, FG2.ColIndex("PercentCost")) = val(FG2.TextMatrix(i, FG2.ColIndex("Cost"))) / mCostTotal * 100
        Next
    End If
 End Sub



Public Function DeleteTransactiomsVoucher2(Transaction_ID As Double)
Dim StrSQL  As String
Dim StrSqlDel  As String
If Transaction_ID = 0 Then Exit Function
        
    StrSQL = "Delete From DOUBLE_ENTREY_VOUCHERS Where Notes_ID = (SELECT NoteID FROM Transactions Where Transaction_ID=" & Transaction_ID & ")"
    Cn.Execute StrSQL, , adExecuteNoRecords
    
    
    StrSqlDel = "delete From Notes where NoteID= (SELECT NoteID FROM Transactions Where Transaction_ID=" & Transaction_ID & ")"
    Cn.Execute StrSqlDel, , adExecuteNoRecords
    
    StrSqlDel = "delete From Transaction_Details  where Transaction_ID=" & Transaction_ID
    Cn.Execute StrSqlDel, , adExecuteNoRecords
        
    StrSqlDel = "delete From Transactions  where Transaction_ID=" & Transaction_ID
    Cn.Execute StrSqlDel, , adExecuteNoRecords
        
    
    
    
   
               
     
End Function




Public Sub RetriveOrder(Optional order_no As String = "", _
                        Optional Transaction_Type As Integer = 0)
    Dim RsDetails As New ADODB.Recordset
    Dim StrSQL As String
    Dim RsNotes As New ADODB.Recordset
    Dim RsTemp As ADODB.Recordset
    Dim rs As ADODB.Recordset
    Dim Dcombos As ClsDataCombos
    Set Dcombos = New ClsDataCombos
    Dim Num As Long
    Dim StoreId2 As Double
    Dim issuedQty As Double
    Dim rsDummy As New ADODB.Recordset
    Dim mCostPrice As Double
    Dim s As String
    Dim mTransType As Long
   On Error GoTo ErrTrap
    FG2.Clear flexClearScrollable, flexClearEverything
    FG2.Rows = 2
    FG.Clear flexClearScrollable, flexClearEverything
    FG.Rows = 2
    FG.Refresh
    Dcombos.GetItemsNames DcboItemID1
   
        StrSQL = "Select * from transactions where  Transaction_Type=" & Transaction_Type & " and NoteSerial1='" & order_no & "' "
      
 

    Set rs = New ADODB.Recordset
    rs.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

    If rs.RecordCount < 1 Then
 
        Exit Sub
    Else
        DBCboClientName.BoundText = IIf(IsNull(rs("CusID").value), "", rs("CusID").value)
       ' Me.BoundText = IIf(IsNull(rs("Currency_id").value), "", rs("Currency_id").value)
        
        Me.DCboStoreName.BoundText = IIf(IsNull(rs("storeid").value), "", rs("storeid").value)
        DCboStore2Name.BoundText = IIf(IsNull(rs("storeid").value), "", rs("storeid").value)
        
        
        Me.dcBranch.BoundText = IIf(IsNull(rs("Branchid").value), "", rs("Branchid").value)


        
        
        'txt_Currency_rate.text = IIf(IsNull(rs("Currency_rate").value), 1, (rs("Currency_rate").value))
    End If

    If rs.EOF Or rs.BOF Then
        Exit Sub
    End If
    txtOrderID = rs!Transaction_ID & ""


    Screen.MousePointer = vbArrowHourglass
    If SystemOptions.InsertItemManualOut Then Exit Sub
    
 '   StrSQL = "Select * from transactions where  Transaction_Type=" & 26 & " and order_no='" & order_no & "' and CBoBasedON = 1"
 '   StrSQL = StrSQL & " Union all"
 '   StrSQL = StrSQL & " Select * from transactions where  Transaction_Type=" & 10 & " and order_no='" & order_no & "' and BillBasedOn = 4"
 '   StrSQL = StrSQL & " Union all"
    StrSQL = StrSQL & " Select * from transactions where  Transaction_Type=" & 6 & " and NoteSerial1='" & order_no & "' and BillBasedOn = 4"
    
     XPTxtSum.Text = ""
    Set rs = New ADODB.Recordset
    rs.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText
    Do While Not rs.EOF
        
        mTransType = val(rs!Transaction_Type & "")
        mTransactionID4 = 0
        mNoteSerial14 = ""
        mTransactionID5 = 0
        mNoteSerial15 = ""

        If mTransType = 10 Then
            mTransactionID5 = rs!Transaction_ID & ""
            mNoteSerial15 = rs!NoteSerial1 & ""
        Else
            mTransactionID4 = rs!Transaction_ID & ""
            mNoteSerial14 = rs!NoteSerial1 & ""
        End If
       
        StrSQL = "SELECT TblItems.HaveSerial, * FROM TblItems INNER JOIN Transaction_Details " & "ON TblItems.ItemID = Transaction_Details.Item_ID INNER JOIN dbo.TblUnites ON dbo.Transaction_Details.UnitID = dbo.TblUnites.UnitID"
        StrSQL = StrSQL + " where Transaction_ID=" & val(rs("Transaction_ID").value)
    
    
        StrSQL = StrSQL & " order by Transaction_Details.id "
        Set RsDetails = New ADODB.Recordset
        
        RsDetails.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText
       
    
        If Not (RsDetails.EOF Or RsDetails.BOF) Then
            FG2.Rows = FG2.Rows + RsDetails.RecordCount + 1
    
            For Num = 1 To RsDetails.RecordCount
                DcboItemID1.BoundText = IIf(IsNull(RsDetails("Item_ID")), "", (RsDetails("Item_ID").value))
                txtQty1 = IIf(IsNull(RsDetails("showqty")), "", (RsDetails("showqty").value))
             
                DcbUnit.BoundText = IIf(IsNull(RsDetails("UnitID")), "", (RsDetails("UnitID").value))
    '            If Transaction_Type = 42 Then
    '                s = "SELECT T2.* "
    '                s = s & " from  Transactions AS t "
    '                s = s & " Inner Join Transaction_Details T2 On T2.Transaction_ID = t.Transaction_ID"
    '                s = s & " WHERE t.Transaction_Type = 26 and t.OrderID =  " & val(txtOrderID)
    '                s = s & " and  T2.Item_ID = " & val(RsDetails("Item_ID").value & "")
    '                s = s & " and T2.UnitId= " & val(RsDetails("UnitId").value & "")
    '                Set rsDummy = New ADODB.Recordset
    '
    '                rsDummy.Open s, Cn, adOpenStatic, adLockReadOnly, adCmdText
    '                If rsDummy.EOF Then
    '                    mCostPrice = 0
    '                Else
    '                    mCostPrice = val(rsDummy!ShowPrice & "")
    '                End If
    '
    '            End If
    '
    '            If mCostPrice <> 0 Then
    '                txtPrice = mCostPrice
    '            Else
    '                txtPrice = ModItemCostPrice.GetCostItemPrice(DcboItemID1.BoundText, 0, , , SystemOptions.SysMainStockCostMethod, , , XPDtbBill.value, val(Me.XPTxtBillID), val(FG.Cell(flexcpData, Num, FG.ColIndex("UnitID"))), val(Me.DCboStoreName.BoundText))
    '            End If
    '            FG.TextMatrix(Num, FG.ColIndex("SalesPrice")) = GetItemPrice(val(FG.TextMatrix(Num, FG.ColIndex("Code"))), 0, val(FG.Cell(flexcpData, Num, FG.ColIndex("UnitID"))))
    '            FG.TextMatrix(Num, FG.ColIndex("TotalSalesPrice")) = val(FG.TextMatrix(Num, FG.ColIndex("SalesPrice"))) * val(FG.TextMatrix(Num, FG.ColIndex("Count")))
                DcboItemID1_Validate False
                'cmdAdd_Click
                cmdAdd__Click
               ' cmdAdd_Click
                RsDetails.MoveNext
                Debug.Print Num
    
    
            Next Num
    
        End If
        rs.MoveNext
    Loop
    TxtFillData.Text = "F"
    Screen.MousePointer = vbDefault
'    XPTxtCurrent.Caption = rs.AbsolutePosition
'    XPTxtCount.Caption = rs.RecordCount
Dcombos.GetItemsNames DcboItemID1, -1, -1, 1
    Exit Sub
ErrTrap:
    Screen.MousePointer = vbDefault
End Sub






 

 
 

