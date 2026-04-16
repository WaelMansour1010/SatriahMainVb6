VERSION 5.00
Object = "{0AFE7BE0-11B7-4A3E-978D-D4501E9A57FE}#1.0#0"; "c1sizer.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{FE5DCFAD-BC1D-11D2-94CF-004005455FAA}#1.4#0"; "ImpulseButton.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Object = "{49003D3A-66CD-11D7-A449-E937BE2D9041}#1.0#0"; "ALLBUTTONS.ocx"
Begin VB.Form FrmOptions 
   BackColor       =   &H00E2E9E9&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "«⁄œ«œ«  «·‰Ÿ«„"
   ClientHeight    =   9750
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   12630
   Icon            =   "FrmOptions.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   RightToLeft     =   -1  'True
   ScaleHeight     =   9750
   ScaleWidth      =   12630
   ShowInTaskbar   =   0   'False
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   8
      Charset         =   178
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Begin ALLButtonS.ALLButton ALLButton8 
      Height          =   255
      Left            =   13200
      TabIndex        =   9
      Top             =   8400
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   450
      BTYPE           =   8
      TX              =   "«··«∆Õ… «·œ«Œ·Ì…"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   2
      FOCUSR          =   -1  'True
      BCOL            =   12648447
      BCOLO           =   12648447
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   12648447
      MPTR            =   1
      MICON           =   "FrmOptions.frx":038A
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin ALLButtonS.ALLButton ALLButton3 
      Height          =   255
      Left            =   13200
      TabIndex        =   4
      Top             =   7560
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   450
      BTYPE           =   8
      TX              =   " ﬂÊÌœ «·„” ‰œ« "
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   2
      FOCUSR          =   -1  'True
      BCOL            =   12648447
      BCOLO           =   12648447
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   12648447
      MPTR            =   1
      MICON           =   "FrmOptions.frx":03A6
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin ALLButtonS.ALLButton ALLButton4 
      Height          =   255
      Left            =   13200
      TabIndex        =   5
      Top             =   7935
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   450
      BTYPE           =   8
      TX              =   " ﬂÊÌœ «·ÕﬁÊ·"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   2
      FOCUSR          =   -1  'True
      BCOL            =   12648447
      BCOLO           =   12648447
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   12648447
      MPTR            =   1
      MICON           =   "FrmOptions.frx":03C2
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin ALLButtonS.ALLButton ALLButton2 
      Height          =   255
      Left            =   13200
      TabIndex        =   255
      Top             =   6840
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   450
      BTYPE           =   8
      TX              =   "  ﬂÊÌœ «·Õ”«»« "
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   2
      FOCUSR          =   -1  'True
      BCOL            =   12648447
      BCOLO           =   12648447
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   12648447
      MPTR            =   1
      MICON           =   "FrmOptions.frx":03DE
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin ALLButtonS.ALLButton ALLButton7 
      Height          =   255
      Left            =   12720
      TabIndex        =   8
      Top             =   840
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   450
      BTYPE           =   8
      TX              =   "«·«‰‘ÿ… Ê «·«›—⁄"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   2
      FOCUSR          =   -1  'True
      BCOL            =   12648447
      BCOLO           =   12648447
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   12648447
      MPTR            =   1
      MICON           =   "FrmOptions.frx":03FA
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin ALLButtonS.ALLButton ALLButton5 
      Height          =   255
      Left            =   13200
      TabIndex        =   6
      Top             =   7215
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   450
      BTYPE           =   8
      TX              =   "«‰Ê«⁄ «·„” ‰œ« "
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   2
      FOCUSR          =   -1  'True
      BCOL            =   12648447
      BCOLO           =   12648447
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   12648447
      MPTR            =   1
      MICON           =   "FrmOptions.frx":0416
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin ALLButtonS.ALLButton ALLButton6 
      Height          =   255
      Left            =   12720
      TabIndex        =   7
      Top             =   1200
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   450
      BTYPE           =   8
      TX              =   "—»ÿ «·Õ”«»« "
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   2
      FOCUSR          =   -1  'True
      BCOL            =   12648447
      BCOLO           =   12648447
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   12648447
      MPTR            =   1
      MICON           =   "FrmOptions.frx":0432
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin ALLButtonS.ALLButton ALLButton1 
      Height          =   255
      Left            =   12840
      TabIndex        =   2
      Top             =   480
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   450
      BTYPE           =   8
      TX              =   " «·› —«  «·„Õ«”»Ì…"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   2
      FOCUSR          =   -1  'True
      BCOL            =   12648447
      BCOLO           =   12648447
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   12648447
      MPTR            =   1
      MICON           =   "FrmOptions.frx":044E
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin ImpulseButton.ISButton CmdCancel 
      Height          =   255
      Left            =   180
      TabIndex        =   0
      Top             =   9480
      Width           =   825
      _ExtentX        =   1455
      _ExtentY        =   450
      ButtonStyle     =   1
      ButtonPositionImage=   1
      Caption         =   "≈·€«¡"
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
      ButtonImage     =   "FrmOptions.frx":046A
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
   Begin ImpulseButton.ISButton CmdOk 
      Height          =   255
      Left            =   960
      TabIndex        =   1
      Top             =   9480
      Width           =   825
      _ExtentX        =   1455
      _ExtentY        =   450
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
      BackStyle       =   0
      ButtonImage     =   "FrmOptions.frx":0804
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
   Begin C1SizerLibCtl.C1Tab tb 
      Height          =   9465
      Left            =   -1560
      TabIndex        =   10
      Top             =   0
      Width           =   14160
      _cx             =   24977
      _cy             =   16695
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
      Caption         =   $"FrmOptions.frx":0B9E
      Align           =   0
      CurrTab         =   0
      FirstTab        =   0
      Style           =   3
      Position        =   6
      AutoSwitch      =   -1  'True
      AutoScroll      =   -1  'True
      TabPreview      =   -1  'True
      ShowFocusRect   =   -1  'True
      TabsPerPage     =   0
      BorderWidth     =   0
      BoldCurrent     =   0   'False
      DogEars         =   -1  'True
      MultiRow        =   0   'False
      MultiRowOffset  =   200
      CaptionStyle    =   0
      TabHeight       =   0
      TabCaptionPos   =   7
      TabPicturePos   =   1
      CaptionEmpty    =   ""
      Separators      =   0   'False
      AccessibleName  =   ""
      AccessibleDescription=   ""
      AccessibleValue =   ""
      AccessibleRole  =   37
      Begin C1SizerLibCtl.C1Elastic Ele 
         Height          =   9375
         Index           =   0
         Left            =   45
         TabIndex        =   11
         TabStop         =   0   'False
         Top             =   45
         Width           =   12225
         _cx             =   21564
         _cy             =   16536
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
         Begin VB.TextBox txtActivityName 
            Alignment       =   1  'Right Justify
            Height          =   375
            Left            =   2010
            MaxLength       =   50
            RightToLeft     =   -1  'True
            TabIndex        =   630
            Top             =   3720
            Width           =   2205
         End
         Begin VB.TextBox txtNoOFDigitUser 
            Alignment       =   1  'Right Justify
            Height          =   315
            Index           =   3
            Left            =   6660
            RightToLeft     =   -1  'True
            TabIndex        =   626
            Top             =   5760
            Width           =   1365
         End
         Begin VB.TextBox txtNoOFDigitUser 
            Alignment       =   1  'Right Justify
            Height          =   315
            Index           =   5
            Left            =   8940
            MaxLength       =   4
            RightToLeft     =   -1  'True
            TabIndex        =   625
            Top             =   5760
            Width           =   1365
         End
         Begin VB.TextBox txtNoOFDigitUser 
            Alignment       =   1  'Right Justify
            Height          =   315
            Index           =   8
            Left            =   4140
            RightToLeft     =   -1  'True
            TabIndex        =   624
            Top             =   5760
            Width           =   1365
         End
         Begin VB.TextBox txtNoOFDigitUser 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00C0C0FF&
            Height          =   315
            Index           =   7
            Left            =   8940
            MaxLength       =   5
            RightToLeft     =   -1  'True
            TabIndex        =   621
            Tag             =   "5 digit at least"
            Top             =   5430
            Width           =   1365
         End
         Begin VB.TextBox txtNoOFDigitUser 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00C0C0FF&
            Height          =   315
            Index           =   10
            Left            =   6660
            MaxLength       =   2
            RightToLeft     =   -1  'True
            TabIndex        =   620
            Top             =   5430
            Width           =   1365
         End
         Begin VB.TextBox txtNoOFDigitUser 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00C0C0FF&
            Height          =   315
            Index           =   2
            Left            =   6660
            RightToLeft     =   -1  'True
            TabIndex        =   615
            Top             =   5055
            Width           =   1365
         End
         Begin VB.TextBox txtNoOFDigitUser 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00C0C0FF&
            Height          =   315
            Index           =   4
            Left            =   8940
            MaxLength       =   4
            RightToLeft     =   -1  'True
            TabIndex        =   614
            Tag             =   "4 digit at least"
            Top             =   5055
            Width           =   1365
         End
         Begin VB.TextBox txtNoOFDigitUser 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00C0C0FF&
            Height          =   315
            Index           =   6
            Left            =   1740
            RightToLeft     =   -1  'True
            TabIndex        =   613
            Top             =   5010
            Width           =   1365
         End
         Begin VB.TextBox txtNoOFDigitUser 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00C0C0FF&
            Height          =   315
            Index           =   9
            Left            =   4140
            RightToLeft     =   -1  'True
            TabIndex        =   612
            Top             =   5010
            Width           =   1365
         End
         Begin VB.TextBox txtDomainData 
            Alignment       =   1  'Right Justify
            Height          =   375
            Left            =   5040
            MaxLength       =   50
            RightToLeft     =   -1  'True
            TabIndex        =   579
            Top             =   6210
            Width           =   5235
         End
         Begin VB.CheckBox Chkbarcode 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "«·„‘€·"
            Height          =   285
            Index           =   199
            Left            =   8310
            RightToLeft     =   -1  'True
            TabIndex        =   563
            Top             =   8010
            Width           =   3645
         End
         Begin VB.CheckBox Chkbarcode 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "«· ﬂÊÌœ »ÿ—ﬁ «·œ›⁄"
            Height          =   285
            Index           =   185
            Left            =   8820
            RightToLeft     =   -1  'True
            TabIndex        =   540
            Top             =   7680
            Width           =   3135
         End
         Begin VB.TextBox txtNoOFDigitUser 
            Alignment       =   1  'Right Justify
            Height          =   315
            Index           =   1
            Left            =   6990
            RightToLeft     =   -1  'True
            TabIndex        =   516
            Top             =   8910
            Width           =   1365
         End
         Begin VB.TextBox txtNoOFDigitUser 
            Alignment       =   1  'Right Justify
            Height          =   315
            Index           =   0
            Left            =   6990
            RightToLeft     =   -1  'True
            TabIndex        =   514
            Top             =   8550
            Width           =   1365
         End
         Begin VB.CheckBox Chkbarcode 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "«·”—Ì«· ÿ»ﬁ« ··„” Œœ„ ›Ï «·ﬁÌÊœ"
            Height          =   285
            Index           =   163
            Left            =   7710
            RightToLeft     =   -1  'True
            TabIndex        =   513
            Top             =   7320
            Width           =   4245
         End
         Begin VB.CheckBox Chkbarcode 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "«·”—Ì«· ÿ»ﬁ« ··„” Œœ„ ›Ï «·Õ—ﬂ« "
            Height          =   285
            Index           =   162
            Left            =   7230
            RightToLeft     =   -1  'True
            TabIndex        =   512
            Top             =   7080
            Width           =   4725
         End
         Begin VB.TextBox XPTxtComment 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   3
            Left            =   5130
            MaxLength       =   50
            RightToLeft     =   -1  'True
            TabIndex        =   391
            Top             =   2040
            Width           =   5235
         End
         Begin VB.TextBox XPTxtComment 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   2
            Left            =   5040
            MaxLength       =   50
            RightToLeft     =   -1  'True
            TabIndex        =   359
            Top             =   1680
            Width           =   5235
         End
         Begin VB.TextBox XPTxtComment 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   1
            Left            =   5040
            MaxLength       =   50
            RightToLeft     =   -1  'True
            TabIndex        =   358
            Top             =   1320
            Width           =   5235
         End
         Begin VB.TextBox XPTxtCompanye 
            Alignment       =   1  'Right Justify
            Height          =   375
            Left            =   5040
            LinkTimeout     =   255
            MaxLength       =   255
            RightToLeft     =   -1  'True
            TabIndex        =   287
            Top             =   480
            Width           =   5235
         End
         Begin VB.TextBox XPTxtCompany 
            Alignment       =   1  'Right Justify
            Height          =   375
            Left            =   5040
            LinkTimeout     =   255
            MaxLength       =   255
            RightToLeft     =   -1  'True
            TabIndex        =   24
            Top             =   120
            Width           =   5235
         End
         Begin VB.TextBox XPTxtAddress 
            Alignment       =   1  'Right Justify
            Height          =   615
            Left            =   5040
            MaxLength       =   50
            MultiLine       =   -1  'True
            RightToLeft     =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   23
            Top             =   2460
            Width           =   5235
         End
         Begin VB.TextBox XPTxtPhone 
            Alignment       =   1  'Right Justify
            Height          =   375
            Left            =   7560
            MaxLength       =   50
            RightToLeft     =   -1  'True
            TabIndex        =   22
            Top             =   3630
            Width           =   2715
         End
         Begin VB.TextBox XPTxtMobile 
            Alignment       =   1  'Right Justify
            Height          =   375
            Left            =   5040
            MaxLength       =   50
            RightToLeft     =   -1  'True
            TabIndex        =   21
            Top             =   4140
            Width           =   1515
         End
         Begin VB.TextBox XPTxtMail 
            Alignment       =   1  'Right Justify
            Height          =   375
            Left            =   5040
            MaxLength       =   50
            RightToLeft     =   -1  'True
            TabIndex        =   20
            Top             =   4620
            Width           =   5235
         End
         Begin VB.TextBox XPTxtResponsable 
            Alignment       =   1  'Right Justify
            Height          =   375
            Left            =   7560
            MaxLength       =   50
            RightToLeft     =   -1  'True
            TabIndex        =   19
            Top             =   4140
            Width           =   2715
         End
         Begin VB.TextBox XPTxtComment 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   0
            Left            =   5040
            MaxLength       =   50
            RightToLeft     =   -1  'True
            TabIndex        =   18
            Top             =   930
            Width           =   5235
         End
         Begin VB.Frame Fra 
            BackColor       =   &H00E2E9E9&
            Caption         =   "‘⁄«— «·‘—ﬂ…"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   178
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   2715
            Index           =   2
            Left            =   1740
            RightToLeft     =   -1  'True
            TabIndex        =   14
            Top             =   6630
            Width           =   5235
            Begin VB.CheckBox Chkbarcode 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "—ﬁ„ «· VAT ÿ»ﬁ« ··‰‘«ÿ "
               Height          =   285
               Index           =   66
               Left            =   3120
               RightToLeft     =   -1  'True
               TabIndex        =   403
               Top             =   1080
               Width           =   2025
            End
            Begin VB.CheckBox Chkbarcode 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "⁄—÷ «··ÊÃÊ ÿ»ﬁ« ··›—⁄"
               Height          =   285
               Index           =   8
               Left            =   3120
               RightToLeft     =   -1  'True
               TabIndex        =   327
               Top             =   720
               Width           =   2025
            End
            Begin VB.CheckBox Chk 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "≈ŸÂ«— ‘⁄«— «·‘—ﬂ… ›Ï «· ﬁ«—Ì—"
               Height          =   435
               Left            =   2970
               RightToLeft     =   -1  'True
               TabIndex        =   16
               Top             =   270
               Width           =   2175
            End
            Begin ImpulseButton.ISButton Cmd 
               Height          =   375
               Left            =   240
               TabIndex        =   15
               Top             =   2310
               Width           =   1725
               _ExtentX        =   3043
               _ExtentY        =   661
               ButtonStyle     =   1
               ButtonPositionImage=   1
               Caption         =   "«Œ Ì«— «·‘⁄«—"
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
               ButtonImage     =   "FrmOptions.frx":0D39
               ColorButton     =   14871017
               DrawFocusRectangle=   0   'False
            End
            Begin Dynamic_Byte.NewViewBox ImgPic 
               Height          =   1875
               Left            =   60
               TabIndex        =   17
               Top             =   240
               Width           =   2865
               _ExtentX        =   5054
               _ExtentY        =   3307
            End
         End
         Begin VB.TextBox TxtFax 
            Alignment       =   1  'Right Justify
            Height          =   375
            Left            =   5040
            MaxLength       =   50
            RightToLeft     =   -1  'True
            TabIndex        =   13
            Top             =   3630
            Width           =   1515
         End
         Begin VB.TextBox TxtEmails 
            Alignment       =   1  'Right Justify
            Height          =   375
            Left            =   5040
            MaxLength       =   50
            RightToLeft     =   -1  'True
            TabIndex        =   12
            Top             =   3120
            Width           =   5235
         End
         Begin MSComDlg.CommonDialog Cdg 
            Left            =   570
            Top             =   1680
            _ExtentX        =   847
            _ExtentY        =   847
            _Version        =   393216
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "«·‰‘«ÿ"
            Height          =   375
            Index           =   66
            Left            =   3390
            RightToLeft     =   -1  'True
            TabIndex        =   631
            Top             =   3780
            Width           =   1395
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "«”„ «·‘«—⁄2"
            Height          =   375
            Index           =   46
            Left            =   7980
            RightToLeft     =   -1  'True
            TabIndex        =   629
            Top             =   5880
            Width           =   1005
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "—ﬁ„ «·„Œÿÿ"
            Height          =   375
            Index           =   48
            Left            =   10980
            RightToLeft     =   -1  'True
            TabIndex        =   628
            Top             =   5880
            Width           =   1005
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "«·„œÌ‰… «·›—⁄Ì…"
            Height          =   375
            Index           =   51
            Left            =   5580
            RightToLeft     =   -1  'True
            TabIndex        =   627
            Top             =   5880
            Width           =   1005
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "«·—„“ «·»—ÌœÌ*"
            Height          =   375
            Index           =   50
            Left            =   10740
            RightToLeft     =   -1  'True
            TabIndex        =   623
            Top             =   5550
            Width           =   1245
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "ﬂÊœ «·œÊ·…*"
            Height          =   375
            Index           =   53
            Left            =   8100
            RightToLeft     =   -1  'True
            TabIndex        =   622
            Top             =   5550
            Width           =   765
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "«”„ «·‘«—⁄*"
            Height          =   375
            Index           =   45
            Left            =   7860
            RightToLeft     =   -1  'True
            TabIndex        =   619
            Top             =   5175
            Width           =   1005
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "—ﬁ„ «·„»‰Ï*"
            Height          =   375
            Index           =   47
            Left            =   10980
            RightToLeft     =   -1  'True
            TabIndex        =   618
            Top             =   5115
            Width           =   1005
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "«·„œÌ‰…*"
            Height          =   375
            Index           =   49
            Left            =   2940
            RightToLeft     =   -1  'True
            TabIndex        =   617
            Top             =   5130
            Width           =   1005
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "«·ÕÌ*"
            Height          =   375
            Index           =   52
            Left            =   5580
            RightToLeft     =   -1  'True
            TabIndex        =   616
            Top             =   5130
            Width           =   765
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "„Êﬁ⁄ «·»Ì«‰« "
            Height          =   375
            Index           =   39
            Left            =   10230
            RightToLeft     =   -1  'True
            TabIndex        =   580
            Top             =   6210
            Width           =   1725
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "⁄œœ «·Œ«‰« "
            Height          =   315
            Index           =   33
            Left            =   8520
            RightToLeft     =   -1  'True
            TabIndex        =   517
            Top             =   8850
            Width           =   1035
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "⁄œœ «·Œ«‰« "
            Height          =   315
            Index           =   32
            Left            =   8520
            RightToLeft     =   -1  'True
            TabIndex        =   515
            Top             =   8520
            Width           =   1035
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "—ﬁ„ «· ”ÃÌ· VAT"
            Height          =   375
            Index           =   31
            Left            =   10200
            RightToLeft     =   -1  'True
            TabIndex        =   392
            Top             =   2040
            Width           =   1725
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "—ﬁ„ «·Õ«”»"
            Height          =   375
            Index           =   29
            Left            =   10200
            RightToLeft     =   -1  'True
            TabIndex        =   360
            Top             =   1680
            Width           =   1725
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "—ﬁ„ «·⁄÷ÊÌ…"
            Height          =   375
            Index           =   28
            Left            =   10200
            RightToLeft     =   -1  'True
            TabIndex        =   357
            Top             =   1320
            Width           =   1725
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "«·›«ﬂ”"
            Height          =   375
            Index           =   18
            Left            =   6480
            RightToLeft     =   -1  'True
            TabIndex        =   356
            Top             =   3720
            Width           =   885
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "«”„ «·„”ƒ·"
            Height          =   375
            Index           =   2
            Left            =   10230
            RightToLeft     =   -1  'True
            TabIndex        =   355
            Top             =   4260
            Width           =   1725
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "—ﬁ„ «·”Ã·"
            Height          =   375
            Index           =   21
            Left            =   10230
            RightToLeft     =   -1  'True
            TabIndex        =   286
            Top             =   960
            Width           =   1725
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "«·ÃÊ«·"
            Height          =   375
            Index           =   1
            Left            =   6030
            RightToLeft     =   -1  'True
            TabIndex        =   31
            Top             =   4260
            Width           =   1245
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "«·»—Ìœ «·«·ﬂ —Ê‰Ì"
            Height          =   375
            Index           =   3
            Left            =   10230
            RightToLeft     =   -1  'True
            TabIndex        =   30
            Top             =   4650
            Width           =   1725
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "«·Â« ›"
            Height          =   375
            Index           =   5
            Left            =   10230
            RightToLeft     =   -1  'True
            TabIndex        =   29
            Top             =   3750
            Width           =   1725
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "«”„ «·‘—ﬂ… «‰Ã·Ì“Ì"
            Height          =   375
            Index           =   8
            Left            =   10230
            RightToLeft     =   -1  'True
            TabIndex        =   28
            Top             =   570
            Width           =   1725
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "«·⁄‰Ê«‰"
            Height          =   375
            Index           =   9
            Left            =   10230
            RightToLeft     =   -1  'True
            TabIndex        =   27
            Top             =   2580
            Width           =   1725
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "«”„ «·‘—ﬂ… ⁄—»Ì"
            Height          =   375
            Index           =   10
            Left            =   10230
            RightToLeft     =   -1  'True
            TabIndex        =   26
            Top             =   120
            Width           =   1725
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "«·„Êﬁ⁄ «·«·ﬂ —Ê‰Ì"
            Height          =   375
            Index           =   20
            Left            =   10230
            RightToLeft     =   -1  'True
            TabIndex        =   25
            Top             =   3120
            Width           =   1725
         End
      End
      Begin C1SizerLibCtl.C1Elastic Ele 
         Height          =   9375
         Index           =   6
         Left            =   14805
         TabIndex        =   32
         TabStop         =   0   'False
         Top             =   45
         Width           =   12225
         _cx             =   21564
         _cy             =   16536
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
         Height          =   9375
         Index           =   7
         Left            =   15105
         TabIndex        =   33
         TabStop         =   0   'False
         Top             =   45
         Width           =   12225
         _cx             =   21564
         _cy             =   16536
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
         Height          =   9375
         Index           =   8
         Left            =   15405
         TabIndex        =   34
         TabStop         =   0   'False
         Top             =   45
         Width           =   12225
         _cx             =   21564
         _cy             =   16536
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
         Height          =   9375
         Index           =   9
         Left            =   15705
         TabIndex        =   35
         TabStop         =   0   'False
         Top             =   45
         Width           =   12225
         _cx             =   21564
         _cy             =   16536
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
         Begin VB.Frame Frame44 
            Caption         =   "ŒÌ«—«  «·„Œ«“‰"
            Height          =   8655
            Left            =   1680
            RightToLeft     =   -1  'True
            TabIndex        =   178
            Top             =   0
            Width           =   10305
            Begin VB.CheckBox CHKStore 
               Alignment       =   1  'Right Justify
               Caption         =   "«⁄ »«— «Œ— Ã—œ «⁄«œ…  ﬁÌÌ„"
               Height          =   285
               Index           =   8
               Left            =   6450
               RightToLeft     =   -1  'True
               TabIndex        =   651
               Top             =   4350
               Width           =   3765
            End
            Begin VB.CheckBox CHKStore 
               Alignment       =   1  'Right Justify
               Caption         =   "«· ﬂ·›…  Õ”» „‰ «Œ— Ã—œ"
               Height          =   285
               Index           =   7
               Left            =   6450
               RightToLeft     =   -1  'True
               TabIndex        =   650
               Top             =   4050
               Width           =   3765
            End
            Begin VB.CheckBox Chkbarcode 
               Alignment       =   1  'Right Justify
               Caption         =   "«·€«¡ «·ﬂ„Ì… «·„› —÷…"
               Height          =   285
               Index           =   205
               Left            =   -30
               RightToLeft     =   -1  'True
               TabIndex        =   571
               Top             =   5460
               Width           =   4575
            End
            Begin VB.CheckBox Chkbarcode 
               Alignment       =   1  'Right Justify
               Caption         =   "«·«’‰«› ÿ»ﬁ« ··›—⁄ ›Ï ‰ﬁ«ÿ «·»Ì⁄"
               Height          =   285
               Index           =   202
               Left            =   3300
               RightToLeft     =   -1  'True
               TabIndex        =   566
               Top             =   3720
               Width           =   2985
            End
            Begin VB.CheckBox CHKStore 
               Alignment       =   1  'Right Justify
               Caption         =   "«·”‰œ«  «·„Œ“‰Ì…. Ã„Ì⁄ «·«’‰«› ⁄·Ï „” ÊÏ «·”ÿ—"
               Height          =   285
               Index           =   6
               Left            =   360
               RightToLeft     =   -1  'True
               TabIndex        =   561
               Top             =   8280
               Width           =   9795
            End
            Begin VB.CheckBox CHKStore 
               Alignment       =   1  'Right Justify
               Caption         =   "«· ÕÊÌ· «·„Œ“‰Ì  Ã„Ì⁄ «·«’‰«› ⁄·Ï „” ÊÏ «·”ÿ—"
               Height          =   285
               Index           =   5
               Left            =   360
               RightToLeft     =   -1  'True
               TabIndex        =   560
               Top             =   8040
               Width           =   9795
            End
            Begin VB.CheckBox Chkbarcode 
               Alignment       =   1  'Right Justify
               Caption         =   "⁄œ„ «” Œœ«„ «·ÿ·»«  «·„Œ“‰Ì… «ﬂÀ— „‰ „—…"
               Height          =   285
               Index           =   196
               Left            =   1320
               RightToLeft     =   -1  'True
               TabIndex        =   559
               Top             =   7260
               Width           =   8895
            End
            Begin VB.CheckBox Chkbarcode 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "«· ⁄«„· »«·»«ﬁÌ ›Ì ”‰œ ’—› «·„Ê«œ »‰«¡ ⁄·Ï ›« Ê—… „»Ì⁄« "
               Height          =   285
               Index           =   108
               Left            =   1440
               RightToLeft     =   -1  'True
               TabIndex        =   558
               Top             =   4440
               Width           =   4845
            End
            Begin VB.CheckBox Chkbarcode 
               Alignment       =   1  'Right Justify
               Caption         =   "⁄œ„ «·”Õ» ⁄·Ï «·„ﬂ‘Ê› ›Ï «·ÿ·»«  «·œ«Œ·Ì…"
               Height          =   285
               Index           =   194
               Left            =   1320
               RightToLeft     =   -1  'True
               TabIndex        =   556
               Top             =   6990
               Width           =   8895
            End
            Begin VB.CheckBox Chkbarcode 
               Alignment       =   1  'Right Justify
               Caption         =   "”‰œ «·’—› »‰«¡ ⁄·Ì ›« Ê—… «·„»Ì⁄«  ·« Ìﬁ»· «Ì «’‰«› «Œ—Ì"
               Height          =   285
               Index           =   145
               Left            =   0
               RightToLeft     =   -1  'True
               TabIndex        =   552
               Top             =   4080
               Width           =   6285
            End
            Begin VB.TextBox TXTReturnSallingIntervalCount 
               Height          =   285
               Index           =   5
               Left            =   240
               TabIndex        =   548
               Top             =   4800
               Width           =   1215
            End
            Begin VB.CheckBox CHKStore 
               Alignment       =   1  'Right Justify
               Caption         =   "⁄„· «”„  ·ﬁ«∆Ï ··’‰›"
               Height          =   285
               Index           =   4
               Left            =   420
               RightToLeft     =   -1  'True
               TabIndex        =   541
               Top             =   7650
               Width           =   9795
            End
            Begin VB.CheckBox Chkbarcode 
               Alignment       =   1  'Right Justify
               Caption         =   "ﬂÊœ «·„Ã„Ê⁄Â ÌÕ ÊÌ ⁄·Ì  «· ”·”· «·‘Ã—Ì ··„ÃÊ⁄« "
               Height          =   285
               Index           =   38
               Left            =   210
               RightToLeft     =   -1  'True
               TabIndex        =   366
               Top             =   5370
               Width           =   10005
            End
            Begin VB.CheckBox Chkbarcode 
               Alignment       =   1  'Right Justify
               Caption         =   "«ﬁ›«· ”‰œ«  «· ÕÊÌ· ›Ì ›Ê« Ì— «·„»Ì⁄« "
               Height          =   285
               Index           =   130
               Left            =   1320
               RightToLeft     =   -1  'True
               TabIndex        =   471
               Top             =   6690
               Width           =   8895
            End
            Begin VB.CheckBox Chkbarcode 
               Alignment       =   1  'Right Justify
               Caption         =   "”‰œ «·«” ·«„ ⁄·Ì Õ”«» «·„Ê—œ"
               Height          =   285
               Index           =   96
               Left            =   1410
               RightToLeft     =   -1  'True
               TabIndex        =   436
               Top             =   6450
               Width           =   8805
            End
            Begin VB.TextBox TXTReturnSallingIntervalCount 
               Alignment       =   2  'Center
               Height          =   285
               Index           =   4
               Left            =   6480
               TabIndex        =   422
               Top             =   6210
               Width           =   1245
            End
            Begin VB.CheckBox Chkbarcode 
               Alignment       =   1  'Right Justify
               Caption         =   "ﬂ· «·«’‰«› Œ«÷⁄… ·· VAT"
               Height          =   285
               Index           =   82
               Left            =   1410
               RightToLeft     =   -1  'True
               TabIndex        =   421
               Top             =   6210
               Width           =   8805
            End
            Begin VB.CheckBox Chkbarcode 
               Alignment       =   1  'Right Justify
               Caption         =   " ⁄œÌ· «·ﬂ„Ì«  ÌœÊÌ« ›Ì «·„ﬂ”"
               Height          =   285
               Index           =   73
               Left            =   390
               RightToLeft     =   -1  'True
               TabIndex        =   412
               Top             =   5970
               Width           =   9825
            End
            Begin VB.CheckBox CHKStore 
               Alignment       =   1  'Right Justify
               Caption         =   "«· ⁄«„· »«ﬂÀ— „‰ „Œ“‰ ›Ì «·„»Ì⁄«  Ê«·„‘ —Ì« "
               Height          =   285
               Index           =   3
               Left            =   150
               RightToLeft     =   -1  'True
               TabIndex        =   405
               Top             =   5730
               Width           =   10065
            End
            Begin VB.TextBox TXTReturnSallingIntervalCount 
               Height          =   285
               Index           =   1
               Left            =   240
               TabIndex        =   383
               Top             =   5130
               Width           =   1215
            End
            Begin VB.CheckBox Chkbarcode 
               Alignment       =   1  'Right Justify
               Caption         =   " ﬂ·›… ÿ»ﬁ« ··”Ì—Ì«·"
               Height          =   285
               Index           =   41
               Left            =   720
               RightToLeft     =   -1  'True
               TabIndex        =   369
               Top             =   840
               Width           =   2985
            End
            Begin VB.CheckBox Chkbarcode 
               Alignment       =   1  'Right Justify
               Caption         =   " ﬂ·›… „ €Ì—…"
               Height          =   285
               Index           =   40
               Left            =   720
               RightToLeft     =   -1  'True
               TabIndex        =   368
               Top             =   600
               Width           =   2985
            End
            Begin VB.CheckBox Chkbarcode 
               Alignment       =   1  'Right Justify
               Caption         =   "ÿ»ﬁ« ··„Œ“‰"
               Height          =   285
               Index           =   39
               Left            =   720
               RightToLeft     =   -1  'True
               TabIndex        =   367
               Top             =   360
               Width           =   2985
            End
            Begin VB.CheckBox Chkbarcode 
               Alignment       =   1  'Right Justify
               Caption         =   "Ì „ œ„Ã ﬂÊœ «·„Ã„Ê⁄Â „⁄ ﬂÊœ «·’‰›"
               Height          =   285
               Index           =   10
               Left            =   5700
               RightToLeft     =   -1  'True
               TabIndex        =   339
               Top             =   5130
               Width           =   4515
            End
            Begin VB.CheckBox ChkCostStarting 
               Alignment       =   1  'Right Justify
               Caption         =   "«· ﬂ·›…  Õ”» „‰ »œ«Ì… «·› —… «·Õ«·Ì…"
               Height          =   285
               Left            =   4980
               RightToLeft     =   -1  'True
               TabIndex        =   294
               Top             =   4890
               Width           =   5235
            End
            Begin VB.CheckBox CHKStore 
               Alignment       =   1  'Right Justify
               Caption         =   "⁄‰œ › Õ „Œ“‰ Ì „ › Õ Õ”«»    Âœ«Ì« Ê⁄Ì‰«  ··„Œ“‰"
               Height          =   285
               Index           =   2
               Left            =   5580
               RightToLeft     =   -1  'True
               TabIndex        =   293
               Top             =   4650
               Width           =   4635
            End
            Begin VB.CheckBox CHKStore 
               Alignment       =   1  'Right Justify
               Caption         =   "⁄‰œ › Õ „Œ“‰ Ì „ › Õ Õ”«» ›ﬁœ Ê ·› ··„Œ“‰"
               Height          =   285
               Index           =   1
               Left            =   4980
               RightToLeft     =   -1  'True
               TabIndex        =   292
               Top             =   4650
               Width           =   5235
            End
            Begin VB.CheckBox CHKStore 
               Alignment       =   1  'Right Justify
               Caption         =   "Õ”«» «· ”ÊÌ«  «·Ã—œÌ… Ì »⁄ «·„Œ“Ê‰"
               Height          =   285
               Index           =   0
               Left            =   420
               RightToLeft     =   -1  'True
               TabIndex        =   291
               Top             =   3780
               Width           =   9795
            End
            Begin VB.ComboBox CboMainStockType 
               Height          =   315
               ItemData        =   "FrmOptions.frx":10D3
               Left            =   3840
               List            =   "FrmOptions.frx":10D5
               RightToLeft     =   -1  'True
               Style           =   2  'Dropdown List
               TabIndex        =   188
               Top             =   360
               Width           =   2745
            End
            Begin VB.CheckBox Chk2 
               Alignment       =   1  'Right Justify
               Caption         =   " «·”Õ» ⁄·Ï «·„ﬂ‘Ê› „‰ «·„Œ“‰"
               Height          =   285
               Left            =   5760
               RightToLeft     =   -1  'True
               TabIndex        =   187
               Top             =   720
               Width           =   2745
            End
            Begin VB.Frame Frame37 
               Caption         =   "«·ﬂ„Ì… «·„ «Õ…  ”«ÊÌ"
               ForeColor       =   &H000000FF&
               Height          =   855
               Left            =   2640
               RightToLeft     =   -1  'True
               TabIndex        =   184
               Top             =   2850
               Width           =   5895
               Begin VB.OptionButton OptCurrQty 
                  Alignment       =   1  'Right Justify
                  Caption         =   "«·ﬂ„Ì… «·„ÊÃÊœ… - «·ﬂ„Ì… «·„ÕÃÊ“…"
                  Height          =   195
                  Index           =   1
                  Left            =   120
                  RightToLeft     =   -1  'True
                  TabIndex        =   186
                  Top             =   480
                  Width           =   5055
               End
               Begin VB.OptionButton OptCurrQty 
                  Alignment       =   1  'Right Justify
                  Caption         =   "«·ﬂ„Ì… «·„ÊÃÊœ…  ﬂ«„·…"
                  Height          =   195
                  Index           =   0
                  Left            =   120
                  RightToLeft     =   -1  'True
                  TabIndex        =   185
                  Top             =   240
                  Value           =   -1  'True
                  Width           =   5055
               End
            End
            Begin VB.Frame Frame20 
               Caption         =   "«· ⁄«„· „⁄ «·«’‰«›"
               Height          =   1755
               Left            =   1560
               RightToLeft     =   -1  'True
               TabIndex        =   179
               Top             =   1080
               Width           =   6945
               Begin VB.CheckBox Chkbarcode 
                  Alignment       =   1  'Right Justify
                  Caption         =   "—»ÿ «·«’‰«› ÿ»ﬁ« ··‰‘«ÿ"
                  ForeColor       =   &H000000FF&
                  Height          =   285
                  Index           =   192
                  Left            =   0
                  RightToLeft     =   -1  'True
                  TabIndex        =   553
                  Top             =   480
                  Width           =   3795
               End
               Begin VB.CheckBox Chkbarcode 
                  Alignment       =   1  'Right Justify
                  Caption         =   "«· ⁄«„· »«·‰”» «·Â‰œ”Ì…"
                  Height          =   285
                  Index           =   160
                  Left            =   4200
                  RightToLeft     =   -1  'True
                  TabIndex        =   510
                  Top             =   1350
                  Width           =   2625
               End
               Begin VB.CheckBox Chkbarcode 
                  Alignment       =   1  'Right Justify
                  Caption         =   "«· ⁄«„· «··Êÿ Ê «—ÌŒ «·«‰ Â«¡ FIFO"
                  Height          =   285
                  Index           =   52
                  Left            =   1230
                  RightToLeft     =   -1  'True
                  TabIndex        =   380
                  Top             =   1320
                  Width           =   2625
               End
               Begin VB.CheckBox Chkbarcode 
                  Alignment       =   1  'Right Justify
                  Caption         =   "«· ⁄«„· »«·„”«ÕÂ"
                  Height          =   285
                  Index           =   49
                  Left            =   4200
                  RightToLeft     =   -1  'True
                  TabIndex        =   377
                  Top             =   1080
                  Width           =   2625
               End
               Begin VB.CheckBox Chkbarcode 
                  Alignment       =   1  'Right Justify
                  Caption         =   "«· ⁄«„· »«·«”„ «·„Œ ’—"
                  Height          =   285
                  Index           =   37
                  Left            =   120
                  RightToLeft     =   -1  'True
                  TabIndex        =   365
                  Top             =   1080
                  Width           =   1905
               End
               Begin VB.CheckBox Chkbarcode 
                  Alignment       =   1  'Right Justify
                  Caption         =   "Ì⁄„·  » Õ·Ì· «·«’‰«›"
                  Height          =   195
                  Index           =   31
                  Left            =   240
                  RightToLeft     =   -1  'True
                  TabIndex        =   350
                  Top             =   840
                  Width           =   1785
               End
               Begin VB.CheckBox Chkbarcode 
                  Alignment       =   1  'Right Justify
                  Caption         =   "«· ‰»ƒ »√”„ «·’‰› ÿ»ﬁ« ··„Ã„Ê⁄Â"
                  Height          =   285
                  Index           =   12
                  Left            =   4200
                  RightToLeft     =   -1  'True
                  TabIndex        =   308
                  Top             =   840
                  Width           =   2625
               End
               Begin VB.CheckBox Chkbarcode 
                  Alignment       =   1  'Right Justify
                  Caption         =   "—»ÿ «·«’‰«› »«·„Œ«“‰"
                  ForeColor       =   &H000000FF&
                  Height          =   285
                  Index           =   7
                  Left            =   4800
                  RightToLeft     =   -1  'True
                  TabIndex        =   305
                  Top             =   480
                  Width           =   2025
               End
               Begin VB.CheckBox Chkbarcode 
                  Alignment       =   1  'Right Justify
                  Caption         =   "«·”„«Õ » ﬂ—«— «·«”„«¡"
                  Height          =   285
                  Index           =   1
                  Left            =   1830
                  RightToLeft     =   -1  'True
                  TabIndex        =   299
                  Top             =   1080
                  Width           =   2025
               End
               Begin VB.CheckBox Chkbarcode 
                  Alignment       =   1  'Right Justify
                  Caption         =   "«· ⁄«„· »«·»«—ﬂÊœ"
                  Height          =   285
                  Index           =   0
                  Left            =   2190
                  RightToLeft     =   -1  'True
                  TabIndex        =   297
                  Top             =   840
                  Width           =   1665
               End
               Begin VB.CheckBox ChkitemsWorkWithColor 
                  Alignment       =   1  'Right Justify
                  Caption         =   "Ì⁄„· »«··Ê‰"
                  Height          =   195
                  Left            =   4200
                  RightToLeft     =   -1  'True
                  TabIndex        =   183
                  Top             =   240
                  Width           =   1095
               End
               Begin VB.CheckBox ChkitemsWorkWithSize 
                  Alignment       =   1  'Right Justify
                  Caption         =   "Ì⁄„·  »«·„ﬁ«”"
                  Height          =   195
                  Left            =   5520
                  RightToLeft     =   -1  'True
                  TabIndex        =   182
                  Top             =   240
                  Width           =   1335
               End
               Begin VB.CheckBox ChkitemsWorkWithDates 
                  Alignment       =   1  'Right Justify
                  Caption         =   "Ì⁄„·  » «—ÌŒ «·’·«ÕÌ…"
                  Height          =   285
                  Left            =   240
                  RightToLeft     =   -1  'True
                  TabIndex        =   181
                  Top             =   240
                  Width           =   1785
               End
               Begin VB.CheckBox ChkitemsWorkWithClass 
                  Alignment       =   1  'Right Justify
                  Caption         =   "Ì⁄„·  »«·›—“"
                  Height          =   195
                  Left            =   2640
                  RightToLeft     =   -1  'True
                  TabIndex        =   180
                  Top             =   240
                  Width           =   1215
               End
            End
            Begin VB.Label Label25 
               Alignment       =   1  'Right Justify
               Caption         =   "«·ﬂ„ÌÂ «·«› —«÷Ì… ··‘«‘« "
               Height          =   375
               Left            =   1500
               RightToLeft     =   -1  'True
               TabIndex        =   549
               Top             =   4800
               Width           =   3015
            End
            Begin VB.Label Label20 
               Alignment       =   1  'Right Justify
               Caption         =   "Õœœ ›«’·  ﬂÊÌœ «·ÕﬁÊ· ··«’‰«›"
               Height          =   495
               Left            =   1440
               RightToLeft     =   -1  'True
               TabIndex        =   340
               Top             =   5130
               Width           =   3015
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               Caption         =   "ÿ—Ìﬁ… Õ”«»  ﬂ·›… «·„Œ“Ê‰"
               Height          =   285
               Index           =   12
               Left            =   6630
               RightToLeft     =   -1  'True
               TabIndex        =   189
               Top             =   360
               Width           =   1875
            End
         End
      End
      Begin C1SizerLibCtl.C1Elastic Ele 
         Height          =   9375
         Index           =   10
         Left            =   16005
         TabIndex        =   36
         TabStop         =   0   'False
         Top             =   45
         Width           =   12225
         _cx             =   21564
         _cy             =   16536
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
         Begin C1SizerLibCtl.C1Elastic Ele 
            Height          =   9285
            Index           =   21
            Left            =   1860
            TabIndex        =   89
            TabStop         =   0   'False
            Top             =   30
            Width           =   10185
            _cx             =   17965
            _cy             =   16378
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
            Begin VB.CheckBox Chkbarcode 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "—ﬁ„ «·›«  €Ì— «·“«„Ì «·“«„Ï ··⁄„·«¡ «·‰ﬁœÌ"
               Height          =   285
               Index           =   214
               Left            =   2520
               RightToLeft     =   -1  'True
               TabIndex        =   632
               Top             =   8520
               Width           =   3225
            End
            Begin VB.CheckBox Chkbarcode 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "«ŸÂ«—  ›«’Ì· «·⁄œ”«  ›Ï «·„»Ì⁄« "
               Height          =   285
               Index           =   191
               Left            =   7380
               RightToLeft     =   -1  'True
               TabIndex        =   547
               Top             =   6090
               Width           =   2745
            End
            Begin VB.CheckBox Chkbarcode 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "”‰œ ﬁ»÷ ⁄«„ «·‰ﬁÿÂ „—Â Ê«Õœ… ÌÊ„Ì« ·ﬂ· ‰ﬁÿÂ"
               Height          =   405
               Index           =   186
               Left            =   -120
               RightToLeft     =   -1  'True
               TabIndex        =   542
               Top             =   8760
               Width           =   2145
            End
            Begin VB.CheckBox Chkbarcode 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "«ŸÂ«— —’Ìœ «·⁄„Ì· ›Ï ›« Ê—… «·„»Ì⁄« "
               Height          =   285
               Index           =   183
               Left            =   6660
               RightToLeft     =   -1  'True
               TabIndex        =   538
               Top             =   9000
               Width           =   3465
            End
            Begin VB.CheckBox Chkbarcode 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "—ﬁ„ «·ÃÊ«· «·“«„Ï ›Ï ‘«‘… «·⁄„·«¡"
               Height          =   405
               Index           =   178
               Left            =   0
               RightToLeft     =   -1  'True
               TabIndex        =   533
               Top             =   8040
               Width           =   1905
            End
            Begin VB.CheckBox Chkbarcode 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   " — Ì» «·›« Ê—… ÿ»ﬁ« ·Êﬁ  «·«œŒ«·"
               Height          =   285
               Index           =   175
               Left            =   -120
               RightToLeft     =   -1  'True
               TabIndex        =   530
               Top             =   1110
               Width           =   2745
            End
            Begin VB.TextBox MYTEXT 
               Alignment       =   2  'Center
               Height          =   285
               Index           =   0
               Left            =   330
               TabIndex        =   526
               Top             =   720
               Width           =   615
            End
            Begin VB.CheckBox Chkbarcode 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "«‰‘«¡ ”‰œ ﬁ»÷ ¬·Ï »«·„œ›Ê⁄ ›Ï ›« Ê—… «·»Ì⁄"
               Height          =   285
               Index           =   168
               Left            =   2130
               RightToLeft     =   -1  'True
               TabIndex        =   522
               Top             =   8970
               Width           =   3465
            End
            Begin VB.CheckBox Chkbarcode 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "„—œÊœ«  ‰ﬁ«ÿ «·»Ì⁄ »«·»«—ﬂÊœ"
               Height          =   405
               Index           =   166
               Left            =   0
               RightToLeft     =   -1  'True
               TabIndex        =   520
               Top             =   7680
               Width           =   2295
            End
            Begin VB.CheckBox Chkbarcode 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "⁄œ„  Ã«Ê“ «„— «·‘—«¡ ›Ï «·„»Ì⁄« "
               Height          =   525
               Index           =   152
               Left            =   60
               RightToLeft     =   -1  'True
               TabIndex        =   502
               Top             =   7260
               Width           =   2295
            End
            Begin VB.CheckBox Chkbarcode 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "—»ÿ «”⁄«— «·⁄„·«¡ „⁄ «·„‰«œÌ»  "
               ForeColor       =   &H00FF0000&
               Height          =   285
               Index           =   149
               Left            =   6660
               RightToLeft     =   -1  'True
               TabIndex        =   499
               Top             =   8280
               Width           =   3465
            End
            Begin VB.CheckBox Chkbarcode 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "Ì „ «·’—› »”⁄— «· ﬂ·›… ÿ»ﬁ« ··„ﬂ” ›« Ê—… «·„»Ì⁄« "
               Height          =   285
               Index           =   142
               Left            =   930
               RightToLeft     =   -1  'True
               TabIndex        =   493
               Top             =   4440
               Width           =   4815
            End
            Begin VB.CheckBox Chkbarcode 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E0E0E0&
               Caption         =   "”‰œ «·ﬁ»÷ «·⁄«„ Ì‰‘√ ﬁÌœ «·«ﬁ›«·"
               ForeColor       =   &H00FF0000&
               Height          =   405
               Index           =   140
               Left            =   0
               RightToLeft     =   -1  'True
               TabIndex        =   491
               Top             =   6840
               Width           =   2355
            End
            Begin VB.CheckBox Chkbarcode 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "ÿ»«⁄Â «·›« Ê—… ÿ»ﬁ« ··›—⁄"
               ForeColor       =   &H000000FF&
               Height          =   285
               Index           =   136
               Left            =   240
               RightToLeft     =   -1  'True
               TabIndex        =   480
               Top             =   6600
               Width           =   2115
            End
            Begin VB.CheckBox Chkbarcode 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "«ŸÂ«— «’‰«› «·»Ì⁄ ›ﬁÿ"
               Height          =   285
               Index           =   135
               Left            =   270
               RightToLeft     =   -1  'True
               TabIndex        =   479
               Top             =   6360
               Width           =   2115
            End
            Begin VB.TextBox txtLimitDefaultCreditDays 
               Alignment       =   1  'Right Justify
               Height          =   375
               Left            =   6060
               RightToLeft     =   -1  'True
               TabIndex        =   475
               Top             =   1830
               Width           =   975
            End
            Begin VB.TextBox txtLimitDefaultCredit 
               Alignment       =   1  'Right Justify
               Height          =   375
               Left            =   6090
               RightToLeft     =   -1  'True
               TabIndex        =   473
               Top             =   1410
               Width           =   975
            End
            Begin VB.CheckBox Chkbarcode 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "⁄œ„ «·”„«Õ » ⁄œÌ· «·„‰œÊ» ›Ì ›« Ê—Â «·„»Ì⁄« "
               Height          =   315
               Index           =   131
               Left            =   1020
               RightToLeft     =   -1  'True
               TabIndex        =   472
               Top             =   8760
               Width           =   4725
            End
            Begin VB.CheckBox Chkbarcode 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "«·”„«Õ » ⁄œÌ· «·›Ê« Ì— «· Ì ⁄·ÌÂ« „—œÊœ« "
               Height          =   525
               Index           =   128
               Left            =   -90
               RightToLeft     =   -1  'True
               TabIndex        =   468
               Top             =   5880
               Width           =   2475
            End
            Begin VB.CheckBox Chkbarcode 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "«·”„«Õ » ⁄œÌ· «·›Ê« Ì— «· Ì ⁄·”Â« «‘⁄«—« "
               Height          =   285
               Index           =   127
               Left            =   2400
               RightToLeft     =   -1  'True
               TabIndex        =   467
               Top             =   6360
               Width           =   3345
            End
            Begin VB.CheckBox Chkbarcode 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   " ⁄«„· «·„—œÊœ«  »«· FIFO"
               Height          =   285
               Index           =   122
               Left            =   30
               RightToLeft     =   -1  'True
               TabIndex        =   462
               Top             =   5160
               Width           =   2355
            End
            Begin VB.CheckBox Chkbarcode 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "«·”„«Õ » ﬂ—«— —ﬁ„ «·›« Ê—…"
               Height          =   285
               Index           =   121
               Left            =   30
               RightToLeft     =   -1  'True
               TabIndex        =   460
               Top             =   4920
               Width           =   2355
            End
            Begin VB.CheckBox Chkbarcode 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "«·”⁄— ÿ»ﬁ« ··„ﬁ«”"
               Height          =   285
               Index           =   113
               Left            =   30
               RightToLeft     =   -1  'True
               TabIndex        =   452
               Top             =   4680
               Width           =   2355
            End
            Begin VB.CheckBox Chkbarcode 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "«· ⁄«„· »Õ”«» «·œ›⁄«  «·„ﬁœ„… ›Ì „ﬁ»Ê÷«  «·⁄„Ì·"
               Height          =   285
               Index           =   106
               Left            =   900
               RightToLeft     =   -1  'True
               TabIndex        =   446
               Top             =   8280
               Width           =   4845
            End
            Begin VB.CheckBox Chkbarcode 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "„·› «·⁄„·«¡  —ﬁ„ «·”Ã· €Ì— «·“«„Ì"
               Height          =   285
               Index           =   105
               Left            =   2520
               RightToLeft     =   -1  'True
               TabIndex        =   445
               Top             =   4680
               Width           =   3225
            End
            Begin VB.CheckBox Chkbarcode 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "«· ⁄«„· »‰ﬁ«ÿ «·⁄„·«¡"
               Height          =   285
               Index           =   103
               Left            =   6660
               RightToLeft     =   -1  'True
               TabIndex        =   443
               Top             =   8730
               Width           =   3465
            End
            Begin VB.CheckBox Chkbarcode 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "«·ﬁÌ„… ‘«„·… ··÷—Ì»…"
               Height          =   285
               Index           =   102
               Left            =   7380
               RightToLeft     =   -1  'True
               TabIndex        =   442
               Top             =   6360
               Width           =   2745
            End
            Begin VB.CheckBox Chkbarcode 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "«ŸÂ«— ‰ﬁÿ… Ê«ÃÂÂ —ﬁ„ 2"
               ForeColor       =   &H00FF0000&
               Height          =   285
               Index           =   87
               Left            =   6660
               RightToLeft     =   -1  'True
               TabIndex        =   427
               Top             =   7290
               Width           =   3465
            End
            Begin VB.CheckBox Chkbarcode 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "«—”«· ··«⁄ „«œ ⁄‰œ  ŒÿÌ «·Õœ «·√ „«‰Ì ··⁄„Ì·"
               Height          =   285
               Index           =   85
               Left            =   1200
               RightToLeft     =   -1  'True
               TabIndex        =   425
               Top             =   7800
               Width           =   4545
            End
            Begin VB.CheckBox Chkbarcode 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "«·—ﬁ„ «·ÌœÊÌ «·“«„Ì ›Ì ›« Ê—… «·„»Ì⁄« "
               Height          =   285
               Index           =   81
               Left            =   1110
               RightToLeft     =   -1  'True
               TabIndex        =   420
               Top             =   7560
               Width           =   4635
            End
            Begin VB.CheckBox Chkbarcode 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "«„ﬂ«‰Ì… ⁄„· „—œÊœ«  »œÊ‰  ﬂ·›…"
               Height          =   285
               Index           =   76
               Left            =   3240
               RightToLeft     =   -1  'True
               TabIndex        =   415
               Top             =   5880
               Width           =   2505
            End
            Begin VB.CheckBox Chkbarcode 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "«·Õ”«» ÿ»ﬁ« ··«”«” «·‰ﬁœÌ"
               Height          =   285
               Index           =   74
               Left            =   2520
               RightToLeft     =   -1  'True
               TabIndex        =   413
               Top             =   6120
               Width           =   3225
            End
            Begin VB.CheckBox Chkbarcode 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   " Ã„Ì⁄ «·’‰› ⁄·Ï „” ÊÏ «·”ÿ—"
               Height          =   285
               Index           =   72
               Left            =   6960
               RightToLeft     =   -1  'True
               TabIndex        =   411
               Top             =   6810
               Width           =   3165
            End
            Begin VB.CheckBox Chkbarcode 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "«ﬁ —«Õ «Œ— ”⁄— ·»Ì⁄ «·’‰›"
               Height          =   285
               Index           =   71
               Left            =   150
               RightToLeft     =   -1  'True
               TabIndex        =   410
               Top             =   5640
               Width           =   2235
            End
            Begin VB.CheckBox Chkbarcode 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "„»Ì⁄«  «·»÷«⁄… «·«„«‰…  ƒÀ— „»«‘—… ›Ì Õ”«» «·„Ê—œ"
               Height          =   285
               Index           =   68
               Left            =   2460
               RightToLeft     =   -1  'True
               TabIndex        =   406
               Top             =   6840
               Width           =   4425
            End
            Begin VB.CheckBox Chkbarcode 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "’—› „Ê«œ «·Œ«„ ÿ»ﬁ« ··„ﬂ”"
               Height          =   285
               Index           =   64
               Left            =   2490
               RightToLeft     =   -1  'True
               TabIndex        =   401
               Top             =   5400
               Width           =   3255
            End
            Begin VB.CheckBox Chkbarcode 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "«ŸÂ«— «·«’‰«› ÿ»ﬁ« ··⁄„Ì·"
               Height          =   285
               Index           =   63
               Left            =   150
               RightToLeft     =   -1  'True
               TabIndex        =   400
               Top             =   5400
               Width           =   2235
            End
            Begin VB.CheckBox Chkbarcode 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "«„— «·»Ì⁄ „⁄ ›« Ê—… «·„»Ì⁄«  «· ⁄«„· »«·»«ﬁÌ"
               Height          =   285
               Index           =   61
               Left            =   6660
               RightToLeft     =   -1  'True
               TabIndex        =   398
               Top             =   8490
               Width           =   3465
            End
            Begin VB.CheckBox Chkbarcode 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "«·⁄„· »”Ì«”…  ⁄ÃÌ· «·œ›⁄"
               Height          =   285
               Index           =   51
               Left            =   2370
               RightToLeft     =   -1  'True
               TabIndex        =   379
               Top             =   5640
               Width           =   3375
            End
            Begin VB.CheckBox Chkbarcode 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "ﬁÌœ «·⁄„Ê·… ›Ì «·›Ì“« ÌŒ’„ „‰ «·«Ã„«·Ì"
               Height          =   285
               Index           =   48
               Left            =   780
               RightToLeft     =   -1  'True
               TabIndex        =   376
               Top             =   7080
               Width           =   4965
            End
            Begin VB.CheckBox Chkbarcode 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "«”„ «·⁄„Ì· «·‰ﬁœÌ Ê—ﬁ„ «· ·Ì›Ê‰ «·“«„Ì ›Ì ›« Ê—… «·»Ì⁄"
               Height          =   285
               Index           =   45
               Left            =   2460
               RightToLeft     =   -1  'True
               TabIndex        =   373
               Top             =   6600
               Width           =   4425
            End
            Begin VB.CheckBox Chkbarcode 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "«· ⁄«„· »ÿ—ﬁ œ›⁄ „ ⁄œœ… ›Ì ›« Ê—… «·»Ì⁄"
               Height          =   285
               Index           =   44
               Left            =   930
               RightToLeft     =   -1  'True
               TabIndex        =   372
               Top             =   7320
               Width           =   4815
            End
            Begin VB.CheckBox Chkbarcode 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "«· ⁄«„· »√⁄„«— «·œÌÊ‰"
               Height          =   285
               Index           =   34
               Left            =   2520
               RightToLeft     =   -1  'True
               TabIndex        =   354
               Top             =   5160
               Width           =   3225
            End
            Begin VB.CheckBox Chkbarcode 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "«·Œ’„ ›Ì ›Ê« Ì— «·„»Ì⁄«  Ì‰‘Ï¡ ﬁÌœ"
               Height          =   285
               Index           =   28
               Left            =   1080
               RightToLeft     =   -1  'True
               TabIndex        =   330
               Top             =   8040
               Width           =   4665
            End
            Begin VB.CheckBox Chkbarcode 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "«· ⁄«„· »« ›«ﬁÌ«  «·⁄„·«¡"
               ForeColor       =   &H00FF0000&
               Height          =   285
               Index           =   22
               Left            =   6660
               RightToLeft     =   -1  'True
               TabIndex        =   322
               Top             =   8010
               Width           =   3465
            End
            Begin VB.CheckBox Chkbarcode 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "«·Õ«·Â «·€«·»Â ··’‰› «·„·Õﬁ „Ã«‰Ì"
               ForeColor       =   &H000080FF&
               Height          =   285
               Index           =   17
               Left            =   2550
               RightToLeft     =   -1  'True
               TabIndex        =   318
               Top             =   4920
               Width           =   3195
            End
            Begin VB.CheckBox Chkbarcode 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "«ŸÂ«—  Õ–Ì—«   ﬂ·›… «·«’‰«› »«·«·Ê«‰"
               Height          =   285
               Index           =   18
               Left            =   6660
               RightToLeft     =   -1  'True
               TabIndex        =   317
               Top             =   7770
               Width           =   3465
            End
            Begin VB.CheckBox Chkbarcode 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "⁄œ„ «·”„«Õ ··„‰œÊ» » ŒÿÌ ‰”»… «·Œ’„"
               Height          =   285
               Index           =   15
               Left            =   6660
               RightToLeft     =   -1  'True
               TabIndex        =   310
               Top             =   7530
               Width           =   3465
            End
            Begin VB.CheckBox Chkbarcode 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "«·›« Ê—… «·«› —«÷Ì… «Ã·…"
               Height          =   285
               Index           =   13
               Left            =   60
               RightToLeft     =   -1  'True
               TabIndex        =   309
               Top             =   120
               Width           =   2325
            End
            Begin VB.CheckBox Chkbarcode 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "«ŸÂ«— ‰ﬁÿ… »Ì⁄  Ã«—Ì…"
               ForeColor       =   &H00C00000&
               Height          =   285
               Index           =   2
               Left            =   8070
               RightToLeft     =   -1  'True
               TabIndex        =   300
               Top             =   7050
               Width           =   2115
            End
            Begin VB.CheckBox ChkItemsattachedzero 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "«·«’‰«› «·„·Õﬁ…  ŸÂ— »œÊ‰ ”⁄—"
               ForeColor       =   &H000080FF&
               Height          =   285
               Left            =   6900
               RightToLeft     =   -1  'True
               TabIndex        =   296
               Top             =   6570
               Width           =   3225
            End
            Begin VB.Frame Fra 
               BackColor       =   &H00E2E9E9&
               Caption         =   "⁄‰œ  ”ÃÌ· ›« Ê—… ÃœÌœ…"
               ForeColor       =   &H000000FF&
               Height          =   1065
               Index           =   0
               Left            =   0
               RightToLeft     =   -1  'True
               TabIndex        =   111
               Top             =   1260
               Width           =   5745
               Begin VB.OptionButton Opt 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   " «—ÌŒ «·›« Ê—… ÂÊ «· «—ÌŒ «·Õ«·Ï ›Ï «·ÃÂ«“ "
                  Height          =   255
                  Index           =   0
                  Left            =   90
                  RightToLeft     =   -1  'True
                  TabIndex        =   115
                  Top             =   240
                  Value           =   -1  'True
                  Width           =   5565
               End
               Begin VB.OptionButton Opt 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   " «—ÌŒ «Œ— ›« Ê—… »Ì⁄ „”Ã·… ﬁ»·Â«"
                  Height          =   255
                  Index           =   1
                  Left            =   90
                  RightToLeft     =   -1  'True
                  TabIndex        =   114
                  Top             =   480
                  Width           =   5565
               End
               Begin VB.OptionButton Opt 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   " «—ÌŒ «·ÃÂ«“ «·”Ì—›— ( „ «Õ ›ﬁÿ ›Ï Õ«·… ⁄„·  «·‘»ﬂ…)"
                  Height          =   255
                  Index           =   2
                  Left            =   90
                  RightToLeft     =   -1  'True
                  TabIndex        =   113
                  Top             =   780
                  Width           =   5535
               End
               Begin VB.Frame Frame29 
                  Caption         =   "Frame29"
                  Height          =   15
                  Index           =   0
                  Left            =   120
                  RightToLeft     =   -1  'True
                  TabIndex        =   112
                  Top             =   1440
                  Width           =   3975
               End
            End
            Begin VB.Frame Frame28 
               BackColor       =   &H00E2E9E9&
               Caption         =   "«Ê·ÊÌ«  Œ’Ê„«  «·»Ì⁄"
               ForeColor       =   &H000000FF&
               Height          =   1725
               Left            =   5910
               RightToLeft     =   -1  'True
               TabIndex        =   102
               Top             =   4380
               Width           =   3615
               Begin VB.TextBox TxtSaleDiscount3 
                  Alignment       =   1  'Right Justify
                  Height          =   285
                  Left            =   240
                  RightToLeft     =   -1  'True
                  TabIndex        =   488
                  Top             =   1080
                  Width           =   975
               End
               Begin VB.TextBox TxtSaleDiscount1 
                  Alignment       =   1  'Right Justify
                  Height          =   285
                  Left            =   240
                  RightToLeft     =   -1  'True
                  TabIndex        =   105
                  Top             =   360
                  Width           =   975
               End
               Begin VB.TextBox TxtSaleDiscount2 
                  Alignment       =   1  'Right Justify
                  Height          =   285
                  Left            =   240
                  RightToLeft     =   -1  'True
                  TabIndex        =   104
                  Top             =   720
                  Width           =   975
               End
               Begin VB.TextBox TxtSaleDiscount4 
                  Alignment       =   1  'Right Justify
                  Height          =   285
                  Left            =   240
                  RightToLeft     =   -1  'True
                  TabIndex        =   103
                  Top             =   1410
                  Width           =   975
               End
               Begin VB.Label Label3 
                  Alignment       =   1  'Right Justify
                  BackStyle       =   0  'Transparent
                  Caption         =   "Œ’„ «·’‰›"
                  Height          =   255
                  Left            =   840
                  RightToLeft     =   -1  'True
                  TabIndex        =   110
                  Top             =   360
                  Width           =   1815
               End
               Begin VB.Label Label4 
                  Alignment       =   1  'Right Justify
                  BackStyle       =   0  'Transparent
                  Caption         =   "Œ’„ „Ã„Ê⁄Â  «·’‰›"
                  Height          =   255
                  Left            =   840
                  RightToLeft     =   -1  'True
                  TabIndex        =   109
                  Top             =   720
                  Width           =   1815
               End
               Begin VB.Label Label6 
                  Alignment       =   1  'Right Justify
                  BackStyle       =   0  'Transparent
                  Caption         =   "Œ’„ «·⁄„Ì·"
                  Height          =   255
                  Left            =   840
                  RightToLeft     =   -1  'True
                  TabIndex        =   108
                  Top             =   1080
                  Width           =   1815
               End
               Begin VB.Label Label1 
                  Alignment       =   1  'Right Justify
                  BackStyle       =   0  'Transparent
                  Caption         =   "Œ’„ «·„‰œÊ»"
                  Height          =   255
                  Index           =   0
                  Left            =   840
                  RightToLeft     =   -1  'True
                  TabIndex        =   107
                  Top             =   1440
                  Width           =   1815
               End
               Begin VB.Label Label2 
                  Alignment       =   1  'Right Justify
                  BackStyle       =   0  'Transparent
                  Caption         =   "«· — Ì»"
                  Height          =   255
                  Left            =   240
                  RightToLeft     =   -1  'True
                  TabIndex        =   106
                  Top             =   120
                  Width           =   735
               End
            End
            Begin VB.Frame Fra 
               BackColor       =   &H00E2E9E9&
               Caption         =   "”Ì«”Â «·«” —Ã«⁄ Ê «·«” »œ«·"
               ClipControls    =   0   'False
               ForeColor       =   &H000000FF&
               Height          =   1185
               Index           =   3
               Left            =   840
               RightToLeft     =   -1  'True
               TabIndex        =   92
               Top             =   3240
               Width           =   8655
               Begin VB.CheckBox Chkbarcode 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "«·„— Ã⁄ »«·»«—ﬂÊœ ›ﬁÿ"
                  Height          =   285
                  Index           =   47
                  Left            =   2760
                  RightToLeft     =   -1  'True
                  TabIndex        =   375
                  Top             =   840
                  Width           =   4665
               End
               Begin VB.Frame Frame30 
                  Caption         =   "Frame29"
                  Height          =   15
                  Left            =   120
                  RightToLeft     =   -1  'True
                  TabIndex        =   97
                  Top             =   1440
                  Width           =   3975
               End
               Begin VB.OptionButton Opt 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "«·„— Ã⁄«  »œÊ‰ ›« Ê—…  Ê»›« Ê—…"
                  Height          =   255
                  Index           =   6
                  Left            =   4590
                  RightToLeft     =   -1  'True
                  TabIndex        =   96
                  Top             =   300
                  Width           =   2865
               End
               Begin VB.OptionButton Opt 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "«·„— Ã⁄«  »›« Ê—… ›ﬁÿ "
                  Height          =   255
                  Index           =   7
                  Left            =   4710
                  RightToLeft     =   -1  'True
                  TabIndex        =   95
                  Top             =   600
                  Width           =   2745
               End
               Begin VB.TextBox TXTReturnSallingIntervalCount 
                  Height          =   285
                  Index           =   0
                  Left            =   3000
                  TabIndex        =   94
                  Top             =   600
                  Width           =   615
               End
               Begin VB.TextBox TXTReturnSallingIntervalCount1 
                  Height          =   285
                  Left            =   480
                  TabIndex        =   93
                  Top             =   600
                  Width           =   615
               End
               Begin VB.Label Label5 
                  Alignment       =   1  'Right Justify
                  BackStyle       =   0  'Transparent
                  Caption         =   "ÌÊ„"
                  Height          =   255
                  Left            =   0
                  RightToLeft     =   -1  'True
                  TabIndex        =   101
                  Top             =   600
                  Width           =   375
               End
               Begin VB.Label Label7 
                  Alignment       =   1  'Right Justify
                  BackStyle       =   0  'Transparent
                  Caption         =   "› —… «·„— Ã⁄"
                  Height          =   255
                  Left            =   3600
                  RightToLeft     =   -1  'True
                  TabIndex        =   100
                  Top             =   600
                  Width           =   1095
               End
               Begin VB.Label Label9 
                  Alignment       =   1  'Right Justify
                  BackStyle       =   0  'Transparent
                  Caption         =   "› —… «·«” »œ«·"
                  Height          =   255
                  Left            =   1080
                  RightToLeft     =   -1  'True
                  TabIndex        =   99
                  Top             =   600
                  Width           =   1335
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackStyle       =   0  'Transparent
                  Caption         =   "ÌÊ„"
                  Height          =   255
                  Index           =   37
                  Left            =   2520
                  RightToLeft     =   -1  'True
                  TabIndex        =   98
                  Top             =   600
                  Width           =   375
               End
            End
            Begin VB.Frame Fra 
               BackColor       =   &H00E2E9E9&
               Caption         =   "«·„»Ì⁄«  Ê«–‰ «·’—› "
               ForeColor       =   &H000000FF&
               Height          =   975
               Index           =   7
               Left            =   45
               RightToLeft     =   -1  'True
               TabIndex        =   90
               Top             =   2250
               Width           =   9495
               Begin VB.CheckBox Chkbarcode 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "«·»«—ﬂÊœ ÿ»ﬁ« ··ÊÕœ…"
                  ForeColor       =   &H00FF0000&
                  Height          =   285
                  Index           =   169
                  Left            =   30
                  RightToLeft     =   -1  'True
                  TabIndex        =   638
                  Top             =   120
                  Width           =   2325
               End
               Begin VB.TextBox TXTReturnSallingIntervalCount 
                  Height          =   285
                  Index           =   6
                  Left            =   1680
                  TabIndex        =   567
                  Top             =   720
                  Width           =   615
               End
               Begin VB.CheckBox Chkbarcode 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "⁄‰œ ⁄„· „—œÊœ«  „»Ì⁄«  ·« Ì „ «‰‘«¡ ”‰œ «” ·«„ «·Ì"
                  Height          =   195
                  Index           =   158
                  Left            =   -2520
                  RightToLeft     =   -1  'True
                  TabIndex        =   508
                  Top             =   480
                  Width           =   6915
               End
               Begin VB.CheckBox Chkbarcode 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "›« Ê—… «·„»Ì⁄«  ·« ‰‘√ ﬁÌœ"
                  Height          =   285
                  Index           =   141
                  Left            =   5580
                  RightToLeft     =   -1  'True
                  TabIndex        =   492
                  Top             =   720
                  Width           =   3675
               End
               Begin VB.OptionButton Opt_OrderOut 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "⁄„· «–‰ ’—› ÌﬁÊ„ »«·Œ’„ „‰ «·„Œ“‰ À„ ⁄„· ›« Ê—… ·«Õﬁ… ·… "
                  Height          =   195
                  Left            =   1620
                  RightToLeft     =   -1  'True
                  TabIndex        =   490
                  Top             =   240
                  Width           =   7635
               End
               Begin VB.CheckBox ChKautoIssueVoucher 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "⁄‰œ «‰‘«¡ ›« Ê—Â  „»Ì⁄«  Ì „ «‰‘«¡ ”‰œ ’—› «·Ì"
                  Height          =   195
                  Left            =   1620
                  RightToLeft     =   -1  'True
                  TabIndex        =   489
                  Top             =   480
                  Width           =   7635
               End
               Begin VB.OptionButton Opt_Sal 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "⁄„· ›« Ê—… „»Ì⁄«  ·Ì „ ⁄„· Œ’„ «·„Œ“‰ „‰Â« Ê«·„⁄«„·«  «·„«·Ì… "
                  Height          =   195
                  Left            =   1200
                  RightToLeft     =   -1  'True
                  TabIndex        =   91
                  Top             =   1080
                  Visible         =   0   'False
                  Width           =   5055
               End
               Begin VB.Label Label12 
                  Alignment       =   1  'Right Justify
                  BackStyle       =   0  'Transparent
                  Caption         =   "⁄œœ ‰”Œ ÿ»«⁄… «·›« Ê—…"
                  Height          =   255
                  Left            =   2760
                  RightToLeft     =   -1  'True
                  TabIndex        =   568
                  Top             =   720
                  Width           =   1695
               End
            End
            Begin MSDataListLib.DataCombo DBCboClientName 
               Height          =   315
               Left            =   2910
               TabIndex        =   116
               Top             =   150
               Width           =   4230
               _ExtentX        =   7461
               _ExtentY        =   556
               _Version        =   393216
               ListField       =   "6"
               Text            =   ""
               RightToLeft     =   -1  'True
            End
            Begin MSDataListLib.DataCombo DCboStoreName 
               Height          =   315
               Index           =   0
               Left            =   2880
               TabIndex        =   117
               Top             =   600
               Width           =   4230
               _ExtentX        =   7461
               _ExtentY        =   556
               _Version        =   393216
               Text            =   ""
               RightToLeft     =   -1  'True
            End
            Begin MSDataListLib.DataCombo DcboBox 
               Height          =   315
               Left            =   2880
               TabIndex        =   486
               Top             =   960
               Width           =   4230
               _ExtentX        =   7461
               _ExtentY        =   556
               _Version        =   393216
               Text            =   ""
               RightToLeft     =   -1  'True
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "⁄œœ ‰”Œ ÿ»«⁄Â «·›« Ê—…"
               Height          =   495
               Index           =   34
               Left            =   1200
               RightToLeft     =   -1  'True
               TabIndex        =   527
               Top             =   720
               Width           =   1335
            End
            Begin VB.Label Label19 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "«·„œ… «·«∆ „«‰Ì… »«·«Ì«„ «·«› —«÷Ì…"
               Height          =   255
               Left            =   7140
               RightToLeft     =   -1  'True
               TabIndex        =   476
               Top             =   1890
               Width           =   2175
            End
            Begin VB.Label Label16 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "ŒÌ«— «·Õœ «·«∆ „«‰Ï «·«› —«÷Ì"
               Height          =   255
               Left            =   7170
               RightToLeft     =   -1  'True
               TabIndex        =   474
               Top             =   1440
               Width           =   2085
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "«·⁄„Ì· «·«› —«÷Ì"
               Height          =   285
               Index           =   7
               Left            =   7590
               RightToLeft     =   -1  'True
               TabIndex        =   120
               Top             =   150
               Width           =   1575
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "«”„ «·„Œ“‰"
               Height          =   270
               Index           =   4
               Left            =   7515
               RightToLeft     =   -1  'True
               TabIndex        =   119
               Top             =   600
               Width           =   1620
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "«”„ «·Œ“‰…"
               Height          =   270
               Index           =   11
               Left            =   7590
               RightToLeft     =   -1  'True
               TabIndex        =   118
               Top             =   1050
               Width           =   1620
            End
         End
      End
      Begin C1SizerLibCtl.C1Elastic Ele 
         Height          =   9375
         Index           =   11
         Left            =   16305
         TabIndex        =   37
         TabStop         =   0   'False
         Top             =   45
         Width           =   12225
         _cx             =   21564
         _cy             =   16536
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
         Begin C1SizerLibCtl.C1Elastic Ele 
            Height          =   8655
            Index           =   18
            Left            =   3240
            TabIndex        =   49
            TabStop         =   0   'False
            Top             =   -120
            Width           =   9045
            _cx             =   15954
            _cy             =   15266
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
            Begin VB.CheckBox Chkbarcode 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   " Ê“Ì⁄ «·Œ’„ ÿ»ﬁ« ··ﬂ„Ì… Ê«·”⁄— "
               Height          =   285
               Index           =   220
               Left            =   1200
               RightToLeft     =   -1  'True
               TabIndex        =   639
               Top             =   6510
               Width           =   7185
            End
            Begin VB.CheckBox Chkbarcode 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "«·ﬁÌœ ›Ï «·„‘ —Ì«  ⁄·Ï ”ÿ— «·«’‰«›"
               Height          =   285
               Index           =   210
               Left            =   1200
               RightToLeft     =   -1  'True
               TabIndex        =   577
               Top             =   6150
               Width           =   7185
            End
            Begin VB.CheckBox Chkbarcode 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "⁄œ„ «·”„«Õ » ﬂ—«— «·—ﬁ„ «·ÌœÊÌ ·›« Ê—… «·„Ê—œ ﬁÌ ›« Ê—Â «·„‘ —Ì« "
               Height          =   285
               Index           =   189
               Left            =   1200
               RightToLeft     =   -1  'True
               TabIndex        =   545
               Top             =   5880
               Width           =   7185
            End
            Begin VB.CheckBox Chkbarcode 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   " Ê“Ì⁄ «·„’—Ê›«  ÿ»ﬁ« ··ﬂ„Ì«  ›ﬁÿ"
               Height          =   285
               Index           =   171
               Left            =   1200
               RightToLeft     =   -1  'True
               TabIndex        =   524
               Top             =   5640
               Width           =   7185
            End
            Begin VB.CheckBox Chkbarcode 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "«·„—œÊœ«   «·«› —«÷Ì… «Ã·…"
               Height          =   285
               Index           =   137
               Left            =   0
               RightToLeft     =   -1  'True
               TabIndex        =   481
               Top             =   360
               Width           =   2265
            End
            Begin VB.CheckBox Chkbarcode 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "›« Ê—… «·„‘ —Ì«  —»ÿ «·„Ê—œ »«·«’‰«›"
               Height          =   405
               Index           =   134
               Left            =   1200
               RightToLeft     =   -1  'True
               TabIndex        =   478
               Top             =   5040
               Width           =   7185
            End
            Begin VB.CheckBox Chkbarcode 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "«·”„«Õ » ⁄œÌ· ›Ê« Ì— «·„‘ —Ì«  «·„— »ÿ… »”‰œ«  ’—›"
               Height          =   405
               Index           =   77
               Left            =   1200
               RightToLeft     =   -1  'True
               TabIndex        =   416
               Top             =   4680
               Width           =   7185
            End
            Begin VB.CheckBox Chkbarcode 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "«· ⁄«„· »ÿ—ﬁ œ›⁄ „ ⁄œœ… ›Ì ›« Ê—… «·‘—«¡"
               Height          =   285
               Index           =   50
               Left            =   1200
               RightToLeft     =   -1  'True
               TabIndex        =   378
               Top             =   4440
               Width           =   7185
            End
            Begin VB.CheckBox Chkbarcode 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "«·›« Ê—… «·«› —«÷Ì… «Ã·…"
               Height          =   285
               Index           =   46
               Left            =   240
               RightToLeft     =   -1  'True
               TabIndex        =   374
               Top             =   120
               Width           =   2025
            End
            Begin VB.CheckBox Chkbarcode 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "«·ﬂ„Ì… «·„ÿ·Ê»… ›Ì «·ÿ·» «·œ«Œ·Ì  ŸÂ— ﬂ«„·… ›Ì ÿ·» «·‘—«¡"
               Height          =   285
               Index           =   27
               Left            =   1200
               RightToLeft     =   -1  'True
               TabIndex        =   329
               Top             =   5400
               Width           =   7185
            End
            Begin VB.CheckBox Chkbarcode 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "«„— «·‘—«¡ Ì‰ Ã ﬁÌœ ‰Ÿ«„Ì"
               Height          =   285
               Index           =   26
               Left            =   1200
               RightToLeft     =   -1  'True
               TabIndex        =   328
               Top             =   3840
               Width           =   7185
            End
            Begin VB.CheckBox Chkbarcode 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "«· ⁄«„· »« ›«ﬁÌ«  «·„Ê—œÌ‰"
               Height          =   285
               Index           =   25
               Left            =   1200
               RightToLeft     =   -1  'True
               TabIndex        =   325
               Top             =   3480
               Width           =   7185
            End
            Begin VB.CheckBox Chkbarcode 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "«·„‘ —Ì«  »œÊ‰ ⁄·«„«  ⁄‘—Ì…"
               Height          =   285
               Index           =   21
               Left            =   60
               RightToLeft     =   -1  'True
               TabIndex        =   321
               Top             =   3120
               Width           =   8325
            End
            Begin VB.Frame Fra 
               BackColor       =   &H00E2E9E9&
               Caption         =   "⁄‰œ  ”ÃÌ· ›« Ê—… ÃœÌœ…"
               ForeColor       =   &H000000FF&
               Height          =   1065
               Index           =   1
               Left            =   1080
               RightToLeft     =   -1  'True
               TabIndex        =   54
               Top             =   1050
               Width           =   7545
               Begin VB.OptionButton Opt 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   " «—ÌŒ «·ÃÂ«“ «·”Ì—›— ( „ «Õ ›ﬁÿ ›Ï Õ«·… ⁄„·  «·‘»ﬂ…)"
                  Height          =   255
                  Index           =   3
                  Left            =   75
                  RightToLeft     =   -1  'True
                  TabIndex        =   57
                  Top             =   750
                  Width           =   7035
               End
               Begin VB.OptionButton Opt 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   " «—ÌŒ «Œ— ›« Ê—… ‘—«¡ „”Ã·… ﬁ»·Â«"
                  Height          =   255
                  Index           =   4
                  Left            =   120
                  RightToLeft     =   -1  'True
                  TabIndex        =   56
                  Top             =   480
                  Width           =   7005
               End
               Begin VB.OptionButton Opt 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   " «—ÌŒ «·›« Ê—… ÂÊ «· «—ÌŒ «·Õ«·Ï ›Ï «·ÃÂ«“ "
                  ForeColor       =   &H00000000&
                  Height          =   255
                  Index           =   5
                  Left            =   150
                  RightToLeft     =   -1  'True
                  TabIndex        =   55
                  Top             =   240
                  Value           =   -1  'True
                  Width           =   6975
               End
            End
            Begin VB.Frame Frame2 
               BackColor       =   &H00E2E9E9&
               Caption         =   "«·„‘ —Ì«  Ê«–‰ «·«÷«›… "
               ForeColor       =   &H000000FF&
               Height          =   735
               Left            =   990
               RightToLeft     =   -1  'True
               TabIndex        =   50
               Top             =   2280
               Width           =   7665
               Begin VB.CheckBox ChKautoReseiveVoucher 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "⁄‰œ «‰‘«¡ ›« Ê—Â  „‘ —Ì«  Ì „ «‰‘«¡ ”‰œ «÷«›Â «·Ì"
                  Height          =   195
                  Left            =   450
                  RightToLeft     =   -1  'True
                  TabIndex        =   53
                  Top             =   480
                  Width           =   6795
               End
               Begin VB.OptionButton Opt_OrderInpo 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "⁄„· «–‰ «÷«›… ÌﬁÊ„ »«·«÷«›… ›Ï «·„Œ“‰ À„ ⁄„· ›« Ê—… ·«Õﬁ… ·… "
                  Height          =   195
                  Left            =   270
                  RightToLeft     =   -1  'True
                  TabIndex        =   52
                  Top             =   240
                  Width           =   6975
               End
               Begin VB.OptionButton Opt_Bey 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "⁄„· ›« Ê—… „‘ —Ì«  ·Ì „ «·«÷«›… ›Ï «·„Œ“‰ Êﬂ· «·„⁄«„·«  «·„«·Ì… "
                  Height          =   195
                  Left            =   480
                  RightToLeft     =   -1  'True
                  TabIndex        =   51
                  Top             =   720
                  Width           =   5055
               End
            End
            Begin MSDataListLib.DataCombo DBCboSupName 
               Height          =   315
               Left            =   3030
               TabIndex        =   58
               Top             =   150
               Width           =   3510
               _ExtentX        =   6191
               _ExtentY        =   556
               _Version        =   393216
               Style           =   2
               Text            =   ""
               RightToLeft     =   -1  'True
            End
            Begin MSDataListLib.DataCombo DCboStoreName 
               Height          =   315
               Index           =   1
               Left            =   3030
               TabIndex        =   59
               Top             =   600
               Width           =   3510
               _ExtentX        =   6191
               _ExtentY        =   556
               _Version        =   393216
               Text            =   ""
               RightToLeft     =   -1  'True
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "«·„Ê—œ «·«› —«÷Ì"
               Height          =   270
               Index           =   6
               Left            =   7245
               RightToLeft     =   -1  'True
               TabIndex        =   61
               Top             =   150
               Width           =   1380
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "«”„ «·„Œ“‰"
               Height          =   270
               Index           =   0
               Left            =   7245
               RightToLeft     =   -1  'True
               TabIndex        =   60
               Top             =   600
               Width           =   1260
            End
         End
      End
      Begin C1SizerLibCtl.C1Elastic Ele 
         Height          =   9375
         Index           =   1
         Left            =   16605
         TabIndex        =   38
         TabStop         =   0   'False
         Top             =   45
         Width           =   12225
         _cx             =   21564
         _cy             =   16536
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
         Begin VB.Frame Frame45 
            Caption         =   "ŒÌ«—«  «·„⁄«„·«  «·„«·Ì…"
            Height          =   9015
            Left            =   1980
            RightToLeft     =   -1  'True
            TabIndex        =   230
            Top             =   0
            Width           =   10035
            Begin VB.CheckBox Chkbarcode 
               Alignment       =   1  'Right Justify
               Caption         =   "«Œ›«¡ «·—’Ìœ „‰  ’›Ì… «·⁄Âœ…"
               Height          =   405
               Index           =   219
               Left            =   4710
               RightToLeft     =   -1  'True
               TabIndex        =   637
               Top             =   8550
               Width           =   4755
            End
            Begin VB.CheckBox Chkbarcode 
               Alignment       =   1  'Right Justify
               Caption         =   "⁄œ„ «·”„«Õ » ﬂ—«— —ﬁ„ «·ÕÊ«·Â ›Ì ”‰œ«  «·ﬁ»÷"
               Height          =   285
               Index           =   188
               Left            =   0
               RightToLeft     =   -1  'True
               TabIndex        =   544
               Top             =   8280
               Width           =   4005
            End
            Begin VB.CheckBox Chkbarcode 
               Alignment       =   1  'Right Justify
               Caption         =   "⁄œ„ „—«Ã⁄Â « “«‰ «·ﬁÌœ »«·‰”»… ··ÃÊ«—Ì"
               Height          =   285
               Index           =   174
               Left            =   0
               RightToLeft     =   -1  'True
               TabIndex        =   529
               Top             =   7920
               Width           =   4005
            End
            Begin VB.CheckBox Chkbarcode 
               Alignment       =   1  'Right Justify
               Caption         =   " Ê“Ì⁄ «·Ì ··Õ”«»«  «·Ã«—ÌÂ »ﬁÌœ «· ”ÊÌ… «·ÌœÊÌ"
               Height          =   285
               Index           =   173
               Left            =   0
               RightToLeft     =   -1  'True
               TabIndex        =   528
               Top             =   7560
               Width           =   4005
            End
            Begin VB.CheckBox Chkbarcode 
               Alignment       =   1  'Right Justify
               Caption         =   "⁄œ„  Ê“Ì⁄ «·ÃÊ«—Ì »”‰œ ’—› «·„œ›Ê⁄« "
               Height          =   285
               Index           =   167
               Left            =   0
               RightToLeft     =   -1  'True
               TabIndex        =   521
               Top             =   7200
               Width           =   4005
            End
            Begin VB.CheckBox Chkbarcode 
               Alignment       =   1  'Right Justify
               Caption         =   "›Ï ”‰œ«  ’—›  Õ·Ì·Ì «·„’—Ê›«  œ„Ã «·ﬁÌ„Â «·„÷«›…"
               Height          =   285
               Index           =   161
               Left            =   60
               RightToLeft     =   -1  'True
               TabIndex        =   511
               Top             =   6810
               Width           =   4005
            End
            Begin VB.CheckBox Chkbarcode 
               Alignment       =   1  'Right Justify
               Caption         =   " ⁄«„· ﬁÌœ «· ”ÊÌ… »«· FIFO"
               Height          =   285
               Index           =   124
               Left            =   690
               RightToLeft     =   -1  'True
               TabIndex        =   464
               Top             =   6480
               Visible         =   0   'False
               Width           =   3375
            End
            Begin VB.CheckBox Chkbarcode 
               Alignment       =   1  'Right Justify
               Caption         =   " ⁄«„· «·Œ’„ «·„”„ÊÕ »Â »«· FIFO"
               Height          =   285
               Index           =   123
               Left            =   600
               RightToLeft     =   -1  'True
               TabIndex        =   463
               Top             =   6120
               Width           =   3465
            End
            Begin VB.CheckBox Chkbarcode 
               Alignment       =   1  'Right Justify
               Caption         =   "⁄œ„ «Õ ”«» «·ﬁÌ„… «·„÷«›… ›Ï «·„⁄«„·«  «·„«·Ì…"
               Height          =   405
               Index           =   107
               Left            =   480
               RightToLeft     =   -1  'True
               TabIndex        =   447
               Top             =   5730
               Width           =   3585
            End
            Begin VB.CheckBox Chkbarcode 
               Alignment       =   1  'Right Justify
               Caption         =   "ﬁÌœ  ›’Ì·Ì ›Ì  ’›Ì… «·⁄Âœ…"
               Height          =   285
               Index           =   83
               Left            =   540
               RightToLeft     =   -1  'True
               TabIndex        =   423
               Top             =   5400
               Width           =   3525
            End
            Begin VB.CheckBox Chkbarcode 
               Alignment       =   1  'Right Justify
               Caption         =   "—»ÿ «·„” Œœ„Ì‰ »ÿ—ﬁ «·œ›⁄"
               Height          =   285
               Index           =   65
               Left            =   2640
               RightToLeft     =   -1  'True
               TabIndex        =   402
               Top             =   8280
               Width           =   6825
            End
            Begin VB.CheckBox Chkbarcode 
               Alignment       =   1  'Right Justify
               Caption         =   "«· ⁄«„· »ÿ—ﬁ œ›⁄ „ ⁄œœ… ›Ì «·„ﬁ»Ê÷« "
               Height          =   195
               Index           =   59
               Left            =   2640
               RightToLeft     =   -1  'True
               TabIndex        =   396
               Top             =   8040
               Width           =   6825
            End
            Begin VB.CheckBox Chkbarcode 
               Alignment       =   1  'Right Justify
               Caption         =   "≈Œ›«¡ «·⁄Âœ „‰ ”‰œ «·ﬁ»÷"
               Height          =   285
               Index           =   57
               Left            =   2550
               RightToLeft     =   -1  'True
               TabIndex        =   394
               Top             =   7680
               Width           =   6915
            End
            Begin VB.CheckBox Chkbarcode 
               Alignment       =   1  'Right Justify
               Caption         =   "≈ŸÂ«— «·Õ”«»«  ÿ»ﬁ« ·›—⁄ «·„” Œœ„"
               Height          =   405
               Index           =   54
               Left            =   4680
               RightToLeft     =   -1  'True
               TabIndex        =   388
               Top             =   6480
               Width           =   4785
            End
            Begin VB.CheckBox Chkbarcode 
               Alignment       =   1  'Right Justify
               Caption         =   "≈„ﬂ«‰Ì…  ⁄œÌ· «·Õ”«»«  «·«·Ì…"
               Height          =   405
               Index           =   55
               Left            =   3750
               RightToLeft     =   -1  'True
               TabIndex        =   387
               Top             =   6840
               Width           =   5715
            End
            Begin VB.CheckBox Chkbarcode 
               Alignment       =   1  'Right Justify
               Caption         =   " Õ·Ì· «·„’—Ê› ÿ»ﬁ« ·ÃÂÂ «·’—›-›Ì ”‰œ ’—›  Õ·Ì·Ì „’—Ê›« "
               Height          =   405
               Index           =   42
               Left            =   4320
               RightToLeft     =   -1  'True
               TabIndex        =   370
               Top             =   6120
               Width           =   5145
            End
            Begin VB.ComboBox CboChasingStatus 
               Height          =   315
               Left            =   4440
               RightToLeft     =   -1  'True
               Style           =   2  'Dropdown List
               TabIndex        =   332
               Top             =   7320
               Width           =   2745
            End
            Begin VB.CheckBox Chkbarcode 
               Alignment       =   1  'Right Justify
               Caption         =   "› Õ Õ”«» ⁄Ã“ Ê“Ì«œ… ›Ì «·‰ﬁœÌ… ·ﬂ· ’‰œÊﬁ/⁄Âœ…"
               Height          =   285
               Index           =   16
               Left            =   1320
               RightToLeft     =   -1  'True
               TabIndex        =   314
               Top             =   360
               Width           =   4305
            End
            Begin VB.CheckBox Chkbarcode 
               Alignment       =   1  'Right Justify
               Caption         =   "«·”„«Õ » ⁄œÌ· ›Ê« Ì— «·„»Ì⁄«  «·„— »ÿ… »”‰œ«  ﬁ»÷"
               Height          =   405
               Index           =   3
               Left            =   4650
               RightToLeft     =   -1  'True
               TabIndex        =   304
               Top             =   5760
               Width           =   4815
            End
            Begin VB.CheckBox Chkbarcode 
               Alignment       =   1  'Right Justify
               Caption         =   "«·”„«Õ » ⁄œÌ· ”‰œ«  «·ﬁ»÷ «· Ì ⁄·ÌÂ« Õ—ﬂ«  «Ìœ«⁄« "
               Height          =   405
               Index           =   4
               Left            =   4710
               RightToLeft     =   -1  'True
               TabIndex        =   301
               Top             =   5400
               Width           =   4755
            End
            Begin VB.Frame Frame54 
               Caption         =   "«·⁄„·«¡"
               ForeColor       =   &H000000FF&
               Height          =   735
               Left            =   2010
               RightToLeft     =   -1  'True
               TabIndex        =   280
               Top             =   2160
               Width           =   7605
               Begin VB.CheckBox chkIsCreateOpenBalnceMan 
                  Alignment       =   1  'Right Justify
                  Caption         =   "«‰‘«¡ —’Ìœ «›  «ÕÏ ··⁄„·«¡ Ê«·„Ê—œÌ‰ „»«‘—…"
                  ForeColor       =   &H00000000&
                  Height          =   405
                  Left            =   1680
                  RightToLeft     =   -1  'True
                  TabIndex        =   643
                  Top             =   390
                  Width           =   5715
               End
               Begin VB.CheckBox Chkbarcode 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "«·⁄„Ì· Ì‰‘√ «—»⁄ Õ”«»« "
                  Height          =   405
                  Index           =   208
                  Left            =   0
                  RightToLeft     =   -1  'True
                  TabIndex        =   576
                  Top             =   90
                  Width           =   2475
               End
               Begin VB.CheckBox chkCustomerhavethreeAccounts 
                  Alignment       =   1  'Right Justify
                  Caption         =   "«· ⁄«„· »‘Ìﬂ«   Õ  «· Õ’Ì· ··⁄„·«¡ Ê «·„œ›Ê⁄«   «·„ﬁœ„… ··⁄„·«¡"
                  ForeColor       =   &H00000000&
                  Height          =   405
                  Left            =   1680
                  RightToLeft     =   -1  'True
                  TabIndex        =   281
                  Top             =   150
                  Width           =   5715
               End
            End
            Begin VB.CheckBox Chk1 
               Alignment       =   1  'Right Justify
               Caption         =   " «·”Õ» ⁄·Ï «·„ﬂ‘Ê› „‰ «·Œ“‰…"
               Height          =   285
               Left            =   6960
               RightToLeft     =   -1  'True
               TabIndex        =   241
               Top             =   360
               Width           =   2505
            End
            Begin VB.Frame Frame36 
               Caption         =   "«·»‰Êﬂ"
               ForeColor       =   &H000000FF&
               Height          =   1455
               Left            =   2010
               RightToLeft     =   -1  'True
               TabIndex        =   237
               Top             =   720
               Width           =   7605
               Begin VB.CheckBox chkIsCheque 
                  Alignment       =   1  'Right Justify
                  Caption         =   "›Ï Õ«·… «” Œœ«„ ‰Ê⁄ «·ﬁ»÷ «Ê «·œ›⁄ (‘Ìﬂ ) ·« Ì‰‘√ ﬁÌœ"
                  Height          =   405
                  Left            =   180
                  RightToLeft     =   -1  'True
                  TabIndex        =   578
                  Top             =   1020
                  Width           =   7245
               End
               Begin VB.CheckBox ChkChequeBox 
                  Alignment       =   1  'Right Justify
                  Caption         =   "Õ«›Ÿ… «·‘Ìﬂ« "
                  Height          =   405
                  Left            =   300
                  RightToLeft     =   -1  'True
                  TabIndex        =   240
                  Top             =   450
                  Width           =   7125
               End
               Begin VB.CheckBox ChkBankComm 
                  Alignment       =   1  'Right Justify
                  Caption         =   "«· ⁄«„· »«·⁄„Ê·Â «·»‰ﬂÌÂ"
                  Height          =   525
                  Left            =   210
                  RightToLeft     =   -1  'True
                  TabIndex        =   239
                  Top             =   690
                  Width           =   7215
               End
               Begin VB.CheckBox Chk3 
                  Alignment       =   1  'Right Justify
                  Caption         =   "«· ⁄«„· „⁄ Õ”«»«  «·‘Ìﬂ«  «·„ƒÃ·… Ê«·‘Ìﬂ«   Õ  «· Õ’Ì·"
                  Height          =   405
                  Left            =   180
                  RightToLeft     =   -1  'True
                  TabIndex        =   238
                  Top             =   150
                  Width           =   7245
               End
            End
            Begin VB.Frame Frame42 
               Caption         =   " ﬂÊÌœ ”‰œ«  «· ÕÊÌ·«  «·„«·Ì…"
               ForeColor       =   &H000000FF&
               Height          =   735
               Left            =   2070
               RightToLeft     =   -1  'True
               TabIndex        =   235
               Top             =   3960
               Width           =   7545
               Begin VB.CheckBox ChkExpensesCoding2 
                  Alignment       =   1  'Right Justify
                  Caption         =   "”‰œ«  «· ÕÊÌ· Ê ”‰œ«  «·’—› ‰›” «·”Ì—Ì«·"
                  ForeColor       =   &H00000000&
                  Height          =   405
                  Left            =   600
                  RightToLeft     =   -1  'True
                  TabIndex        =   236
                  Top             =   240
                  Width           =   6735
               End
            End
            Begin VB.Frame Frame55 
               Caption         =   " ﬂÊÌœ ”‰œ«  «·’—›"
               ForeColor       =   &H000000FF&
               Height          =   975
               Left            =   2010
               RightToLeft     =   -1  'True
               TabIndex        =   233
               Top             =   2880
               Width           =   7605
               Begin VB.CheckBox Chkbarcode 
                  Alignment       =   1  'Right Justify
                  Caption         =   "”‰œ ’—› «·„œ›Ê⁄«  Ê”‰œ ’—› «·‘Ìﬂ«  ÿ»«⁄Â „Œ ·›…"
                  Height          =   285
                  Index           =   29
                  Left            =   90
                  RightToLeft     =   -1  'True
                  TabIndex        =   338
                  Top             =   600
                  Width           =   7305
               End
               Begin VB.CheckBox ChkExpensesCoding 
                  Alignment       =   1  'Right Justify
                  Caption         =   "”‰œ«  «·’—› Ê«·„œ›Ê⁄«  ‰›” «·”Ì—Ì«·"
                  ForeColor       =   &H00000000&
                  Height          =   405
                  Left            =   180
                  RightToLeft     =   -1  'True
                  TabIndex        =   234
                  Top             =   240
                  Width           =   7215
               End
            End
            Begin VB.Frame Frame56 
               Caption         =   " ﬂÊÌœ ”‰œ«  «·«ﬁ”«ÿ"
               ForeColor       =   &H000000FF&
               Height          =   735
               Left            =   2040
               RightToLeft     =   -1  'True
               TabIndex        =   231
               Top             =   4680
               Width           =   7575
               Begin VB.CheckBox chkInstallmntsvchrCoding 
                  Alignment       =   1  'Right Justify
                  Caption         =   "”‰œ«  «·«ﬁ”«ÿ Ê «·„ﬁ»Ê÷«  ‰›” «·”Ì—Ì«·"
                  ForeColor       =   &H00000000&
                  Height          =   405
                  Left            =   1080
                  RightToLeft     =   -1  'True
                  TabIndex        =   232
                  Top             =   240
                  Width           =   6285
               End
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               BackStyle       =   0  'Transparent
               Caption         =   "«·Õ«·Â «·«› —«÷Ì… ›Ì ”‰œ «·ﬁ»÷"
               Height          =   375
               Index           =   25
               Left            =   7080
               RightToLeft     =   -1  'True
               TabIndex        =   331
               Top             =   7320
               Width           =   2325
            End
         End
      End
      Begin C1SizerLibCtl.C1Elastic Ele 
         Height          =   9375
         Index           =   2
         Left            =   16905
         TabIndex        =   39
         TabStop         =   0   'False
         Top             =   45
         Width           =   12225
         _cx             =   21564
         _cy             =   16536
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
         Begin VB.Frame Frame9 
            BackColor       =   &H00E2E9E9&
            Caption         =   "ŒÌ«—«  ‘∆Ê‰ «·„ÊŸ›Ì‰"
            Height          =   9135
            Left            =   1770
            RightToLeft     =   -1  'True
            TabIndex        =   80
            Top             =   0
            Width           =   10245
            Begin VB.CheckBox Chkbarcode 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "«‰‘«¡ «·Õ”«»«  Ê›ﬁ« ··«œ«—…"
               Height          =   405
               Index           =   217
               Left            =   1230
               RightToLeft     =   -1  'True
               TabIndex        =   635
               Top             =   8490
               Width           =   7185
            End
            Begin VB.CheckBox Chkbarcode 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   " Ê“Ì⁄ ﬁÌœ «·„Œ’’«  ÿ»ﬁ« ··„⁄œ« "
               Height          =   405
               Index           =   165
               Left            =   1200
               RightToLeft     =   -1  'True
               TabIndex        =   519
               Top             =   8160
               Width           =   7185
            End
            Begin VB.CheckBox Chkbarcode 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "›Ì „”Ì— «·—Ê« » ·«»œ „‰ «Œ Ì«— ›—⁄"
               ForeColor       =   &H000000FF&
               Height          =   405
               Index           =   157
               Left            =   5520
               RightToLeft     =   -1  'True
               TabIndex        =   507
               Top             =   360
               Width           =   2985
            End
            Begin VB.CheckBox Chkbarcode 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "ﬁÌœ «·„Œ’’«  ÿ»ﬁ« ··«œ«—« "
               Height          =   405
               Index           =   129
               Left            =   1200
               RightToLeft     =   -1  'True
               TabIndex        =   470
               Top             =   7800
               Width           =   7185
            End
            Begin VB.TextBox TxtEmpSalaryDigts 
               Alignment       =   2  'Center
               Height          =   285
               Left            =   4920
               MaxLength       =   1
               RightToLeft     =   -1  'True
               TabIndex        =   469
               Top             =   7200
               Width           =   375
            End
            Begin VB.CheckBox Chkbarcode 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "«ŸÂ«— —’Ìœ –„„ «·„ÊŸ› ›Ì «·„”Ì—"
               Height          =   405
               Index           =   125
               Left            =   1200
               RightToLeft     =   -1  'True
               TabIndex        =   465
               Top             =   7440
               Width           =   7185
            End
            Begin VB.CheckBox Chkbarcode 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "≈‰‘«¡ «·ﬁÌœ ›Ì «·«Ã«“… «·⁄«—÷…"
               Height          =   405
               Index           =   101
               Left            =   1200
               RightToLeft     =   -1  'True
               TabIndex        =   441
               Top             =   6720
               Width           =   7185
            End
            Begin VB.CheckBox Chkbarcode 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "«·›—⁄ «·“«„Ì ›Ì «·„”Ì—"
               Height          =   405
               Index           =   97
               Left            =   1200
               RightToLeft     =   -1  'True
               TabIndex        =   437
               Top             =   6360
               Width           =   7185
            End
            Begin VB.CheckBox Chkbarcode 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "ﬁÌœ «·—Ê« »  Õ·Ì·Ì ÿ»ﬁ« ··„⁄œ…"
               Height          =   285
               Index           =   86
               Left            =   390
               RightToLeft     =   -1  'True
               TabIndex        =   426
               Top             =   6120
               Width           =   7995
            End
            Begin VB.CheckBox Chkbarcode 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "ﬁÌœ «·—Ê« » ÿ»ﬁ« ··«œ«—…"
               Height          =   285
               Index           =   79
               Left            =   3360
               RightToLeft     =   -1  'True
               TabIndex        =   418
               Top             =   5760
               Width           =   5025
            End
            Begin VB.CheckBox Chkbarcode 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "‘«‘Â „” Õﬁ«  «·«Ã«“…  ·« ‰ŸÂ— «·—Ê« » «· Ì ·„  ”œœ "
               Height          =   285
               Index           =   75
               Left            =   360
               RightToLeft     =   -1  'True
               TabIndex        =   414
               Top             =   5400
               Width           =   8025
            End
            Begin VB.CheckBox Chkbarcode 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "‰Â«Ì… «·Œœ„… «Õ ”«» 5 «·”‰Ê«  "
               Height          =   285
               Index           =   62
               Left            =   4440
               RightToLeft     =   -1  'True
               TabIndex        =   399
               Top             =   5160
               Width           =   3945
            End
            Begin VB.CheckBox Chkbarcode 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "‘«‘Â «·«—’œ… «·«›  «ÕÌ… ··„ÊŸ›Ì‰  ⁄—÷ ﬂ· «·„ÊŸ›Ì‰"
               Height          =   285
               Index           =   60
               Left            =   960
               RightToLeft     =   -1  'True
               TabIndex        =   397
               Top             =   3720
               Width           =   7425
            End
            Begin VB.CheckBox Chkbarcode 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "≈ﬁ›«· «·„”Ì—"
               Height          =   285
               Index           =   58
               Left            =   4440
               RightToLeft     =   -1  'True
               TabIndex        =   395
               Top             =   4800
               Width           =   3945
            End
            Begin VB.CheckBox Chkbarcode 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "«·⁄„· »‰Ÿ«„ «Ê· »’„Â Ê«Œ— »’„…"
               Height          =   285
               Index           =   36
               Left            =   4440
               RightToLeft     =   -1  'True
               TabIndex        =   362
               Top             =   4440
               Width           =   3945
            End
            Begin VB.CheckBox Chkbarcode 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "«·”„«Õ  » ⁄œÌ· Õ«·Â «·„ÊŸ› „‰ „·› «·„ÊŸ›Ì‰"
               Height          =   285
               Index           =   6
               Left            =   4440
               RightToLeft     =   -1  'True
               TabIndex        =   303
               Top             =   4080
               Width           =   3945
            End
            Begin VB.Frame Frame35 
               BackColor       =   &H00E2E9E9&
               Caption         =   "‘∆Ê‰ «·„ÊŸ›Ì‰"
               ForeColor       =   &H000000FF&
               Height          =   2055
               Left            =   3120
               RightToLeft     =   -1  'True
               TabIndex        =   84
               Top             =   720
               Width           =   5415
               Begin VB.CheckBox Chkbarcode 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "Ì „ ‰ﬁ· „Œ’’«  «·„ÊŸ› ⁄‰œ ⁄„·Ì… «·‰ﬁ·"
                  Height          =   285
                  Index           =   35
                  Left            =   1320
                  RightToLeft     =   -1  'True
                  TabIndex        =   361
                  Top             =   1680
                  Width           =   3945
               End
               Begin VB.CheckBox Chkbarcode 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "ﬁÌœ «·—Ê« » ﬁÌœ »”Ìÿ ⁄·Ì Õ”«» Ê«Õœ"
                  Height          =   285
                  Index           =   30
                  Left            =   1320
                  RightToLeft     =   -1  'True
                  TabIndex        =   344
                  Top             =   1080
                  Width           =   3945
               End
               Begin VB.CheckBox Chkbarcode 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "«· ⁄«„· »„Œ’’«  «· –«ﬂ—"
                  Height          =   285
                  Index           =   24
                  Left            =   1320
                  RightToLeft     =   -1  'True
                  TabIndex        =   324
                  Top             =   840
                  Width           =   3945
               End
               Begin VB.TextBox TxtEmpComponentDigts 
                  Alignment       =   2  'Center
                  Height          =   285
                  Left            =   600
                  MaxLength       =   1
                  RightToLeft     =   -1  'True
                  TabIndex        =   87
                  Top             =   1440
                  Width           =   375
               End
               Begin VB.CheckBox chkMonthIs30days 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "«·‘Â— ÌÕ”» ⁄·Ï «”«” 30 ÌÊ„"
                  Height          =   405
                  Left            =   2280
                  RightToLeft     =   -1  'True
                  TabIndex        =   86
                  Top             =   240
                  Width           =   2985
               End
               Begin VB.CheckBox Chkemployeeaccounts 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "⁄‰œ «‰‘«¡ „ÊŸ› ÃœÌœ Ì „ —»ÿÂ »«·Õ”«»« "
                  Height          =   285
                  Left            =   720
                  RightToLeft     =   -1  'True
                  TabIndex        =   85
                  Top             =   600
                  Width           =   4545
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "⁄œœ «·«—ﬁ«„ »⁄œ «·⁄·«„… «·⁄‘—Ì… ··„›—œ«  «·„ €Ì—… «Ì«„ Ê”«⁄« "
                  Height          =   255
                  Index           =   38
                  Left            =   840
                  RightToLeft     =   -1  'True
                  TabIndex        =   88
                  Top             =   1440
                  Width           =   4455
               End
            End
            Begin VB.Frame Frame48 
               BackColor       =   &H00E2E9E9&
               Caption         =   "ÿ—Ìﬁ… Õ” «» «·„Œ’’"
               ForeColor       =   &H000000FF&
               Height          =   615
               Left            =   3240
               RightToLeft     =   -1  'True
               TabIndex        =   81
               Top             =   2760
               Width           =   5415
               Begin VB.OptionButton ChkEmpRes 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "Õ”«» «·„Œ’’«  ”‰ÊÌ«"
                  ForeColor       =   &H00000000&
                  Height          =   255
                  Index           =   1
                  Left            =   480
                  RightToLeft     =   -1  'True
                  TabIndex        =   83
                  Top             =   240
                  Width           =   2175
               End
               Begin VB.OptionButton ChkEmpRes 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "Õ”«» «·„Œ’’«  ‘Â—Ì«"
                  ForeColor       =   &H00000000&
                  Height          =   255
                  Index           =   0
                  Left            =   3000
                  RightToLeft     =   -1  'True
                  TabIndex        =   82
                  Top             =   240
                  Width           =   2175
               End
            End
            Begin VB.Label Label13 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "⁄œœ «·«—ﬁ«„ »⁄œ «·⁄·«„… «·⁄‘—Ì… ·„”Ì— «·—Ê« »"
               Height          =   255
               Left            =   5400
               RightToLeft     =   -1  'True
               TabIndex        =   461
               Top             =   7200
               Width           =   3015
            End
         End
      End
      Begin C1SizerLibCtl.C1Elastic Ele 
         Height          =   9375
         Index           =   3
         Left            =   17205
         TabIndex        =   40
         TabStop         =   0   'False
         Top             =   45
         Width           =   12225
         _cx             =   21564
         _cy             =   16536
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
         Begin C1SizerLibCtl.C1Elastic Ele 
            Height          =   8655
            Index           =   19
            Left            =   2640
            TabIndex        =   62
            TabStop         =   0   'False
            Top             =   0
            Width           =   8805
            _cx             =   15531
            _cy             =   15266
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
            Begin VB.CheckBox Chkbarcode 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "›Ï ”‰œ «· Ã„Ì⁄ «·ﬂ„Ì… ÿ»ﬁ« ··”„ﬂ"
               Height          =   285
               Index           =   198
               Left            =   840
               RightToLeft     =   -1  'True
               TabIndex        =   562
               Top             =   6000
               Width           =   7635
            End
            Begin VB.CheckBox Chkbarcode 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   " ⁄œœ «·«’‰«› ›Ï ”‰œ «· Ã„Ì⁄"
               ForeColor       =   &H000000FF&
               Height          =   285
               Index           =   132
               Left            =   120
               RightToLeft     =   -1  'True
               TabIndex        =   554
               Top             =   5640
               Width           =   8325
            End
            Begin VB.CheckBox Chkbarcode 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "Õ«·… «·œ›⁄ «Ã· ›Ï ”‰œ «· Ã„Ì⁄"
               Height          =   285
               Index           =   182
               Left            =   4890
               RightToLeft     =   -1  'True
               TabIndex        =   537
               Top             =   5400
               Width           =   3615
            End
            Begin VB.CheckBox Chkbarcode 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "«· ﬂ·›… »‰«¡« ⁄·Ï «„— «·«‰ «Ã Õ«·… ⁄—÷ «·”⁄—"
               Height          =   285
               Index           =   179
               Left            =   4890
               RightToLeft     =   -1  'True
               TabIndex        =   534
               Top             =   5040
               Width           =   3615
            End
            Begin VB.CheckBox Chkbarcode 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "«· ÕÊÌ· ·« Ì‰‘√ ›« Ê—… ›Ï ”‰œ «· Ã„Ì⁄"
               Height          =   285
               Index           =   177
               Left            =   4890
               RightToLeft     =   -1  'True
               TabIndex        =   532
               Top             =   4680
               Width           =   3615
            End
            Begin VB.CheckBox Chkbarcode 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "«· ﬂ·›… ›Ï «„— «·«‰ «Ã »‰«¡« ⁄·Ï ”‰œ «·’—›"
               Height          =   285
               Index           =   176
               Left            =   4890
               RightToLeft     =   -1  'True
               TabIndex        =   531
               Top             =   4320
               Width           =   3615
            End
            Begin VB.CheckBox Chkbarcode 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   " Ã„Ì⁄ «·«Ê«„— »‰«¡« ⁄·Ï «Ê«„— «· Ã„Ì⁄"
               Height          =   285
               Index           =   153
               Left            =   60
               RightToLeft     =   -1  'True
               TabIndex        =   503
               Top             =   3990
               Width           =   8445
            End
            Begin VB.CheckBox Chkbarcode 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "«Œ›«¡ «·Œ«‰«  «· ›’Ì·Ì… ›Ï ”‰œ «· Ã„Ì⁄"
               Height          =   285
               Index           =   147
               Left            =   60
               RightToLeft     =   -1  'True
               TabIndex        =   497
               Top             =   3660
               Width           =   8445
            End
            Begin VB.CheckBox Chkbarcode 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "›« Ê—… «·»Ì⁄ ›Ï ”‰œ «· Ã„Ì⁄ ·« ‰‘√ ”‰œ ’—›"
               Height          =   285
               Index           =   144
               Left            =   60
               RightToLeft     =   -1  'True
               TabIndex        =   495
               Top             =   3000
               Width           =   8445
            End
            Begin VB.CheckBox Chkbarcode 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "”‰œ «· Ã„Ì⁄ ·«Ì‰‘√ ”‰œ ’—›"
               Height          =   285
               Index           =   143
               Left            =   60
               RightToLeft     =   -1  'True
               TabIndex        =   494
               Top             =   2700
               Width           =   8445
            End
            Begin VB.CheckBox Chkbarcode 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "«·»« ‘ ÌﬁÊ„ »⁄„· ⁄œœ «Ê«„— «‰ «Ã"
               Height          =   285
               Index           =   133
               Left            =   60
               RightToLeft     =   -1  'True
               TabIndex        =   477
               Top             =   2400
               Width           =   8445
            End
            Begin VB.CheckBox Chkbarcode 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "›Ì ”‰œ ’—› «·„Ê«œ «·Œ«„ Ì „ ’—› «„— «·«‰ «Ã „—Â Ê«ÕœÂ"
               Height          =   285
               Index           =   78
               Left            =   540
               RightToLeft     =   -1  'True
               TabIndex        =   417
               Top             =   2160
               Width           =   7965
            End
            Begin VB.CheckBox Chkbarcode 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "›Ì «„— «·«‰ «Ã ’—› „Ê«œ «·Œ«„ ÿ»ﬁ« ··„ﬂ”"
               Height          =   285
               Index           =   70
               Left            =   870
               RightToLeft     =   -1  'True
               TabIndex        =   409
               Top             =   1920
               Width           =   7635
            End
            Begin VB.Frame Frame32 
               BackColor       =   &H00E2E9E9&
               Caption         =   "”Ì«”Â «·«‰ «Ã"
               ForeColor       =   &H000000FF&
               Height          =   615
               Left            =   2640
               RightToLeft     =   -1  'True
               TabIndex        =   66
               Top             =   240
               Width           =   5895
               Begin VB.CheckBox ChkTypicalProduction 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "Ì „ «· ⁄«„· »«·«‰ «Ã «·‰„ÿÌ"
                  ForeColor       =   &H00000000&
                  Height          =   285
                  Left            =   2760
                  RightToLeft     =   -1  'True
                  TabIndex        =   67
                  Top             =   240
                  Width           =   2745
               End
            End
            Begin VB.Frame Fra 
               BackColor       =   &H00E2E9E9&
               Caption         =   "«· ﬂ«·Ì› €Ì— «·„»«‘—…"
               ForeColor       =   &H000000FF&
               Height          =   945
               Index           =   6
               Left            =   1650
               RightToLeft     =   -1  'True
               TabIndex        =   63
               Top             =   960
               Width           =   6885
               Begin VB.CheckBox chkExpProduction 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "„’«—Ì› Œÿ «·«‰ «Ã"
                  ForeColor       =   &H00000000&
                  Height          =   285
                  Left            =   1770
                  RightToLeft     =   -1  'True
                  TabIndex        =   484
                  Top             =   570
                  Width           =   1725
               End
               Begin VB.CheckBox chkItemProduction 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "„Ê«œ"
                  ForeColor       =   &H00000000&
                  Height          =   285
                  Left            =   3870
                  RightToLeft     =   -1  'True
                  TabIndex        =   483
                  Top             =   600
                  Width           =   1185
               End
               Begin VB.CheckBox chkEmpProduction 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "⁄„«·…"
                  ForeColor       =   &H00000000&
                  Height          =   285
                  Left            =   5580
                  RightToLeft     =   -1  'True
                  TabIndex        =   482
                  Top             =   570
                  Width           =   855
               End
               Begin VB.CheckBox ChkAllowIndirectCost 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   " Õ„Ì· ‰”»… À«» … ⁄·Ï «„— «·«‰ «Ã"
                  ForeColor       =   &H00000000&
                  Height          =   285
                  Left            =   2760
                  RightToLeft     =   -1  'True
                  TabIndex        =   65
                  Top             =   240
                  Width           =   2745
               End
               Begin VB.TextBox TxtIndirectCostPercentage 
                  Alignment       =   2  'Center
                  Height          =   345
                  Left            =   1320
                  RightToLeft     =   -1  'True
                  TabIndex        =   64
                  Tag             =   "«·⁄œœ «·„ Êﬁ⁄ ··«—ﬁ«„ ›Ì «·ﬁÌœ Ê÷·ﬂ · ”ÂÌ· «· — Ì» „À«· 2011010001 Â‰«  „ «Œ Ì«— 3 ·–·ﬂ ŸÂ— 001"
                  Top             =   240
                  Width           =   615
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "Õœœ «·‰”»…"
                  Height          =   375
                  Index           =   36
                  Left            =   1800
                  RightToLeft     =   -1  'True
                  TabIndex        =   551
                  Top             =   240
                  Width           =   1005
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "%"
                  Height          =   375
                  Index           =   35
                  Left            =   240
                  RightToLeft     =   -1  'True
                  TabIndex        =   550
                  Top             =   240
                  Width           =   1005
               End
            End
         End
      End
      Begin C1SizerLibCtl.C1Elastic Ele 
         Height          =   9375
         Index           =   4
         Left            =   17505
         TabIndex        =   41
         TabStop         =   0   'False
         Top             =   45
         Width           =   12225
         _cx             =   21564
         _cy             =   16536
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
         Begin VB.Frame Frame60 
            Height          =   8655
            Left            =   2640
            RightToLeft     =   -1  'True
            TabIndex        =   211
            Top             =   0
            Width           =   8805
            Begin VB.CheckBox Chkbarcode 
               Alignment       =   1  'Right Justify
               Caption         =   "«·‰ﬁ·Ì«  »ﬂÊœ «·„œÌ‰…"
               ForeColor       =   &H000000FF&
               Height          =   285
               Index           =   222
               Left            =   90
               RightToLeft     =   -1  'True
               TabIndex        =   644
               Top             =   6600
               Width           =   8535
            End
            Begin VB.CheckBox Chkbarcode 
               Alignment       =   1  'Right Justify
               Caption         =   "—›⁄ «ÃÊ— «·Ìœ ··ÂÌ∆…"
               ForeColor       =   &H000000FF&
               Height          =   285
               Index           =   221
               Left            =   90
               RightToLeft     =   -1  'True
               TabIndex        =   640
               Top             =   6300
               Width           =   8535
            End
            Begin VB.CheckBox Chkbarcode 
               Alignment       =   1  'Right Justify
               Caption         =   "«Œ›«¡  ›«’Ì· «·›« Ê—…"
               ForeColor       =   &H000000FF&
               Height          =   285
               Index           =   211
               Left            =   90
               RightToLeft     =   -1  'True
               TabIndex        =   581
               Top             =   5940
               Width           =   8535
            End
            Begin VB.CheckBox Chkbarcode 
               Alignment       =   1  'Right Justify
               Caption         =   "«·’Ì«‰… «’‰«›"
               ForeColor       =   &H000000FF&
               Height          =   285
               Index           =   206
               Left            =   90
               RightToLeft     =   -1  'True
               TabIndex        =   572
               Top             =   5520
               Width           =   8535
            End
            Begin VB.CheckBox Chkbarcode 
               Alignment       =   1  'Right Justify
               Caption         =   "Õ”«» «·«Ì—«œ  ·ﬁ«∆Ì ›Ê ›Ê« Ì— «·‰ﬁ·Ì« "
               ForeColor       =   &H000000FF&
               Height          =   285
               Index           =   184
               Left            =   90
               RightToLeft     =   -1  'True
               TabIndex        =   539
               Top             =   5160
               Width           =   8535
            End
            Begin VB.CheckBox Chkbarcode 
               Alignment       =   1  'Right Justify
               Caption         =   "—ﬁ„ «„— «·‘€· ·« Ì” Œœ„ ·«ﬂÀ— „‰ ›« Ê—… „‘ —Ì« "
               ForeColor       =   &H000000FF&
               Height          =   285
               Index           =   181
               Left            =   90
               RightToLeft     =   -1  'True
               TabIndex        =   536
               Top             =   4890
               Width           =   8535
            End
            Begin VB.CheckBox Chkbarcode 
               Alignment       =   1  'Right Justify
               Caption         =   "—ﬁ„ «„— «·‘€· ·« Ì” Œœ„ ·«ﬂÀ— „‰ ›« Ê—… „»Ì⁄« "
               ForeColor       =   &H000000FF&
               Height          =   285
               Index           =   180
               Left            =   90
               RightToLeft     =   -1  'True
               TabIndex        =   535
               Top             =   4530
               Width           =   8535
            End
            Begin VB.CheckBox Chkbarcode 
               Alignment       =   1  'Right Justify
               Caption         =   "⁄œ„  Õ„Ì· «·„’—Ê›«  «·Ì« „‰ ‘«‘Â «·„”«›«  »Ì‰ «·„œ‰"
               ForeColor       =   &H000000FF&
               Height          =   285
               Index           =   170
               Left            =   90
               RightToLeft     =   -1  'True
               TabIndex        =   523
               Top             =   4200
               Width           =   8535
            End
            Begin VB.CheckBox Chkbarcode 
               Alignment       =   1  'Right Justify
               Caption         =   "«·”„«Õ » ﬂ—«— «·„⁄œÂ/«·”Ì«—… Õ«· «Œ ·«› «·‘«”ÌÂ"
               Height          =   285
               Index           =   164
               Left            =   450
               RightToLeft     =   -1  'True
               TabIndex        =   518
               Top             =   3930
               Width           =   8175
            End
            Begin VB.CheckBox Chkbarcode 
               Alignment       =   1  'Right Justify
               Caption         =   "Õ”«» «·«ÃÊ— «·„” Õﬁ… ÌÃÌ «·Ì« ›Ì ÃÂÂ «·’—› ··—Õ·« "
               Height          =   285
               Index           =   150
               Left            =   480
               RightToLeft     =   -1  'True
               TabIndex        =   500
               Top             =   3600
               Width           =   8145
            End
            Begin VB.CheckBox Chkbarcode 
               Alignment       =   1  'Right Justify
               Caption         =   "«·Œ’„ ›Ì ›« Ê—… «·‰ﬁ· „»«‘—… „‰ Õ”«» «·⁄„Ì·"
               Height          =   285
               Index           =   148
               Left            =   480
               RightToLeft     =   -1  'True
               TabIndex        =   498
               Top             =   3240
               Width           =   8145
            End
            Begin VB.CheckBox Chkbarcode 
               Alignment       =   1  'Right Justify
               Caption         =   "ﬁÌœ ›Ê« Ì— «·‰ﬁ· ÿ»ﬁ« ··„«·ﬂ"
               Height          =   285
               Index           =   146
               Left            =   480
               RightToLeft     =   -1  'True
               TabIndex        =   496
               Top             =   2880
               Width           =   8145
            End
            Begin VB.CheckBox Chkbarcode 
               Alignment       =   1  'Right Justify
               Caption         =   "›Ì ‘«‘… ⁄„·«¡ «·‰ﬁ· «·ﬁÌœ «Ã„«·Ì"
               Height          =   285
               Index           =   139
               Left            =   480
               RightToLeft     =   -1  'True
               TabIndex        =   487
               Top             =   2640
               Width           =   8145
            End
            Begin VB.CheckBox Chkbarcode 
               Alignment       =   1  'Right Justify
               Caption         =   "›Ì ‘«‘… ⁄„·«¡ «·‰ﬁ· «·”⁄— ⁄·Ï „” ÊÏ «·Ã—Ìœ"
               Height          =   285
               Index           =   116
               Left            =   480
               RightToLeft     =   -1  'True
               TabIndex        =   455
               Top             =   2400
               Width           =   8145
            End
            Begin VB.CheckBox Chkbarcode 
               Alignment       =   1  'Right Justify
               Caption         =   "—»ÿ «·⁄„Ì· »”Ì«—« Â ›ﬁÿ"
               Height          =   285
               Index           =   114
               Left            =   450
               RightToLeft     =   -1  'True
               TabIndex        =   453
               Top             =   2160
               Width           =   8175
            End
            Begin VB.CheckBox Chkbarcode 
               Alignment       =   1  'Right Justify
               Caption         =   " ”ÃÌ· «·—Õ·«  »‰«¡ ⁄·Ì «„—  Õ„Ì· ›ﬁÿ"
               Height          =   285
               Index           =   112
               Left            =   570
               RightToLeft     =   -1  'True
               TabIndex        =   451
               Top             =   1920
               Width           =   8055
            End
            Begin VB.CheckBox Chkbarcode 
               Alignment       =   1  'Right Justify
               Caption         =   "»Ì«‰«  «·—Õ·«  Ì „ «÷«›… «·»Ì«‰«  «·Ï «·Ã—Ìœ  ·ﬁ«∆Ì"
               Height          =   285
               Index           =   109
               Left            =   570
               RightToLeft     =   -1  'True
               TabIndex        =   448
               Top             =   1680
               Width           =   8055
            End
            Begin VB.CheckBox Chkbarcode 
               Alignment       =   1  'Right Justify
               Caption         =   "«„— «·‘€· ··’Ì«‰… »‰«¡ ⁄·Ì ÿ·» «Ê Œÿ… ›ﬁÿ"
               Height          =   285
               Index           =   100
               Left            =   990
               RightToLeft     =   -1  'True
               TabIndex        =   440
               Top             =   1440
               Width           =   7635
            End
            Begin VB.CheckBox Chkbarcode 
               Alignment       =   1  'Right Justify
               Caption         =   "‘«‘Â »Ì«‰«  «·—Õ·«  «·Õ›Ÿ »œÊ‰ «·„’—Ê›« "
               Height          =   285
               Index           =   84
               Left            =   390
               RightToLeft     =   -1  'True
               TabIndex        =   424
               Top             =   1200
               Width           =   8235
            End
            Begin VB.CheckBox Chkbarcode 
               Alignment       =   1  'Right Justify
               Caption         =   "‘«‘Â «·„—ﬂ»«   ŸÂ— «·”«∆ﬁÌ‰ ›ﬁÿ"
               Height          =   285
               Index           =   43
               Left            =   570
               RightToLeft     =   -1  'True
               TabIndex        =   371
               Top             =   960
               Width           =   8055
            End
            Begin VB.CheckBox chkDriverEra 
               Alignment       =   1  'Right Justify
               Caption         =   "⁄‰œ «‰‘«¡  ”«∆ﬁ ÃœÌœ Ì „ › Õ Õ”«» ⁄ÂœÂ ·Â"
               Height          =   285
               Left            =   630
               RightToLeft     =   -1  'True
               TabIndex        =   213
               Top             =   720
               Width           =   7995
            End
            Begin VB.CheckBox ChkDriverBox 
               Alignment       =   1  'Right Justify
               Caption         =   "⁄‰œ «‰‘«¡  ”«∆ﬁ ÃœÌœ Ì „ › Õ Õ”«» ’‰œÊﬁ ·Â"
               Height          =   285
               Left            =   630
               RightToLeft     =   -1  'True
               TabIndex        =   212
               Top             =   480
               Width           =   7995
            End
         End
      End
      Begin C1SizerLibCtl.C1Elastic Ele 
         Height          =   9375
         Index           =   5
         Left            =   17805
         TabIndex        =   42
         TabStop         =   0   'False
         Top             =   45
         Width           =   12225
         _cx             =   21564
         _cy             =   16536
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
         Begin VB.CheckBox Chkbarcode 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "„”„Ï «·›« Ê—… «·Œœ„Ì… (›« Ê—… ÷—Ì»Ì…)"
            Height          =   285
            Index           =   224
            Left            =   -2640
            RightToLeft     =   -1  'True
            TabIndex        =   648
            Top             =   5940
            Width           =   10455
         End
         Begin VB.CheckBox Chkbarcode 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "«·”⁄Ï Ê«·ﬂÂ—»«¡ Ê«·„Ì«Â «Ã»«—Ï ›Ï «·⁄ﬁœ"
            Height          =   285
            Index           =   204
            Left            =   -2640
            RightToLeft     =   -1  'True
            TabIndex        =   570
            Top             =   5640
            Width           =   10455
         End
         Begin VB.CheckBox Chkbarcode 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "«·—ﬁ„ «·„ÊÕœ «·“«„Ì ›Ì ⁄ﬁœ «·«ÌÃ«—"
            ForeColor       =   &H000000FF&
            Height          =   285
            Index           =   203
            Left            =   0
            RightToLeft     =   -1  'True
            TabIndex        =   569
            Top             =   0
            Width           =   10455
         End
         Begin VB.CheckBox Chkbarcode 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "«·—ﬁ„ «·„ÊÕœ «·“«„Ì ›Ì ⁄ﬁœ «·«ÌÃ«—"
            ForeColor       =   &H000000FF&
            Height          =   285
            Index           =   195
            Left            =   -2640
            RightToLeft     =   -1  'True
            TabIndex        =   557
            Top             =   5400
            Width           =   10455
         End
         Begin VB.CheckBox Chkbarcode 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "‘«‘Â «À»«  «·«Ì—«œ ··«ÌÃ«— ›ﬁÿ"
            ForeColor       =   &H000000FF&
            Height          =   285
            Index           =   193
            Left            =   -2640
            RightToLeft     =   -1  'True
            TabIndex        =   555
            Top             =   5160
            Width           =   10455
         End
         Begin VB.CheckBox Chkbarcode 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "«· ’›Ì… »—ﬁ„ «·⁄ﬁœ "
            Height          =   285
            Index           =   159
            Left            =   -2640
            RightToLeft     =   -1  'True
            TabIndex        =   509
            Top             =   4890
            Width           =   10455
         End
         Begin VB.CheckBox Chkbarcode 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "›Ì ”‰œ ’—›  Õ·Ì·Ì ··⁄ﬁ«—«   «·€«¡ «„ﬂ«‰ÌÂ  ⁄œÌ· «·⁄ﬁ«—  "
            Height          =   285
            Index           =   156
            Left            =   -2640
            RightToLeft     =   -1  'True
            TabIndex        =   506
            Top             =   4560
            Width           =   10455
         End
         Begin VB.CheckBox Chkbarcode 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "”‰œ ’—› «·„œ›Ê⁄«  ··„«·ﬂ ÌŸÂ— «· ’›Ì«  «·„”œœÂ ›ﬁÿ"
            Height          =   285
            Index           =   155
            Left            =   -2640
            RightToLeft     =   -1  'True
            TabIndex        =   505
            Top             =   4320
            Width           =   10455
         End
         Begin VB.CheckBox Chkbarcode 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "«„ﬂ«‰Ì… «·”œ«œ «·Ã“∆Ì ··ﬁÌ„… «·„÷«›… "
            Height          =   285
            Index           =   154
            Left            =   -2640
            RightToLeft     =   -1  'True
            TabIndex        =   504
            Top             =   4020
            Width           =   10455
         End
         Begin VB.CheckBox Chkbarcode 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "›Ì ‘«‘Â «·«ÌÃ«—«  «·„” Õﬁ… «·«” Õﬁ«ﬁ ÿ»ﬁ« · «—ÌŒ «·œ›⁄Â Ê·Ì”  «—ÌŒ ”‰œ «·«” Õﬁ«ﬁ"
            Height          =   285
            Index           =   151
            Left            =   3960
            RightToLeft     =   -1  'True
            TabIndex        =   501
            Top             =   8760
            Width           =   8145
         End
         Begin VB.CheckBox Chkbarcode 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "› Õ Õ”«» ·ﬂ· ⁄ﬁ«—"
            ForeColor       =   &H00000000&
            Height          =   405
            Index           =   138
            Left            =   7320
            RightToLeft     =   -1  'True
            TabIndex        =   485
            Top             =   1320
            Width           =   4785
         End
         Begin VB.CheckBox Chkbarcode 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "‰Ê⁄ «·⁄ﬁœ «·Ì« „‰ ‘«‘… «·⁄ﬁ«—"
            Height          =   285
            Index           =   120
            Left            =   1650
            RightToLeft     =   -1  'True
            TabIndex        =   459
            Top             =   8520
            Width           =   10455
         End
         Begin VB.CheckBox Chkbarcode 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "«‰‘«¡ ﬁÌœ ⁄„Ê·«  «·„‰«œÌ» ›Ì «·„ﬁ»Ê÷« "
            Height          =   285
            Index           =   119
            Left            =   2130
            RightToLeft     =   -1  'True
            TabIndex        =   458
            Top             =   8160
            Width           =   9975
         End
         Begin VB.CheckBox Chkbarcode 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "› Õ Õ”«» ﬁÌ„… „÷«›… ·ﬂ· „«·ﬂ"
            Height          =   285
            Index           =   118
            Left            =   7800
            RightToLeft     =   -1  'True
            TabIndex        =   457
            Top             =   7920
            Width           =   4305
         End
         Begin VB.CheckBox Chkbarcode 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "⁄œ„ «‰‘«¡ «·ﬁÌœ ›Ì ⁄ﬁÊœ «·«ÌÃ«—"
            Height          =   405
            Index           =   117
            Left            =   7800
            RightToLeft     =   -1  'True
            TabIndex        =   456
            Top             =   7560
            Width           =   4305
         End
         Begin VB.CheckBox Chkbarcode 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "⁄œ„ «‰‘«¡ «·ﬁÌœ ¬·Ì« ›Ï «· ⁄«ﬁœ« "
            Height          =   405
            Index           =   110
            Left            =   7800
            RightToLeft     =   -1  'True
            TabIndex        =   450
            Top             =   6840
            Visible         =   0   'False
            Width           =   4305
         End
         Begin VB.CheckBox Chkbarcode 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "«·”„«Õ »«· ⁄œÌ· ›Ï «·œ›⁄«  ›Ï «·⁄ﬁÊœ"
            Height          =   405
            Index           =   111
            Left            =   7800
            RightToLeft     =   -1  'True
            TabIndex        =   449
            Top             =   7200
            Width           =   4305
         End
         Begin VB.CheckBox Chkbarcode 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "«·”„«Õ » ŒÿÌ «·œ›⁄Â"
            ForeColor       =   &H00000000&
            Height          =   405
            Index           =   98
            Left            =   7320
            RightToLeft     =   -1  'True
            TabIndex        =   438
            Top             =   6480
            Width           =   4785
         End
         Begin VB.CheckBox Chkbarcode 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "«” Õﬁ«ﬁ «·⁄„Ê·Â"
            ForeColor       =   &H00000000&
            Height          =   405
            Index           =   95
            Left            =   7320
            RightToLeft     =   -1  'True
            TabIndex        =   435
            Top             =   5760
            Width           =   4785
         End
         Begin VB.CheckBox Chkbarcode 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "«·⁄„Ê·…  Õ„· ⁄·Ì «·„«·ﬂ"
            ForeColor       =   &H00FF0000&
            Height          =   405
            Index           =   94
            Left            =   7320
            RightToLeft     =   -1  'True
            TabIndex        =   434
            Top             =   6120
            Width           =   4785
         End
         Begin VB.CheckBox Chkbarcode 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "«” Õﬁ«ﬁ «·Œœ„« "
            ForeColor       =   &H00000000&
            Height          =   285
            Index           =   93
            Left            =   7320
            RightToLeft     =   -1  'True
            TabIndex        =   433
            Top             =   5520
            Width           =   4785
         End
         Begin VB.CheckBox Chkbarcode 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "«” Õﬁ«ﬁ «·ﬂÂ—»«¡"
            ForeColor       =   &H00000000&
            Height          =   405
            Index           =   92
            Left            =   7320
            RightToLeft     =   -1  'True
            TabIndex        =   432
            Top             =   5160
            Width           =   4785
         End
         Begin VB.CheckBox Chkbarcode 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "«” Õﬁ«ﬁ «·„«¡"
            ForeColor       =   &H00000000&
            Height          =   405
            Index           =   91
            Left            =   7320
            RightToLeft     =   -1  'True
            TabIndex        =   431
            Top             =   4800
            Width           =   4785
         End
         Begin VB.CheckBox Chkbarcode 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "«” Õﬁ«ﬁ «·”⁄Ì"
            ForeColor       =   &H00000000&
            Height          =   405
            Index           =   90
            Left            =   7320
            RightToLeft     =   -1  'True
            TabIndex        =   430
            Top             =   4440
            Width           =   4785
         End
         Begin VB.CheckBox Chkbarcode 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "«·„«¡ Ê«·ﬂÂ—»«¡ Ê«·Œœ„«  ··„«·ﬂ"
            ForeColor       =   &H00000000&
            Height          =   405
            Index           =   89
            Left            =   7320
            RightToLeft     =   -1  'True
            TabIndex        =   429
            Top             =   4080
            Width           =   4785
         End
         Begin VB.CheckBox Chkbarcode 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "«· «„Ì‰ ··„«·ﬂ"
            ForeColor       =   &H00000000&
            Height          =   405
            Index           =   88
            Left            =   7320
            RightToLeft     =   -1  'True
            TabIndex        =   428
            Top             =   3720
            Width           =   4785
         End
         Begin VB.TextBox TXTReturnSallingIntervalCount 
            Height          =   285
            Index           =   2
            Left            =   9000
            TabIndex        =   384
            Top             =   3360
            Width           =   615
         End
         Begin VB.TextBox Text17 
            Alignment       =   2  'Center
            Height          =   285
            Left            =   0
            MaxLength       =   1
            RightToLeft     =   -1  'True
            TabIndex        =   382
            Top             =   0
            Width           =   375
         End
         Begin VB.CheckBox Chkbarcode 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "Ì „ › Õ Õ”«» «ÌÃ«—«  „” Õﬁ… ·ﬂ· „«·ﬂ"
            ForeColor       =   &H00000000&
            Height          =   405
            Index           =   33
            Left            =   6600
            RightToLeft     =   -1  'True
            TabIndex        =   353
            Top             =   360
            Width           =   5505
         End
         Begin VB.CheckBox chkCustomerhavethreeAccounts1 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "«· ⁄«„· »‘Ìﬂ«    „ƒÃ·… ··„”«Â„Ì‰ Ê «·„œ›Ê⁄«   «·„ﬁœ„… ··„”«Â„Ì‰ "
            ForeColor       =   &H00000000&
            Height          =   405
            Left            =   6600
            RightToLeft     =   -1  'True
            TabIndex        =   351
            Top             =   2880
            Width           =   5505
         End
         Begin VB.Frame Fra 
            BackColor       =   &H00E2E9E9&
            Caption         =   "⁄‰œ  ”ÃÌ· ”‰œ ﬁ»÷ ÃœÌœ  "
            ForeColor       =   &H000000FF&
            Height          =   1065
            Index           =   5
            Left            =   8040
            RightToLeft     =   -1  'True
            TabIndex        =   345
            Top             =   1680
            Width           =   4185
            Begin VB.Frame Frame29 
               Caption         =   "Frame29"
               Height          =   15
               Index           =   1
               Left            =   120
               RightToLeft     =   -1  'True
               TabIndex        =   349
               Top             =   1440
               Width           =   3975
            End
            Begin VB.OptionButton Opt 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   " «—ÌŒ «·ÃÂ«“ «·”Ì—›— ( „ «Õ ›ﬁÿ ›Ï Õ«·… ⁄„·  «·‘»ﬂ…)"
               Height          =   255
               Index           =   10
               Left            =   30
               RightToLeft     =   -1  'True
               TabIndex        =   348
               Top             =   780
               Width           =   4065
            End
            Begin VB.OptionButton Opt 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   " «—ÌŒ «Œ—  ”‰œ ﬁ»÷ „”Ã·  ﬁ»·Â"
               Height          =   255
               Index           =   9
               Left            =   750
               RightToLeft     =   -1  'True
               TabIndex        =   347
               Top             =   480
               Width           =   3375
            End
            Begin VB.OptionButton Opt 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   " «—ÌŒ  ”‰œ ﬁ»÷ ÂÊ «· «—ÌŒ «·Õ«·Ï ›Ï «·ÃÂ«“ "
               Height          =   255
               Index           =   8
               Left            =   510
               RightToLeft     =   -1  'True
               TabIndex        =   346
               Top             =   240
               Value           =   -1  'True
               Width           =   3615
            End
         End
         Begin VB.CheckBox Chkbarcode 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "› Õ Õ”«»  √„Ì‰«  ··€Ì—ﬂ· ⁄„Ì· "
            ForeColor       =   &H00000000&
            Height          =   405
            Index           =   11
            Left            =   7320
            RightToLeft     =   -1  'True
            TabIndex        =   307
            Top             =   960
            Width           =   4785
         End
         Begin VB.CheckBox Chkbarcode 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "«” Õﬁ«ﬁ «·⁄ﬁœ ⁄·Ï «·œ›⁄Â «·«Ê·Ì ›ﬁÿ"
            Height          =   285
            Index           =   9
            Left            =   8160
            RightToLeft     =   -1  'True
            TabIndex        =   306
            Top             =   720
            Width           =   3945
         End
         Begin VB.Label Label8 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "⁄œœ „—«  «·ÕÃ“ ··⁄„Ì·"
            Height          =   255
            Left            =   9720
            RightToLeft     =   -1  'True
            TabIndex        =   385
            Top             =   3360
            Width           =   2415
         End
      End
      Begin C1SizerLibCtl.C1Elastic Ele 
         Height          =   9375
         Index           =   12
         Left            =   18105
         TabIndex        =   43
         TabStop         =   0   'False
         Top             =   45
         Width           =   12225
         _cx             =   21564
         _cy             =   16536
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
         Begin VB.Frame Frame31 
            BackColor       =   &H00E2E9E9&
            Caption         =   "«·«’Ê· «·À«» Â"
            ForeColor       =   &H000000FF&
            Height          =   8655
            Left            =   2400
            RightToLeft     =   -1  'True
            TabIndex        =   175
            Top             =   0
            Width           =   8805
            Begin VB.CheckBox Chkbarcode 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "⁄‰œ «÷«›… «·«’Ê· Ì „ › Õ Õ”«» ··«÷«›«  "
               Height          =   405
               Index           =   32
               Left            =   90
               RightToLeft     =   -1  'True
               TabIndex        =   352
               Top             =   840
               Width           =   8295
            End
            Begin VB.CheckBox ChkAssetAccount1 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "⁄‰œ «‰‘«¡ „Ã„Ê⁄Â «’· Ì „ › Ã Õ”«» «—»«Õ ÊŒ”«∆— ·ﬂ· „Ã„Ê⁄Â «·Ì«"
               ForeColor       =   &H00000000&
               Height          =   525
               Left            =   150
               RightToLeft     =   -1  'True
               TabIndex        =   177
               Top             =   480
               Width           =   8235
            End
            Begin VB.CheckBox ChkAssetAccount 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "Õ”«» «·„Ã„⁄ Ì »⁄ «·«’Ê·"
               Height          =   285
               Left            =   540
               RightToLeft     =   -1  'True
               TabIndex        =   176
               Top             =   240
               Width           =   7845
            End
         End
      End
      Begin C1SizerLibCtl.C1Elastic Ele 
         Height          =   9375
         Index           =   13
         Left            =   18405
         TabIndex        =   44
         TabStop         =   0   'False
         Top             =   45
         Width           =   12225
         _cx             =   21564
         _cy             =   16536
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
         Begin C1SizerLibCtl.C1Elastic Ele 
            Height          =   9255
            Index           =   20
            Left            =   2400
            TabIndex        =   68
            TabStop         =   0   'False
            Top             =   0
            Width           =   8805
            _cx             =   15531
            _cy             =   16325
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
            Begin VB.CheckBox Chkbarcode 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "«·„ﬁ«Ê· Ì‰‘√ «—»⁄ Õ”«»« "
               Height          =   405
               Index           =   209
               Left            =   1320
               RightToLeft     =   -1  'True
               TabIndex        =   575
               Top             =   3480
               Width           =   2175
            End
            Begin VB.CheckBox Chkbarcode 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "”‰œ «·ﬁ»÷ „‰ „‘—Ê⁄ »Ì”„⁄ ›Ì ﬂ‘› Õ”«» «·„‘—Ê⁄"
               Height          =   405
               Index           =   126
               Left            =   810
               RightToLeft     =   -1  'True
               TabIndex        =   466
               Top             =   8160
               Width           =   7335
            End
            Begin VB.CheckBox Chkbarcode 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "«„ﬂ«‰Ì…  ⁄œÌ· ”‰œ«  «·ﬁ»÷ «·„—»ÊÿÂ »«·„‘«—Ì⁄"
               Height          =   405
               Index           =   115
               Left            =   750
               RightToLeft     =   -1  'True
               TabIndex        =   454
               Top             =   7800
               Width           =   7395
            End
            Begin VB.CheckBox Chkbarcode 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "«·ﬁÌœ  Õ·Ì·Ì ›Ì „” Œ·’«  «·„‘«—Ì⁄"
               Height          =   405
               Index           =   104
               Left            =   630
               RightToLeft     =   -1  'True
               TabIndex        =   444
               Top             =   7440
               Width           =   7515
            End
            Begin VB.CheckBox Chkbarcode 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "«·”„«Õ » ⁄œÌ· «·”⁄— «·„⁄ „œ ›Ì «·„” Œ·’« "
               Height          =   405
               Index           =   99
               Left            =   60
               RightToLeft     =   -1  'True
               TabIndex        =   439
               Top             =   7080
               Width           =   8085
            End
            Begin VB.CheckBox Chkbarcode 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "«· ⁄«„· »Õ”«» Õ”‰ «·«œ«¡"
               Height          =   405
               Index           =   80
               Left            =   3840
               RightToLeft     =   -1  'True
               TabIndex        =   419
               Top             =   6720
               Width           =   4305
            End
            Begin VB.TextBox TXTReturnSallingIntervalCount 
               Height          =   285
               Index           =   3
               Left            =   1710
               TabIndex        =   408
               Top             =   6480
               Width           =   1695
            End
            Begin VB.CheckBox Chkbarcode 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "⁄œœ Œ«‰«  «· ﬁ—Ì» ›Ì „” Œ·’«  «·„‘«—Ì⁄"
               Height          =   405
               Index           =   69
               Left            =   3840
               RightToLeft     =   -1  'True
               TabIndex        =   407
               Top             =   6360
               Width           =   4305
            End
            Begin VB.CheckBox Chkbarcode 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "„‘ —Ì«  «·„‘«—Ì⁄ ·« ‰‘√ ”‰œ «” ·«„ „Œ“‰Ì"
               Height          =   405
               Index           =   67
               Left            =   60
               RightToLeft     =   -1  'True
               TabIndex        =   404
               Top             =   6000
               Width           =   8085
            End
            Begin VB.CheckBox Chkbarcode 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "›Ì »Ì«‰«  «·„‘«—Ì⁄ «·Õ«·… «·€«·»…  Õ  «· ‰›Ì–"
               Height          =   405
               Index           =   56
               Left            =   60
               RightToLeft     =   -1  'True
               TabIndex        =   393
               Top             =   5640
               Width           =   8085
            End
            Begin VB.CheckBox Chkbarcode 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "›Ì ›Ê« Ì— «·„‘«—Ì⁄ ›’·  ”—Ì«· «·⁄„Ì· ⁄‰ ”—Ì«· «·„ﬁ«Ê·"
               Height          =   405
               Index           =   53
               Left            =   60
               RightToLeft     =   -1  'True
               TabIndex        =   381
               Top             =   5280
               Width           =   8085
            End
            Begin VB.OptionButton OPTdISCOUNT 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "«·Œ’Ê„«   ⁄·Ì «·«Ì—«œ« "
               Height          =   195
               Index           =   1
               Left            =   4560
               RightToLeft     =   -1  'True
               TabIndex        =   337
               Top             =   4800
               Width           =   3135
            End
            Begin VB.OptionButton OPTdISCOUNT 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "«·Œ’Ê„«   Œ›÷ «·„’—Ê›"
               Height          =   195
               Index           =   0
               Left            =   4560
               RightToLeft     =   -1  'True
               TabIndex        =   336
               Top             =   4560
               Width           =   3135
            End
            Begin VB.CheckBox Chkbarcode 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "ﬁÌœ «·—Ê« » ÌÊ“⁄ ⁄·Ï «·„‘«—Ì⁄"
               Height          =   285
               Index           =   20
               Left            =   5640
               RightToLeft     =   -1  'True
               TabIndex        =   320
               Top             =   3960
               Width           =   2625
            End
            Begin VB.CheckBox Chkbarcode 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "Ì „ «· ⁄«„· „⁄ „ﬁ«Ê· «·»«ÿ‰ »Õ”«»  œ›⁄«  „ﬁœ„… Ê÷„«‰ «·«⁄„«· "
               Height          =   285
               Index           =   19
               Left            =   2880
               RightToLeft     =   -1  'True
               TabIndex        =   319
               Top             =   3600
               Width           =   5385
            End
            Begin VB.CheckBox Chkbarcode 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "«· ⁄«„· »«·«Ì—«œ«  «·„” Õﬁ…"
               Height          =   405
               Index           =   5
               Left            =   3960
               RightToLeft     =   -1  'True
               TabIndex        =   302
               Top             =   3240
               Width           =   4305
            End
            Begin VB.Frame Frame18 
               BackColor       =   &H00E2E9E9&
               Caption         =   "ÿ—Ìﬁ… «· ⁄«„· „⁄ «·„‘«—Ì⁄"
               ForeColor       =   &H000000FF&
               Height          =   975
               Left            =   3120
               RightToLeft     =   -1  'True
               TabIndex        =   77
               Top             =   240
               Width           =   5415
               Begin VB.OptionButton OptionItemsTotal 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "»‰Êœ «Ã„«·Ì…"
                  Height          =   195
                  Left            =   120
                  RightToLeft     =   -1  'True
                  TabIndex        =   79
                  Top             =   240
                  Width           =   5055
               End
               Begin VB.OptionButton OptionOperation 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "⁄„·Ì«   ›’Ì·Ì…"
                  Height          =   195
                  Left            =   120
                  RightToLeft     =   -1  'True
                  TabIndex        =   78
                  Top             =   600
                  Width           =   5055
               End
            End
            Begin VB.Frame Frame19 
               BackColor       =   &H00E2E9E9&
               Caption         =   "«· √ÀÌ— ⁄·Ì Õ”«»«  «·„‘—Ê⁄"
               ForeColor       =   &H000000FF&
               Height          =   975
               Left            =   3120
               RightToLeft     =   -1  'True
               TabIndex        =   74
               Top             =   1440
               Width           =   5415
               Begin VB.OptionButton glgeneral 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "·« ÌÊÃœ  √ÀÌ—"
                  Height          =   195
                  Left            =   120
                  RightToLeft     =   -1  'True
                  TabIndex        =   76
                  Top             =   600
                  Width           =   5055
               End
               Begin VB.OptionButton GlDetails 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   " ﬁ›· ›Ì Õ”«»«  «·„‘—Ê⁄"
                  Height          =   195
                  Left            =   120
                  RightToLeft     =   -1  'True
                  TabIndex        =   75
                  Top             =   240
                  Width           =   5055
               End
            End
            Begin VB.Frame Frame21 
               BackColor       =   &H00E2E9E9&
               Caption         =   "„œ… «·⁄„·Ì«  „Õœœ… »‹‹‹"
               ForeColor       =   &H000000FF&
               Height          =   615
               Left            =   3120
               RightToLeft     =   -1  'True
               TabIndex        =   69
               Top             =   2520
               Width           =   5415
               Begin VB.OptionButton Optday 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "ÌÊ„"
                  Height          =   195
                  Left            =   4320
                  RightToLeft     =   -1  'True
                  TabIndex        =   73
                  Top             =   240
                  Width           =   855
               End
               Begin VB.OptionButton OptMonth 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "‘Â—"
                  Height          =   195
                  Left            =   1560
                  RightToLeft     =   -1  'True
                  TabIndex        =   72
                  Top             =   240
                  Width           =   975
               End
               Begin VB.OptionButton OptYear 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "”‰…"
                  Height          =   195
                  Left            =   120
                  RightToLeft     =   -1  'True
                  TabIndex        =   71
                  Top             =   240
                  Width           =   975
               End
               Begin VB.OptionButton Optweek 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "«”»Ê⁄"
                  Height          =   195
                  Left            =   2880
                  RightToLeft     =   -1  'True
                  TabIndex        =   70
                  Top             =   240
                  Width           =   975
               End
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "”Ì«”… «·Œ’Ê„« "
               Height          =   375
               Index           =   27
               Left            =   6240
               RightToLeft     =   -1  'True
               TabIndex        =   335
               Top             =   4320
               Width           =   1965
            End
         End
      End
      Begin C1SizerLibCtl.C1Elastic Ele 
         Height          =   9375
         Index           =   14
         Left            =   18705
         TabIndex        =   45
         TabStop         =   0   'False
         Top             =   45
         Width           =   12225
         _cx             =   21564
         _cy             =   16536
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
         Begin VB.Frame Frame22 
            Caption         =   "„ «»⁄Â «·«”Â„"
            Height          =   8655
            Left            =   2520
            RightToLeft     =   -1  'True
            TabIndex        =   223
            Top             =   0
            Width           =   8805
            Begin VB.Frame Frame33 
               Caption         =   "—»ÿ «·Õ”«»«  »«·«”Â„"
               ForeColor       =   &H000000FF&
               Height          =   855
               Left            =   3120
               RightToLeft     =   -1  'True
               TabIndex        =   227
               Top             =   480
               Width           =   5415
               Begin VB.OptionButton OptArrowGroup 
                  Alignment       =   1  'Right Justify
                  Caption         =   "—»ÿ ⁄·Ï „” ÊÏ „Ã„Ê⁄«  «·«”Â„"
                  ForeColor       =   &H00000000&
                  Height          =   195
                  Left            =   120
                  RightToLeft     =   -1  'True
                  TabIndex        =   229
                  Top             =   480
                  Width           =   5055
               End
               Begin VB.OptionButton OptArrowBranch 
                  Alignment       =   1  'Right Justify
                  Caption         =   "—»ÿ ⁄·Ï „” ÊÏ «·›—⁄"
                  Height          =   195
                  Left            =   120
                  RightToLeft     =   -1  'True
                  TabIndex        =   228
                  Top             =   240
                  Width           =   5055
               End
            End
            Begin VB.Frame Frame23 
               Caption         =   "ÿ—Ìﬁ…  ﬁÌÌ„ «·«”Â„"
               ForeColor       =   &H000000FF&
               Height          =   855
               Left            =   3120
               RightToLeft     =   -1  'True
               TabIndex        =   224
               Top             =   1560
               Width           =   5415
               Begin VB.OptionButton Option1 
                  Alignment       =   1  'Right Justify
                  Caption         =   "„ Ê”ÿ ”⁄— «·‘—«¡"
                  Height          =   195
                  Left            =   120
                  RightToLeft     =   -1  'True
                  TabIndex        =   226
                  Top             =   240
                  Width           =   5055
               End
               Begin VB.OptionButton Option2 
                  Alignment       =   1  'Right Justify
                  Caption         =   "ÿ»ﬁ« ·«”⁄«— «·‘—«¡"
                  ForeColor       =   &H00000000&
                  Height          =   195
                  Left            =   120
                  RightToLeft     =   -1  'True
                  TabIndex        =   225
                  Top             =   480
                  Width           =   5055
               End
            End
         End
      End
      Begin C1SizerLibCtl.C1Elastic Ele 
         Height          =   9375
         Index           =   15
         Left            =   19005
         TabIndex        =   46
         TabStop         =   0   'False
         Top             =   45
         Width           =   12225
         _cx             =   21564
         _cy             =   16536
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
         Begin VB.Frame Frame57 
            Height          =   8655
            Left            =   2640
            RightToLeft     =   -1  'True
            TabIndex        =   214
            Top             =   120
            Width           =   8805
            Begin VB.CheckBox Chkbarcode 
               Alignment       =   1  'Right Justify
               Caption         =   " ⁄·Ì„Ì"
               Height          =   405
               Index           =   215
               Left            =   4290
               RightToLeft     =   -1  'True
               TabIndex        =   633
               Top             =   5040
               Width           =   4305
            End
            Begin VB.TextBox TXTData 
               Alignment       =   1  'Right Justify
               Height          =   285
               Index           =   2
               Left            =   5280
               RightToLeft     =   -1  'True
               TabIndex        =   390
               Top             =   4560
               Width           =   975
            End
            Begin VB.TextBox TXTData 
               Alignment       =   1  'Right Justify
               Height          =   285
               Index           =   1
               Left            =   5280
               RightToLeft     =   -1  'True
               TabIndex        =   389
               Top             =   3720
               Width           =   975
            End
            Begin VB.TextBox TXTData 
               Alignment       =   1  'Right Justify
               Height          =   285
               Index           =   0
               Left            =   5280
               RightToLeft     =   -1  'True
               TabIndex        =   386
               Top             =   3240
               Width           =   975
            End
            Begin VB.CheckBox Chkbarcode 
               Alignment       =   1  'Right Justify
               Caption         =   " ﬂÊÌœ «·ﬁÌÊœ ÿ»ﬁ« ··›—⁄"
               Height          =   405
               Index           =   14
               Left            =   4290
               RightToLeft     =   -1  'True
               TabIndex        =   341
               Top             =   2640
               Width           =   4305
            End
            Begin VB.CheckBox Chkbarcode 
               Alignment       =   1  'Right Justify
               Caption         =   "›Ì «·«⁄ „«œ«  Ì „ «Œ›«¡ «·„” ‰œ «·„—›Ê÷ „‰ ﬂ· «·„” ÊÌ« "
               Height          =   285
               Index           =   23
               Left            =   180
               RightToLeft     =   -1  'True
               TabIndex        =   323
               Top             =   3960
               Width           =   8445
            End
            Begin VB.Frame Frame49 
               BackColor       =   &H00E2E9E9&
               Caption         =   "«·«—ﬁ«„"
               ForeColor       =   &H000000FF&
               Height          =   1695
               Left            =   2640
               RightToLeft     =   -1  'True
               TabIndex        =   215
               Top             =   240
               Width           =   5895
               Begin VB.TextBox txt_ACCOUNT_digit 
                  Alignment       =   2  'Center
                  Height          =   285
                  Left            =   270
                  MaxLength       =   1
                  RightToLeft     =   -1  'True
                  TabIndex        =   315
                  Tag             =   "«·⁄œœ «·„ Êﬁ⁄ ··«—ﬁ«„ ›Ì «·ﬁÌœ Ê÷·ﬂ · ”ÂÌ· «· — Ì» „À«· 2011010001 Â‰«  „ «Œ Ì«— 3 ·–·ﬂ ŸÂ— 001"
                  Top             =   1320
                  Width           =   615
               End
               Begin VB.TextBox TxtQtyDigts 
                  Alignment       =   2  'Center
                  Height          =   285
                  Left            =   270
                  MaxLength       =   1
                  RightToLeft     =   -1  'True
                  TabIndex        =   218
                  Top             =   600
                  Width           =   615
               End
               Begin VB.TextBox TxtPriceDigts 
                  Alignment       =   2  'Center
                  Height          =   285
                  Left            =   270
                  MaxLength       =   1
                  RightToLeft     =   -1  'True
                  TabIndex        =   217
                  Top             =   240
                  Width           =   615
               End
               Begin VB.TextBox TxtPriceDigtsInst 
                  Alignment       =   2  'Center
                  Height          =   285
                  Left            =   270
                  MaxLength       =   1
                  RightToLeft     =   -1  'True
                  TabIndex        =   216
                  Top             =   960
                  Width           =   615
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   " ⁄œœ «·√—ﬁ«„ »⁄œ «·⁄·«„… «·⁄‘—Ì… ·›Ê« Ì— «·„»Ì⁄« "
                  Height          =   345
                  Index           =   16
                  Left            =   1050
                  RightToLeft     =   -1  'True
                  TabIndex        =   316
                  Top             =   1320
                  Width           =   4605
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "⁄œœ «·√—ﬁ«„ »⁄œ «·⁄·«„… «·⁄‘—Ì… ··ﬂ„Ì…"
                  Height          =   285
                  Index           =   14
                  Left            =   1020
                  RightToLeft     =   -1  'True
                  TabIndex        =   221
                  Top             =   600
                  Width           =   4575
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "⁄œœ «·√—ﬁ«„ »⁄œ «·⁄·«„… «·⁄‘—Ì… ··⁄„·…"
                  Height          =   285
                  Index           =   13
                  Left            =   1020
                  RightToLeft     =   -1  'True
                  TabIndex        =   220
                  Top             =   240
                  Width           =   4575
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "⁄œœ «·√—ﬁ«„ »⁄œ «·⁄·«„… «·⁄‘—Ì… ··«ﬁ”«ÿ"
                  ForeColor       =   &H00000000&
                  Height          =   285
                  Index           =   17
                  Left            =   1020
                  RightToLeft     =   -1  'True
                  TabIndex        =   219
                  Top             =   960
                  Width           =   4665
               End
            End
            Begin VB.Label Label23 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "⁄œœ Œ«‰«  «·›—⁄"
               ForeColor       =   &H00000000&
               Height          =   375
               Left            =   6360
               TabIndex        =   343
               Top             =   3240
               Width           =   1575
            End
            Begin VB.Label Label24 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "⁄œœ Œ«‰«  «·„Œ«“‰"
               ForeColor       =   &H00000000&
               Height          =   375
               Left            =   6360
               TabIndex        =   342
               Top             =   3720
               Width           =   1575
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               BackStyle       =   0  'Transparent
               Caption         =   "⁄œœ «·«—ﬁ«„ «·„ Êﬁ⁄Â ··ﬁÌœ"
               Height          =   345
               Index           =   15
               Left            =   5520
               RightToLeft     =   -1  'True
               TabIndex        =   222
               Top             =   4560
               Visible         =   0   'False
               Width           =   2985
            End
         End
      End
      Begin C1SizerLibCtl.C1Elastic Ele 
         Height          =   9375
         Index           =   16
         Left            =   19305
         TabIndex        =   47
         TabStop         =   0   'False
         Top             =   45
         Width           =   12225
         _cx             =   21564
         _cy             =   16536
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
         Begin C1SizerLibCtl.C1Elastic Ele 
            Height          =   8655
            Index           =   22
            Left            =   1440
            TabIndex        =   121
            TabStop         =   0   'False
            Top             =   0
            Width           =   10845
            _cx             =   19129
            _cy             =   15266
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
            Begin VB.Frame Frame8 
               Caption         =   " ‰»ÌÂ«  «·«⁄ „«œ«  «·„” ‰œÌ…/«·÷„«‰«  «·»‰ﬂÌ…"
               Height          =   615
               Left            =   120
               RightToLeft     =   -1  'True
               TabIndex        =   252
               Top             =   3600
               Width           =   4095
               Begin VB.ComboBox Combo14 
                  Height          =   315
                  ItemData        =   "FrmOptions.frx":10D7
                  Left            =   1440
                  List            =   "FrmOptions.frx":10E4
                  TabIndex        =   298
                  Top             =   240
                  Width           =   855
               End
               Begin VB.TextBox Text7 
                  Height          =   285
                  Left            =   2400
                  TabIndex        =   254
                  Top             =   240
                  Width           =   495
               End
               Begin VB.CheckBox CheckLC 
                  Alignment       =   1  'Right Justify
                  Caption         =   " ‰»Ì… ﬁ»· "
                  Height          =   285
                  Left            =   3000
                  RightToLeft     =   -1  'True
                  TabIndex        =   253
                  Top             =   240
                  Width           =   1005
               End
            End
            Begin VB.Frame Frame5 
               Caption         =   " ‰»ÌÂ«  «·«’‰«›"
               Height          =   1095
               Left            =   120
               RightToLeft     =   -1  'True
               TabIndex        =   249
               Top             =   2400
               Width           =   4095
               Begin VB.CheckBox ChkGuranAlram 
                  Alignment       =   1  'Right Justify
                  Caption         =   "⁄—÷  ‰»ÌÂ «·√’‰«› «· Ï œŒ·  ›Ï «·„œ… «·Õ—Ã… ··÷„«‰"
                  Height          =   525
                  Left            =   120
                  RightToLeft     =   -1  'True
                  TabIndex        =   251
                  Top             =   480
                  Visible         =   0   'False
                  Width           =   3825
               End
               Begin VB.CheckBox ChkShow 
                  Alignment       =   1  'Right Justify
                  Caption         =   "≈ŸÂ«—  ‰»ÌÂ«  «·√’‰«› «· Ì »·€  Õœ «·ÿ·»"
                  Height          =   285
                  Left            =   720
                  RightToLeft     =   -1  'True
                  TabIndex        =   250
                  Top             =   240
                  Width           =   3195
               End
            End
            Begin VB.Frame Frame4 
               Caption         =   "«⁄œ«œ«  ⁄«„…"
               Height          =   975
               Left            =   360
               RightToLeft     =   -1  'True
               TabIndex        =   245
               Top             =   120
               Width           =   3855
               Begin VB.CheckBox ChkHideAllAlarms 
                  Alignment       =   1  'Right Justify
                  Caption         =   "⁄œ„ «ŸÂ«— ‘«‘… «· ‰»ÌÂ«   ⁄‰œ «·œŒÊ·"
                  ForeColor       =   &H000000FF&
                  Height          =   285
                  Left            =   720
                  RightToLeft     =   -1  'True
                  TabIndex        =   248
                  Top             =   600
                  Width           =   2985
               End
               Begin VB.TextBox Text14 
                  Height          =   285
                  Left            =   1440
                  TabIndex        =   247
                  Top             =   240
                  Width           =   615
               End
               Begin VB.CheckBox Check1 
                  Alignment       =   1  'Right Justify
                  Caption         =   " –ﬂÌ— »«· ‰»ÌÂ ﬂ·"
                  ForeColor       =   &H00FF0000&
                  Height          =   285
                  Left            =   2160
                  RightToLeft     =   -1  'True
                  TabIndex        =   246
                  Top             =   240
                  Width           =   1545
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  BackStyle       =   0  'Transparent
                  Caption         =   "œﬁÌﬁ…"
                  Height          =   375
                  Index           =   24
                  Left            =   240
                  RightToLeft     =   -1  'True
                  TabIndex        =   326
                  Top             =   240
                  Width           =   1125
               End
            End
            Begin VB.Frame Frame3 
               Caption         =   " ‰»ÌÂ  «·„‘«—Ì⁄"
               Height          =   975
               Left            =   120
               RightToLeft     =   -1  'True
               TabIndex        =   242
               Top             =   1320
               Width           =   4095
               Begin VB.CheckBox ChKProjectsAlarm2 
                  Alignment       =   1  'Right Justify
                  Caption         =   "≈ŸÂ«—  ‰»ÌÂ«  «·„” Œ·’«  «· Ì ·„  ”œœ Õ Ï «·«‰  "
                  Height          =   285
                  Left            =   240
                  RightToLeft     =   -1  'True
                  TabIndex        =   244
                  Top             =   480
                  Width           =   3765
               End
               Begin VB.CheckBox ChKProjectsAlarm1 
                  Alignment       =   1  'Right Justify
                  Caption         =   "≈ŸÂ«—  ‰»ÌÂ«   «·»‰Êœ «· Ì  Œÿ  «·„ Êﬁ⁄ ·Â«  "
                  Height          =   285
                  Left            =   600
                  RightToLeft     =   -1  'True
                  TabIndex        =   243
                  Top             =   240
                  Width           =   3405
               End
            End
            Begin VB.Frame Frame43 
               BackColor       =   &H00E2E9E9&
               Height          =   1335
               Left            =   4320
               RightToLeft     =   -1  'True
               TabIndex        =   166
               Top             =   0
               Width           =   6255
               Begin VB.Frame Fra 
                  Caption         =   "«ŸÂ«—   ﬁ»·"
                  Height          =   615
                  Index           =   14
                  Left            =   240
                  RightToLeft     =   -1  'True
                  TabIndex        =   172
                  Top             =   120
                  Width           =   2055
                  Begin VB.ComboBox Combo1 
                     Height          =   315
                     ItemData        =   "FrmOptions.frx":10F7
                     Left            =   120
                     List            =   "FrmOptions.frx":1104
                     TabIndex        =   174
                     Top             =   240
                     Width           =   855
                  End
                  Begin VB.TextBox Text1 
                     Height          =   285
                     Left            =   1080
                     TabIndex        =   173
                     Top             =   240
                     Width           =   615
                  End
               End
               Begin VB.Frame Fra 
                  Caption         =   "«ŸÂ«—   ﬁ»·"
                  Height          =   615
                  Index           =   8
                  Left            =   240
                  RightToLeft     =   -1  'True
                  TabIndex        =   169
                  Top             =   720
                  Width           =   2055
                  Begin VB.TextBox Text2 
                     Height          =   285
                     Left            =   1080
                     TabIndex        =   171
                     Top             =   240
                     Width           =   615
                  End
                  Begin VB.ComboBox Combo2 
                     Height          =   315
                     ItemData        =   "FrmOptions.frx":1117
                     Left            =   120
                     List            =   "FrmOptions.frx":1124
                     TabIndex        =   170
                     Top             =   240
                     Width           =   855
                  End
               End
               Begin VB.CheckBox ChkDelayVal 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "≈ŸÂ«—  ‰»ÌÂ«  «·√Ê—«ﬁ «·„«·Ì… «·„” Õﬁ…"
                  Height          =   285
                  Left            =   3315
                  RightToLeft     =   -1  'True
                  TabIndex        =   168
                  Top             =   480
                  Width           =   2865
               End
               Begin VB.CheckBox ChkInstallmentMustPayed 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "≈ŸÂ«—  ‰»ÌÂ«  «·√ﬁ”«ÿ «· Ì Õ«‰ Êﬁ  ”œ«œÂ«"
                  Height          =   285
                  Left            =   2760
                  RightToLeft     =   -1  'True
                  TabIndex        =   167
                  Top             =   900
                  Width           =   3405
               End
            End
            Begin VB.Frame Frame38 
               BackColor       =   &H00E2E9E9&
               Caption         =   " ‰»ÌÂ«   ﬁÿ«⁄ «·‰ﬁ·Ì« "
               ForeColor       =   &H000000FF&
               Height          =   1935
               Left            =   4440
               RightToLeft     =   -1  'True
               TabIndex        =   153
               Top             =   6360
               Width           =   6255
               Begin VB.Frame Frame41 
                  Caption         =   "«ŸÂ«—   ﬁ»·"
                  Height          =   615
                  Left            =   240
                  RightToLeft     =   -1  'True
                  TabIndex        =   163
                  Top             =   1320
                  Width           =   2055
                  Begin VB.ComboBox Combo9 
                     Height          =   315
                     ItemData        =   "FrmOptions.frx":1137
                     Left            =   120
                     List            =   "FrmOptions.frx":1144
                     TabIndex        =   165
                     Top             =   240
                     Width           =   855
                  End
                  Begin VB.TextBox Text13 
                     Height          =   285
                     Left            =   1080
                     TabIndex        =   164
                     Top             =   240
                     Width           =   615
                  End
               End
               Begin VB.CheckBox ChkExpireLicense 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "«ŸÂ«—  ‰»ÌÂ   «·„⁄œ« /«·”Ì«—«  «·‰Ì ” ‰ ÂÏ «” „«— Â« ﬁ»·"
                  ForeColor       =   &H00000000&
                  Height          =   285
                  Left            =   2250
                  RightToLeft     =   -1  'True
                  TabIndex        =   162
                  Top             =   240
                  Width           =   3885
               End
               Begin VB.Frame Frame40 
                  Caption         =   "«ŸÂ«—   ﬁ»·"
                  Height          =   615
                  Left            =   240
                  RightToLeft     =   -1  'True
                  TabIndex        =   159
                  Top             =   720
                  Width           =   2055
                  Begin VB.TextBox Text12 
                     Height          =   285
                     Left            =   1080
                     TabIndex        =   161
                     Top             =   240
                     Width           =   615
                  End
                  Begin VB.ComboBox Combo8 
                     Height          =   315
                     ItemData        =   "FrmOptions.frx":1157
                     Left            =   120
                     List            =   "FrmOptions.frx":1164
                     TabIndex        =   160
                     Top             =   240
                     Width           =   855
                  End
               End
               Begin VB.CheckBox ChkExpireInsurance 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "«ŸÂ«—  ‰»ÌÂ   «·„⁄œ« /«·”Ì«—«  «·‰Ì  ”Ì‰ ÂÏ  √„Ì‰Â« ﬁ»·"
                  Height          =   285
                  Left            =   2370
                  RightToLeft     =   -1  'True
                  TabIndex        =   158
                  Top             =   840
                  Width           =   3765
               End
               Begin VB.Frame Frame39 
                  Caption         =   "«ŸÂ«—   ﬁ»·"
                  Height          =   615
                  Left            =   240
                  RightToLeft     =   -1  'True
                  TabIndex        =   155
                  Top             =   120
                  Width           =   2055
                  Begin VB.TextBox Text11 
                     Height          =   285
                     Left            =   1080
                     TabIndex        =   157
                     Top             =   240
                     Width           =   615
                  End
                  Begin VB.ComboBox Combo7 
                     Height          =   315
                     ItemData        =   "FrmOptions.frx":1177
                     Left            =   120
                     List            =   "FrmOptions.frx":1184
                     TabIndex        =   156
                     Top             =   240
                     Width           =   855
                  End
               End
               Begin VB.CheckBox ChkExpireTest 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "«ŸÂ«—  ‰»ÌÂ   «·„⁄œ« /«·”Ì«—«  «·‰Ì  ”Ì‰ ÂÏ ›Õ’Â« ﬁ»·"
                  Height          =   285
                  Left            =   2490
                  RightToLeft     =   -1  'True
                  TabIndex        =   154
                  Top             =   1440
                  Width           =   3645
               End
            End
            Begin VB.Frame Frame24 
               BackColor       =   &H00E2E9E9&
               Caption         =   " ‰»ÌÂ«  «œ«—Â «·«„·«ﬂ"
               ForeColor       =   &H000000FF&
               Height          =   1935
               Left            =   4440
               RightToLeft     =   -1  'True
               TabIndex        =   140
               Top             =   4320
               Width           =   6255
               Begin VB.CheckBox Check7 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "«ŸÂ«—  ‰»ÌÂ   «·⁄ﬁÊœ   «·„‰ÂÌÂ ﬁ»·"
                  Height          =   285
                  Left            =   2730
                  RightToLeft     =   -1  'True
                  TabIndex        =   152
                  Top             =   1440
                  Width           =   3405
               End
               Begin VB.Frame Frame27 
                  Caption         =   "«ŸÂ«—   ﬁ»·"
                  Height          =   615
                  Left            =   240
                  RightToLeft     =   -1  'True
                  TabIndex        =   149
                  Top             =   1320
                  Width           =   2055
                  Begin VB.TextBox Text10 
                     Height          =   285
                     Left            =   1080
                     TabIndex        =   151
                     Top             =   240
                     Width           =   615
                  End
                  Begin VB.ComboBox Combo12 
                     Height          =   315
                     ItemData        =   "FrmOptions.frx":1197
                     Left            =   120
                     List            =   "FrmOptions.frx":11A4
                     TabIndex        =   150
                     Top             =   240
                     Width           =   855
                  End
               End
               Begin VB.CheckBox Check6 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "«ŸÂ«—  ‰»ÌÂ   «·⁄ﬁÊœ «· Ì ” ‰ ÂÌ ﬁ»·"
                  Height          =   285
                  Left            =   2730
                  RightToLeft     =   -1  'True
                  TabIndex        =   148
                  Top             =   840
                  Width           =   3405
               End
               Begin VB.Frame Frame26 
                  Caption         =   "«ŸÂ«—   ﬁ»·"
                  Height          =   615
                  Left            =   240
                  RightToLeft     =   -1  'True
                  TabIndex        =   145
                  Top             =   720
                  Width           =   2055
                  Begin VB.ComboBox Combo11 
                     Height          =   315
                     ItemData        =   "FrmOptions.frx":11B7
                     Left            =   120
                     List            =   "FrmOptions.frx":11C4
                     TabIndex        =   147
                     Top             =   240
                     Width           =   855
                  End
                  Begin VB.TextBox Text9 
                     Height          =   285
                     Left            =   1080
                     TabIndex        =   146
                     Top             =   240
                     Width           =   615
                  End
               End
               Begin VB.CheckBox chkRentInstallments 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "«ŸÂ«—  ‰»ÌÂ «·«ÌÃ«—«  «·„” Õﬁ… ﬁ»·"
                  Height          =   285
                  Left            =   2730
                  RightToLeft     =   -1  'True
                  TabIndex        =   144
                  Top             =   240
                  Width           =   3405
               End
               Begin VB.Frame Frame25 
                  Caption         =   "«ŸÂ«—   ﬁ»·"
                  Height          =   615
                  Left            =   240
                  RightToLeft     =   -1  'True
                  TabIndex        =   141
                  Top             =   120
                  Width           =   2055
                  Begin VB.TextBox Text8 
                     Height          =   285
                     Left            =   1080
                     TabIndex        =   143
                     Top             =   240
                     Width           =   615
                  End
                  Begin VB.ComboBox Combo10 
                     Height          =   315
                     ItemData        =   "FrmOptions.frx":11D7
                     Left            =   120
                     List            =   "FrmOptions.frx":11E4
                     TabIndex        =   142
                     Top             =   240
                     Width           =   855
                  End
               End
            End
            Begin VB.CheckBox ChKHR 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "≈ŸÂ«—  ‰»ÌÂ«   ‘∆Ê‰ «·„ÊŸ›Ì‰"
               ForeColor       =   &H000000FF&
               Height          =   285
               Left            =   8160
               RightToLeft     =   -1  'True
               TabIndex        =   139
               Top             =   1320
               Width           =   2445
            End
            Begin VB.Frame Fra 
               BackColor       =   &H00E2E9E9&
               Height          =   3015
               Index           =   10
               Left            =   4320
               RightToLeft     =   -1  'True
               TabIndex        =   122
               Top             =   1320
               Width           =   6255
               Begin VB.CheckBox ChkExpirepoket 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "≈ŸÂ«—  ‰»ÌÂ«  Õ«›Ÿ… «·‰›Ê” «· Ì  ” ‰ ÂÌ"
                  Height          =   285
                  Left            =   2730
                  RightToLeft     =   -1  'True
                  TabIndex        =   138
                  Top             =   2040
                  Width           =   3405
               End
               Begin VB.Frame Fra 
                  Caption         =   "«ŸÂ«—   ﬁ»·"
                  Height          =   615
                  Index           =   13
                  Left            =   240
                  RightToLeft     =   -1  'True
                  TabIndex        =   135
                  Top             =   1920
                  Width           =   2055
                  Begin VB.TextBox Text6 
                     Height          =   285
                     Left            =   1080
                     TabIndex        =   137
                     Top             =   240
                     Width           =   615
                  End
                  Begin VB.ComboBox Combo6 
                     Height          =   315
                     ItemData        =   "FrmOptions.frx":11F7
                     Left            =   120
                     List            =   "FrmOptions.frx":1204
                     TabIndex        =   136
                     Top             =   240
                     Width           =   855
                  End
               End
               Begin VB.Frame Fra 
                  Caption         =   "«ŸÂ«—   ﬁ»·"
                  Height          =   615
                  Index           =   12
                  Left            =   240
                  RightToLeft     =   -1  'True
                  TabIndex        =   132
                  Top             =   1320
                  Width           =   2055
                  Begin VB.ComboBox Combo5 
                     Height          =   315
                     ItemData        =   "FrmOptions.frx":1217
                     Left            =   120
                     List            =   "FrmOptions.frx":1224
                     TabIndex        =   134
                     Top             =   240
                     Width           =   855
                  End
                  Begin VB.TextBox Text5 
                     Height          =   285
                     Left            =   1080
                     TabIndex        =   133
                     Top             =   240
                     Width           =   615
                  End
               End
               Begin VB.CheckBox ChkExpirepas 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "≈ŸÂ«—  ‰»ÌÂ«  «·ÃÊ«“«  «· Ì  ” ‰ ÂÌ"
                  Height          =   285
                  Left            =   2730
                  RightToLeft     =   -1  'True
                  TabIndex        =   131
                  Top             =   1440
                  Width           =   3405
               End
               Begin VB.CheckBox ChkExpireLicence 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "≈ŸÂ«—  ‰»ÌÂ«  «·—Œ’ «· Ì  ” ‰ ÂÌ"
                  Height          =   285
                  Left            =   2730
                  RightToLeft     =   -1  'True
                  TabIndex        =   130
                  Top             =   840
                  Width           =   3405
               End
               Begin VB.Frame Fra 
                  Caption         =   "«ŸÂ«—   ﬁ»·"
                  Height          =   615
                  Index           =   11
                  Left            =   240
                  RightToLeft     =   -1  'True
                  TabIndex        =   127
                  Top             =   720
                  Width           =   2055
                  Begin VB.TextBox Text4 
                     Height          =   285
                     Left            =   1080
                     TabIndex        =   129
                     Top             =   240
                     Width           =   615
                  End
                  Begin VB.ComboBox Combo4 
                     Height          =   315
                     ItemData        =   "FrmOptions.frx":1237
                     Left            =   120
                     List            =   "FrmOptions.frx":1244
                     TabIndex        =   128
                     Top             =   240
                     Width           =   855
                  End
               End
               Begin VB.Frame Fra 
                  Caption         =   "«ŸÂ«—   ﬁ»·"
                  Height          =   615
                  Index           =   9
                  Left            =   240
                  RightToLeft     =   -1  'True
                  TabIndex        =   124
                  Top             =   120
                  Width           =   2055
                  Begin VB.ComboBox Combo3 
                     Height          =   315
                     ItemData        =   "FrmOptions.frx":1257
                     Left            =   120
                     List            =   "FrmOptions.frx":1264
                     TabIndex        =   126
                     Top             =   240
                     Width           =   855
                  End
                  Begin VB.TextBox Text3 
                     Height          =   285
                     Left            =   1080
                     TabIndex        =   125
                     Top             =   240
                     Width           =   615
                  End
               End
               Begin VB.CheckBox ChkExpireEkama 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "≈ŸÂ«—  ‰»ÌÂ«  «·«ﬁ«„«   «· Ì  ” ‰ ÂÌ"
                  Height          =   285
                  Left            =   2730
                  RightToLeft     =   -1  'True
                  TabIndex        =   123
                  Top             =   240
                  Width           =   3405
               End
            End
         End
      End
      Begin C1SizerLibCtl.C1Elastic Ele 
         Height          =   9375
         Index           =   17
         Left            =   19605
         TabIndex        =   48
         TabStop         =   0   'False
         Top             =   45
         Width           =   12225
         _cx             =   21564
         _cy             =   16536
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
         Begin VB.Frame Frame58 
            Caption         =   "ŒÌ«—«  «·⁄—÷"
            Height          =   8655
            Left            =   2880
            RightToLeft     =   -1  'True
            TabIndex        =   190
            Top             =   120
            Width           =   8805
            Begin VB.CheckBox Chkbarcode 
               Alignment       =   1  'Right Justify
               Caption         =   "«ŸÂ«— ŒÌ«—«  «·ÿ«»⁄… "
               ForeColor       =   &H00000000&
               Height          =   285
               Index           =   225
               Left            =   -180
               RightToLeft     =   -1  'True
               TabIndex        =   649
               Top             =   1050
               Width           =   3345
            End
            Begin VB.TextBox TxtImagesPath 
               Alignment       =   2  'Center
               Height          =   285
               IMEMode         =   3  'DISABLE
               Index           =   3
               Left            =   1320
               PasswordChar    =   "*"
               TabIndex        =   641
               Text            =   "n20172018"
               Top             =   360
               Width           =   1815
            End
            Begin VB.CheckBox Chkbarcode 
               Alignment       =   1  'Right Justify
               Caption         =   "«·ÂÌœ— ›Ï «·ÿ»«⁄…"
               ForeColor       =   &H00000000&
               Height          =   285
               Index           =   207
               Left            =   240
               RightToLeft     =   -1  'True
               TabIndex        =   574
               Top             =   6120
               Width           =   3345
            End
            Begin VB.CheckBox Chkbarcode 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Height          =   285
               Index           =   197
               Left            =   0
               RightToLeft     =   -1  'True
               TabIndex        =   573
               Top             =   0
               Visible         =   0   'False
               Width           =   75
            End
            Begin VB.CheckBox Chkbarcode 
               Alignment       =   1  'Right Justify
               Caption         =   "Q"
               ForeColor       =   &H00000000&
               Height          =   285
               Index           =   201
               Left            =   240
               RightToLeft     =   -1  'True
               TabIndex        =   565
               Top             =   7680
               Width           =   3345
            End
            Begin VB.CheckBox Chkbarcode 
               Alignment       =   1  'Right Justify
               Caption         =   "ÿ»«⁄… «„— «·»Ì⁄ «·«Ê·"
               ForeColor       =   &H00000000&
               Height          =   285
               Index           =   200
               Left            =   240
               RightToLeft     =   -1  'True
               TabIndex        =   564
               Top             =   5880
               Width           =   3345
            End
            Begin VB.CheckBox Chkbarcode 
               Alignment       =   1  'Right Justify
               Caption         =   "«· √ﬂœ „‰ ’Ì€… —ﬁ„ «·ÃÊ«·"
               ForeColor       =   &H00000000&
               Height          =   285
               Index           =   190
               Left            =   240
               RightToLeft     =   -1  'True
               TabIndex        =   546
               Top             =   5520
               Width           =   3345
            End
            Begin VB.CheckBox Chkbarcode 
               Alignment       =   1  'Right Justify
               Caption         =   "«· √ﬂœ „‰ ’Ì€… «· «—ÌŒ"
               ForeColor       =   &H00000000&
               Height          =   285
               Index           =   187
               Left            =   240
               RightToLeft     =   -1  'True
               TabIndex        =   543
               Top             =   5160
               Width           =   3345
            End
            Begin VB.CheckBox Chkbarcode 
               Alignment       =   1  'Right Justify
               Caption         =   "«„ﬂ«‰ÌÂ  »œÌ· «·ÿ«»⁄Â «À‰«¡ «·ÿ»«⁄Â"
               ForeColor       =   &H00000000&
               Height          =   285
               Index           =   172
               Left            =   240
               RightToLeft     =   -1  'True
               TabIndex        =   525
               Top             =   4800
               Visible         =   0   'False
               Width           =   3345
            End
            Begin VB.TextBox TxtImagesPath 
               Alignment       =   2  'Center
               Height          =   285
               IMEMode         =   3  'DISABLE
               Index           =   2
               Left            =   1320
               PasswordChar    =   "*"
               TabIndex        =   363
               Text            =   "n20172018"
               Top             =   2280
               Width           =   1815
            End
            Begin VB.TextBox TxtImagesPath 
               Alignment       =   2  'Center
               Height          =   285
               Index           =   1
               Left            =   5280
               TabIndex        =   334
               Text            =   "Stander"
               Top             =   3240
               Width           =   1815
            End
            Begin VB.TextBox TxtImagesPath 
               Alignment       =   2  'Center
               Height          =   285
               Index           =   0
               Left            =   5280
               TabIndex        =   312
               Text            =   "Images"
               Top             =   2760
               Width           =   1815
            End
            Begin VB.CheckBox chkuserCode 
               Alignment       =   1  'Right Justify
               Caption         =   "œŒÊ· «·„” Œœ„Ì‰ »«·ﬂÊœ ›ﬁÿ"
               Height          =   285
               Left            =   5040
               RightToLeft     =   -1  'True
               TabIndex        =   295
               Top             =   2400
               Width           =   3465
            End
            Begin VB.TextBox TxtLogoheight 
               Alignment       =   2  'Center
               Height          =   285
               Left            =   5520
               TabIndex        =   285
               Text            =   "1500"
               Top             =   8040
               Width           =   615
            End
            Begin VB.TextBox TxtLogoWidth 
               Alignment       =   2  'Center
               Height          =   285
               Left            =   5520
               TabIndex        =   283
               Text            =   "4000"
               Top             =   7680
               Width           =   615
            End
            Begin VB.TextBox TxtZoom 
               Alignment       =   2  'Center
               Height          =   285
               Left            =   5520
               TabIndex        =   279
               Text            =   "100"
               Top             =   7320
               Width           =   615
            End
            Begin VB.CheckBox CHECK_OPEN_NEW_SCREEN 
               Alignment       =   1  'Right Justify
               Caption         =   "⁄‰œ › Õ «Ì ‘«‘…  »œ√ »ÃœÌœ ¬·Ì«"
               Height          =   285
               Left            =   5760
               RightToLeft     =   -1  'True
               TabIndex        =   210
               Top             =   480
               Width           =   2865
            End
            Begin VB.Frame Frame59 
               Caption         =   "ŒÌ«—«  «·Õ›Ÿ"
               ForeColor       =   &H000000FF&
               Height          =   1335
               Left            =   3360
               RightToLeft     =   -1  'True
               TabIndex        =   205
               Top             =   960
               Width           =   5295
               Begin VB.OptionButton Option4 
                  Alignment       =   1  'Right Justify
                  Caption         =   "«·Õ›Ÿ  Ê «·ÿ»«⁄Â  «·Ï «·ÿ«»⁄Â «·«› —«÷Ì…"
                  Height          =   195
                  Index           =   2
                  Left            =   0
                  RightToLeft     =   -1  'True
                  TabIndex        =   209
                  Top             =   720
                  Width           =   5055
               End
               Begin VB.OptionButton Option4 
                  Alignment       =   1  'Right Justify
                  Caption         =   "«·Õ›Ÿ  Ê «·ÿ»«⁄Â  «·Ï «·ÿ«»⁄Â «·«› —«÷Ì… Ê› Õ ‘«‘… ÃœÌœ…"
                  Height          =   195
                  Index           =   3
                  Left            =   0
                  RightToLeft     =   -1  'True
                  TabIndex        =   208
                  Top             =   960
                  Width           =   5055
               End
               Begin VB.OptionButton Option4 
                  Alignment       =   1  'Right Justify
                  Caption         =   " «·Õ›Ÿ  Ê «·ÿ»«⁄Â  ⁄·Ï «·‘«‘Â"
                  Height          =   195
                  Index           =   1
                  Left            =   0
                  RightToLeft     =   -1  'True
                  TabIndex        =   207
                  Top             =   480
                  Width           =   5055
               End
               Begin VB.OptionButton Option4 
                  Alignment       =   1  'Right Justify
                  Caption         =   "«·Õ›Ÿ ›ﬁÿ"
                  Height          =   195
                  Index           =   0
                  Left            =   0
                  RightToLeft     =   -1  'True
                  TabIndex        =   206
                  Top             =   240
                  Width           =   5055
               End
            End
            Begin VB.CheckBox chkshortCuts 
               Alignment       =   1  'Right Justify
               Caption         =   "⁄—÷ ‘—Ìÿ «·«Œ ’«—« "
               Height          =   195
               Left            =   240
               RightToLeft     =   -1  'True
               TabIndex        =   204
               Top             =   3360
               Visible         =   0   'False
               Width           =   2295
            End
            Begin VB.CheckBox Chktree 
               Alignment       =   1  'Right Justify
               Caption         =   "⁄—÷ ‘Ã—… «·«’‰«›"
               Height          =   195
               Left            =   720
               RightToLeft     =   -1  'True
               TabIndex        =   203
               Top             =   2520
               Visible         =   0   'False
               Width           =   2295
            End
            Begin VB.CheckBox ChkCalender 
               Alignment       =   1  'Right Justify
               Caption         =   "⁄—÷  «·‰ ÌÃÂ"
               Height          =   195
               Left            =   0
               RightToLeft     =   -1  'True
               TabIndex        =   202
               Top             =   2880
               Visible         =   0   'False
               Width           =   2295
            End
            Begin VB.CheckBox Chkgraphic 
               Alignment       =   1  'Right Justify
               Caption         =   "⁄—÷  „·Œ’ «·Õ—ﬂ… «·ÌÊ„Ì…"
               Height          =   195
               Left            =   -1200
               RightToLeft     =   -1  'True
               TabIndex        =   201
               Top             =   3840
               Visible         =   0   'False
               Width           =   4335
            End
            Begin VB.CheckBox ChkMessnger 
               Alignment       =   1  'Right Justify
               Caption         =   "⁄—÷ «·»—Ìœ «·œ«Œ·Ì «·Ì«"
               Height          =   195
               Left            =   4080
               RightToLeft     =   -1  'True
               TabIndex        =   200
               Top             =   3870
               Width           =   4335
            End
            Begin VB.CheckBox ChkViewAging 
               Alignment       =   1  'Right Justify
               Caption         =   "⁄—÷ «⁄„«— «·œÌÊ‰ ›Ì ﬂ‘Ê› «·Õ”«»« "
               Height          =   285
               Left            =   1590
               RightToLeft     =   -1  'True
               TabIndex        =   199
               Top             =   4200
               Width           =   6795
            End
            Begin VB.CheckBox ChkShowToolTip 
               Alignment       =   1  'Right Justify
               Caption         =   "⁄—÷ «·‘—Ìÿ «·œ⁄«∆Ì"
               Height          =   285
               Left            =   6600
               RightToLeft     =   -1  'True
               TabIndex        =   198
               Top             =   4560
               Width           =   1785
            End
            Begin VB.CheckBox ChkAsk 
               Alignment       =   1  'Right Justify
               Caption         =   "≈ŸÂ«— ŒÌ«—«  «·ÿ»«⁄…"
               ForeColor       =   &H00000000&
               Height          =   285
               Left            =   6600
               RightToLeft     =   -1  'True
               TabIndex        =   197
               Top             =   4920
               Width           =   1785
            End
            Begin VB.CheckBox ChkPrintBranchINGE 
               Alignment       =   1  'Right Justify
               Caption         =   "⁄—÷ «”„ «·›—⁄  »Ã«‰» «·Õ”«» ›Ì «·ﬁÌœ"
               Height          =   285
               Left            =   4440
               RightToLeft     =   -1  'True
               TabIndex        =   196
               Top             =   5280
               Width           =   3945
            End
            Begin VB.CheckBox ChkChartPrintinAS 
               Alignment       =   1  'Right Justify
               Caption         =   "⁄—÷  «·—”„ «·»Ì«‰Ì ›Ì ﬂ‘› «·Õ”«»"
               Height          =   285
               Left            =   4440
               RightToLeft     =   -1  'True
               TabIndex        =   195
               Top             =   6000
               Width           =   3945
            End
            Begin VB.CheckBox ChkPrintCCinGE 
               Alignment       =   1  'Right Justify
               Caption         =   "⁄—÷  „—ﬂ“ «· ﬂ·›… ›Ì «·ﬁÌœ"
               Height          =   285
               Left            =   4440
               RightToLeft     =   -1  'True
               TabIndex        =   194
               Top             =   5640
               Width           =   3945
            End
            Begin VB.Frame Fra 
               Caption         =   "«· «—ÌŒ «·«› —«÷Ì"
               ForeColor       =   &H000000FF&
               Height          =   825
               Index           =   4
               Left            =   3120
               RightToLeft     =   -1  'True
               TabIndex        =   191
               Top             =   6360
               Width           =   5385
               Begin VB.OptionButton DateOpt 
                  Alignment       =   1  'Right Justify
                  Caption         =   "„Ì·«œÌ"
                  Height          =   255
                  Index           =   0
                  Left            =   750
                  RightToLeft     =   -1  'True
                  TabIndex        =   193
                  Top             =   240
                  Value           =   -1  'True
                  Width           =   4335
               End
               Begin VB.OptionButton DateOpt 
                  Alignment       =   1  'Right Justify
                  Caption         =   "ÂÃ—Ì"
                  Height          =   255
                  Index           =   1
                  Left            =   750
                  RightToLeft     =   -1  'True
                  TabIndex        =   192
                  Top             =   480
                  Width           =   4335
               End
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               BackStyle       =   0  'Transparent
               Caption         =   "Vat Password"
               Height          =   375
               Index           =   40
               Left            =   3120
               RightToLeft     =   -1  'True
               TabIndex        =   642
               Top             =   330
               Width           =   1485
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               BackStyle       =   0  'Transparent
               Caption         =   "«·»«”Ê—œ «·—∆Ì”Ì"
               Height          =   375
               Index           =   30
               Left            =   3360
               RightToLeft     =   -1  'True
               TabIndex        =   364
               Top             =   2400
               Width           =   1485
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               BackStyle       =   0  'Transparent
               Caption         =   "„Ã·œ Õ›Ÿ «· ﬁ«—Ì— «·„ Œ’’…"
               Height          =   375
               Index           =   26
               Left            =   6840
               RightToLeft     =   -1  'True
               TabIndex        =   333
               Top             =   3120
               Width           =   1485
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               BackStyle       =   0  'Transparent
               Caption         =   "%"
               Height          =   375
               Index           =   23
               Left            =   3720
               RightToLeft     =   -1  'True
               TabIndex        =   313
               Top             =   7320
               Width           =   1725
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               BackStyle       =   0  'Transparent
               Caption         =   "„Ã·œ Õ›Ÿ «·’Ê—"
               Height          =   375
               Index           =   22
               Left            =   6840
               RightToLeft     =   -1  'True
               TabIndex        =   311
               Top             =   2760
               Width           =   1485
            End
            Begin VB.Label Label22 
               Alignment       =   1  'Right Justify
               Caption         =   "ÕÃ„ «··ÊÃÊ ›Ì «· ﬁ«—Ì— - «·«— ›«⁄"
               Height          =   255
               Left            =   6360
               RightToLeft     =   -1  'True
               TabIndex        =   284
               Top             =   8040
               Width           =   2295
            End
            Begin VB.Label Label21 
               Alignment       =   1  'Right Justify
               Caption         =   "ÕÃ„ «··ÊÃÊ ›Ì «· ﬁ«—Ì— - «·⁄—÷"
               Height          =   255
               Left            =   6360
               RightToLeft     =   -1  'True
               TabIndex        =   282
               Top             =   7680
               Width           =   2295
            End
            Begin VB.Label Label18 
               Alignment       =   1  'Right Justify
               Caption         =   "«· ﬂ»Ì— «·«› —«÷Ì ·· ﬁ«—Ì—"
               Height          =   255
               Left            =   6480
               RightToLeft     =   -1  'True
               TabIndex        =   278
               Top             =   7320
               Width           =   2055
            End
         End
      End
      Begin C1SizerLibCtl.C1Elastic Ele 
         Height          =   9375
         Index           =   23
         Left            =   19905
         TabIndex        =   256
         TabStop         =   0   'False
         Top             =   45
         Width           =   12225
         _cx             =   21564
         _cy             =   16536
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
         Begin C1SizerLibCtl.C1Elastic Ele 
            Height          =   8655
            Index           =   24
            Left            =   2640
            TabIndex        =   257
            TabStop         =   0   'False
            Top             =   0
            Width           =   8805
            _cx             =   15531
            _cy             =   15266
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
            Begin VB.Frame Frame47 
               BackColor       =   &H00E2E9E9&
               Caption         =   "«· ﬂ«·Ì› €Ì— «·„»«‘—…"
               ForeColor       =   &H000000FF&
               Height          =   855
               Left            =   2640
               RightToLeft     =   -1  'True
               TabIndex        =   260
               Top             =   960
               Width           =   5895
               Begin VB.TextBox Text15 
                  Alignment       =   2  'Center
                  Height          =   345
                  Left            =   1320
                  RightToLeft     =   -1  'True
                  TabIndex        =   262
                  Tag             =   "«·⁄œœ «·„ Êﬁ⁄ ··«—ﬁ«„ ›Ì «·ﬁÌœ Ê÷·ﬂ · ”ÂÌ· «· — Ì» „À«· 2011010001 Â‰«  „ «Œ Ì«— 3 ·–·ﬂ ŸÂ— 001"
                  Top             =   240
                  Width           =   615
               End
               Begin VB.CheckBox Check3 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   " Õ„Ì· ‰”»… À«» … ⁄·Ï «„— «·«‰ «Ã"
                  ForeColor       =   &H00000000&
                  Height          =   285
                  Left            =   2760
                  RightToLeft     =   -1  'True
                  TabIndex        =   261
                  Top             =   240
                  Width           =   2745
               End
               Begin VB.Label Label15 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "%"
                  Height          =   255
                  Left            =   -120
                  RightToLeft     =   -1  'True
                  TabIndex        =   264
                  Top             =   240
                  Width           =   735
               End
               Begin VB.Label Label14 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "Õœœ «·‰”»…"
                  Height          =   255
                  Left            =   1920
                  RightToLeft     =   -1  'True
                  TabIndex        =   263
                  Top             =   240
                  Width           =   855
               End
            End
            Begin VB.Frame Frame46 
               BackColor       =   &H00E2E9E9&
               Caption         =   "”Ì«”Â «·«‰ «Ã"
               ForeColor       =   &H000000FF&
               Height          =   615
               Left            =   2640
               RightToLeft     =   -1  'True
               TabIndex        =   258
               Top             =   240
               Width           =   5895
               Begin VB.CheckBox Check2 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "Ì „ «· ⁄«„· »«·«‰ «Ã «·‰„ÿÌ"
                  ForeColor       =   &H00000000&
                  Height          =   285
                  Left            =   2760
                  RightToLeft     =   -1  'True
                  TabIndex        =   259
                  Top             =   240
                  Width           =   2745
               End
            End
         End
      End
      Begin C1SizerLibCtl.C1Elastic Ele 
         Height          =   9375
         Index           =   25
         Left            =   20205
         TabIndex        =   265
         TabStop         =   0   'False
         Top             =   45
         Width           =   12225
         _cx             =   21564
         _cy             =   16536
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
         Begin C1SizerLibCtl.C1Elastic Ele 
            Height          =   8655
            Index           =   26
            Left            =   2640
            TabIndex        =   266
            TabStop         =   0   'False
            Top             =   0
            Width           =   8805
            _cx             =   15531
            _cy             =   15266
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
            Begin VB.Frame Frame51 
               BackColor       =   &H00E2E9E9&
               Caption         =   "”Ì«”Â «·«‰ «Ã"
               ForeColor       =   &H000000FF&
               Height          =   615
               Left            =   2640
               RightToLeft     =   -1  'True
               TabIndex        =   271
               Top             =   240
               Width           =   5895
               Begin VB.CheckBox Check9 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "Ì „ «· ⁄«„· »«·«‰ «Ã «·‰„ÿÌ"
                  ForeColor       =   &H00000000&
                  Height          =   285
                  Left            =   2760
                  RightToLeft     =   -1  'True
                  TabIndex        =   272
                  Top             =   240
                  Width           =   2745
               End
            End
            Begin VB.Frame Frame50 
               BackColor       =   &H00E2E9E9&
               Caption         =   "«· ﬂ«·Ì› €Ì— «·„»«‘—…"
               ForeColor       =   &H000000FF&
               Height          =   855
               Left            =   2640
               RightToLeft     =   -1  'True
               TabIndex        =   267
               Top             =   960
               Width           =   5895
               Begin VB.CheckBox Check8 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   " Õ„Ì· ‰”»… À«» … ⁄·Ï «„— «·«‰ «Ã"
                  ForeColor       =   &H00000000&
                  Height          =   285
                  Left            =   2760
                  RightToLeft     =   -1  'True
                  TabIndex        =   269
                  Top             =   240
                  Width           =   2745
               End
               Begin VB.TextBox Text16 
                  Alignment       =   2  'Center
                  Height          =   345
                  Left            =   960
                  RightToLeft     =   -1  'True
                  TabIndex        =   268
                  Tag             =   "«·⁄œœ «·„ Êﬁ⁄ ··«—ﬁ«„ ›Ì «·ﬁÌœ Ê÷·ﬂ · ”ÂÌ· «· — Ì» „À«· 2011010001 Â‰«  „ «Œ Ì«— 3 ·–·ﬂ ŸÂ— 001"
                  Top             =   240
                  Width           =   735
               End
               Begin VB.Label Label17 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "Õœœ «·‰”»…%"
                  Height          =   255
                  Left            =   1920
                  RightToLeft     =   -1  'True
                  TabIndex        =   270
                  Top             =   240
                  Width           =   855
               End
            End
         End
      End
      Begin C1SizerLibCtl.C1Elastic Ele 
         Height          =   9375
         Index           =   27
         Left            =   20505
         TabIndex        =   273
         TabStop         =   0   'False
         Top             =   45
         Width           =   12225
         _cx             =   21564
         _cy             =   16536
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
         Height          =   9375
         Index           =   28
         Left            =   20805
         TabIndex        =   274
         TabStop         =   0   'False
         Top             =   45
         Width           =   12225
         _cx             =   21564
         _cy             =   16536
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
         Height          =   9375
         Index           =   29
         Left            =   21105
         TabIndex        =   275
         TabStop         =   0   'False
         Top             =   45
         Width           =   12225
         _cx             =   21564
         _cy             =   16536
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
         Height          =   9375
         Index           =   30
         Left            =   21405
         TabIndex        =   276
         TabStop         =   0   'False
         Top             =   45
         Width           =   12225
         _cx             =   21564
         _cy             =   16536
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
         Begin VB.TextBox XPTxtComment 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   16
            Left            =   1740
            MaxLength       =   50
            RightToLeft     =   -1  'True
            TabIndex        =   647
            Top             =   8940
            Width           =   2325
         End
         Begin VB.TextBox XPTxtComment 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   12
            Left            =   1740
            MaxLength       =   50
            RightToLeft     =   -1  'True
            TabIndex        =   646
            Top             =   8520
            Width           =   2325
         End
         Begin VB.CheckBox Chkbarcode 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   " ›⁄Ì· «·„—Õ·… «·À«‰Ì… ÊÌ»"
            ForeColor       =   &H000000FF&
            Height          =   285
            Index           =   223
            Left            =   5250
            RightToLeft     =   -1  'True
            TabIndex        =   645
            Top             =   8490
            Width           =   2205
         End
         Begin VB.CheckBox Chkbarcode 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "—›⁄ «·›Ê« Ì— Ê›ﬁ« ··›—⁄"
            ForeColor       =   &H000000FF&
            Height          =   285
            Index           =   218
            Left            =   4470
            RightToLeft     =   -1  'True
            TabIndex        =   636
            Top             =   8880
            Width           =   2955
         End
         Begin VB.CheckBox Chkbarcode 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "—›⁄ «·›Ê« Ì— Ê›ﬁ« ··‰‘«ÿ"
            ForeColor       =   &H000000FF&
            Height          =   285
            Index           =   216
            Left            =   8160
            RightToLeft     =   -1  'True
            TabIndex        =   634
            Top             =   8940
            Width           =   2955
         End
         Begin VB.TextBox XPTxtComment 
            Height          =   375
            Index           =   6
            Left            =   2970
            MaxLength       =   50
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   610
            Top             =   2640
            Width           =   7755
         End
         Begin VB.CheckBox Chkbarcode 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   " ›⁄Ì· «·„—Õ·… «·À«‰Ì…"
            ForeColor       =   &H000000FF&
            Height          =   285
            Index           =   213
            Left            =   8160
            RightToLeft     =   -1  'True
            TabIndex        =   609
            Top             =   8520
            Width           =   2955
         End
         Begin VB.CheckBox Chkbarcode 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "«·›« Ê—…   —”· «·Ì« »„Ã—œ «·Õ›Ÿ"
            Height          =   285
            Index           =   212
            Left            =   7920
            RightToLeft     =   -1  'True
            TabIndex        =   608
            Top             =   8520
            Visible         =   0   'False
            Width           =   4035
         End
         Begin VB.TextBox XPTxtComment 
            Height          =   735
            Index           =   15
            Left            =   2910
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   605
            Top             =   7110
            Width           =   9075
         End
         Begin VB.ComboBox SendingMode 
            Height          =   315
            ItemData        =   "FrmOptions.frx":1277
            Left            =   2910
            List            =   "FrmOptions.frx":1279
            Style           =   2  'Dropdown List
            TabIndex        =   604
            Top             =   8025
            Width           =   7545
         End
         Begin VB.TextBox XPTxtComment 
            Height          =   735
            Index           =   13
            Left            =   2910
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   601
            Top             =   5520
            Width           =   9075
         End
         Begin VB.TextBox XPTxtComment 
            Height          =   735
            Index           =   14
            Left            =   2910
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   600
            Top             =   6285
            Width           =   9075
         End
         Begin VB.ComboBox Invoicetype 
            Height          =   315
            ItemData        =   "FrmOptions.frx":127B
            Left            =   7200
            List            =   "FrmOptions.frx":127D
            RightToLeft     =   -1  'True
            Style           =   2  'Dropdown List
            TabIndex        =   598
            Top             =   3180
            Width           =   3465
         End
         Begin VB.TextBox XPTxtComment 
            Height          =   375
            Index           =   10
            Left            =   2940
            MaxLength       =   50
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   593
            Top             =   3720
            Width           =   7755
         End
         Begin VB.TextBox XPTxtComment 
            Height          =   375
            Index           =   11
            Left            =   2940
            MaxLength       =   50
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   592
            Top             =   4200
            Width           =   7755
         End
         Begin VB.TextBox XPTxtComment 
            Height          =   735
            Index           =   9
            Left            =   2940
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   591
            Top             =   4680
            Width           =   9075
         End
         Begin VB.ComboBox DefaultInvoicetype 
            Height          =   315
            ItemData        =   "FrmOptions.frx":127F
            Left            =   2970
            List            =   "FrmOptions.frx":1281
            RightToLeft     =   -1  'True
            Style           =   2  'Dropdown List
            TabIndex        =   590
            Top             =   3150
            Width           =   2625
         End
         Begin VB.TextBox XPTxtComment 
            Height          =   375
            Index           =   4
            Left            =   2970
            MaxLength       =   150
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   585
            Top             =   1650
            Width           =   7755
         End
         Begin VB.TextBox XPTxtComment 
            Height          =   375
            Index           =   5
            Left            =   2970
            MaxLength       =   150
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   584
            Top             =   2130
            Width           =   7755
         End
         Begin VB.TextBox XPTxtComment 
            Height          =   375
            Index           =   7
            Left            =   2970
            MaxLength       =   150
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   583
            Top             =   1170
            Width           =   7755
         End
         Begin VB.TextBox XPTxtComment 
            Height          =   375
            Index           =   8
            Left            =   3000
            MaxLength       =   50
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   582
            Top             =   690
            Width           =   7755
         End
         Begin VB.CommandButton Command1 
            Caption         =   "«⁄œ«œ«   ⁄«„Â  ·‘—Ìÿ «·„⁄·Ê„« "
            Height          =   375
            Index           =   0
            Left            =   6330
            RightToLeft     =   -1  'True
            TabIndex        =   290
            Top             =   60
            Visible         =   0   'False
            Width           =   3135
         End
         Begin VB.CommandButton Command1 
            Caption         =   " ⁄—Ì› «·—”«∆· «·„ƒﬁ …"
            Height          =   375
            Index           =   2
            Left            =   3960
            RightToLeft     =   -1  'True
            TabIndex        =   289
            Top             =   60
            Visible         =   0   'False
            Width           =   2355
         End
         Begin VB.CommandButton Command1 
            Caption         =   "‰„«–Ã «·—”«∆·"
            Height          =   375
            Index           =   1
            Left            =   1590
            RightToLeft     =   -1  'True
            TabIndex        =   288
            Top             =   60
            Visible         =   0   'False
            Width           =   2325
         End
         Begin VB.Label lbl 
            BackColor       =   &H00E2E9E9&
            Caption         =   "Org name"
            Height          =   375
            Index           =   84
            Left            =   1650
            RightToLeft     =   -1  'True
            TabIndex        =   611
            Top             =   2715
            Width           =   1725
         End
         Begin VB.Label lbl 
            BackColor       =   &H00E2E9E9&
            Caption         =   "Secret Key"
            Height          =   375
            Index           =   82
            Left            =   1590
            RightToLeft     =   -1  'True
            TabIndex        =   607
            Top             =   7185
            Width           =   1725
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "«·Ê÷⁄ «·Õ«·Ì"
            Height          =   375
            Index           =   89
            Left            =   1050
            RightToLeft     =   -1  'True
            TabIndex        =   606
            Top             =   8070
            Width           =   1485
         End
         Begin VB.Label lbl 
            BackColor       =   &H00E2E9E9&
            Caption         =   "Private Key"
            Height          =   375
            Index           =   80
            Left            =   1590
            RightToLeft     =   -1  'True
            TabIndex        =   603
            Top             =   5760
            Width           =   1725
         End
         Begin VB.Label lbl 
            BackColor       =   &H00E2E9E9&
            Caption         =   "Public key certpem"
            Height          =   495
            Index           =   81
            Left            =   1590
            RightToLeft     =   -1  'True
            TabIndex        =   602
            Top             =   6480
            Width           =   1725
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "«·‰Ê⁄ «·«› —«÷Ì"
            Height          =   375
            Index           =   90
            Left            =   5520
            RightToLeft     =   -1  'True
            TabIndex        =   599
            Top             =   3180
            Width           =   1485
         End
         Begin VB.Label lbl 
            BackColor       =   &H00E2E9E9&
            Caption         =   "CSR"
            Height          =   375
            Index           =   79
            Left            =   1620
            RightToLeft     =   -1  'True
            TabIndex        =   597
            Top             =   4800
            Width           =   1725
         End
         Begin VB.Label lbl 
            BackColor       =   &H00E2E9E9&
            Caption         =   "Invoice type"
            Height          =   375
            Index           =   86
            Left            =   1620
            RightToLeft     =   -1  'True
            TabIndex        =   596
            Top             =   3240
            Width           =   1725
         End
         Begin VB.Label lbl 
            BackColor       =   &H00E2E9E9&
            Caption         =   "Location"
            Height          =   375
            Index           =   87
            Left            =   1620
            RightToLeft     =   -1  'True
            TabIndex        =   595
            Top             =   3720
            Width           =   1725
         End
         Begin VB.Label lbl 
            BackColor       =   &H00E2E9E9&
            Caption         =   "Industry"
            Height          =   255
            Index           =   88
            Left            =   1620
            RightToLeft     =   -1  'True
            TabIndex        =   594
            Top             =   4200
            Width           =   1725
         End
         Begin VB.Label lbl 
            BackColor       =   &H00E2E9E9&
            Caption         =   "Common Name"
            Height          =   375
            Index           =   67
            Left            =   1650
            RightToLeft     =   -1  'True
            TabIndex        =   589
            Top             =   690
            Width           =   1725
         End
         Begin VB.Label lbl 
            BackColor       =   &H00E2E9E9&
            Caption         =   "Serial Number"
            Height          =   375
            Index           =   68
            Left            =   1650
            RightToLeft     =   -1  'True
            TabIndex        =   588
            Top             =   1170
            Width           =   1725
         End
         Begin VB.Label lbl 
            BackColor       =   &H00E2E9E9&
            Caption         =   "Org Identifier"
            Height          =   255
            Index           =   69
            Left            =   1650
            RightToLeft     =   -1  'True
            TabIndex        =   587
            Top             =   1650
            Width           =   1725
         End
         Begin VB.Label lbl 
            BackColor       =   &H00E2E9E9&
            Caption         =   "Org Unit name"
            Height          =   375
            Index           =   83
            Left            =   1650
            RightToLeft     =   -1  'True
            TabIndex        =   586
            Top             =   2205
            Width           =   1725
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "«·„—Õ·… «·À«‰·Ì…"
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
            Left            =   6480
            RightToLeft     =   -1  'True
            TabIndex        =   277
            Top             =   240
            Width           =   5760
         End
      End
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   "«·Â« ›"
      Height          =   375
      Index           =   19
      Left            =   7560
      RightToLeft     =   -1  'True
      TabIndex        =   3
      Top             =   1920
      Width           =   1245
   End
End
Attribute VB_Name = "FrmOptions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Dim rs As ADODB.Recordset
Dim Askinterval As String
Dim Askcount As Integer
Dim checksave As Boolean

Private Sub ALLButton1_Click()

    If checkApility("FrmyaersData") = False Then
        Exit Sub
    End If

    If bigUser = False Then
        MsgBox "€Ì— „”„ÊÕ ·ﬂ »«· ⁄«„· „⁄ Â–Â «·‰«›–…", vbOKOnly + vbCritical + vbMsgBoxRight + vbMsgBoxRtlReading, "’·«ÕÌ«  «·„” Œœ„Ì‰"
        Exit Sub
    End If
            
    FrmyaersData.show
    
  '  DB_CreateField "TblEmployee", "chkAllowEditPaymentCont", adBoolean, adColNullable, , "0", "        ", False, True
    
    'FrmAccountIntervals.Show
End Sub

Private Sub Check3_Click()
        
    mdifrmmain.DockingPane1.FindPane(DockingPanesIDs.ItemsTreeID).Closed = Not mdifrmmain.DockingPane1.FindPane(DockingPanesIDs.ItemsTreeID).Closed
    
End Sub

Private Sub Chkbarcode_Click(Index As Integer)
If Index = 197 Then
Chkbarcode(197).Visible = False
ElseIf Index = 209 Then
    If Chkbarcode(209).value Then
        Chkbarcode(19).value = False
    End If
ElseIf Index = 19 Then
    If Chkbarcode(19).value Then
        Chkbarcode(209).value = False
    End If
ElseIf Index = 208 Then
    If Chkbarcode(208).value Then
        chkCustomerhavethreeAccounts.value = False
    End If


End If

End Sub

Private Sub ChkCalender_Click()
      
    mdifrmmain.DockingPane1.FindPane(DockingPanesIDs.CalendarPaneID).Closed = IIf(ChkCalender.value = 1, 0, 1) 'Not MDIFrmMain.DockingPane1.FindPane(DockingPanesIDs.CalendarPaneID).Closed
End Sub

Private Sub chkCustomerhavethreeAccounts_Click()

    If chkCustomerhavethreeAccounts.value Then
        Chkbarcode(208).value = False
    End If
End Sub

Private Sub Chkgraphic_Click()
    mdifrmmain.DockingPane1.FindPane(DockingPanesIDs.NewsBarPaneID).Closed = IIf(Chkgraphic.value = 1, 0, 1)
End Sub

Private Sub ChKHR_Click()

    If ChKHR.value = Checked Then
        Fra(10).Enabled = True
    Else
        Fra(10).Enabled = False
    End If

End Sub

Private Sub chkshortCuts_Click()
     mdifrmmain.DockingPane1.FindPane(DockingPanesIDs.OutBarPaneID).Closed = IIf(chkshortCuts.value = 1, 0, 1) ' Not MDIFrmMain.DockingPane1.FindPane(DockingPanesIDs.OutBarPaneID).Closed
End Sub

Private Sub ChkTax_Click()
'    On Error GoTo ErrTrap

'    If mdifrmmain.ActiveForm Is Nothing Then Exit Sub
'    If mdifrmmain.ActiveForm.name = "FrmSaleBill" Then
'        If ChkTax.value = vbChecked Then
'            frmsalebill.Ele(4).Visible = True
'        Else
'            frmsalebill.Ele(4).Visible = False
'        End If

  '  ElseIf mdifrmmain.ActiveForm.name = "FrmReports" Then
'
'        If ChkTax.value = vbChecked Then
'            FrmReports.C1TabMain.TabVisible(3) = True
'        Else
'            FrmReports.C1TabMain.TabVisible(3) = False
'        End If
'    End If
'
'    Exit Sub
'ErrTrap:
End Sub

Private Sub ChkShowToolTip_Click()
mdifrmmain.DockingPane1.FindPane(DockingPanesIDs.DynamicHelp).Closed = IIf(ChkShowToolTip.value = 1, 0, 1)
End Sub

Private Sub Chktree_Click()
    mdifrmmain.DockingPane1.FindPane(DockingPanesIDs.ItemsTreeID).Closed = IIf(Chktree.value = 1, 0, 1) 'Not MDIFrmMain.DockingPane1.FindPane(DockingPanesIDs.ItemsTreeID).Closed
End Sub

Private Sub Cmd_Click()

    With cdg
        '*.jpg,*.jpeg,*.jpe,*.jfif
        .CancelError = False
        .DialogTitle = " ≈Œ Ì«— ’Ê—…"
        'Set The Filter to show pictures only
        .filter = "Bitmap (*.bmp)|*.bmp|JPEG(*.JPG,*.JPEG,*.JPE,*.JFIF)|*.jpg;*.jpeg;*.jpe;*.jfif|" & "GIF (*.gif)|*.gif|All Files|*.*" ' choose formats to include
        .ShowOpen

        If .FileName <> "" Then
            Set Me.ImgPic.Picture = LoadPicture(.FileName)
        End If

    End With

End Sub

Private Sub CmdCancel_Click()
    Unload Me
End Sub

Function CuurentLogdata(Optional Currentmode As String)
'Exit Function
ScreenNameArabic = Me.Caption
ScreenNameEnglish = "Options "


    LogTextA = "    ‘«‘…  " & ScreenNameArabic & CHR(13)
    LogTextA = LogTextA & " " & lbl(10).Caption & " " & XPTxtCompany.text & CHR(13)
    LogTextA = LogTextA & " " & lbl(8).Caption & " " & XPTxtCompanye.text & CHR(13)
    LogTextA = LogTextA & " " & lbl(21).Caption & " " & XPTxtComment(0).text & CHR(13)
    LogTextA = LogTextA & " " & lbl(9).Caption & " " & XPTxtAddress.text & CHR(13)
    LogTextA = LogTextA & " " & lbl(20).Caption & " " & TxtEmails.text & CHR(13)
    LogTextA = LogTextA & " " & lbl(5).Caption & " " & xptxtphone.text & CHR(13)
    LogTextA = LogTextA & " " & lbl(18).Caption & " " & TxtFax.text & CHR(13)
    LogTextA = LogTextA & " " & lbl(2).Caption & " " & XPTxtResponsable.text & CHR(13)
    
   LogTextA = LogTextA & " " & lbl(3).Caption & " " & XPTxtMail.text & CHR(13)
   LogTextA = LogTextA & " " & lbl(1).Caption & " " & XPTxtmobile.text & CHR(13)
   
 If chk.value = vbChecked Then
   LogTextA = LogTextA & " " & chk.Caption & " ‰⁄„" & CHR(13)
  Else
  LogTextA = LogTextA & " " & chk.Caption & " ·«" & CHR(13)
  End If
  
    LogTextA = LogTextA & " " & lbl(2).Caption & " " & CboMainStockType.text & CHR(13)
   
   
   LogTextA = LogTextA & " " & lbl(25).Caption & " " & CboChasingStatus.text & CHR(13)
   
   
   
  
   
   If chk2.value = vbChecked Then
   LogTextA = LogTextA & " " & chk2.Caption & " ‰⁄„" & CHR(13)
  Else
  LogTextA = LogTextA & " " & chk2.Caption & " ·«" & CHR(13)
  End If
  
   If ChkitemsWorkWithSize.value = vbChecked Then
   LogTextA = LogTextA & " " & ChkitemsWorkWithSize.Caption & " ‰⁄„" & CHR(13)
  Else
  LogTextA = LogTextA & " " & ChkitemsWorkWithSize.Caption & " ·«" & CHR(13)
  End If
  
     If ChkitemsWorkWithColor.value = vbChecked Then
   LogTextA = LogTextA & " " & ChkitemsWorkWithColor.Caption & " ‰⁄„" & CHR(13)
  Else
  LogTextA = LogTextA & " " & ChkitemsWorkWithColor.Caption & " ·«" & CHR(13)
  End If
  
       If ChkitemsWorkWithClass.value = vbChecked Then
   LogTextA = LogTextA & " " & ChkitemsWorkWithClass.Caption & " ‰⁄„" & CHR(13)
  Else
  LogTextA = LogTextA & " " & ChkitemsWorkWithClass.Caption & " ·«" & CHR(13)
  End If
  
  
        If ChkitemsWorkWithDates.value = vbChecked Then
   LogTextA = LogTextA & " " & ChkitemsWorkWithDates.Caption & " ‰⁄„" & CHR(13)
  Else
  LogTextA = LogTextA & " " & ChkitemsWorkWithDates.Caption & " ·«" & CHR(13)
  End If
  
          If Chkbarcode(0).value = vbChecked Then
   LogTextA = LogTextA & " " & Chkbarcode(0).Caption & " ‰⁄„" & CHR(13)
  Else
  LogTextA = LogTextA & " " & Chkbarcode(0).Caption & " ·«" & CHR(13)
  End If
  
          If Chkbarcode(38).value = vbChecked Then
   LogTextA = LogTextA & " " & Chkbarcode(38).Caption & " ‰⁄„" & CHR(13)
  Else
  LogTextA = LogTextA & " " & Chkbarcode(38).Caption & " ·«" & CHR(13)
  End If
    
    
             If Chkbarcode(7).value = vbChecked Then
   LogTextA = LogTextA & " " & Chkbarcode(7).Caption & " ‰⁄„" & CHR(13)
  Else
  LogTextA = LogTextA & " " & Chkbarcode(7).Caption & " ·«" & CHR(13)
  End If
  
  
              If Chkbarcode(1).value = vbChecked Then
   LogTextA = LogTextA & " " & Chkbarcode(1).Caption & " ‰⁄„" & CHR(13)
  Else
  LogTextA = LogTextA & " " & Chkbarcode(1).Caption & " ·«" & CHR(13)
  End If
  
               If Chkbarcode(12).value = vbChecked Then
   LogTextA = LogTextA & " " & Chkbarcode(12).Caption & " ‰⁄„" & CHR(13)
  Else
  LogTextA = LogTextA & " " & Chkbarcode(12).Caption & " ·«" & CHR(13)
  End If
  
               If Chkbarcode(7).value = vbChecked Then
   LogTextA = LogTextA & " " & Chkbarcode(7).Caption & " ‰⁄„" & CHR(13)
  Else
  LogTextA = LogTextA & " " & Chkbarcode(7).Caption & " ·«" & CHR(13)
  End If
  
 
   LogTextA = LogTextA & " " & Frame37.Caption & CHR(13)
   
   If OptCurrQty(0).value = True Then
    LogTextA = LogTextA & " " & OptCurrQty(0).value & CHR(13)
   Else
    LogTextA = LogTextA & " " & OptCurrQty(1).value & CHR(13)
   End If
   
   
   
 
                 If chkStore(0).value = vbChecked Then
   LogTextA = LogTextA & " " & chkStore(0).Caption & " ‰⁄„" & CHR(13)
  Else
  LogTextA = LogTextA & " " & chkStore(0).Caption & " ·«" & CHR(13)
  End If
  
   
                 If chkStore(1).value = vbChecked Then
   LogTextA = LogTextA & " " & chkStore(1).Caption & " ‰⁄„" & CHR(13)
  Else
  LogTextA = LogTextA & " " & chkStore(1).Caption & " ·«" & CHR(13)
  End If
  
   
    If chkStore(2).value = vbChecked Then
   LogTextA = LogTextA & " " & chkStore(2).Caption & " ‰⁄„" & CHR(13)
  Else
  LogTextA = LogTextA & " " & chkStore(2).Caption & " ·«" & CHR(13)
  End If
If chkStore(3).value = vbChecked Then
   LogTextA = LogTextA & " " & chkStore(3).Caption & " ‰⁄„" & CHR(13)
  Else
  LogTextA = LogTextA & " " & chkStore(3).Caption & " ·«" & CHR(13)
  End If
  If chkStore(5).value = vbChecked Then
   LogTextA = LogTextA & " " & chkStore(5).Caption & " ‰⁄„" & CHR(13)
  Else
  LogTextA = LogTextA & " " & chkStore(5).Caption & " ·«" & CHR(13)
  End If
  If chkStore(6).value = vbChecked Then
   LogTextA = LogTextA & " " & chkStore(6).Caption & " ‰⁄„" & CHR(13)
  Else
  LogTextA = LogTextA & " " & chkStore(6).Caption & " ·«" & CHR(13)
  End If
   
                 If ChkCostStarting.value = vbChecked Then
   LogTextA = LogTextA & " " & ChkCostStarting.Caption & " ‰⁄„" & CHR(13)
  Else
  LogTextA = LogTextA & " " & ChkCostStarting.Caption & " ·«" & CHR(13)
  End If
  
  
  LogTextA = LogTextA & " " & lbl(7).Caption & " " & DBCboClientName.text & CHR(13)
   LogTextA = LogTextA & " " & lbl(4).Caption & " " & DCboStoreName(0).text & CHR(13)
  LogTextA = LogTextA & " " & lbl(11).Caption & " " & DcboBox.text & CHR(13)
  
  
                 If Chkbarcode(13).value = vbChecked Then
   LogTextA = LogTextA & " " & Chkbarcode(13).Caption & " ‰⁄„" & CHR(13)
  Else
  LogTextA = LogTextA & " " & Chkbarcode(13).Caption & " ·«" & CHR(13)
  End If
   If Chkbarcode(137).value = vbChecked Then
   LogTextA = LogTextA & " " & Chkbarcode(137).Caption & " ‰⁄„" & CHR(13)
  Else
  LogTextA = LogTextA & " " & Chkbarcode(137).Caption & " ·«" & CHR(13)
  End If
If Chkbarcode(138).value = vbChecked Then
   LogTextA = LogTextA & " " & Chkbarcode(138).Caption & " ‰⁄„" & CHR(13)
  Else
  LogTextA = LogTextA & " " & Chkbarcode(138).Caption & " ·«" & CHR(13)
  End If
  If Chkbarcode(139).value = vbChecked Then
   LogTextA = LogTextA & " " & Chkbarcode(139).Caption & " ‰⁄„" & CHR(13)
  Else
  LogTextA = LogTextA & " " & Chkbarcode(139).Caption & " ·«" & CHR(13)
  End If
                   If Chkbarcode(46).value = vbChecked Then
   LogTextA = LogTextA & " " & Chkbarcode(46).Caption & " ‰⁄„" & CHR(13)
  Else
  LogTextA = LogTextA & " " & Chkbarcode(46).Caption & " ·«" & CHR(13)
  End If
  
                   If Chkbarcode(47).value = vbChecked Then
   LogTextA = LogTextA & " " & Chkbarcode(47).Caption & " ‰⁄„" & CHR(13)
  Else
  LogTextA = LogTextA & " " & Chkbarcode(47).Caption & " ·«" & CHR(13)
  End If
  
 
   LogTextA = LogTextA & " " & Fra(0).Caption & CHR(13)
   
   If opt(0).value = True Then
    LogTextA = LogTextA & " " & opt(0).value & CHR(13)
  ElseIf opt(1).value = True Then
    LogTextA = LogTextA & " " & opt(1).value & CHR(13)
    
   Else
    LogTextA = LogTextA & " " & opt(2).value & CHR(13)
   End If
   
   
      If opt(8).value = True Then
    LogTextA = LogTextA & " " & opt(8).value & CHR(13)
  ElseIf opt(9).value = True Then
    LogTextA = LogTextA & " " & opt(9).value & CHR(13)
    
   Else
    LogTextA = LogTextA & " " & opt(10).value & CHR(13)
   End If
   
   
   
   
  LogTextA = LogTextA & " " & Fra(7).Caption & CHR(13)
  If Opt_OrderOut.value = True Then
    LogTextA = LogTextA & " " & Opt_OrderOut.value & CHR(13)

   End If
   
   
   
   If ChKautoIssueVoucher.value = vbChecked Then
   LogTextA = LogTextA & " " & ChKautoIssueVoucher.Caption & " ‰⁄„" & CHR(13)
  Else
  LogTextA = LogTextA & " " & ChKautoIssueVoucher.Caption & " ·«" & CHR(13)
  End If
  
  
     LogTextA = LogTextA & " " & Fra(3).Caption & CHR(13)
   
   If opt(6).value = True Then
    LogTextA = LogTextA & " " & opt(6).value & CHR(13)
  Else: opt(7).value = True
    LogTextA = LogTextA & " " & opt(7).value & CHR(13)
    
   LogTextA = LogTextA & " " & Label7.Caption & " " & TXTReturnSallingIntervalCount(0).text & " " & lbl(37).Caption & CHR(13)
   LogTextA = LogTextA & " " & Label9.Caption & " " & TXTReturnSallingIntervalCount1.text & " " & Label5.Caption & CHR(13)
   LogTextA = LogTextA & " " & Chkbarcode(69).Caption & " " & TXTReturnSallingIntervalCount(3).text
   LogTextA = LogTextA & " " & Chkbarcode(82).Caption & " " & TXTReturnSallingIntervalCount(4).text
   End If
   
  
     LogTextA = LogTextA & " " & Frame28.Caption & CHR(13)
   
     LogTextA = LogTextA & " " & Label3.Caption & " " & TxtSaleDiscount1.text
    LogTextA = LogTextA & " " & Label4.Caption & " " & TxtSaleDiscount2.text
    LogTextA = LogTextA & " " & Label6.Caption & " " & TxtSaleDiscount3.text
    LogTextA = LogTextA & " " & Label1(0).Caption & " " & TxtSaleDiscount4.text
    
    LogTextA = LogTextA & " " & Label16.Caption & " " & txtLimitDefaultCredit.text
    LogTextA = LogTextA & " " & Label19.Caption & " " & txtLimitDefaultCreditDays.text
    
    
     If ChkItemsattachedzero.value = vbChecked Then
   LogTextA = LogTextA & " " & ChkItemsattachedzero.Caption & " ‰⁄„" & CHR(13)
  Else
  LogTextA = LogTextA & " " & ChkItemsattachedzero.Caption & " ·«" & CHR(13)
  End If
  
   
     If Chkbarcode(17).value = vbChecked Then
   LogTextA = LogTextA & " " & Chkbarcode(17).Caption & " ‰⁄„" & CHR(13)
  Else
  LogTextA = LogTextA & " " & Chkbarcode(17).Caption & " ·«" & CHR(13)
  End If
     
     If Chkbarcode(34).value = vbChecked Then
   LogTextA = LogTextA & " " & Chkbarcode(34).Caption & " ‰⁄„" & CHR(13)
  Else
  LogTextA = LogTextA & " " & Chkbarcode(34).Caption & " ·«" & CHR(13)
  End If
  
  
     If Chkbarcode(2).value = vbChecked Then
   LogTextA = LogTextA & " " & Chkbarcode(2).Caption & " ‰⁄„" & CHR(13)
  Else
  LogTextA = LogTextA & " " & Chkbarcode(2).Caption & " ·«" & CHR(13)
  End If
  
  If Chkbarcode(87).value = vbChecked Then
   LogTextA = LogTextA & " " & Chkbarcode(87).Caption & " ‰⁄„" & CHR(13)
  Else
  LogTextA = LogTextA & " " & Chkbarcode(87).Caption & " ·«" & CHR(13)
  End If
  If Chkbarcode(88).value = vbChecked Then
   LogTextA = LogTextA & " " & Chkbarcode(88).Caption & " ‰⁄„" & CHR(13)
  Else
  LogTextA = LogTextA & " " & Chkbarcode(88).Caption & " ·«" & CHR(13)
  End If
  If Chkbarcode(89).value = vbChecked Then
   LogTextA = LogTextA & " " & Chkbarcode(89).Caption & " ‰⁄„" & CHR(13)
  Else
  LogTextA = LogTextA & " " & Chkbarcode(89).Caption & " ·«" & CHR(13)
  End If
  If Chkbarcode(90).value = vbChecked Then
   LogTextA = LogTextA & " " & Chkbarcode(90).Caption & " ‰⁄„" & CHR(13)
  Else
  LogTextA = LogTextA & " " & Chkbarcode(90).Caption & " ·«" & CHR(13)
  End If
    If Chkbarcode(91).value = vbChecked Then
   LogTextA = LogTextA & " " & Chkbarcode(91).Caption & " ‰⁄„" & CHR(13)
  Else
  LogTextA = LogTextA & " " & Chkbarcode(91).Caption & " ·«" & CHR(13)
  End If
    If Chkbarcode(92).value = vbChecked Then
   LogTextA = LogTextA & " " & Chkbarcode(92).Caption & " ‰⁄„" & CHR(13)
  Else
  LogTextA = LogTextA & " " & Chkbarcode(92).Caption & " ·«" & CHR(13)
  End If
If Chkbarcode(93).value = vbChecked Then
   LogTextA = LogTextA & " " & Chkbarcode(93).Caption & " ‰⁄„" & CHR(13)
  Else
  LogTextA = LogTextA & " " & Chkbarcode(93).Caption & " ·«" & CHR(13)
  End If
If Chkbarcode(94).value = vbChecked Then
   LogTextA = LogTextA & " " & Chkbarcode(94).Caption & " ‰⁄„" & CHR(13)
  Else
  LogTextA = LogTextA & " " & Chkbarcode(94).Caption & " ·«" & CHR(13)
  End If
  If Chkbarcode(95).value = vbChecked Then
   LogTextA = LogTextA & " " & Chkbarcode(95).Caption & " ‰⁄„" & CHR(13)
  Else
  LogTextA = LogTextA & " " & Chkbarcode(95).Caption & " ·«" & CHR(13)
  End If
  If Chkbarcode(96).value = vbChecked Then
   LogTextA = LogTextA & " " & Chkbarcode(96).Caption & " ‰⁄„" & CHR(13)
  Else
  LogTextA = LogTextA & " " & Chkbarcode(96).Caption & " ·«" & CHR(13)
  End If
    If Chkbarcode(97).value = vbChecked Then
   LogTextA = LogTextA & " " & Chkbarcode(97).Caption & " ‰⁄„" & CHR(13)
  Else
  LogTextA = LogTextA & " " & Chkbarcode(97).Caption & " ·«" & CHR(13)
  End If
 If Chkbarcode(98).value = vbChecked Then
   LogTextA = LogTextA & " " & Chkbarcode(98).Caption & " ‰⁄„" & CHR(13)
  Else
  LogTextA = LogTextA & " " & Chkbarcode(98).Caption & " ·«" & CHR(13)
  End If
  If Chkbarcode(99).value = vbChecked Then
   LogTextA = LogTextA & " " & Chkbarcode(99).Caption & " ‰⁄„" & CHR(13)
  Else
   LogTextA = LogTextA & " " & Chkbarcode(99).Caption & " ·«" & CHR(13)
  End If
  If Chkbarcode(101).value = vbChecked Then
   LogTextA = LogTextA & " " & Chkbarcode(101).Caption & " ‰⁄„" & CHR(13)
  Else
   LogTextA = LogTextA & " " & Chkbarcode(101).Caption & " ·«" & CHR(13)
  End If
  If Chkbarcode(102).value = vbChecked Then
   LogTextA = LogTextA & " " & Chkbarcode(102).Caption & " ‰⁄„" & CHR(13)
  Else
   LogTextA = LogTextA & " " & Chkbarcode(102).Caption & " ·«" & CHR(13)
  End If
  If Chkbarcode(103).value = vbChecked Then
   LogTextA = LogTextA & " " & Chkbarcode(103).Caption & " ‰⁄„" & CHR(13)
  Else
   LogTextA = LogTextA & " " & Chkbarcode(103).Caption & " ·«" & CHR(13)
  End If
    If Chkbarcode(104).value = vbChecked Then
   LogTextA = LogTextA & " " & Chkbarcode(104).Caption & " ‰⁄„" & CHR(13)
  Else
   LogTextA = LogTextA & " " & Chkbarcode(104).Caption & " ·«" & CHR(13)
  End If
    If Chkbarcode(105).value = vbChecked Then
   LogTextA = LogTextA & " " & Chkbarcode(105).Caption & " ‰⁄„" & CHR(13)
  Else
   LogTextA = LogTextA & " " & Chkbarcode(105).Caption & " ·«" & CHR(13)
  End If
If Chkbarcode(106).value = vbChecked Then
   LogTextA = LogTextA & " " & Chkbarcode(106).Caption & " ‰⁄„" & CHR(13)
  Else
   LogTextA = LogTextA & " " & Chkbarcode(106).Caption & " ·«" & CHR(13)
  End If
  If Chkbarcode(15).value = vbChecked Then
   LogTextA = LogTextA & " " & Chkbarcode(15).Caption & " ‰⁄„" & CHR(13)
  Else
   LogTextA = LogTextA & " " & Chkbarcode(15).Caption & " ·«" & CHR(13)
  End If

       
    If Chkbarcode(18).value = vbChecked Then
   LogTextA = LogTextA & " " & Chkbarcode(18).Caption & " ‰⁄„" & CHR(13)
  Else
  LogTextA = LogTextA & " " & Chkbarcode(18).Caption & " ·«" & CHR(13)
  End If
  
  
    If Chkbarcode(22).value = vbChecked Then
   LogTextA = LogTextA & " " & Chkbarcode(22).Caption & " ‰⁄„" & CHR(13)
  Else
  LogTextA = LogTextA & " " & Chkbarcode(22).Caption & " ·«" & CHR(13)
  End If
  
    
     LogTextA = LogTextA & " " & lbl(6).Caption & " " & DBCboSupName.text & CHR(13)
    LogTextA = LogTextA & " " & lbl(0).Caption & " " & DCboStoreName(1).text & CHR(13)
    
          
          
          LogTextA = LogTextA & " " & Fra(1).Caption & CHR(13)
 If opt(5).value = True Then
    LogTextA = LogTextA & " " & opt(5).value & CHR(13)
  ElseIf opt(4).value = True Then
    LogTextA = LogTextA & " " & opt(4).value & CHR(13)
    
   Else
    LogTextA = LogTextA & " " & opt(3).value & CHR(13)
   End If
   
   
             LogTextA = LogTextA & " " & Frame2.Caption & CHR(13)
             LogTextA = LogTextA & " " & Opt_OrderInpo.Caption & CHR(13)
             
      If ChKautoReseiveVoucher.value = vbChecked Then
   LogTextA = LogTextA & " " & ChKautoReseiveVoucher.Caption & " ‰⁄„" & CHR(13)
  Else
  LogTextA = LogTextA & " " & ChKautoReseiveVoucher.Caption & " ·«" & CHR(13)
  End If
       
       
        If Chkbarcode(21).value = vbChecked Then
   LogTextA = LogTextA & " " & Chkbarcode(21).Caption & " ‰⁄„" & CHR(13)
  Else
  LogTextA = LogTextA & " " & Chkbarcode(21).Caption & " ·«" & CHR(13)
  End If
  
  
        If Chkbarcode(25).value = vbChecked Then
   LogTextA = LogTextA & " " & Chkbarcode(25).Caption & " ‰⁄„" & CHR(13)
  Else
  LogTextA = LogTextA & " " & Chkbarcode(25).Caption & " ·«" & CHR(13)
  End If
  
        If Chkbarcode(26).value = vbChecked Then
   LogTextA = LogTextA & " " & Chkbarcode(26).Caption & " ‰⁄„" & CHR(13)
  Else
  LogTextA = LogTextA & " " & Chkbarcode(26).Caption & " ·«" & CHR(13)
  End If
  
        If Chkbarcode(27).value = vbChecked Then
   LogTextA = LogTextA & " " & Chkbarcode(27).Caption & " ‰⁄„" & CHR(13)
  Else
  LogTextA = LogTextA & " " & Chkbarcode(27).Caption & " ·«" & CHR(13)
  End If
   
 If Chkbarcode(114).value = vbChecked Then
   LogTextA = LogTextA & " " & Chkbarcode(114).Caption & " ‰⁄„" & CHR(13)
  Else
  LogTextA = LogTextA & " " & Chkbarcode(114).Caption & " ·«" & CHR(13)
  End If
  
 If Chkbarcode(115).value = vbChecked Then
   LogTextA = LogTextA & " " & Chkbarcode(115).Caption & " ‰⁄„" & CHR(13)
  Else
  LogTextA = LogTextA & " " & Chkbarcode(115).Caption & " ·«" & CHR(13)
  End If
   If Chkbarcode(116).value = vbChecked Then
   LogTextA = LogTextA & " " & Chkbarcode(116).Caption & " ‰⁄„" & CHR(13)
  Else
  LogTextA = LogTextA & " " & Chkbarcode(116).Caption & " ·«" & CHR(13)
  End If
   If Chkbarcode(117).value = vbChecked Then
   LogTextA = LogTextA & " " & Chkbarcode(117).Caption & " ‰⁄„" & CHR(13)
  Else
  LogTextA = LogTextA & " " & Chkbarcode(117).Caption & " ·«" & CHR(13)
  End If
  
 If Chkbarcode(118).value = vbChecked Then
   LogTextA = LogTextA & " " & Chkbarcode(118).Caption & " ‰⁄„" & CHR(13)
  Else
  LogTextA = LogTextA & " " & Chkbarcode(118).Caption & " ·«" & CHR(13)
  End If
   If Chkbarcode(119).value = vbChecked Then
   LogTextA = LogTextA & " " & Chkbarcode(119).Caption & " ‰⁄„" & CHR(13)
  Else
  LogTextA = LogTextA & " " & Chkbarcode(119).Caption & " ·«" & CHR(13)
  End If
   If Chkbarcode(120).value = vbChecked Then
   LogTextA = LogTextA & " " & Chkbarcode(120).Caption & " ‰⁄„" & CHR(13)
  Else
  LogTextA = LogTextA & " " & Chkbarcode(120).Caption & " ·«" & CHR(13)
  End If
 If Chkbarcode(121).value = vbChecked Then
   LogTextA = LogTextA & " " & Chkbarcode(121).Caption & " ‰⁄„" & CHR(13)
  Else
  LogTextA = LogTextA & " " & Chkbarcode(121).Caption & " ·«" & CHR(13)
  End If
  
   If Chkbarcode(122).value = vbChecked Then
   LogTextA = LogTextA & " " & Chkbarcode(122).Caption & " ‰⁄„" & CHR(13)
  Else
  LogTextA = LogTextA & " " & Chkbarcode(122).Caption & " ·«" & CHR(13)
  End If
   If Chkbarcode(123).value = vbChecked Then
   LogTextA = LogTextA & " " & Chkbarcode(123).Caption & " ‰⁄„" & CHR(13)
  Else
  LogTextA = LogTextA & " " & Chkbarcode(123).Caption & " ·«" & CHR(13)
  End If
   If Chkbarcode(124).value = vbChecked Then
   LogTextA = LogTextA & " " & Chkbarcode(124).Caption & " ‰⁄„" & CHR(13)
  Else
  LogTextA = LogTextA & " " & Chkbarcode(124).Caption & " ·«" & CHR(13)
  End If
If Chkbarcode(125).value = vbChecked Then
   LogTextA = LogTextA & " " & Chkbarcode(125).Caption & " ‰⁄„" & CHR(13)
  Else
  LogTextA = LogTextA & " " & Chkbarcode(125).Caption & " ·«" & CHR(13)
  End If
  If Chkbarcode(126).value = vbChecked Then
   LogTextA = LogTextA & " " & Chkbarcode(126).Caption & " ‰⁄„" & CHR(13)
  Else
  LogTextA = LogTextA & " " & Chkbarcode(126).Caption & " ·«" & CHR(13)
  End If
    If Chkbarcode(127).value = vbChecked Then
   LogTextA = LogTextA & " " & Chkbarcode(127).Caption & " ‰⁄„" & CHR(13)
  Else
  LogTextA = LogTextA & " " & Chkbarcode(127).Caption & " ·«" & CHR(13)
  End If
If Chkbarcode(128).value = vbChecked Then
   LogTextA = LogTextA & " " & Chkbarcode(128).Caption & " ‰⁄„" & CHR(13)
  Else
  LogTextA = LogTextA & " " & Chkbarcode(128).Caption & " ·«" & CHR(13)
  End If
  
If Chkbarcode(129).value = vbChecked Then
   LogTextA = LogTextA & " " & Chkbarcode(129).Caption & " ‰⁄„" & CHR(13)
  Else
  LogTextA = LogTextA & " " & Chkbarcode(129).Caption & " ·«" & CHR(13)
  End If
  
  If Chkbarcode(130).value = vbChecked Then
   LogTextA = LogTextA & " " & Chkbarcode(130).Caption & " ‰⁄„" & CHR(13)
  Else
  LogTextA = LogTextA & " " & Chkbarcode(130).Caption & " ·«" & CHR(13)
  End If
  
  If Chkbarcode(131).value = vbChecked Then
   LogTextA = LogTextA & " " & Chkbarcode(131).Caption & " ‰⁄„" & CHR(13)
  Else
  LogTextA = LogTextA & " " & Chkbarcode(131).Caption & " ·«" & CHR(13)
  End If
  
  
  If Chkbarcode(133).value = vbChecked Then
   LogTextA = LogTextA & " " & Chkbarcode(133).Caption & " ‰⁄„" & CHR(13)
  Else
  LogTextA = LogTextA & " " & Chkbarcode(133).Caption & " ·«" & CHR(13)
  End If
 
 '134             136
  
 
  
          If Chkbarcode(28).value = vbChecked Then
   LogTextA = LogTextA & " " & Chkbarcode(28).Caption & " ‰⁄„" & CHR(13)
  Else
  LogTextA = LogTextA & " " & Chkbarcode(28).Caption & " ·«" & CHR(13)
  End If
  
          If Chkbarcode(39).value = vbChecked Then
   LogTextA = LogTextA & " " & Chkbarcode(39).Caption & " ‰⁄„" & CHR(13)
  Else
  LogTextA = LogTextA & " " & Chkbarcode(39).Caption & " ·«" & CHR(13)
  End If
  
            If Chkbarcode(40).value = vbChecked Then
   LogTextA = LogTextA & " " & Chkbarcode(40).Caption & " ‰⁄„" & CHR(13)
  Else
  LogTextA = LogTextA & " " & Chkbarcode(40).Caption & " ·«" & CHR(13)
  End If
  
  
          If Chkbarcode(29).value = vbChecked Then
   LogTextA = LogTextA & " " & Chkbarcode(29).Caption & " ‰⁄„" & CHR(13)
  Else
  LogTextA = LogTextA & " " & Chkbarcode(29).Caption & " ·«" & CHR(13)
  End If
  
            If Chkbarcode(30).value = vbChecked Then
   LogTextA = LogTextA & " " & Chkbarcode(30).Caption & " ‰⁄„" & CHR(13)
  Else
  LogTextA = LogTextA & " " & Chkbarcode(30).Caption & " ·«" & CHR(13)
  End If
  
  
            If Chkbarcode(31).value = vbChecked Then
   LogTextA = LogTextA & " " & Chkbarcode(31).Caption & " ‰⁄„" & CHR(13)
  Else
  LogTextA = LogTextA & " " & Chkbarcode(31).Caption & " ·«" & CHR(13)
  End If
  
           
            If Chkbarcode(32).value = vbChecked Then
   LogTextA = LogTextA & " " & Chkbarcode(32).Caption & " ‰⁄„" & CHR(13)
  Else
  LogTextA = LogTextA & " " & Chkbarcode(32).Caption & " ·«" & CHR(13)
  End If
                     If Chkbarcode(33).value = vbChecked Then
   LogTextA = LogTextA & " " & Chkbarcode(33).Caption & " ‰⁄„" & CHR(13)
  Else
  LogTextA = LogTextA & " " & Chkbarcode(33).Caption & " ·«" & CHR(13)
  End If
                 
                 
             
             LogTextA = LogTextA & " " & Frame45.Caption & CHR(13)
         If Chk1.value = vbChecked Then
   LogTextA = LogTextA & " " & Chk1.Caption & " ‰⁄„" & CHR(13)
  Else
  LogTextA = LogTextA & " " & Chk1.Caption & " ·«" & CHR(13)
  End If
  
        If Chkbarcode(16).value = vbChecked Then
   LogTextA = LogTextA & " " & Chkbarcode(16).Caption & " ‰⁄„" & CHR(13)
  Else
  LogTextA = LogTextA & " " & Chkbarcode(16).Caption & " ·«" & CHR(13)
  End If
    
    
        If chk3.value = vbChecked Then
   LogTextA = LogTextA & " " & chk3.Caption & " ‰⁄„" & CHR(13)
  Else
  LogTextA = LogTextA & " " & chk3.Caption & " ·«" & CHR(13)
  End If
        
              If chkChequeBox.value = vbChecked Then
   LogTextA = LogTextA & " " & chkChequeBox.Caption & " ‰⁄„" & CHR(13)
  Else
  LogTextA = LogTextA & " " & chkChequeBox.Caption & " ·«" & CHR(13)
  End If
        
        
                If ChkBankComm.value = vbChecked Then
   LogTextA = LogTextA & " " & ChkBankComm.Caption & " ‰⁄„" & CHR(13)
  Else
  LogTextA = LogTextA & " " & ChkBankComm.Caption & " ·«" & CHR(13)
  End If
          
                If chkCustomerhavethreeAccounts.value = vbChecked Then
   LogTextA = LogTextA & " " & chkCustomerhavethreeAccounts.Caption & " ‰⁄„" & CHR(13)
  Else
  LogTextA = LogTextA & " " & chkCustomerhavethreeAccounts.Caption & " ·«" & CHR(13)
  End If
  
  
                If chkCustomerhavethreeAccounts1.value = vbChecked Then
   LogTextA = LogTextA & " " & chkCustomerhavethreeAccounts1.Caption & " ‰⁄„" & CHR(13)
  Else
  LogTextA = LogTextA & " " & chkCustomerhavethreeAccounts1.Caption & " ·«" & CHR(13)
  End If
  
  
     If ChkExpensesCoding.value = vbChecked Then
   LogTextA = LogTextA & " " & ChkExpensesCoding.Caption & " ‰⁄„" & CHR(13)
  Else
  LogTextA = LogTextA & " " & ChkExpensesCoding.Caption & " ·«" & CHR(13)
  End If
  
  
     If ChkExpensesCoding2.value = vbChecked Then
   LogTextA = LogTextA & " " & ChkExpensesCoding2.Caption & " ‰⁄„" & CHR(13)
  Else
  LogTextA = LogTextA & " " & ChkExpensesCoding2.Caption & " ·«" & CHR(13)
  End If
    
    
     If chkInstallmntsvchrCoding.value = vbChecked Then
   LogTextA = LogTextA & " " & chkInstallmntsvchrCoding.Caption & " ‰⁄„" & CHR(13)
  Else
  LogTextA = LogTextA & " " & chkInstallmntsvchrCoding.Caption & " ·«" & CHR(13)
  End If
        
        
        If Chkbarcode(4).value = vbChecked Then
   LogTextA = LogTextA & " " & Chkbarcode(4).Caption & " ‰⁄„" & CHR(13)
  Else
  LogTextA = LogTextA & " " & Chkbarcode(4).Caption & " ·«" & CHR(13)
  End If
        
    If Chkbarcode(3).value = vbChecked Then
   LogTextA = LogTextA & " " & Chkbarcode(3).Caption & " ‰⁄„" & CHR(13)
  Else
  LogTextA = LogTextA & " " & Chkbarcode(3).Caption & " ·«" & CHR(13)
  End If
  
  If Chkbarcode(107).value = vbChecked Then
    LogTextA = LogTextA & " " & Chkbarcode(107).Caption & " ‰⁄„" & CHR(13)
  Else
    LogTextA = LogTextA & " " & Chkbarcode(107).Caption & " ·«" & CHR(13)
  End If
   If Chkbarcode(108).value = vbChecked Then
    LogTextA = LogTextA & " " & Chkbarcode(108).Caption & " ‰⁄„" & CHR(13)
  Else
    LogTextA = LogTextA & " " & Chkbarcode(108).Caption & " ·«" & CHR(13)
  End If
    If Chkbarcode(109).value = vbChecked Then
    LogTextA = LogTextA & " " & Chkbarcode(109).Caption & " ‰⁄„" & CHR(13)
  Else
    LogTextA = LogTextA & " " & Chkbarcode(109).Caption & " ·«" & CHR(13)
  End If
  
  
    If Chkbarcode(110).value = vbChecked Then
    LogTextA = LogTextA & " " & Chkbarcode(110).Caption & " ‰⁄„" & CHR(13)
  Else
    LogTextA = LogTextA & " " & Chkbarcode(110).Caption & " ·«" & CHR(13)
  End If
    
    If Chkbarcode(111).value = vbChecked Then
    LogTextA = LogTextA & " " & Chkbarcode(111).Caption & " ‰⁄„" & CHR(13)
  Else
    LogTextA = LogTextA & " " & Chkbarcode(111).Caption & " ·«" & CHR(13)
  End If
  
    
    
  If Chkbarcode(77).value = vbChecked Then
   LogTextA = LogTextA & " " & Chkbarcode(77).Caption & " ‰⁄„" & CHR(13)
  Else
  LogTextA = LogTextA & " " & Chkbarcode(77).Caption & " ·«" & CHR(13)
  End If
 If Chkbarcode(78).value = vbChecked Then
   LogTextA = LogTextA & " " & Chkbarcode(78).Caption & " ‰⁄„" & CHR(13)
  Else
  LogTextA = LogTextA & " " & Chkbarcode(78).Caption & " ·«" & CHR(13)
  End If
  If Chkbarcode(79).value = vbChecked Then
   LogTextA = LogTextA & " " & Chkbarcode(79).Caption & " ‰⁄„" & CHR(13)
  Else
  LogTextA = LogTextA & " " & Chkbarcode(79).Caption & " ·«" & CHR(13)
  End If
  If Chkbarcode(80).value = vbChecked Then
   LogTextA = LogTextA & " " & Chkbarcode(80).Caption & " ‰⁄„" & CHR(13)
  Else
  LogTextA = LogTextA & " " & Chkbarcode(80).Caption & " ·«" & CHR(13)
  End If
  
  If Chkbarcode(81).value = vbChecked Then
   LogTextA = LogTextA & " " & Chkbarcode(81).Caption & " ‰⁄„" & CHR(13)
  Else
  LogTextA = LogTextA & " " & Chkbarcode(81).Caption & " ·«" & CHR(13)
  End If
If Chkbarcode(82).value = vbChecked Then
   LogTextA = LogTextA & " " & Chkbarcode(82).Caption & " ‰⁄„" & CHR(13)
  Else
  LogTextA = LogTextA & " " & Chkbarcode(82).Caption & " ·«" & CHR(13)
  End If
  If Chkbarcode(83).value = vbChecked Then
   LogTextA = LogTextA & " " & Chkbarcode(83).Caption & " ‰⁄„" & CHR(13)
  Else
  LogTextA = LogTextA & " " & Chkbarcode(83).Caption & " ·«" & CHR(13)
  End If
 If Chkbarcode(84).value = vbChecked Then
   LogTextA = LogTextA & " " & Chkbarcode(84).Caption & " ‰⁄„" & CHR(13)
  Else
  LogTextA = LogTextA & " " & Chkbarcode(84).Caption & " ·«" & CHR(13)
  End If
  
  
  If Chkbarcode(100).value = vbChecked Then
   LogTextA = LogTextA & " " & Chkbarcode(100).Caption & " ‰⁄„" & CHR(13)
  Else
  LogTextA = LogTextA & " " & Chkbarcode(100).Caption & " ·«" & CHR(13)
  End If
  
 If Chkbarcode(85).value = vbChecked Then
   LogTextA = LogTextA & " " & Chkbarcode(85).Caption & " ‰⁄„" & CHR(13)
  Else
  LogTextA = LogTextA & " " & Chkbarcode(85).Caption & " ·«" & CHR(13)
  End If
  If Chkbarcode(86).value = vbChecked Then
   LogTextA = LogTextA & " " & Chkbarcode(86).Caption & " ‰⁄„" & CHR(13)
  Else
  LogTextA = LogTextA & " " & Chkbarcode(86).Caption & " ·«" & CHR(13)
  End If
  
      If Chkbarcode(42).value = vbChecked Then
   LogTextA = LogTextA & " " & Chkbarcode(42).Caption & " ‰⁄„" & CHR(13)
  Else
  LogTextA = LogTextA & " " & Chkbarcode(42).Caption & " ·«" & CHR(13)
  End If
  
     If Chkbarcode(43).value = vbChecked Then
   LogTextA = LogTextA & " " & Chkbarcode(43).Caption & " ‰⁄„" & CHR(13)
  Else
  LogTextA = LogTextA & " " & Chkbarcode(43).Caption & " ·«" & CHR(13)
  End If
  
  
     If Chkbarcode(44).value = vbChecked Then
   LogTextA = LogTextA & " " & Chkbarcode(44).Caption & " ‰⁄„" & CHR(13)
  Else
  LogTextA = LogTextA & " " & Chkbarcode(44).Caption & " ·«" & CHR(13)
  End If
If Chkbarcode(59).value = vbChecked Then
   LogTextA = LogTextA & " " & Chkbarcode(59).Caption & " ‰⁄„" & CHR(13)
  Else
  LogTextA = LogTextA & " " & Chkbarcode(59).Caption & " ·«" & CHR(13)
  End If
  
    If Chkbarcode(45).value = vbChecked Then
   LogTextA = LogTextA & " " & Chkbarcode(45).Caption & " ‰⁄„" & CHR(13)
  Else
  LogTextA = LogTextA & " " & Chkbarcode(45).Caption & " ·«" & CHR(13)
  End If
    If Chkbarcode(48).value = vbChecked Then
   LogTextA = LogTextA & " " & Chkbarcode(48).Caption & " ‰⁄„" & CHR(13)
  Else
  LogTextA = LogTextA & " " & Chkbarcode(48).Caption & " ·«" & CHR(13)
  End If
  
If Chkbarcode(49).value = vbChecked Then
   LogTextA = LogTextA & " " & Chkbarcode(49).Caption & " ‰⁄„" & CHR(13)
  Else
  LogTextA = LogTextA & " " & Chkbarcode(49).Caption & " ·«" & CHR(13)
  End If
   If Chkbarcode(51).value = vbChecked Then
   LogTextA = LogTextA & " " & Chkbarcode(51).Caption & " ‰⁄„" & CHR(13)
  Else
  LogTextA = LogTextA & " " & Chkbarcode(51).Caption & " ·«" & CHR(13)
  End If
 If Chkbarcode(52).value = vbChecked Then
   LogTextA = LogTextA & " " & Chkbarcode(52).Caption & " ‰⁄„" & CHR(13)
  Else
  LogTextA = LogTextA & " " & Chkbarcode(52).Caption & " ·«" & CHR(13)
  End If
  If Chkbarcode(53).value = vbChecked Then
   LogTextA = LogTextA & " " & Chkbarcode(53).Caption & " ‰⁄„" & CHR(13)
  Else
  LogTextA = LogTextA & " " & Chkbarcode(53).Caption & " ·«" & CHR(13)
  End If
If Chkbarcode(56).value = vbChecked Then
   LogTextA = LogTextA & " " & Chkbarcode(56).Caption & " ‰⁄„" & CHR(13)
  Else
  LogTextA = LogTextA & " " & Chkbarcode(56).Caption & " ·«" & CHR(13)
  End If
If Chkbarcode(54).value = vbChecked Then
   LogTextA = LogTextA & " " & Chkbarcode(54).Caption & " ‰⁄„" & CHR(13)
  Else
  LogTextA = LogTextA & " " & Chkbarcode(54).Caption & " ·«" & CHR(13)
  End If
  If Chkbarcode(55).value = vbChecked Then
   LogTextA = LogTextA & " " & Chkbarcode(55).Caption & " ‰⁄„" & CHR(13)
  Else
  LogTextA = LogTextA & " " & Chkbarcode(55).Caption & " ·«" & CHR(13)
  End If
    If Chkbarcode(57).value = vbChecked Then
   LogTextA = LogTextA & " " & Chkbarcode(57).Caption & " ‰⁄„" & CHR(13)
  Else
  LogTextA = LogTextA & " " & Chkbarcode(57).Caption & " ·«" & CHR(13)
  End If
      If Chkbarcode(58).value = vbChecked Then
   LogTextA = LogTextA & " " & Chkbarcode(58).Caption & " ‰⁄„" & CHR(13)
  Else
  LogTextA = LogTextA & " " & Chkbarcode(58).Caption & " ·«" & CHR(13)
  End If
  
  If Chkbarcode(132).value = vbChecked Then
   LogTextA = LogTextA & " " & Chkbarcode(132).Caption & " ‰⁄„" & CHR(13)
  Else
  LogTextA = LogTextA & " " & Chkbarcode(132).Caption & " ·«" & CHR(13)
  End If
  
If Chkbarcode(50).value = vbChecked Then
   LogTextA = LogTextA & " " & Chkbarcode(50).Caption & " ‰⁄„" & CHR(13)
  Else
  LogTextA = LogTextA & " " & Chkbarcode(50).Caption & " ·«" & CHR(13)
  End If
  
  LogTextA = LogTextA & " " & Frame35.Caption & CHR(13)
  
              If chkMonthIs30days.value = vbChecked Then
   LogTextA = LogTextA & " " & chkMonthIs30days.Caption & " ‰⁄„" & CHR(13)
  Else
  LogTextA = LogTextA & " " & chkMonthIs30days.Caption & " ·«" & CHR(13)
  End If
  
  
   If Chkemployeeaccounts.value = vbChecked Then
   LogTextA = LogTextA & " " & Chkemployeeaccounts.Caption & " ‰⁄„" & CHR(13)
  Else
  LogTextA = LogTextA & " " & Chkemployeeaccounts.Caption & " ·«" & CHR(13)
  End If
        
  If Chkbarcode(24).value = vbChecked Then
   LogTextA = LogTextA & " " & Chkbarcode(24).Caption & " ‰⁄„" & CHR(13)
  Else
  LogTextA = LogTextA & " " & Chkbarcode(24).Caption & " ·«" & CHR(13)
  End If
  
  LogTextA = LogTextA & " " & lbl(38).Caption & " " & TxtEmpComponentDigts.text & CHR(13)
  LogTextA = LogTextA & " " & Label13.Caption & " " & TxtEmpSalaryDigts.text & CHR(13)
  LogTextA = LogTextA & " " & Frame48.Caption & CHR(13)
  
     If ChkEmpRes(0).value = True Then
    LogTextA = LogTextA & " " & ChkEmpRes(0).value & CHR(13)
  ElseIf ChkEmpRes(1).value = True Then
    LogTextA = LogTextA & " " & ChkEmpRes(1).value & CHR(13)
 
   End If
   
     If Chkbarcode(6).value = vbChecked Then
   LogTextA = LogTextA & " " & Chkbarcode(6).Caption & " ‰⁄„" & CHR(13)
  Else
  LogTextA = LogTextA & " " & Chkbarcode(6).Caption & " ·«" & CHR(13)
  End If
  
  
 If Chkbarcode(60).value = vbChecked Then
   LogTextA = LogTextA & " " & Chkbarcode(60).Caption & " ‰⁄„" & CHR(13)
  Else
  LogTextA = LogTextA & " " & Chkbarcode(60).Caption & " ·«" & CHR(13)
  End If
 If Chkbarcode(61).value = vbChecked Then
   LogTextA = LogTextA & " " & Chkbarcode(61).Caption & " ‰⁄„" & CHR(13)
  Else
  LogTextA = LogTextA & " " & Chkbarcode(61).Caption & " ·«" & CHR(13)
  End If
  If Chkbarcode(62).value = vbChecked Then
   LogTextA = LogTextA & " " & Chkbarcode(62).Caption & " ‰⁄„" & CHR(13)
  Else
  LogTextA = LogTextA & " " & Chkbarcode(62).Caption & " ·«" & CHR(13)
  End If
  
    If Chkbarcode(75).value = vbChecked Then
   LogTextA = LogTextA & " " & Chkbarcode(75).Caption & " ‰⁄„" & CHR(13)
  Else
  LogTextA = LogTextA & " " & Chkbarcode(75).Caption & " ·«" & CHR(13)
  End If
  
      If Chkbarcode(76).value = vbChecked Then
   LogTextA = LogTextA & " " & Chkbarcode(76).Caption & " ‰⁄„" & CHR(13)
  Else
  LogTextA = LogTextA & " " & Chkbarcode(76).Caption & " ·«" & CHR(13)
  End If
  
  
  If Chkbarcode(63).value = vbChecked Then
   LogTextA = LogTextA & " " & Chkbarcode(63).Caption & " ‰⁄„" & CHR(13)
  Else
  LogTextA = LogTextA & " " & Chkbarcode(63).Caption & " ·«" & CHR(13)
  End If
  If Chkbarcode(64).value = vbChecked Then
   LogTextA = LogTextA & " " & Chkbarcode(64).Caption & " ‰⁄„" & CHR(13)
  Else
  LogTextA = LogTextA & " " & Chkbarcode(64).Caption & " ·«" & CHR(13)
  End If
  If Chkbarcode(65).value = vbChecked Then
   LogTextA = LogTextA & " " & Chkbarcode(65).Caption & " ‰⁄„" & CHR(13)
  Else
  LogTextA = LogTextA & " " & Chkbarcode(65).Caption & " ·«" & CHR(13)
  End If
    If Chkbarcode(66).value = vbChecked Then
   LogTextA = LogTextA & " " & Chkbarcode(66).Caption & " ‰⁄„" & CHR(13)
  Else
  LogTextA = LogTextA & " " & Chkbarcode(66).Caption & " ·«" & CHR(13)
  End If
If Chkbarcode(67).value = vbChecked Then
   LogTextA = LogTextA & " " & Chkbarcode(67).Caption & " ‰⁄„" & CHR(13)
  Else
  LogTextA = LogTextA & " " & Chkbarcode(67).Caption & " ·«" & CHR(13)
  End If
If Chkbarcode(68).value = vbChecked Then
   LogTextA = LogTextA & " " & Chkbarcode(68).Caption & " ‰⁄„" & CHR(13)
  Else
  LogTextA = LogTextA & " " & Chkbarcode(68).Caption & " ·«" & CHR(13)
  End If
  If Chkbarcode(69).value = vbChecked Then
   LogTextA = LogTextA & " " & Chkbarcode(69).Caption & " ‰⁄„" & CHR(13)
  Else
  LogTextA = LogTextA & " " & Chkbarcode(69).Caption & " ·«" & CHR(13)
  End If
 If Chkbarcode(70).value = vbChecked Then
   LogTextA = LogTextA & " " & Chkbarcode(70).Caption & " ‰⁄„" & CHR(13)
  Else
  LogTextA = LogTextA & " " & Chkbarcode(70).Caption & " ·«" & CHR(13)
  End If
   If Chkbarcode(71).value = vbChecked Then
   LogTextA = LogTextA & " " & Chkbarcode(71).Caption & " ‰⁄„" & CHR(13)
  Else
  LogTextA = LogTextA & " " & Chkbarcode(71).Caption & " ·«" & CHR(13)
  End If
     If Chkbarcode(72).value = vbChecked Then
   LogTextA = LogTextA & " " & Chkbarcode(72).Caption & " ‰⁄„" & CHR(13)
  Else
  LogTextA = LogTextA & " " & Chkbarcode(72).Caption & " ·«" & CHR(13)
  End If
  If Chkbarcode(73).value = vbChecked Then
   LogTextA = LogTextA & " " & Chkbarcode(73).Caption & " ‰⁄„" & CHR(13)
  Else
  LogTextA = LogTextA & " " & Chkbarcode(73).Caption & " ·«" & CHR(13)
  End If
  If Chkbarcode(74).value = vbChecked Then
   LogTextA = LogTextA & " " & Chkbarcode(74).Caption & " ‰⁄„" & CHR(13)
  Else
  LogTextA = LogTextA & " " & Chkbarcode(74).Caption & " ·«" & CHR(13)
  End If
  If Chkbarcode(36).value = vbChecked Then
   LogTextA = LogTextA & " " & Chkbarcode(36).Caption & " ‰⁄„" & CHR(13)
  Else
  LogTextA = LogTextA & " " & Chkbarcode(36).Caption & " ·«" & CHR(13)
  End If
  
    
      If Chkbarcode(37).value = vbChecked Then
   LogTextA = LogTextA & " " & Chkbarcode(37).Caption & " ‰⁄„" & CHR(13)
  Else
  LogTextA = LogTextA & " " & Chkbarcode(37).Caption & " ·«" & CHR(13)
  End If
  
  
         If ChkTypicalProduction.value = vbChecked Then
   LogTextA = LogTextA & " " & ChkTypicalProduction.Caption & " ‰⁄„" & CHR(13)
  Else
  LogTextA = LogTextA & " " & ChkTypicalProduction.Caption & " ·«" & CHR(13)
  End If
  
    
        
         If ChkAllowIndirectCost.value = vbChecked Then
   LogTextA = LogTextA & " " & ChkAllowIndirectCost.Caption & " ‰⁄„" & CHR(13)
    LogTextA = LogTextA & " " & TxtIndirectCostPercentage.text & CHR(13)
    
  Else
  LogTextA = LogTextA & " " & ChkAllowIndirectCost.Caption & " ·«" & CHR(13)
  End If
  
'            If chkEmpProduction.value = vbChecked Then
'   LogTextA = LogTextA & " " & chkEmpProduction.Caption & " ‰⁄„" & CHR(13)
'    LogTextA = LogTextA & " " & TxtIndirectCostPercentage.Text & CHR(13)
'
'  Else
'  LogTextA = LogTextA & " " & chkEmpProduction.Caption & " ·«" & CHR(13)
'  End If
'
'
'              If chkExpProduction.value = vbChecked Then
'   LogTextA = LogTextA & " " & chkExpProduction.Caption & " ‰⁄„" & CHR(13)
'    LogTextA = LogTextA & " " & TxtIndirectCostPercentage.Text & CHR(13)
'
'  Else
'  LogTextA = LogTextA & " " & chkExpProduction.Caption & " ·«" & CHR(13)
'  End If
'
'              If chkItemProduction.value = vbChecked Then
'   LogTextA = LogTextA & " " & chkItemProduction.Caption & " ‰⁄„" & CHR(13)
'    LogTextA = LogTextA & " " & TxtIndirectCostPercentage.Text & CHR(13)
'
'  Else
'  LogTextA = LogTextA & " " & chkItemProduction.Caption & " ·«" & CHR(13)
'  End If
   
    
   
   
   
           If ChkDriverBox.value = vbChecked Then
   LogTextA = LogTextA & " " & ChkDriverBox.Caption & " ‰⁄„" & CHR(13)
  Else
  LogTextA = LogTextA & " " & ChkDriverBox.Caption & " ·«" & CHR(13)
  End If
  
  
           If chkDriverEra.value = vbChecked Then
   LogTextA = LogTextA & " " & chkDriverEra.Caption & " ‰⁄„" & CHR(13)
  Else
  LogTextA = LogTextA & " " & chkDriverEra.Caption & " ·«" & CHR(13)
  End If
    
    
           If Chkbarcode(9).value = vbChecked Then
   LogTextA = LogTextA & " " & Chkbarcode(9).Caption & " ‰⁄„" & CHR(13)
  Else
  LogTextA = LogTextA & " " & Chkbarcode(9).Caption & " ·«" & CHR(13)
  End If
        
           If Chkbarcode(11).value = vbChecked Then
   LogTextA = LogTextA & " " & Chkbarcode(11).Caption & " ‰⁄„" & CHR(13)
  Else
  LogTextA = LogTextA & " " & Chkbarcode(11).Caption & " ·«" & CHR(13)
  End If
                
                
                If ChkAssetAccount.value = vbChecked Then
   LogTextA = LogTextA & " " & ChkAssetAccount.Caption & " ‰⁄„" & CHR(13)
  Else
  LogTextA = LogTextA & " " & ChkAssetAccount.Caption & " ·«" & CHR(13)
  End If
               
               
                If ChkAssetAccount1.value = vbChecked Then
   LogTextA = LogTextA & " " & ChkAssetAccount1.Caption & " ‰⁄„" & CHR(13)
  Else
  LogTextA = LogTextA & " " & ChkAssetAccount1.Caption & " ·«" & CHR(13)
  End If
                              
 
  
  LogTextA = LogTextA & " " & Frame18.Caption & CHR(13)
                              
 'OptionItemsTotal
     If OptionItemsTotal.value = True Then
    LogTextA = LogTextA & " " & OptionItemsTotal.value & CHR(13)
  ElseIf OptionOperation.value = True Then
    LogTextA = LogTextA & " " & OptionOperation.value & CHR(13)
 
   End If
   

   
        If OPTdISCOUNT(0).value = True Then
    LogTextA = LogTextA & " " & OPTdISCOUNT(0).value & CHR(13)
  ElseIf OPTdISCOUNT(1).value = True Then
    LogTextA = LogTextA & " " & OPTdISCOUNT(1).value & CHR(13)
 
   End If
   
   
   
     
    
   
  LogTextA = LogTextA & " " & Frame19.Caption & CHR(13)
   
     If GlDetails.value = True Then
    LogTextA = LogTextA & " " & GlDetails.value & CHR(13)
  ElseIf glgeneral.value = True Then
    LogTextA = LogTextA & " " & glgeneral.value & CHR(13)
 
   End If
   
   
     LogTextA = LogTextA & " " & Frame21.Caption & CHR(13)
                              
 'OptionItemsTotal
     If Optday.value = True Then
    LogTextA = LogTextA & " " & Optday.value & CHR(13)
  ElseIf Optweek.value = True Then
    LogTextA = LogTextA & " " & Optweek.value & CHR(13)
 
  ElseIf OptMonth.value = True Then
    LogTextA = LogTextA & " " & OptMonth.value & CHR(13)
    
   ElseIf OptYear.value = True Then
    LogTextA = LogTextA & " " & OptYear.value & CHR(13)
    
    
   End If
   
   
  
                  If Chkbarcode(5).value = vbChecked Then
   LogTextA = LogTextA & " " & Chkbarcode(5).Caption & " ‰⁄„" & CHR(13)
  Else
  LogTextA = LogTextA & " " & Chkbarcode(5).Caption & " ·«" & CHR(13)
  End If
  
                  If Chkbarcode(19).value = vbChecked Then
   LogTextA = LogTextA & " " & Chkbarcode(19).Caption & " ‰⁄„" & CHR(13)
  Else
  LogTextA = LogTextA & " " & Chkbarcode(19).Caption & " ·«" & CHR(13)
  End If
  
  
                    If Chkbarcode(20).value = vbChecked Then
   LogTextA = LogTextA & " " & Chkbarcode(20).Caption & " ‰⁄„" & CHR(13)
  Else
  LogTextA = LogTextA & " " & Chkbarcode(20).Caption & " ·«" & CHR(13)
  End If
  
  
       LogTextA = LogTextA & " " & Frame33.Caption & CHR(13)
                              
 'OptionItemsTotal
     If OptArrowBranch.value = True Then
    LogTextA = LogTextA & " " & OptArrowBranch.value & CHR(13)
  ElseIf OptArrowGroup.value = True Then
    LogTextA = LogTextA & " " & OptArrowGroup.value & CHR(13)
  
    
   End If
   
   
       LogTextA = LogTextA & " " & Frame23.Caption & CHR(13)
                              
 'OptionItemsTotal
     If Option1.value = True Then
    LogTextA = LogTextA & " " & Option1.value & CHR(13)
  ElseIf Option2.value = True Then
    LogTextA = LogTextA & " " & Option2.value & CHR(13)
    End If
      
      LogTextA = LogTextA & " " & lbl(13).Caption & " " & TxtPriceDigts.text & CHR(13)
      LogTextA = LogTextA & " " & lbl(14).Caption & " " & TxtQtyDigts.text & CHR(13)
      LogTextA = LogTextA & " " & lbl(17).Caption & " " & TxtPriceDigtsInst.text & CHR(13)
      LogTextA = LogTextA & " " & lbl(16).Caption & " " & txt_ACCOUNT_digit.text & CHR(13)
      
        If Chkbarcode(23).value = vbChecked Then
   LogTextA = LogTextA & " " & Chkbarcode(23).Caption & " ‰⁄„" & CHR(13)
  Else
  LogTextA = LogTextA & " " & Chkbarcode(23).Caption & " ·«" & CHR(13)
  End If
  
  
        If ChkDelayVal.value = vbChecked Then
   LogTextA = LogTextA & " " & ChkDelayVal.Caption & " ‰⁄„" & CHR(13)
   LogTextA = LogTextA & " " & Text1.text & " " & Combo1.text & CHR(13)
   
  Else
  LogTextA = LogTextA & " " & ChkDelayVal.Caption & " ·«" & CHR(13)
  End If
    
  
  
  If ChkInstallmentMustPayed.value = vbChecked Then
   LogTextA = LogTextA & " " & ChkInstallmentMustPayed.Caption & " ‰⁄„" & CHR(13)
   LogTextA = LogTextA & " " & Text2.text & " " & Combo2.text & CHR(13)
   
  Else
  LogTextA = LogTextA & " " & ChkInstallmentMustPayed.Caption & " ·«" & CHR(13)
  End If
  
    
  If ChkExpireEkama.value = vbChecked Then
   LogTextA = LogTextA & " " & ChkExpireEkama.Caption & " ‰⁄„" & CHR(13)
   LogTextA = LogTextA & " " & Text3.text & " " & Combo3.text & CHR(13)
   
  Else
  LogTextA = LogTextA & " " & ChkExpireEkama.Caption & " ·«" & CHR(13)
  End If
      
      
     If ChkExpireLicence.value = vbChecked Then
   LogTextA = LogTextA & " " & ChkExpireLicence.Caption & " ‰⁄„" & CHR(13)
   LogTextA = LogTextA & " " & Text4.text & " " & Combo4.text & CHR(13)
   
  Else
  LogTextA = LogTextA & " " & ChkExpireLicence.Caption & " ·«" & CHR(13)
  End If
  
       If ChkExpirepas.value = vbChecked Then
   LogTextA = LogTextA & " " & ChkExpirepas.Caption & " ‰⁄„" & CHR(13)
   LogTextA = LogTextA & " " & Text5.text & " " & Combo5.text & CHR(13)
   
  Else
  LogTextA = LogTextA & " " & ChkExpirepas.Caption & " ·«" & CHR(13)
  End If
  
       If ChkExpirepoket.value = vbChecked Then
   LogTextA = LogTextA & " " & ChkExpirepoket.Caption & " ‰⁄„" & CHR(13)
   LogTextA = LogTextA & " " & Text6.text & " " & Combo6.text & CHR(13)
   
  Else
  LogTextA = LogTextA & " " & ChkExpirepoket.Caption & " ·«" & CHR(13)
  End If
  
  
       If chkRentInstallments.value = vbChecked Then
   LogTextA = LogTextA & " " & chkRentInstallments.Caption & " ‰⁄„" & CHR(13)
   LogTextA = LogTextA & " " & Text8.text & " " & Combo10.text & CHR(13)
   
  Else
  LogTextA = LogTextA & " " & chkRentInstallments.Caption & " ·«" & CHR(13)
  End If
  
      If Check6.value = vbChecked Then
   LogTextA = LogTextA & " " & Check6.Caption & " ‰⁄„" & CHR(13)
   LogTextA = LogTextA & " " & Text9.text & " " & Combo11.text & CHR(13)
   
  Else
  LogTextA = LogTextA & " " & Check6.Caption & " ·«" & CHR(13)
  End If
    
      
      If Check6.value = vbChecked Then
   LogTextA = LogTextA & " " & Check6.Caption & " ‰⁄„" & CHR(13)
   LogTextA = LogTextA & " " & Text9.text & " " & Combo11.text & CHR(13)
   
  Else
  LogTextA = LogTextA & " " & Check6.Caption & " ·«" & CHR(13)
  End If
  
  
          If Check7.value = vbChecked Then
   LogTextA = LogTextA & " " & Check7.Caption & " ‰⁄„" & CHR(13)
   LogTextA = LogTextA & " " & Text10.text & " " & Combo12.text & CHR(13)
   
  Else
  LogTextA = LogTextA & " " & Check7.Caption & " ·«" & CHR(13)
  End If
  
  
          If ChkExpireLicense.value = vbChecked Then
   LogTextA = LogTextA & " " & ChkExpireLicense.Caption & " ‰⁄„" & CHR(13)
   LogTextA = LogTextA & " " & Text11.text & " " & Combo7.text & CHR(13)
   
  Else
  LogTextA = LogTextA & " " & ChkExpireLicense.Caption & " ·«" & CHR(13)
  End If
    
    
    
            If ChkExpireInsurance.value = vbChecked Then
   LogTextA = LogTextA & " " & ChkExpireInsurance.Caption & " ‰⁄„" & CHR(13)
   LogTextA = LogTextA & " " & Text12.text & " " & Combo8.text & CHR(13)
   
  Else
  LogTextA = LogTextA & " " & ChkExpireInsurance.Caption & " ·«" & CHR(13)
  End If
  
            If ChkExpireTest.value = vbChecked Then
   LogTextA = LogTextA & " " & ChkExpireTest.Caption & " ‰⁄„" & CHR(13)
   LogTextA = LogTextA & " " & Text13.text & " " & Combo9.text & CHR(13)
   
  Else
  LogTextA = LogTextA & " " & ChkExpireTest.Caption & " ·«" & CHR(13)
  End If
  
      
               If Check1.value = vbChecked Then
   LogTextA = LogTextA & " " & Check1.Caption & " ‰⁄„" & CHR(13)
   LogTextA = LogTextA & " " & Text14.text & " " & lbl(24).Caption & CHR(13)
   
  Else
  LogTextA = LogTextA & " " & Check1.Caption & " ·«" & CHR(13)
  End If
 
 
             If ChkHideAllAlarms.value = vbChecked Then
   LogTextA = LogTextA & " " & ChkHideAllAlarms.Caption & " ‰⁄„" & CHR(13)
    
  Else
  LogTextA = LogTextA & " " & ChkHideAllAlarms.Caption & " ·«" & CHR(13)
  End If
  
             If ChKProjectsAlarm1.value = vbChecked Then
   LogTextA = LogTextA & " " & ChKProjectsAlarm1.Caption & " ‰⁄„" & CHR(13)
    
  Else
  LogTextA = LogTextA & " " & ChKProjectsAlarm1.Caption & " ·«" & CHR(13)
  End If
  
             If ChKProjectsAlarm2.value = vbChecked Then
   LogTextA = LogTextA & " " & ChKProjectsAlarm2.Caption & " ‰⁄„" & CHR(13)
    
  Else
  LogTextA = LogTextA & " " & ChKProjectsAlarm2.Caption & " ·«" & CHR(13)
  End If
  
             If ChkShow.value = vbChecked Then
   LogTextA = LogTextA & " " & ChkShow.Caption & " ‰⁄„" & CHR(13)
    
  Else
  LogTextA = LogTextA & " " & ChkShow.Caption & " ·«" & CHR(13)
  End If
  
             If CheckLC.value = vbChecked Then
   LogTextA = LogTextA & " " & CheckLC.Caption & " ‰⁄„" & CHR(13)
   LogTextA = LogTextA & " " & Text7.text & " " & Combo14.text & CHR(13)
   
    
  Else
  LogTextA = LogTextA & " " & CheckLC.Caption & " ·«" & CHR(13)
  End If
  
  
              If CHECK_OPEN_NEW_SCREEN.value = vbChecked Then
   LogTextA = LogTextA & " " & CHECK_OPEN_NEW_SCREEN.Caption & " ‰⁄„" & CHR(13)
    
  Else
  LogTextA = LogTextA & " " & CHECK_OPEN_NEW_SCREEN.Caption & " ·«" & CHR(13)
  End If
  

     LogTextA = LogTextA & " " & Frame59.Caption & CHR(13)
                              
 'OptionItemsTotal
     If Option4(0).value = True Then
    LogTextA = LogTextA & " " & Option4(0).value & CHR(13)
  ElseIf Option4(1).value = True Then
    LogTextA = LogTextA & " " & Option4(1).value & CHR(13)
      ElseIf Option4(2).value = True Then
    LogTextA = LogTextA & " " & Option4(2).value & CHR(13)
      ElseIf Option4(3).value = True Then
    LogTextA = LogTextA & " " & Option4(3).value & CHR(13)
    
    End If
    
    
    If chkshortCuts.value = vbChecked Then
   LogTextA = LogTextA & " " & chkshortCuts.Caption & " ‰⁄„" & CHR(13)
    
  Else
  LogTextA = LogTextA & " " & chkshortCuts.Caption & " ·«" & CHR(13)
  End If
  
    If chkuserCode.value = vbChecked Then
   LogTextA = LogTextA & " " & chkuserCode.Caption & " ‰⁄„" & CHR(13)
    
  Else
  LogTextA = LogTextA & " " & chkuserCode.Caption & " ·«" & CHR(13)
  End If
  
  
    If Chktree.value = vbChecked Then
   LogTextA = LogTextA & " " & Chktree.Caption & " ‰⁄„" & CHR(13)
    
  Else
  LogTextA = LogTextA & " " & Chktree.Caption & " ·«" & CHR(13)
  End If
    
    If ChkCalender.value = vbChecked Then
   LogTextA = LogTextA & " " & ChkCalender.Caption & " ‰⁄„" & CHR(13)
    
  Else
  LogTextA = LogTextA & " " & ChkCalender.Caption & " ·«" & CHR(13)
  End If
  
    
    If Chkgraphic.value = vbChecked Then
   LogTextA = LogTextA & " " & Chkgraphic.Caption & " ‰⁄„" & CHR(13)
    
  Else
  LogTextA = LogTextA & " " & Chkgraphic.Caption & " ·«" & CHR(13)
  End If
      
      
    If ChkMessnger.value = vbChecked Then
   LogTextA = LogTextA & " " & ChkMessnger.Caption & " ‰⁄„" & CHR(13)
    
  Else
  LogTextA = LogTextA & " " & ChkMessnger.Caption & " ·«" & CHR(13)
  End If
  
      
    If ChkViewAging.value = vbChecked Then
   LogTextA = LogTextA & " " & ChkViewAging.Caption & " ‰⁄„" & CHR(13)
    
  Else
  LogTextA = LogTextA & " " & ChkViewAging.Caption & " ·«" & CHR(13)
  End If
        
    If ChkShowToolTip.value = vbChecked Then
   LogTextA = LogTextA & " " & ChkShowToolTip.Caption & " ‰⁄„" & CHR(13)
    
  Else
  LogTextA = LogTextA & " " & ChkShowToolTip.Caption & " ·«" & CHR(13)
  End If
  
    If ChkAsk.value = vbChecked Then
   LogTextA = LogTextA & " " & ChkAsk.Caption & " ‰⁄„" & CHR(13)
    
  Else
  LogTextA = LogTextA & " " & ChkAsk.Caption & " ·«" & CHR(13)
  End If
  
      If ChkPrintBranchINGE.value = vbChecked Then
   LogTextA = LogTextA & " " & ChkPrintBranchINGE.Caption & " ‰⁄„" & CHR(13)
    
  Else
  LogTextA = LogTextA & " " & ChkPrintBranchINGE.Caption & " ·«" & CHR(13)
  End If
  
      If ChkPrintCCinGE.value = vbChecked Then
   LogTextA = LogTextA & " " & ChkPrintCCinGE.Caption & " ‰⁄„" & CHR(13)
    
  Else
  LogTextA = LogTextA & " " & ChkPrintCCinGE.Caption & " ·«" & CHR(13)
  End If
          
     If ChkChartPrintinAS.value = vbChecked Then
   LogTextA = LogTextA & " " & ChkChartPrintinAS.Caption & " ‰⁄„" & CHR(13)
    
  Else
  LogTextA = LogTextA & " " & ChkChartPrintinAS.Caption & " ·«" & CHR(13)
  End If
                    
                    
                    
         LogTextA = LogTextA & " " & Fra(4).Caption & CHR(13)
                              
     If DateOpt(0).value = True Then
    LogTextA = LogTextA & " " & DateOpt(0).value & CHR(13)
  ElseIf DateOpt(1).value = True Then
    LogTextA = LogTextA & " " & DateOpt(1).value & CHR(13)
 
    
    End If
    
    
  LogTextA = LogTextA & " " & lbl(22).Caption & " " & TxtImagesPath(0).text & CHR(13)
  LogTextA = LogTextA & " " & lbl(26).Caption & " " & TxtImagesPath(1).text & CHR(13)
  LogTextA = LogTextA & " " & lbl(30).Caption & " " & TxtImagesPath(2).text & CHR(13)
  
  
      
  LogTextA = LogTextA & " " & Label18.Caption & " " & TxtZoom.text & CHR(13)
  LogTextA = LogTextA & " " & Label21.Caption & " " & TxtLogoWidth.text & CHR(13)
  LogTextA = LogTextA & " " & Label22.Caption & " " & TxtLogoheight.text & CHR(13)
          
FollowCurrentLogData
    
End Function
Sub FollowCurrentLogData()
             If Chkbarcode(8).value = vbChecked Then
   LogTextA = LogTextA & " " & Chkbarcode(8).Caption & " ‰⁄„" & CHR(13)
    
  Else
  LogTextA = LogTextA & " " & Chkbarcode(8).Caption & " ·«" & CHR(13)
  End If
  
  
             If Chkbarcode(14).value = vbChecked Then
   LogTextA = LogTextA & " " & Chkbarcode(14).Caption & " ‰⁄„" & CHR(13)
    
  Else
  LogTextA = LogTextA & " " & Chkbarcode(14).Caption & " ·«" & CHR(13)
  End If
  
  LogTextA = LogTextA & " " & Label23.Caption & " " & TxtData(0).text & CHR(13)
  LogTextA = LogTextA & " " & Label24.Caption & " " & TxtData(1).text & CHR(13)
  
 LogTextA = LogTextA & " " & Label20.Caption & " " & TXTReturnSallingIntervalCount(1).text & CHR(13)
 LogTextA = LogTextA & " " & Label8.Caption & " " & TXTReturnSallingIntervalCount(2).text & CHR(13)
       
                 If Chkbarcode(10).value = vbChecked Then
   LogTextA = LogTextA & " " & Chkbarcode(10).Caption & " ‰⁄„" & CHR(13)
    
  Else
  LogTextA = LogTextA & " " & Chkbarcode(10).Caption & " ·«" & CHR(13)
  End If
  LogTexte = LogTextA
 
    If Currentmode <> "D" Then
        AddToLogFile CInt(user_id), 0, Date, Time, LogTextA, LogTexte, Me.Name, "E", "", , 0, ""
    Else
        AddToLogFile CInt(user_id), 0, Date, Time, LogTextA, LogTexte, Me.Name, "D", "", 0, "", ""
    End If
End Sub

Private Sub checkss()
   If IsNumeric(TxtZoom.text) Then
                SaveSetting StrAppRegPath, "Setting", "ReportZoom", val(TxtZoom.text)
            Else
                SaveSetting StrAppRegPath, "Setting", "ReportZoom", 100
            End If
            
            
            
            
    If Me.chk.value = vbChecked Then
        If Me.ImgPic.Picture = 0 Then
            Msg = "ÌÃ» ≈Œ Ì«— ’Ê—… ‘⁄«— «·‘—ﬂ… ·⁄—÷Â« ›Ï «· ﬁ«—Ì—..!!"
            MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
          NewTab = 0
            Exit Sub
        End If
    End If

    If IsNumeric(Me.TxtPriceDigts.text) = False Then
        Msg = "«·ﬁÌ„… «·„œŒ·… ··⁄·«„«    «·⁄‘—Ì… ··⁄„·… €Ì— ’ÕÌÕ…...!!!"
        MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
        Exit Sub
    End If

    If IsNumeric(Me.TxtPriceDigtsInst.text) = False Then
        Msg = "«·ﬁÌ„… «·„œŒ·… ··⁄·«„«    «·⁄‘—Ì… ·· ﬁ”Ìÿ €Ì— ’ÕÌÕ…...!!!"
        MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
        Exit Sub
    End If

    If IsNumeric(Me.TxtEmpComponentDigts.text) = False Then
        Msg = "«·ﬁÌ„… «·„œŒ·… ··⁄œœ   «·⁄‘—Ì… ··„›—œ«  «·„ €Ì—… ··„ÊŸ›Ì‰ €Ì— ’ÕÌÕ…...!!!"
        MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
        Exit Sub
    End If

'    If ChkAllowIndirectCost.value = vbChecked Then
'        If IsNumeric(Me.TxtIndirectCostPercentage.text) = False Or val(TxtIndirectCostPercentage.text) = 0 Then
'            Msg = "«·ﬁÌ„… «·„œŒ·… ·‰”»… «· ﬂ«·Ì› €Ì— «·„»«‘—…       €Ì— ’ÕÌÕ…...!!!"
'            MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
'            Exit Sub
'        End If
'
'    Else
'        TxtIndirectCostPercentage = 0
'
'    End If

    If IsNumeric(Me.txtLimitDefaultCredit.text) = False Then
        Msg = "«·ﬁÌ„… «·„œŒ·…   ··Õœ «·«∆ „«‰Ï «·«› —«÷Ï €Ì— ’ÕÌÕ…...!!!"
        MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
        Exit Sub
    End If


    If IsNumeric(Me.txtLimitDefaultCreditDays.text) = False Then
        Msg = "«·ﬁÌ„… «·„œŒ·…  ··„œ… «·«∆ „«‰Ì… «·«› —«÷Ì… €Ì— ’ÕÌÕ…...!!!"
        MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
        Exit Sub
    End If

    If IsNumeric(Me.TxtSaleDiscount1.text) = False Then
        Msg = "«·ﬁÌ„… «·„œŒ·…        ·Œ’„ «·’‰›  €Ì— ’ÕÌÕ…...!!!"
        MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
        Exit Sub
    End If


    If IsNumeric(Me.TxtSaleDiscount2.text) = False Then
        Msg = "«·ﬁÌ„… «·„œŒ·…        ·Œ’„ „Ã„Ê⁄Â «·’‰›  €Ì— ’ÕÌÕ…...!!!"
        MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
        Exit Sub
    End If

    If IsNumeric(Me.TxtSaleDiscount3.text) = False Then
        Msg = "«·ﬁÌ„… «·„œŒ·…        ·Œ’„ «·⁄„Ì·  €Ì— ’ÕÌÕ…...!!!"
        MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
        Exit Sub
    End If

    If IsNumeric(Me.TxtSaleDiscount4.text) = False Then
        Msg = "«·ﬁÌ„… «·„œŒ·…        ·Œ’„ «·„‰œÊ»  €Ì— ’ÕÌÕ…...!!!"
        MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
        Exit Sub
    End If





    If val(Me.TxtLogoWidth.text) < 1000 Then
        Msg = "«·ﬁÌ„… «·„œŒ·…  ·⁄—÷ «··ÊÃÊ ·«»œ «‰  ﬂÊ‰ «ﬂ»— „‰ «Ê Ì”«ÊÌ 1000  ...!!!"
        MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
        Exit Sub
    End If
    
     If val(Me.TxtLogoheight.text) < 1000 Then
        Msg = "«·ﬁÌ„… «·„œŒ·…  ·«— ›«⁄  «··ÊÃÊ ·«»œ «‰  ﬂÊ‰ «ﬂ»— „‰ «Ê Ì”«ÊÌ 1000  ...!!!"
        MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
        Exit Sub
    End If
       
End Sub

Private Sub CmdOk_Click()

If checkEeinvoice = False Then Exit Sub


checksave = True
checkss
    Dim Msg  As String
    '  On Error GoTo ErrTrap

        
       
        rs("LogoWidth").value = val(Me.TxtLogoWidth.text)
        
            rs("Logoheight").value = val(Me.TxtLogoheight.text)
            
    rs("Company_Arabic_Name").value = XPTxtCompany.text
    rs("Company_Name_Eng").value = XPTxtCompanye.text
    rs("ImagesPath").value = IIf(TxtImagesPath(0).text = "", "Images", TxtImagesPath(0).text)
    rs("reportPath").value = IIf(TxtImagesPath(1).text = "", "Stander", TxtImagesPath(1).text)
    rs("BigUserPw").value = IIf(TxtImagesPath(2).text = "", "n20172018", TxtImagesPath(2).text)
    rs("BigUserPw2").value = IIf(TxtImagesPath(3).text = "", "123456", TxtImagesPath(3).text)
    
    
   
   rs("CountPrint").value = val(TXTReturnSallingIntervalCount(6).text)
    rs("NoRoudProjectInvoices").value = val(TXTReturnSallingIntervalCount(3).text)
    rs("VATItems").value = val(TXTReturnSallingIntervalCount(4).text)
    rs("NoBooking").value = val(TXTReturnSallingIntervalCount(2).text)
    rs("itemSeprator").value = TXTReturnSallingIntervalCount(1).text
    
    rs("DefaultQtyTrans").value = val(TXTReturnSallingIntervalCount(5).text)
    
    
    rs("Fax").value = TxtFax.text
    rs("Company_Comment").value = XPTxtComment(0).text
    rs("MembershipNo").value = XPTxtComment(1).text
    rs("ComputerNo").value = XPTxtComment(2).text
    rs("VATRegNo").value = XPTxtComment(3).text
    rs("WEBSite").value = TxtEmails.text
    
        rs("IsSerialByUserTrans").value = val(txtNoOFDigitUser(0))
        rs("NoOFDigitUserVouc").value = val(txtNoOFDigitUser(1))

 
    
    rs("Company_Address").value = XPTxtAddress.text
    rs("Company_Phone").value = xptxtphone.text
    rs("Company_Mobile").value = XPTxtmobile.text
    rs("Company_Maile").value = XPTxtMail.text
    rs("DomainData").value = txtDomainData.text
    
    rs("Company_Responsable").value = XPTxtResponsable.text
    rs("SalesBoxID").value = val(Me.DcboBox.BoundText)

    If opt(0).value = True Then
        rs("InvDate").value = 0
    ElseIf opt(1).value = True Then
        rs("InvDate").value = 1
    ElseIf opt(2).value = True Then
        rs("InvDate").value = 2
    End If



    If opt(5).value = True Then
        rs("PurDate").value = 0
    ElseIf opt(4).value = True Then
        rs("PurDate").value = 1
    ElseIf opt(3).value = True Then
        rs("PurDate").value = 2
    End If
    
    
    If opt(8).value = True Then
        rs("CashDate").value = 0
    ElseIf opt(9).value = True Then
        rs("CashDate").value = 1
    ElseIf opt(10).value = True Then
        rs("CashDate").value = 2
    End If
    
    
    If DateOpt(0).value = True Then
        rs("DateOpt").value = 0
    ElseIf DateOpt(1).value = True Then
        rs("DateOpt").value = 1
    End If
rs("NOOFPRINTCOPIESSALES").value = val(MYTEXT(0).text)
    If opt(6).value = True Then
        rs("ReturnSallingOption").value = 0
         
        rs("ReturnSallingIntervalCount").value = val(TXTReturnSallingIntervalCount(0).text)
        rs("ReturnSallingIntervalCount1").value = val(TXTReturnSallingIntervalCount1.text)
        
        
    ElseIf opt(7).value = True Then
        rs("ReturnSallingOption").value = 1
        rs("ReturnSallingIntervalCount").value = val(TXTReturnSallingIntervalCount(0).text)
        rs("ReturnSallingIntervalCount1").value = val(TXTReturnSallingIntervalCount1.text)
     
    End If

    If Option4(0).value = True Then
        rs("Save_options").value = 0
    ElseIf Option4(1).value = True Then
        rs("Save_options").value = 1
    ElseIf Option4(2).value = True Then
        rs("Save_options").value = 2
    ElseIf Option4(3).value = True Then
        rs("Save_options").value = 3
    End If

    rs("ShowLogoInReports").value = IIf(Me.chk.value = vbChecked, True, False)

    If Me.ImgPic.Picture <> 0 Then
        SavePictureToDB ImgPic, rs, "CompanyLogo"
    End If

    If Me.CboMainStockType.ListIndex = 0 Then
        rs("MainStockCostType").value = 0
    ElseIf Me.CboMainStockType.ListIndex = 1 Then
        rs("MainStockCostType").value = 2
    ElseIf Me.CboMainStockType.ListIndex = 2 Then
        rs("MainStockCostType").value = 4
    ElseIf Me.CboMainStockType.ListIndex = 3 Then
        rs("MainStockCostType").value = 5
    End If
    
        If Me.CboChasingStatus.ListIndex = 0 Then
        rs("ChasingStatus").value = 3
    ElseIf Me.CboChasingStatus.ListIndex = 1 Then
        rs("ChasingStatus").value = 1
    ElseIf Me.CboChasingStatus.ListIndex = 2 Then
        rs("ChasingStatus").value = 2
    ElseIf Me.CboChasingStatus.ListIndex = 3 Then
        rs("ChasingStatus").value = 7
    End If
    
    
    

    If Me.Chk1.value = vbChecked Then
        rs("AllowBoxNegative").value = 1
    ElseIf Me.Chk1.value = vbUnchecked Then
        rs("AllowBoxNegative").value = 0
    End If

    If Me.chk3.value = vbChecked Then
        rs("banks_Accounts").value = 1
    ElseIf Me.chk3.value = vbUnchecked Then
        rs("banks_Accounts").value = 0
    End If

    If Me.ChkBankComm.value = vbChecked Then
        rs("BankComm").value = 1
    ElseIf Me.ChkBankComm.value = vbUnchecked Then
        rs("BankComm").value = 0
    End If

    If Me.chkChequeBox.value = vbChecked Then
        rs("ChequeBox").value = 1
    ElseIf Me.ChkBankComm.value = vbUnchecked Then
        rs("ChequeBox").value = 0
    End If
    

    If Me.chkIsCheque.value = vbChecked Then
        rs("IsCheque").value = 1
    ElseIf Me.chkIsCheque.value = vbUnchecked Then
        rs("IsCheque").value = 0
    End If
        
    

    If Me.chkCustomerhavethreeAccounts.value = vbChecked Then
        rs("CustomerhavethreeAccounts").value = 1
    ElseIf Me.chkCustomerhavethreeAccounts.value = vbUnchecked Then
        rs("CustomerhavethreeAccounts").value = 0
    End If
    
    
   
    If Me.chkIsCreateOpenBalnceMan.value = vbChecked Then
        rs("IsCreateOpenBalnceMan").value = 1
    ElseIf Me.chkIsCreateOpenBalnceMan.value = vbUnchecked Then
        rs("IsCreateOpenBalnceMan").value = 0
    End If
     
    
    
    If Me.chkCustomerhavethreeAccounts1.value = vbChecked Then
        rs("CustomerhavethreeAccounts1").value = 1
    ElseIf Me.chkCustomerhavethreeAccounts1.value = vbUnchecked Then
        rs("CustomerhavethreeAccounts1").value = 0
    End If
    
    
    If Me.ChkTypicalProduction.value = vbChecked Then
        rs("TypicalProduction").value = 1
    ElseIf Me.ChkTypicalProduction.value = vbUnchecked Then
        rs("TypicalProduction").value = 0
    End If

    If Me.ChkExpensesCoding.value = vbChecked Then
        rs("ExpensesCoding").value = 1
    ElseIf Me.ChkExpensesCoding.value = vbUnchecked Then
        rs("ExpensesCoding").value = 0
    End If

    If Me.ChkExpensesCoding2.value = vbChecked Then
        rs("ExpensesCoding2").value = 1
    ElseIf Me.ChkExpensesCoding2.value = vbUnchecked Then
        rs("ExpensesCoding2").value = 0
    End If

    If Me.chkInstallmntsvchrCoding.value = vbChecked Then
        rs("InstallmntsvchrCoding").value = 1
    ElseIf Me.chkInstallmntsvchrCoding.value = vbUnchecked Then
        rs("InstallmntsvchrCoding").value = 0
    End If

    If Me.ChkAssetAccount.value = vbChecked Then
        rs("AssetAccount").value = 1
    ElseIf Me.ChkAssetAccount.value = vbUnchecked Then
        rs("AssetAccount").value = 0
    End If


'*******************************************************
    If Me.chkStore(0).value = vbChecked Then
        rs("StoreAccountHaveSettelment").value = 1
    ElseIf Me.chkStore(0).value = vbUnchecked Then
        rs("StoreAccountHaveSettelment").value = 0
    End If
    
   
    If Me.chkStore(7).value = vbChecked Then
        rs("CostStartingGard").value = 1
    ElseIf Me.chkStore(7).value = vbUnchecked Then
        rs("CostStartingGard").value = 0
    End If
    
  If Me.chkStore(8).value = vbChecked Then
        rs("TreatUncountedItemsAsZeroQty").value = 1
    ElseIf Me.chkStore(8).value = vbUnchecked Then
        rs("TreatUncountedItemsAsZeroQty").value = 0
    End If
     
    
    If Me.chkStore(4).value = vbChecked Then
        rs("IsAutoNameItems").value = 1
    ElseIf Me.chkStore(4).value = vbUnchecked Then
        rs("IsAutoNameItems").value = 0
    End If
    
       If Me.chkStore(5).value = vbChecked Then
        rs("AllowItemByRowMove").value = 1
    ElseIf Me.chkStore(5).value = vbUnchecked Then
        rs("AllowItemByRowMove").value = 0
    End If
        If Me.chkStore(6).value = vbChecked Then
        rs("AllowItemByRowOut").value = 1
    ElseIf Me.chkStore(6).value = vbUnchecked Then
        rs("AllowItemByRowOut").value = 0
    End If
    
    
    If Me.chkStore(1).value = vbChecked Then
        rs("eachStoreHaveLossAccount").value = 1
    ElseIf Me.chkStore(1).value = vbUnchecked Then
        rs("eachStoreHaveLossAccount").value = 0
    End If
    
    
    
    If Me.chkStore(2).value = vbChecked Then
        rs("eachStoreHaveGiftAccount").value = 1
    ElseIf Me.chkStore(2).value = vbUnchecked Then
        rs("eachStoreHaveGiftAccount").value = 0
    End If
       If Me.chkStore(3).value = vbChecked Then
        rs("MultyStore").value = 1
    ElseIf Me.chkStore(3).value = vbUnchecked Then
        rs("MultyStore").value = 0
    End If
    
'*********************************************************



    If Me.ChkAssetAccount1.value = vbChecked Then
        rs("AssetAccount1").value = 1
    ElseIf Me.ChkAssetAccount1.value = vbUnchecked Then
        rs("AssetAccount1").value = 0
    End If

    If Me.ChkAllowIndirectCost.value = vbChecked Then
        rs("AllowIndirectCost").value = 1
    ElseIf Me.ChkAllowIndirectCost.value = vbUnchecked Then
        rs("AllowIndirectCost").value = 0
    End If



    If Me.chkEmpProduction.value = vbChecked Then
        rs("EmpProduction").value = 1
    ElseIf Me.chkEmpProduction.value = vbUnchecked Then
        rs("EmpProduction").value = 0
    End If



    If Me.chkItemProduction.value = vbChecked Then
        rs("ItemProduction").value = 1
    ElseIf Me.chkItemProduction.value = vbUnchecked Then
        rs("ItemProduction").value = 0
    End If


    If Me.chkExpProduction.value = vbChecked Then
        rs("ExpProduction").value = 1
    ElseIf Me.chkExpProduction.value = vbUnchecked Then
        rs("ExpProduction").value = 0
    End If



    If Me.chk2.value = vbChecked Then
        rs("AllowStockNegative").value = 1
    ElseIf Me.chk2.value = vbUnchecked Then
        rs("AllowStockNegative").value = 0
    End If

    rs("checkout").value = IIf(Opt_OrderOut.value = True, 1, 0)
    rs("EmpRes").value = IIf(ChkEmpRes(0).value = True, 0, 1)

    rs("Checksal").value = IIf(Opt_Sal.value = True, 1, 0)
    rs("checkinpo").value = IIf(Opt_OrderInpo.value = True, 1, 0)

    If OptionItemsTotal.value = True Then
        rs("Items_or_operation").value = 0
    ElseIf OptionOperation.value = True Then
        rs("Items_or_operation").value = 1
    Else
        rs("Items_or_operation").value = Null
    End If
    
    
        If OPTdISCOUNT(0).value = True Then
        rs("ProjectDiscountPolicy").value = 0
    Else
        rs("ProjectDiscountPolicy").value = 1
    
    End If
    
    
    
    'OPT ProjectDiscountPolicy
    

    If GlDetails.value = True Then
        rs("gl_detaila_or_total").value = 1
    ElseIf glgeneral.value = True Then
        rs("gl_detaila_or_total").value = 0
    Else
        rs("gl_detaila_or_total").value = Null
    End If

    If Me.Optday.value = True Then
        rs("ProcessPeriodType").value = 0
    ElseIf Me.OptMonth.value = True Then
        rs("ProcessPeriodType").value = 1
    ElseIf Me.OptYear.value = True Then
        rs("ProcessPeriodType").value = 2

    ElseIf Me.Optweek.value = True Then
        rs("ProcessPeriodType").value = 3
    Else
        rs("ProcessPeriodType").value = Null
    End If

    rs("checkbey").value = IIf(Opt_Bey.value = True, 1, 0)

'    rs("Opt_branch").value = IIf(opt_Branch.value = True, 1, 0)

'    rs("opt_group").value = IIf(opt_group.value = True, 1, 0)

    If OptArrowBranch.value = True Then
        rs("Arrows_group").value = 0
    Else
        rs("Arrows_group").value = 1
    End If

'    rs("Opt_Inventory_create_account").value = IIf(Opt_Inventory_create_account.value = True, 1, 0)
'    rs("opt_inv_and_branch_create_account").value = IIf(opt_inv_and_branch_create_account.value = True, 1, 0)

    'If Me.Chk3.Value = vbChecked Then
    '     Rs("OUT").Value = 1
    'ElseIf Me.Chk3.Value = vbUnchecked Then
    '     Rs("OUT").Value = 0
    'End If

    'If Me.Check1.Value = vbChecked Then
    '     Rs("Inp").Value = 1
    'ElseIf Me.Check1.Value = vbUnchecked Then
    '     Rs("Inp").Value = 0
    'End If

    rs("CurrencyDigts").value = val(Me.TxtPriceDigts.text)
    rs("PriceDigtsInst").value = val(Me.TxtPriceDigtsInst.text)

    rs("EmpComponentDigts").value = val(Me.TxtEmpComponentDigts.text)
    rs("EmpSalaryDigts").value = val(Me.TxtEmpSalaryDigts.text)
    rs("IndirectCostPercentage").value = val(Me.TxtIndirectCostPercentage.text)
    
    rs("StoreDigit").value = IIf(val(Me.TxtData(1).text) > 0, val(Me.TxtData(1).text), 1)
     rs("BranchDigit").value = IIf(val(Me.TxtData(0).text) > 0, val(Me.TxtData(0).text), 1)
    

    rs("Ked_digit").value = val(Me.TxtData(2).text)
    rs("Count_ACCOUNT_digit").value = val(Me.txt_ACCOUNT_digit.text)
 
    
    rs("LimitDefaultCredit").value = val(Me.txtLimitDefaultCredit.text)
    rs("LimitDefaultCreditDays").value = val(Me.txtLimitDefaultCreditDays.text)
        
    rs("SaleDiscount1").value = val(Me.TxtSaleDiscount1.text)
    rs("SaleDiscount2").value = val(Me.TxtSaleDiscount2.text)
    rs("SaleDiscount3").value = val(Me.TxtSaleDiscount3.text)
    rs("SaleDiscount4").value = val(Me.TxtSaleDiscount4.text)

    rs("QtyDigts").value = val(Me.TxtQtyDigts.text)

    If Me.Chkemployeeaccounts.value = vbChecked Then
        rs("Create_employee_account").value = 1
    ElseIf Me.Chkemployeeaccounts.value = vbUnchecked Then
        rs("Create_employee_account").value = 0
    End If

    If Me.ChkDriverBox.value = vbChecked Then
        rs("CreateDriverBox").value = 1
    ElseIf Me.ChkDriverBox.value = vbUnchecked Then
        rs("CreateDriverBox").value = 0
    End If

    If Me.chkDriverEra.value = vbChecked Then
        rs("CreateDriverEra").value = 1
    ElseIf Me.chkDriverEra.value = vbUnchecked Then
        rs("CreateDriverEra").value = 0
    End If

    If Me.ChkitemsWorkWithSize.value = vbChecked Then
        rs("itemsWorkWithSize").value = 1
    ElseIf Me.ChkitemsWorkWithSize.value = vbUnchecked Then
        rs("itemsWorkWithSize").value = 0
    End If


    If Me.Chkbarcode(0).value = vbChecked Then
        rs("WorkWithBarCode").value = 1
    ElseIf Me.Chkbarcode(0).value = vbUnchecked Then
        rs("WorkWithBarCode").value = 0
    End If

    If Me.Chkbarcode(38).value = vbChecked Then
        rs("WorkWithBarCodeParent").value = 1
    ElseIf Me.Chkbarcode(38).value = vbUnchecked Then
        rs("WorkWithBarCodeParent").value = 0
    End If
    
  If Me.Chkbarcode(13).value = vbChecked Then
        rs("DefaultIsCreditSales").value = 1
    ElseIf Me.Chkbarcode(13).value = vbUnchecked Then
        rs("DefaultIsCreditSales").value = 0
    End If
      If Me.Chkbarcode(46).value = vbChecked Then
        rs("DefaultIsCreditPurchase").value = 1
    ElseIf Me.Chkbarcode(46).value = vbUnchecked Then
        rs("DefaultIsCreditPurchase").value = 0
    End If
    
    
    If Me.Chkbarcode(137).value = vbChecked Then
        rs("DefaultIsCreditPurchaseRet").value = 1
    ElseIf Me.Chkbarcode(137).value = vbUnchecked Then
        rs("DefaultIsCreditPurchaseRet").value = 0
    End If
    If Me.Chkbarcode(138).value = vbChecked Then
        rs("OpenAccountAqar").value = 1
    ElseIf Me.Chkbarcode(138).value = vbUnchecked Then
        rs("OpenAccountAqar").value = 0
    End If
    If Me.Chkbarcode(139).value = vbChecked Then
        rs("InvoiceTransferJLTotal").value = 1
    ElseIf Me.Chkbarcode(139).value = vbUnchecked Then
        rs("InvoiceTransferJLTotal").value = 0
    End If
    
   If Me.Chkbarcode(146).value = vbChecked Then
        rs("CarsRevenuePerOwner").value = 1
    ElseIf Me.Chkbarcode(146).value = vbUnchecked Then
        rs("CarsRevenuePerOwner").value = 0
    End If
    
   If Me.Chkbarcode(147).value = vbChecked Then
        rs("DontShowMoreDetailsCompItem").value = 1
    ElseIf Me.Chkbarcode(147).value = vbUnchecked Then
        rs("DontShowMoreDetailsCompItem").value = 0
    End If
    
   If Me.Chkbarcode(153).value = vbChecked Then
        rs("CompilingBasedTable").value = 1
    ElseIf Me.Chkbarcode(153).value = vbUnchecked Then
        rs("CompilingBasedTable").value = 0
    End If
    
   If Me.Chkbarcode(154).value = vbChecked Then
        rs("CanPartialpayment").value = 1
    ElseIf Me.Chkbarcode(154).value = vbUnchecked Then
        rs("CanPartialpayment").value = 0
    End If
    
    
   If Me.Chkbarcode(155).value = vbChecked Then
        rs("EndRentifPayed").value = 1
    ElseIf Me.Chkbarcode(155).value = vbUnchecked Then
        rs("EndRentifPayed").value = 0
    End If
        
        
   If Me.Chkbarcode(156).value = vbChecked Then
        rs("cantCahngeAkarinExpenses").value = 1
    ElseIf Me.Chkbarcode(156).value = vbUnchecked Then
        rs("cantCahngeAkarinExpenses").value = 0
    End If
    
    
  If Me.Chkbarcode(157).value = vbChecked Then
        rs("EmployeeSalaryBYBranch").value = 1
    ElseIf Me.Chkbarcode(157).value = vbUnchecked Then
        rs("EmployeeSalaryBYBranch").value = 0
    End If
       
       
  If Me.Chkbarcode(158).value = vbChecked Then
        rs("returnnotcreatvoucher").value = 1
    ElseIf Me.Chkbarcode(158).value = vbUnchecked Then
        rs("returnnotcreatvoucher").value = 0
    End If
              
              
             If Me.Chkbarcode(186).value = vbChecked Then
        rs("OnlyOneCashingVchr").value = 1
    ElseIf Me.Chkbarcode(186).value = vbUnchecked Then
        rs("OnlyOneCashingVchr").value = 0
    End If
                 
                 
                          If Me.Chkbarcode(187).value = vbChecked Then
        rs("CheckDateFormatCorrect").value = 1
    ElseIf Me.Chkbarcode(187).value = vbUnchecked Then
        rs("CheckDateFormatCorrect").value = 0
    End If
                  
                                    If Me.Chkbarcode(190).value = vbChecked Then
        rs("CheckMobileFormatCorrect").value = 1
    ElseIf Me.Chkbarcode(190).value = vbUnchecked Then
        rs("CheckMobileFormatCorrect").value = 0
    End If
                          

    If Me.Chkbarcode(191).value = vbChecked Then
        rs("IsShowLensesDetails").value = 1
    ElseIf Me.Chkbarcode(191).value = vbUnchecked Then
        rs("IsShowLensesDetails").value = 0
    End If
                          


                  
                                        If Me.Chkbarcode(188).value = vbChecked Then
        rs("CantRepetttransferNoforCashing").value = 1
    ElseIf Me.Chkbarcode(188).value = vbUnchecked Then
        rs("CantRepetttransferNoforCashing").value = 0
    End If
                  
                                                If Me.Chkbarcode(189).value = vbChecked Then
        rs("DontDuplicateManulaNoInPurchase").value = 1
    ElseIf Me.Chkbarcode(189).value = vbUnchecked Then
        rs("DontDuplicateManulaNoInPurchase").value = 0
    End If
                                
                      
                      

   If Me.Chkbarcode(159).value = vbChecked Then
        rs("WaiverSetByContract").value = 1
    ElseIf Me.Chkbarcode(159).value = vbUnchecked Then
        rs("WaiverSetByContract").value = 0
    End If
                  
                  
   If Me.Chkbarcode(160).value = vbChecked Then
        rs("IsGeometricProportions").value = 1
    ElseIf Me.Chkbarcode(160).value = vbUnchecked Then
        rs("IsGeometricProportions").value = 0
    End If
                  
                  
                  
     If Me.Chkbarcode(162).value = vbChecked Then
        rs("IsSerialByUserTrans").value = 1
    ElseIf Me.Chkbarcode(162).value = vbUnchecked Then
        rs("IsSerialByUserTrans").value = 0
    End If
                                 
                  
     If Me.Chkbarcode(164).value = vbChecked Then
        rs("AllowRepeateCar").value = 1
    ElseIf Me.Chkbarcode(164).value = vbUnchecked Then
        rs("AllowRepeateCar").value = 0
    End If
                                 
                                 
                  
     If Me.Chkbarcode(163).value = vbChecked Then
        rs("IsSerialByUserVouch").value = 1
    ElseIf Me.Chkbarcode(163).value = vbUnchecked Then
        rs("IsSerialByUserVouch").value = 0
    End If
                                 
              
   If Me.Chkbarcode(148).value = vbChecked Then
        rs("traveDiscountFromCustomerDirect").value = 1
    ElseIf Me.Chkbarcode(148).value = vbUnchecked Then
        rs("traveDiscountFromCustomerDirect").value = 0
    End If
    
    
   If Me.Chkbarcode(149).value = vbChecked Then
        rs("IsCustSalesManCashRelated").value = 1
    ElseIf Me.Chkbarcode(149).value = vbUnchecked Then
        rs("IsCustSalesManCashRelated").value = 0
    End If
    

   If Me.Chkbarcode(150).value = vbChecked Then
        rs("showEmployeeAccountIntrip").value = 1
    ElseIf Me.Chkbarcode(150).value = vbUnchecked Then
        rs("showEmployeeAccountIntrip").value = 0
    End If
    
    
   If Me.Chkbarcode(151).value = vbChecked Then
        rs("DUEDOCUMENTbyinstallDate").value = 1
    ElseIf Me.Chkbarcode(151).value = vbUnchecked Then
        rs("DUEDOCUMENTbyinstallDate").value = 0
    End If
        
        
   If Me.Chkbarcode(152).value = vbChecked Then
        rs("CanSkipPurchOrder").value = 1
    ElseIf Me.Chkbarcode(152).value = vbUnchecked Then
        rs("CanSkipPurchOrder").value = 0
    End If
        
        
        
          If Me.Chkbarcode(47).value = vbChecked Then
        rs("returnByBarCodeOnly").value = 1
    ElseIf Me.Chkbarcode(47).value = vbUnchecked Then
        rs("returnByBarCodeOnly").value = 0
    End If
    
    
 If Me.Chkbarcode(14).value = vbChecked Then
        rs("JLCodeBasedOnBranch").value = 1
    ElseIf Me.Chkbarcode(14).value = vbUnchecked Then
        rs("JLCodeBasedOnBranch").value = 0
    End If
    
 If Me.Chkbarcode(15).value = vbChecked Then
        rs("EmpNotExcceedDiscount").value = 1
    ElseIf Me.Chkbarcode(15).value = vbUnchecked Then
        rs("EmpNotExcceedDiscount").value = 0
    End If
    
 If Me.Chkbarcode(16).value = vbChecked Then
        rs("BoxLossandIncreae").value = 1
    ElseIf Me.Chkbarcode(16).value = vbUnchecked Then
        rs("BoxLossandIncreae").value = 0
    End If
    
    
     If Me.Chkbarcode(17).value = vbChecked Then
        rs("attacheditemsisfree").value = 1
    ElseIf Me.Chkbarcode(17).value = vbUnchecked Then
        rs("attacheditemsisfree").value = 0
    End If
    
    If Me.Chkbarcode(34).value = vbChecked Then
        rs("EnableCustomerAging").value = 1
    ElseIf Me.Chkbarcode(34).value = vbUnchecked Then
        rs("EnableCustomerAging").value = 0
    End If
    
    
    If Me.Chkbarcode(18).value = vbChecked Then
        rs("showcostColorininvoice").value = 1
    ElseIf Me.Chkbarcode(18).value = vbUnchecked Then
        rs("showcostColorininvoice").value = 0
    End If
        
        
    
       If Me.Chkbarcode(19).value = vbChecked Then
        rs("SubContactorHave3Account").value = 1
    ElseIf Me.Chkbarcode(19).value = vbUnchecked Then
        rs("SubContactorHave3Account").value = 0
    End If
         
         
      If Me.Chkbarcode(20).value = vbChecked Then
        rs("ProjectEmployeeGV").value = 1
    ElseIf Me.Chkbarcode(20).value = vbUnchecked Then
        rs("ProjectEmployeeGV").value = 0
    End If
    
       If Me.Chkbarcode(21).value = vbChecked Then
        rs("PursgaseWithoutDecimal").value = 1
    ElseIf Me.Chkbarcode(21).value = vbUnchecked Then
        rs("PursgaseWithoutDecimal").value = 0
    End If
    
    
           If Me.Chkbarcode(22).value = vbChecked Then
        rs("workWithCustomerContract").value = 1
    ElseIf Me.Chkbarcode(22).value = vbUnchecked Then
        rs("workWithCustomerContract").value = 0
    End If
    
          If Me.Chkbarcode(25).value = vbChecked Then
        rs("workWithVendorContract").value = 1
    ElseIf Me.Chkbarcode(25).value = vbUnchecked Then
        rs("workWithVendorContract").value = 0
    End If
    
          If Me.Chkbarcode(26).value = vbChecked Then
        rs("PoCreateVoucher").value = 1
    ElseIf Me.Chkbarcode(26).value = vbUnchecked Then
        rs("PoCreateVoucher").value = 0
    End If
    
     
             If Me.Chkbarcode(27).value = vbChecked Then
        rs("poWithatotalQty").value = 1
    ElseIf Me.Chkbarcode(27).value = vbUnchecked Then
        rs("poWithatotalQty").value = 0
    End If
    
    
             If Me.Chkbarcode(28).value = vbChecked Then
        rs("DiscountSalesCreateVchr").value = 1
    ElseIf Me.Chkbarcode(28).value = vbUnchecked Then
        rs("DiscountSalesCreateVchr").value = 0
    End If
        
        
                     If Me.Chkbarcode(39).value = vbChecked Then
        rs("AllowCostPerStore").value = 1
    ElseIf Me.Chkbarcode(39).value = vbUnchecked Then
        rs("AllowCostPerStore").value = 0
    End If
        
        
                     If Me.Chkbarcode(40).value = vbChecked Then
        rs("AllowCostnNewShape").value = 1
    ElseIf Me.Chkbarcode(40).value = vbUnchecked Then
        rs("AllowCostnNewShape").value = 0
    End If
    
    
                         If Me.Chkbarcode(41).value = vbChecked Then
        rs("AllowCostBySerial").value = 1
    ElseIf Me.Chkbarcode(41).value = vbUnchecked Then
        rs("AllowCostBySerial").value = 0
    End If
    
    
    
        
             If Me.Chkbarcode(29).value = vbChecked Then
        rs("PaymentDifferent").value = 1
    ElseIf Me.Chkbarcode(29).value = vbUnchecked Then
        rs("PaymentDifferent").value = 0
    End If
                
                
                   If Me.Chkbarcode(30).value = vbChecked Then
        rs("PayrollOneAccount").value = 1
    ElseIf Me.Chkbarcode(30).value = vbUnchecked Then
        rs("PayrollOneAccount").value = 0
    End If
    
    
    
                   If Me.Chkbarcode(31).value = vbChecked Then
        rs("WorkWithItemsDetails").value = 1
    ElseIf Me.Chkbarcode(31).value = vbUnchecked Then
        rs("WorkWithItemsDetails").value = 0
    End If
        
        
            
                   If Me.Chkbarcode(32).value = vbChecked Then
        rs("FAAddtionCreateAccount").value = 1
    ElseIf Me.Chkbarcode(32).value = vbUnchecked Then
        rs("FAAddtionCreateAccount").value = 0
    End If
        
                           If Me.Chkbarcode(33).value = vbChecked Then
        rs("Create2account4Supp").value = 1
    ElseIf Me.Chkbarcode(33).value = vbUnchecked Then
        rs("Create2account4Supp").value = 0
    End If
        
        
          If Me.Chkbarcode(23).value = vbChecked Then
        rs("cancellAllApprove").value = 1
    ElseIf Me.Chkbarcode(23).value = vbUnchecked Then
        rs("cancellAllApprove").value = 0
    End If
         
         
        
          If Me.Chkbarcode(24).value = vbChecked Then
        rs("workwithticketAllocation").value = 1
    ElseIf Me.Chkbarcode(24).value = vbUnchecked Then
        rs("workwithticketAllocation").value = 0
    End If
              
              
         
     If Me.Chkbarcode(10).value = vbChecked Then
        rs("WorkWithGroupCode").value = 1
    ElseIf Me.Chkbarcode(10).value = vbUnchecked Then
        rs("WorkWithGroupCode").value = 0
    End If
    
     If Me.Chkbarcode(9).value = vbChecked Then
        rs("WorkWithFirstInstallOnly").value = 1
    ElseIf Me.Chkbarcode(9).value = vbUnchecked Then
        rs("WorkWithFirstInstallOnly").value = 0
    End If
    
    
    If Me.Chkbarcode(11).value = vbChecked Then
        rs("CreateInsuranceAccountForCustomers").value = 1
    ElseIf Me.Chkbarcode(11).value = vbUnchecked Then
        rs("CreateInsuranceAccountForCustomers").value = 0
    End If
    
    If Me.Chkbarcode(12).value = vbChecked Then
        rs("DecideItemName").value = 1
    ElseIf Me.Chkbarcode(12).value = vbUnchecked Then
        rs("DecideItemName").value = 0
    End If
    
    
     If Me.Chkbarcode(7).value = vbChecked Then
        rs("WorkWithLINKEDiTEMS").value = 1
    ElseIf Me.Chkbarcode(7).value = vbUnchecked Then
        rs("WorkWithLINKEDiTEMS").value = 0
    End If
    
    
     If Me.Chkbarcode(192).value = vbChecked Then
        rs("WorkWithLINKEDiActivity").value = 1
    ElseIf Me.Chkbarcode(192).value = vbUnchecked Then
        rs("WorkWithLINKEDiActivity").value = 0
    End If
    
         If Me.Chkbarcode(193).value = vbChecked Then
        rs("amlaketbatrentOnly").value = 1
    ElseIf Me.Chkbarcode(193).value = vbUnchecked Then
        rs("amlaketbatrentOnly").value = 0
    End If
    
        
    If Me.Chkbarcode(194).value = vbChecked Then
        rs("NotAllowStockNegativeInternal").value = 1
    ElseIf Me.Chkbarcode(194).value = vbUnchecked Then
        rs("NotAllowStockNegativeInternal").value = 0
    End If
    
    
    
        If Me.Chkbarcode(195).value = vbChecked Then
        rs("MustEnterNewNo").value = 1
    ElseIf Me.Chkbarcode(195).value = vbUnchecked Then
        rs("MustEnterNewNo").value = 0
    End If
    
    
        If Me.Chkbarcode(196).value = vbChecked Then
        rs("IsInternalMultiOrder").value = 1
    ElseIf Me.Chkbarcode(196).value = vbUnchecked Then
        rs("IsInternalMultiOrder").value = 0
    End If
    
    
    
    If Me.Chkbarcode(197).value = vbChecked Then
        rs("IsBlue").value = 1
    ElseIf Me.Chkbarcode(197).value = vbUnchecked Then
        rs("IsBlue").value = 0
    End If
    
    
   
    
    
    If Me.Chkbarcode(198).value = vbChecked Then
        rs("Isthickness").value = 1
    ElseIf Me.Chkbarcode(198).value = vbUnchecked Then
        rs("Isthickness").value = 0
    End If
    
    
    
    If Me.Chkbarcode(199).value = vbChecked Then
        rs("IsMashghal").value = 1
    ElseIf Me.Chkbarcode(199).value = vbUnchecked Then
        rs("IsMashghal").value = 0
    End If
    

        If Me.Chkbarcode(200).value = vbChecked Then
        rs("IsSalesOrder").value = 1
    ElseIf Me.Chkbarcode(200).value = vbUnchecked Then
        rs("IsSalesOrder").value = 0
    End If
    
     
    

    If Me.Chkbarcode(201).value = vbChecked Then
        rs("IsQrCodePrint").value = 1
    ElseIf Me.Chkbarcode(201).value = vbUnchecked Then
        rs("IsQrCodePrint").value = 0
    End If
    
   
    If Me.Chkbarcode(202).value = vbChecked Then
        rs("IsShowItemsBranch").value = 1
    ElseIf Me.Chkbarcode(202).value = vbUnchecked Then
        rs("IsShowItemsBranch").value = 0
    End If
     
    
    If Me.Chkbarcode(204).value = vbChecked Then
        rs("IsElecWaterCont").value = 1
    ElseIf Me.Chkbarcode(204).value = vbUnchecked Then
        rs("IsElecWaterCont").value = 0
    End If
        
    
    If Me.Chkbarcode(205).value = vbChecked Then
        rs("IsDogeMode").value = 1
    ElseIf Me.Chkbarcode(205).value = vbUnchecked Then
        rs("IsDogeMode").value = 0
    End If
        
    
    
    If Me.Chkbarcode(206).value = vbChecked Then
        rs("IsMaintItemMode").value = 1
    ElseIf Me.Chkbarcode(206).value = vbUnchecked Then
        rs("IsMaintItemMode").value = 0
    End If
        
        
  
    If Me.Chkbarcode(211).value = vbChecked Then
        rs("IsHiddenTransportInv").value = 1
    ElseIf Me.Chkbarcode(211).value = vbUnchecked Then
        rs("IsHiddenTransportInv").value = 0
    End If
         
        
    If Me.Chkbarcode(207).value = vbChecked Then
        rs("IsHeaderPrint").value = 1
    ElseIf Me.Chkbarcode(207).value = vbUnchecked Then
        rs("IsHeaderPrint").value = 0
    End If
            
                    
               

    
    
    If Me.Chkbarcode(8).value = vbChecked Then
        rs("WorkWithBranchLogo").value = 1
    ElseIf Me.Chkbarcode(8).value = vbUnchecked Then
        rs("WorkWithBranchLogo").value = 0
    End If
    
    If Me.Chkbarcode(1).value = vbChecked Then
        rs("DuplicateitemsNames").value = 1
    ElseIf Me.Chkbarcode(1).value = vbUnchecked Then
        rs("DuplicateitemsNames").value = 0
    End If
    
    
        If Me.Chkbarcode(2).value = vbChecked Then
        rs("TradingPOS").value = 1
    ElseIf Me.Chkbarcode(2).value = vbUnchecked Then
        rs("TradingPOS").value = 0
    End If
    
    If Me.Chkbarcode(87).value = vbChecked Then
        rs("posshape2").value = 1
    ElseIf Me.Chkbarcode(87).value = vbUnchecked Then
        rs("posshape2").value = 0
    End If
    
    If Me.Chkbarcode(88).value = vbChecked Then
        rs("InsuranceOnOwner").value = 1
    ElseIf Me.Chkbarcode(88).value = vbUnchecked Then
        rs("InsuranceOnOwner").value = 0
    End If
     If Me.Chkbarcode(89).value = vbChecked Then
        rs("ServicesOnOwner").value = 1
    ElseIf Me.Chkbarcode(89).value = vbUnchecked Then
        rs("ServicesOnOwner").value = 0
    End If
    If Me.Chkbarcode(90).value = vbChecked Then
        rs("DueComm").value = 1
    ElseIf Me.Chkbarcode(90).value = vbUnchecked Then
        rs("DueComm").value = 0
    End If
     If Me.Chkbarcode(91).value = vbChecked Then
        rs("DueWater").value = 1
    ElseIf Me.Chkbarcode(91).value = vbUnchecked Then
        rs("DueWater").value = 0
    End If
    If Me.Chkbarcode(92).value = vbChecked Then
        rs("DueElectr").value = 1
    ElseIf Me.Chkbarcode(92).value = vbUnchecked Then
        rs("DueElectr").value = 0
    End If
    If Me.Chkbarcode(93).value = vbChecked Then
        rs("DueService").value = 1
    ElseIf Me.Chkbarcode(93).value = vbUnchecked Then
        rs("DueService").value = 0
    End If
    If Me.Chkbarcode(94).value = vbChecked Then
        rs("CommissionOnOwner").value = 1
    ElseIf Me.Chkbarcode(94).value = vbUnchecked Then
        rs("CommissionOnOwner").value = 0
    End If
    
    If Me.Chkbarcode(95).value = vbChecked Then
        rs("CommissionDue").value = 1
    ElseIf Me.Chkbarcode(95).value = vbUnchecked Then
        rs("CommissionDue").value = 0
    End If
    
    If Me.Chkbarcode(96).value = vbChecked Then
        rs("SupplierReciveGE").value = 1
    ElseIf Me.Chkbarcode(96).value = vbUnchecked Then
        rs("SupplierReciveGE").value = 0
    End If
    
    If Me.Chkbarcode(97).value = vbChecked Then
        rs("BranchmustimSalary").value = 1
    ElseIf Me.Chkbarcode(97).value = vbUnchecked Then
        rs("BranchmustimSalary").value = 0
    End If
        If Me.Chkbarcode(98).value = vbChecked Then
        rs("AllowSkipPayment").value = 1
    ElseIf Me.Chkbarcode(98).value = vbUnchecked Then
        rs("AllowSkipPayment").value = 0
    End If
      If Me.Chkbarcode(99).value = vbChecked Then
        rs("AllowChangePriceApprove").value = 1
    ElseIf Me.Chkbarcode(99).value = vbUnchecked Then
        rs("AllowChangePriceApprove").value = 0
    End If

    If Me.Chkbarcode(101).value = vbChecked Then
        rs("CreateJLVactionAratha").value = 1
    ElseIf Me.Chkbarcode(101).value = vbUnchecked Then
        rs("CreateJLVactionAratha").value = 0
    End If
    
    If Me.Chkbarcode(102).value = vbChecked Then
        rs("PriceWithVAT").value = 1
    ElseIf Me.Chkbarcode(102).value = vbUnchecked Then
        rs("PriceWithVAT").value = 0
    End If
    If Me.Chkbarcode(103).value = vbChecked Then
        rs("AllowWorkCustomerPoints").value = 1
    ElseIf Me.Chkbarcode(103).value = vbUnchecked Then
        rs("AllowWorkCustomerPoints").value = 0
    End If
    If Me.Chkbarcode(104).value = vbChecked Then
        rs("ProjectInvoiceAnalysisJL").value = 1
    ElseIf Me.Chkbarcode(104).value = vbUnchecked Then
        rs("ProjectInvoiceAnalysisJL").value = 0
    End If
   If Me.Chkbarcode(105).value = vbChecked Then
        rs("CustomerRecordNoIsnotManda").value = 1
    ElseIf Me.Chkbarcode(105).value = vbUnchecked Then
        rs("CustomerRecordNoIsnotManda").value = 0
    End If
     If Me.Chkbarcode(106).value = vbChecked Then
        rs("DealingWithPrepayAccount").value = 1
    ElseIf Me.Chkbarcode(106).value = vbUnchecked Then
        rs("DealingWithPrepayAccount").value = 0
    End If
     If Me.Chkbarcode(3).value = vbChecked Then
        rs("CanChanegeLinkedSsalesnvoice").value = 1
    ElseIf Me.Chkbarcode(3).value = vbUnchecked Then
        rs("CanChanegeLinkedSsalesnvoice").value = 0
    End If
     
     If Me.Chkbarcode(107).value = vbChecked Then
        rs("NotAllowedCalcVata").value = 1
    ElseIf Me.Chkbarcode(107).value = vbUnchecked Then
        rs("NotAllowedCalcVata").value = 0
    End If
         If Me.Chkbarcode(108).value = vbChecked Then
        rs("IssueVoucherWorkWithRemain").value = 1
    ElseIf Me.Chkbarcode(108).value = vbUnchecked Then
        rs("IssueVoucherWorkWithRemain").value = 0
    End If
         If Me.Chkbarcode(109).value = vbChecked Then
        rs("TripDateInsertDefulat").value = 1
    ElseIf Me.Chkbarcode(109).value = vbUnchecked Then
        rs("TripDateInsertDefulat").value = 0
    End If
     
     
         If Me.Chkbarcode(112).value = vbChecked Then
        rs("TripwithorderOnly").value = 1
    ElseIf Me.Chkbarcode(112).value = vbUnchecked Then
        rs("TripwithorderOnly").value = 0
    End If
    
    
     If Me.Chkbarcode(113).value = vbChecked Then
        rs("AllowPriceWithWidth").value = 1
    ElseIf Me.Chkbarcode(113).value = vbUnchecked Then
        rs("AllowPriceWithWidth").value = 0
    End If
    If Me.Chkbarcode(114).value = vbChecked Then
        rs("LinkCustomerWithCars").value = 1
    ElseIf Me.Chkbarcode(114).value = vbUnchecked Then
        rs("LinkCustomerWithCars").value = 0
    End If
    If Me.Chkbarcode(115).value = vbChecked Then
        rs("AllowEditCashingLinkProj").value = 1
    ElseIf Me.Chkbarcode(115).value = vbUnchecked Then
        rs("AllowEditCashingLinkProj").value = 0
    End If
    
    If Me.Chkbarcode(116).value = vbChecked Then
        rs("TransBillPriceByGrid").value = 1
    ElseIf Me.Chkbarcode(116).value = vbUnchecked Then
        rs("TransBillPriceByGrid").value = 0
    End If
    If Me.Chkbarcode(117).value = vbChecked Then
        rs("NoCreatJLInRentContract").value = 1
    ElseIf Me.Chkbarcode(117).value = vbUnchecked Then
        rs("NoCreatJLInRentContract").value = 0
    End If
    If Me.Chkbarcode(118).value = vbChecked Then
        rs("OpenVATAccountOwner").value = 1
    ElseIf Me.Chkbarcode(118).value = vbUnchecked Then
        rs("OpenVATAccountOwner").value = 0
    End If
    If Me.Chkbarcode(119).value = vbChecked Then
        rs("CreateJLEmpCommissions").value = 1
    ElseIf Me.Chkbarcode(119).value = vbUnchecked Then
        rs("CreateJLEmpCommissions").value = 0
    End If
    If Me.Chkbarcode(120).value = vbChecked Then
        rs("TypeContractAutoFromIqar").value = 1
    ElseIf Me.Chkbarcode(120).value = vbUnchecked Then
        rs("TypeContractAutoFromIqar").value = 0
    End If
    If Me.Chkbarcode(121).value = vbChecked Then
        rs("AllowRepeatInvoiceNo").value = 1
    ElseIf Me.Chkbarcode(121).value = vbUnchecked Then
        rs("AllowRepeatInvoiceNo").value = 0
    End If
    
    If Me.Chkbarcode(122).value = vbChecked Then
        rs("AllowReturnFIFO").value = 1
    ElseIf Me.Chkbarcode(122).value = vbUnchecked Then
        rs("AllowReturnFIFO").value = 0
    End If
    If Me.Chkbarcode(123).value = vbChecked Then
        rs("AllowDiscountAllowedFIFO").value = 1
    ElseIf Me.Chkbarcode(123).value = vbUnchecked Then
        rs("AllowDiscountAllowedFIFO").value = 0
    End If
    If Me.Chkbarcode(124).value = vbChecked Then
        rs("AllowJLManualFIFO").value = 1
    ElseIf Me.Chkbarcode(124).value = vbUnchecked Then
        rs("AllowJLManualFIFO").value = 0
    End If
    
    
    If Me.Chkbarcode(161).value = vbChecked Then
        rs("IsMergeVat").value = 1
    ElseIf Me.Chkbarcode(124).value = vbUnchecked Then
        rs("IsMergeVat").value = 0
    End If
 
    
    If Me.Chkbarcode(125).value = vbChecked Then
        rs("ShowBalanceOfEmpInSalary").value = 1
    ElseIf Me.Chkbarcode(125).value = vbUnchecked Then
        rs("ShowBalanceOfEmpInSalary").value = 0
    End If
    If Me.Chkbarcode(126).value = vbChecked Then
        rs("PaymentIntoAccouStat").value = 1
    ElseIf Me.Chkbarcode(126).value = vbUnchecked Then
        rs("PaymentIntoAccouStat").value = 0
    End If
    
       If Me.Chkbarcode(127).value = vbChecked Then
        rs("AllowEditInvoiceNoticeDiscount").value = 1
    ElseIf Me.Chkbarcode(127).value = vbUnchecked Then
        rs("AllowEditInvoiceNoticeDiscount").value = 0
    End If
        If Me.Chkbarcode(128).value = vbChecked Then
        rs("AllowEditInvoiceOfReturn").value = 1
    ElseIf Me.Chkbarcode(128).value = vbUnchecked Then
        rs("AllowEditInvoiceOfReturn").value = 0
    End If
    
    
       If Me.Chkbarcode(129).value = vbChecked Then
        rs("ProvisionsByManagement").value = 1
    ElseIf Me.Chkbarcode(129).value = vbUnchecked Then
        rs("ProvisionsByManagement").value = 0
    End If
      
      
             If Me.Chkbarcode(165).value = vbChecked Then
        rs("ProvisionsByıEQuipments").value = 1
    ElseIf Me.Chkbarcode(165).value = vbUnchecked Then
        rs("ProvisionsByıEQuipments").value = 0
    End If
    
      
     If Me.Chkbarcode(166).value = vbChecked Then
        rs("ReturnSAlesByBarcode").value = 1
    ElseIf Me.Chkbarcode(166).value = vbUnchecked Then
        rs("ReturnSAlesByBarcode").value = 0
    End If
    
    
         If Me.Chkbarcode(167).value = vbChecked Then
        rs("DontDistributeLegalACC").value = 1
    ElseIf Me.Chkbarcode(167).value = vbUnchecked Then
        rs("DontDistributeLegalACC").value = 0
    End If
    
    
    If Me.Chkbarcode(168).value = vbChecked Then
        rs("CreatePayOrderSales").value = 1
    ElseIf Me.Chkbarcode(168).value = vbUnchecked Then
        rs("CreatePayOrderSales").value = 0
    End If
    
    
    If Me.Chkbarcode(169).value = vbChecked Then
        rs("IsBarCodeByUnit").value = 1
    ElseIf Me.Chkbarcode(169).value = vbUnchecked Then
        rs("IsBarCodeByUnit").value = 0
    End If
    
    
            If Me.Chkbarcode(170).value = vbChecked Then
        rs("TripnotUploadExpenses").value = 1
    ElseIf Me.Chkbarcode(170).value = vbUnchecked Then
        rs("TripnotUploadExpenses").value = 0
    End If
      
      
            If Me.Chkbarcode(171).value = vbChecked Then
        rs("ExpensesByQtyOnly").value = 1
    ElseIf Me.Chkbarcode(171).value = vbUnchecked Then
        rs("ExpensesByQtyOnly").value = 0
    End If
            
            If Me.Chkbarcode(172).value = vbChecked Then
        rs("ShowPrinterDialoge").value = 1
    ElseIf Me.Chkbarcode(172).value = vbUnchecked Then
        rs("ShowPrinterDialoge").value = 0
    End If
    
    
            If Me.Chkbarcode(173).value = vbChecked Then
        rs("AllowDynamicAutoSus").value = 1
    ElseIf Me.Chkbarcode(173).value = vbUnchecked Then
        rs("AllowDynamicAutoSus").value = 0
    End If
    
                If Me.Chkbarcode(174).value = vbChecked Then
        rs("AllowUnbalncedByBranchAccount").value = 1
    ElseIf Me.Chkbarcode(174).value = vbUnchecked Then
        rs("AllowUnbalncedByBranchAccount").value = 0
    End If
    
            
    If Me.Chkbarcode(175).value = vbChecked Then
        rs("SortInvoiceByEntry").value = 1
    ElseIf Me.Chkbarcode(175).value = vbUnchecked Then
        rs("SortInvoiceByEntry").value = 0
    End If
    
            
    
    If Me.Chkbarcode(176).value = vbChecked Then
        rs("CostProductOrderByOut").value = 1
    ElseIf Me.Chkbarcode(176).value = vbUnchecked Then
        rs("CostProductOrderByOut").value = 0
    End If
                
    If Me.Chkbarcode(179).value = vbChecked Then
        rs("CostByProduction").value = 1
    ElseIf Me.Chkbarcode(179).value = vbUnchecked Then
        rs("CostByProduction").value = 0
    End If
    If Me.Chkbarcode(180).value = vbChecked Then
        rs("MaintOrderCantRepeatSales").value = 1
    ElseIf Me.Chkbarcode(180).value = vbUnchecked Then
        rs("MaintOrderCantRepeatSales").value = 0
    End If
    If Me.Chkbarcode(181).value = vbChecked Then
        rs("MaintOrderCantRepeatBillBuy").value = 1
    ElseIf Me.Chkbarcode(181).value = vbUnchecked Then
        rs("MaintOrderCantRepeatBillBuy").value = 0
    End If
                
                
                 If Me.Chkbarcode(184).value = vbChecked Then
        rs("TripRevenueAuto").value = 1
    ElseIf Me.Chkbarcode(184).value = vbUnchecked Then
        rs("TripRevenueAuto").value = 0
    End If
                   
                   
    If Me.Chkbarcode(185).value = vbChecked Then
        rs("IsByNewCoding").value = 1
    ElseIf Me.Chkbarcode(185).value = vbUnchecked Then
        rs("IsByNewCoding").value = 0
    End If
                   
                   
                   

    If Me.Chkbarcode(182).value = vbChecked Then
        rs("PaymentMethLaterCompItem").value = 1
    ElseIf Me.Chkbarcode(182).value = vbUnchecked Then
        rs("PaymentMethLaterCompItem").value = 0
    End If
                

    If Me.Chkbarcode(183).value = vbChecked Then
        rs("ShowBalanceCustInv").value = 1
    ElseIf Me.Chkbarcode(183).value = vbUnchecked Then
        rs("ShowBalanceCustInv").value = 0
    End If
                

                
    If Me.Chkbarcode(177).value = vbChecked Then
        rs("TransferNotInvItemDef").value = 1
    ElseIf Me.Chkbarcode(177).value = vbUnchecked Then
        rs("TransferNotInvItemDef").value = 0
    End If
         
    
    If Me.Chkbarcode(178).value = vbChecked Then
        rs("CustMobNoMandatory").value = 1
    ElseIf Me.Chkbarcode(178).value = vbUnchecked Then
        rs("CustMobNoMandatory").value = 0
    End If
                            
                            
    If Me.Chkbarcode(214).value = vbChecked Then
        rs("CustVatNoMandatory").value = 1
    ElseIf Me.Chkbarcode(214).value = vbUnchecked Then
        rs("CustVatNoMandatory").value = 0
    End If
                            
            
  If Me.Chkbarcode(215).value = vbChecked Then
        rs("AllowScInterface2").value = 1
    ElseIf Me.Chkbarcode(215).value = vbUnchecked Then
        rs("AllowScInterface2").value = 0
    End If
             
            
             If Me.Chkbarcode(130).value = vbChecked Then
        rs("CloseMovingVchrinSales").value = 1
    ElseIf Me.Chkbarcode(130).value = vbUnchecked Then
        rs("CloseMovingVchrinSales").value = 0
    End If
      
               If Me.Chkbarcode(132).value = vbChecked Then
        rs("IsMultiItemsInCompItem").value = 1
    ElseIf Me.Chkbarcode(132).value = vbUnchecked Then
        rs("IsMultiItemsInCompItem").value = 0
    End If
      
          
      
             If Me.Chkbarcode(131).value = vbChecked Then
        rs("CantChangeSalesPerson").value = 1
    ElseIf Me.Chkbarcode(131).value = vbUnchecked Then
        rs("CantChangeSalesPerson").value = 0
    End If
            
            
      
    If Me.Chkbarcode(133).value = vbChecked Then
        rs("BatchCreateManyworkOrder").value = 1
    ElseIf Me.Chkbarcode(133).value = vbUnchecked Then
        rs("BatchCreateManyworkOrder").value = 0
    End If
    If Me.Chkbarcode(134).value = vbChecked Then
        rs("LinkSupplerWithItem").value = 1
    ElseIf Me.Chkbarcode(134).value = vbUnchecked Then
        rs("LinkSupplerWithItem").value = 0
    End If
                        
      If Me.Chkbarcode(135).value = vbChecked Then
        rs("ShowOnlyItemsOfSales").value = 1
    ElseIf Me.Chkbarcode(135).value = vbUnchecked Then
        rs("ShowOnlyItemsOfSales").value = 0
    End If
                        
      
      If Me.Chkbarcode(136).value = vbChecked Then
        rs("PrintInvoiceByBranch").value = 1
    ElseIf Me.Chkbarcode(136).value = vbUnchecked Then
        rs("PrintInvoiceByBranch").value = 0
    End If
                              
             If Me.Chkbarcode(140).value = vbChecked Then
        rs("GeneralVoucherCreateSalesGE").value = 1
    ElseIf Me.Chkbarcode(140).value = vbUnchecked Then
        rs("GeneralVoucherCreateSalesGE").value = 0
    End If
                                           
             If Me.Chkbarcode(141).value = vbChecked Then
        rs("SalesNotCreateGe").value = 1
    ElseIf Me.Chkbarcode(141).value = vbUnchecked Then
        rs("SalesNotCreateGe").value = 0
    End If
                              
    If Me.Chkbarcode(77).value = vbChecked Then
        rs("CanChanegeLinkedPurcahsenvoice").value = 1
    ElseIf Me.Chkbarcode(77).value = vbUnchecked Then
        rs("CanChanegeLinkedPurcahsenvoice").value = 0
    End If
       If Me.Chkbarcode(78).value = vbChecked Then
        rs("AllowProductOrderOne").value = 1
    ElseIf Me.Chkbarcode(78).value = vbUnchecked Then
        rs("AllowProductOrderOne").value = 0
    End If
     If Me.Chkbarcode(79).value = vbChecked Then
        rs("SalaryJLByManagement").value = 1
    ElseIf Me.Chkbarcode(79).value = vbUnchecked Then
        rs("SalaryJLByManagement").value = 0
    End If
    If Me.Chkbarcode(80).value = vbChecked Then
        rs("AllowGoodPerfAccount").value = 1
    ElseIf Me.Chkbarcode(80).value = vbUnchecked Then
        rs("AllowGoodPerfAccount").value = 0
    End If
    If Me.Chkbarcode(83).value = vbChecked Then
        rs("AllowAnalyticJL").value = 1
    ElseIf Me.Chkbarcode(83).value = vbUnchecked Then
        rs("AllowAnalyticJL").value = 0
    End If
    If Me.Chkbarcode(84).value = vbChecked Then
        rs("AllowSaveTripWithoutExpen").value = 1
    ElseIf Me.Chkbarcode(84).value = vbUnchecked Then
        rs("AllowSaveTripWithoutExpen").value = 0
    End If
    
    
    
    If Me.Chkbarcode(110).value = vbChecked Then
        rs("CreateEntryManual").value = 1
    ElseIf Me.Chkbarcode(110).value = vbUnchecked Then
        rs("CreateEntryManual").value = 0
    End If
    
    
    If Me.Chkbarcode(208).value = vbChecked Then
        rs("CustCreat4Acc").value = 1
    ElseIf Me.Chkbarcode(208).value = vbUnchecked Then
        rs("CustCreat4Acc").value = 0
    End If
   
       
    If Me.Chkbarcode(209).value = vbChecked Then
        rs("SuppCreat4Acc").value = 1
    ElseIf Me.Chkbarcode(209).value = vbUnchecked Then
        rs("SuppCreat4Acc").value = 0
    End If
   
   
     
    
   
    If Me.Chkbarcode(111).value = vbChecked Then
        rs("chkAllowEditPaymentCont").value = 1
    ElseIf Me.Chkbarcode(111).value = vbUnchecked Then
        rs("chkAllowEditPaymentCont").value = 0
    End If
    
    
           
    If Me.Chkbarcode(210).value = vbChecked Then
        rs("CreateEntryBillItems").value = 1
    ElseIf Me.Chkbarcode(210).value = vbUnchecked Then
        rs("CreateEntryBillItems").value = 0
    End If
   
     
    
    
    If Me.Chkbarcode(100).value = vbChecked Then
        rs("SAVEMAINTENANCEJOBWITHORDERORPLANONLY").value = 1
    ElseIf Me.Chkbarcode(100).value = vbUnchecked Then
        rs("SAVEMAINTENANCEJOBWITHORDERORPLANONLY").value = 0
    End If
    
    
    
    If Me.Chkbarcode(85).value = vbChecked Then
        rs("SendToAprovedSalesBill").value = 1
    ElseIf Me.Chkbarcode(85).value = vbUnchecked Then
        rs("SendToAprovedSalesBill").value = 0
    End If
     If Me.Chkbarcode(86).value = vbChecked Then
        rs("SalaryJLByAnalyEqup").value = 1
    ElseIf Me.Chkbarcode(86).value = vbUnchecked Then
        rs("SalaryJLByAnalyEqup").value = 0
    End If
    If Me.Chkbarcode(81).value = vbChecked Then
        rs("ManualSalesInvoiceMust").value = 1
    ElseIf Me.Chkbarcode(81).value = vbUnchecked Then
        rs("ManualSalesInvoiceMust").value = 0
    End If
    If Me.Chkbarcode(82).value = vbChecked Then
        rs("AllItemInVAT").value = 1
    ElseIf Me.Chkbarcode(82).value = vbUnchecked Then
        rs("AllItemInVAT").value = 0
    End If
    
       If Me.Chkbarcode(42).value = vbChecked Then
        rs("AnalyticPaymentVouchr").value = 1
    ElseIf Me.Chkbarcode(42).value = vbUnchecked Then
        rs("AnalyticPaymentVouchr").value = 0
    End If
    
    
       If Me.Chkbarcode(43).value = vbChecked Then
        rs("ShowDriverOnly").value = 1
    ElseIf Me.Chkbarcode(43).value = vbUnchecked Then
        rs("ShowDriverOnly").value = 0
    End If
    
    
        If Me.Chkbarcode(44).value = vbChecked Then
        rs("AllowSalesMultyPayed").value = 1
    ElseIf Me.Chkbarcode(44).value = vbUnchecked Then
        rs("AllowSalesMultyPayed").value = 0
    End If
    If Me.Chkbarcode(59).value = vbChecked Then
        rs("AllowAccountMultyPayed").value = 1
    ElseIf Me.Chkbarcode(59).value = vbUnchecked Then
        rs("AllowAccountMultyPayed").value = 0
    End If
    
         If Me.Chkbarcode(50).value = vbChecked Then
        rs("AllowPurchasesMultyPayed").value = 1
    ElseIf Me.Chkbarcode(50).value = vbUnchecked Then
        rs("AllowPurchasesMultyPayed").value = 0
    End If
    
      If Me.Chkbarcode(45).value = vbChecked Then
        rs("CashCustomerNameMustenter").value = 1
    ElseIf Me.Chkbarcode(45).value = vbUnchecked Then
        rs("CashCustomerNameMustenter").value = 0
    End If
    
    If Me.Chkbarcode(48).value = vbChecked Then
        rs("AllowCommtionJEFromValueVisa").value = 1
    ElseIf Me.Chkbarcode(48).value = vbUnchecked Then
        rs("AllowCommtionJEFromValueVisa").value = 0
    End If
    
    
    If Me.Chkbarcode(49).value = vbChecked Then
        rs("AllowWorkWithArea").value = 1
    ElseIf Me.Chkbarcode(49).value = vbUnchecked Then
        rs("AllowWorkWithArea").value = 0
    End If
       If Me.Chkbarcode(51).value = vbChecked Then
        rs("AllowAcceleratepayment").value = 1
    ElseIf Me.Chkbarcode(51).value = vbUnchecked Then
        rs("AllowAcceleratepayment").value = 0
    End If
      If Me.Chkbarcode(52).value = vbChecked Then
        rs("AllowExperDateFIFO").value = 1
    ElseIf Me.Chkbarcode(52).value = vbUnchecked Then
        rs("AllowExperDateFIFO").value = 0
    End If
      If Me.Chkbarcode(53).value = vbChecked Then
        rs("AllowProjectBill2Serial").value = 1
    ElseIf Me.Chkbarcode(53).value = vbUnchecked Then
        rs("AllowProjectBill2Serial").value = 0
    End If
    If Me.Chkbarcode(54).value = vbChecked Then
        rs("ViewAccountsbyBranch").value = 1
    ElseIf Me.Chkbarcode(54).value = vbUnchecked Then
        rs("ViewAccountsbyBranch").value = 0
    End If
     If Me.Chkbarcode(55).value = vbChecked Then
        rs("AllowEditeAccounts").value = 1
    ElseIf Me.Chkbarcode(55).value = vbUnchecked Then
        rs("AllowEditeAccounts").value = 0
    End If
    If Me.Chkbarcode(56).value = vbChecked Then
        rs("ProjectUnderImplemen").value = 1
    ElseIf Me.Chkbarcode(56).value = vbUnchecked Then
        rs("ProjectUnderImplemen").value = 0
    End If
        If Me.Chkbarcode(57).value = vbChecked Then
        rs("AllowHideAssest").value = 1
    ElseIf Me.Chkbarcode(57).value = vbUnchecked Then
        rs("AllowHideAssest").value = 0
    End If
     If Me.Chkbarcode(58).value = vbChecked Then
        rs("LockSalary").value = 1
    ElseIf Me.Chkbarcode(58).value = vbUnchecked Then
        rs("LockSalary").value = 0
    End If
            If Me.Chkbarcode(4).value = vbChecked Then
        rs("updatecashvchrifdeposite").value = 1
    ElseIf Me.Chkbarcode(4).value = vbUnchecked Then
        rs("updatecashvchrifdeposite").value = 0
    End If
    
    
    If Me.Chkbarcode(5).value = vbChecked Then
        rs("Revenueowed").value = 1
    ElseIf Me.Chkbarcode(5).value = vbUnchecked Then
        rs("Revenueowed").value = 0
    End If
    
    
        If Me.Chkbarcode(6).value = vbChecked Then
        rs("AllowupdateJobStatus").value = 1
    ElseIf Me.Chkbarcode(6).value = vbUnchecked Then
        rs("AllowupdateJobStatus").value = 0
    End If
    
    
    
    
    If Me.Chkbarcode(60).value = vbChecked Then
         rs("OpeningEmployeeShowAll").value = 1
    ElseIf Me.Chkbarcode(60).value = vbUnchecked Then
         rs("OpeningEmployeeShowAll").value = 0
    End If
        
    If Me.Chkbarcode(61).value = vbChecked Then
        rs("SellOrderBalance").value = 1
    ElseIf Me.Chkbarcode(61).value = vbUnchecked Then
        rs("SellOrderBalance").value = 0
    End If
    If Me.Chkbarcode(62).value = vbChecked Then
        rs("EndServiceMore5Year").value = 1
    ElseIf Me.Chkbarcode(62).value = vbUnchecked Then
        rs("EndServiceMore5Year").value = 0
    End If
    
    
    
        If Me.Chkbarcode(75).value = vbChecked Then
        rs("VacstionShowOldSalaries").value = 1
    ElseIf Me.Chkbarcode(75).value = vbUnchecked Then
        rs("VacstionShowOldSalaries").value = 0
    End If
    
            If Me.Chkbarcode(76).value = vbChecked Then
        rs("AllowReturnWithoutCost").value = 1
    ElseIf Me.Chkbarcode(76).value = vbUnchecked Then
        rs("AllowReturnWithoutCost").value = 0
    End If
    
    If Me.Chkbarcode(63).value = vbChecked Then
        rs("ShowItemByCustomer").value = 1
    ElseIf Me.Chkbarcode(63).value = vbUnchecked Then
        rs("ShowItemByCustomer").value = 0
    End If
     If Me.Chkbarcode(64).value = vbChecked Then
        rs("RawMaterMix").value = 1
    ElseIf Me.Chkbarcode(64).value = vbUnchecked Then
        rs("RawMaterMix").value = 0
    End If
    
     If Me.Chkbarcode(142).value = vbChecked Then
        rs("RawMaterMix2").value = 1
    ElseIf Me.Chkbarcode(142).value = vbUnchecked Then
        rs("RawMaterMix2").value = 0
    End If
    
     If Me.Chkbarcode(143).value = vbChecked Then
        rs("DontCreateOut").value = 1
    ElseIf Me.Chkbarcode(143).value = vbUnchecked Then
        rs("DontCreateOut").value = 0
    End If
        

     If Me.Chkbarcode(144).value = vbChecked Then
        rs("DontCreateOut2").value = 1
    ElseIf Me.Chkbarcode(144).value = vbUnchecked Then
        rs("DontCreateOut2").value = 0
    End If
                
                
        
     If Me.Chkbarcode(145).value = vbChecked Then
        rs("InsertItemManualOut").value = 1
    ElseIf Me.Chkbarcode(145).value = vbUnchecked Then
        rs("InsertItemManualOut").value = 0
    End If
                
    
    If Me.Chkbarcode(65).value = vbChecked Then
        rs("LinkUsersWithPayment").value = 1
    ElseIf Me.Chkbarcode(65).value = vbUnchecked Then
        rs("LinkUsersWithPayment").value = 0
    End If
    If Me.Chkbarcode(66).value = vbChecked Then
        rs("VATNoAccordActivity").value = 1
    ElseIf Me.Chkbarcode(66).value = vbUnchecked Then
        rs("VATNoAccordActivity").value = 0
    End If
       If Me.Chkbarcode(67).value = vbChecked Then
        rs("NotCrtResvVouchProjects").value = 1
    ElseIf Me.Chkbarcode(67).value = vbUnchecked Then
        rs("NotCrtResvVouchProjects").value = 0
    End If
    If Me.Chkbarcode(68).value = vbChecked Then
        rs("SalesTrustsAffectVedorCode").value = 1
    ElseIf Me.Chkbarcode(68).value = vbUnchecked Then
        rs("SalesTrustsAffectVedorCode").value = 0
    End If
    If Me.Chkbarcode(69).value = vbChecked Then
        rs("AllowNoRoudProjectInvoices").value = 1
    ElseIf Me.Chkbarcode(69).value = vbUnchecked Then
        rs("AllowNoRoudProjectInvoices").value = 0
    End If
    If Me.Chkbarcode(70).value = vbChecked Then
        rs("ProductionRawMaterMix").value = 1
    ElseIf Me.Chkbarcode(70).value = vbUnchecked Then
        rs("ProductionRawMaterMix").value = 0
    End If
    If Me.Chkbarcode(71).value = vbChecked Then
        rs("AllowLastPrice").value = 1
    ElseIf Me.Chkbarcode(71).value = vbUnchecked Then
        rs("AllowLastPrice").value = 0
    End If
    If Me.Chkbarcode(72).value = vbChecked Then
        rs("AllowItemByRow").value = 1
    ElseIf Me.Chkbarcode(72).value = vbUnchecked Then
        rs("AllowItemByRow").value = 0
    End If
     If Me.Chkbarcode(73).value = vbChecked Then
        rs("AllowChangManualQtyMix").value = 1
    ElseIf Me.Chkbarcode(73).value = vbUnchecked Then
        rs("AllowChangManualQtyMix").value = 0
    End If
     If Me.Chkbarcode(74).value = vbChecked Then
        rs("AccountAccordingCash").value = 1
    ElseIf Me.Chkbarcode(74).value = vbUnchecked Then
        rs("AccountAccordingCash").value = 0
    End If
     If Me.Chkbarcode(36).value = vbChecked Then
        rs("AllowTowShift").value = 1
    ElseIf Me.Chkbarcode(36).value = vbUnchecked Then
        rs("AllowTowShift").value = 0
    End If
    If Me.Chkbarcode(37).value = vbChecked Then
        rs("AllowItemsShortName").value = 1
    ElseIf Me.Chkbarcode(37).value = vbUnchecked Then
        rs("AllowItemsShortName").value = 0
    End If
    
    

    If Me.ChkCostStarting.value = vbChecked Then
        rs("CostStarting").value = 1
    ElseIf Me.ChkCostStarting.value = vbUnchecked Then
        rs("CostStarting").value = 0
    End If
    

 If Me.chkuserCode.value = vbChecked Then
        rs("chkuserCode").value = 1
    ElseIf Me.chkuserCode.value = vbUnchecked Then
        rs("chkuserCode").value = 0
    End If
    
    
     If Me.ChkItemsattachedzero.value = vbChecked Then
        rs("Itemsattachedzero").value = 1
    ElseIf Me.ChkItemsattachedzero.value = vbUnchecked Then
        rs("Itemsattachedzero").value = 0
    End If
    

    If Me.ChKautoIssueVoucher.value = vbChecked Then
        rs("autoIssueVoucher").value = 1
    ElseIf Me.ChKautoIssueVoucher.value = vbUnchecked Then
        rs("autoIssueVoucher").value = 0
    End If

    If Me.chkMonthIs30days.value = vbChecked Then
        rs("MonthIs30days").value = 1
    ElseIf Me.chkMonthIs30days.value = vbUnchecked Then
        rs("MonthIs30days").value = 0
    End If

   
    rs.update
chkSaveRR





FllowCmdOk

 rs.update
cheks
 
     Exit Sub
ErrTrap:
    Msg = "ÕœÀ  „‘ﬂ·… ≈À‰«¡ Õ›Ÿ «·≈⁄œ«œ« ...!!!!"
 
 
  
    MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title

End Sub

Sub chkSaveRR()
    If ChkAsk.value = Checked Then
        SaveSetting StrAppRegPath, "View_Type", "ShowMe", False
    Else
        SaveSetting StrAppRegPath, "View_Type", "ShowMe", True
    End If

  '  If ChkTax.value = Checked Then
  '      SaveSetting StrAppRegPath, "SallBill", "HaveTaxOnSalles", True
  '  Else
  '      SaveSetting StrAppRegPath, "SallBill", "HaveTaxOnSalles", False
  '  End If

    If ChkDelayVal.value = Checked Then
        SaveSetting StrAppRegPath, "View_Type", "ShowPayment", True
    
        If Combo1.ListIndex <> -1 Then
           
            If Combo1.ListIndex = 0 Then
                SaveSetting StrAppRegPath, "Setting", "INTERVAL_ShowPayment", "D"
            ElseIf Combo1.ListIndex = 1 Then
                SaveSetting StrAppRegPath, "Setting", "INTERVAL_ShowPayment", "m"
            ElseIf Combo1.ListIndex = 2 Then
                SaveSetting StrAppRegPath, "Setting", "INTERVAL_ShowPayment", "yyyy"
            End If
        
            If IsNumeric(Text1.text) Then
                SaveSetting StrAppRegPath, "Setting", "Count_ShowPayment", val(Text1.text)
            Else
                SaveSetting StrAppRegPath, "Setting", "Count_ShowPayment", 0
            End If
        
        End If

    Else
        SaveSetting StrAppRegPath, "View_Type", "ShowPayment", False
    End If

    If ChkShow.value = Checked Then
        SaveSetting StrAppRegPath, "View_Type", "ShowRequest", True
    Else
        SaveSetting StrAppRegPath, "View_Type", "ShowRequest", False
    End If

    If ChKHR.value = Checked Then
        SaveSetting StrAppRegPath, "Setting", "Showhr", True
    Else
        SaveSetting StrAppRegPath, "Setting", "Showhr", False
    End If
 
    If ChKProjectsAlarm1.value = Checked Then
        SaveSetting StrAppRegPath, "View_Type", "ShowProjectsAlarm1", True
    Else
        SaveSetting StrAppRegPath, "View_Type", "ShowProjectsAlarm1", False
    End If

    If ChKProjectsAlarm2.value = Checked Then
        SaveSetting StrAppRegPath, "View_Type", "ShowProjectsAlarm2", True
    Else
        SaveSetting StrAppRegPath, "View_Type", "ShowProjectsAlarm2", False
    End If

    If ChkInstallmentMustPayed.value = Checked Then
        SaveSetting StrAppRegPath, "View_Type", "InstallmentMustPayed", True
        
        If Combo2.ListIndex <> -1 Then
           
            If Combo2.ListIndex = 0 Then
                SaveSetting StrAppRegPath, "Setting", "INTERVAL_InstallmentMustPayed", "D"
            ElseIf Combo2.ListIndex = 1 Then
                SaveSetting StrAppRegPath, "Setting", "INTERVAL_InstallmentMustPayed", "M"
            ElseIf Combo2.ListIndex = 2 Then
                SaveSetting StrAppRegPath, "Setting", "INTERVAL_InstallmentMustPayed", "yyyy"
            End If
        
            If IsNumeric(Text2.text) Then
                SaveSetting StrAppRegPath, "Setting", "Count_InstallmentMustPayed", val(Text2.text)
            Else
                SaveSetting StrAppRegPath, "Setting", "Count_InstallmentMustPayed", 0
            End If
        
        End If

    Else
        SaveSetting StrAppRegPath, "View_Type", "InstallmentMustPayed", False
    End If
End Sub
Sub FllowCmdOk()


    If Me.Chkbarcode(225).value = vbChecked Then
        rs("ShowPrinterDialoge2").value = 1
    ElseIf Me.Chkbarcode(225).value = vbUnchecked Then
        rs("ShowPrinterDialoge2").value = 0
    End If
    


   If Me.Chkbarcode(221).value = vbChecked Then
        rs("ZacatHandW").value = 1
    ElseIf Me.Chkbarcode(221).value = vbUnchecked Then
        rs("ZacatHandW").value = 0
    End If
    

     If Me.Chkbarcode(220).value = vbChecked Then
        rs("DiscountByQtyOnly").value = 1
    ElseIf Me.Chkbarcode(220).value = vbUnchecked Then
        rs("DiscountByQtyOnly").value = 0
    End If
    
    If Me.Chkbarcode(222).value = vbChecked Then
        rs("IsTransferByCode").value = 1
    ElseIf Me.Chkbarcode(220).value = vbUnchecked Then
        rs("IsTransferByCode").value = 0
    End If
    

     If Me.ChKautoReseiveVoucher.value = vbChecked Then
        rs("autoReseiveVoucher").value = 1
    ElseIf Me.ChKautoReseiveVoucher.value = vbUnchecked Then
        rs("autoReseiveVoucher").value = 0
    End If
    If Me.ChkitemsWorkWithColor.value = vbChecked Then
        rs("itemsWorkWithColor").value = 1
    ElseIf Me.ChkitemsWorkWithColor.value = vbUnchecked Then
        rs("itemsWorkWithColor").value = 0
    End If
    If Me.ChkitemsWorkWithDates.value = vbChecked Then
        rs("itemsWorkWithDates").value = 1
    ElseIf Me.ChkitemsWorkWithDates.value = vbUnchecked Then
        rs("itemsWorkWithDates").value = 0
    End If

    If Me.ChkitemsWorkWithClass.value = vbChecked Then
        rs("itemsWorkWithClass").value = 1
    ElseIf Me.ChkitemsWorkWithClass.value = vbUnchecked Then
        rs("itemsWorkWithClass").value = 0
    End If
                
                
                
                
    rs("Commonname").value = XPTxtComment(8)
    rs("SerialNumber").value = XPTxtComment(7)
    rs("OrganizationName").value = XPTxtComment(6)
        
            rs("Invoicetype").value = Me.Invoicetype.ListIndex
            rs("DefaultInvoicetype").value = Me.DefaultInvoicetype.ListIndex
            rs("SendingMode").value = Me.SendingMode.ListIndex
            
             
   
    rs("industrey").value = XPTxtComment(11)
    rs("CSR").value = XPTxtComment(9)
    rs("Privatekey").value = XPTxtComment(13)
    rs("PublickeycertPem").value = XPTxtComment(14)
    rs("SecretKey").value = XPTxtComment(15)
 savepart2
 rs.update
    '%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
    If ChkExpireLicense.value = Checked Then
        SaveSetting StrAppRegPath, "View_Type", "ExpireLicense", True
        
        If Combo7.ListIndex <> -1 Then
            If Combo7.ListIndex = 0 Then
                SaveSetting StrAppRegPath, "Setting", "INTERVAL_ExpireLicense", "D"
            ElseIf Combo7.ListIndex = 1 Then
                SaveSetting StrAppRegPath, "Setting", "INTERVAL_ExpireLicense", "m"
            ElseIf Combo7.ListIndex = 2 Then
                SaveSetting StrAppRegPath, "Setting", "INTERVAL_ExpireLicense", "yyyy"
            End If
        
            If IsNumeric(Text11.text) Then
                SaveSetting StrAppRegPath, "Setting", "Count_ExpireLicense", val(Text11.text)
            Else
                SaveSetting StrAppRegPath, "Setting", "Count_ExpireLicense", 0
            End If
        
        End If

    Else
        SaveSetting StrAppRegPath, "View_Type", "ExpireLicense", False:          SaveSetting StrAppRegPath, "Setting", "Count_ExpireLicense", 0
    End If

    
    If ChkExpireInsurance.value = Checked Then
        SaveSetting StrAppRegPath, "View_Type", "ExpireInsurance", True
        
        If Combo8.ListIndex <> -1 Then
           
            If Combo8.ListIndex = 0 Then
                SaveSetting StrAppRegPath, "Setting", "INTERVAL_ExpireInsurance", "D"
            ElseIf Combo8.ListIndex = 1 Then
                SaveSetting StrAppRegPath, "Setting", "INTERVAL_ExpireInsurance", "m"
            ElseIf Combo8.ListIndex = 2 Then
                SaveSetting StrAppRegPath, "Setting", "INTERVAL_ExpireInsurance", "yyyy"
            End If
        
            If IsNumeric(Text12.text) Then
                SaveSetting StrAppRegPath, "Setting", "Count_ExpireInsurance", val(Text12.text)
            Else
                SaveSetting StrAppRegPath, "Setting", "Count_ExpireInsurance", 0
            End If
        
        End If

    Else
        SaveSetting StrAppRegPath, "View_Type", "ExpireInsurance", False
        SaveSetting StrAppRegPath, "Setting", "Count_ExpireInsurance", 0
    End If
End Sub
Private Sub cheks()
    '%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
    


    If chkRentInstallments.value = Checked Then
        SaveSetting StrAppRegPath, "View_Type", "RentInstallments", True
        
        If Combo10.ListIndex <> -1 Then
           
            If Combo10.ListIndex = 0 Then
                SaveSetting StrAppRegPath, "Setting", "INTERVAL_RentInstallments", "D"
            ElseIf Combo10.ListIndex = 1 Then
                SaveSetting StrAppRegPath, "Setting", "INTERVAL_RentInstallments", "M"
            ElseIf Combo10.ListIndex = 2 Then
                SaveSetting StrAppRegPath, "Setting", "INTERVAL_RentInstallments", "yyyy"
            End If
        
            If IsNumeric(Text2.text) Then
                SaveSetting StrAppRegPath, "Setting", "Count_RentInstallments", val(Text8.text)
            Else
                SaveSetting StrAppRegPath, "Setting", "Count_RentInstallments", 0
            End If
        
        End If

    Else
        SaveSetting StrAppRegPath, "View_Type", "RentInstallments", False
    End If
    

    If ChkExpireEkama.value = Checked Then
        SaveSetting StrAppRegPath, "View_Type", "ExpireEkama", True
        
        If Combo3.ListIndex <> -1 Then
           
                        If Combo3.ListIndex = 0 Then
                            SaveSetting StrAppRegPath, "Setting", "INTERVAL_ExpireEkama", "D"
                        ElseIf Combo3.ListIndex = 1 Then
                            SaveSetting StrAppRegPath, "Setting", "INTERVAL_ExpireEkama", "m"
                        ElseIf Combo3.ListIndex = 2 Then
                            SaveSetting StrAppRegPath, "Setting", "INTERVAL_ExpireEkama", "yyyy"
                        End If
                                    
                            If IsNumeric(Text3.text) Then
                                SaveSetting StrAppRegPath, "Setting", "Count_ExpireEkama", val(Text3.text)
                            Else
                                SaveSetting StrAppRegPath, "Setting", "Count_ExpireEkama", 0
                            End If
        
        End If

    Else
        SaveSetting StrAppRegPath, "View_Type", "ExpireEkama", False
    End If

    
    
    If CheckLC.value = Checked Then
        SaveSetting StrAppRegPath, "View_Type", "LC", True
        
        If Combo14.ListIndex <> -1 Then
           
                        If Combo14.ListIndex = 0 Then
                            SaveSetting StrAppRegPath, "Setting", "INTERVAL_LC", "D"
                        ElseIf Combo14.ListIndex = 1 Then
                            SaveSetting StrAppRegPath, "Setting", "INTERVAL_LC", "m"
                        ElseIf Combo14.ListIndex = 2 Then
                            SaveSetting StrAppRegPath, "Setting", "INTERVAL_LC", "yyyy"
                        End If
                                    
                            If IsNumeric(Text7.text) Then
                                SaveSetting StrAppRegPath, "Setting", "Count_LC", val(Text7.text)
                            Else
                                SaveSetting StrAppRegPath, "Setting", "Count_LC", 0
                            End If
        
        End If

    Else
        SaveSetting StrAppRegPath, "View_Type", "LC", False
    End If
'*************
    
    If ChkExpireTest.value = Checked Then
        SaveSetting StrAppRegPath, "View_Type", "ExpireTest", True
        
        If Combo9.ListIndex <> -1 Then
           
            If Combo9.ListIndex = 0 Then
                SaveSetting StrAppRegPath, "Setting", "INTERVAL_ExpireTest", "D"
            ElseIf Combo9.ListIndex = 1 Then
                SaveSetting StrAppRegPath, "Setting", "INTERVAL_ExpireTest", "m"
            ElseIf Combo9.ListIndex = 2 Then
                SaveSetting StrAppRegPath, "Setting", "INTERVAL_ExpireTest", "yyyy"
            End If
        
            If IsNumeric(Text13.text) Then
                SaveSetting StrAppRegPath, "Setting", "Count_ExpireTest", val(Text13.text)
            Else
                SaveSetting StrAppRegPath, "Setting", "Count_ExpireTest", 0
            End If
        
        End If

    Else
        SaveSetting StrAppRegPath, "View_Type", "ExpireTest", False
        SaveSetting StrAppRegPath, "Setting", "Count_ExpireTest", 0
    End If

    If ChkExpireLicence.value = Checked Then
        SaveSetting StrAppRegPath, "View_Type", "ExpireLicence", True
        
        If Combo4.ListIndex <> -1 Then
           
            If Combo4.ListIndex = 0 Then
                SaveSetting StrAppRegPath, "Setting", "INTERVAL_ExpireLicence", "D"
            ElseIf Combo4.ListIndex = 1 Then
                SaveSetting StrAppRegPath, "Setting", "INTERVAL_ExpireLicence", "m"
            ElseIf Combo4.ListIndex = 2 Then
                SaveSetting StrAppRegPath, "Setting", "INTERVAL_ExpireLicence", "yyyy"
            End If
        
            If IsNumeric(Text4.text) Then
                SaveSetting StrAppRegPath, "Setting", "Count_ExpireLicence", val(Text4.text)
            Else
                SaveSetting StrAppRegPath, "Setting", "Count_ExpireLicence", 0
            End If
        
        End If

    Else
        SaveSetting StrAppRegPath, "View_Type", "ExpireLicence", False
    End If

    If ChkExpirepas.value = Checked Then
        SaveSetting StrAppRegPath, "View_Type", "Expirepas", True
        
        If Combo5.ListIndex <> -1 Then
           
            If Combo5.ListIndex = 0 Then
                SaveSetting StrAppRegPath, "Setting", "INTERVAL_Expirepas", "D"
            ElseIf Combo5.ListIndex = 1 Then
                SaveSetting StrAppRegPath, "Setting", "INTERVAL_Expirepas", "m"
            ElseIf Combo5.ListIndex = 2 Then
                SaveSetting StrAppRegPath, "Setting", "INTERVAL_Expirepas", "yyyy"
            End If
        
            If IsNumeric(Text5.text) Then
                SaveSetting StrAppRegPath, "Setting", "Count_Expirepas", val(Text5.text)
            Else
                SaveSetting StrAppRegPath, "Setting", "Count_Expirepas", 0
            End If
        
        End If

    Else
        SaveSetting StrAppRegPath, "View_Type", "Expirepas", False
    End If

    If ChkExpirepoket.value = Checked Then
        SaveSetting StrAppRegPath, "View_Type", "Expirepoket", True
        
        If Combo6.ListIndex <> -1 Then
           
            If Combo6.ListIndex = 0 Then
                SaveSetting StrAppRegPath, "Setting", "INTERVAL_Expirepoket", "D"
            ElseIf Combo6.ListIndex = 1 Then
                SaveSetting StrAppRegPath, "Setting", "INTERVAL_Expirepoket", "m"
            ElseIf Combo6.ListIndex = 2 Then
                SaveSetting StrAppRegPath, "Setting", "INTERVAL_Expirepoket", "yyyy"
            End If
        
            If IsNumeric(Text6.text) Then
                SaveSetting StrAppRegPath, "Setting", "Count_Expirepoket", val(Text6.text)
            Else
                SaveSetting StrAppRegPath, "Setting", "Count_Expirepoket", 0
            End If
        
        End If

    Else
        SaveSetting StrAppRegPath, "View_Type", "Expirepoket", False
    End If

    If DBCboClientName.text <> "" Then
        SaveSetting StrAppRegPath, "DefaultOptions", "DefaultClient", DBCboClientName.BoundText
    End If

    If DBCboSupName.text <> "" Then
        SaveSetting StrAppRegPath, "DefaultOptions", "DefaultSup", DBCboSupName.BoundText
    End If

    If DCboStoreName(0).text <> "" Then
        SaveSetting StrAppRegPath, "DefaultOptions", "DefaultSaleStore", DCboStoreName(0).BoundText
    End If


    If DCboStoreName(1).text <> "" Then
        SaveSetting StrAppRegPath, "DefaultOptions", "DefaultPurchaseStore", DCboStoreName(1).BoundText
    End If

    '«· ·„ÌÕ «·ÌÊ„Ì
    If ChkShowToolTip.value = Checked Then
        SaveSetting StrAppRegPath, "View_Type", "ShowToolTip", True
    Else
        SaveSetting StrAppRegPath, "View_Type", "ShowToolTip", False
    End If

    '‘—Ìÿ «·«Œ ’«—« 
    If chkshortCuts.value = Checked Then
        SaveSetting StrAppRegPath, "View_Type", "shortCuts", True
    Else
        SaveSetting StrAppRegPath, "View_Type", "shortCuts", False
    End If

    '  ‘Ã—… «·«’‰«›
    If Chktree.value = Checked Then
        SaveSetting StrAppRegPath, "View_Type", "tree", True
    Else
        SaveSetting StrAppRegPath, "View_Type", "tree", False
    End If

    '    ⁄—÷ «·‰ ÌÃ…
    If ChkCalender.value = Checked Then
        SaveSetting StrAppRegPath, "View_Type", "Calender", True
    Else
        SaveSetting StrAppRegPath, "View_Type", "Calender", False
    End If

    '   ⁄‰œ › Õ «Ì ‘«‘…  › Õ ÃœÌœ «·Ì«
    If CHECK_OPEN_NEW_SCREEN.value = Checked Then
        SaveSetting StrAppRegPath, "View_Type", "OPEN_NEW_SCREEN", True
    Else
        SaveSetting StrAppRegPath, "View_Type", "OPEN_NEW_SCREEN", False
    End If

    '⁄—÷ «⁄„«— «·œÌÊ‰ ›Ì ﬂ‘› «·Õ”«»
    If ChkViewAging.value = Checked Then
        SaveSetting StrAppRegPath, "View_Type", "ViewAging", True
    Else
        SaveSetting StrAppRegPath, "View_Type", "ViewAging", False
    End If

    '    › Õ «·—”«∆·
    If ChkMessnger.value = Checked Then
        SaveSetting StrAppRegPath, "View_Type", "Messnger", True
    Else
        SaveSetting StrAppRegPath, "View_Type", "Messnger", False
    End If

    '    ⁄—÷ «”„ «·›—⁄  »Ã«‰» «·Õ”«» ›Ì «·ﬁÌœ
    If ChkPrintBranchINGE.value = Checked Then
        SaveSetting StrAppRegPath, "View_Type", "PrintBranchINGE", True
    Else
        SaveSetting StrAppRegPath, "View_Type", "PrintBranchINGE", False
    End If
    If ChkPrintCCinGE.value = Checked Then
        SaveSetting StrAppRegPath, "View_Type", "PrintCCinGE", True
    Else
        SaveSetting StrAppRegPath, "View_Type", "PrintCCinGE", False
    End If

    ' ⁄—÷  «·—”„ «·»Ì«‰Ì ›Ì ﬂ‘› «·Õ”«»
    If ChkChartPrintinAS.value = Checked Then SaveSetting StrAppRegPath, "View_Type", "ChartPrintinAS", True Else:                    SaveSetting StrAppRegPath, "View_Type", "ChartPrintinAS", False
   

    '«Œ›«¡ ﬂ· «· ‰»ÌÂ« 
    If ChkHideAllAlarms.value = Checked Then
        SaveSetting StrAppRegPath, "View_Type", "HideAllAlarms", True
    Else
        SaveSetting StrAppRegPath, "View_Type", "HideAllAlarms", False
    End If

   If IsNumeric(Text14.text) Then
                SaveSetting StrAppRegPath, "Setting", "CountAlarmMinutes", val(Text14.text)
            Else
                SaveSetting StrAppRegPath, "Setting", "CountAlarmMinutes", 0
            End If
    LoadMainSystemOptions
    CuurentLogdata
    If SystemOptions.UserInterface = ArabicInterface Then MsgBox " „ «·Õ›Ÿ ", vbInformation Else:              MsgBox "Saved Successfully   ", vbInformation

End Sub
Private Sub Command1_Click(Index As Integer)
Select Case Index
Case 0
frmInfoSettings.show
Case 1
FrmMessageTempltes.show
Case 2
frmInfoSettings1.show
End Select
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, _
                         Shift As Integer)
    On Error GoTo ErrTrap

    If Shift = 2 Then
        tb.SetFocus

        If KeyCode = vbKeyTab Then
            If tb.CurrTab < 5 Then
                tb.CurrTab = tb.CurrTab + 1
            Else
              NewTab = 0
            End If
        End If
    End If

    Exit Sub
ErrTrap:
End Sub

Private Sub ChangeLang()
Chkbarcode(100).Caption = "Job with order or plan only"
Chkbarcode(110).Caption = "Do not create the entry automatically in the contracts"
Chkbarcode(111).Caption = "Allow adjustment in payments in contracts"
Chkbarcode(84).Caption = "Allow Trips Save Without Expenses"
Chkbarcode(85).Caption = "Send To Approve "
Chkbarcode(133).Caption = "Batch Create many work order"
lbl(31).Caption = "VAT Reg. No'"
Chkbarcode(112).Caption = "Trip with order Only"
Chkbarcode(112).Caption = "Price according to size"
Chkbarcode(71).Caption = "Get Last Price"
Chkbarcode(72).Caption = "work with mix on each Line"
Chkbarcode(74).Caption = "Work with Cahing Method"
Chkbarcode(76).Caption = "Return Without Cost"
Chkbarcode(69).Caption = "Project Invoice Decimal"
Chkbarcode(70).Caption = "Production Work with Mix"
Chkbarcode(78).Caption = "only one issue Voucher for Prod. Order"
Chkbarcode(73).Caption = "Allow Edit Mix Qty"
Label8.Caption = "No of Reservation"
Chkbarcode(58).Caption = "Lock Salary"
Chkbarcode(57).Caption = "Allow Hide Assest"
Chkbarcode(55).Caption = "Allow Edit Account"
Chkbarcode(56).Caption = "Worl With Under implementation"
Chkbarcode(80).Caption = "Allow Good Performance Account"
Chkbarcode(81).Caption = "Muste Enter Manual number in invoice"
Chkbarcode(82).Caption = "All Items In VAT"
Chkbarcode(54).Caption = "View Account With Branch"
Chkbarcode(38).Caption = "Work with Prent Group Code"
Chkbarcode(0).Caption = "Work with BarCode"
Chkbarcode(7).Caption = "Work with Linked Items"
Chkbarcode(8).Caption = "Work with Branch Logo"

Chkbarcode(143).Caption = "The bond of consolidation does not constitute a bond of exchange"
Chkbarcode(144).Caption = "The bill of sale in the bond of exchange does not constitute a bond"
Chkbarcode(145).Caption = "Items are entered manually into the exchange voucher built on a sales invoice"


Chkbarcode(9).Caption = "Work with First Installmets in contract"
Chkbarcode(11).Caption = "Create insurance Account For Each Customers"
Chkbarcode(12).Caption = "Decide Item Name According To Group Name "
Chkbarcode(13).Caption = "Default is Credit Sales "
Chkbarcode(46).Caption = "Default is Credit Purshase "
Chkbarcode(137).Caption = "Default is Credit RET "
Chkbarcode(47).Caption = "Return By Barcode Only "
Chkbarcode(14).Caption = "Jl Coding According To Branch "
Chkbarcode(15).Caption = "Sales Person Not Exceed Discount %"
Chkbarcode(16).Caption = "Create Box loss and increase account for each box "
Frame58.Caption = "View Options"
lbl(22).Caption = "Image Folder"
lbl(26).Caption = "Report Folder"
lbl(28).Caption = "Membership No."
lbl(29).Caption = "Computer No."
chkRentInstallments.Caption = "View Rent Alarms"
lbl(24).Caption = "Minute"
Chkbarcode(28).Caption = "Invoice Disc. Create GE"
Chkbarcode(29).Caption = "Cash & Cheque Different Print"
Chkbarcode(30).Caption = "Payroll work With one Accont"
Chkbarcode(31).Caption = "work With Items Details "
Chkbarcode(32).Caption = "F.A Addition Create Account "
Chkbarcode(33).Caption = "Owner have 2   Account2 "

Chkbarcode(39).Caption = "According to the store"
Chkbarcode(41).Caption = "Cost according to the order"
Chkbarcode(49).Caption = "Dealing with the area"
Chkbarcode(25).Caption = "Fifo"
Chkbarcode(96).Caption = "Receipt of receipt at the supplier's account"
Chkbarcode(130).Caption = "Closing the transfer bills in the sales invoices"
Chkbarcode(132).Caption = "Multiple items in the collection document"
Label16.Caption = "Default credit limit option"
Label19.Caption = "Credit period on default days"
Frame56.Caption = "Encoding premium bonds"
Frame55.Caption = "Encapsulation of exchange bonds"
Frame45.Caption = "Financial Transaction Options"

Chkbarcode(105).Caption = "Client file The record number is not mandatory"
Chkbarcode(113).Caption = "Price according to size"
Chkbarcode(121).Caption = "Allow duplicate invoice number"
Chkbarcode(122).Caption = "The returns are treated with FIFO"
Chkbarcode(51).Caption = "Accelerate payment policy"
Chkbarcode(127).Caption = "Allows modifi of invoices that are marked by alerts"
Chkbarcode(45).Caption = "Cash customer,phone are mandatory in the invoice"
Chkbarcode(128).Caption = "Allows modification of invoices with refunds"
Chkbarcode(102).Caption = "Taxable"
Chkbarcode(48).Caption = "The commission on the visa is deducted from the total"
Chkbarcode(44).Caption = "Dealing with multiple payment methods in the invoice"
Chkbarcode(106).Caption = "Dealing with the prepay account in the customer's receipts"
Chkbarcode(108).Caption = "Dealing with the rest in the bill of exchange of materials based on the invoice"
Chkbarcode(131).Caption = "Do not allow the delegate to modify his sales bill"
Chkbarcode(103).Caption = "Dealing with customer points"
Chkbarcode(50).Caption = "Dealing with multiple payment methods in your purchase invoice"
Chkbarcode(83).Caption = "Underline the filter in the count"
Chkbarcode(123).Caption = "FIFO discount allowed"
Chkbarcode(124).Caption = "Dealing with multiple payment methods in receipts"
Chkbarcode(161).Caption = "The bonds analytic disbursement expenses integrate value added"
 
Chkbarcode(35).Caption = "Employee benefits are transferred during the transfer"
Chkbarcode(79).Caption = "Underwriting according to management"
Chkbarcode(86).Caption = " salary entry is analytical according to the stomach"
Chkbarcode(97).Caption = "Mandatory branch in salary"
Label13.Caption = "Number of digits after the decimal point of the salary "
Fra(5).Caption = "When a new voucher is registered"
Chkbarcode(101).Caption = "Establish the restriction on the casual leave"
Chkbarcode(125).Caption = "Show the employee's balance in the payroll"
Chkbarcode(129).Caption = "Provision is made in accordance with the administrations"
Chkbarcode(109).Caption = "Flight data The data is added to the spreadsheet automatically"
Chkbarcode(114).Caption = "Connect the customer to his cars only"
Chkbarcode(116).Caption = "On the Transfer Clients screen, the price is at the journal level"
Chkbarcode(88).Caption = "Insurance for the owner"
Chkbarcode(89).Caption = "Water, electricity and services for the owner"
Chkbarcode(90).Caption = "The maturity of the quest"
Chkbarcode(91).Caption = "Water entitlement"
Chkbarcode(92).Caption = "Electricity maturity"
Chkbarcode(93).Caption = "Service entitlement"
Chkbarcode(95).Caption = "The maturity of the commission"
Chkbarcode(94).Caption = "The commission is borne by the owner"
Chkbarcode(98).Caption = "Allow batch override"
Chkbarcode(110).Caption = "Do not create the restriction automatically in the contracts"
Chkbarcode(111).Caption = "Allow adjustment in payments in contracts"
Chkbarcode(117).Caption = "Not to establish the restriction in the tenancy contracts"
Chkbarcode(118).Caption = "Open a value-added account for each owner"
Chkbarcode(119).Caption = "The establishment of the commissioning of monies in the receipts"
Chkbarcode(120).Caption = "The type of contract automatically from the screen of the property"
Chkbarcode(53).Caption = "In the project invoices separate the surreal client from the surreal contractor"
Chkbarcode(99).Caption = "Allow adjustment of the price adopted in the abstracts"
Chkbarcode(104).Caption = "Analytical constraint in project abstracts"
Chkbarcode(115).Caption = "The possibility of modifying the bonds of arrest linked to projects"
Chkbarcode(126).Caption = "The bill of arrest of Bismuth project in the project statement"

Frame49.Caption = "Numbers"
lbl(13).Caption = "Number of digits after the decimal point of the currency"
lbl(14).Caption = "Number of digits after decimal point of quantity"
lbl(17).Caption = "Number of digits after the decimal point of premiums"
lbl(16).Caption = "The number of expected sales invoices"
Check1.Caption = "Reminder alert all"
Frame5.Caption = "Alerts items"
Frame8.Caption = "Documentary Credits / Bank Guarantees"
CHECK_OPEN_NEW_SCREEN.Caption = "When you open any screen you start automatically"
Option4(0).Caption = "Save only"
Option4(1).Caption = "Save and print on screen"
Option4(2).Caption = "Save and print to virtual printer"
Option4(3).Caption = "Save and print to the default printer and open a new screen"
lbl(30).Caption = "The main password"
ChkMessnger.Caption = "View internal mail automatically"
ChkViewAging.Caption = "View debt reconstruction in the statements of accounts"
ChkPrintBranchINGE.Caption = "Display the branch name next to the account in the constraint"







Frame44.Caption = "Store options"

'Create2account4Supp
lbl(25).Caption = "Cashing Default value"
 lbl(27).Caption = "Discount Policy"
Chkbarcode(27).Caption = "Qty In PO is total Qty from Internal Order "
Chkbarcode(17).Caption = "Attached items is free "
Chkbarcode(34).Caption = "Enable Customer Aging"
Chkbarcode(18).Caption = "Show items Cost Alarm in invoices "
Chkbarcode(19).Caption = "Sub-Contractor have 3 Accounts"
Chkbarcode(20).Caption = "Create Gv for Employee"
Chkbarcode(21).Caption = "Purchase Without decimal"
Chkbarcode(22).Caption = "work With Customer Contract "
Chkbarcode(23).Caption = "Cancel All Approve "
Chkbarcode(24).Caption = "work with Allocation "
Chkbarcode(25).Caption = "work with Vendor Contract "
Chkbarcode(26).Caption = "Po Create Voucher "
Chkbarcode(26).Caption = "Sales Discount create voucher "

Chkbarcode(10).Caption = "Work with GroupCode in items "
Chkbarcode(1).Caption = "Allow Repeated name"
ChkItemsattachedzero.Caption = "Attached Item With no price"
Chkbarcode(2).Caption = "View Tradational POS"
Chkbarcode(87).Caption = "View POS Shape 2"

Chkbarcode(3).Caption = "Allow Update Sales Invoice "
Chkbarcode(107).Caption = "The value added is not calculated in financial transactions "
Chkbarcode(77).Caption = "Allow Update Purchase Invoice "

Chkbarcode(42).Caption = "Analytic Payment Voucher "
Chkbarcode(43).Caption = "Show Driver Only "

Chkbarcode(4).Caption = "Allow Update Cashing Voucher"

Chkbarcode(6).Caption = "Allow Update Employee Status"
Chkbarcode(36).Caption = "Allow First and last Finger Print"
Chkbarcode(37).Caption = "Allow Short Name"
Chkbarcode(60).Caption = "Employee Opening Balance Show All"
Chkbarcode(61).Caption = "Sales orders with sales invoices dealing with the remaining"
Chkbarcode(62).Caption = "End of service calculation of 5 years"
Chkbarcode(75).Caption = "Vacation Show old Salaries"
Chkbarcode(76).Caption = "Allow Return Without Cost"

Chkbarcode(63).Caption = "Show items according to customer"
Chkbarcode(64).Caption = "Raw materials according to mixture"
Chkbarcode(142).Caption = "Raw materials according to mixture Sales invoice"
Chkbarcode(65).Caption = "Link users with payment methods"
Chkbarcode(66).Caption = "VAT No. According to Activity"
Chkbarcode(67).Caption = "Not Create  Reseive Voucher For Purchase Project"
Chkbarcode(68).Caption = "Sales of trust affected by the resource account"
Chkbarcode(5).Caption = "Dealing revenues owed"
Label18.Caption = "No Booking"
Label23.Caption = "Branch Digit No"
Label24.Caption = "Store Digit No"
Label20.Caption = "Define Item Code Seperator"

Frame54.Caption = "Customer"
Command1(2).Caption = "Define Temperory Messages"

Label1(0).Caption = "SalesPerson Discount"

Label18.Caption = "Report Zoom"
Label21.Caption = "Logo Width"
Label22.Caption = "Logo Hight"



Cmd.Caption = "Select Logo"
chkuserCode.Caption = "Login with Code Only"
    lbl(20).Caption = "Web Site"
    lbl(18).Caption = "Fax "
    ChkDriverBox.Caption = "Create Driver Box Account"
    chkDriverEra.Caption = "Create Driver Era Account"
    ChkItemsattachedzero.Caption = "show Items attached with No Price"
    Combo1.Clear
    Combo1.AddItem "Day"
    Combo1.AddItem "Month"
    Combo1.AddItem "Year"
    ChkHideAllAlarms.Caption = "Hide Alarms Screen"
    Frame38.Caption = "Transportation Alarms"
    Frame37.Caption = "Reserved Qty="
    OptCurrQty(0).Caption = "CurrQty"
    OptCurrQty(1).Caption = "CurrQty+res"
    lbl(38).Caption = "Salary Components Decimals"
    Fra(6).Caption = "Undirect Cost"
    ChkPrintCCinGE.Caption = "Show Cost Center In GL"
    Fra(4).Caption = "Defalut Date"
lbl(21).Caption = "Remarks"
    Combo2.Clear
    Combo2.AddItem "Day"
    Combo2.AddItem "Month"
    Combo2.AddItem "Year"
    ChkChartPrintinAS.Caption = "Print Chart In AS"
    Frame36.Caption = "Installments Vchr Coding"
'    lblIndirectCostPercentage.Caption = "Percentage"
    Frame32.Caption = "Productions"
    ChkTypicalProduction.Caption = "Typical Productions"
    Frame31.Caption = "Fixed Assets"
    Label7.Caption = "Return Period"
    Label9.Caption = "change Period"
    Label5.Caption = "Day"
    lbl(37).Caption = "Day"
    Frame33.Caption = "Expenses Vchr Coding"
    ChkExpensesCoding.Caption = "Expenses And payement "
    ChkExpensesCoding2.Caption = "Transfer And expenses "

    chkInstallmntsvchrCoding.Caption = "Installment And Payment have same code"
    Frame42.Caption = "Transfer And expenses  have same code"

    Combo3.Clear
    Combo3.AddItem "Day"
    Combo3.AddItem "Month"
    Combo3.AddItem "Year"

    Combo4.Clear
    Combo4.AddItem "Day"
    Combo4.AddItem "Month"
    Combo4.AddItem "Year"

    Combo5.Clear
    Combo5.AddItem "Day"
    Combo5.AddItem "Month"
    Combo5.AddItem "Year"

    Combo6.Clear
    Combo6.AddItem "Day"
    Combo6.AddItem "Month"
    Combo6.AddItem "Year"

    Combo10.Clear
    Combo10.AddItem "Day"
    Combo10.AddItem "Month"
    Combo10.AddItem "Year"

    Combo11.Clear
    Combo11.AddItem "Day"
    Combo11.AddItem "Month"
    Combo11.AddItem "Year"

    Combo12.Clear
    Combo12.AddItem "Day"
    Combo12.AddItem "Month"
    Combo12.AddItem "Year"

    Frame20.Caption = "Work With Items"
    ChkitemsWorkWithSize.Caption = "Size"
    'ChkitemsWorkWithSize.Caption = "Work With Barcode"
    Command1(0).Caption = "Info Bar Settings"
    
    ChkCostStarting.Caption = "Cost Starting at this interval"
    ChkitemsWorkWithColor.Caption = "Color"
    ChkitemsWorkWithDates.Caption = "Pro. & Exp Dates"
    ChkitemsWorkWithClass.Caption = "Class"
    ChKautoIssueVoucher.Caption = "Create Auto Issue Voucher For Sales Invoice"
 
    ChKautoReseiveVoucher.Caption = "Create Auto Reseive Voucher For Purchase Invoice"
    ChkBankComm.Caption = "Work With Bank Commission"
    chkChequeBox.Caption = "Enable Cheque Box"
    chkCustomerhavethreeAccounts.Caption = "Enable Due Check To Customers"
    chkCustomerhavethreeAccounts1.Caption = "Enable Due Check To Vendors"
    
    ChkAssetAccount.Caption = "Accum Accounts in Assets"
    ChkAssetAccount1.Caption = "Create Profit and Lose Acc. For Each Group"
'********************************************

 chkStore(0).Caption = "Account inventory adjustments following the inventory"
 chkStore(1).Caption = " Each store has its expense and damage Account"
 chkStore(2).Caption = "Each store has its Gifts and sample Account"
 chkStore(3).Caption = "Dealing with more than one store"
'*********************************************
    ChkAllowIndirectCost.Caption = "Allow Indirect Cost"
    chkEmpProduction.Caption = "Employment"
    chkItemProduction.Caption = "Raw materials"
    chkExpProduction.Caption = "Outlay"
    chkMonthIs30days.Caption = "Month Is 30 days"
    Me.Frame35.Caption = "Hr Mangement"

    Frame21.Caption = "Arrows Follow"
    Frame22.Caption = "Arrows Linked With GL"
    OptArrowBranch.Caption = "Linked With Branch"
    OptArrowGroup.Caption = "Linked With Arrows Groups"
    Frame23.Caption = "Arrows Evaluation"
    Option1.Caption = "Purchase Price Avearge"
    Option2.Caption = "According to purchase Price"

    ChkAsk.Caption = "Show Print Options"
    Me.Caption = "Options"
    Me.Optday.Caption = "Days"
    Optweek.Caption = "Weeks"
    Me.OptMonth.Caption = "Months"
    Me.OptYear.Caption = "Years"

    Frame19.Caption = "Process Period"

    lbl(10).Caption = "Company Name"
    lbl(8).Caption = "Title"
    lbl(9).Caption = "Address"
    lbl(5).Caption = "Tel"
    lbl(1).Caption = "Mob"
    lbl(3).Caption = "E-mail"
    lbl(2).Caption = "Contact person"
    lbl(9).Caption = "Address"
    Fra(2).Caption = "Logo"
    chk.Caption = "View Logo in reports"
    cmdOK.Caption = "Save"
    cmdCancel.Caption = "Exit"
    '777777777777777
    Fra(14).Caption = "Working With Project Type"
    OptionItemsTotal.Caption = "Total Items"
    OptionOperation.Caption = "Detailed Operaion"
Me.OPTdISCOUNT(0).Caption = "Discount effet Expenses"
Me.OPTdISCOUNT(1).Caption = "Discount effet Revenue"


    Frame18.Caption = "Project Expenses JL  "
    GlDetails.Caption = "Detailed Jl On Expenses"
    glgeneral.Caption = "Jl On Project Accounts"
    Fra(3).Caption = "Policy replacement of the goods"
    opt(6).Caption = "Without or With  bill"
    opt(7).Caption = " With  bill Only"
    Label5.Caption = "Days"
    Frame28.Caption = "Discount Priorities "
    Label2.Caption = "Sort"
    Label3.Caption = "Item Discount"
    Label4.Caption = "Item Group Discount"
    Label6.Caption = "Customers Discount"
    'Label1.Caption = "Delegate Discount"


    Frame24.Caption = "RealState Mangement Alarms"
'    chkExpProduction.Caption = "Show alerts rents before"
    Check6.Caption = "Show  Contracts will be Expire  before"
    Check7.Caption = "Show  Expired Contracts  before"
    Frame25.Caption = "Show Before"
    Frame26.Caption = "Show Before"
    Frame27.Caption = "Show Before"
'    chkEmpProduction.Caption = "Arrows Pecentage Of Profit And Loss Alarm"
 
    '77777777
    Me.tb.TabCaption(0) = "Company. Info"
    Me.tb.TabCaption(1) = "fiscal Years"
    Me.tb.TabCaption(2) = "Activity and Branches"
    Me.tb.TabCaption(3) = "Accounts Integration"
    Me.tb.TabCaption(4) = "Inventory Options"
    Me.tb.TabCaption(5) = "Sale Options"
    Me.tb.TabCaption(6) = "Purchase Options"
    Me.tb.TabCaption(7) = "Financial Options"
    Me.tb.TabCaption(8) = "HR Options"

    Me.tb.TabCaption(9) = "Production Options"
    Me.tb.TabCaption(10) = "Transportation Options"
    Me.tb.TabCaption(11) = "RealEstate Options"
    Me.tb.TabCaption(12) = "Fixed Asssets Options"
    Me.tb.TabCaption(13) = "Projects Options"
    Me.tb.TabCaption(14) = "Arrows Follow"
    Me.tb.TabCaption(15) = "General options"
    Me.tb.TabCaption(16) = "Alarms Manger"
    Me.tb.TabCaption(17) = "View Options"

Me.tb.TabCaption(18) = "Accounts Coding"
Me.tb.TabCaption(19) = "Documents Types"
Me.tb.TabCaption(20) = "Documents Coding"
Me.tb.TabCaption(21) = "Field Coding"
Me.tb.TabCaption(22) = "Internal Rules"
Me.tb.TabCaption(23) = "Info Bar "
Label1(2).Caption = "Info Bar Settings"
'Command1(10).Caption = "Settings"
Command1(1).Caption = "Forms"
Command1(2).Caption = "Temporary Msg"

Frame59.Caption = "Save Options"

    ALLButton1.Caption = "Periods "
    ALLButton7.Caption = "Branches"
    ALLButton6.Caption = "Acc. Link"
    ALLButton2.Caption = "Acc. Coding"
    ALLButton5.Caption = "DocS Types"
    ALLButton3.Caption = "VchrS. Coding"
    ALLButton4.Caption = "Fileds Coding"
 
    lbl(7).Caption = "Default Custoomer"
    lbl(4).Caption = "Default Inventorty"
    lbl(11).Caption = "Default Box"
   ' ChkTax.Caption = "Add Tax To Sales invoice"
    Fra(0).Caption = "When Add New Bill"
    opt(0).Caption = "today Date is the Default date "
    opt(8).Caption = "today Date is the Default date "
    
    DateOpt(0).Caption = "G Date"
    DateOpt(1).Caption = "Higri Date"
    

    opt(1).Caption = "Last Sales invoice Date is the Default date "
    opt(2).Caption = "Server  Date is the Default date - in network only "

   opt(9).Caption = "Last Sales   Date is the Default date "
    opt(10).Caption = "Server  Date is the Default date - in network only "


    lbl(6).Caption = "Default Vendor"
    lbl(0).Caption = "Default Inventorty"
    Fra(1).Caption = "When Add New Bill"
    opt(5).Caption = "today Date is the Default date "
    opt(4).Caption = "Last Purchase invoice Date is the Default date "
    opt(3).Caption = "Server  Date is the Default date - in network only "

    Frame9.Caption = "Before"
    Fra(8).Caption = "Before"
    Fra(9).Caption = "Before"
    Fra(11).Caption = "Before"
    Fra(12).Caption = "Before"
    Fra(13).Caption = "Before"

    Frame39.Caption = "Before"
    Frame40.Caption = "Before"
    Frame41.Caption = "Before"

    ChkDelayVal.Caption = "Securities outstanding"
    ChkInstallmentMustPayed.Caption = "Premium payable"
    ChKProjectsAlarm1.Caption = "Show Projects Alarms"
    ChKProjectsAlarm2.Caption = "Show Projects Bill  Alarms"

    ChKHR.Caption = "Show Employees Alarms"
    Fra(10).Caption = "Employees Alarms"
    ChkExpireEkama.Caption = "Show End Ekama Before"
    Me.CheckLC.Caption = "LC Alarms"
    
    ChkExpireLicence.Caption = "Show End License"
    ChkExpireLicense.Caption = "Show End License Before"
    ChkExpireInsurance.Caption = "Show End nsurance  Before"
    ChkExpireTest.Caption = "Show End Test  Before"

    ChkExpirepas.Caption = "Show End Passport"
    ChkExpirepoket.Caption = "Show End Saudi ID"
    ChkShow.Caption = "show  that amounted to a demand Alarm"
    ChkShowToolTip.Caption = "Show daily help"
    ChkGuranAlram.Caption = "show  that amounted to a Risk alarm"

    lbl(12).Caption = "Cost price Type"
    Chk1.Caption = "Allow Negative Box"
    chk2.Caption = "Allow Negative warehouse"
    Fra(7).Caption = "Sales Invoice and Issue Voucher"
    Opt_OrderOut.Caption = "Allow to make Issue voucher and make bill at another time"
    ChkEmpRes(0).Caption = "Monthly Reserv"
    ChkEmpRes(1).Caption = "Yearly Reserv"
    Opt_Sal.Caption = "DisAllow to make Issue voucher  "
    Frame2.Caption = "Purcahase  Invoice and Recieve Voucher"

    Opt_OrderInpo.Caption = "Allow to make Recieve voucher and make bill at another time"
    Opt_Bey.Caption = "DisAllow to make Recieve voucher  "

    Frame3.Caption = "banks"
    chk3.Caption = "Create 3 Account For Each Bank"
'    Frame16.Caption = "Employee Accounts"
    Chkemployeeaccounts.Caption = "Create 3 Account For Each Employee"
    Frame4.Caption = "wareHouse And Groups"
'    opt_Branch.Caption = "Join with Branch"
'    opt_group.Caption = "Join with Group"
'    Opt_Inventory_create_account.Caption = "Inventory only"
'    opt_inv_and_branch_create_account.Caption = "Inventory And Branch"
    lbl(13).Caption = "Decimal Places For Currency"
    lbl(14).Caption = "Decimal Places For Quantity"
    lbl(15).Caption = "Gl predect Numbers 000"
    lbl(16).Caption = "Account predect Numbers 000"
    lbl(17).Caption = "Installments Decimals"
'    Frame6.Caption = "Accounting periods"
    ALLButton1.Caption = "Setup"
    CHECK_OPEN_NEW_SCREEN.Caption = "Any Screen Start with new"
    ChkPrintBranchINGE.Caption = "Print Branch Name With A.c. in GL"
    ChkChartPrintinAS.Caption = "Print CC in GE"

    ChkViewAging.Caption = "View Aging In Account Statement"
    ChkMessnger.Caption = "Allow Messnger"
    Frame8.Caption = "Saving options"
    Option4(0).Caption = "Save only"
    Option4(1).Caption = "Save and print preview"
    Option4(2).Caption = "Save and print to default printer"
    Option4(3).Caption = "Save and print to default printer , New operation"
    chkshortCuts.Caption = "view Shortcuts"
    Chktree.Caption = "view Items Tree"
    ChkCalender.Caption = "view Calender"
    Chkgraphic.Caption = "view Daily movement summary"

    Frame44.Caption = "Invenotory Settings"
    Frame45.Caption = "Financial Transactions Settings"
    Frame55.Caption = "Expenses Voucher Coding"
    Frame56.Caption = "Installments Voucher Coding"
    Frame48.Caption = "The method of calculating allocations"
    Frame49.Caption = "General Numbers"
    Check1.Caption = "Show Alaram"

'    With Combo13
  '      .Clear
  '      .AddItem "Minutes"
  '      .AddItem "Hours"
  '      .AddItem "Day"

  '  End With

    Frame58.Caption = "View Settings"
'    Label13.Caption = "Percentage"
    Frame4.Caption = "General Settings"
    Frame3.Caption = "Projects"
    Frame5.Caption = "Ittems"
    Frame8.Caption = "Stocks Management"
    ALLButton8.Caption = "Internal Rules"
    
    
    Chkbarcode(39).Caption = "Based on Store"
    Chkbarcode(49).Caption = "Dealing by Area"
    Chkbarcode(52).Caption = "Dealing by Lot and Expiration Date"
    Chkbarcode(41).Caption = "Cost According to The Serial"
    Chkbarcode(59).Caption = "Dealing by Multi-payment in Sales Invoice"
    Chkbarcode(44).Caption = "Dealing by Multi-payment in Sales Invoice"
    Chkbarcode(45).Caption = "The Name and Phone Number of Cash Client is a Must"
    Chkbarcode(48).Caption = "The Commission on The Visa is Deducted From the Total"
    Chkbarcode(51).Caption = "Accelerated Payment Policy"
    Chkbarcode(50).Caption = "Dealing by Multiple Payment Methods in Purchase Invoice"
    Chkbarcode(35).Caption = "Employee Benefits are Transferred During the Transfer"
    Fra(5).Caption = "When a New Receiving Voucher is Registered"
    Chkbarcode(53).Caption = "In the Project Invoices, the Serial Customer Will be Separated From the Serial Contractor"
    lbl(30).Caption = "Primary Password"
    
End Sub
Private Sub SetComboBox()
 Dim Dcombos As ClsDataCombos
     Set Dcombos = New ClsDataCombos
    Dcombos.GetBoxes Me.DcboBox

    If SystemOptions.UserInterface = EnglishInterface Then
        SetInterface Me
        ChangeLang

    End If
    
    
With Me.DefaultInvoicetype
            .Clear
            
             


            .AddItem " ›« Ê—… ÷—Ì»Ì…  "
            .ItemData(0) = 0
     
            .AddItem " ›« Ê—… „»”ÿ… "
            .ItemData(1) = 2
         
        End With
        
   With Me.Invoicetype
            .Clear
            
             


            .AddItem "standard  Invoices only  ›« Ê—… ÷—Ì»Ì…  ›ﬁÿ"
            .ItemData(0) = 0
            .AddItem "standard & Simplified Invoices  ›« Ê—… ÷—Ì»Ì… Ê„»”ÿ…"
            .ItemData(1) = 1
            .AddItem "Simplified Invoices only  ›« Ê—… „»”ÿ… ›ﬁÿ"
            .ItemData(2) = 2
         
        End With
        
        
                With Me.SendingMode
            .Clear
            
             


            .AddItem "dev"
            .ItemData(0) = 0
            .AddItem " Simulation"
            .ItemData(1) = 1
            .AddItem "production"
            .ItemData(2) = 2
         
        End With
        
        
             
    If SystemOptions.UserInterface = EnglishInterface Then

        With Me.CboMainStockType
            .Clear
            .AddItem "Last Purchase Price"
            .ItemData(0) = 0
            .AddItem "FiFO"
            .ItemData(1) = 2
            .AddItem "AverageCost  in bill"
            .ItemData(2) = 4
            .AddItem "Average Cost"
            .ItemData(2) = 5
        End With


        With Me.CboChasingStatus
            .Clear
            .AddItem "Advanced Payment"
            .ItemData(0) = 3
            .AddItem "FiFO"
            .ItemData(1) = 1
            .AddItem "Select Invoice"
            .ItemData(2) = 2
            .AddItem "Old Projects"
            .ItemData(2) = 7
        End With
        
        
    Else

        With Me.CboMainStockType
            .Clear
            .AddItem "√Œ— ”⁄— ‘—«¡"
            .ItemData(0) = 0
            .AddItem "«·Ê«œ— «Ê·« Ì’—› «Ê·«"
            .ItemData(1) = 2
            .AddItem "«·„ Ê”ÿ «·„—ÃÕ ⁄·Ï «·› —…"
            .ItemData(2) = 4
            .AddItem "«·„ Ê”ÿ «·„—ÃÕ «·„⁄œ·"
            .ItemData(2) = 5
        End With


        With Me.CboChasingStatus
            .Clear
            .AddItem "œ›⁄Â „ﬁœ„…"
            .ItemData(0) = 3
            .AddItem "FiFO"
            .ItemData(1) = 1
            .AddItem " ÕœÌœ ›Ê« Ì—"
            .ItemData(2) = 2
            .AddItem "„‘«—Ì⁄ ”«»ﬁ…  "
            .ItemData(2) = 7
        End With
        
        
        
    End If

End Sub
Private Sub Form_Load()
  On Error GoTo ErrTrap
    Dim AskOption As Boolean
    Dim StrSQL As String
    Dim intDef As Integer
    Dim Dcombos As ClsDataCombos
         Set Dcombos = New ClsDataCombos
    
checksave = False
    CenterForm Me
tb.CurrTab = 0
SetComboBox
    FormPostion Me, GetPostion
  NewTab = 0

    Set rs = New ADODB.Recordset
    rs.Open "[TblOptions]", Cn, adOpenStatic, adLockOptimistic, adCmdTable

    If Not (rs.EOF Or rs.BOF) Then
 
 
 
txtNoOFDigitUser(2).text = IIf(IsNull(rs("StreetName").value), "", rs("StreetName").value)
txtNoOFDigitUser(4).text = IIf(IsNull(rs("BuildingNumber").value), "", rs("BuildingNumber").value)
txtNoOFDigitUser(9).text = IIf(IsNull(rs("CitySubdivisionName").value), "", rs("CitySubdivisionName").value)
txtNoOFDigitUser(6).text = IIf(IsNull(rs("CityName").value), "", rs("CityName").value)
txtNoOFDigitUser(7).text = IIf(IsNull(rs("PostalZone").value), "", rs("PostalZone").value)
txtNoOFDigitUser(10).text = IIf(IsNull(rs("IdentificationCode").value), "SA", rs("IdentificationCode").value)
If txtNoOFDigitUser(10).text = "" Then txtNoOFDigitUser(10).text = "SA"
txtNoOFDigitUser(5).text = IIf(IsNull(rs("PlotIdentification").value), "", rs("PlotIdentification").value)
txtNoOFDigitUser(3).text = IIf(IsNull(rs("AdditionalStreetName").value), "", rs("AdditionalStreetName").value)
txtNoOFDigitUser(8).text = IIf(IsNull(rs("CountrySubentity").value), "", rs("CountrySubentity").value)


        Me.Invoicetype.ListIndex = IIf(IsNull(rs("Invoicetype").value), 0, rs("Invoicetype").value)
            Me.DefaultInvoicetype.ListIndex = IIf(IsNull(rs("DefaultInvoicetype").value), 0, rs("DefaultInvoicetype").value)
            
            Me.SendingMode.ListIndex = IIf(IsNull(rs("SendingMode").value), 0, rs("SendingMode").value)
            
            TxtLogoWidth.text = IIf(IsNull(rs("LogoWidth").value), 4000, rs("LogoWidth").value)
            TxtLogoheight.text = IIf(IsNull(rs("Logoheight").value), 1500, rs("Logoheight").value)
            
        XPTxtCompany.text = IIf(IsNull(rs("Company_Arabic_Name").value), "", rs("Company_Arabic_Name").value)
        XPTxtCompanye.text = IIf(IsNull(rs("Company_Name_Eng").value), "", rs("Company_Name_Eng").value)
        
        TxtImagesPath(0).text = IIf(IsNull(rs("ImagesPath").value), "Images", rs("ImagesPath").value)
        TxtImagesPath(1).text = IIf(IsNull(rs("reportPath").value), "Stander", rs("reportPath").value)
        TxtImagesPath(2).text = IIf(IsNull(rs("BigUserPw").value), "n20172018", rs("BigUserPw").value)
        TxtImagesPath(3).text = IIf(IsNull(rs("BigUserPw2").value), "123456", rs("BigUserPw2").value)
        
        'BigUserPw
        'Stander
        TXTReturnSallingIntervalCount(4).text = IIf(IsNull(rs("VATItems").value), "", rs("VATItems").value)
         TXTReturnSallingIntervalCount(3).text = IIf(IsNull(rs("NoRoudProjectInvoices").value), "", rs("NoRoudProjectInvoices").value)
         
         TXTReturnSallingIntervalCount(6).text = IIf(IsNull(rs("CountPrint").value), "", rs("CountPrint").value)
         
         
        TXTReturnSallingIntervalCount(2).text = IIf(IsNull(rs("NoBooking").value), "", rs("NoBooking").value)
        TXTReturnSallingIntervalCount(1).text = IIf(IsNull(rs("itemSeprator").value), "", rs("itemSeprator").value)
        TXTReturnSallingIntervalCount(5).text = IIf(IsNull(rs("DefaultQtyTrans").value), 1, rs("DefaultQtyTrans").value)
        
        TxtFax.text = IIf(IsNull(rs("Fax").value), "", rs("Fax").value)
        
        txtNoOFDigitUser(0) = IIf(val(rs!IsSerialByUserTrans & "") = 0, 2, val(rs!IsSerialByUserTrans & ""))
        txtNoOFDigitUser(1) = IIf(val(rs!NoOFDigitUserVouc & "") = 0, 2, val(rs!NoOFDigitUserVouc & ""))
        
        XPTxtComment(0).text = IIf(IsNull(rs("Company_Comment").value), "", rs("Company_Comment").value)
        XPTxtComment(2).text = IIf(IsNull(rs("ComputerNo").value), "", rs("ComputerNo").value)
        XPTxtComment(3).text = IIf(IsNull(rs("VATRegNo").value), "", rs("VATRegNo").value)
        
       TxtEmails.text = IIf(IsNull(rs("WEBSite").value), "", rs("WEBSite").value)
        
        XPTxtComment(1).text = IIf(IsNull(rs("MembershipNo").value), "", rs("MembershipNo").value)
        XPTxtAddress.text = IIf(IsNull(rs("Company_Address").value), "", rs("Company_Address").value)
        xptxtphone.text = IIf(IsNull(rs("Company_Phone").value), "", rs("Company_Phone").value)
        XPTxtmobile.text = IIf(IsNull(rs("Company_Mobile").value), "", rs("Company_Mobile").value)
        XPTxtMail.text = IIf(IsNull(rs("Company_Maile").value), "", rs("Company_Maile").value)
        txtDomainData.text = IIf(IsNull(rs("DomainData").value), "", rs("DomainData").value)
        
        XPTxtResponsable.text = IIf(IsNull(rs("Company_Responsable").value), "", rs("Company_Responsable").value)
        Me.DcboBox.BoundText = IIf(IsNull(rs("SalesBoxID").value), "", rs("SalesBoxID").value)

        If rs("DateOpt").value = 0 Then
            Me.DateOpt(0).value = True
        ElseIf rs("DateOpt").value = 1 Then
            Me.DateOpt(1).value = True
        Else
            Me.DateOpt(0).value = True
        End If

        If rs("MainStockCostType").value = 0 Then
            Me.CboMainStockType.ListIndex = 0
        ElseIf rs("MainStockCostType").value = 2 Then
            Me.CboMainStockType.ListIndex = 1
        ElseIf rs("MainStockCostType").value = 4 Then
            Me.CboMainStockType.ListIndex = 2
        ElseIf rs("MainStockCostType").value = 5 Then
            Me.CboMainStockType.ListIndex = 3
        End If
        
        If IsNull(rs("CashDate").value) Then
        opt(8).value = True
        Else
                If rs("CashDate").value = 0 Then
                opt(8).value = True
                ElseIf rs("CashDate").value = 1 Then
                opt(9).value = True
                 ElseIf rs("CashDate").value = 2 Then
                opt(10).value = True
                End If
                
                
         End If
         
        
        
        
        If IsNull(rs("InvDate").value) Then
        opt(0).value = True
        Else
                If rs("InvDate").value = 0 Then
                opt(0).value = True
                ElseIf rs("InvDate").value = 1 Then
                opt(1).value = True
                 ElseIf rs("InvDate").value = 2 Then
                opt(2).value = True
                End If
                     
         End If
         
         
       If IsNull(rs("PurDate").value) Then
              opt(5).value = True
        Else
                If rs("PurDate").value = 0 Then
                opt(5).value = True
                ElseIf rs("PurDate").value = 1 Then
                opt(4).value = True
                 ElseIf rs("PurDate").value = 2 Then
                opt(3).value = True
                End If
                     
         End If
         
  'ChasingStatus
  
  If IsNull(rs("ChasingStatus").value) Then
  Me.CboChasingStatus.ListIndex = 1
  Else
            If rs("ChasingStatus").value = 3 Then
            Me.CboChasingStatus.ListIndex = 0
        ElseIf rs("ChasingStatus").value = 1 Then
            Me.CboChasingStatus.ListIndex = 1
        ElseIf rs("ChasingStatus").value = 2 Then
            Me.CboChasingStatus.ListIndex = 2
        ElseIf rs("ChasingStatus").value = 7 Then
            Me.CboChasingStatus.ListIndex = 3
        End If
            
  End If
        
        
    End If

    If IsNull(rs("ShowLogoInReports").value) Then
        Me.chk.value = vbUnchecked
    Else
        Me.chk.value = IIf((rs("ShowLogoInReports").value = True), vbChecked, vbUnchecked)
    End If

    If Not (IsNull(rs("CompanyLogo").value)) Then
        LoadPictureFromDB ImgPic, rs, "CompanyLogo"
    End If

    If rs("AllowBoxNegative").value = 1 Then
        Me.Chk1.value = vbChecked
    Else
        Me.Chk1.value = vbUnchecked
    End If

    If rs("banks_Accounts").value = vbTrue Then
        Me.chk3.value = vbChecked
    Else
        Me.chk3.value = vbUnchecked
    End If

    If rs("BankComm").value = vbTrue Then
        Me.ChkBankComm.value = vbChecked
    Else
        Me.ChkBankComm.value = vbUnchecked
    End If

    If rs("ChequeBox").value = vbTrue Then
        Me.chkChequeBox.value = vbChecked
    Else
        Me.chkChequeBox.value = vbUnchecked
    End If
    
    
   
    If rs("IsCheque").value = vbTrue Then
        Me.chkIsCheque.value = vbChecked
    Else
        Me.chkIsCheque.value = vbUnchecked
    End If
    
    

    
    If rs("CustomerhavethreeAccounts").value = vbTrue Then
        Me.chkCustomerhavethreeAccounts.value = vbChecked
    Else
        Me.chkCustomerhavethreeAccounts.value = vbUnchecked
    End If
        
        
    
    If rs("IsCreateOpenBalnceMan").value = vbTrue Then
        Me.chkIsCreateOpenBalnceMan.value = vbChecked
    Else
        Me.chkIsCreateOpenBalnceMan.value = vbUnchecked
    End If
        
            
   

     
        
    If rs("CustomerhavethreeAccounts1").value = vbTrue Then
        Me.chkCustomerhavethreeAccounts1.value = vbChecked
    Else
        Me.chkCustomerhavethreeAccounts1.value = vbUnchecked
    End If
        
        
    If rs("TypicalProduction").value = vbTrue Then
        Me.ChkTypicalProduction.value = vbChecked
    Else
        Me.ChkTypicalProduction.value = vbUnchecked
    End If

    If rs("ExpensesCoding").value = vbTrue Then
        Me.ChkExpensesCoding.value = vbChecked
    Else
        Me.ChkExpensesCoding.value = vbUnchecked
    End If

    If rs("ExpensesCoding2").value = vbTrue Then
        Me.ChkExpensesCoding2.value = vbChecked
    Else
        Me.ChkExpensesCoding2.value = vbUnchecked
    End If

    If rs("InstallmntsvchrCoding").value = vbTrue Then
        Me.chkInstallmntsvchrCoding.value = vbChecked
    Else
        Me.chkInstallmntsvchrCoding.value = vbUnchecked
    End If

    If rs("AssetAccount").value = vbTrue Then
        Me.ChkAssetAccount.value = vbChecked
    Else
        Me.ChkAssetAccount.value = vbUnchecked
    End If
'****************************************************************************
   If rs("StoreAccountHaveSettelment").value = vbTrue Then
        Me.chkStore(0).value = vbChecked
    Else
        Me.chkStore(0).value = vbUnchecked
    End If
    
   If rs("eachStoreHaveLossAccount").value = vbTrue Then
        Me.chkStore(1).value = vbChecked
    Else
        Me.chkStore(1).value = vbUnchecked
    End If
    
       If rs("eachStoreHaveGiftAccount").value = vbTrue Then
        Me.chkStore(2).value = vbChecked
    Else
        Me.chkStore(2).value = vbUnchecked
    End If
    
    If rs("MultyStore").value = vbTrue Then
        Me.chkStore(3).value = vbChecked
    Else
        Me.chkStore(3).value = vbUnchecked
    End If
    
             
    If rs("IsAutoNameItems").value = vbTrue Then
        Me.chkStore(4).value = vbChecked
    Else
        Me.chkStore(4).value = vbUnchecked
    End If
    
     If rs("AllowItemByRowMove").value = vbTrue Then
        Me.chkStore(5).value = vbChecked
    Else
        Me.chkStore(5).value = vbUnchecked
    End If
    
    If rs("AllowItemByRowOut").value = vbTrue Then
        Me.chkStore(6).value = vbChecked
    Else
        Me.chkStore(6).value = vbUnchecked
    End If
    
'*****************************************************************************
    If rs("AssetAccount1").value = vbTrue Then
        Me.ChkAssetAccount1.value = vbChecked
    Else
        Me.ChkAssetAccount1.value = vbUnchecked
    End If

    If rs("AllowIndirectCost").value = vbTrue Then
        Me.ChkAllowIndirectCost.value = vbChecked
    Else
        Me.ChkAllowIndirectCost.value = vbUnchecked
    End If

    If rs("EmpProduction").value = vbTrue Then
        Me.chkEmpProduction.value = vbChecked
    Else
        Me.chkEmpProduction.value = vbUnchecked
    End If
    
    
    If rs("ItemProduction").value = vbTrue Then
        Me.chkItemProduction.value = vbChecked
    Else
        Me.chkItemProduction.value = vbUnchecked
    End If
    
    
    If rs("ExpProduction").value = vbTrue Then
        Me.chkExpProduction.value = vbChecked
    Else
        Me.chkExpProduction.value = vbUnchecked
    End If
    
    
    
    

    If rs("AllowStockNegative").value = 1 Then
        Me.chk2.value = vbChecked
    Else
        Me.chk2.value = vbUnchecked
    End If

    'If Rs("Out").Value = 1 Then
    '    Me.Chk3.Value = vbChecked
    'Else
    '    Me.Chk3.Value = vbUnchecked
    'End If
    '
    'If Rs("inp").Value = 1 Then
    '    Me.Check1.Value = vbChecked
    'Else
    '    Me.Check1.Value = vbUnchecked
    'End If

    If IsNumeric(rs("Save_options").value) Then
    
        Option4(rs("Save_options").value).value = True
         
    End If
TXTReturnSallingIntervalCount(4).text = IIf(IsNull(rs("VATItems").value), 0, rs("VATItems").value)
 TXTReturnSallingIntervalCount(3).text = IIf(IsNull(rs("NoRoudProjectInvoices").value), 0, rs("NoRoudProjectInvoices").value)
 TXTReturnSallingIntervalCount(6).text = IIf(IsNull(rs("CountPrint").value), 0, rs("CountPrint").value)
 
 
 MYTEXT(0).text = IIf(IsNull(rs("NOOFPRINTCOPIESSALES").value), 0, rs("NOOFPRINTCOPIESSALES").value)
 
    If IsNull(rs("ReturnSallingOption").value) Then
     
        opt(6).value = True
        TXTReturnSallingIntervalCount(0).text = 0
        TXTReturnSallingIntervalCount1.text = 0
    Else
         
        If rs("ReturnSallingOption").value = True Then
            opt(7).value = True
            TXTReturnSallingIntervalCount(0).text = IIf(IsNull(rs("ReturnSallingIntervalCount").value), 0, rs("ReturnSallingIntervalCount").value)
            TXTReturnSallingIntervalCount1.text = IIf(IsNull(rs("ReturnSallingIntervalCount1").value), 0, rs("ReturnSallingIntervalCount1").value)
                       
        ElseIf rs("ReturnSallingOption").value = False Then
            opt(6).value = True
            TXTReturnSallingIntervalCount(0).text = IIf(IsNull(rs("ReturnSallingIntervalCount").value), 0, rs("ReturnSallingIntervalCount").value)
            TXTReturnSallingIntervalCount1.text = IIf(IsNull(rs("ReturnSallingIntervalCount1").value), 0, rs("ReturnSallingIntervalCount1").value)
                      
        End If
         
    End If
       
    If IsNull(rs("EmpRes").value) Then
        ChkEmpRes(0).value = True
    Else
     
        If rs("EmpRes").value = 0 Then
            ChkEmpRes(0).value = True
        Else
            ChkEmpRes(1).value = True
        End If
    End If
     
    If rs("checkout").value = True Then
        Opt_OrderOut.value = True
    End If
        
    If rs("Checksal").value = True Then
        Opt_Sal.value = True
    End If
        
    If rs("checkinpo").value = True Then
        Opt_OrderInpo.value = True
    End If
        
    If rs("Items_or_operation").value = 0 Then
        OptionItemsTotal.value = True
    End If
        
      If IsNull(rs("ProjectDiscountPolicy").value) Then
      OPTdISCOUNT(0).value = True
      Else
                If rs("ProjectDiscountPolicy").value = 0 Then
                OPTdISCOUNT(0).value = True
                Else
                OPTdISCOUNT(1).value = True
                End If
      
      End If
      
 
 
    
    
    If rs("Items_or_operation").value = 1 Then
        OptionOperation.value = True
    End If
        
    If rs("gl_detaila_or_total").value = 1 Then
        GlDetails.value = True
    End If
        
    If rs("gl_detaila_or_total").value = 0 Then
        glgeneral.value = True
    End If
        
    If rs("ProcessPeriodType").value = 0 Then 'ÌÊ„Ì
        Me.Optday.value = True
    End If

    If rs("ProcessPeriodType").value = 1 Then '‘Â—Ì
        Me.OptMonth.value = True
    End If
 
    If rs("ProcessPeriodType").value = 2 Then '”‰ÊÌ
        Me.OptYear.value = True
    End If
        
    If rs("ProcessPeriodType").value = 3 Then '«”»Ê⁄Ì
        Me.Optweek.value = True
    End If
        
    If rs("checkbey").value = True Then
        Opt_Bey.value = True
    End If

    If rs("Opt_branch").value = True Then
        'opt_Branch.value = True
        Frame5.Visible = False
    End If
        
    If rs("Create_employee_account").value = True Then
        Me.Chkemployeeaccounts.value = vbChecked
    Else
        Me.Chkemployeeaccounts.value = Unchecked
    End If
        
    If rs("CreateDriverBox").value = True Then
        Me.ChkDriverBox.value = vbChecked
    Else
        Me.ChkDriverBox.value = Unchecked
    End If
         
    If rs("CreateDriverEra").value = True Then
        Me.chkDriverEra.value = vbChecked
    Else
        Me.chkDriverEra.value = Unchecked
    End If
        
    If rs("itemsWorkWithSize").value = True Then
        Me.ChkitemsWorkWithSize.value = vbChecked
    Else
        Me.ChkitemsWorkWithSize.value = Unchecked
    End If
        
      If rs("WorkWithBarCodeParent").value = True Then
        Me.Chkbarcode(38).value = vbChecked
    Else
        Me.Chkbarcode(38).value = Unchecked
    End If
    
    If rs("workWithBarcode").value = True Then
        Me.Chkbarcode(0).value = vbChecked
    Else
        Me.Chkbarcode(0).value = Unchecked
    End If
        
      If rs("WorkWithLINKEDiTEMS").value = True Then
        Me.Chkbarcode(7).value = vbChecked
    Else
        Me.Chkbarcode(7).value = Unchecked
    End If
    
      If rs("WorkWithLINKEDiActivity").value = True Then
        Me.Chkbarcode(192).value = vbChecked
    Else
        Me.Chkbarcode(192).value = Unchecked
    End If
    
    
          If rs("amlaketbatrentOnly").value = True Then
        Me.Chkbarcode(193).value = vbChecked
    Else
        Me.Chkbarcode(193).value = Unchecked
    End If
    
    If rs("NotAllowStockNegativeInternal").value = True Then
        Me.Chkbarcode(194).value = vbChecked
    Else
        Me.Chkbarcode(194).value = Unchecked
    End If
    
    If rs("MustEnterNewNo").value = True Then
        Me.Chkbarcode(195).value = vbChecked
    Else
        Me.Chkbarcode(195).value = Unchecked
    End If
    
    If rs("IsInternalMultiOrder").value = True Then
        Me.Chkbarcode(196).value = vbChecked
    Else
        Me.Chkbarcode(196).value = Unchecked
    End If
    
    
    If rs("IsBlue").value = True Then
        Me.Chkbarcode(197).value = vbChecked
    Else
        Me.Chkbarcode(197).value = Unchecked
    End If
    

    
    If rs("Isthickness").value = True Then
        Me.Chkbarcode(198).value = vbChecked
    Else
        Me.Chkbarcode(198).value = Unchecked
    End If
    
    
    
    If rs("IsMashghal").value = True Then
        Me.Chkbarcode(199).value = vbChecked
    Else
        Me.Chkbarcode(199).value = Unchecked
    End If
    
    
    
    If rs("IsSalesOrder").value = True Then
        Me.Chkbarcode(200).value = vbChecked
    Else
        Me.Chkbarcode(200).value = Unchecked
    End If
    
    
    If rs("IsShowItemsBranch").value = True Then
        Me.Chkbarcode(202).value = vbChecked
    Else
        Me.Chkbarcode(202).value = Unchecked
    End If
     
    
    
    If rs("IsQrCodePrint").value = True Then
        Me.Chkbarcode(201).value = vbChecked
    Else
        Me.Chkbarcode(201).value = Unchecked
    End If
     
         
   
    If rs("IsElecWaterCont").value = True Then
        Me.Chkbarcode(204).value = vbChecked
    Else
        Me.Chkbarcode(204).value = Unchecked
    End If
     

   
    If rs("IsDogeMode").value = True Then
        Me.Chkbarcode(205).value = vbChecked
    Else
        Me.Chkbarcode(205).value = Unchecked
    End If
              
              
    If rs("IsMaintItemMode").value = True Then
        Me.Chkbarcode(206).value = vbChecked
    Else
        Me.Chkbarcode(206).value = Unchecked
    End If
              
              
                
              
    If rs("IsHiddenTransportInv").value = True Then
        Me.Chkbarcode(211).value = vbChecked
    Else
        Me.Chkbarcode(211).value = Unchecked
    End If
              
                
  
         
       
              
    If rs("IsHeaderPrint").value = True Then
        Me.Chkbarcode(207).value = vbChecked
    Else
        Me.Chkbarcode(207).value = Unchecked
    End If
              
              
         

     

    
            If rs("WorkWithBranchLogo").value = True Then
        Me.Chkbarcode(8).value = vbChecked
    Else
        Me.Chkbarcode(8).value = Unchecked
    End If
                If rs("WorkWithFirstInstallOnly").value = True Then
        Me.Chkbarcode(9).value = vbChecked
    Else
        Me.Chkbarcode(9).value = Unchecked
    End If
    
                If rs("CreateInsuranceAccountForCustomers").value = True Then
        Me.Chkbarcode(11).value = vbChecked
    Else
        Me.Chkbarcode(11).value = Unchecked
    End If
    
                    If rs("DecideItemName").value = True Then
        Me.Chkbarcode(12).value = vbChecked
    Else
        Me.Chkbarcode(12).value = Unchecked
    End If
    
                   If rs("DefaultIsCreditSales").value = True Then
        Me.Chkbarcode(13).value = vbChecked
    Else
        Me.Chkbarcode(13).value = Unchecked
    End If
    
                     If rs("DefaultIsCreditPurchase").value = True Then
        Me.Chkbarcode(46).value = vbChecked
    Else
        Me.Chkbarcode(46).value = Unchecked
    End If
    
    If rs("DefaultIsCreditPurchaseRet").value = True Then
        Me.Chkbarcode(137).value = vbChecked
    Else
        Me.Chkbarcode(137).value = Unchecked
    End If
    
    If rs("OpenAccountAqar").value = True Then
        Me.Chkbarcode(138).value = vbChecked
    Else
        Me.Chkbarcode(138).value = Unchecked
    End If
    If rs("InvoiceTransferJLTotal").value = True Then
        Me.Chkbarcode(139).value = vbChecked
    Else
        Me.Chkbarcode(139).value = Unchecked
    End If
    
  If rs("CarsRevenuePerOwner").value = True Then
        Me.Chkbarcode(146).value = vbChecked
    Else
        Me.Chkbarcode(146).value = Unchecked
    End If
      
      
    If rs("DontShowMoreDetailsCompItem").value = True Then
        Me.Chkbarcode(147).value = vbChecked
    Else
        Me.Chkbarcode(147).value = Unchecked
    End If

    If rs("CompilingBasedTable").value = True Then
        Me.Chkbarcode(153).value = vbChecked
    Else
        Me.Chkbarcode(153).value = Unchecked
    End If

 If rs("CanPartialpayment").value = True Then
        Me.Chkbarcode(154).value = vbChecked
    Else
        Me.Chkbarcode(154).value = Unchecked
    End If


 If rs("EndRentifPayed").value = True Then
        Me.Chkbarcode(155).value = vbChecked
    Else
        Me.Chkbarcode(155).value = Unchecked
    End If

 If rs("cantCahngeAkarinExpenses").value = True Then
        Me.Chkbarcode(156).value = vbChecked
    Else
        Me.Chkbarcode(156).value = Unchecked
    End If
    
    
 If rs("EmployeeSalaryBYBranch").value = True Then
        Me.Chkbarcode(157).value = vbChecked
    Else
        Me.Chkbarcode(157).value = Unchecked
    End If
    
    
 If rs("returnnotcreatvoucher").value = True Then
        Me.Chkbarcode(158).value = vbChecked
    Else
        Me.Chkbarcode(158).value = Unchecked
    End If
    
    
    
 If rs("OnlyOneCashingVchr").value = True Then
        Me.Chkbarcode(186).value = vbChecked
    Else
        Me.Chkbarcode(186).value = Unchecked
    End If
    
 If rs("CheckDateFormatCorrect").value = True Then
        Me.Chkbarcode(187).value = vbChecked
    Else
        Me.Chkbarcode(187).value = Unchecked
    End If
    
 If rs("CheckMobileFormatCorrect").value = True Then
        Me.Chkbarcode(190).value = vbChecked
    Else
        Me.Chkbarcode(190).value = Unchecked
    End If
        
                                  
 If rs("IsShowLensesDetails").value = True Then
        Me.Chkbarcode(191).value = vbChecked
    Else
        Me.Chkbarcode(191).value = Unchecked
    End If
                                  

    If Me.Chkbarcode(191).value = vbChecked Then
        rs("IsShowLensesDetails").value = 1
    ElseIf Me.Chkbarcode(191).value = vbUnchecked Then
        rs("IsShowLensesDetails").value = 0
    End If
                          

 
 If rs("CantRepetttransferNoforCashing").value = True Then
        Me.Chkbarcode(188).value = vbChecked
    Else
        Me.Chkbarcode(188).value = Unchecked
    End If
    
    
 If rs("DontDuplicateManulaNoInPurchase").value = True Then
        Me.Chkbarcode(189).value = vbChecked
    Else
        Me.Chkbarcode(189).value = Unchecked
    End If
    If rs("WaiverSetByContract").value = True Then
        Me.Chkbarcode(159).value = vbChecked
    Else
        Me.Chkbarcode(159).value = Unchecked
    End If
            

    If rs("IsGeometricProportions").value = True Then
        Me.Chkbarcode(160).value = vbChecked
    Else
        Me.Chkbarcode(160).value = Unchecked
    End If
            
'
                  
    If rs("IsSerialByUserTrans").value = True Then
        Me.Chkbarcode(162).value = vbChecked
    Else
        Me.Chkbarcode(162).value = Unchecked
    End If
                              
    If rs("IsSerialByUserVouch").value = True Then
        Me.Chkbarcode(163).value = vbChecked
    Else
        Me.Chkbarcode(163).value = Unchecked
    End If
                                          
                  

    If rs("AllowRepeateCar").value = True Then
        Me.Chkbarcode(164).value = vbChecked
    Else
        Me.Chkbarcode(164).value = Unchecked
    End If
    
                               


    If rs("traveDiscountFromCustomerDirect").value = True Then
        Me.Chkbarcode(148).value = vbChecked
    Else
        Me.Chkbarcode(148).value = Unchecked
    End If
        If rs("IsCustSalesManCashRelated").value = True Then
        Me.Chkbarcode(149).value = vbChecked
    Else
        Me.Chkbarcode(149).value = Unchecked
    End If
    
        If rs("showEmployeeAccountIntrip").value = True Then
        Me.Chkbarcode(150).value = vbChecked
    Else
        Me.Chkbarcode(150).value = Unchecked
    End If
              
              
        If rs("DUEDOCUMENTbyinstallDate").value = True Then
        Me.Chkbarcode(151).value = vbChecked
    Else
        Me.Chkbarcode(151).value = Unchecked
    End If
                            
        If rs("CanSkipPurchOrder").value = True Then
        Me.Chkbarcode(152).value = vbChecked
    Else
        Me.Chkbarcode(152).value = Unchecked
    End If
                            
                            
                            
          If rs("returnByBarCodeOnly").value = True Then
        Me.Chkbarcode(47).value = vbChecked
    Else
        Me.Chkbarcode(47).value = Unchecked
    End If
    
    
                  If rs("JLCodeBasedOnBranch").value = True Then
        Me.Chkbarcode(14).value = vbChecked
    Else
        Me.Chkbarcode(14).value = Unchecked
    End If
    
                 If rs("EmpNotExcceedDiscount").value = True Then
        Me.Chkbarcode(15).value = vbChecked
    Else
        Me.Chkbarcode(15).value = Unchecked
    End If
    
                    
                 If rs("BoxLossandIncreae").value = True Then
        Me.Chkbarcode(16).value = vbChecked
    Else
        Me.Chkbarcode(16).value = Unchecked
    End If
                        
                 If rs("attacheditemsisfree").value = True Then
        Me.Chkbarcode(17).value = vbChecked
    Else
        Me.Chkbarcode(17).value = Unchecked
    End If
    
                If rs("EnableCustomerAging").value = 0 Then
        Me.Chkbarcode(34).value = Unchecked
    Else
        Me.Chkbarcode(34).value = vbChecked
    End If
                         
                         
      If rs("showcostColorininvoice").value = True Then
        Me.Chkbarcode(18).value = vbChecked
    Else
        Me.Chkbarcode(18).value = Unchecked
    End If
                     
                     
        If rs("SubContactorHave3Account").value = True Then
        Me.Chkbarcode(19).value = vbChecked
    Else
        Me.Chkbarcode(19).value = Unchecked
    End If
    
                     
                If rs("ProjectEmployeeGV").value = True Then
        Me.Chkbarcode(20).value = vbChecked
    Else
        Me.Chkbarcode(20).value = Unchecked
    End If
       
       
       If rs("PursgaseWithoutDecimal").value = True Then
        Me.Chkbarcode(21).value = vbChecked
    Else
        Me.Chkbarcode(21).value = Unchecked
    End If
    
    
      
       If rs("workWithCustomerContract").value = True Then
        Me.Chkbarcode(22).value = vbChecked
    Else
        Me.Chkbarcode(22).value = Unchecked
    End If
    
      If rs("workWithvendorContract").value = True Then
        Me.Chkbarcode(25).value = vbChecked
    Else
        Me.Chkbarcode(25).value = Unchecked
    End If
    
      If rs("PoCreateVoucher").value = True Then
        Me.Chkbarcode(26).value = vbChecked
    Else
        Me.Chkbarcode(26).value = Unchecked
    End If
    
        If rs("poWithatotalQty").value = True Then
        Me.Chkbarcode(27).value = vbChecked
    Else
        Me.Chkbarcode(27).value = Unchecked
    End If
    
        If rs("DiscountSalesCreateVchr").value = True Then
        Me.Chkbarcode(28).value = vbChecked
    Else
        Me.Chkbarcode(28).value = Unchecked
    End If
    
    If rs("AllowCostPerStore").value = True Then
        Me.Chkbarcode(39).value = vbChecked
    Else
        Me.Chkbarcode(39).value = Unchecked
    End If
    
    If rs("AllowCostnNewShape").value = True Then
        Me.Chkbarcode(40).value = vbChecked
    Else
        Me.Chkbarcode(40).value = Unchecked
    End If
    
    If rs("AllowCostBySerial").value = True Then
        Me.Chkbarcode(41).value = vbChecked
    Else
        Me.Chkbarcode(41).value = Unchecked
    End If
    
    'AllowCostBySerial
    'AllowCostnNewShape
    'AllowCostPerStore
        
        If rs("PaymentDifferent").value = True Then
        Me.Chkbarcode(29).value = vbChecked
    Else
        Me.Chkbarcode(29).value = Unchecked
    End If
    
    
            If rs("PayrollOneAccount").value = True Then
        Me.Chkbarcode(30).value = vbChecked
    Else
        Me.Chkbarcode(30).value = Unchecked
    End If
    
    
    
                If rs("WorkWithItemsDetails").value = True Then
        Me.Chkbarcode(31).value = vbChecked
    Else
        Me.Chkbarcode(31).value = Unchecked
    End If
    
                If rs("FAAddtionCreateAccount").value = True Then
        Me.Chkbarcode(32).value = vbChecked
    Else
        Me.Chkbarcode(32).value = Unchecked
    End If
    
   If rs("Create2account4Supp").value = True Then
        Me.Chkbarcode(33).value = vbChecked
    Else
        Me.Chkbarcode(33).value = Unchecked
    End If
    
    '
        
    
       If rs("cancellAllApprove").value = True Then
        Me.Chkbarcode(23).value = vbChecked
    Else
        Me.Chkbarcode(23).value = Unchecked
    End If
        
       If rs("workwithticketAllocation").value = True Then
        Me.Chkbarcode(24).value = vbChecked
    Else
        Me.Chkbarcode(24).value = Unchecked
    End If
    
        
                    If rs("WorkWithGroupCode").value = True Then
        Me.Chkbarcode(10).value = vbChecked
    Else
        Me.Chkbarcode(10).value = Unchecked
    End If
    
    
       If rs("TradingPOS").value = True Then
        Me.Chkbarcode(2).value = vbChecked
    Else
        Me.Chkbarcode(2).value = Unchecked
    End If
    
    
    If rs("posshape2").value = True Then
        Me.Chkbarcode(87).value = vbChecked
    Else
        Me.Chkbarcode(87).value = Unchecked
    End If
    If rs("InsuranceOnOwner").value = True Then
        Me.Chkbarcode(88).value = vbChecked
    Else
        Me.Chkbarcode(88).value = Unchecked
    End If
    If rs("ServicesOnOwner").value = True Then
        Me.Chkbarcode(89).value = vbChecked
    Else
        Me.Chkbarcode(89).value = Unchecked
    End If
    If rs("DueComm").value = True Then
        Me.Chkbarcode(90).value = vbChecked
    Else
        Me.Chkbarcode(90).value = Unchecked
    End If
     If rs("DueWater").value = True Then
        Me.Chkbarcode(91).value = vbChecked
    Else
        Me.Chkbarcode(91).value = Unchecked
    End If
        If rs("DueElectr").value = True Then
        Me.Chkbarcode(92).value = vbChecked
    Else
        Me.Chkbarcode(92).value = Unchecked
    End If
    If rs("DueService").value = True Then
        Me.Chkbarcode(93).value = vbChecked
    Else
        Me.Chkbarcode(93).value = Unchecked
    End If
     If rs("CommissionOnOwner").value = True Then
        Me.Chkbarcode(94).value = vbChecked
    Else
        Me.Chkbarcode(94).value = Unchecked
    End If
    
    
     If rs("CommissionDue").value = True Then
        Me.Chkbarcode(95).value = vbChecked
    Else
        Me.Chkbarcode(95).value = Unchecked
    End If
    
    
    If rs("SupplierReciveGE").value = True Then
        Me.Chkbarcode(96).value = vbChecked
    Else
        Me.Chkbarcode(96).value = Unchecked
    End If
    
    If rs("BranchmustimSalary").value = True Then
        Me.Chkbarcode(97).value = vbChecked
    Else
        Me.Chkbarcode(97).value = Unchecked
    End If
    If rs("AllowSkipPayment").value = True Then
        Me.Chkbarcode(98).value = vbChecked
    Else
        Me.Chkbarcode(98).value = Unchecked
    End If
        If rs("AllowChangePriceApprove").value = True Then
        Me.Chkbarcode(99).value = vbChecked
    Else
        Me.Chkbarcode(99).value = Unchecked
    End If
    If rs("CreateJLVactionAratha").value = True Then
        Me.Chkbarcode(101).value = vbChecked
    Else
        Me.Chkbarcode(101).value = Unchecked
    End If
    If rs("PriceWithVAT").value = True Then
        Me.Chkbarcode(102).value = vbChecked
    Else
        Me.Chkbarcode(102).value = Unchecked
    End If
    If rs("AllowWorkCustomerPoints").value = True Then
        Me.Chkbarcode(103).value = vbChecked
    Else
        Me.Chkbarcode(103).value = Unchecked
    End If
       If rs("ProjectInvoiceAnalysisJL").value = True Then
        Me.Chkbarcode(104).value = vbChecked
    Else
        Me.Chkbarcode(104).value = Unchecked
    End If
    If rs("CustomerRecordNoIsnotManda").value = True Then
        Me.Chkbarcode(105).value = vbChecked
    Else
        Me.Chkbarcode(105).value = Unchecked
    End If
    If rs("DealingWithPrepayAccount").value = True Then
        Me.Chkbarcode(106).value = vbChecked
    Else
        Me.Chkbarcode(106).value = Unchecked
    End If
    
    If rs("CanChanegeLinkedSsalesnvoice").value = True Then
        Me.Chkbarcode(3).value = vbChecked
    Else
        Me.Chkbarcode(3).value = Unchecked
    End If
    
    
    If rs("NotAllowedCalcVata").value = True Then
        Me.Chkbarcode(107).value = vbChecked
    Else
        Me.Chkbarcode(107).value = Unchecked
    End If
    
      If rs("IssueVoucherWorkWithRemain").value = True Then
        Me.Chkbarcode(108).value = vbChecked
    Else
        Me.Chkbarcode(108).value = Unchecked
    End If
      If rs("TripDateInsertDefulat").value = True Then
        Me.Chkbarcode(109).value = vbChecked
    Else
        Me.Chkbarcode(109).value = Unchecked
    End If
    
    
      If rs("TripwithorderOnly").value = True Then
        Me.Chkbarcode(112).value = vbChecked
    Else
        Me.Chkbarcode(112).value = Unchecked
    End If
    
    If rs("AllowPriceWithWidth").value = True Then
        Me.Chkbarcode(113).value = vbChecked
    Else
        Me.Chkbarcode(113).value = Unchecked
    End If
    If rs("LinkCustomerWithCars").value = True Then
        Me.Chkbarcode(114).value = vbChecked
    Else
        Me.Chkbarcode(114).value = Unchecked
    End If
    If rs("AllowEditCashingLinkProj").value = True Then
        Me.Chkbarcode(115).value = vbChecked
    Else
        Me.Chkbarcode(115).value = Unchecked
    End If
    If rs("TransBillPriceByGrid").value = True Then
        Me.Chkbarcode(116).value = vbChecked
    Else
        Me.Chkbarcode(116).value = Unchecked
    End If
    If rs("NoCreatJLInRentContract").value = True Then
        Me.Chkbarcode(117).value = vbChecked
    Else
        Me.Chkbarcode(117).value = Unchecked
    End If
        If rs("OpenVATAccountOwner").value = True Then
        Me.Chkbarcode(118).value = vbChecked
    Else
        Me.Chkbarcode(118).value = Unchecked
    End If
    If rs("CreateJLEmpCommissions").value = True Then
        Me.Chkbarcode(119).value = vbChecked
    Else
        Me.Chkbarcode(119).value = Unchecked
    End If
    If rs("TypeContractAutoFromIqar").value = True Then
        Me.Chkbarcode(120).value = vbChecked
    Else
        Me.Chkbarcode(120).value = Unchecked
    End If
    
    If rs("AllowRepeatInvoiceNo").value = True Then
        Me.Chkbarcode(121).value = vbChecked
    Else
        Me.Chkbarcode(121).value = Unchecked
    End If
    
    If rs("AllowReturnFIFO").value = True Then
        Me.Chkbarcode(122).value = vbChecked
    Else
        Me.Chkbarcode(122).value = Unchecked
    End If
        If rs("AllowDiscountAllowedFIFO").value = True Then
        Me.Chkbarcode(123).value = vbChecked
    Else
        Me.Chkbarcode(123).value = Unchecked
    End If
    If rs("AllowJLManualFIFO").value = True Then
        Me.Chkbarcode(124).value = vbChecked
    Else
        Me.Chkbarcode(124).value = Unchecked
    End If
    
    If rs("IsMergeVat").value = True Then
        Me.Chkbarcode(161).value = vbChecked
    Else
        Me.Chkbarcode(161).value = Unchecked
    End If
    

 
    
    If rs("ShowBalanceOfEmpInSalary").value = True Then
        Me.Chkbarcode(125).value = vbChecked
    Else
        Me.Chkbarcode(125).value = Unchecked
    End If
       If rs("PaymentIntoAccouStat").value = True Then
        Me.Chkbarcode(126).value = vbChecked
    Else
        Me.Chkbarcode(126).value = Unchecked
    End If
    If rs("AllowEditInvoiceNoticeDiscount").value = True Then
        Me.Chkbarcode(127).value = vbChecked
    Else
        Me.Chkbarcode(127).value = Unchecked
    End If
    If rs("AllowEditInvoiceOfReturn").value = True Then
        Me.Chkbarcode(128).value = vbChecked
    Else
        Me.Chkbarcode(128).value = Unchecked
    End If
    
    
    If rs("ProvisionsByManagement").value = True Then
        Me.Chkbarcode(129).value = vbChecked
    Else
        Me.Chkbarcode(129).value = Unchecked
    End If
     
     
    If rs("ProvisionsByıEQuipments").value = True Then
        Me.Chkbarcode(165).value = vbChecked
    Else
        Me.Chkbarcode(165).value = Unchecked
    End If
    
        If rs("ReturnSAlesByBarcode").value = True Then
        Me.Chkbarcode(166).value = vbChecked
    Else
        Me.Chkbarcode(166).value = Unchecked
    End If
    
        If rs("DontDistributeLegalACC").value = True Then
        Me.Chkbarcode(167).value = vbChecked
    Else
        Me.Chkbarcode(167).value = Unchecked
    End If
    
    If rs("CreatePayOrderSales").value = True Then
        Me.Chkbarcode(168).value = vbChecked
    Else
        Me.Chkbarcode(168).value = Unchecked
    End If
        
    If rs("IsBarCodeByUnit").value = True Then
        Me.Chkbarcode(169).value = vbChecked
    Else
        Me.Chkbarcode(169).value = Unchecked
    End If
    
    If rs("TripnotUploadExpenses").value = True Then
        Me.Chkbarcode(170).value = vbChecked
    Else
        Me.Chkbarcode(170).value = Unchecked
    End If
    
    
    If rs("ExpensesByQtyOnly").value = True Then
        Me.Chkbarcode(171).value = vbChecked
    Else
        Me.Chkbarcode(171).value = Unchecked
    End If
        
    If rs("DiscountByQtyOnly").value = True Then
        Me.Chkbarcode(220).value = vbChecked
    Else
        Me.Chkbarcode(220).value = Unchecked
    End If
         

    If rs("IsTransferByCode").value = True Then
        Me.Chkbarcode(222).value = vbChecked
    Else
        Me.Chkbarcode(222).value = Unchecked
    End If
                  
       
    If rs("ZacatHandW").value = True Then
        Me.Chkbarcode(221).value = vbChecked
    Else
        Me.Chkbarcode(221).value = Unchecked
    End If
         

    If rs("ShowPrinterDialoge").value = True Then
        Me.Chkbarcode(172).value = vbChecked
    Else
        Me.Chkbarcode(172).value = Unchecked
    End If
        
        
            If rs("AllowDynamicAutoSus").value = True Then
        Me.Chkbarcode(173).value = vbChecked
    Else
        Me.Chkbarcode(173).value = Unchecked
    End If
        
        
    If rs("AllowUnbalncedByBranchAccount").value = True Then
        Me.Chkbarcode(174).value = vbChecked
    Else
        Me.Chkbarcode(174).value = Unchecked
    End If
    
    If rs("SortInvoiceByEntry").value = True Then
        Me.Chkbarcode(175).value = vbChecked
    Else
        Me.Chkbarcode(175).value = Unchecked
    End If
    
    If rs("CostProductOrderByOut").value = True Then
        Me.Chkbarcode(176).value = vbChecked
    Else
        Me.Chkbarcode(176).value = Unchecked
    End If
        
    
    If rs("CostByProduction").value = True Then
        Me.Chkbarcode(179).value = vbChecked
    Else
        Me.Chkbarcode(179).value = Unchecked
    End If

    If rs("MaintOrderCantRepeatSales").value = True Then
        Me.Chkbarcode(180).value = vbChecked
    Else
        Me.Chkbarcode(180).value = Unchecked
    End If
   
                 
    If rs("MaintOrderCantRepeatBillBuy").value = True Then
        Me.Chkbarcode(181).value = vbChecked
    Else
        Me.Chkbarcode(181).value = Unchecked
    End If
   
   
     If rs("TripRevenueAuto").value = True Then
        Me.Chkbarcode(184).value = vbChecked
    Else
        Me.Chkbarcode(184).value = Unchecked
    End If
   
  
     If rs("IsByNewCoding").value = True Then
        Me.Chkbarcode(185).value = vbChecked
    Else
        Me.Chkbarcode(185).value = Unchecked
    End If
   
 
                 
                 
    If rs("PaymentMethLaterCompItem").value = True Then
        Me.Chkbarcode(182).value = vbChecked
    Else
        Me.Chkbarcode(182).value = Unchecked
    End If
   
        
       
    If rs("ShowBalanceCustInv").value = True Then
        Me.Chkbarcode(183).value = vbChecked
    Else
        Me.Chkbarcode(183).value = Unchecked
    End If
    
        
                  
   
    
    If rs("TransferNotInvItemDef").value = True Then
        Me.Chkbarcode(177).value = vbChecked
    Else
        Me.Chkbarcode(177).value = Unchecked
    End If
        
    
    If rs("CustMobNoMandatory").value = True Then
        Me.Chkbarcode(178).value = vbChecked
    Else
        Me.Chkbarcode(178).value = Unchecked
    End If
        


    If rs("CustVatNoMandatory").value = True Then
        Me.Chkbarcode(214).value = vbChecked
    Else
        Me.Chkbarcode(214).value = Unchecked
    End If
                                    



    If rs("AllowScInterface2").value = True Then
        Me.Chkbarcode(215).value = vbChecked
    Else
        Me.Chkbarcode(215).value = Unchecked
    End If
               

   
     
    If rs("CloseMovingVchrinSales").value = True Then
        Me.Chkbarcode(130).value = vbChecked
    Else
        Me.Chkbarcode(130).value = Unchecked
    End If
          
   
    If rs("IsMultiItemsInCompItem").value = True Then
        Me.Chkbarcode(132).value = vbChecked
    Else
        Me.Chkbarcode(132).value = Unchecked
    End If
           
    
          
     If rs("CantChangeSalesPerson").value = 1 Then
        Me.Chkbarcode(131).value = vbChecked
    Else
        Me.Chkbarcode(131).value = Unchecked
    End If
                    
                    
     If rs("BatchCreateManyworkOrder").value = 1 Then
        Me.Chkbarcode(133).value = vbChecked
    Else
        Me.Chkbarcode(133).value = Unchecked
    End If
     If rs("LinkSupplerWithItem").value = 1 Then
        Me.Chkbarcode(134).value = vbChecked
    Else
        Me.Chkbarcode(134).value = Unchecked
    End If
   If rs("ShowOnlyItemsOfSales").value = 1 Then
        Me.Chkbarcode(135).value = vbChecked
    Else
        Me.Chkbarcode(135).value = Unchecked
    End If
                                        
                                        
   If rs("PrintInvoiceByBranch").value = 1 Then
        Me.Chkbarcode(136).value = vbChecked
    Else
        Me.Chkbarcode(136).value = Unchecked
    End If
    
    
   If rs("GeneralVoucherCreateSalesGE").value = 1 Then
        Me.Chkbarcode(140).value = vbChecked
    Else
        Me.Chkbarcode(140).value = Unchecked
    End If
        
   If rs("SalesNotCreateGe").value = 1 Then
        Me.Chkbarcode(141).value = vbChecked
    Else
        Me.Chkbarcode(141).value = Unchecked
    End If
    
        
    If rs("CanChanegeLinkedPurcahsenvoice").value = True Then
        Me.Chkbarcode(77).value = vbChecked
    Else
        Me.Chkbarcode(77).value = Unchecked
    End If
    If rs("AllowProductOrderOne").value = True Then
        Me.Chkbarcode(78).value = vbChecked
    Else
        Me.Chkbarcode(78).value = Unchecked
    End If
    If rs("SalaryJLByManagement").value = True Then
        Me.Chkbarcode(79).value = vbChecked
    Else
        Me.Chkbarcode(79).value = Unchecked
    End If
    If rs("AllowGoodPerfAccount").value = True Then
        Me.Chkbarcode(80).value = vbChecked
    Else
        Me.Chkbarcode(80).value = Unchecked
    End If
     If rs("AllowAnalyticJL").value = True Then
        Me.Chkbarcode(83).value = vbChecked
    Else
        Me.Chkbarcode(83).value = Unchecked
    End If
    If rs("AllowSaveTripWithoutExpen").value = True Then
        Me.Chkbarcode(84).value = vbChecked
    Else
        Me.Chkbarcode(84).value = Unchecked
    End If
    

    If rs("CreateEntryManual").value = True Then
        Me.Chkbarcode(110).value = vbChecked
    Else
        Me.Chkbarcode(110).value = Unchecked
    End If
    
    If rs("chkAllowEditPaymentCont").value = True Then
        Me.Chkbarcode(111).value = vbChecked
    Else
        Me.Chkbarcode(111).value = Unchecked
    End If
    
    If rs("CustCreat4Acc").value = True Then
        Me.Chkbarcode(208).value = vbChecked
    Else
        Me.Chkbarcode(208).value = Unchecked
    End If
        
    If rs("SuppCreat4Acc").value = True Then
        Me.Chkbarcode(209).value = vbChecked
    Else
        Me.Chkbarcode(209).value = Unchecked
    End If
    
    
    If rs("CreateEntryBillItems").value = True Then
        Me.Chkbarcode(210).value = vbChecked
    Else
        Me.Chkbarcode(210).value = Unchecked
    End If
    
    If rs("SAVEMAINTENANCEJOBWITHORDERORPLANONLY").value = True Then
        Me.Chkbarcode(100).value = vbChecked
    Else
        Me.Chkbarcode(100).value = Unchecked
    End If
     
     
     
        If rs("SendToAprovedSalesBill").value = True Then
        Me.Chkbarcode(85).value = vbChecked
    Else
        Me.Chkbarcode(85).value = Unchecked
    End If
    If rs("SalaryJLByAnalyEqup").value = True Then
        Me.Chkbarcode(86).value = vbChecked
    Else
        Me.Chkbarcode(86).value = Unchecked
    End If
    
    If rs("ManualSalesInvoiceMust").value = True Then
        Me.Chkbarcode(81).value = vbChecked
    Else
        Me.Chkbarcode(81).value = Unchecked
    End If
    If rs("AllItemInVAT").value = True Then
        Me.Chkbarcode(82).value = vbChecked
    Else
        Me.Chkbarcode(82).value = Unchecked
    End If
    
       If rs("AnalyticPaymentVouchr").value = True Then
        Me.Chkbarcode(42).value = vbChecked
    Else
        Me.Chkbarcode(42).value = Unchecked
    End If
        
               If rs("ShowDriverOnly").value = True Then
        Me.Chkbarcode(43).value = vbChecked
    Else
        Me.Chkbarcode(43).value = Unchecked
    End If
    
    If rs("AllowSalesMultyPayed").value = True Then
        Me.Chkbarcode(44).value = vbChecked
    Else
        Me.Chkbarcode(44).value = Unchecked
    End If
    If rs("AllowAccountMultyPayed").value = True Then
        Me.Chkbarcode(59).value = vbChecked
    Else
        Me.Chkbarcode(59).value = Unchecked
    End If
       If rs("AllowPurchasesMultyPayed").value = True Then
        Me.Chkbarcode(50).value = vbChecked
    Else
        Me.Chkbarcode(50).value = Unchecked
    End If
      If rs("CashCustomerNameMustenter").value = True Then
        Me.Chkbarcode(45).value = vbChecked
    Else
        Me.Chkbarcode(45).value = Unchecked
    End If
    If rs("AllowCommtionJEFromValueVisa").value = True Then
        Me.Chkbarcode(48).value = vbChecked
    Else
        Me.Chkbarcode(48).value = Unchecked
    End If
    
        If rs("AllowWorkWithArea").value = True Then
        Me.Chkbarcode(49).value = vbChecked
    Else
        Me.Chkbarcode(49).value = Unchecked
    End If
    If rs("AllowAcceleratepayment").value = True Then
        Me.Chkbarcode(51).value = vbChecked
    Else
        Me.Chkbarcode(51).value = Unchecked
    End If
        If rs("AllowExperDateFIFO").value = True Then
        Me.Chkbarcode(52).value = vbChecked
    Else
        Me.Chkbarcode(52).value = Unchecked
    End If
     If rs("AllowProjectBill2Serial").value = True Then
        Me.Chkbarcode(53).value = vbChecked
    Else
        Me.Chkbarcode(53).value = Unchecked
    End If
      If rs("ViewAccountsbyBranch").value = True Then
        Me.Chkbarcode(54).value = vbChecked
    Else
        Me.Chkbarcode(54).value = Unchecked
    End If
     If rs("AllowEditeAccounts").value = True Then
        Me.Chkbarcode(55).value = vbChecked
    Else
        Me.Chkbarcode(55).value = Unchecked
    End If
     If rs("ProjectUnderImplemen").value = True Then
        Me.Chkbarcode(56).value = vbChecked
    Else
        Me.Chkbarcode(56).value = Unchecked
    End If
    If rs("AllowHideAssest").value = True Then
        Me.Chkbarcode(57).value = vbChecked
    Else
        Me.Chkbarcode(57).value = Unchecked
    End If
    If rs("LockSalary").value = True Then
        Me.Chkbarcode(58).value = vbChecked
    Else
        Me.Chkbarcode(58).value = Unchecked
    End If
      If rs("updatecashvchrifdeposite").value = True Then
        Me.Chkbarcode(4).value = vbChecked
    Else
        Me.Chkbarcode(4).value = Unchecked
    End If
    
    
    If rs("Revenueowed").value = True Then
        Me.Chkbarcode(5).value = vbChecked
    Else
        Me.Chkbarcode(5).value = Unchecked
    End If
    
        If rs("AllowupdateJobStatus").value = True Then
        Me.Chkbarcode(6).value = vbChecked
    Else
        Me.Chkbarcode(6).value = Unchecked
    End If
    
    
    If rs("OpeningEmployeeShowAll").value = True Then
          Me.Chkbarcode(60).value = vbChecked
      Else
          Me.Chkbarcode(60).value = Unchecked
      End If
    If rs("SellOrderBalance").value = True Then
        Me.Chkbarcode(61).value = vbChecked
    Else
        Me.Chkbarcode(61).value = Unchecked
    End If
      If rs("EndServiceMore5Year").value = True Then
        Me.Chkbarcode(62).value = vbChecked
    Else
        Me.Chkbarcode(62).value = Unchecked
    End If
    
          If rs("VacstionShowOldSalaries").value = True Then
        Me.Chkbarcode(75).value = vbChecked
    Else
        Me.Chkbarcode(75).value = Unchecked
    End If
    
          If rs("AllowReturnWithoutCost").value = True Then
        Me.Chkbarcode(76).value = vbChecked
    Else
        Me.Chkbarcode(76).value = Unchecked
    End If
    
     If rs("ShowItemByCustomer").value = True Then
        Me.Chkbarcode(63).value = vbChecked
    Else
        Me.Chkbarcode(63).value = Unchecked
    End If
    If rs("RawMaterMix").value = True Then
        Me.Chkbarcode(64).value = vbChecked
    Else
        Me.Chkbarcode(64).value = Unchecked
    End If
    
    If rs("RawMaterMix2").value = True Then
        Me.Chkbarcode(142).value = vbChecked
    Else
        Me.Chkbarcode(142).value = Unchecked
    End If
    
    
    If rs("DontCreateOut").value = True Then
        Me.Chkbarcode(143).value = vbChecked
    Else
        Me.Chkbarcode(143).value = Unchecked
    End If
        
        
        
    If rs("DontCreateOut2").value = True Then
        Me.Chkbarcode(144).value = vbChecked
    Else
        Me.Chkbarcode(144).value = Unchecked
    End If
        
        
        
    If rs("InsertItemManualOut").value = True Then
        Me.Chkbarcode(145).value = vbChecked
    Else
        Me.Chkbarcode(145).value = Unchecked
    End If
        
    
    If rs("LinkUsersWithPayment").value = True Then
        Me.Chkbarcode(65).value = vbChecked
    Else
        Me.Chkbarcode(65).value = Unchecked
    End If
    If rs("VATNoAccordActivity").value = True Then
        Me.Chkbarcode(66).value = vbChecked
    Else
        Me.Chkbarcode(66).value = Unchecked
    End If
        If rs("NotCrtResvVouchProjects").value = True Then
        Me.Chkbarcode(67).value = vbChecked
    Else
        Me.Chkbarcode(67).value = Unchecked
    End If
      If rs("SalesTrustsAffectVedorCode").value = True Then
        Me.Chkbarcode(68).value = vbChecked
    Else
        Me.Chkbarcode(68).value = Unchecked
    End If
     If rs("AllowNoRoudProjectInvoices").value = True Then
        Me.Chkbarcode(69).value = vbChecked
    Else
        Me.Chkbarcode(69).value = Unchecked
    End If
    If rs("ProductionRawMaterMix").value = True Then
        Me.Chkbarcode(70).value = vbChecked
    Else
        Me.Chkbarcode(70).value = Unchecked
    End If
    If rs("AllowLastPrice").value = True Then
        Me.Chkbarcode(71).value = vbChecked
    Else
        Me.Chkbarcode(71).value = Unchecked
    End If
    
     If rs("AllowItemByRow").value = True Then
        Me.Chkbarcode(72).value = vbChecked
    Else
        Me.Chkbarcode(72).value = Unchecked
    End If
      If rs("AllowChangManualQtyMix").value = True Then
        Me.Chkbarcode(73).value = vbChecked
    Else
        Me.Chkbarcode(73).value = Unchecked
    End If
    If rs("AccountAccordingCash").value = True Then
        Me.Chkbarcode(74).value = vbChecked
    Else
        Me.Chkbarcode(74).value = Unchecked
    End If
    
       If rs("AllowTowShift").value = True Then
        Me.Chkbarcode(36).value = vbChecked
    Else
        Me.Chkbarcode(36).value = Unchecked
    End If
    
    
       If rs("AllowItemsShortName").value = True Then
        Me.Chkbarcode(37).value = vbChecked
    Else
        Me.Chkbarcode(37).value = Unchecked
    End If
    
    
    
      If rs("DuplicateitemsNames").value = True Then
        Me.Chkbarcode(1).value = vbChecked
    Else
        Me.Chkbarcode(1).value = Unchecked
    End If
    
        
            If rs("CostStarting").value = True Then
        Me.ChkCostStarting.value = vbChecked
    Else
        Me.ChkCostStarting.value = Unchecked
    End If
    
                If rs("CostStartingGard").value = True Then
        Me.chkStore(7).value = vbChecked
    Else
        Me.chkStore(7).value = Unchecked
    End If
    
                    If rs("TreatUncountedItemsAsZeroQty").value = True Then
        Me.chkStore(8).value = vbChecked
    Else
        Me.chkStore(8).value = Unchecked
    End If
     
     

        
        
            If rs("chkuserCode").value = True Then
        Me.chkuserCode.value = vbChecked
    Else
        Me.chkuserCode.value = Unchecked
    End If
    
         
                     If rs("Itemsattachedzero").value = True Then
        Me.ChkItemsattachedzero.value = vbChecked
    Else
        Me.ChkItemsattachedzero.value = Unchecked
    End If
    
    
         
         
    If rs("autoIssueVoucher").value = True Then
        Me.ChKautoIssueVoucher.value = vbChecked
    Else
        Me.ChKautoIssueVoucher.value = Unchecked
    End If
       
    If rs("MonthIs30days").value = True Then
        Me.chkMonthIs30days.value = vbChecked
    Else
        Me.chkMonthIs30days.value = Unchecked
    End If
        
    If rs("autoReseiveVoucher").value = True Then
        Me.ChKautoReseiveVoucher.value = vbChecked
    Else
        Me.ChKautoReseiveVoucher.value = Unchecked
    End If
        
    If rs("itemsWorkWithColor").value = True Then
        Me.ChkitemsWorkWithColor.value = vbChecked
    Else
        Me.ChkitemsWorkWithColor.value = Unchecked
    End If
        
    If rs("itemsWorkWithDates").value = True Then
        Me.ChkitemsWorkWithDates.value = vbChecked
    Else
        Me.ChkitemsWorkWithDates.value = Unchecked
    End If
        
    If rs("itemsWorkWithClass").value = True Then
        Me.ChkitemsWorkWithClass.value = vbChecked
    Else
        Me.ChkitemsWorkWithClass.value = Unchecked
    End If
 
    If rs("Arrows_group").value = False Then
        OptArrowBranch.value = True
    Else
        OptArrowGroup.value = True
    End If
        
'    If rs("opt_group").value = True Then
'        opt_group.value = True
'        Frame5.Visible = True
'    End If
        
'    If rs("Opt_Inventory_create_account").value = 1 Then
'        Opt_Inventory_create_account.value = True
'    End If
        
'    If rs("opt_inv_and_branch_create_account").value = 1 Then
'        opt_inv_and_branch_create_account.value = True
'    End If

    Me.TxtPriceDigts.text = IIf(IsNull(rs("CurrencyDigts").value), 0, rs("CurrencyDigts").value)
    Me.TxtPriceDigtsInst.text = IIf(IsNull(rs("PriceDigtsInst").value), 0, rs("PriceDigtsInst").value)
    TxtEmpSalaryDigts.text = IIf(IsNull(rs("EmpSalaryDigts").value), 0, rs("EmpSalaryDigts").value)
    Me.TxtEmpComponentDigts.text = IIf(IsNull(rs("EmpComponentDigts").value), 0, rs("EmpComponentDigts").value)
   ' Me.TxtCustNoBooking.Text = IIf(IsNull(rs("CustNoBooking").value), 0, rs("CustNoBooking").value)
    Me.TxtIndirectCostPercentage.text = IIf(IsNull(rs("IndirectCostPercentage").value), 0, rs("IndirectCostPercentage").value)
    
    Me.TxtData(1).text = IIf(IsNull(rs("StoreDigit").value), 1, rs("StoreDigit").value)
    Me.TxtData(0).text = IIf(IsNull(rs("BranchDigit").value), 1, rs("BranchDigit").value)
    

    Me.TxtQtyDigts.text = IIf(IsNull(rs("QtyDigts").value), 0, rs("QtyDigts").value)
    Me.TxtData(2).text = IIf(IsNull(rs("Ked_digit").value), 0, rs("Ked_digit").value)
    Me.txt_ACCOUNT_digit.text = IIf(IsNull(rs("Count_ACCOUNT_digit").value), 0, rs("Count_ACCOUNT_digit").value)

    
    Me.txtLimitDefaultCredit.text = IIf(IsNull(rs("LimitDefaultCredit").value), 0, rs("LimitDefaultCredit").value)
    Me.txtLimitDefaultCreditDays.text = IIf(IsNull(rs("LimitDefaultCreditDays").value), 0, rs("LimitDefaultCreditDays").value)
    
    Me.TxtSaleDiscount1.text = IIf(IsNull(rs("SaleDiscount1").value), 0, rs("SaleDiscount1").value)
    Me.TxtSaleDiscount2.text = IIf(IsNull(rs("SaleDiscount2").value), 0, rs("SaleDiscount2").value)
    Me.TxtSaleDiscount3.text = IIf(IsNull(rs("SaleDiscount3").value), 0, rs("SaleDiscount3").value)
    Me.TxtSaleDiscount4.text = IIf(IsNull(rs("SaleDiscount4").value), 0, rs("SaleDiscount4").value)

    AskOption = GetSetting(StrAppRegPath, "View_Type", "ShowMe", False)

    If AskOption = False Then
        ChkAsk.value = Checked
    Else
        ChkAsk.value = Unchecked
    End If

    AskOption = GetSetting(StrAppRegPath, "View_Type", "ShowRequest", True)

    If AskOption = False Then
        ChkShow.value = Unchecked
    Else
        ChkShow.value = Checked
    End If

    AskOption = GetSetting(StrAppRegPath, "View_Type", "ShowPayment", True)

    If AskOption = False Then
        ChkDelayVal.value = Unchecked
    Else
        ChkDelayVal.value = Checked
    End If
    
    Askinterval = GetSetting(StrAppRegPath, "Setting", "INTERVAL_ShowPayment", "")

    If Askinterval = "D" Then
        Combo1.ListIndex = 0
    ElseIf Askinterval = "m" Then
        Combo1.ListIndex = 1
    ElseIf Askinterval = "yyyy" Then
        Combo1.ListIndex = 2
    Else
        Combo1.ListIndex = -1
    End If
    
    Askcount = GetSetting(StrAppRegPath, "Setting", "count_ShowPayment", 0)
        
    Text1.text = Askcount
    
    
   Askcount = GetSetting(StrAppRegPath, "Setting", "CountAlarmMinutes", 5)
        
    Text14.text = Askcount
     
     
    
   '****************************************************
     AskOption = GetSetting(StrAppRegPath, "View_Type", "RentInstallments", True)

    If AskOption = False Then
        chkRentInstallments.value = Unchecked
    Else
        chkRentInstallments.value = Checked
    End If
    
    Askinterval = GetSetting(StrAppRegPath, "Setting", "INTERVAL_RentInstallments", "")

    If Askinterval = "D" Then
        Combo10.ListIndex = 0
    ElseIf Askinterval = "M" Then
        Combo10.ListIndex = 1
    ElseIf Askinterval = "YYYY" Then
        Combo10.ListIndex = 2
    Else
        Combo10.ListIndex = -1
    End If
    
    Askcount = GetSetting(StrAppRegPath, "Setting", "Count_RentInstallments", 0)
        
    Text8.text = Askcount
       '****************************************************

    LoadData
    
    
    AskOption = GetSetting(StrAppRegPath, "View_Type", "InstallmentMustPayed", True)

    If AskOption = False Then
        ChkInstallmentMustPayed.value = Unchecked
    Else
        ChkInstallmentMustPayed.value = Checked
    End If

    AskOption = GetSetting(StrAppRegPath, "View_Type", "ShowProjectsAlarm1", True)

    If AskOption = False Then
        ChKProjectsAlarm1.value = Unchecked
    Else
        ChKProjectsAlarm1.value = Checked
    End If
    
    AskOption = GetSetting(StrAppRegPath, "View_Type", "ShowProjectsAlarm2", True)

    If AskOption = False Then
        ChKProjectsAlarm2.value = Unchecked
    Else
        ChKProjectsAlarm2.value = Checked
    End If

    Askinterval = GetSetting(StrAppRegPath, "Setting", "INTERVAL_InstallmentMustPayed", "")

    If Askinterval = "D" Then
        Combo2.ListIndex = 0
    ElseIf Askinterval = "M" Then
        Combo2.ListIndex = 1
    ElseIf Askinterval = "YYYY" Then
        Combo2.ListIndex = 2
    Else
        Combo2.ListIndex = -1
    End If

    Askcount = GetSetting(StrAppRegPath, "Setting", "count_InstallmentMustPayed", 0)
    Text2.text = Askcount

    AskOption = GetSetting(StrAppRegPath, "View_Type", "ExpireEkama", "")

    If AskOption = False Then
        ChkExpireEkama.value = Unchecked
    Else
        ChkExpireEkama.value = Checked
    End If

    Askinterval = GetSetting(StrAppRegPath, "Setting", "INTERVAL_ExpireEkama", "")

    If Askinterval = "D" Then
        Combo3.ListIndex = 0
    ElseIf Askinterval = "m" Then
        Combo3.ListIndex = 1
    ElseIf Askinterval = "yyyy" Then
        Combo3.ListIndex = 2
    Else
        Combo3.ListIndex = -1
    End If

    Askcount = GetSetting(StrAppRegPath, "Setting", "count_ExpireEkama", 0)
    Text3.text = Askcount
'****************************************************
    AskOption = GetSetting(StrAppRegPath, "View_Type", "LC", "")

    If AskOption = False Then
        CheckLC.value = Unchecked
    Else
        CheckLC.value = Checked
    End If

    Askinterval = GetSetting(StrAppRegPath, "Setting", "INTERVAL_LC", "")

    If Askinterval = "D" Then
        Combo14.ListIndex = 0
    ElseIf Askinterval = "m" Then
        Combo14.ListIndex = 1
    ElseIf Askinterval = "yyyy" Then
        Combo14.ListIndex = 2
    Else
        Combo14.ListIndex = -1
    End If

    Askcount = GetSetting(StrAppRegPath, "Setting", "count_LC", 0)
    Text7.text = Askcount
'******************************************************


 Askcount = GetSetting(StrAppRegPath, "Setting", "ReportZoom", 100)
        
    TxtZoom.text = Askcount
  
  
    '    Askcount = GetSetting(StrAppRegPath, "Setting", "Count_Ked_digit", 0)
    '      TXTData(2).text = Askcount
        
    '   Askcount = GetSetting(StrAppRegPath, "Setting", "COUNT_ACCOUNT_digit", 0)
    '      txt_ACCOUNT_digit.text = Askcount
        
    '%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
    AskOption = GetSetting(StrAppRegPath, "View_Type", "ExpireLicense", "")

    If AskOption = False Then
        ChkExpireLicense.value = Unchecked
    Else
        ChkExpireLicense.value = Checked
    End If

    Askinterval = GetSetting(StrAppRegPath, "Setting", "INTERVAL_ExpireLicense", "")

    If Askinterval = "D" Then
        Combo7.ListIndex = 0
    ElseIf Askinterval = "m" Then
        Combo7.ListIndex = 1
    ElseIf Askinterval = "yyyy" Then
        Combo7.ListIndex = 2
    Else
        Combo7.ListIndex = -1
    End If

    Askcount = GetSetting(StrAppRegPath, "Setting", "count_ExpireLicense", 0)
    Text11.text = Askcount
        
    '%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
    AskOption = GetSetting(StrAppRegPath, "View_Type", "ExpireInsurance", "")

    If AskOption = False Then
        ChkExpireInsurance.value = Unchecked
    Else
        ChkExpireInsurance.value = Checked
    End If

    Askinterval = GetSetting(StrAppRegPath, "Setting", "INTERVAL_ExpireInsurance", "")

    If Askinterval = "D" Then
        Combo8.ListIndex = 0
    ElseIf Askinterval = "m" Then
        Combo8.ListIndex = 1
    ElseIf Askinterval = "yyyy" Then
        Combo8.ListIndex = 2
    Else
        Combo8.ListIndex = -1
    End If

    Askcount = GetSetting(StrAppRegPath, "Setting", "count_ExpireInsurance", 0)
    Text12.text = Askcount
        
    '%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
    AskOption = GetSetting(StrAppRegPath, "View_Type", "ExpireTest", "")

    If AskOption = False Then
        ChkExpireTest.value = Unchecked
    Else
        ChkExpireTest.value = Checked
    End If

    Askinterval = GetSetting(StrAppRegPath, "Setting", "INTERVAL_ExpireTest", "")

    If Askinterval = "D" Then
        Combo9.ListIndex = 0
    ElseIf Askinterval = "m" Then
        Combo9.ListIndex = 1
    ElseIf Askinterval = "yyyy" Then
        Combo9.ListIndex = 2
    Else
        Combo9.ListIndex = -1
    End If

    Askcount = GetSetting(StrAppRegPath, "Setting", "count_ExpireTest", 0)
    Text13.text = Askcount
        
    '%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
       
    AskOption = GetSetting(StrAppRegPath, "View_Type", "ExpireLicence", True)

    If AskOption = False Then
        ChkExpireLicence.value = Unchecked
    Else
        ChkExpireLicence.value = Checked
    End If

    Askinterval = GetSetting(StrAppRegPath, "Setting", "INTERVAL_ExpireLicence", "")

    If Askinterval = "D" Then
        Combo4.ListIndex = 0
    ElseIf Askinterval = "m" Then
        Combo4.ListIndex = 1
    ElseIf Askinterval = "yyyy" Then
        Combo4.ListIndex = 2
    Else
        Combo4.ListIndex = -1
    End If

    Askcount = GetSetting(StrAppRegPath, "Setting", "count_ExpireLicence", 0)
    Text4.text = Askcount
    
    AskOption = GetSetting(StrAppRegPath, "View_Type", "Expirepas", True)
    
    If AskOption = False Then
        ChkExpirepas.value = Unchecked
    Else
        ChkExpirepas.value = Checked
    End If

    Askinterval = GetSetting(StrAppRegPath, "Setting", "INTERVAL_Expirepas", "")

    If Askinterval = "D" Then
        Combo5.ListIndex = 0
    ElseIf Askinterval = "m" Then
        Combo5.ListIndex = 1
    ElseIf Askinterval = "yyyy" Then
        Combo5.ListIndex = 2
    Else
        Combo5.ListIndex = -1
    End If

    Askcount = GetSetting(StrAppRegPath, "Setting", "count_Expirepas", 0)
    Text5.text = Askcount

    AskOption = GetSetting(StrAppRegPath, "View_Type", "Expirepoket", True)
    
    If AskOption = False Then
        ChkExpirepoket.value = Unchecked
    Else
        ChkExpirepoket.value = Checked
    End If

    Askinterval = GetSetting(StrAppRegPath, "Setting", "INTERVAL_Expirepoket", "")

    If Askinterval = "D" Then
        Combo6.ListIndex = 0
    ElseIf Askinterval = "m" Then
        Combo6.ListIndex = 1
    ElseIf Askinterval = "yyyy" Then
        Combo6.ListIndex = 2
    Else
        Combo6.ListIndex = -1
    End If

    Askcount = GetSetting(StrAppRegPath, "Setting", "count_Expirepoket", 0)
    Text6.text = Askcount

    AskOption = GetSetting(StrAppRegPath, "Setting", "showhr", False)

    If AskOption = False Then
        ChKHR.value = Unchecked
    Else
        ChKHR.value = Checked
    End If

    AskOption = GetSetting(StrAppRegPath, "SallBill", "HaveTaxOnSalles", False)

 '   If AskOption = False Then
 '       ChkTax.value = Unchecked
 '   Else
 '       ChkTax.value = Checked
 '   End If

    '«· ·„ÌÕ «·ÌÊ„Ì
    AskOption = GetSetting(StrAppRegPath, "View_Type", "ShowToolTip", True)

    If AskOption = False Then
        ChkShowToolTip.value = Unchecked
    Else
        ChkShowToolTip.value = Checked
    End If

    '‘—Ìÿ «·«Œ ’«—« 
    AskOption = GetSetting(StrAppRegPath, "View_Type", "shortCuts", True)

    If AskOption = False Then
        chkshortCuts.value = Unchecked
    Else
        chkshortCuts.value = Checked
    End If

    '  ‘Ã—… «·«’‰«›
    AskOption = GetSetting(StrAppRegPath, "View_Type", "tree", True)

    If AskOption = False Then
        Chktree.value = Unchecked
    Else
        Chktree.value = Checked
    End If

    '    ⁄—÷ «·‰ ÌÃ…
    AskOption = GetSetting(StrAppRegPath, "View_Type", "Calender", True)

    If AskOption = False Then
        ChkCalender.value = Unchecked
    Else
        ChkCalender.value = Checked
    End If

    '    ⁄‰œ › Õ ‘«‘… ÃœÌœ… Ì› Õ ⁄·Ï ÃœÌœ «·Ì«
    AskOption = GetSetting(StrAppRegPath, "View_Type", "OPEN_NEW_SCREEN", False)

    If AskOption = False Then
        CHECK_OPEN_NEW_SCREEN.value = Unchecked
    Else
        CHECK_OPEN_NEW_SCREEN.value = Checked
    End If

    '  ⁄„— «·œÌ‰
    AskOption = GetSetting(StrAppRegPath, "View_Type", "ViewAging", False)

    If AskOption = False Then
        ChkViewAging.value = Unchecked
    Else
        ChkViewAging.value = Checked
    End If

    '⁄—÷ «”„ «·›—⁄  »Ã«‰» «·Õ”«» ›Ì «·ﬁÌœ
    AskOption = GetSetting(StrAppRegPath, "View_Type", "PrintBranchINGE", True)

    If AskOption = False Then
        ChkPrintBranchINGE.value = Unchecked
    Else
        ChkPrintBranchINGE.value = Checked
    End If

    '⁄—÷       „—ﬂ“ «· ﬂ·›…  ›Ì «·ﬁÌœ
    AskOption = GetSetting(StrAppRegPath, "View_Type", "PrintCCinGE", True)

    If AskOption = False Then
        ChkPrintCCinGE.value = Unchecked
    Else
        ChkPrintCCinGE.value = Checked
    End If

    '⁄—÷  «·—”„ «·»Ì«‰Ì ›Ì ﬂ‘› «·Õ”«»
    AskOption = GetSetting(StrAppRegPath, "View_Type", "ChartPrintinAS", True)

    If AskOption = False Then
        ChkChartPrintinAS.value = Unchecked
    Else
        ChkChartPrintinAS.value = Checked
    End If


    If rs("ShowPrinterDialoge2").value = True Then
        Me.Chkbarcode(225).value = vbChecked
    Else
        Me.Chkbarcode(225).value = Unchecked
    End If
        
        
    '«Œ›«¡ ﬂ· «· ‰»ÌÂ« 

    AskOption = GetSetting(StrAppRegPath, "View_Type", "HideAllAlarms", False)

    If AskOption = False Then
        ChkHideAllAlarms.value = Unchecked
    Else
        ChkHideAllAlarms.value = Checked
    End If

    ' ›⁄Ì· «·—”«∆·
    AskOption = GetSetting(StrAppRegPath, "View_Type", "Messnger", False)

    If AskOption = False Then
        ChkMessnger.value = Unchecked
    Else
        ChkMessnger.value = Checked
    End If
      intDef = GetSetting(StrAppRegPath, "DefaultOptions", "DefaultClient", 2)
    DBCboClientName.BoundText = intDef
    intDef = GetSetting(StrAppRegPath, "DefaultOptions", "DefaultSup", 1)
    DBCboSupName.BoundText = intDef
    intDef = GetSetting(StrAppRegPath, "DefaultOptions", "DefaultSaleStore", 1)
    DCboStoreName(0).BoundText = intDef
    intDef = GetSetting(StrAppRegPath, "DefaultOptions", "DefaultPurchaseStore", 1)
    DCboStoreName(1).BoundText = intDef
followload
    Exit Sub
ErrTrap:
End Sub
Private Sub followload()
Dim StrSQL As String
    StrSQL = "SELECT * From TblCustemers where Type=1"
    fill_combo Me.DBCboClientName, StrSQL
    StrSQL = "SELECT * From TblCustemers where Type=2"
    fill_combo Me.DBCboSupName, StrSQL
    StrSQL = "SELECT * From TblStore"
    fill_combo DCboStoreName(0), StrSQL
    fill_combo DCboStoreName(1), StrSQL
  

    If SystemOptions.SysDataBaseType = SQLServerDataBase Then
        opt(2).Enabled = True
    ElseIf SystemOptions.SysDataBaseType = AccessDataBase Then
        opt(2).Enabled = False
    End If

End Sub
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
If checksave = False Then

    Dim IntResult As String
    Dim StrMSG As String
    On Error GoTo ErrTrap
 
 
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
  
 

        IntResult = MsgBox(StrMSG, vbMsgBoxRight + vbYesNoCancel + vbMsgBoxRtlReading + vbQuestion, App.Title)

        Select Case IntResult

            Case vbYes
                Cancel = True
       
           CmdOk_Click

            Case vbCancel
                Cancel = True
        End Select

    End If

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

    Exit Sub
ErrTrap:
End Sub

Private Sub Label22_Click()
Chkbarcode(197).Visible = True
End Sub

Private Sub MYTEXT_KeyPress(Index As Integer, KeyAscii As Integer)
  KeyAscii = KeyAscii_Num(KeyAscii, MYTEXT(Index).text, 1)
End Sub

Private Sub Opt_Bey_Click()
    'If Opt_Bey.Value = True Then Opt_Sal.Value = True

End Sub

Private Sub opt_Branch_Click()
    Frame5.Visible = False
End Sub

Private Sub opt_group_Click()
 
    Frame5.Visible = True
 
End Sub

Private Sub Opt_OrderInpo_Click()
    'If Opt_OrderInpo.Value = True Then Opt_OrderOut.Value = True
End Sub

Private Sub Opt_OrderOut_Click()
    'If Opt_OrderOut.Value = True Then Opt_OrderInpo.Value = True
End Sub

Private Sub Opt_Sal_Click()
    'If Opt_Sal.Value = True Then Opt_Bey.Value = True

End Sub


Private Sub OptCurrQty_Click(Index As Integer)
UpdateCostriceProcedure
UpdateCostriceProcedureByStores
End Sub
  Function UpdateCostriceProcedure()
On Error Resume Next
Dim sql As String

    sql = "    DROP FUNCTION QryItemsTransactionsTotals" & CHR(13)
    Cn.Execute sql

    sql = " CREATE FUNCTION QryItemsTransactionsTotals(@TransType int =0,@TransType2 int=0,@TransType3 int=0,@FromDate datetime ,@ToDate datetime ,@ItemID  as integer ,@Transaction_ID as float=null )" & CHR(13)
    sql = sql & "RETURNS @xTable TABLE" & CHR(13)
    sql = sql & "(" & CHR(13)
    sql = sql & "ItemID int," & CHR(13)
    sql = sql & "ItemCode nvarchar(50)," & CHR(13)
    sql = sql & "ItemName nvarchar(4000)," & CHR(13)
    sql = sql & "GroupID  int," & CHR(13)
    sql = sql & "Total   money," & CHR(13)
    sql = sql & "totalqty Float" & CHR(13)
    sql = sql & ")" & CHR(13)
    sql = sql & "AS" & CHR(13)
    sql = sql & "Begin" & CHR(13)

    sql = sql & "INSERT @xTable" & CHR(13)
    sql = sql & "   Select ItemID,ItemCode,ItemName,GroupID,Sum(Total) as Totals,Sum(Quantity) as TotalQty" & CHR(13)
    sql = sql & "from" & CHR(13)
    sql = sql & "(" & CHR(13)
    sql = sql & "SELECT TblItems.ItemID,TblItems.ItemCode, TblItems.ItemName,TblItems.GroupID," & CHR(13)
    sql = sql & " Total= Transaction_Details.Quantity*Transaction_Details.Price " & CHR(13)
      sql = sql & ",Transaction_Details.Quantity" & CHR(13)
    sql = sql & "FROM dbo.TblItems INNER JOIN  dbo.Transaction_Details ON dbo.TblItems.ItemID = dbo.Transaction_Details.Item_ID INNER JOIN" & CHR(13)
    sql = sql & "dbo.Transactions ON dbo.Transaction_Details.Transaction_ID = dbo.Transactions.Transaction_ID" & CHR(13)
    sql = sql & "WHERE (Transactions.Transaction_Type=@TransType  OR Transactions.Transaction_Type=@TransType2 OR Transactions.Transaction_Type=@TransType3" & CHR(13)
    sql = sql & "or   Transactions.Transaction_Type=34 or   Transactions.Transaction_Type=11 or     Transactions.Transaction_Type=15 or     Transactions.Transaction_Type=39 )" & CHR(13)
    sql = sql & "AND" & CHR(13)
    sql = sql & "Transactions.Transaction_Date >=@FromDate" & CHR(13)
    sql = sql & "AND" & CHR(13)
    sql = sql & "Transactions.Transaction_Date <=@TODate" & CHR(13)
      sql = sql & "AND" & CHR(13)
   sql = sql & "  TblItems.ItemID =@ItemID" & CHR(13)
   sql = sql & " and  Transactions.Transaction_ID<>isnull(@Transaction_ID,Transactions.Transaction_ID)" & CHR(13)

   sql = sql & ")DrivTable" & CHR(13)
    
    
    sql = sql & "Group By ItemID,ItemCode,ItemName,GroupID" & CHR(13)
    sql = sql & "Return" & CHR(13)
    sql = sql & " End" & CHR(13)
    db_createOrUpdateFuctionSQL "QryItemsTransactionsTotals", sql

End Function

 

Private Sub tb_Switch(OldTab As Integer, _
                      NewTab As Integer, _
                      Cancel As Integer)

    If NewTab = 1 Then
        If checkApility("FrmyaersData") = False Then
            Exit Sub
        End If

        If bigUser = False Then
                If SystemOptions.UserInterface = ArabicInterface Then
                    MsgBox "€Ì— „”„ÊÕ ·ﬂ »«· ⁄«„· „⁄ Â–Â «·‰«›–…", vbOKOnly + vbCritical + vbMsgBoxRight + vbMsgBoxRtlReading, "’·«ÕÌ«  «·„” Œœ„Ì‰"
                 Else
                    MsgBox "Not Allowed", vbOKOnly + vbCritical + vbMsgBoxRight + vbMsgBoxRtlReading, "  Users Privligies"
                 End If
            Exit Sub
        End If
              
 
          FrmyaersData.show
 NewTab = 0
    ElseIf NewTab = 2 Then
 
        If checkApility("FrmBranchesData") = False Then
            Exit Sub
        End If
             
        If bigUser = False Then
         If SystemOptions.UserInterface = ArabicInterface Then
                    MsgBox "€Ì— „”„ÊÕ ·ﬂ »«· ⁄«„· „⁄ Â–Â «·‰«›–…", vbOKOnly + vbCritical + vbMsgBoxRight + vbMsgBoxRtlReading, "’·«ÕÌ«  «·„” Œœ„Ì‰"
                 Else
                    MsgBox "Not Allowed", vbOKOnly + vbCritical + vbMsgBoxRight + vbMsgBoxRtlReading, "  Users Privligies"
                 End If
             
             
             Exit Sub
        End If

         FrmBranchesData.show
    NewTab = 0
    ElseIf NewTab = 3 Then

        If bigUser = False Then
'            MsgBox "€Ì— „”„ÊÕ ·ﬂ »«· ⁄«„· „⁄ Â–Â «·‰«›–…", vbOKOnly + vbCritical + vbMsgBoxRight + vbMsgBoxRtlReading, "’·«ÕÌ«  «·„” Œœ„Ì‰"
              If SystemOptions.UserInterface = ArabicInterface Then
                    MsgBox "€Ì— „”„ÊÕ ·ﬂ »«· ⁄«„· „⁄ Â–Â «·‰«›–…", vbOKOnly + vbCritical + vbMsgBoxRight + vbMsgBoxRtlReading, "’·«ÕÌ«  «·„” Œœ„Ì‰"
                 Else
                    MsgBox "Not Allowed", vbOKOnly + vbCritical + vbMsgBoxRight + vbMsgBoxRtlReading, "  Users Privligies"
                 End If
            
            Exit Sub
        End If
            
        If checkApility("baranches") = False Then
            Exit Sub
        Else
            baranches.show
        End If
   NewTab = 0
    ElseIf NewTab = 18 Then

        If checkApility("FrmAccountsSeetting") = False Then
            Exit Sub
        End If
        
         FrmAccountsSeetting.show
             NewTab = 0
    ElseIf NewTab = 19 Then

        If checkApility("FrmDocType") = False Then
            Exit Sub
        End If
         
         FrmDocType.show
   NewTab = 0
    ElseIf NewTab = 20 Then

        If checkApility("System_manger2") = False Then
            Exit Sub
        End If

       ' System_manger2.show
          System_manger3.show
   NewTab = 0
    ElseIf NewTab = 21 Then

        If checkApility("coding") = False Then
            Exit Sub
        End If

         Coding.show
 NewTab = 0
    ElseIf NewTab = 22 Then

        If checkApility("FrmWorkFollowDesc") = False Then
            Exit Sub
        End If

         FrmWorkFollowDesc.show
    NewTab = 0
             
    End If

End Sub

Private Sub txt_ACCOUNT_digit_KeyPress(KeyAscii As Integer)
    KeyAscii = KeyAscii_Num(KeyAscii, txt_ACCOUNT_digit.text, 1)
End Sub

 



Private Sub TXTData_KeyPress(Index As Integer, KeyAscii As Integer)
  KeyAscii = KeyAscii_Num(KeyAscii, TxtData(Index).text, 1)
End Sub


Private Sub TxtEmpComponentDigts_KeyPress(KeyAscii As Integer)
    KeyAscii = KeyAscii_Num(KeyAscii, TxtEmpComponentDigts.text, 1)
End Sub

Private Sub TxtEmpSalaryDigts_KeyPress(KeyAscii As Integer)
KeyAscii = KeyAscii_Num(KeyAscii, TxtEmpSalaryDigts.text, 1)
End Sub

Private Sub TxtIndirectCostPercentage_KeyPress(KeyAscii As Integer)
    KeyAscii = KeyAscii_Num(KeyAscii, TxtIndirectCostPercentage.text, 1)
End Sub

Private Sub TxtLogoheight_KeyPress(KeyAscii As Integer)
KeyAscii = KeyAscii_Num(KeyAscii, TxtLogoheight.text, 1)
End Sub

Private Sub TxtLogoWidth_KeyPress(KeyAscii As Integer)
KeyAscii = KeyAscii_Num(KeyAscii, TxtLogoWidth.text, 1)
End Sub

Private Sub TxtPriceDigts_KeyPress(KeyAscii As Integer)
    KeyAscii = KeyAscii_Num(KeyAscii, TxtPriceDigts.text, 1)
End Sub

Private Sub TxtPriceDigtsInst_KeyPress(KeyAscii As Integer)
    KeyAscii = KeyAscii_Num(KeyAscii, TxtPriceDigtsInst.text, 1)
End Sub

Private Sub TxtQtyDigts_KeyPress(KeyAscii As Integer)
    KeyAscii = KeyAscii_Num(KeyAscii, TxtQtyDigts.text, 1)
End Sub

 
Private Sub TXTReturnSallingIntervalCount_KeyPress(Index As Integer, KeyAscii As Integer)
  KeyAscii = KeyAscii_Num(KeyAscii, TXTReturnSallingIntervalCount(Index).text, 1)

End Sub

Private Sub TxtZoom_KeyPress(KeyAscii As Integer)
   KeyAscii = KeyAscii_Num(KeyAscii, TxtZoom.text, 1)
End Sub



Private Sub LoadData()
    
   
'einvoice
  If rs("IsBluee").value = True Then
        Me.Chkbarcode(212).value = vbChecked
    Else
        Me.Chkbarcode(212).value = Unchecked
    End If
    
    
        If rs("ApplyEinvoice").value = 1 Then
        Me.Chkbarcode(213).value = vbChecked
    Else
        Me.Chkbarcode(213).value = Unchecked
    End If
   
    
    If rs("CanUploadZakatOpt").value Then
        Me.Chkbarcode(223).value = vbChecked
    Else
        Me.Chkbarcode(223).value = Unchecked
    End If
       If rs("IsCahngeServiceInvoice").value Then
        Me.Chkbarcode(224).value = vbChecked
    Else
        Me.Chkbarcode(224).value = Unchecked
    End If
   
   
   
    
     XPTxtComment(12).text = Trim(rs("ServerNameW").value & "")
    XPTxtComment(16).text = Trim(rs("DbNameW").value & "")
    



      If rs("EmpAccountByDep").value = 1 Then
        Me.Chkbarcode(217).value = vbChecked
    Else
        Me.Chkbarcode(217).value = Unchecked
    End If
 
   
    If rs("ApplyEinvoiceWithActive").value = True Then
        Me.Chkbarcode(216).value = vbChecked
    Else
        Me.Chkbarcode(216).value = Unchecked
    End If
    
       If rs("EmpAccountByDep").value = 1 Then
        Me.Chkbarcode(217).value = vbChecked
    Else
        Me.Chkbarcode(217).value = Unchecked
    End If
 
 
    
       If rs("ApplyEinvoiceWithBranch").value = True Then
        Me.Chkbarcode(218).value = vbChecked
    Else
        Me.Chkbarcode(218).value = Unchecked
    End If
  
      If rs("HiddenBalanceFromBox").value = 1 Then
        Me.Chkbarcode(219).value = vbChecked
    Else
        Me.Chkbarcode(219).value = Unchecked
    End If
  
 
    
    
XPTxtComment(4) = XPTxtComment(3)
XPTxtComment(5) = XPTxtComment(0)
'Wael
'XPTxtComment(12) = txtNoOFDigitUser(10)
XPTxtComment(10) = txtNoOFDigitUser(6)

XPTxtComment(8) = IIf(IsNull(rs("Commonname").value), "", rs("Commonname").value)
XPTxtComment(7) = IIf(IsNull(rs("SerialNumber").value), "", rs("SerialNumber").value)
XPTxtComment(6) = IIf(IsNull(rs("OrganizationName").value), "", rs("OrganizationName").value)
 
XPTxtComment(11) = IIf(IsNull(rs("industrey").value), "", rs("industrey").value)
XPTxtComment(9) = IIf(IsNull(rs("CSR").value), "", rs("CSR").value)
XPTxtComment(13) = IIf(IsNull(rs("Privatekey").value), "", rs("Privatekey").value)
XPTxtComment(14) = IIf(IsNull(rs("PublickeycertPem").value), "", rs("PublickeycertPem").value)
XPTxtComment(15) = IIf(IsNull(rs("SecretKey").value), "", rs("SecretKey").value)
    
  intDef = GetSetting(StrAppRegPath, "DefaultOptions", "DefaultClient", 2)
    DBCboClientName.BoundText = intDef
    intDef = GetSetting(StrAppRegPath, "DefaultOptions", "DefaultSup", 1)
    DBCboSupName.BoundText = intDef
    intDef = GetSetting(StrAppRegPath, "DefaultOptions", "DefaultSaleStore", 1)
    DCboStoreName(0).BoundText = intDef
    intDef = GetSetting(StrAppRegPath, "DefaultOptions", "DefaultPurchaseStore", 1)
    DCboStoreName(1).BoundText = intDef
    
    If SystemOptions.SysDataBaseType = SQLServerDataBase Then
        opt(2).Enabled = True
    ElseIf SystemOptions.SysDataBaseType = AccessDataBase Then
        opt(2).Enabled = False
    End If
    
      
''e invoice





End Sub

Function checkEeinvoice() As Boolean
checkEeinvoice = False

If XPTxtCompany.text = "" Or XPTxtCompanye.text = "" Then
      If SystemOptions.UserInterface = ArabicInterface Then
                    MsgBox " «”„ «·‘—ﬂ… ⁄—»Ì Ê«‰Ã·Ì“Ì «·“«„Ì", vbCritical
                Else
                MsgBox "enter COMPANY NAME  ", vbCritical
       End If
        checkEeinvoice = False
        Exit Function
End If


If XPTxtComment(0).text = "" Then
      If SystemOptions.UserInterface = ArabicInterface Then
                    MsgBox "—ﬁ„ «·”Ã· «·“«„Ì", vbCritical
                Else
                MsgBox "enter CRN ", vbCritical
       End If
        checkEeinvoice = False
        Exit Function
End If


If XPTxtComment(3).text = "" Or Len(XPTxtComment(3)) < 15 Then
      If SystemOptions.UserInterface = ArabicInterface Then
                    MsgBox "  «·—ﬁ„ «·÷—ÌÌ 15 Œ«‰…  «·“«„Ì", vbCritical
                Else
                MsgBox "Vat No 15 digit ", vbCritical
       End If
        checkEeinvoice = False
        Exit Function
End If



If txtNoOFDigitUser(4).text = "" Or Len(txtNoOFDigitUser(4)) < 4 Then
      If SystemOptions.UserInterface = ArabicInterface Then
                    MsgBox "     —ﬁ„ «·„»‰Ì 4 Œ«‰…  «·“«„Ì", vbCritical
                Else
                MsgBox "bulding no 4 digit ", vbCritical
       End If
        checkEeinvoice = False
        Exit Function
End If


If txtNoOFDigitUser(7).text = "" Or Len(txtNoOFDigitUser(7)) < 5 Then
      If SystemOptions.UserInterface = ArabicInterface Then
                    MsgBox "     «·—„“ «·»—ÌœÌ   5 Œ«‰…  «·“«„Ì", vbCritical
                Else
                MsgBox "Zib no 5 digit ", vbCritical
       End If
        checkEeinvoice = False
        Exit Function
End If


If txtNoOFDigitUser(2).text = "" Then
      If SystemOptions.UserInterface = ArabicInterface Then
                    MsgBox "    «”„ «·‘«—⁄  «·“«„Ì", vbCritical
                Else
                MsgBox "enter street name ", vbCritical
       End If
        checkEeinvoice = False
        Exit Function
End If

If txtNoOFDigitUser(10).text = "" Or Len(txtNoOFDigitUser(10)) < 2 Then
      If SystemOptions.UserInterface = ArabicInterface Then
                    MsgBox "   ﬂÊœ «·œÊ·…  «·“«„Ì 2 Œ«‰…", vbCritical
                Else
                MsgBox "must enter country code Code ", vbCritical
       End If
        checkEeinvoice = False
        Exit Function
End If


 


If (txtNoOFDigitUser(6)) = "" Then
      If SystemOptions.UserInterface = ArabicInterface Then
                    MsgBox "     «·„œÌ‰…  «·“«„Ì   ", vbCritical
                Else
                MsgBox "must enter city  ", vbCritical
       End If
        checkEeinvoice = False
        Exit Function
End If

If (txtNoOFDigitUser(9)) = "" Then
      If SystemOptions.UserInterface = ArabicInterface Then
                    MsgBox "     «·ÕÌ  «·“«„Ì   ", vbCritical
                Else
                MsgBox "must enter distict  ", vbCritical
       End If
        checkEeinvoice = False
        Exit Function
End If
checkEeinvoice = True


End Function


Function savepart2()
        If Me.Chkbarcode(212).value = vbChecked Then
        rs("IsBluee").value = 1
    ElseIf Me.Chkbarcode(212).value = vbUnchecked Then
        rs("IsBluee").value = 0
    End If
    
    If Me.Chkbarcode(223).value = vbChecked Then
        rs("CanUploadZakatOpt").value = 1
    ElseIf Me.Chkbarcode(223).value = vbUnchecked Then
        rs("CanUploadZakatOpt").value = 0
    End If
    If Me.Chkbarcode(224).value = vbChecked Then
        rs("IsCahngeServiceInvoice").value = 1
    ElseIf Me.Chkbarcode(224).value = vbUnchecked Then
        rs("IsCahngeServiceInvoice").value = 0
    End If
    
    rs("ServerNameW").value = XPTxtComment(12).text
    rs("DbNameW").value = XPTxtComment(16).text
    

 
    If Me.Chkbarcode(217).value = vbChecked Then
        rs("EmpAccountByDep").value = 1
    ElseIf Me.Chkbarcode(217).value = vbUnchecked Then
        rs("EmpAccountByDep").value = 0
    End If
 
 
 
    If Me.Chkbarcode(216).value = vbChecked Then
        rs("ApplyEinvoiceWithActive").value = 1
    ElseIf Me.Chkbarcode(216).value = vbUnchecked Then
        rs("ApplyEinvoiceWithActive").value = 0
    End If
 
 

 
    If Me.Chkbarcode(218).value = vbChecked Then
        rs("ApplyEinvoiceWithBranch").value = 1
    ElseIf Me.Chkbarcode(218).value = vbUnchecked Then
        rs("ApplyEinvoiceWithBranch").value = 0
    End If
  

    If Me.Chkbarcode(219).value = vbChecked Then
        rs("HiddenBalanceFromBox").value = 1
    ElseIf Me.Chkbarcode(218).value = vbUnchecked Then
        rs("HiddenBalanceFromBox").value = 0
    End If
   
 
    
    If Me.Chkbarcode(213).value = vbChecked Then
        rs("ApplyEinvoice").value = 1
    ElseIf Me.Chkbarcode(213).value = vbUnchecked Then
        rs("ApplyEinvoice").value = 0
    End If
    
       rs("StreetName").value = txtNoOFDigitUser(2).text
        rs("BuildingNumber").value = txtNoOFDigitUser(4).text
         rs("CitySubdivisionName").value = txtNoOFDigitUser(9).text
          rs("CityName").value = txtNoOFDigitUser(6).text
           rs("PostalZone").value = txtNoOFDigitUser(7).text
            rs("IdentificationCode").value = txtNoOFDigitUser(10).text
             rs("PlotIdentification").value = txtNoOFDigitUser(5).text
              rs("AdditionalStreetName").value = txtNoOFDigitUser(3).text
              rs("CountrySubentity").value = txtNoOFDigitUser(8).text
              
          rs("ActivityName").value = txtActivityName.text
                  
    
        
End Function

