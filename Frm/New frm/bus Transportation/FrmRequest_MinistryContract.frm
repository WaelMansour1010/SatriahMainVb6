VERSION 5.00
Object = "{0AFE7BE0-11B7-4A3E-978D-D4501E9A57FE}#1.0#0"; "c1sizer.ocx"
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{FE5DCFAD-BC1D-11D2-94CF-004005455FAA}#1.4#0"; "ImpulseButton.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form FrmRequest_MinistryContract 
   BackColor       =   &H00E2E9E9&
   Caption         =   " «À»«  «·«” Õﬁ«ﬁ«  «·‘Â—Ì… ·⁄ﬁÊœ «·Ê“«—…"
   ClientHeight    =   10380
   ClientLeft      =   120
   ClientTop       =   405
   ClientWidth     =   15780
   Icon            =   "FrmRequest_MinistryContract.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   RightToLeft     =   -1  'True
   ScaleHeight     =   10380
   ScaleWidth      =   15780
   WindowState     =   2  'Maximized
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   8
      Charset         =   178
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Begin C1SizerLibCtl.C1Elastic Main_CLE 
      Height          =   10380
      Left            =   0
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   0
      Width           =   15780
      _cx             =   27834
      _cy             =   18309
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
      Begin C1SizerLibCtl.C1Elastic C1Elastic3 
         Height          =   1320
         Left            =   0
         TabIndex        =   2
         TabStop         =   0   'False
         Top             =   8844
         Width           =   15768
         _cx             =   27808
         _cy             =   2328
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
         Begin MSComDlg.CommonDialog cd 
            Left            =   3120
            Top             =   120
            _ExtentX        =   847
            _ExtentY        =   847
            _Version        =   393216
         End
         Begin VB.CommandButton Command2 
            Caption         =   " ’œÌ—«·Ï «·«ﬂ”Ì·"
            Height          =   375
            Left            =   6252
            RightToLeft     =   -1  'True
            TabIndex        =   51
            Top             =   240
            Width           =   1488
         End
         Begin VB.Frame Frame9 
            Caption         =   "»Ì«‰«  „Õ«”»Ì…"
            Height          =   744
            Left            =   8304
            RightToLeft     =   -1  'True
            TabIndex        =   43
            Top             =   0
            Width           =   7308
            Begin VB.CommandButton Command8 
               Caption         =   "ﬂ‘› Õ”«»"
               Height          =   375
               Left            =   120
               RightToLeft     =   -1  'True
               TabIndex        =   47
               Top             =   240
               Visible         =   0   'False
               Width           =   1215
            End
            Begin VB.CommandButton Command9 
               Caption         =   "ÿ»«⁄Â «·ﬁÌœ"
               Height          =   375
               Left            =   1560
               RightToLeft     =   -1  'True
               TabIndex        =   46
               Top             =   240
               Width           =   1095
            End
            Begin VB.TextBox TxtNoteSerial 
               Alignment       =   1  'Right Justify
               Height          =   405
               Left            =   3720
               RightToLeft     =   -1  'True
               TabIndex        =   45
               Top             =   240
               Width           =   2415
            End
            Begin VB.TextBox TxtNoteID 
               Alignment       =   1  'Right Justify
               Height          =   285
               Left            =   2760
               RightToLeft     =   -1  'True
               TabIndex        =   44
               Top             =   120
               Visible         =   0   'False
               Width           =   855
            End
            Begin VB.Label Label1 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               BackStyle       =   0  'Transparent
               Caption         =   "—ﬁ„ «·ﬁÌœ"
               Height          =   195
               Index           =   35
               Left            =   5880
               RightToLeft     =   -1  'True
               TabIndex        =   48
               Top             =   240
               Width           =   990
            End
         End
         Begin ImpulseButton.ISButton Cmd 
            Height          =   396
            Index           =   0
            Left            =   14052
            TabIndex        =   3
            Top             =   804
            Width           =   1488
            _ExtentX        =   2619
            _ExtentY        =   688
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
            Height          =   396
            Index           =   1
            Left            =   12456
            TabIndex        =   4
            Top             =   804
            Width           =   1488
            _ExtentX        =   2619
            _ExtentY        =   688
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
            ColorShadow     =   -2147483637
            ColorOutline    =   0
            DrawFocusRectangle=   0   'False
            ColorToggledHoverText=   16711680
            ColorTextShadow =   -2147483637
         End
         Begin ImpulseButton.ISButton Cmd 
            Height          =   396
            Index           =   2
            Left            =   10920
            TabIndex        =   5
            Top             =   804
            Width           =   1500
            _ExtentX        =   2646
            _ExtentY        =   688
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
            Height          =   396
            Index           =   3
            Left            =   9384
            TabIndex        =   6
            Top             =   804
            Width           =   1488
            _ExtentX        =   2619
            _ExtentY        =   688
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
            Height          =   396
            Index           =   4
            Left            =   7788
            TabIndex        =   7
            Top             =   804
            Width           =   1512
            _ExtentX        =   2672
            _ExtentY        =   688
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
            Height          =   396
            Index           =   6
            Left            =   1620
            TabIndex        =   9
            Top             =   804
            Width           =   1488
            _ExtentX        =   2619
            _ExtentY        =   688
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
         Begin ImpulseButton.ISButton CmdAttach 
            Height          =   396
            Left            =   60
            TabIndex        =   10
            Top             =   804
            Width           =   1488
            _ExtentX        =   2619
            _ExtentY        =   688
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
            ColorShadow     =   -2147483637
            ColorOutline    =   0
            DrawFocusRectangle=   0   'False
            ColorToggledHoverText=   16711680
            ColorTextShadow =   -2147483637
         End
         Begin ImpulseButton.ISButton Cmd 
            Height          =   396
            Index           =   7
            Left            =   4632
            TabIndex        =   8
            Top             =   804
            Width           =   1488
            _ExtentX        =   2619
            _ExtentY        =   688
            ButtonPositionImage=   1
            Caption         =   "ÿ»«⁄Â"
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
            Height          =   396
            Index           =   5
            Left            =   3120
            TabIndex        =   49
            Top             =   804
            Width           =   1500
            _ExtentX        =   2646
            _ExtentY        =   688
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
            Height          =   396
            Index           =   8
            Left            =   6252
            TabIndex        =   50
            Top             =   804
            Width           =   1488
            _ExtentX        =   2619
            _ExtentY        =   688
            ButtonPositionImage=   1
            Caption         =   "≈·€«¡ «·œ›⁄« "
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
         Begin VB.Label XPTxtCurrent 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Height          =   336
            Left            =   4260
            RightToLeft     =   -1  'True
            TabIndex        =   21
            Top             =   240
            Width           =   684
         End
         Begin VB.Label XPTxtCount 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Height          =   336
            Left            =   168
            RightToLeft     =   -1  'True
            TabIndex        =   20
            Top             =   252
            Width           =   648
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   " «·”Ã· «·Õ«·Ì:"
            Height          =   336
            Index           =   2
            Left            =   4992
            RightToLeft     =   -1  'True
            TabIndex        =   19
            Top             =   252
            Width           =   1248
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   " ⁄œœ «·”Ã·« :"
            Height          =   336
            Index           =   4
            Left            =   876
            RightToLeft     =   -1  'True
            TabIndex        =   18
            Top             =   252
            Width           =   1296
         End
      End
      Begin C1SizerLibCtl.C1Elastic C1Elastic2 
         Height          =   7140
         Left            =   0
         TabIndex        =   17
         TabStop         =   0   'False
         Top             =   1635
         Width           =   15765
         _cx             =   27808
         _cy             =   12594
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
         Begin VB.TextBox total 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFFC0&
            Height          =   315
            Left            =   11520
            Locked          =   -1  'True
            TabIndex        =   56
            Top             =   6720
            Width           =   2460
         End
         Begin VB.TextBox TxtFATYou 
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
            Height          =   315
            Left            =   8160
            Locked          =   -1  'True
            MaxLength       =   50
            RightToLeft     =   -1  'True
            TabIndex        =   55
            Top             =   6735
            Width           =   1860
         End
         Begin VB.TextBox TxtFATValue 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFC0&
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   4080
            Locked          =   -1  'True
            MaxLength       =   50
            RightToLeft     =   -1  'True
            TabIndex        =   54
            Top             =   6735
            Width           =   2460
         End
         Begin VB.TextBox TxtTotalValue 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFC0&
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   240
            Locked          =   -1  'True
            MaxLength       =   50
            RightToLeft     =   -1  'True
            TabIndex        =   53
            Top             =   6735
            Width           =   2820
         End
         Begin VB.CheckBox chkChooseAll 
            Alignment       =   1  'Right Justify
            Caption         =   "«Œ Ì«— «·ﬂ·"
            Height          =   255
            Left            =   14460
            RightToLeft     =   -1  'True
            TabIndex        =   41
            Top             =   0
            Width           =   1125
         End
         Begin VSFlex8Ctl.VSFlexGrid Grid 
            Height          =   6240
            Left            =   0
            TabIndex        =   1
            Top             =   360
            Width           =   15810
            _cx             =   27887
            _cy             =   11007
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
            AllowBigSelection=   0   'False
            AllowUserResizing=   0
            SelectionMode   =   1
            GridLines       =   1
            GridLinesFixed  =   2
            GridLineWidth   =   1
            Rows            =   50
            Cols            =   40
            FixedRows       =   2
            FixedCols       =   1
            RowHeightMin    =   320
            RowHeightMax    =   0
            ColWidthMin     =   0
            ColWidthMax     =   0
            ExtendLastCol   =   0   'False
            FormatString    =   $"FrmRequest_MinistryContract.frx":038A
            ScrollTrack     =   0   'False
            ScrollBars      =   3
            ScrollTips      =   0   'False
            MergeCells      =   0
            MergeCompare    =   0
            AutoResize      =   0   'False
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
            Editable        =   1
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
         Begin MSComctlLib.ProgressBar ProgressBar1 
            Height          =   372
            Left            =   4212
            TabIndex        =   42
            Top             =   -2040
            Visible         =   0   'False
            Width           =   8796
            _ExtentX        =   15505
            _ExtentY        =   661
            _Version        =   393216
            Appearance      =   0
         End
         Begin VB.Label Label36 
            Alignment       =   2  'Center
            BackColor       =   &H00E2E9E9&
            Caption         =   "’«›Ì «·ﬁÌ„…"
            ForeColor       =   &H00C00000&
            Height          =   300
            Left            =   13980
            TabIndex        =   60
            Top             =   6735
            Width           =   1530
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            BackStyle       =   0  'Transparent
            Caption         =   "‰”»…«·›« "
            ForeColor       =   &H00C00000&
            Height          =   300
            Index           =   66
            Left            =   10365
            RightToLeft     =   -1  'True
            TabIndex        =   59
            Top             =   6735
            Width           =   690
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            BackStyle       =   0  'Transparent
            Caption         =   "«·«Ã„«·Ì"
            ForeColor       =   &H00C00000&
            Height          =   300
            Index           =   68
            Left            =   3165
            RightToLeft     =   -1  'True
            TabIndex        =   58
            Top             =   6735
            Width           =   810
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            BackStyle       =   0  'Transparent
            Caption         =   "ﬁÌ„… «·›« "
            ForeColor       =   &H00C00000&
            Height          =   300
            Index           =   67
            Left            =   6720
            RightToLeft     =   -1  'True
            TabIndex        =   57
            Top             =   6735
            Width           =   930
         End
      End
      Begin C1SizerLibCtl.C1Elastic EleHeader 
         Height          =   732
         Left            =   0
         TabIndex        =   11
         TabStop         =   0   'False
         Top             =   0
         Width           =   15816
         _cx             =   27887
         _cy             =   1296
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
         Appearance      =   0
         MousePointer    =   0
         Version         =   801
         BackColor       =   16777215
         ForeColor       =   4210688
         FloodColor      =   6553600
         ForeColorDisabled=   -2147483631
         Caption         =   "     «À»«  «·«” Õﬁ«ﬁ«  «·‘Â—Ì… ·⁄ﬁÊœ «·Ê“«—…   "
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
            TabIndex        =   12
            Top             =   180
            Visible         =   0   'False
            Width           =   855
         End
         Begin ImpulseButton.ISButton XPBtnMove 
            Height          =   345
            Index           =   0
            Left            =   1155
            TabIndex        =   13
            Top             =   120
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
            ButtonImage     =   "FrmRequest_MinistryContract.frx":096C
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
            TabIndex        =   14
            Top             =   120
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
            ButtonImage     =   "FrmRequest_MinistryContract.frx":0D06
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
            TabIndex        =   15
            Top             =   120
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
            ButtonImage     =   "FrmRequest_MinistryContract.frx":10A0
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
            TabIndex        =   16
            Top             =   120
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
            ButtonImage     =   "FrmRequest_MinistryContract.frx":143A
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
      Begin C1SizerLibCtl.C1Elastic C1Elastic1 
         Height          =   675
         Left            =   0
         TabIndex        =   22
         TabStop         =   0   'False
         Top             =   840
         Width           =   15840
         _cx             =   27940
         _cy             =   1191
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
         Begin VB.TextBox txtID 
            Alignment       =   1  'Right Justify
            Enabled         =   0   'False
            Height          =   324
            Left            =   13725
            Locked          =   -1  'True
            MaxLength       =   10
            RightToLeft     =   -1  'True
            TabIndex        =   26
            Top             =   270
            Width           =   1272
         End
         Begin VB.TextBox txtCode 
            Alignment       =   1  'Right Justify
            Height          =   324
            Left            =   9408
            MaxLength       =   50
            RightToLeft     =   -1  'True
            TabIndex        =   25
            Top             =   1035
            Visible         =   0   'False
            Width           =   1080
         End
         Begin VB.ComboBox cbType 
            Height          =   315
            Left            =   6708
            RightToLeft     =   -1  'True
            TabIndex        =   24
            Top             =   1035
            Visible         =   0   'False
            Width           =   1704
         End
         Begin VB.CommandButton Command1 
            Caption         =   "⁄—÷"
            Height          =   528
            Left            =   165
            RightToLeft     =   -1  'True
            TabIndex        =   23
            Top             =   120
            Width           =   1008
         End
         Begin MSDataListLib.DataCombo DcDur 
            Height          =   288
            Left            =   3372
            TabIndex        =   27
            Tag             =   "⁄›Ê« Ì—ÃÏ «œŒ«· √”„ «·ÕÌ"
            Top             =   276
            Width           =   1500
            _ExtentX        =   2646
            _ExtentY        =   556
            _Version        =   393216
            Text            =   ""
            RightToLeft     =   -1  'True
         End
         Begin MSDataListLib.DataCombo dcMontth 
            Height          =   288
            Left            =   1272
            TabIndex        =   28
            Tag             =   "⁄›Ê« Ì—ÃÏ «œŒ«· √”„ «·ÕÌ"
            Top             =   276
            Width           =   1560
            _ExtentX        =   2752
            _ExtentY        =   556
            _Version        =   393216
            Text            =   ""
            RightToLeft     =   -1  'True
         End
         Begin MSComCtl2.DTPicker RecDate 
            Height          =   348
            Left            =   11388
            TabIndex        =   29
            TabStop         =   0   'False
            Top             =   240
            Width           =   1368
            _ExtentX        =   2408
            _ExtentY        =   609
            _Version        =   393216
            CalendarBackColor=   12648447
            CalendarTitleBackColor=   10383715
            CustomFormat    =   "yyyy/M/d"
            Format          =   99155971
            CurrentDate     =   37140
         End
         Begin Dynamic_Byte.NourHijriCal DateH 
            Height          =   348
            Left            =   10572
            TabIndex        =   30
            Top             =   240
            Visible         =   0   'False
            Width           =   816
            _ExtentX        =   1429
            _ExtentY        =   609
         End
         Begin MSDataListLib.DataCombo dcBranch 
            Height          =   315
            Left            =   8280
            TabIndex        =   31
            Tag             =   "⁄›Ê« Ì—ÃÏ «œŒ«· √”„ «·ÕÌ"
            Top             =   240
            Width           =   1335
            _ExtentX        =   2355
            _ExtentY        =   556
            _Version        =   393216
            Enabled         =   0   'False
            Text            =   ""
            RightToLeft     =   -1  'True
         End
         Begin MSDataListLib.DataCombo dcM 
            Height          =   288
            Left            =   5736
            TabIndex        =   39
            Tag             =   "⁄›Ê« Ì—ÃÏ «œŒ«· √”„ «·ÕÌ"
            Top             =   240
            Width           =   1248
            _ExtentX        =   2223
            _ExtentY        =   556
            _Version        =   393216
            Text            =   ""
            RightToLeft     =   -1  'True
         End
         Begin MSDataListLib.DataCombo AccountVat 
            Bindings        =   "FrmRequest_MinistryContract.frx":17D4
            Height          =   315
            Left            =   0
            TabIndex        =   52
            Top             =   240
            Visible         =   0   'False
            Width           =   3450
            _ExtentX        =   6085
            _ExtentY        =   556
            _Version        =   393216
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
            Caption         =   "«·⁄ﬁœ «·Ê“«—Ï"
            Height          =   312
            Index           =   6
            Left            =   6948
            RightToLeft     =   -1  'True
            TabIndex        =   40
            Top             =   240
            Width           =   1212
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "‰Ê⁄ «·’—›"
            Height          =   390
            Index           =   0
            Left            =   8610
            RightToLeft     =   -1  'True
            TabIndex        =   38
            Top             =   1035
            Visible         =   0   'False
            Width           =   585
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "—ﬁ„ «·”‰œ"
            Height          =   336
            Index           =   8
            Left            =   14520
            RightToLeft     =   -1  'True
            TabIndex        =   37
            Top             =   276
            Width           =   1116
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "«·—ﬁ„ «·ÌœÊÏ"
            Height          =   300
            Index           =   9
            Left            =   10344
            RightToLeft     =   -1  'True
            TabIndex        =   36
            Top             =   1308
            Visible         =   0   'False
            Width           =   864
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "«·”‰… «·œ—«”Ì…"
            Height          =   435
            Index           =   3
            Left            =   4920
            RightToLeft     =   -1  'True
            TabIndex        =   35
            Top             =   150
            Width           =   750
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "«·› —…"
            Height          =   312
            Index           =   1
            Left            =   2760
            RightToLeft     =   -1  'True
            TabIndex        =   34
            Top             =   276
            Width           =   480
         End
         Begin VB.Label Label3 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   " «—ÌŒ «·”‰œ"
            ForeColor       =   &H00000000&
            Height          =   285
            Left            =   12735
            TabIndex        =   33
            Top             =   240
            Width           =   870
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "«·›—⁄"
            Height          =   312
            Index           =   5
            Left            =   9696
            RightToLeft     =   -1  'True
            TabIndex        =   32
            Top             =   240
            Width           =   732
         End
      End
   End
End
Attribute VB_Name = "FrmRequest_MinistryContract"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rs As ADODB.Recordset
Dim RsTemp As ADODB.Recordset
Dim RsTemp2 As ADODB.Recordset
Dim RsTemp3 As ADODB.Recordset
Dim Rs_Temp As ADODB.Recordset
Dim TTP As clstooltip
Dim Account_Code_dynamic As String



Private Sub chkChooseAll_Click()

If Me.TxtModFlg.Text = "E" Or Me.TxtModFlg.Text = "N" Then
            Else
                  Exit Sub
  End If

Dim i As Integer

For i = 1 To Grid.Rows - 1
    If Grid.TextMatrix(i, Grid.ColIndex("Due_Date")) <> "" Then
            If chkChooseAll.value = 1 Then
                    Grid.TextMatrix(i, Grid.ColIndex("Status")) = 1
            Else
                    Grid.TextMatrix(i, Grid.ColIndex("Status")) = 0
            End If
    End If
Next
ClculteVAT
Relain
End Sub

Private Sub Cmd_Click(Index As Integer)
 '    On Error GoTo ErrTrap

    Select Case Index
        Case 0
            If DoPremis(Do_New, Me.Name, True) = False Then
                Exit Sub
            End If
            TxtModFlg.Text = "N"
            clear_all Me
            TxtId.Text = CStr(new_id("TblRequest_MinistryContract", "ID", "", True))
          '  TXTid.SetFocus
             Grid.Rows = Grid.FixedRows
             ClculteVAT
   
        Case 1
        
               If ChekClodePeriod(Me.RecDate.value) = True Then
               If SystemOptions.UserInterface = ArabicInterface Then
                MsgBox "Ì—ÃÏ  €ÌÌ—  «—ÌŒ «·Õ—ﬂ… ·«‰ Â–Â «·› —… „€·ﬁ…"
               Else
               MsgBox "Please Change Date Becouse This is Period is Closed"
              End If
              Exit Sub
              End If
              

            If DoPremis(Do_Edit, Me.Name, True) = False Then
                Exit Sub
            End If
                TxtModFlg.Text = "E"
                C1Elastic1.Enabled = False
        Case 2
            If ChekClodePeriod(Me.RecDate.value) = True Then
               If SystemOptions.UserInterface = ArabicInterface Then
                MsgBox "Ì—ÃÏ  €ÌÌ—  «—ÌŒ «·Õ—ﬂ… ·«‰ Â–Â «·› —… „€·ﬁ…"
               Else
               MsgBox "Please Change Date Becouse This is Period is Closed"
              End If
              Exit Sub
              End If
              
                Account_Code_dynamic = get_account_code_branch(105, my_branch)
        
    If Account_Code_dynamic = "NO branch" Then
        If SystemOptions.UserInterface = ArabicInterface Then
            MsgBox "·„ Ì „ «‰‘«¡ «·›—⁄", vbCritical
        Else
            MsgBox "Branch Not Created", vbCritical
        End If

 Exit Sub
    ElseIf Account_Code_dynamic = "NO account" Then

        If SystemOptions.UserInterface = ArabicInterface Then
            MsgBox "·„ Ì „  ÕœÌœ «Ì—«œ«·    «·‰ﬁ·  «·„»Ì⁄«   ›Ì «·›—⁄ ·Â–… «·⁄„·Ì…", vbCritical
        Else
            MsgBox "Sales Account Not Defined in this Branch", vbCritical
        End If

        Exit Sub
         
    End If
'Dim AccountVATDept As String
'If AccountVat.BoundText = "" And True= True And CheckAnyVAT = True Then
'MsgBox "Ì—ÃÏ ÷»ÿ «⁄œ«œ  «·ﬁÌ„… «·„÷«›…"
'Exit Sub
'End If
            SaveData

        Case 3
            Undo

        Case 4
         If ChekClodePeriod(Me.RecDate.value) = True Then
               If SystemOptions.UserInterface = ArabicInterface Then
                MsgBox "Ì—ÃÏ  €ÌÌ—  «—ÌŒ «·Õ—ﬂ… ·«‰ Â–Â «·› —… „€·ﬁ…"
               Else
               MsgBox "Please Change Date Becouse This is Period is Closed"
              End If
              Exit Sub
              End If
              
              
            If DoPremis(Do_Delete, Me.Name, True) = False Then
                Exit Sub
            End If

            Del_Company

        Case 5
                Unload FrmSearch_Request
                 FrmSearch_Request.SendForm = "MR_MR"
                 FrmSearch_Request.show
        Case 6
            Unload Me
         Case 7
         print_report
          Case 8
          
               Dim Msg As String
               If SystemOptions.UserInterface = ArabicInterface Then
                        Msg = " ”Ì „ Õ–› «·„œ›Ê⁄«  ··”‰œ ° Â·  —Ìœ «” ﬂ„«· ⁄„·Ì… «·Õ–› ø "
               Else
                        Msg = " This action will delete Paid for this receipt "
               End If
               If MsgBox(Msg, vbOKCancel) = vbOK Then
                    Cancel_Paid
                    Command1_Click
                    C1Elastic1.Enabled = True
                    C1Elastic2.Enabled = True
                End If
                
    End Select

    Exit Sub
ErrTrap:
End Sub

Private Sub Cancel_Paid()

Dim str As String, AllID As String, i As Integer
If rs.RecordCount > 0 Then
        AllID = IIf(IsNull(rs("AllID").value), "", rs("AllID").value)
        rs("AllID") = Null
        rs.update
End If


If AllID = "" Then
    Exit Sub
End If

str = " select * from TblMinistryContract_Installment where id in (  " & AllID & "  )"
Set Rs_Temp = New ADODB.Recordset
Rs_Temp.Open str, Cn, adOpenStatic, adLockOptimistic, adCmdText

If Rs_Temp.RecordCount > 0 Then
        For i = 0 To Rs_Temp.RecordCount - 1
                Rs_Temp("Paid") = Null
                Rs_Temp("RID") = Null
                Rs_Temp.update
                Rs_Temp.MoveNext
        Next
End If



End Sub

Private Sub CmdAttach_Click()
            On Error Resume Next
'ShowAttachments XPTxtBoxID, "0701201405"
 
     On Error Resume Next
          If DoPremis(Do_Attach, Me.Name, True) = False Then
                Exit Sub
            End If
            
ShowAttachments TxtId, "15062020002"



End Sub

Private Sub CmdHelp_Click()
    SystemOptions.SysHelp.HHTopicID = Me.HelpContextID
    SystemOptions.SysHelp.HHDisplayTopicID Me.hwnd
End Sub

Private Sub Command2_Click()
    On Error Resume Next
    Dim StrFileName As String
    StrFileName = App.path & "\Report1.xls"

    If Dir(StrFileName) <> "" Then
        Kill StrFileName
    End If
'Grid.RightToLeft = True
  '  Me.Grid.saveGrid StrFileName, flexFileExcel, True
  '  OpenFile StrFileName
    
         On Error Resume Next
      cd.CancelError = True 'allow escape key/cancel
     cd.filename = "Report"
    cd.ShowSave     'show the dialog screen
    If Err <> 32755 Then    ' User didn't chose Cancel.
   Else
       Exit Sub
    End If
 StrFileName = cd.filename & ".xls"
Me.Grid.saveGrid StrFileName, flexFileCustomText, True
   
    OpenFile StrFileName
    
End Sub

Private Sub Command9_Click()
ShowGL_cc Me.TxtNoteSerial.Text, , 200
End Sub
Public Function CREATE_VOUCHER_GE(general_noteid As Long, BranchID As Integer, UserID As Long _
, NoteDate As Date)
Dim BasicSalaryAccount As String
Dim StrSQL As String
         StrSQL = "Delete From DOUBLE_ENTREY_VOUCHERS Where Notes_ID=" & general_noteid
        Cn.Execute StrSQL, , adExecuteNoRecords

'TxtNoteSerial.text = ""

    Dim i As Integer
    Dim LngDevID As Long
    Dim Msg As String
   
    Dim Employee_account As String
    Dim StrAccountCode As String
    Dim X As Integer
    Dim rs As New ADODB.Recordset
    Dim notes_serial As String
    Dim notes_id As String
    Dim j As Integer
    Dim ColumnName As String
    Dim SalaryAccount As String
    Dim BonusAccount As String
    Dim DiscountAccount As String
        Msg = EleHeader.Caption & " —ﬁ„ " & TxtId & " » «—ÌŒ" & Date
 
 
        
 
 notes_id = general_noteid

  
   
    LngDevID = new_id("DOUBLE_ENTREY_VOUCHERS", "Double_Entry_Vouchers_ID", "")

    Dim line_no As Integer
    line_no = 1
                
    'C???? C??I?? C?C?C?CE
     
    Dim CValue As Double
    Dim Branch As Integer
    Dim ProjectID As Integer
    Dim CustomerBranchId As Integer
        Dim DeptSide1 As String
                 Dim credit_side1 As String
                 
    BranchID = 1
    
    With Grid


line_no = 1
        For i = .FixedRows To .Rows - 1
    BranchID = val(dcBranch.BoundText)
    
            If .TextMatrix(i, .ColIndex("Value")) > 0 And .TextMatrix(i, .ColIndex("Account_Code")) <> "" And .Cell(flexcpChecked, i, .ColIndex("Status")) = flexChecked Then   'C?C??? C???E??E IC??
                'Employee_account = get_EMPLOYEE_Account(val(.TextMatrix(i, .ColIndex("Emp_ID"))), "Account_Code1") 'C?C??? C???E??E
                StrAccountCode = .TextMatrix(i, .ColIndex("Account_Code"))
             CustomerBranchId = val(.TextMatrix(i, .ColIndex("BranchId")))
    
                
                
              
            
                If CustomerBranchId <> BranchID Then
                     
                     DeptSide1 = getBranchCurrentAccount(BranchID)
                       credit_side1 = getBranchCurrentAccount(CustomerBranchId)



            If ModAccounts.AddNewDev(LngDevID, line_no, StrAccountCode, val(.TextMatrix(i, .ColIndex("FATValue"))) + val(.TextMatrix(i, .ColIndex("Value"))), 0, Msg & "  ··⁄ﬁœ  " & dcM.Text & "  ··œ›⁄Â  " & .TextMatrix(i, .ColIndex("InstallmentNo")), val(notes_id), , , , NoteDate, user_id, , , , , , , , , setfoxy_Line, , , , , , , , , CustomerBranchId) = False Then
                    GoTo ErrTrap
                End If

                line_no = line_no + 1
                
                
               If ModAccounts.AddNewDev(LngDevID, line_no, Account_Code_dynamic, .TextMatrix(i, .ColIndex("Value")), 1, Msg & "  ··⁄ﬁœ  " & dcM.Text & "  ··œ›⁄Â  " & .TextMatrix(i, .ColIndex("InstallmentNo")), val(notes_id), , , , NoteDate, user_id, , , , , , , , , setfoxy_Line, , , , , , , , , BranchID) = False Then
                    GoTo ErrTrap
                End If
If val(.TextMatrix(i, .ColIndex("FATValue"))) <> 0 Then
                line_no = line_no + 1
               If ModAccounts.AddNewDev(LngDevID, line_no, .TextMatrix(i, .ColIndex("AccountCodeVat")), .TextMatrix(i, .ColIndex("FATValue")), 1, "Õ”«» «·ﬁÌ„… «·„÷«›…" & "  ··⁄ﬁœ  " & dcM.Text & "  ··œ›⁄Â  " & .TextMatrix(i, .ColIndex("InstallmentNo")), val(notes_id), , , , NoteDate, user_id, , , , , , , , , setfoxy_Line, , , , , , , , , BranchID) = False Then
                    GoTo ErrTrap
                End If
   End If
                line_no = line_no + 1
                
                
                
                           If ModAccounts.AddNewDev(LngDevID, line_no, credit_side1, val(.TextMatrix(i, .ColIndex("Value"))) + val(.TextMatrix(i, .ColIndex("FATValue"))), 0, Msg & "  ··⁄ﬁœ  " & dcM.Text & "  ··œ›⁄Â  " & .TextMatrix(i, .ColIndex("InstallmentNo")), val(notes_id), , , , NoteDate, user_id, , , , , , , , , setfoxy_Line, , , , , , , , , BranchID) = False Then
                    GoTo ErrTrap
                End If

                line_no = line_no + 1
                
                              If ModAccounts.AddNewDev(LngDevID, line_no, DeptSide1, val(.TextMatrix(i, .ColIndex("Value"))) + val(.TextMatrix(i, .ColIndex("FATValue"))), 1, Msg & "  ··⁄ﬁœ  " & dcM.Text & "  ··œ›⁄Â  " & .TextMatrix(i, .ColIndex("InstallmentNo")), val(notes_id), , , , NoteDate, user_id, , , , , , , , , setfoxy_Line, , , , , , , , , CustomerBranchId) = False Then
                    GoTo ErrTrap
                End If

                line_no = line_no + 1
                
                
     
                
Else
            If ModAccounts.AddNewDev(LngDevID, line_no, StrAccountCode, val(.TextMatrix(i, .ColIndex("Value"))) + val(.TextMatrix(i, .ColIndex("FATValue"))), 0, Msg & "  ··⁄ﬁœ  " & dcM.Text & "  ··œ›⁄Â  " & .TextMatrix(i, .ColIndex("InstallmentNo")), val(notes_id), , , , NoteDate, user_id, , , , , , , , , setfoxy_Line, , , , , , , , , BranchID) = False Then
                    GoTo ErrTrap
                End If

                line_no = line_no + 1
                
                
               If ModAccounts.AddNewDev(LngDevID, line_no, Account_Code_dynamic, .TextMatrix(i, .ColIndex("Value")), 1, Msg & "  ··⁄ﬁœ  " & dcM.Text & "  ··œ›⁄Â  " & .TextMatrix(i, .ColIndex("InstallmentNo")), val(notes_id), , , , NoteDate, user_id, , , , , , , , , setfoxy_Line, , , , , , , , , BranchID) = False Then
                    GoTo ErrTrap
                End If

                line_no = line_no + 1
                
                If val(.TextMatrix(i, .ColIndex("FATYou"))) <> 0 Then
               
               If ModAccounts.AddNewDev(LngDevID, line_no, .TextMatrix(i, .ColIndex("AccountCodeVat")), .TextMatrix(i, .ColIndex("FATValue")), 1, "Õ”«» «·ﬁÌ„… «·„÷«›…" & "  ··⁄ﬁœ  " & "Õ”«» «·ﬁÌ„… «·„÷«›…" & "  ··œ›⁄Â  " & .TextMatrix(i, .ColIndex("InstallmentNo")), val(notes_id), , , , NoteDate, user_id, , , , , , , , , setfoxy_Line, , , , , , , , , BranchID) = False Then
                    GoTo ErrTrap
                End If
                 line_no = line_no + 1
   End If
                
                End If
                
                
                
                
                
            End If
     
     
     Next i
     
     End With
           
    updateNotesValueAndNobytext (val(notes_id))
    Exit Function
ErrTrap:
   End Function
Function createVoucher()
Dim NoteID As Long
Dim NoteDate As Date
Dim NoteSerial As String
Dim Notevalue As Double
Dim notytype As Integer
Dim des As String
des = EleHeader.Caption & " —ﬁ„ " & TxtId & " » «—ÌŒ" & RecDate
Dim tablename As String
Dim Filedname As String
Dim NoteSerial1 As Long
Dim BranchID As Integer
 

Dim sql As String
tablename = "TblRequest_MinistryContract"
Filedname = "ID"
NoteSerial1 = val(TxtId)
Notevalue = 0

 notytype = 8067
'Notevalue = val(total)
 

 BranchID = val(dcBranch.BoundText)
 
 Dim i As Integer
    With Grid


 
        For i = .FixedRows To .Rows - 1
     
    
            If .TextMatrix(i, .ColIndex("Value")) > 0 And .TextMatrix(i, .ColIndex("Account_Code")) <> "" And .Cell(flexcpChecked, i, .ColIndex("Status")) = flexChecked Then   'C?C??? C???E??E IC??
 NoteDate = .TextMatrix(i, .ColIndex("Due_Date"))
        End If
        
        Next i
        
    End With
 
'NoteDate = Me.Date.value
 
'If Notevalue > 0 Then
                                If Me.TxtModFlg = "N" Then
                                      CreateNotes NoteID, NoteDate, BranchID, notytype, Notevalue, NoteSerial, (NoteSerial1), tablename, Filedname, NoteSerial1, des         ', recordDateH.value
                                              TXTNoteID.Text = NoteID
                                                     TxtNoteSerial.Text = NoteSerial
                                     Else
                                                 If TXTNoteID.Text = "" Or TxtNoteSerial.Text = "" Then
                                            CreateNotes NoteID, NoteDate, BranchID, notytype, Notevalue, NoteSerial, (NoteSerial1), tablename, Filedname, NoteSerial1, des   ', recordDateH.value
                                                                 TXTNoteID.Text = NoteID
                                                                TxtNoteSerial.Text = NoteSerial
                                                   Else
                                                                 sql = "update notes  set Note_Value=" & Notevalue & ",note_value_by_characters='" & WriteNo(val(Notevalue), 0, True) & "'"
                                                                sql = sql & ",NoteSerial1='" & (NoteSerial1) & "'"
                                                                   sql = sql & " where NoteID=" & val(TXTNoteID.Text)
                                                                   Cn.Execute sql
                                                               
                                                 End If
                                       
                                End If

CREATE_VOUCHER_GE val(TXTNoteID.Text), BranchID, user_id, NoteDate
rs.Resync adAffectCurrent
 

'     End If

End Function

Private Sub Date_Change()
'If Me.TxtModFlg.text = "N" Or Me.TxtModFlg.text = "E" Then
'TxtNoteSerial.text = ""
'End If
End Sub

Private Sub Dcbranch_Click(Area As Integer)
Dcbranch_Change
'Command1_Click
End Sub


Private Sub Dcbranch_Change()
If Me.TxtModFlg.Text = "N" Or Me.TxtModFlg.Text = "E" Then
TxtNoteSerial.Text = ""
End If


End Sub

 



Private Sub dcBranch_KeyUp(KeyCode As Integer, _
                           Shift As Integer)

 
End Sub

Private Sub DCEmP_KeyUp(KeyCode As Integer, _
                        Shift As Integer)

 

End Sub

Private Sub Option1_Click()
 
End Sub


 
Private Sub Command1_Click()
ProgressBar1.Visible = True
'If check_reg = True Then
' Exit Sub
' End If
Fill_Grid
ClculteVAT
ProgressBar1.Visible = False
ProgressBar1.value = 0
End Sub

Private Function check_reg() As Boolean

Dim query As String, count As Integer, count1 As Integer, str As String, j As Integer


    str = "  select   h.type , h.id HID , h.FromDate HFromDate ,h.FromDateH  HFromDateH , h.ToDate HToDate , h.TODateH HTODateH ,"
    str = str & "  d.id DID , d.FromDate ,d.FromDateH ,d.ToDate , d.TODateH   from TblDurations  h , TblDurations_Details  d"
    str = str & "  where h.ID = d.DID "
    
    If dcMontth.BoundText <> "" Then
            str = str & " and d.id  = " & val(dcMontth.BoundText)
    Else
          
    End If
        
    Set RsTemp2 = New ADODB.Recordset
    RsTemp2.Open str, Cn, adOpenStatic, adLockOptimistic, adCmdText
    Dim type_ As Integer, FromDate   As String, ToDate As String, FromDateH As String, todateH As String
    
    If RsTemp2.RecordCount > 0 Then
    
         type_ = IIf(IsNull(RsTemp2("Type").value), 0, RsTemp2("Type").value)
               FromDate = IIf(IsNull(RsTemp2("FromDate").value), "", RsTemp2("FromDate").value)
               ToDate = IIf(IsNull(RsTemp2("toDate").value), "", RsTemp2("toDate").value)
               FromDateH = IIf(IsNull(RsTemp2("FromDateH").value), "", RsTemp2("FromDateH").value)
               todateH = IIf(IsNull(RsTemp2("ToDateH").value), "", RsTemp2("ToDateH").value)
       
        For j = 0 To RsTemp2.RecordCount - 1
       
                query = Fill_Query(type_, FromDate, ToDate, FromDateH, todateH)
                Set RsTemp = New ADODB.Recordset
                RsTemp.Open query, Cn, adOpenStatic, adLockOptimistic, adCmdText
                If RsTemp.RecordCount > 0 Then
                    count = count + RsTemp.RecordCount
                End If
        Next
    End If
    

str = " select * from TblRequest_MinistryContract  where durationid = " & val(dcDur.BoundText) & "  and Month =   " & val(dcMontth.BoundText)
Set RsTemp = New ADODB.Recordset
RsTemp.Open str, Cn, adOpenStatic, adLockOptimistic, adCmdText
If RsTemp.RecordCount > 0 Then
        count1 = RsTemp.RecordCount
        If count = count1 Then
            MsgBox (" „ «À»«  «·«” Õﬁ«ﬁ ·Â–Â «·› —… ")
            check_reg = True
        End If
Else
        check_reg = False
End If
End Function

Private Sub DcDur_Change()
Dim i As Integer, j As Integer, str As String
    i = val(dcDur.BoundText)
    
    If i > 0 Then
        str = "  select id , Name  from TblDurations_Details where did =   " & i
        fill_combo dcMontth, str
    Else
        str = "  select id , Name  from TblDurations_Details where did =   " & -1
        fill_combo dcMontth, str
    End If
End Sub

Private Sub dcDur_Click(Area As Integer)

'Command1_Click
End Sub

Private Function Fill_Query(type_ As Integer, FromDate As String, ToDate As String, FromDateH As String, todateH As String) As String
Dim str As String


                
             str = str & "      SELECT dbo.TblMinistryContract_Installment.InstallmentNo, dbo.TblMinistryContract_Installment.Value, dbo.TblMinistryContract_Installment.Type,"
             str = str & "       dbo.TblMinistryContract_Installment.Due_DateH, dbo.TblMinistryContract_Installment.Due_Date, dbo.TblMinistryContract_Installment.ID,"
             str = str & "        dbo.TblMinistryContract_Installment.IDMC, dbo.TblMinistryContract.Name, dbo.TblMinistryContract.VendorID, dbo.TblMinistryContract.FromDate,"
             str = str & "         dbo.TblMinistryContract.FromDateH, dbo.TblMinistryContract.ToDate, dbo.TblMinistryContract.ToDateH, dbo.TblMinistryContract.StartContractDate,"
             str = str & "         dbo.TblMinistryContract.EndContractDate, dbo.TblMinistryContract.BranchID, dbo.TblMinistryContract.ClientID, dbo.TblCustemers.CusName, dbo.TblCustemers.CusNamee,"
             str = str & "           dbo.TblCustemers.Account_Code , dbo.ACCOUNTS.account_serial, TblBranchesData.branch_ID , dbo.TblBranchesData.branch_name, dbo.TblBranchesData.branch_nameE"
             str = str & "   FROM     dbo.TblBranchesData RIGHT OUTER JOIN"
             str = str & "                 dbo.TblCustemers ON dbo.TblBranchesData.branch_id = dbo.TblCustemers.BranchId RIGHT OUTER JOIN"
             str = str & "             dbo.TblMinistryContract_Installment INNER JOIN"
             str = str & "             dbo.TblMinistryContract ON dbo.TblMinistryContract_Installment.IDMC = dbo.TblMinistryContract.IDMC ON"
             str = str & "           dbo.TblCustemers.CusID = dbo.TblMinistryContract.ClientID LEFT OUTER JOIN"
             str = str & "            dbo.ACCOUNTS ON dbo.TblCustemers.Account_Code = dbo.ACCOUNTS.Account_Code"
             str = str & "   Where ( TblMinistryContract_Installment.paid is null or TblMinistryContract_Installment.paid  = 0 ) and  (dbo.TblMinistryContract_Installment.Type = 1)"
                                
               str = str & "  and TblMinistryContract.idmc =  " & val(dcM.BoundText)
                
               If type_ = 0 Then
                            str = str & " and TblMinistryContract_Installment.Due_Date >= '" & FromDate & "'  and  TblMinistryContract_Installment.Due_Date <= '" & ToDate & "'"
               ElseIf type_ = 1 Then
                             str = str & " and TblMinistryContract_Installment.Due_DateH >= '" & FromDateH & "'  and  TblMinistryContract_Installment.Due_DateH <= '" & todateH & "'"
               End If
               
              If dcBranch.BoundText <> "" Then
                      str = str & "  and    TblMinistryContract.BranchID  = " & val(dcBranch.BoundText)
              End If

Fill_Query = str


End Function


Private Sub Fill_Grid()

    Dim i As Integer, j As Integer, str As String
    i = val(dcDur.BoundText)
   
    Grid.Rows = Grid.FixedRows
    
    
    str = "  select   h.type , h.id HID , h.FromDate HFromDate ,h.FromDateH  HFromDateH , h.ToDate HToDate , h.TODateH HTODateH ,"
    str = str & "  d.id DID , d.FromDate ,d.FromDateH ,d.ToDate , d.TODateH   from TblDurations  h , TblDurations_Details  d"
    str = str & "  where h.ID = d.DID "
    
    If dcMontth.BoundText <> "" Then
            str = str & " and d.id  = " & val(dcMontth.BoundText)
    Else
          
    End If
        
    Set RsTemp = New ADODB.Recordset
    RsTemp.Open str, Cn, adOpenStatic, adLockOptimistic, adCmdText
    Dim type_ As Integer, FromDate   As String, ToDate As String, FromDateH As String, todateH As String
    
    If RsTemp.RecordCount > 0 Then
        
        ProgressBar1.Max = RsTemp.RecordCount
                
        For j = 0 To RsTemp.RecordCount - 1
                
                 ProgressBar1.value = j
                
               type_ = IIf(IsNull(RsTemp("Type").value), 0, RsTemp("Type").value)
               FromDate = IIf(IsNull(RsTemp("FromDate").value), "", RsTemp("FromDate").value)
               ToDate = IIf(IsNull(RsTemp("toDate").value), "", RsTemp("toDate").value)
               FromDateH = IIf(IsNull(RsTemp("FromDateH").value), "", RsTemp("FromDateH").value)
               todateH = IIf(IsNull(RsTemp("ToDateH").value), "", RsTemp("ToDateH").value)
               
                
               Set RsTemp2 = New ADODB.Recordset
               RsTemp2.Open Fill_Query(type_, FromDate, ToDate, FromDateH, todateH), Cn, adOpenStatic, adLockOptimistic, adCmdText
               
               Dim s As Integer, m As Integer
               If RsTemp2.RecordCount > 0 Then
                       ' Grid.Rows = Grid.FixedRows
                        m = Grid.Rows
                        Grid.Rows = Grid.Rows + RsTemp2.RecordCount
                        For s = 0 To RsTemp2.RecordCount - 1
                        
                                    With Grid
                                                 
                                                If Registered_Before(IIf(IsNull(RsTemp2("ID").value), 0, RsTemp2("ID").value)) = True Then
                                                    .TextMatrix(m, .ColIndex("status")) = 1
                                                    Else
                                                     .TextMatrix(m, .ColIndex("status")) = 0
                                                End If
                                                 
                                                .TextMatrix(m, .ColIndex("ID")) = IIf(IsNull(RsTemp2("ID").value), "", RsTemp2("ID").value)
                                                .TextMatrix(m, .ColIndex("IDMC")) = IIf(IsNull(RsTemp2("IDMC").value), "", RsTemp2("IDMC").value)
                                                .TextMatrix(m, .ColIndex("Value")) = IIf(IsNull(RsTemp2("Value").value), "", RsTemp2("Value").value)
                                                .TextMatrix(m, .ColIndex("InstallmentNo")) = IIf(IsNull(RsTemp2("InstallmentNo").value), "", RsTemp2("InstallmentNo").value)
                                                                                            
                                                .TextMatrix(m, .ColIndex("StartContractDate")) = IIf(IsNull(RsTemp2("StartContractDate").value), "", RsTemp2("StartContractDate").value)
                                                .TextMatrix(m, .ColIndex("EndContractDate")) = IIf(IsNull(RsTemp2("EndContractDate").value), "", RsTemp2("EndContractDate").value)
                                                .TextMatrix(m, .ColIndex("FromDate")) = IIf(IsNull(RsTemp2("FromDate").value), "", RsTemp2("FromDate").value)
                                                .TextMatrix(m, .ColIndex("Due_Date")) = IIf(IsNull(RsTemp2("Due_Date").value), "", RsTemp2("Due_Date").value)
                                                .TextMatrix(m, .ColIndex("Due_DateH")) = IIf(IsNull(RsTemp2("Due_DateH").value), "", RsTemp2("Due_DateH").value)
                                                
                                                .TextMatrix(m, .ColIndex("Account_Serial")) = IIf(IsNull(RsTemp2("Account_Serial").value), "", RsTemp2("Account_Serial").value)
                                                .TextMatrix(m, .ColIndex("CusName")) = IIf(IsNull(RsTemp2("CusName").value), "", RsTemp2("CusName").value)
                                                .TextMatrix(m, .ColIndex("Account_Code")) = IIf(IsNull(RsTemp2("Account_Code").value), "", RsTemp2("Account_Code").value)
                                                
                                                .TextMatrix(m, .ColIndex("branchID")) = IIf(IsNull(RsTemp2("branch_ID").value), "", RsTemp2("branch_ID").value)
                                                .TextMatrix(m, .ColIndex("branch_name")) = IIf(IsNull(RsTemp2("branch_name").value), "", RsTemp2("branch_name").value)
                                        End With
                             RsTemp2.MoveNext
                             m = m + 1
                         Next
               End If
               RsTemp.MoveNext
        Next
    End If
    
    
calculation



End Sub

Private Function Registered_Before(ID As Integer) As Boolean
 Dim str As String
 str = "  select * from TblRequest_MinistryContract_Detailst where insid =  " & ID
 Set Rs_Temp = New ADODB.Recordset
 Rs_Temp.Open str, Cn, adOpenStatic, adLockOptimistic, adCmdText
 If Rs_Temp.RecordCount > 0 Then
        Registered_Before = True
        Exit Function
 End If
Registered_Before = False
End Function



Private Function GetDayRate(IDMC As Integer, FromDate As String, ToDate As String)

 Dim str As String, days As Integer, net As Double, Operation As String
 str = " select * from TblAttributionContract  where idac = " & IDMC
 Set Rs_Temp = New ADODB.Recordset
 Rs_Temp.Open str, Cn, adOpenStatic, adLockOptimistic, adCmdText
 If Rs_Temp.RecordCount > 0 Then
'        days = DateDiff("d", Rs_Temp("StartContractDate").value, Rs_Temp("EndContractDate").value)

        days = DateDiff("d", FromDate, ToDate)
        If days > 0 Then
                Operation = IIf(IsNull(Rs_Temp("AdditionalType").value), "", Rs_Temp("AdditionalType").value)
                If Operation = "add" Then
                       net = val(Rs_Temp("studentcount").value) * val(Rs_Temp("StudentCustom").value) + val(Rs_Temp("StudentCustom").value)
                ElseIf Operation = "sub" Then
                       net = val(Rs_Temp("studentcount").value) * val(Rs_Temp("StudentCustom").value) - val(Rs_Temp("StudentCustom").value)
                Else
                        net = val(Rs_Temp("studentcount").value) * val(Rs_Temp("StudentCustom").value)
                End If
                GetDayRate = net / days
        End If
 End If
 
End Function

Private Sub GetDeducts(IDMC As Integer, dur As Integer, MonthID As Integer, Row As Integer)

         Dim str As String, i As Integer, j As Integer
         str = "SELECT   dbo.TblConfirmViolation.ID, dbo.TblConfirmViolation.DurationID, dbo.TblConfirmViolation.ViolationID, dbo.TblConfirmViolation.MinistryContractID,"
         str = str & " dbo.TblConfirmViolation.Date , dbo.TblConfirmViolation.value, dbo.TblConfirmViolation.monthid, dbo.TblViolationTypes.name"
         str = str & " FROM     dbo.TblConfirmViolation INNER JOIN   dbo.TblViolationTypes ON dbo.TblConfirmViolation.ViolationID = dbo.TblViolationTypes.ID"
         str = str & " where DurationID = " & dur & " and MonthID = " & MonthID & "  and MinistryContractID =  " & IDMC
         
         Set RsTemp2 = New ADODB.Recordset
         RsTemp2.Open str, Cn, adOpenStatic, adLockOptimistic, adCmdText
            
         With Grid
         If RsTemp2.RecordCount > 0 Then
                For i = 1 To RsTemp2.RecordCount
                        For j = 1 To 10
                                If .TextMatrix(1, .ColIndex("d" & j)) = RsTemp2("Name").value Then
                                        .TextMatrix(Row, .ColIndex("d" & j)) = RsTemp2("Value").value
                                End If
                        Next
                Next
         End If
         End With
End Sub

Private Function GetHold(MonthID As Integer)
    Dim str As String, cunt As Integer
             str = " select count (*)  cunt, DDID  from TblVacationSchedule where ISVac = 1 and  ddid =   " & MonthID & " group by DDID "
             Set RsTemp2 = New ADODB.Recordset
             RsTemp2.Open str, Cn, adOpenStatic, adLockOptimistic, adCmdText
             
             If RsTemp2.RecordCount > 0 Then
                    cunt = IIf(IsNull(RsTemp2("cunt").value), 0, RsTemp2("cunt").value)
             End If
             
    GetHold = cunt
End Function

Private Sub GetVac(IDMC As Integer, MonthID As Integer, ByRef Total As Integer, ByRef daycount As Integer)

        Dim str As String, cunt As Integer, CityID As Integer, DurID As Integer, DayDiff As Integer, j As Integer, i As Integer
        Dim DurationID As Integer, SchoolFileID As Integer
        
        'str = " select CityID , DurationID  from  TblAttributionContract where IDAC = " & IDMC
        str = " select  d.schoolfileid , a.DurationID  from TblAttributionContract a ,  TblVehicleAllocation_Details  d where a.IDAC = d.IDVA   and d.type = 3  and  a.IDAC =  " & IDMC
        str = str & "   group by schoolfileid , DurationID  "
        
        Set RsTemp2 = New ADODB.Recordset
        RsTemp2.Open str, Cn, adOpenStatic, adLockOptimistic, adCmdText
        
       For j = 0 To RsTemp2.RecordCount - 1
                
                DurationID = IIf(IsNull(RsTemp2("DurationID").value), 0, RsTemp2("DurationID").value)
                SchoolFileID = IIf(IsNull(RsTemp2("schoolfileid").value), 0, RsTemp2("schoolfileid").value)
                
                str = " select h.DurationID , h.MonthID ,d.SchoolFileID ,sum (d.daycount) daycount , sum (d.dayvalue)  dayvalue , sum ( (d.daycount * d.dayvalue )) Total"
                str = str & " from TblconfirmVacation  h, TblConfirmVacation_Details d "
                str = str & " where  h.ID = d.HID and  DurationID = " & DurationID & "  and  MonthID = " & MonthID & " and  SchoolFileID = " & SchoolFileID
                str = str & " group by  h.DurationID , h.MonthID ,d.SchoolFileID  "
                Set RsTemp3 = New ADODB.Recordset
                RsTemp3.Open str, Cn, adOpenStatic, adLockOptimistic, adCmdText
                If RsTemp3.RecordCount > 0 Then
                    For i = 0 To RsTemp3.RecordCount - 1
                         Total = Total + IIf(IsNull(RsTemp3("Total").value), 0, RsTemp3("Total").value)
                         daycount = daycount + IIf(IsNull(RsTemp3("daycount").value), 0, RsTemp3("daycount").value)
                         RsTemp3.MoveNext
                   Next
                End If
                RsTemp2.MoveNext
       Next
          
          
     
End Sub

Private Function GetMonthDays(MonthID As Integer)

    Dim str As String, cunt As Integer
             str = " select count (*)  cunt, DDID  from TblVacationSchedule where   ddid =   " & MonthID & " group by DDID "
             Set RsTemp2 = New ADODB.Recordset
             RsTemp2.Open str, Cn, adOpenStatic, adLockOptimistic, adCmdText
             
             If RsTemp2.RecordCount > 0 Then
                    cunt = IIf(IsNull(RsTemp2("cunt").value), 0, RsTemp2("cunt").value)
             End If
             
    GetMonthDays = cunt

End Function

Private Sub dcM_Change()
Dim str As String
Set Rs_Temp = New ADODB.Recordset
str = " select * from TblMinistryContract where IDMC =  " & val(dcM.BoundText)
Rs_Temp.Open str, Cn, adOpenStatic, adLockOptimistic, adCmdText

If Rs_Temp.RecordCount > 0 Then
        dcBranch.BoundText = IIf(IsNull(Rs_Temp("branchid").value), "", Rs_Temp("branchid").value)
End If

End Sub

Private Sub dcM_Click(Area As Integer)
'Command1_Click
End Sub

Private Sub dcM_KeyUp(KeyCode As Integer, Shift As Integer)

If KeyCode = vbKeyF3 Then

        Unload FrmSearch_MinistryContract
        FrmSearch_MinistryContract.SendForm = "R"
        FrmSearch_MinistryContract.show
End If


End Sub

Private Sub dcMontth_Click(Area As Integer)
'Fill_Grid
'Command1_Click
End Sub

Private Sub Form_Activate()
'    XPTxtBoxID.SetFocus
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
    On Error GoTo ErrTrap
    Dim Dcombos As ClsDataCombos

    Set Dcombos = New ClsDataCombos
    'Dcombos.GetEmployees DcboGovernmentID
   ' Dcombos.getCountriesGovernments Me.DcboGovernmentID
    Dcombos.GetBranches dcBranch
    Dcombos.GetAccountingCodes AccountVat
    
    If SystemOptions.UserInterface = EnglishInterface Then
        SetInterface Me
        ChangeLang
    End If

    LogTextA = "   «·œŒÊ· «·Ì ‘«‘… " & "  «À»«  «·«” Õﬁ«ﬁ«  «·‘Â—Ì… ·⁄ﬁÊœ «·Ê“«—…  "
    LogTexte = " Open Window " & "  Boxes Data "
    AddToLogFile CInt(user_id), 0, Date, Time, LogTextA, LogTexte, Me.Name, "O", "", ""

    Dim My_SQL As String
    My_SQL = " Select id , name from  TblDurations "
    fill_combo dcDur, My_SQL
  
   Dim ss As String
   ss = "select IDMC , MinistryContractNo    from TblMinistryContract "
   fill_combo dcM, ss

    Resize_Form Me
    AddTip
    Set rs = New ADODB.Recordset
  
   Dim StrSQL As String
   StrSQL = "SELECT  *  From TblRequest_MinistryContract order by ID"
   rs.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText
      
   With cbType
        If SystemOptions.UserInterface = ArabicInterface Then
                .Clear
                .AddItem ("‰ﬁœÏ")
                .AddItem ("‘Ìﬂ")
        Else
                .Clear
                .AddItem ("Cash")
                .AddItem ("Cheque")
        End If
    End With
        
    Me.TxtModFlg.Text = "R"
    XPBtnMove_Click 2

    If OPEN_NEW_SCREEN = True Then
        Cmd_Click (0)
    End If
Intialize_Deducts

Inatial_Grid



    Exit Sub

ErrTrap:
End Sub

Private Sub Inatial_Grid()

 With Grid

        .MergeCells = flexMergeFixedOnly
        .MergeRow(0) = True
        
       ' .MergeCol(.ColIndex("No")) = True
        .Cell(flexcpText, 0, .ColIndex("No"), 1, .ColIndex("No")) = "—ﬁ„ «·”ÿ—"

        .MergeCol(.ColIndex("cusname")) = True
        .Cell(flexcpText, 0, .ColIndex("cusname"), 1, .ColIndex("cusname")) = "«·«”„"

        .MergeCol(.ColIndex("PayNo")) = True
        .Cell(flexcpText, 0, .ColIndex("PayNo"), 1, .ColIndex("PayNo")) = "—ﬁ„ «·œ›⁄…"

        .MergeCol(.ColIndex("Value")) = True
        .Cell(flexcpText, 0, .ColIndex("Value"), 1, .ColIndex("Value")) = "«·ﬁÌ„…"

        .MergeCol(.ColIndex("Total")) = True
        .Cell(flexcpText, 0, .ColIndex("total"), 1, .ColIndex("total")) = "«Ã„«·Ï «·„” Õﬁ« "
        
        .MergeCol(.ColIndex("Net")) = True
        .Cell(flexcpText, 0, .ColIndex("Net"), 1, .ColIndex("Net")) = "«·’«›Ï «·„” Õﬁ"
        .Cell(flexcpText, 0, .ColIndex("d1"), 0, .ColIndex("d10")) = "Õ”„Ì« "
 
    End With



End Sub



Private Sub Intialize_Deducts()
Dim str As String, i As Integer
Set Rs_Temp = New ADODB.Recordset
str = " select * from TblViolationTypes  "
Rs_Temp.Open str, Cn, adOpenStatic, adLockOptimistic, adCmdText
Rs_Temp.MoveFirst
If Rs_Temp.RecordCount > 0 Then
    For i = 1 To Rs_Temp.RecordCount
        Grid.TextMatrix(1, Grid.ColIndex("d" & i)) = IIf(IsNull(Rs_Temp("Name").value), "", Rs_Temp("Name").value)
        Rs_Temp.MoveNext
    Next
End If


For i = 1 To 10
     
      If Grid.TextMatrix(1, Grid.ColIndex("d" & i)) = "" Then
                Grid.ColWidth(Grid.ColIndex("d" & i)) = 0
      End If
Next


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

 
   lbl(0).Caption = "No."
   lbl(3).Caption = " Name Ar"
   lbl(7).Caption = " Name En"
   'Label3.Caption = "City"
   
  lbl(2).Caption = "Current Record"
  lbl(4).Caption = "Recors Count"
   
    Me.Caption = "Managerial Area"
    EleHeader.Caption = Me.Caption
   
    Me.Cmd(0).Caption = "New"
    Me.Cmd(1).Caption = "Edit"
    Me.Cmd(2).Caption = "Save"
    Me.Cmd(3).Caption = "Undo"
    Me.Cmd(4).Caption = "Delete"
    'Me.Cmd(5).Caption = "Search"
    Me.Cmd(6).Caption = "Exit"
    Me.Cmd(7).Caption = "Print"
    CmdAttach.Caption = "Attachment"

End Sub

Private Sub Form_Unload(Cancel As Integer)
    On Error GoTo ErrTrap
    LogTextA = "     «·Œ—ÊÃ „‰ ‘«‘… " & "   «À»«  «·«” Õﬁ«ﬁ«  «·‘Â—Ì… ·⁄ﬁÊœ «·Ê“«—…   "
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



Private Sub Grid_AfterEdit(ByVal Row As Long, ByVal Col As Long)

     With Grid
            Select Case .ColKey(Col)
                Case "d1", "d2", "d3", "d4", "d5", "d6", "d7", "d8", "d9", "d10", "Value"
                        .TextMatrix(Row, .ColIndex("Total")) = val(.TextMatrix(Row, .ColIndex("d1"))) + val(.TextMatrix(Row, .ColIndex("d2"))) + val(.TextMatrix(Row, .ColIndex("d3"))) + val(.TextMatrix(Row, .ColIndex("d4"))) + val(.TextMatrix(Row, .ColIndex("d5"))) + val(.TextMatrix(Row, .ColIndex("d6"))) + val(.TextMatrix(Row, .ColIndex("d7"))) + val(.TextMatrix(Row, .ColIndex("d8"))) + val(.TextMatrix(Row, .ColIndex("d9"))) + val(.TextMatrix(Row, .ColIndex("d10")))
                        .TextMatrix(Row, .ColIndex("Net")) = val(.TextMatrix(Row, .ColIndex("Value"))) - val(.TextMatrix(Row, .ColIndex("Total")))
            End Select
       End With
ClculteVAT
Relain
End Sub
Sub Relain()
Dim i As Integer
Dim SumVal As Double
Dim SumVAT As Double
SumVal = 0
SumVAT = 0
With Grid
For i = 1 To .Rows - 1
 If .Cell(flexcpChecked, i, .ColIndex("Status")) = flexChecked And .TextMatrix(i, .ColIndex("ID")) <> "" Then
 SumVal = SumVal + val(.TextMatrix(i, .ColIndex("Value")))
 SumVAT = SumVAT + val(.TextMatrix(i, .ColIndex("FATValue")))
 End If
Next i
End With
TxtFATValue.Text = SumVAT
Total.Text = SumVal
TxtTotalValue.Text = SumVal + SumVAT
End Sub
Sub ClculteVAT()
If Me.TxtModFlg.Text <> "R" Then
Dim Percetage As Double
Dim i As Integer
Dim account As String
With Grid
For i = 1 To .Rows - 1
If .TextMatrix(i, .ColIndex("Due_Date")) <> "" Then
PercentgValueAddedAccount_Transec .TextMatrix(i, .ColIndex("Due_Date")), 1, 1, account, Percetage
.TextMatrix(i, .ColIndex("AccountCodeVat")) = account
TxtFATYou.Text = Percetage
.TextMatrix(i, .ColIndex("FATYou")) = Percetage
If val(.TextMatrix(i, .ColIndex("FATYou"))) > 0 Then
.TextMatrix(i, .ColIndex("FATValue")) = (val(.TextMatrix(i, .ColIndex("Value"))) * val(.TextMatrix(i, .ColIndex("FATYou")))) / 100
Else
.TextMatrix(i, .ColIndex("FATValue")) = 0
End If
End If
.TextMatrix(i, .ColIndex("TotalValue")) = val(.TextMatrix(i, .ColIndex("FATValue"))) + val(.TextMatrix(i, .ColIndex("Value")))
Next i
End With
End If
End Sub

Private Sub calculation()
    Dim i As Integer
     With Grid
            For i = 2 To .Rows - 1
                        .TextMatrix(i, .ColIndex("Total")) = val(.TextMatrix(i, .ColIndex("d1"))) + val(.TextMatrix(i, .ColIndex("d2"))) + val(.TextMatrix(i, .ColIndex("d3"))) + val(.TextMatrix(i, .ColIndex("d4"))) + val(.TextMatrix(i, .ColIndex("d5"))) + val(.TextMatrix(i, .ColIndex("d6"))) + val(.TextMatrix(i, .ColIndex("d7"))) + val(.TextMatrix(i, .ColIndex("d8"))) + val(.TextMatrix(i, .ColIndex("d9"))) + val(.TextMatrix(i, .ColIndex("d10"))) + val(.TextMatrix(i, .ColIndex("VacValue")))
                        .TextMatrix(i, .ColIndex("Net")) = val(.TextMatrix(i, .ColIndex("Value"))) - val(.TextMatrix(i, .ColIndex("Total")))
            Next
       End With



End Sub




Private Sub TxtModFlg_Change()
    On Error GoTo ErrTrap

    Select Case Me.TxtModFlg.Text

        Case "R"

            If SystemOptions.UserInterface = ArabicInterface Then
                Me.Caption = " «À»«  «·«” Õﬁ«ﬁ«  «·‘Â—Ì… ·⁄ﬁÊœ «·Ê“«—… "
            Else
                Me.Caption = "Boxes Data"
            End If
            Me.Cmd(8).Enabled = False
            Me.Cmd(2).Enabled = False
            Me.Cmd(3).Enabled = False
        
            Me.Cmd(0).Enabled = True
            Me.Cmd(1).Enabled = True
            Me.Cmd(4).Enabled = True
            Me.Cmd(5).Enabled = True
        
            Me.XPBtnMove(0).Enabled = True
            Me.XPBtnMove(1).Enabled = True
            Me.XPBtnMove(2).Enabled = True
            Me.XPBtnMove(3).Enabled = True
        
            

            If rs.RecordCount < 1 Then
                Me.XPBtnMove(0).Enabled = False
                Me.XPBtnMove(1).Enabled = False
                Me.XPBtnMove(2).Enabled = False
                Me.XPBtnMove(3).Enabled = False
                Me.Cmd(1).Enabled = False
                Me.Cmd(4).Enabled = False
            End If
            
            C1Elastic1.Enabled = False
            'C1Elastic2.Enabled = False
            
        Case "N"

            If SystemOptions.UserInterface = ArabicInterface Then
                Me.Caption = "«À»«  «·«” Õﬁ«ﬁ«  «·‘Â—Ì… ·⁄ﬁÊœ «·Ê“«—… ( ÃœÌœ )"
            Else
                Me.Caption = "Exchange Request (New)"
            End If
            Me.Cmd(8).Enabled = True
            Me.Cmd(2).Enabled = True
            Me.Cmd(3).Enabled = True

            If SystemOptions.UserInterface = ArabicInterface Then
                Me.Caption = "«À»«  «·«” Õﬁ«ﬁ«  «·‘Â—Ì… ·⁄ﬁÊœ «·Ê“«—…( ÃœÌœ )"
            Else
                Me.Caption = "Exchange Request  (New)"
            End If
            Me.Cmd(0).Enabled = False
            Me.Cmd(1).Enabled = False
            Me.Cmd(4).Enabled = False
            Me.Cmd(5).Enabled = False
            
            C1Elastic1.Enabled = True
            C1Elastic2.Enabled = True
            
        Case "E"

            If SystemOptions.UserInterface = ArabicInterface Then
                Me.Caption = "«À»«  «·«” Õﬁ«ﬁ«  «·‘Â—Ì… ·⁄ﬁÊœ «·Ê“«—… (  ⁄œÌ· )"
            Else
                Me.Caption = "Exchange Request (Edit)"
            End If
            Me.Cmd(8).Enabled = True
            Me.Cmd(2).Enabled = True
            Me.Cmd(3).Enabled = True
            Me.Cmd(0).Enabled = False
            Me.Cmd(1).Enabled = False
            Me.Cmd(4).Enabled = False
            Me.Cmd(5).Enabled = False
              
            Me.XPBtnMove(0).Enabled = False
            Me.XPBtnMove(1).Enabled = False
            Me.XPBtnMove(2).Enabled = False
            Me.XPBtnMove(3).Enabled = False
            
            C1Elastic1.Enabled = False
       '     C1Elastic2.Enabled = False
    End Select

    Exit Sub
ErrTrap:
End Sub


Public Sub Retrive(Optional Lngid As Long = 0, Optional NoteID As Long = 0)

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
            rs.find "ID =" & Lngid, , adSearchForward, adBookmarkFirst

            If rs.EOF Or rs.BOF Then
                Exit Sub
            End If
               
        End If
        
        
             If NoteID <> 0 Then
            rs.find "NoteID =" & NoteID, , adSearchForward, adBookmarkFirst

            If rs.EOF Or rs.BOF Then
                Exit Sub
            End If
               
        End If
        
        
    End If
    Grid.Rows = Grid.FixedRows
    
    TxtId.Text = IIf(IsNull(rs("ID").value), "", (rs("ID").value))
    txtCode.Text = IIf(IsNull(rs("code").value), "", Trim(rs("code").value))
    cbType.ListIndex = IIf(IsNull(rs("ExchangeType").value), -1, Trim(rs("ExchangeType").value))
    dcDur.BoundText = IIf(IsNull(rs("DurationID").value), "", Trim(rs("DurationID").value))
    dcMontth.BoundText = IIf(IsNull(rs("Month").value), "", Trim(rs("Month").value))
    XPTxtCurrent.Caption = rs.AbsolutePosition
    XPTxtCount.Caption = rs.RecordCount
    
    dcBranch.BoundText = IIf(IsNull(rs("BranchID").value), "", rs("BranchID").value)
    Me.RecDate.value = IIf(IsNull(rs("Date").value), Date, rs("Date").value)
    Me.DateH.value = IIf(IsNull(rs("Dateh").value), Date, rs("Dateh").value)
    dcM.BoundText = IIf(IsNull(rs("MinstryID").value), "", rs("MinstryID").value)
    Me.TXTNoteID.Text = IIf(IsNull(rs.Fields("NoteID").value), "", rs.Fields("NoteID").value)
    Me.TxtNoteSerial.Text = IIf(IsNull(rs("NoteSerial").value), "", rs("NoteSerial").value)
        
    Me.Total.Text = IIf(IsNull(rs("total").value), "", rs("total").value)
    Me.TxtFATYou.Text = IIf(IsNull(rs("FATYou").value), "", rs("FATYou").value)
    Me.TxtFATValue.Text = IIf(IsNull(rs("FATValue").value), "", rs("FATValue").value)
    Me.TxtTotalValue.Text = IIf(IsNull(rs("TotalValue").value), "", rs("TotalValue").value)
    
   Dim str As String
   str = IIf(IsNull(rs("AllID").value), "", rs("AllID").value)

    If str = "" Then
            Exit Sub
    End If
    
   Dim i As Integer
   Set RsTemp = New ADODB.Recordset
   
  Dim ss As String

ss = "         SELECT   dbo.TblMinistryContract.IDMC, dbo.TblMinistryContract.ClientID, dbo.TblCustemers.BranchId, dbo.TblBranchesData.branch_name,"
ss = ss & "               dbo.TblBranchesData.branch_namee, dbo.TblCustemers.CusName, dbo.TblCustemers.CusNamee, dbo.TblCustemers.Account_Code, dbo.ACCOUNTS.Account_Serial,"
ss = ss & "                dbo.TblMinistryContract_Installment.ID AS IID, dbo.TblMinistryContract_Installment.InstallmentNo, dbo.TblMinistryContract_Installment.Value,"
ss = ss & "                dbo.TblMinistryContract_Installment.Due_Date, dbo.TblMinistryContract_Installment.Due_DateH, dbo.TblMinistryContract.FromDate, dbo.TblMinistryContract.FromDateH,"
ss = ss & "                dbo.TblMinistryContract.EndContractDate, dbo.TblMinistryContract.EndContractDateh, dbo.TblMinistryContract.StartContractDate,"
ss = ss & "               dbo.TblMinistryContract.StartContractDateh , dbo.TblMinistryContract.MinistryContractNo, dbo.TblCustemers.fullcode ,dbo.TblMinistryContract_Installment.FATValue,dbo.TblMinistryContract_Installment.FATYou,dbo.TblMinistryContract_Installment.TotalValue,dbo.TblMinistryContract_Installment.AccountCodeVat"
ss = ss & "     FROM     dbo.TblMinistryContract_Installment INNER JOIN"
ss = ss & "                        dbo.TblMinistryContract ON dbo.TblMinistryContract_Installment.IDMC = dbo.TblMinistryContract.IDMC LEFT OUTER JOIN"
ss = ss & "                       dbo.TblCustemers ON dbo.TblMinistryContract.ClientID = dbo.TblCustemers.CusID LEFT OUTER JOIN"
ss = ss & "                     dbo.TblBranchesData ON dbo.TblCustemers.BranchId = dbo.TblBranchesData.branch_id LEFT OUTER JOIN"
ss = ss & "                      dbo.ACCOUNTS ON dbo.TblCustemers.Account_Code = dbo.ACCOUNTS.Account_Code"
ss = ss & "    WHERE  (1 = 1) AND (dbo.TblMinistryContract_Installment.Type = 1) "
        
        ss = ss & "  and  TblMinistryContract_Installment.id in  ( " & str & "  )"

       ss = ss & "    order by TblMinistryContract.IDMC  "
   
       RsTemp.Open ss, Cn, adOpenStatic, adLockOptimistic, adCmdText
    
   If RsTemp.RecordCount > 0 Then
        With Grid
        RsTemp.MoveFirst
        Grid.Rows = .FixedRows + RsTemp.RecordCount
        For i = .FixedRows To .Rows - 1
                .TextMatrix(i, .ColIndex("Serial")) = i - 1
                .TextMatrix(i, .ColIndex("Status")) = 1
                .TextMatrix(i, .ColIndex("FATValue")) = IIf(IsNull(RsTemp("FATValue").value), 0, RsTemp("FATValue").value)
                .TextMatrix(i, .ColIndex("FATYou")) = IIf(IsNull(RsTemp("FATYou").value), 0, RsTemp("FATYou").value)
                .TextMatrix(i, .ColIndex("TotalValue")) = IIf(IsNull(RsTemp("TotalValue").value), 0, RsTemp("TotalValue").value)
                .TextMatrix(i, .ColIndex("AccountCodeVat")) = IIf(IsNull(RsTemp("AccountCodeVat").value), "", RsTemp("AccountCodeVat").value)
                .TextMatrix(i, .ColIndex("ID")) = IIf(IsNull(RsTemp("IID").value), "", RsTemp("IID").value)
                .TextMatrix(i, .ColIndex("ID")) = IIf(IsNull(RsTemp("IID").value), "", RsTemp("IID").value)
                .TextMatrix(i, .ColIndex("IDMC")) = IIf(IsNull(RsTemp("IDMC").value), "", RsTemp("IDMC").value)
                .TextMatrix(i, .ColIndex("CusID")) = IIf(IsNull(RsTemp("ClientID").value), "", RsTemp("ClientID").value)
                .TextMatrix(i, .ColIndex("fullcode")) = IIf(IsNull(RsTemp("fullcode").value), "", RsTemp("fullcode").value)
                .TextMatrix(i, .ColIndex("cusname")) = IIf(IsNull(RsTemp("cusname").value), "", RsTemp("cusname").value)
                .TextMatrix(i, .ColIndex("InstallmentNo")) = IIf(IsNull(RsTemp("InstallmentNo").value), "", RsTemp("InstallmentNo").value)
                .TextMatrix(i, .ColIndex("Value")) = IIf(IsNull(RsTemp("Value").value), "", RsTemp("Value").value)
                .TextMatrix(i, .ColIndex("Due_Date")) = IIf(IsNull(RsTemp("Due_Date").value), "", RsTemp("Due_Date").value)
                .TextMatrix(i, .ColIndex("Due_DateH")) = IIf(IsNull(RsTemp("Due_DateH").value), "", RsTemp("Due_DateH").value)
                .TextMatrix(i, .ColIndex("EndContractDate")) = IIf(IsNull(RsTemp("EndContractDate").value), "", RsTemp("EndContractDate").value)
                .TextMatrix(i, .ColIndex("StartContractDate")) = IIf(IsNull(RsTemp("StartContractDate").value), "", RsTemp("StartContractDate").value)
                .TextMatrix(i, .ColIndex("FromDate")) = IIf(IsNull(RsTemp("FromDate").value), "", RsTemp("FromDate").value)
                .TextMatrix(i, .ColIndex("Account_Serial")) = IIf(IsNull(RsTemp("Account_Serial").value), "", RsTemp("Account_Serial").value)
                .TextMatrix(i, .ColIndex("Account_Code")) = IIf(IsNull(RsTemp("Account_Code").value), "", RsTemp("Account_Code").value)
                .TextMatrix(i, .ColIndex("CusName")) = IIf(IsNull(RsTemp("CusName").value), "", RsTemp("CusName").value)
                .TextMatrix(i, .ColIndex("branch_name")) = IIf(IsNull(RsTemp("branch_name").value), "", RsTemp("branch_name").value)
                .TextMatrix(i, .ColIndex("branchID")) = IIf(IsNull(RsTemp("branchID").value), "", RsTemp("branchID").value)
                 RsTemp.MoveNext
                 
        Next
        End With
   End If
    
    
    Exit Sub
ErrTrap:
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
            If checkedRow = False Then
                MsgBox ("«Œ — «·œ›⁄«  «Ê·«")
                Exit Sub
        End If
        
        
        If dcDur.BoundText = "" Then
            MsgBox "„‰ ›÷·ﬂ √œŒ· «”„ «·› —… ", vbOKOnly + vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
            dcDur.SetFocus
            Exit Sub
        End If
        
    If dcM.BoundText = "" Then
            MsgBox "„‰ ›÷·ﬂ √œŒ· «·⁄ﬁœ «Ê·« ", vbOKOnly + vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
            dcM.SetFocus
            Exit Sub
        End If
      
      If dcBranch.BoundText = "" Then
            MsgBox "„‰ ›÷·ﬂ √œŒ· «·›—⁄ «Ê·« ", vbOKOnly + vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
            dcBranch.SetFocus
            Exit Sub
        End If
        
         If dcMontth.BoundText = "" Then
            MsgBox "„‰ ›÷·ﬂ √œŒ·  «·› —… «Ê·« ", vbOKOnly + vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
            dcMontth.SetFocus
            Exit Sub
        End If
        
        
        Select Case Me.TxtModFlg.Text
            Case "N"
                 rs.AddNew
                 TxtId.Text = CStr(new_id("TblRequest_MinistryContract", "ID", "", True))
            Case "E"
               ' strSQL = "delete From TblRequest_MinistryContract_Detailst where  HID =" & val(txtID.text)
               ' Cn.Execute strSQL, , adExecuteNoRecords
               Cancel_Paid
               StrSQL = "Delete From DOUBLE_ENTREY_VOUCHERS Where Notes_ID=" & val(Me.TXTNoteID.Text)
               Cn.Execute StrSQL, , adExecuteNoRecords
             End Select

        Cn.BeginTrans
        BeginTrans = True
        rs("ID").value = val(TxtId.Text)
        rs("Code").value = Trim(txtCode.Text)
        rs("ExchangeType").value = IIf(cbType.ListIndex = -1, Null, cbType.ListIndex)
        rs("DurationID").value = val(dcDur.BoundText)
        rs("DurationName").value = dcDur.Text
        rs("MinstryID").value = IIf(dcM.BoundText = "", Null, dcM.BoundText)
        rs("Month").value = IIf(dcMontth.BoundText = "", Null, dcMontth.BoundText)
        rs("BranchID").value = IIf(dcBranch.BoundText = "", Null, dcBranch.BoundText)
        rs("Date").value = Me.RecDate.value
        rs("DateH").value = Me.DateH.value
        rs("total").value = val(Total.Text)
        rs("FATYou").value = val(TxtFATYou.Text)
        rs("FATValue").value = val(TxtFATValue.Text)
        rs("TotalValue").value = val(TxtTotalValue.Text)
        
        
       Dim i As Integer, AllID As String
       With Grid
            For i = .FixedRows To .Rows - 1
               If .Cell(flexcpChecked, i, .ColIndex("Status")) = flexChecked And .TextMatrix(i, .ColIndex("ID")) <> "" Then
                       '  If i = .FixedRows Then
                       '          AllID = AllID & IIf(.TextMatrix(i, .ColIndex("ID")) = "", " ", .TextMatrix(i, .ColIndex("ID")))
                       '  Else
                                AllID = AllID & IIf(.TextMatrix(i, .ColIndex("ID")) = "", " ", ",  " & .TextMatrix(i, .ColIndex("ID")))
                       '  End If
                End If
            Next
        End With
          
        rs("AllID").value = mId$(AllID, 2)
        rs.update
        
        
       With Grid
            For i = .FixedRows To .Rows - 1
               If .Cell(flexcpChecked, i, .ColIndex("Status")) = flexChecked And .TextMatrix(i, .ColIndex("ID")) <> "" Then
                        Set RsTemp = New ADODB.Recordset
                        Dim m As String
                        m = "  select * from TblMinistryContract_Installment where id =  " & val(.TextMatrix(i, .ColIndex("ID")))
                        RsTemp.Open m, Cn, adOpenStatic, adLockOptimistic, adCmdText
                        If RsTemp.RecordCount > 0 Then
                                RsTemp("Paid").value = 1
                                RsTemp("RID").value = val(TxtId.Text)
                                RsTemp("FATValue").value = val(.TextMatrix(i, .ColIndex("FATValue")))
                                RsTemp("FATYou").value = val(.TextMatrix(i, .ColIndex("FATYou")))
                                RsTemp("TotalValue").value = val(.TextMatrix(i, .ColIndex("TotalValue")))
                                RsTemp("AccountCodeVat").value = .TextMatrix(i, .ColIndex("AccountCodeVat"))
                                RsTemp.update
                        End If
                End If
            Next
        End With
        
         
        
        Cn.CommitTrans
        BeginTrans = False
        XPTxtCurrent.Caption = rs.AbsolutePosition
        XPTxtCount.Caption = rs.RecordCount
       'CuurentLogdata
        createVoucher
        Select Case Me.TxtModFlg.Text

            Case "N"
                If SystemOptions.UserInterface = ArabicInterface Then
                    Msg = "  „ Õ›Ÿ «·»Ì«‰«    " & CHR(13)
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

Private Function checkedRow() As Boolean

Dim i As Integer, Check As Boolean
For i = 1 To Grid.Rows - 1
        If Grid.TextMatrix(i, Grid.ColIndex("status")) <> "" Then
                If Grid.TextMatrix(i, Grid.ColIndex("status")) <> 0 Then
                        checkedRow = True
                        Exit Function
                End If
        End If
Next
checkedRow = False

End Function

Private Sub Undo()
    On Error GoTo ErrTrap

    Select Case TxtModFlg.Text

        Case "N"
            clear_all Me
            Me.TxtModFlg.Text = "R"
            XPBtnMove_Click (1)

        Case "E"
            rs.find "ID='" & val(TxtId.Text) & "'", , adSearchForward, adBookmarkFirst

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
            
        If TxtId.Text <> "" Then

    
        Msg = "”Ì „ Õ–› »Ì«‰«   —ﬁ„ " & CHR(13)
        Msg = Msg + (TxtId.Text) & CHR(13)
        Msg = Msg + " Â·  —€» ›Ì Õ–› Â–Â «·»Ì«‰« ø"

        If MsgBox(Msg, vbYesNo + vbQuestion + vbMsgBoxRight + vbMsgBoxRtlReading, App.title) = vbYes Then
            
    
            If Not rs.RecordCount < 1 Then
                 
                 Cancel_Paid
                 
                  StrSQL = "Delete From DOUBLE_ENTREY_VOUCHERS Where Notes_ID=" & val(Me.TXTNoteID.Text)
                   Cn.Execute StrSQL, , adExecuteNoRecords


                StrSQL = "delete From TblRequest_MinistryContract_Detailst where  HID =" & val(TxtId.Text)
                Cn.Execute StrSQL, , adExecuteNoRecords
            
                StrSQL = "delete From TblRequest_MinistryContract where  ID =" & val(TxtId.Text)
                Cn.Execute StrSQL, , adExecuteNoRecords
                
                
                
                
                   StrSQL = "SELECT  *  From TblRequest_MinistryContract "
                    Grid.Rows = Grid.FixedRows
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
        Msg = "Â–Â «·⁄„·Ì… €Ì— „ «Õ… ÕÌÀ √‰Â ·«ÌÊÃœ √Ì ”Ã·« "
        MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
        TxtModFlg_Change
        Exit Sub
    End If

    TxtModFlg_Change
    Exit Sub
ErrTrap:
    'If Err.Number = -2147217887 Then
    Msg = "·«Ì„ﬂ‰ Õ–› Â–« «·”Ã· · ﬂ«„· «·»Ì«‰«  " & CHR(13) & " ÊÃœ »Ì«‰«  „— »ÿ… »Â–« «·Œ“‰… "
    Msg = Msg & CHR(13) & Err.description
    MsgBox Msg, vbMsgBoxRight + vbMsgBoxRtlReading + vbExclamation, App.title
    rs.CancelUpdate
    'End If
End Sub


Private Sub AddTip()
  '  Dim Wrap As String
  '  On Error GoTo ErrTrap
  '  Set TTP = New clstooltip
  '  Wrap = Chr(13) + Chr(10)
'
'    With TTP
'        .Create Me.hWnd, "»Ì«‰«  «·Œ“‰", 1, 15204351, -2147483630
'        .MaxWidth = 4000
'        .VisibleTime = 9000
'        .DelayTime = 600
'        .AddControl Cmd(0), "ÃœÌœ ..." & Wrap & "·«÷«›… »Ì«‰«  Œ“‰… ÃœÌœ" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
''    End With
'
'    With TTP
''        .Create Me.hWnd, "»Ì«‰«  «·Œ“‰", 1, 15204351, -2147483630
'        .MaxWidth = 4000
'        .VisibleTime = 9000
'        .DelayTime = 600
'        .AddControl Cmd(1), " ⁄œÌ· ..." & Wrap & "· ⁄œÌ· »Ì«‰«  «·Œ“‰…" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
'    End With
''
'    With TTP
'        .Create Me.hWnd, "»Ì«‰«  «·Œ“‰", 1, 15204351, -2147483630
'        .MaxWidth = 4000
'        .VisibleTime = 9000
'        .DelayTime = 600
'        .AddControl Cmd(2), "Õ›Ÿ ..." & Wrap & "·Õ›Ÿ »Ì«‰«  «·Œ“‰… «·ÃœÌœ" & Wrap & "·Õ›Ÿ «· ⁄œÌ·« " & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
'    End With
'
'    With TTP
'        .Create Me.hWnd, "»Ì«‰«  «·Œ“‰", 1, 15204351, -2147483630
'        .MaxWidth = 4000
'        .VisibleTime = 9000
'        .DelayTime = 600
'        .AddControl Cmd(3), " —«Ã⁄ ..." & Wrap & "·· —«Ã⁄ ⁄‰ ⁄„·Ì… «·«÷«›…" & Wrap & "··· —«Ã⁄ ⁄‰ ⁄„·Ì… «· ⁄œÌ·" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
'    End With
'
'    With TTP
'        .Create Me.hWnd, "»Ì«‰«  «·Œ“‰", 1, 15204351, -2147483630
'        .MaxWidth = 4000
'        .VisibleTime = 9000
'        .DelayTime = 600
'        .AddControl Cmd(4), "Õ–› ..." & Wrap & "·Õ–› »Ì«‰«  Â–« «·Œ“‰…" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
'    End With
'
'    With TTP
'        .Create Me.hWnd, "»Ì«‰«  «·Œ“‰", 1, 15204351, -2147483630
'        .MaxWidth = 4000
'        .VisibleTime = 9000
'        .DelayTime = 600
'        .AddControl Cmd(5), "»ÕÀ ..." & Wrap & "···»ÕÀ ⁄‰ Œ“‰…" & Wrap & "Ì‰ÿ»ﬁ ⁄·ÌÂ ‘—Êÿ „⁄Ì‰…" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
'    End With
'
'    With TTP
'        .Create Me.hWnd, "»Ì«‰«  «·Œ“‰", 1, 15204351, -2147483630
'        .MaxWidth = 4000
'        .VisibleTime = 9000
'        .DelayTime = 600
'        .AddControl Cmd(6), "Œ—ÊÃ ..." & Wrap & "·«€·«ﬁ Â–Â «·‰«›–…" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
'    End With
'
'    With TTP
'        .Create Me.hWnd, "»Ì«‰«  «·Œ“‰", 1, 15204351, -2147483630
'        .MaxWidth = 4000
'        .VisibleTime = 9000
'        .DelayTime = 600
'        .AddControl XPBtnMove(1), "«·√Ê· ..." & Wrap & "··«‰ ﬁ«· «·Ï √Ê· ”Ã·" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
'    End With
'
'    With TTP
'        .Create Me.hWnd, "»Ì«‰«  «·Œ“‰", 1, 15204351, -2147483630
'        .MaxWidth = 4000
'        .VisibleTime = 9000
'        .DelayTime = 600
'        .AddControl XPBtnMove(0), "«·”«»ﬁ ..." & Wrap & "··«‰ ﬁ«· «·Ï «·”Ã· «·”«»ﬁ" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
'    End With
'
'    With TTP
'        .Create Me.hWnd, "»Ì«‰«  «·Œ“‰", 1, 15204351, -2147483630
'        .MaxWidth = 4000
'        .VisibleTime = 9000
'        .DelayTime = 600
'        .AddControl XPBtnMove(3), "«· «·Ì ..." & Wrap & "··«‰ ﬁ«· «·Ï «·”Ã· «· «·Ì" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
'    End With
'
'    With TTP
'        .Create Me.hWnd, "»Ì«‰«  «·Œ“‰", 1, 15204351, -2147483630
'        .MaxWidth = 4000
'        .VisibleTime = 9000
'        .DelayTime = 600
'        .AddControl XPBtnMove(2), "«·√ŒÌ— ..." & Wrap & "··«‰ ﬁ«· «·Ï ¬Œ— ”Ã·" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
'    End With
'
'    With TTP
'        .Create Me.hWnd, "»Ì«‰«  «·Œ“‰", 1, 15204351, -2147483630
'        .MaxWidth = 4000
'        .VisibleTime = 9000
'        .DelayTime = 600
'    '    .AddControl CmdHelp, "„”«⁄œ… ..." & Wrap & "·· ⁄—› ⁄·Ï ÊŸÌ›… Â–Â «·‰«›–…" & Wrap & "ÊﬂÌ›Ì… «· ⁄«„· „⁄Â«" & Wrap & "≈÷€ÿ Â‰«" & Wrap, True
'    End With
'
    Exit Sub
ErrTrap:
End Sub


Function print_report(Optional NoteSerial As Integer)
    On Error Resume Next
    Dim MySQL As String
    Dim RsData As New ADODB.Recordset
    Dim xApp As New CRAXDRT.Application
    Dim xReport As CRAXDRT.Report
    Dim CViewer As ClsReportViewer
    Dim StrReportTitle As String
    Dim StrFileName As String
    Dim Msg As String
    MySQL = " SELECT     dbo.TblMinistryContract_Installment.RID, dbo.TblRequest_MinistryContract.ID, dbo.TblRequest_MinistryContract.[Date], dbo.TblRequest_MinistryContract.DurationID, "
    MySQL = MySQL & "                  dbo.TblRequest_MinistryContract.[Month], dbo.TblRequest_MinistryContract.DateH, TblBranchesData_1.branch_id, dbo.TblCustemers.CusID,"
    MySQL = MySQL & "                  dbo.TblCustemers.CusName, dbo.TblCustemers.CusNamee, dbo.TblDurations_Details.Name AS MonthName, dbo.TblDurations.Name AS DurName,"
    MySQL = MySQL & "                  TblBranchesData_1.branch_name, TblBranchesData_1.branch_namee, dbo.TblMinistryContract.IDMC, dbo.TblMinistryContract_Installment.[Value],"
    MySQL = MySQL & "                  dbo.ACCOUNTS.Account_Serial, TblBranchesData_1.branch_name AS CBN, TblBranchesData_1.branch_namee AS CBNE,"
    MySQL = MySQL & "                  dbo.TblMinistryContract_Installment.InstallmentNo, dbo.TblCustemers.Account_Code, dbo.TblMinistryContract_Installment.Due_Date,"
    MySQL = MySQL & "                  dbo.TblMinistryContract_Installment.Due_DateH, dbo.TblMinistryContract.FromDate, dbo.TblMinistryContract.FromDateH, dbo.TblMinistryContract.StartContractDate,"
    MySQL = MySQL & "                  dbo.TblMinistryContract.StartContractDateh, dbo.TblMinistryContract.EndContractDate, dbo.TblMinistryContract.EndContractDateh, dbo.TblCustemers.Fullcode,"
    MySQL = MySQL & "                  dbo.TblMinistryContract_Installment.FATYou, dbo.TblMinistryContract_Installment.FATValue, dbo.TblMinistryContract_Installment.TotalValue,"
    MySQL = MySQL & "                  dbo.TblRequest_MinistryContract.total, dbo.TblRequest_MinistryContract.FATYou AS HFATYou, dbo.TblRequest_MinistryContract.FATYou AS HFATYou,"
    MySQL = MySQL & "                  dbo.TblRequest_MinistryContract.TotalValue AS HTotalValue , dbo.TblRequest_MinistryContract.FATValue AS TotalVat"
    MySQL = MySQL & "     FROM         dbo.TblBranchesData TblBranchesData_1 RIGHT OUTER JOIN"
    MySQL = MySQL & "                  dbo.TblDurations RIGHT OUTER JOIN"
    MySQL = MySQL & "                  dbo.TblDurations_Details RIGHT OUTER JOIN"
    MySQL = MySQL & "                  dbo.ACCOUNTS RIGHT OUTER JOIN"
    MySQL = MySQL & "                  dbo.TblCustemers INNER JOIN"
    MySQL = MySQL & "                  dbo.TblRequest_MinistryContract INNER JOIN"
    MySQL = MySQL & "                  dbo.TblMinistryContract_Installment ON dbo.TblRequest_MinistryContract.ID = dbo.TblMinistryContract_Installment.RID INNER JOIN"
    MySQL = MySQL & "                  dbo.TblMinistryContract ON dbo.TblRequest_MinistryContract.MinstryID = dbo.TblMinistryContract.IDMC ON"
    MySQL = MySQL & "                  dbo.TblCustemers.CusID = dbo.TblMinistryContract.ClientID ON dbo.ACCOUNTS.Account_Code = dbo.TblCustemers.Account_Code ON"
    MySQL = MySQL & "                  dbo.TblDurations_Details.ID = dbo.TblRequest_MinistryContract.[Month] LEFT OUTER JOIN"
    MySQL = MySQL & "                  dbo.TblBranchesData TblBranchesData_2 ON dbo.TblRequest_MinistryContract.BranchID = TblBranchesData_2.branch_id ON"
    MySQL = MySQL & "                  dbo.TblDurations.ID = dbo.TblRequest_MinistryContract.DurationID ON TblBranchesData_1.branch_id = dbo.TblCustemers.BranchId"
    MySQL = MySQL & "   Where  TblRequest_MinistryContract.ID = " & val(TxtId.Text)
  
     
    If SystemOptions.UserInterface = ArabicInterface Then
        StrFileName = App.path & "\Reports\REPORTS NEW\" & "rpt_MinistrRequestReceipt.rpt"
    Else
        StrFileName = App.path & "\Reports\REPORTS NEW\" & "rpt_MinistrRequestReceipt.rpt"
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
    
    Else
 
         xReport.ParameterFields(1).AddCurrentValue cCompanyInfo.ArabCompanyName ' RPTCompany_Name_Eng
   
    End If
    If SystemOptions.VATNoAccordActivity = False Then
    xReport.ParameterFields(2).AddCurrentValue cCompanyInfo.VATRegNo
    Else
    xReport.ParameterFields(2).AddCurrentValue GetRegVATNo(val(dcBranch.BoundText))
    End If
    
    xReport.ParameterFields(6).AddCurrentValue WriteNo(val(Me.TxtTotalValue), 0, True)
    
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


Private Sub Grid_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
With Grid
Select Case .ColKey(Col)

Case "Status"
            If Me.TxtModFlg.Text = "E" Or Me.TxtModFlg.Text = "N" Then
            Else
                    Cancel = True
            End If

Case "FATValue"
        Cancel = True
        
Case "FATYou"
        Cancel = True
        
Case "TotalValue"
        Cancel = True
         
Case "IDAC"
        Cancel = True
Case "MA"
         Cancel = True
Case "fullcode"
          Cancel = True

Case "fullcode"
          Cancel = True
          
          
Case "cusname"
          Cancel = True
          
Case "Account_Serial"
          Cancel = True
          
          
Case "BankAccount"
          Cancel = True

Case "IBAN"
          Cancel = True

Case "recordno"
          Cancel = True


Case "Car"
          Cancel = True


Case "DayRate"
          Cancel = True


Case "WorkDay"
          Cancel = True


Case "Value"
          Cancel = True


Case "VacDay"
          Cancel = True


Case "VacValue"
          Cancel = True


Case "AbsenceCount"
          Cancel = True

Case "Avalue"
          Cancel = True


Case "StartContractDate"
          Cancel = True


Case "EndContractDate"
          Cancel = True


Case "FromDate"
          Cancel = True

          
Case "Total"
          Cancel = True
          
          
Case "Net"
          Cancel = True
          
        
          
End Select
End With


End Sub


