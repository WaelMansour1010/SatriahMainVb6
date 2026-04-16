VERSION 5.00
Object = "{C115893A-A3BF-43AF-B28D-69DB846077F3}#1.0#0"; "vsflex8u.ocx"
Object = "{0AFE7BE0-11B7-4A3E-978D-D4501E9A57FE}#1.0#0"; "c1sizer.ocx"
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{28D9BF84-BC20-11D2-94CF-004005455FAA}#1.1#0"; "ImpulseAniLabel.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{FE5DCFAD-BC1D-11D2-94CF-004005455FAA}#1.4#0"; "ImpulseButton.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{396F7AC0-A0DD-11D3-93EC-00C0DFE7442A}#1.0#0"; "vbalIml6.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#12.0#0"; "SuiteCtrls.ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{49003D3A-66CD-11D7-A449-E937BE2D9041}#1.0#0"; "ALLBUTTONS.ocx"
Begin VB.Form FrmCashing1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "«·„ﬁ»Ê÷« "
   ClientHeight    =   8595
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   18780
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   178
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "FrmCashing1.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   8595
   ScaleWidth      =   18780
   Begin VB.CommandButton CMDSENDSMS 
      Caption         =   "«—”«· —”«·Â"
      Height          =   375
      Left            =   5040
      RightToLeft     =   -1  'True
      TabIndex        =   381
      Top             =   8160
      Width           =   855
   End
   Begin VB.TextBox txtoldvalue 
      Alignment       =   1  'Right Justify
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   0
      MaxLength       =   10
      RightToLeft     =   -1  'True
      TabIndex        =   132
      Top             =   8640
      Visible         =   0   'False
      Width           =   2685
   End
   Begin C1SizerLibCtl.C1Elastic Ele 
      Height          =   585
      Index           =   1
      Left            =   0
      TabIndex        =   26
      TabStop         =   0   'False
      Top             =   0
      Width           =   18825
      _cx             =   33205
      _cy             =   1032
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
      BackColor       =   12648447
      ForeColor       =   4210688
      FloodColor      =   6553600
      ForeColorDisabled=   -2147483631
      Caption         =   "«·„ﬁ»Ê÷«  "
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
      Begin VB.TextBox oldtxtNoteSerial1 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   7320
         RightToLeft     =   -1  'True
         TabIndex        =   123
         Top             =   120
         Visible         =   0   'False
         Width           =   1425
      End
      Begin VB.TextBox TxtModFlg 
         Alignment       =   1  'Right Justify
         BackColor       =   &H000000FF&
         Height          =   345
         Left            =   4950
         RightToLeft     =   -1  'True
         TabIndex        =   28
         Top             =   60
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.TextBox XPTxtID 
         Alignment       =   1  'Right Justify
         BackColor       =   &H000000FF&
         Height          =   315
         Left            =   5460
         Locked          =   -1  'True
         RightToLeft     =   -1  'True
         TabIndex        =   27
         TabStop         =   0   'False
         Top             =   90
         Visible         =   0   'False
         Width           =   495
      End
      Begin ImpulseButton.ISButton XPBtnMove 
         Height          =   375
         Index           =   0
         Left            =   1125
         TabIndex        =   29
         Top             =   90
         Width           =   495
         _ExtentX        =   873
         _ExtentY        =   661
         ButtonStyle     =   1
         ButtonPositionImage=   4
         Caption         =   ""
         FontName        =   "Arial"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ButtonImage     =   "FrmCashing1.frx":000C
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
         Height          =   375
         Index           =   2
         Left            =   60
         TabIndex        =   30
         Top             =   90
         Width           =   495
         _ExtentX        =   873
         _ExtentY        =   661
         ButtonStyle     =   1
         ButtonPositionImage=   4
         Caption         =   ""
         FontName        =   "Arial"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ButtonImage     =   "FrmCashing1.frx":03A6
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
         Height          =   375
         Index           =   1
         Left            =   1650
         TabIndex        =   31
         Top             =   90
         Width           =   495
         _ExtentX        =   873
         _ExtentY        =   661
         ButtonStyle     =   1
         ButtonPositionImage=   4
         Caption         =   ""
         FontName        =   "Arial"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ButtonImage     =   "FrmCashing1.frx":0740
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
         Height          =   375
         Index           =   3
         Left            =   585
         TabIndex        =   32
         Top             =   90
         Width           =   495
         _ExtentX        =   873
         _ExtentY        =   661
         ButtonStyle     =   1
         ButtonPositionImage=   4
         Caption         =   ""
         FontName        =   "Arial"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ButtonImage     =   "FrmCashing1.frx":0ADA
         ColorHighlight  =   4194304
         ColorHoverText  =   16777215
         ColorShadow     =   -2147483631
         ColorOutline    =   -2147483631
         DrawFocusRectangle=   0   'False
         DisabledImageStyle=   1
         ColorToggledHoverText=   16777215
         ColorTextShadow =   16777215
      End
      Begin ImpulseButton.ISButton Cmd 
         Height          =   375
         Index           =   8
         Left            =   2400
         TabIndex        =   33
         Top             =   1500
         Visible         =   0   'False
         Width           =   1545
         _ExtentX        =   2725
         _ExtentY        =   661
         ButtonStyle     =   1
         ButtonPositionImage=   1
         Caption         =   "«·⁄—÷ «·ÃœÊ·Ï"
         BackColor       =   14871017
         FontName        =   "Arial"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
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
      Begin MSAdodcLib.Adodc numbering 
         Height          =   585
         Left            =   1680
         Top             =   1320
         Visible         =   0   'False
         Width           =   2040
         _ExtentX        =   3598
         _ExtentY        =   1032
         ConnectMode     =   0
         CursorLocation  =   3
         IsolationLevel  =   -1
         ConnectionTimeout=   15
         CommandTimeout  =   30
         CursorType      =   3
         LockType        =   3
         CommandType     =   8
         CursorOptions   =   0
         CacheSize       =   50
         MaxRecords      =   0
         BOFAction       =   0
         EOFAction       =   0
         ConnectStringType=   1
         Appearance      =   1
         BackColor       =   -2147483643
         ForeColor       =   -2147483640
         Orientation     =   0
         Enabled         =   -1
         Connect         =   ""
         OLEDBString     =   ""
         OLEDBFile       =   ""
         DataSourceName  =   ""
         OtherAttributes =   ""
         UserName        =   ""
         Password        =   ""
         RecordSource    =   ""
         Caption         =   " Õ—Ìﬂ"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         _Version        =   393216
      End
      Begin MSAdodcLib.Adodc detect_no 
         Height          =   585
         Left            =   -360
         Top             =   1680
         Visible         =   0   'False
         Width           =   2040
         _ExtentX        =   3598
         _ExtentY        =   1032
         ConnectMode     =   0
         CursorLocation  =   3
         IsolationLevel  =   -1
         ConnectionTimeout=   15
         CommandTimeout  =   30
         CursorType      =   3
         LockType        =   3
         CommandType     =   8
         CursorOptions   =   0
         CacheSize       =   50
         MaxRecords      =   0
         BOFAction       =   0
         EOFAction       =   0
         ConnectStringType=   1
         Appearance      =   1
         BackColor       =   -2147483643
         ForeColor       =   -2147483640
         Orientation     =   0
         Enabled         =   -1
         Connect         =   ""
         OLEDBString     =   ""
         OLEDBFile       =   ""
         DataSourceName  =   ""
         OtherAttributes =   ""
         UserName        =   ""
         Password        =   ""
         RecordSource    =   ""
         Caption         =   " Õ—Ìﬂ"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         _Version        =   393216
      End
      Begin VB.Image ImgFavorites 
         Height          =   390
         Left            =   9600
         Picture         =   "FrmCashing1.frx":0E74
         Stretch         =   -1  'True
         Top             =   0
         Width           =   525
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
         Height          =   345
         Index           =   11
         Left            =   4200
         RightToLeft     =   -1  'True
         TabIndex        =   34
         Top             =   60
         Visible         =   0   'False
         Width           =   705
      End
   End
   Begin C1SizerLibCtl.C1Tab XPTab301 
      Height          =   7545
      Left            =   0
      TabIndex        =   21
      Top             =   480
      Width           =   18810
      _cx             =   33179
      _cy             =   13309
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
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
      Caption         =   "«·„ﬁ»Ê÷« |«Œ Ì«—  „” Œ·’«  «·„‘«—Ì⁄|«·œ›⁄« "
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
      Picture(0)      =   "FrmCashing1.frx":4ADC
      Flags(1)        =   2
      Begin C1SizerLibCtl.C1Elastic Ele 
         Height          =   7080
         Index           =   12
         Left            =   45
         TabIndex        =   22
         TabStop         =   0   'False
         Top             =   45
         Width           =   18720
         _cx             =   33020
         _cy             =   12488
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
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
         Begin VB.TextBox TxtTotalInsurances 
            Alignment       =   1  'Right Justify
            Height          =   315
            Left            =   5640
            RightToLeft     =   -1  'True
            TabIndex        =   388
            Top             =   1680
            Width           =   1515
         End
         Begin VB.Frame Frame13 
            Caption         =   "»Ì«‰«  «· ’›Ì«  Ê›Ê« Ì— «·ﬂÂ—»«¡"
            Height          =   975
            Left            =   390
            RightToLeft     =   -1  'True
            TabIndex        =   365
            Top             =   90
            Visible         =   0   'False
            Width           =   9255
            Begin VB.TextBox TxtPrice3 
               Alignment       =   1  'Right Justify
               Height          =   315
               Left            =   4320
               Locked          =   -1  'True
               RightToLeft     =   -1  'True
               TabIndex        =   375
               Top             =   360
               Width           =   975
            End
            Begin VB.TextBox TxtPrice2 
               Alignment       =   1  'Right Justify
               Height          =   315
               Left            =   6360
               Locked          =   -1  'True
               RightToLeft     =   -1  'True
               TabIndex        =   374
               Top             =   360
               Width           =   975
            End
            Begin VB.TextBox TxtPrice 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00FFFFC0&
               Height          =   315
               Left            =   120
               RightToLeft     =   -1  'True
               TabIndex        =   373
               Top             =   360
               Width           =   1215
            End
            Begin VB.TextBox TxtRemPrice 
               Alignment       =   1  'Right Justify
               Height          =   315
               Left            =   2400
               Locked          =   -1  'True
               RightToLeft     =   -1  'True
               TabIndex        =   372
               Top             =   360
               Width           =   975
            End
            Begin VB.TextBox Text39 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00C0C0FF&
               Height          =   315
               Left            =   2520
               Locked          =   -1  'True
               RightToLeft     =   -1  'True
               TabIndex        =   371
               Top             =   5040
               Width           =   975
            End
            Begin VB.TextBox Text20 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00C0C0FF&
               ForeColor       =   &H00000000&
               Height          =   315
               Left            =   360
               Locked          =   -1  'True
               RightToLeft     =   -1  'True
               TabIndex        =   370
               TabStop         =   0   'False
               Top             =   5760
               Width           =   1215
            End
            Begin VB.TextBox Text19 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00C0C0FF&
               Height          =   315
               Left            =   360
               Locked          =   -1  'True
               RightToLeft     =   -1  'True
               TabIndex        =   369
               TabStop         =   0   'False
               Top             =   5880
               Width           =   1215
            End
            Begin VB.TextBox Text18 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00C0C0FF&
               Height          =   315
               Left            =   360
               Locked          =   -1  'True
               RightToLeft     =   -1  'True
               TabIndex        =   368
               TabStop         =   0   'False
               Top             =   5880
               Width           =   1215
            End
            Begin VB.TextBox Text16 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00C0C0FF&
               Height          =   315
               Left            =   2640
               Locked          =   -1  'True
               RightToLeft     =   -1  'True
               TabIndex        =   367
               TabStop         =   0   'False
               Top             =   6120
               Width           =   1215
            End
            Begin VB.TextBox Text15 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00C0C0FF&
               Height          =   315
               Left            =   5280
               Locked          =   -1  'True
               RightToLeft     =   -1  'True
               TabIndex        =   366
               TabStop         =   0   'False
               Top             =   6120
               Width           =   1215
            End
            Begin VB.Label Label10 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "X"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   178
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H000000FF&
               Height          =   255
               Left            =   8280
               RightToLeft     =   -1  'True
               TabIndex        =   380
               Top             =   240
               Width           =   495
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               Caption         =   "«·„”œœ"
               Height          =   255
               Index           =   107
               Left            =   5280
               RightToLeft     =   -1  'True
               TabIndex        =   379
               ToolTipText     =   "Ì „  Õ„Ì· Â–« «·„’—Ê› ⁄·Ï «·⁄„Ê·«  «·»‰ﬂÌ…"
               Top             =   360
               Width           =   735
            End
            Begin VB.Label Label6 
               Alignment       =   2  'Center
               Caption         =   "«·ﬁÌ„…"
               Height          =   255
               Left            =   7440
               RightToLeft     =   -1  'True
               TabIndex        =   378
               ToolTipText     =   "Ì „  Õ„Ì· Â–« «·„’—Ê› ⁄·Ï «·⁄„Ê·«  «·»‰ﬂÌ…"
               Top             =   360
               Width           =   1455
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               Caption         =   "«·„œ›Ê⁄"
               Height          =   255
               Index           =   106
               Left            =   1440
               RightToLeft     =   -1  'True
               TabIndex        =   377
               ToolTipText     =   "Ì „  Õ„Ì· Â–« «·„’—Ê› ⁄·Ï «·⁄„Ê·«  «·»‰ﬂÌ…"
               Top             =   360
               Width           =   735
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               Caption         =   "«·„ »ﬁÌ"
               Height          =   255
               Index           =   105
               Left            =   3360
               RightToLeft     =   -1  'True
               TabIndex        =   376
               ToolTipText     =   "Ì „  Õ„Ì· Â–« «·„’—Ê› ⁄·Ï «·⁄„Ê·«  «·»‰ﬂÌ…"
               Top             =   360
               Width           =   735
            End
         End
         Begin VB.TextBox TxtAccount 
            Alignment       =   1  'Right Justify
            Height          =   315
            Left            =   9030
            RightToLeft     =   -1  'True
            TabIndex        =   363
            Top             =   5160
            Width           =   1545
         End
         Begin VB.Frame Frame12 
            BackColor       =   &H00E2E9E9&
            Caption         =   "‰Ê⁄ «·Õ—ﬂ…"
            Enabled         =   0   'False
            Height          =   615
            Left            =   9030
            RightToLeft     =   -1  'True
            TabIndex        =   338
            Top             =   1350
            Width           =   2175
            Begin VB.OptionButton RdTypeTrans 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "›Ê« Ì— ﬂÂ—»«¡"
               Height          =   195
               Index           =   1
               Left            =   0
               RightToLeft     =   -1  'True
               TabIndex        =   340
               Top             =   360
               Width           =   1455
            End
            Begin VB.OptionButton RdTypeTrans 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   " ’›Ì« "
               Height          =   195
               Index           =   0
               Left            =   240
               RightToLeft     =   -1  'True
               TabIndex        =   339
               Top             =   120
               Value           =   -1  'True
               Width           =   1215
            End
         End
         Begin VB.Frame Frame11 
            Caption         =   "»Ì«‰«  «· ’›Ì«  Ê›Ê« Ì— «·ﬂÂ—»«¡"
            Height          =   4815
            Left            =   2910
            RightToLeft     =   -1  'True
            TabIndex        =   257
            Top             =   1770
            Visible         =   0   'False
            Width           =   9255
            Begin VB.TextBox TxtRemPaints 
               Alignment       =   1  'Right Justify
               Height          =   315
               Left            =   2640
               Locked          =   -1  'True
               RightToLeft     =   -1  'True
               TabIndex        =   351
               TabStop         =   0   'False
               Top             =   2040
               Width           =   975
            End
            Begin VB.TextBox TxtRemMaintClean 
               Alignment       =   1  'Right Justify
               Height          =   315
               Left            =   2640
               Locked          =   -1  'True
               RightToLeft     =   -1  'True
               TabIndex        =   350
               Top             =   1680
               Width           =   975
            End
            Begin VB.TextBox TxtRemMaintCondition 
               Alignment       =   1  'Right Justify
               Height          =   315
               Left            =   2640
               Locked          =   -1  'True
               RightToLeft     =   -1  'True
               TabIndex        =   349
               Top             =   1320
               Width           =   975
            End
            Begin VB.TextBox TxtRemRemainRent 
               Alignment       =   1  'Right Justify
               Height          =   315
               Left            =   2640
               Locked          =   -1  'True
               RightToLeft     =   -1  'True
               TabIndex        =   348
               Top             =   960
               Width           =   975
            End
            Begin VB.TextBox TxtRemMaintkitchen 
               Alignment       =   1  'Right Justify
               Height          =   315
               Left            =   2640
               Locked          =   -1  'True
               RightToLeft     =   -1  'True
               TabIndex        =   347
               Top             =   2400
               Width           =   975
            End
            Begin VB.TextBox TxtRemElectricity 
               Alignment       =   1  'Right Justify
               Height          =   315
               Left            =   2640
               Locked          =   -1  'True
               RightToLeft     =   -1  'True
               TabIndex        =   346
               Top             =   2760
               Width           =   975
            End
            Begin VB.TextBox TxtRemMaintDoors 
               Alignment       =   1  'Right Justify
               Height          =   315
               Left            =   2640
               Locked          =   -1  'True
               RightToLeft     =   -1  'True
               TabIndex        =   345
               Top             =   3120
               Width           =   975
            End
            Begin VB.TextBox TxtRemWindows 
               Alignment       =   1  'Right Justify
               Height          =   315
               Left            =   2640
               Locked          =   -1  'True
               RightToLeft     =   -1  'True
               TabIndex        =   344
               Top             =   3480
               Width           =   975
            End
            Begin VB.TextBox TxtRemMaintOther 
               Alignment       =   1  'Right Justify
               Height          =   315
               Left            =   2640
               Locked          =   -1  'True
               RightToLeft     =   -1  'True
               TabIndex        =   343
               TabStop         =   0   'False
               Top             =   3840
               Width           =   975
            End
            Begin VB.TextBox TxtRemMaintenance 
               Alignment       =   1  'Right Justify
               Height          =   315
               Left            =   2640
               Locked          =   -1  'True
               RightToLeft     =   -1  'True
               TabIndex        =   341
               Top             =   600
               Width           =   975
            End
            Begin VB.TextBox TxtNet2 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00C0C0FF&
               Height          =   315
               Left            =   5280
               Locked          =   -1  'True
               RightToLeft     =   -1  'True
               TabIndex        =   333
               TabStop         =   0   'False
               Top             =   6120
               Width           =   1215
            End
            Begin VB.TextBox TxtNet3 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00C0C0FF&
               Height          =   315
               Left            =   2640
               Locked          =   -1  'True
               RightToLeft     =   -1  'True
               TabIndex        =   332
               TabStop         =   0   'False
               Top             =   6120
               Width           =   1215
            End
            Begin VB.TextBox TxtNet 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00C0C0FF&
               Height          =   315
               Left            =   2640
               Locked          =   -1  'True
               RightToLeft     =   -1  'True
               TabIndex        =   331
               TabStop         =   0   'False
               Top             =   4200
               Width           =   975
            End
            Begin VB.TextBox TxtTotalAftreIns 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00C0C0FF&
               Height          =   315
               Left            =   360
               Locked          =   -1  'True
               RightToLeft     =   -1  'True
               TabIndex        =   329
               TabStop         =   0   'False
               Top             =   5880
               Width           =   1215
            End
            Begin VB.TextBox TxtTotalAftreIns3 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00C0C0FF&
               Height          =   315
               Left            =   360
               Locked          =   -1  'True
               RightToLeft     =   -1  'True
               TabIndex        =   328
               TabStop         =   0   'False
               Top             =   5880
               Width           =   1215
            End
            Begin VB.TextBox TxtTotalAftreIns2 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00C0C0FF&
               ForeColor       =   &H00000000&
               Height          =   315
               Left            =   360
               Locked          =   -1  'True
               RightToLeft     =   -1  'True
               TabIndex        =   327
               TabStop         =   0   'False
               Top             =   5760
               Width           =   1215
            End
            Begin VB.TextBox TxtTotal 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00C0C0FF&
               Height          =   315
               Left            =   360
               Locked          =   -1  'True
               RightToLeft     =   -1  'True
               TabIndex        =   325
               TabStop         =   0   'False
               Top             =   4200
               Width           =   1215
            End
            Begin VB.TextBox TxtTotal23 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00C0C0FF&
               Height          =   315
               Left            =   4560
               Locked          =   -1  'True
               RightToLeft     =   -1  'True
               TabIndex        =   324
               TabStop         =   0   'False
               Top             =   4200
               Width           =   975
            End
            Begin VB.TextBox TxtTotal22 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00C0C0FF&
               Height          =   315
               Left            =   6600
               Locked          =   -1  'True
               RightToLeft     =   -1  'True
               TabIndex        =   323
               TabStop         =   0   'False
               Top             =   4200
               Width           =   975
            End
            Begin VB.TextBox TxtMaintOther 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00FFFFC0&
               Height          =   315
               Left            =   360
               RightToLeft     =   -1  'True
               TabIndex        =   307
               TabStop         =   0   'False
               Top             =   3840
               Width           =   1215
            End
            Begin VB.TextBox TxtWindows 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00FFFFC0&
               Height          =   315
               Left            =   360
               RightToLeft     =   -1  'True
               TabIndex        =   306
               Top             =   3480
               Width           =   1215
            End
            Begin VB.TextBox TxtMaintDoors 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00FFFFC0&
               Height          =   315
               Left            =   360
               RightToLeft     =   -1  'True
               TabIndex        =   305
               Top             =   3120
               Width           =   1215
            End
            Begin VB.TextBox TxtElectricity1 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00FFFFC0&
               Height          =   315
               Left            =   360
               RightToLeft     =   -1  'True
               TabIndex        =   304
               Top             =   2760
               Width           =   1215
            End
            Begin VB.TextBox TxtMaintkitchen 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00FFFFC0&
               Height          =   315
               Left            =   360
               RightToLeft     =   -1  'True
               TabIndex        =   303
               Top             =   2400
               Width           =   1215
            End
            Begin VB.TextBox TxtMaintOther3 
               Alignment       =   1  'Right Justify
               Height          =   315
               Left            =   4560
               Locked          =   -1  'True
               RightToLeft     =   -1  'True
               TabIndex        =   302
               TabStop         =   0   'False
               Top             =   3840
               Width           =   975
            End
            Begin VB.TextBox TxtWindows3 
               Alignment       =   1  'Right Justify
               Height          =   315
               Left            =   4560
               Locked          =   -1  'True
               RightToLeft     =   -1  'True
               TabIndex        =   301
               Top             =   3480
               Width           =   975
            End
            Begin VB.TextBox TxtMaintDoors3 
               Alignment       =   1  'Right Justify
               Height          =   315
               Left            =   4560
               Locked          =   -1  'True
               RightToLeft     =   -1  'True
               TabIndex        =   300
               Top             =   3120
               Width           =   975
            End
            Begin VB.TextBox TxtElectricity13 
               Alignment       =   1  'Right Justify
               Height          =   315
               Left            =   4560
               Locked          =   -1  'True
               RightToLeft     =   -1  'True
               TabIndex        =   299
               Top             =   2760
               Width           =   975
            End
            Begin VB.TextBox TxtMaintkitchen3 
               Alignment       =   1  'Right Justify
               Height          =   315
               Left            =   4560
               Locked          =   -1  'True
               RightToLeft     =   -1  'True
               TabIndex        =   298
               Top             =   2400
               Width           =   975
            End
            Begin VB.TextBox TxtMaintOther2 
               Alignment       =   1  'Right Justify
               Height          =   315
               Left            =   6600
               Locked          =   -1  'True
               RightToLeft     =   -1  'True
               TabIndex        =   297
               TabStop         =   0   'False
               Top             =   3840
               Width           =   975
            End
            Begin VB.TextBox TxtWindows2 
               Alignment       =   1  'Right Justify
               Height          =   315
               Left            =   6600
               Locked          =   -1  'True
               RightToLeft     =   -1  'True
               TabIndex        =   296
               TabStop         =   0   'False
               Top             =   3480
               Width           =   975
            End
            Begin VB.TextBox TxtMaintDoors2 
               Alignment       =   1  'Right Justify
               Height          =   315
               Left            =   6600
               Locked          =   -1  'True
               RightToLeft     =   -1  'True
               TabIndex        =   295
               TabStop         =   0   'False
               Top             =   3120
               Width           =   975
            End
            Begin VB.TextBox TxtElectricity12 
               Alignment       =   1  'Right Justify
               Height          =   315
               Left            =   6600
               Locked          =   -1  'True
               RightToLeft     =   -1  'True
               TabIndex        =   294
               TabStop         =   0   'False
               Top             =   2760
               Width           =   975
            End
            Begin VB.TextBox TxtMaintkitchen2 
               Alignment       =   1  'Right Justify
               Height          =   315
               Left            =   6600
               Locked          =   -1  'True
               RightToLeft     =   -1  'True
               TabIndex        =   293
               TabStop         =   0   'False
               Top             =   2400
               Width           =   975
            End
            Begin VB.TextBox TxtDiscount 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00C0C0FF&
               Height          =   315
               Left            =   2520
               Locked          =   -1  'True
               RightToLeft     =   -1  'True
               TabIndex        =   274
               Top             =   5040
               Width           =   975
            End
            Begin VB.TextBox TxtMaintCondition 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00FFFFC0&
               Height          =   315
               Left            =   360
               RightToLeft     =   -1  'True
               TabIndex        =   273
               Top             =   1320
               Width           =   1215
            End
            Begin VB.TextBox TxtMaintenance 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00FFFFC0&
               Height          =   315
               Left            =   360
               RightToLeft     =   -1  'True
               TabIndex        =   272
               Top             =   600
               Width           =   1215
            End
            Begin VB.TextBox TxtRemainRent 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00FFFFC0&
               Height          =   315
               Left            =   360
               RightToLeft     =   -1  'True
               TabIndex        =   271
               Top             =   960
               Width           =   1215
            End
            Begin VB.TextBox TxtMaintClean 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00FFFFC0&
               Height          =   315
               Left            =   360
               RightToLeft     =   -1  'True
               TabIndex        =   270
               Top             =   1680
               Width           =   1215
            End
            Begin VB.TextBox TxtPaints 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00FFFFC0&
               Height          =   315
               Left            =   360
               RightToLeft     =   -1  'True
               TabIndex        =   269
               TabStop         =   0   'False
               Top             =   2040
               Width           =   1215
            End
            Begin VB.TextBox TxtPaints2 
               Alignment       =   1  'Right Justify
               Height          =   315
               Left            =   6600
               Locked          =   -1  'True
               RightToLeft     =   -1  'True
               TabIndex        =   268
               TabStop         =   0   'False
               Top             =   2040
               Width           =   975
            End
            Begin VB.TextBox TxtMaintClean2 
               Alignment       =   1  'Right Justify
               Height          =   315
               Left            =   6600
               Locked          =   -1  'True
               RightToLeft     =   -1  'True
               TabIndex        =   267
               Top             =   1680
               Width           =   975
            End
            Begin VB.TextBox TxtMaintCondition2 
               Alignment       =   1  'Right Justify
               Height          =   315
               Left            =   6600
               Locked          =   -1  'True
               RightToLeft     =   -1  'True
               TabIndex        =   266
               Top             =   1320
               Width           =   975
            End
            Begin VB.TextBox TxtRemainRent2 
               Alignment       =   1  'Right Justify
               Height          =   315
               Left            =   6600
               Locked          =   -1  'True
               RightToLeft     =   -1  'True
               TabIndex        =   265
               Top             =   960
               Width           =   975
            End
            Begin VB.TextBox TxtMaintenance2 
               Alignment       =   1  'Right Justify
               Height          =   315
               Left            =   6600
               Locked          =   -1  'True
               RightToLeft     =   -1  'True
               TabIndex        =   264
               Top             =   600
               Width           =   975
            End
            Begin VB.TextBox TxtMaintenance3 
               Alignment       =   1  'Right Justify
               Height          =   315
               Left            =   4560
               Locked          =   -1  'True
               RightToLeft     =   -1  'True
               TabIndex        =   263
               Top             =   600
               Width           =   975
            End
            Begin VB.TextBox TxtRemainRent3 
               Alignment       =   1  'Right Justify
               Height          =   315
               Left            =   4560
               Locked          =   -1  'True
               RightToLeft     =   -1  'True
               TabIndex        =   262
               Top             =   960
               Width           =   975
            End
            Begin VB.TextBox TxtInsurance 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00C0C0FF&
               Height          =   315
               Left            =   2520
               Locked          =   -1  'True
               RightToLeft     =   -1  'True
               TabIndex        =   261
               Top             =   4800
               Width           =   975
            End
            Begin VB.TextBox TxtMaintCondition3 
               Alignment       =   1  'Right Justify
               Height          =   315
               Left            =   4560
               Locked          =   -1  'True
               RightToLeft     =   -1  'True
               TabIndex        =   260
               Top             =   1320
               Width           =   975
            End
            Begin VB.TextBox TxtMaintClean3 
               Alignment       =   1  'Right Justify
               Height          =   315
               Left            =   4560
               Locked          =   -1  'True
               RightToLeft     =   -1  'True
               TabIndex        =   259
               Top             =   1680
               Width           =   975
            End
            Begin VB.TextBox TxtPaints3 
               Alignment       =   1  'Right Justify
               Height          =   315
               Left            =   4560
               Locked          =   -1  'True
               RightToLeft     =   -1  'True
               TabIndex        =   258
               TabStop         =   0   'False
               Top             =   2040
               Width           =   975
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               Caption         =   "«·«Ã„«·Ì"
               Height          =   255
               Index           =   103
               Left            =   3600
               RightToLeft     =   -1  'True
               TabIndex        =   361
               ToolTipText     =   "Ì „  Õ„Ì· Â–« «·„’—Ê› ⁄·Ï «·⁄„Ê·«  «·»‰ﬂÌ…"
               Top             =   4200
               Width           =   735
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               Caption         =   "«·„ »ﬁÌ"
               Height          =   255
               Index           =   102
               Left            =   3600
               RightToLeft     =   -1  'True
               TabIndex        =   360
               ToolTipText     =   "Ì „  Õ„Ì· Â–« «·„’—Ê› ⁄·Ï «·⁄„Ê·«  «·»‰ﬂÌ…"
               Top             =   3840
               Width           =   735
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               Caption         =   "«·„ »ﬁÌ"
               Height          =   255
               Index           =   101
               Left            =   3600
               RightToLeft     =   -1  'True
               TabIndex        =   359
               ToolTipText     =   "Ì „  Õ„Ì· Â–« «·„’—Ê› ⁄·Ï «·⁄„Ê·«  «·»‰ﬂÌ…"
               Top             =   3480
               Width           =   735
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               Caption         =   "«·„ »ﬁÌ"
               Height          =   255
               Index           =   100
               Left            =   3600
               RightToLeft     =   -1  'True
               TabIndex        =   358
               ToolTipText     =   "Ì „  Õ„Ì· Â–« «·„’—Ê› ⁄·Ï «·⁄„Ê·«  «·»‰ﬂÌ…"
               Top             =   3120
               Width           =   735
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               Caption         =   "«·„ »ﬁÌ"
               Height          =   255
               Index           =   99
               Left            =   3600
               RightToLeft     =   -1  'True
               TabIndex        =   357
               ToolTipText     =   "Ì „  Õ„Ì· Â–« «·„’—Ê› ⁄·Ï «·⁄„Ê·«  «·»‰ﬂÌ…"
               Top             =   2760
               Width           =   735
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               Caption         =   "«·„ »ﬁÌ"
               Height          =   255
               Index           =   98
               Left            =   3600
               RightToLeft     =   -1  'True
               TabIndex        =   356
               ToolTipText     =   "Ì „  Õ„Ì· Â–« «·„’—Ê› ⁄·Ï «·⁄„Ê·«  «·»‰ﬂÌ…"
               Top             =   2400
               Width           =   735
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               Caption         =   "«·„ »ﬁÌ"
               Height          =   255
               Index           =   97
               Left            =   3600
               RightToLeft     =   -1  'True
               TabIndex        =   355
               ToolTipText     =   "Ì „  Õ„Ì· Â–« «·„’—Ê› ⁄·Ï «·⁄„Ê·«  «·»‰ﬂÌ…"
               Top             =   2040
               Width           =   735
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               Caption         =   "«·„ »ﬁÌ"
               Height          =   255
               Index           =   96
               Left            =   3600
               RightToLeft     =   -1  'True
               TabIndex        =   354
               ToolTipText     =   "Ì „  Õ„Ì· Â–« «·„’—Ê› ⁄·Ï «·⁄„Ê·«  «·»‰ﬂÌ…"
               Top             =   1680
               Width           =   735
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               Caption         =   "«·„ »ﬁÌ"
               Height          =   255
               Index           =   95
               Left            =   3600
               RightToLeft     =   -1  'True
               TabIndex        =   353
               ToolTipText     =   "Ì „  Õ„Ì· Â–« «·„’—Ê› ⁄·Ï «·⁄„Ê·«  «·»‰ﬂÌ…"
               Top             =   1320
               Width           =   735
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               Caption         =   "«·„ »ﬁÌ"
               Height          =   255
               Index           =   94
               Left            =   3600
               RightToLeft     =   -1  'True
               TabIndex        =   352
               ToolTipText     =   "Ì „  Õ„Ì· Â–« «·„’—Ê› ⁄·Ï «·⁄„Ê·«  «·»‰ﬂÌ…"
               Top             =   960
               Width           =   735
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               Caption         =   "«·„ »ﬁÌ"
               Height          =   255
               Index           =   92
               Left            =   3600
               RightToLeft     =   -1  'True
               TabIndex        =   342
               ToolTipText     =   "Ì „  Õ„Ì· Â–« «·„’—Ê› ⁄·Ï «·⁄„Ê·«  «·»‰ﬂÌ…"
               Top             =   600
               Width           =   735
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               Caption         =   "«·«Ã„«·Ì"
               Height          =   255
               Index           =   93
               Left            =   1680
               RightToLeft     =   -1  'True
               TabIndex        =   337
               ToolTipText     =   "Ì „  Õ„Ì· Â–« «·„’—Ê› ⁄·Ï «·⁄„Ê·«  «·»‰ﬂÌ…"
               Top             =   4200
               Width           =   735
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               Caption         =   "„œ›Ê⁄ „”»ﬁ«"
               Height          =   255
               Index           =   76
               Left            =   3720
               RightToLeft     =   -1  'True
               TabIndex        =   336
               ToolTipText     =   "Ì „  Õ„Ì· Â–« «·„’—Ê› ⁄·Ï «·⁄„Ê·«  «·»‰ﬂÌ…"
               Top             =   5760
               Width           =   1335
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               Caption         =   "„œ›Ê⁄ „”»ﬁ«"
               Height          =   255
               Index           =   75
               Left            =   3720
               RightToLeft     =   -1  'True
               TabIndex        =   335
               ToolTipText     =   "Ì „  Õ„Ì· Â–« «·„’—Ê› ⁄·Ï «·⁄„Ê·«  «·»‰ﬂÌ…"
               Top             =   4800
               Width           =   1335
            End
            Begin VB.Label Label1 
               Alignment       =   2  'Center
               BackColor       =   &H00E2E9E9&
               BackStyle       =   0  'Transparent
               Caption         =   "«·«Ã„«·Ì »⁄œ «·Œ’„"
               Height          =   255
               Index           =   26
               Left            =   6840
               RightToLeft     =   -1  'True
               TabIndex        =   334
               Top             =   6120
               Width           =   1455
            End
            Begin VB.Label Label1 
               Alignment       =   2  'Center
               BackColor       =   &H00E2E9E9&
               BackStyle       =   0  'Transparent
               Caption         =   "’«›Ì «·„œ›Ê⁄"
               Height          =   255
               Index           =   23
               Left            =   7800
               RightToLeft     =   -1  'True
               TabIndex        =   330
               Top             =   5040
               Width           =   1815
            End
            Begin VB.Label Label1 
               Alignment       =   2  'Center
               BackColor       =   &H00E2E9E9&
               BackStyle       =   0  'Transparent
               Caption         =   "«·«Ã„«·Ì"
               Height          =   255
               Index           =   22
               Left            =   8160
               RightToLeft     =   -1  'True
               TabIndex        =   326
               Top             =   4800
               Width           =   1455
            End
            Begin VB.Label Label1 
               Alignment       =   2  'Center
               BackColor       =   &H00E2E9E9&
               BackStyle       =   0  'Transparent
               Caption         =   "‰Ê«›–"
               Height          =   255
               Index           =   21
               Left            =   7680
               RightToLeft     =   -1  'True
               TabIndex        =   322
               Top             =   3480
               Width           =   1455
            End
            Begin VB.Label Label1 
               Alignment       =   2  'Center
               BackColor       =   &H00E2E9E9&
               BackStyle       =   0  'Transparent
               Caption         =   "«Œ—Ï"
               Height          =   255
               Index           =   20
               Left            =   7680
               RightToLeft     =   -1  'True
               TabIndex        =   321
               Top             =   3840
               Width           =   1455
            End
            Begin VB.Label Label1 
               Alignment       =   2  'Center
               BackColor       =   &H00E2E9E9&
               BackStyle       =   0  'Transparent
               Caption         =   "’Ì«‰… «»Ê«»"
               Height          =   255
               Index           =   19
               Left            =   7680
               RightToLeft     =   -1  'True
               TabIndex        =   320
               Top             =   3120
               Width           =   1455
            End
            Begin VB.Label Label1 
               Alignment       =   2  'Center
               BackColor       =   &H00E2E9E9&
               BackStyle       =   0  'Transparent
               Caption         =   "›« Ê—… «·ﬂÂ—»«¡"
               Height          =   255
               Index           =   18
               Left            =   7680
               RightToLeft     =   -1  'True
               TabIndex        =   319
               Top             =   2760
               Width           =   1455
            End
            Begin VB.Label Label1 
               Alignment       =   2  'Center
               BackColor       =   &H00E2E9E9&
               BackStyle       =   0  'Transparent
               Caption         =   "’Ì«‰… „ÿ»Œ"
               Height          =   255
               Index           =   17
               Left            =   7680
               RightToLeft     =   -1  'True
               TabIndex        =   318
               Top             =   2400
               Width           =   1455
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               Caption         =   "«·«Ã„«·Ì"
               Height          =   255
               Index           =   91
               Left            =   5520
               RightToLeft     =   -1  'True
               TabIndex        =   317
               ToolTipText     =   "Ì „  Õ„Ì· Â–« «·„’—Ê› ⁄·Ï «·⁄„Ê·«  «·»‰ﬂÌ…"
               Top             =   4200
               Width           =   735
            End
            Begin VB.Label lbl 
               Alignment       =   2  'Center
               Caption         =   "«·«Ã„«·Ì"
               Height          =   255
               Index           =   90
               Left            =   7680
               RightToLeft     =   -1  'True
               TabIndex        =   316
               ToolTipText     =   "Ì „  Õ„Ì· Â–« «·„’—Ê› ⁄·Ï «·⁄„Ê·«  «·»‰ﬂÌ…"
               Top             =   4200
               Width           =   1455
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               Caption         =   "«·„”œœ"
               Height          =   255
               Index           =   89
               Left            =   5520
               RightToLeft     =   -1  'True
               TabIndex        =   315
               ToolTipText     =   "Ì „  Õ„Ì· Â–« «·„’—Ê› ⁄·Ï «·⁄„Ê·«  «·»‰ﬂÌ…"
               Top             =   3840
               Width           =   735
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               Caption         =   "«·„”œœ"
               Height          =   255
               Index           =   73
               Left            =   5520
               RightToLeft     =   -1  'True
               TabIndex        =   314
               ToolTipText     =   "Ì „  Õ„Ì· Â–« «·„’—Ê› ⁄·Ï «·⁄„Ê·«  «·»‰ﬂÌ…"
               Top             =   3480
               Width           =   735
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               Caption         =   "«·„”œœ"
               Height          =   255
               Index           =   72
               Left            =   5520
               RightToLeft     =   -1  'True
               TabIndex        =   313
               ToolTipText     =   "Ì „  Õ„Ì· Â–« «·„’—Ê› ⁄·Ï «·⁄„Ê·«  «·»‰ﬂÌ…"
               Top             =   3120
               Width           =   735
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               Caption         =   "«·„œ›Ê⁄"
               Height          =   255
               Index           =   71
               Left            =   1680
               RightToLeft     =   -1  'True
               TabIndex        =   312
               ToolTipText     =   "Ì „  Õ„Ì· Â–« «·„’—Ê› ⁄·Ï «·⁄„Ê·«  «·»‰ﬂÌ…"
               Top             =   3480
               Width           =   735
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               Caption         =   "«·„œ›Ê⁄"
               Height          =   255
               Index           =   70
               Left            =   1680
               RightToLeft     =   -1  'True
               TabIndex        =   311
               ToolTipText     =   "Ì „  Õ„Ì· Â–« «·„’—Ê› ⁄·Ï «·⁄„Ê·«  «·»‰ﬂÌ…"
               Top             =   3120
               Width           =   735
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               Caption         =   "«·„œ›Ê⁄"
               Height          =   255
               Index           =   69
               Left            =   1680
               RightToLeft     =   -1  'True
               TabIndex        =   310
               ToolTipText     =   "Ì „  Õ„Ì· Â–« «·„’—Ê› ⁄·Ï «·⁄„Ê·«  «·»‰ﬂÌ…"
               Top             =   600
               Width           =   735
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               Caption         =   "«·„œ›Ê⁄"
               Height          =   255
               Index           =   68
               Left            =   1680
               RightToLeft     =   -1  'True
               TabIndex        =   309
               ToolTipText     =   "Ì „  Õ„Ì· Â–« «·„’—Ê› ⁄·Ï «·⁄„Ê·«  «·»‰ﬂÌ…"
               Top             =   960
               Width           =   735
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               Caption         =   "«·„œ›Ê⁄"
               Height          =   255
               Index           =   67
               Left            =   1680
               RightToLeft     =   -1  'True
               TabIndex        =   308
               ToolTipText     =   "Ì „  Õ„Ì· Â–« «·„’—Ê› ⁄·Ï «·⁄„Ê·«  «·»‰ﬂÌ…"
               Top             =   3840
               Width           =   735
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               Caption         =   "«·„œ›Ê⁄"
               Height          =   255
               Index           =   81
               Left            =   1680
               RightToLeft     =   -1  'True
               TabIndex        =   292
               ToolTipText     =   "Ì „  Õ„Ì· Â–« «·„’—Ê› ⁄·Ï «·⁄„Ê·«  «·»‰ﬂÌ…"
               Top             =   1680
               Width           =   735
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               Caption         =   "«·„œ›Ê⁄"
               Height          =   255
               Index           =   80
               Left            =   1680
               RightToLeft     =   -1  'True
               TabIndex        =   291
               ToolTipText     =   "Ì „  Õ„Ì· Â–« «·„’—Ê› ⁄·Ï «·⁄„Ê·«  «·»‰ﬂÌ…"
               Top             =   1320
               Width           =   735
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               Caption         =   "«·„œ›Ê⁄"
               Height          =   255
               Index           =   79
               Left            =   1680
               RightToLeft     =   -1  'True
               TabIndex        =   290
               ToolTipText     =   "Ì „  Õ„Ì· Â–« «·„’—Ê› ⁄·Ï «·⁄„Ê·«  «·»‰ﬂÌ…"
               Top             =   2760
               Width           =   735
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               Caption         =   "«·„œ›Ê⁄"
               Height          =   255
               Index           =   78
               Left            =   1680
               RightToLeft     =   -1  'True
               TabIndex        =   289
               ToolTipText     =   "Ì „  Õ„Ì· Â–« «·„’—Ê› ⁄·Ï «·⁄„Ê·«  «·»‰ﬂÌ…"
               Top             =   2400
               Width           =   735
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               Caption         =   "«·„œ›Ê⁄"
               Height          =   255
               Index           =   77
               Left            =   1680
               RightToLeft     =   -1  'True
               TabIndex        =   288
               ToolTipText     =   "Ì „  Õ„Ì· Â–« «·„’—Ê› ⁄·Ï «·⁄„Ê·«  «·»‰ﬂÌ…"
               Top             =   2040
               Width           =   735
            End
            Begin VB.Label Label7 
               Alignment       =   2  'Center
               Caption         =   "’Ì«‰… „ﬂÌ›« "
               Height          =   255
               Left            =   7680
               RightToLeft     =   -1  'True
               TabIndex        =   287
               ToolTipText     =   "Ì „  Õ„Ì· Â–« «·„’—Ê› ⁄·Ï «·⁄„Ê·«  «·»‰ﬂÌ…"
               Top             =   1320
               Width           =   1455
            End
            Begin VB.Label Label9 
               Alignment       =   2  'Center
               Caption         =   "‰Ÿ«›…"
               Height          =   255
               Left            =   7680
               RightToLeft     =   -1  'True
               TabIndex        =   286
               ToolTipText     =   "Ì „  Õ„Ì· Â–« «·„’—Ê› ⁄·Ï «·⁄„Ê·«  «·»‰ﬂÌ…"
               Top             =   1680
               Width           =   1455
            End
            Begin VB.Label Label11 
               Alignment       =   2  'Center
               Caption         =   "„ »ﬁÌ «ÌÃ«—"
               Height          =   255
               Left            =   7680
               RightToLeft     =   -1  'True
               TabIndex        =   285
               ToolTipText     =   "Ì „  Õ„Ì· Â–« «·„’—Ê› ⁄·Ï «·⁄„Ê·«  «·»‰ﬂÌ…"
               Top             =   960
               Width           =   1455
            End
            Begin VB.Label Label12 
               Alignment       =   2  'Center
               Caption         =   "’Ì«‰… ﬂÂ—»«¡/”»«ﬂ…"
               Height          =   255
               Left            =   7680
               RightToLeft     =   -1  'True
               TabIndex        =   284
               ToolTipText     =   "Ì „  Õ„Ì· Â–« «·„’—Ê› ⁄·Ï «·⁄„Ê·«  «·»‰ﬂÌ…"
               Top             =   600
               Width           =   1455
            End
            Begin VB.Label lbl 
               Alignment       =   2  'Center
               Caption         =   "œÂ«‰« "
               Height          =   255
               Index           =   74
               Left            =   7680
               TabIndex        =   283
               Top             =   2040
               Width           =   1455
            End
            Begin VB.Label Label40 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "X"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   178
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H000000FF&
               Height          =   255
               Left            =   8280
               RightToLeft     =   -1  'True
               TabIndex        =   282
               Top             =   240
               Width           =   495
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               Caption         =   "«·„”œœ"
               Height          =   255
               Index           =   82
               Left            =   5520
               RightToLeft     =   -1  'True
               TabIndex        =   281
               ToolTipText     =   "Ì „  Õ„Ì· Â–« «·„’—Ê› ⁄·Ï «·⁄„Ê·«  «·»‰ﬂÌ…"
               Top             =   1680
               Width           =   735
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               Caption         =   "«·„”œœ"
               Height          =   255
               Index           =   83
               Left            =   5520
               RightToLeft     =   -1  'True
               TabIndex        =   280
               ToolTipText     =   "Ì „  Õ„Ì· Â–« «·„’—Ê› ⁄·Ï «·⁄„Ê·«  «·»‰ﬂÌ…"
               Top             =   1320
               Width           =   735
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               Caption         =   "«·„”œœ"
               Height          =   255
               Index           =   84
               Left            =   5520
               RightToLeft     =   -1  'True
               TabIndex        =   279
               ToolTipText     =   "Ì „  Õ„Ì· Â–« «·„’—Ê› ⁄·Ï «·⁄„Ê·«  «·»‰ﬂÌ…"
               Top             =   2760
               Width           =   735
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               Caption         =   "«·„”œœ"
               Height          =   255
               Index           =   85
               Left            =   5520
               RightToLeft     =   -1  'True
               TabIndex        =   278
               ToolTipText     =   "Ì „  Õ„Ì· Â–« «·„’—Ê› ⁄·Ï «·⁄„Ê·«  «·»‰ﬂÌ…"
               Top             =   2400
               Width           =   735
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               Caption         =   "«·„”œœ"
               Height          =   255
               Index           =   86
               Left            =   5520
               RightToLeft     =   -1  'True
               TabIndex        =   277
               ToolTipText     =   "Ì „  Õ„Ì· Â–« «·„’—Ê› ⁄·Ï «·⁄„Ê·«  «·»‰ﬂÌ…"
               Top             =   2040
               Width           =   735
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               Caption         =   "«·„”œœ"
               Height          =   255
               Index           =   87
               Left            =   5520
               RightToLeft     =   -1  'True
               TabIndex        =   276
               ToolTipText     =   "Ì „  Õ„Ì· Â–« «·„’—Ê› ⁄·Ï «·⁄„Ê·«  «·»‰ﬂÌ…"
               Top             =   960
               Width           =   735
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               Caption         =   "«·„”œœ"
               Height          =   255
               Index           =   88
               Left            =   5520
               RightToLeft     =   -1  'True
               TabIndex        =   275
               ToolTipText     =   "Ì „  Õ„Ì· Â–« «·„’—Ê› ⁄·Ï «·⁄„Ê·«  «·»‰ﬂÌ…"
               Top             =   600
               Width           =   735
            End
         End
         Begin VB.ComboBox DCboCashType2 
            Height          =   315
            ItemData        =   "FrmCashing1.frx":4E76
            Left            =   8370
            List            =   "FrmCashing1.frx":4E78
            RightToLeft     =   -1  'True
            Style           =   2  'Dropdown List
            TabIndex        =   256
            Top             =   960
            Width           =   2265
         End
         Begin VB.ComboBox CboPaymentType 
            Height          =   315
            Left            =   7920
            RightToLeft     =   -1  'True
            Style           =   2  'Dropdown List
            TabIndex        =   246
            Top             =   2880
            Width           =   2685
         End
         Begin VB.Frame Frame8 
            BackColor       =   &H00E2E9E9&
            Caption         =   "«·› —Â"
            Height          =   1095
            Left            =   120
            RightToLeft     =   -1  'True
            TabIndex        =   233
            Top             =   5280
            Width           =   3495
            Begin Dynamic_Byte.NourHijriCal FrmPriodDateH 
               Height          =   315
               Left            =   120
               TabIndex        =   234
               Top             =   240
               Width           =   1335
               _ExtentX        =   2355
               _ExtentY        =   556
            End
            Begin MSComCtl2.DTPicker FrmPriodDate 
               Height          =   315
               Left            =   1470
               TabIndex        =   235
               Top             =   240
               Width           =   1545
               _ExtentX        =   2725
               _ExtentY        =   556
               _Version        =   393216
               Format          =   163708929
               CurrentDate     =   41640
            End
            Begin Dynamic_Byte.NourHijriCal ToPriodDateH 
               Height          =   315
               Left            =   120
               TabIndex        =   236
               Top             =   720
               Width           =   1335
               _ExtentX        =   2355
               _ExtentY        =   556
            End
            Begin MSComCtl2.DTPicker ToPriodDate 
               Height          =   315
               Left            =   1470
               TabIndex        =   237
               Top             =   720
               Width           =   1545
               _ExtentX        =   2725
               _ExtentY        =   556
               _Version        =   393216
               Format          =   163708929
               CurrentDate     =   41640
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "«·Ï"
               Height          =   285
               Index           =   64
               Left            =   3120
               RightToLeft     =   -1  'True
               TabIndex        =   239
               Top             =   720
               Width           =   285
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "„‰"
               Height          =   285
               Index           =   63
               Left            =   3120
               RightToLeft     =   -1  'True
               TabIndex        =   238
               Top             =   240
               Width           =   285
            End
         End
         Begin VB.TextBox tXtRemarks 
            Alignment       =   1  'Right Justify
            Height          =   555
            Left            =   3930
            MultiLine       =   -1  'True
            RightToLeft     =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   231
            Top             =   5520
            Width           =   6645
         End
         Begin VB.Frame Frame6 
            Height          =   3135
            Left            =   -90
            RightToLeft     =   -1  'True
            TabIndex        =   227
            Top             =   2250
            Width           =   3855
            Begin VSFlex8UCtl.VSFlexGrid VSFlexGrid1 
               Height          =   2565
               Left            =   120
               TabIndex        =   228
               Top             =   240
               Width           =   3645
               _cx             =   6429
               _cy             =   4524
               Appearance      =   1
               BorderStyle     =   1
               Enabled         =   -1  'True
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Arial"
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
               Rows            =   1
               Cols            =   33
               FixedRows       =   1
               FixedCols       =   1
               RowHeightMin    =   320
               RowHeightMax    =   0
               ColWidthMin     =   0
               ColWidthMax     =   0
               ExtendLastCol   =   -1  'True
               FormatString    =   $"FrmCashing1.frx":4E7A
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
               WallPaperAlignment=   9
               AccessibleName  =   ""
               AccessibleDescription=   ""
               AccessibleValue =   ""
               AccessibleRole  =   24
            End
            Begin ImpulseButton.ISButton Cmd 
               Height          =   270
               Index           =   14
               Left            =   3120
               TabIndex        =   229
               Top             =   2760
               Width           =   690
               _ExtentX        =   1217
               _ExtentY        =   476
               ButtonStyle     =   1
               ButtonPositionImage=   1
               Caption         =   "Õ–›"
               FontName        =   "Arial"
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   178
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ButtonImage     =   "FrmCashing1.frx":536C
               DrawFocusRectangle=   0   'False
            End
         End
         Begin VB.Frame Frame9 
            Height          =   3255
            Left            =   0
            RightToLeft     =   -1  'True
            TabIndex        =   224
            Top             =   2160
            Width           =   3855
            Begin VSFlex8UCtl.VSFlexGrid VSFlexGrid2 
               Height          =   2685
               Left            =   0
               TabIndex        =   225
               Top             =   240
               Width           =   3645
               _cx             =   6429
               _cy             =   4736
               Appearance      =   1
               BorderStyle     =   1
               Enabled         =   -1  'True
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Arial"
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
               Rows            =   1
               Cols            =   33
               FixedRows       =   1
               FixedCols       =   1
               RowHeightMin    =   320
               RowHeightMax    =   0
               ColWidthMin     =   0
               ColWidthMax     =   0
               ExtendLastCol   =   -1  'True
               FormatString    =   $"FrmCashing1.frx":5906
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
               WallPaperAlignment=   9
               AccessibleName  =   ""
               AccessibleDescription=   ""
               AccessibleValue =   ""
               AccessibleRole  =   24
            End
            Begin ImpulseButton.ISButton Cmd 
               Height          =   270
               Index           =   13
               Left            =   3000
               TabIndex        =   226
               Top             =   2880
               Width           =   690
               _ExtentX        =   1217
               _ExtentY        =   476
               ButtonStyle     =   1
               ButtonPositionImage=   1
               Caption         =   "Õ–›"
               FontName        =   "Arial"
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   178
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ButtonImage     =   "FrmCashing1.frx":5DEF
               DrawFocusRectangle=   0   'False
            End
         End
         Begin ImpulseButton.ISButton ISButton3 
            Height          =   375
            Left            =   5160
            TabIndex        =   217
            TabStop         =   0   'False
            ToolTipText     =   "«÷€ÿ ·«÷«›… ⁄„Ì· ÃœÌœ"
            Top             =   960
            Width           =   510
            _ExtentX        =   900
            _ExtentY        =   661
            ButtonStyle     =   1
            ButtonPositionImage=   4
            Caption         =   ""
            BackColor       =   14871017
            FontName        =   "Arial"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   178
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BackStyle       =   0
            ButtonImage     =   "FrmCashing1.frx":6389
            ColorButton     =   14871017
            ColorHighlight  =   16777215
            ColorShadow     =   -2147483631
            ColorOutline    =   -2147483631
            DrawFocusRectangle=   0   'False
         End
         Begin VB.TextBox TXtFilter 
            Alignment       =   1  'Right Justify
            Height          =   375
            Left            =   3240
            RightToLeft     =   -1  'True
            TabIndex        =   219
            Top             =   960
            Width           =   1215
         End
         Begin VB.TextBox TxtFilterNo 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFC0&
            Enabled         =   0   'False
            Height          =   315
            Left            =   5760
            RightToLeft     =   -1  'True
            TabIndex        =   216
            Top             =   960
            Width           =   1515
         End
         Begin C1SizerLibCtl.C1Elastic C1Elastic1 
            Height          =   4455
            Left            =   12180
            TabIndex        =   171
            TabStop         =   0   'False
            Top             =   2430
            Width           =   6495
            _cx             =   11456
            _cy             =   7858
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
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
            Caption         =   "FrmCashing1"
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
            Begin VB.TextBox txtVATPayed 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00FFFFFF&
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   178
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   315
               Left            =   3150
               MaxLength       =   50
               RightToLeft     =   -1  'True
               TabIndex        =   386
               Top             =   2850
               Width           =   2025
            End
            Begin VB.OptionButton ComResid 
               Alignment       =   1  'Right Justify
               Caption         =   "”ﬂ‰Ì"
               ForeColor       =   &H00FF0000&
               Height          =   195
               Index           =   0
               Left            =   1440
               RightToLeft     =   -1  'True
               TabIndex        =   385
               Top             =   120
               Width           =   975
            End
            Begin VB.OptionButton ComResid 
               Alignment       =   1  'Right Justify
               Caption         =   " Ã«—Ì"
               ForeColor       =   &H00FF0000&
               Height          =   195
               Index           =   1
               Left            =   480
               RightToLeft     =   -1  'True
               TabIndex        =   384
               Top             =   120
               Width           =   975
            End
            Begin VB.TextBox TxtService 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00FFFFFF&
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   178
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   315
               Left            =   3120
               MaxLength       =   50
               RightToLeft     =   -1  'True
               TabIndex        =   254
               Top             =   2160
               Width           =   2025
            End
            Begin VB.TextBox Txtownerid 
               Height          =   495
               Left            =   5640
               TabIndex        =   253
               Top             =   3480
               Visible         =   0   'False
               Width           =   375
            End
            Begin VB.Frame Frame10 
               Enabled         =   0   'False
               Height          =   615
               Left            =   90
               TabIndex        =   249
               Top             =   3720
               Width           =   5055
               Begin VB.TextBox TxtKickbacks 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00FFFFFF&
                  BeginProperty Font 
                     Name            =   "Arial"
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
                  TabIndex        =   252
                  Top             =   240
                  Width           =   1545
               End
               Begin VB.OptionButton Rd 
                  Alignment       =   1  'Right Justify
                  Caption         =   "⁄„Ê·…"
                  Height          =   435
                  Index           =   1
                  Left            =   1800
                  RightToLeft     =   -1  'True
                  TabIndex        =   251
                  Top             =   120
                  Width           =   1455
               End
               Begin VB.OptionButton Rd 
                  Alignment       =   1  'Right Justify
                  Caption         =   "œ›⁄« "
                  Height          =   435
                  Index           =   0
                  Left            =   3360
                  RightToLeft     =   -1  'True
                  TabIndex        =   250
                  Top             =   120
                  Width           =   1455
               End
            End
            Begin VB.TextBox TxtElectricity 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00FFFFFF&
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   178
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   315
               Left            =   3120
               MaxLength       =   50
               RightToLeft     =   -1  'True
               TabIndex        =   240
               Top             =   2520
               Width           =   2025
            End
            Begin VB.TextBox TxtCommissionOut 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00FFFFFF&
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   178
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   315
               Left            =   120
               MaxLength       =   50
               RightToLeft     =   -1  'True
               TabIndex        =   221
               Top             =   1800
               Width           =   2025
            End
            Begin VB.Frame Frame7 
               Height          =   735
               Left            =   150
               TabIndex        =   210
               Top             =   3000
               Width           =   5055
               Begin XtremeSuiteControls.CheckBox CheckStatusEarnest 
                  Height          =   375
                  Index           =   0
                  Left            =   3720
                  TabIndex        =   211
                  Top             =   120
                  Width           =   1215
                  _Version        =   786432
                  _ExtentX        =   2143
                  _ExtentY        =   661
                  _StockProps     =   79
                  Caption         =   "«·€«¡ «·⁄—»Ê‰"
                  UseVisualStyle  =   -1  'True
                  TextAlignment   =   1
                  RightToLeft     =   -1  'True
               End
               Begin XtremeSuiteControls.CheckBox CheckStatusEarnest 
                  Height          =   375
                  Index           =   1
                  Left            =   0
                  TabIndex        =   212
                  Top             =   120
                  Width           =   1095
                  _Version        =   786432
                  _ExtentX        =   1931
                  _ExtentY        =   661
                  _StockProps     =   79
                  Caption         =   "—œ «·⁄—»Ê‰"
                  Enabled         =   0   'False
                  UseVisualStyle  =   -1  'True
                  TextAlignment   =   1
                  RightToLeft     =   -1  'True
               End
               Begin XtremeSuiteControls.CheckBox CheckStatusEarnest 
                  Height          =   375
                  Index           =   2
                  Left            =   1590
                  TabIndex        =   230
                  Top             =   150
                  Width           =   1575
                  _Version        =   786432
                  _ExtentX        =   2778
                  _ExtentY        =   661
                  _StockProps     =   79
                  Caption         =   "—œ Ã“¡ „‰ «·⁄—»Ê‰"
                  Enabled         =   0   'False
                  UseVisualStyle  =   -1  'True
                  TextAlignment   =   1
                  RightToLeft     =   -1  'True
               End
               Begin XtremeSuiteControls.CheckBox CheckStatusEarnest 
                  Height          =   375
                  Index           =   3
                  Left            =   3720
                  TabIndex        =   390
                  Top             =   390
                  Width           =   1215
                  _Version        =   786432
                  _ExtentX        =   2143
                  _ExtentY        =   661
                  _StockProps     =   79
                  Caption         =   "Œ’„ ⁄—»Ê‰"
                  UseVisualStyle  =   -1  'True
                  TextAlignment   =   1
                  RightToLeft     =   -1  'True
               End
            End
            Begin VB.TextBox TxtTelphone 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00FFFFFF&
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   178
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   315
               Left            =   3120
               MaxLength       =   50
               RightToLeft     =   -1  'True
               TabIndex        =   209
               Top             =   1440
               Width           =   2025
            End
            Begin VB.TextBox TxtCommission 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00FFFFFF&
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   178
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   315
               Left            =   3120
               MaxLength       =   50
               RightToLeft     =   -1  'True
               TabIndex        =   208
               Top             =   1800
               Width           =   2025
            End
            Begin VB.TextBox txtinstrunce 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00FFFFFF&
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   178
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   315
               Left            =   120
               MaxLength       =   50
               RightToLeft     =   -1  'True
               TabIndex        =   188
               Top             =   2160
               Width           =   2025
            End
            Begin VB.TextBox txtWater 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00FFFFFF&
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   178
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   315
               Left            =   120
               MaxLength       =   50
               RightToLeft     =   -1  'True
               TabIndex        =   187
               Top             =   2520
               Width           =   2025
            End
            Begin VB.TextBox TxtRent 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00FFFFFF&
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   178
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   315
               Left            =   120
               MaxLength       =   50
               RightToLeft     =   -1  'True
               TabIndex        =   184
               Top             =   1440
               Width           =   2025
            End
            Begin VB.TextBox txtrenterName 
               Alignment       =   1  'Right Justify
               Height          =   315
               Left            =   150
               MaxLength       =   255
               RightToLeft     =   -1  'True
               TabIndex        =   182
               Top             =   1080
               Width           =   4995
            End
            Begin VB.ComboBox cbointervaltype 
               Height          =   315
               ItemData        =   "FrmCashing1.frx":6786
               Left            =   120
               List            =   "FrmCashing1.frx":6793
               TabIndex        =   181
               Top             =   720
               Width           =   855
            End
            Begin VB.TextBox txtinterval 
               Height          =   285
               Left            =   1080
               TabIndex        =   180
               Top             =   720
               Width           =   615
            End
            Begin VB.TextBox TxtSearch 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00FFFFFF&
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   178
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   315
               Left            =   4080
               MaxLength       =   50
               RightToLeft     =   -1  'True
               TabIndex        =   172
               Top             =   360
               Width           =   1065
            End
            Begin MSDataListLib.DataCombo DcbIqara 
               Height          =   315
               Left            =   120
               TabIndex        =   173
               Tag             =   "⁄›Ê« Ì—ÃÏ «Œ Ì«—√”„ «·⁄ﬁ«—"
               Top             =   360
               Width           =   3915
               _ExtentX        =   6906
               _ExtentY        =   582
               _Version        =   393216
               Text            =   ""
               RightToLeft     =   -1  'True
            End
            Begin MSDataListLib.DataCombo DcbUnitNo 
               Height          =   315
               Left            =   2160
               TabIndex        =   174
               Tag             =   "⁄›Ê« Ì—ÃÏ «œŒ«· √”„ «·ÕÌ"
               Top             =   720
               Width           =   975
               _ExtentX        =   1720
               _ExtentY        =   582
               _Version        =   393216
               Text            =   ""
               RightToLeft     =   -1  'True
            End
            Begin MSDataListLib.DataCombo DcbUnitType 
               Height          =   315
               Left            =   4080
               TabIndex        =   175
               Tag             =   "⁄›Ê« Ì—ÃÏ «œŒ«· √”„ «·ÕÌ"
               Top             =   720
               Width           =   1065
               _ExtentX        =   1879
               _ExtentY        =   582
               _Version        =   393216
               Text            =   ""
               RightToLeft     =   -1  'True
            End
            Begin ImpulseButton.ISButton Cmd 
               Height          =   270
               Index           =   10
               Left            =   5520
               TabIndex        =   207
               Top             =   0
               Width           =   690
               _ExtentX        =   1217
               _ExtentY        =   476
               ButtonStyle     =   1
               ButtonPositionImage=   1
               Caption         =   "«€·«ﬁ"
               FontName        =   "Arial"
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   178
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ButtonImage     =   "FrmCashing1.frx":67A6
               DrawFocusRectangle=   0   'False
            End
            Begin VB.Label Label1 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               BackStyle       =   0  'Transparent
               Caption         =   "ﬁ.„"
               Height          =   195
               Index           =   24
               Left            =   5160
               RightToLeft     =   -1  'True
               TabIndex        =   387
               Top             =   2880
               Width           =   990
            End
            Begin VB.Label Label1 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               BackStyle       =   0  'Transparent
               Caption         =   "«·Œœ„« "
               Height          =   195
               Index           =   16
               Left            =   5040
               RightToLeft     =   -1  'True
               TabIndex        =   255
               Top             =   2160
               Width           =   990
            End
            Begin VB.Label Label1 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               BackStyle       =   0  'Transparent
               Caption         =   "ﬁÌ„… «·ﬂÂ—»«¡"
               Height          =   195
               Index           =   13
               Left            =   5160
               RightToLeft     =   -1  'True
               TabIndex        =   241
               Top             =   2520
               Width           =   990
            End
            Begin VB.Label Label1 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               BackStyle       =   0  'Transparent
               Caption         =   "”⁄Ì „ﬂ » Œ«—ÃÌ"
               Height          =   195
               Index           =   12
               Left            =   2040
               RightToLeft     =   -1  'True
               TabIndex        =   222
               Top             =   1800
               Width           =   990
            End
            Begin VB.Label Label1 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               BackStyle       =   0  'Transparent
               Caption         =   "—ﬁ„ «· ·›Ê‰"
               Height          =   195
               Index           =   11
               Left            =   5160
               RightToLeft     =   -1  'True
               TabIndex        =   213
               Top             =   1440
               Width           =   990
            End
            Begin VB.Label Label1 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               BackStyle       =   0  'Transparent
               Caption         =   "ﬁÌ„… «· «„Ì‰"
               Height          =   195
               Index           =   6
               Left            =   2040
               RightToLeft     =   -1  'True
               TabIndex        =   190
               Top             =   2160
               Width           =   990
            End
            Begin VB.Label Label1 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               BackStyle       =   0  'Transparent
               Caption         =   "ﬁÌ„… «·„Ì«Â"
               Height          =   195
               Index           =   5
               Left            =   2040
               RightToLeft     =   -1  'True
               TabIndex        =   189
               Top             =   2520
               Width           =   990
            End
            Begin VB.Label Label1 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               BackStyle       =   0  'Transparent
               Caption         =   "ﬁÌ„… «·«ÌÃ«—"
               Height          =   195
               Index           =   3
               Left            =   2040
               RightToLeft     =   -1  'True
               TabIndex        =   186
               Top             =   1440
               Width           =   990
            End
            Begin VB.Label Label1 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               BackStyle       =   0  'Transparent
               Caption         =   "ﬁÌ„… «·”⁄Ì"
               Height          =   195
               Index           =   2
               Left            =   5160
               RightToLeft     =   -1  'True
               TabIndex        =   185
               Top             =   1800
               Width           =   990
            End
            Begin VB.Label Label1 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               BackStyle       =   0  'Transparent
               Caption         =   "«·„” √Ã—"
               Height          =   195
               Index           =   1
               Left            =   5160
               RightToLeft     =   -1  'True
               TabIndex        =   183
               Top             =   1080
               Width           =   990
            End
            Begin VB.Label Label5 
               Caption         =   "«·„œ…"
               Height          =   255
               Left            =   1800
               TabIndex        =   179
               Top             =   720
               Width           =   495
            End
            Begin VB.Label Label1 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               BackStyle       =   0  'Transparent
               Caption         =   "«·⁄ﬁ«—"
               Height          =   195
               Index           =   4
               Left            =   5145
               RightToLeft     =   -1  'True
               TabIndex        =   178
               Top             =   360
               Width           =   990
            End
            Begin VB.Label Label1 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               BackStyle       =   0  'Transparent
               Caption         =   "‰Ê⁄ «·ÊÕœ…"
               Height          =   195
               Index           =   15
               Left            =   5280
               RightToLeft     =   -1  'True
               TabIndex        =   177
               Top             =   720
               Width           =   990
            End
            Begin VB.Label Label1 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               BackStyle       =   0  'Transparent
               Caption         =   "—ﬁ„ «·ÊÕœ…"
               ForeColor       =   &H00FF0000&
               Height          =   195
               Index           =   14
               Left            =   3000
               RightToLeft     =   -1  'True
               TabIndex        =   176
               Top             =   720
               Width           =   870
            End
         End
         Begin VB.Frame Frame4 
            Height          =   5775
            Left            =   12120
            TabIndex        =   147
            Top             =   240
            Width           =   6495
            Begin VB.Frame Frame5 
               Height          =   1935
               Left            =   3840
               RightToLeft     =   -1  'True
               TabIndex        =   191
               Top             =   2040
               Visible         =   0   'False
               Width           =   3135
               Begin VB.TextBox txttotal2 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00FFFFFF&
                  BeginProperty Font 
                     Name            =   "Arial"
                     Size            =   8.25
                     Charset         =   178
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   315
                  Left            =   240
                  MaxLength       =   50
                  RightToLeft     =   -1  'True
                  TabIndex        =   198
                  Top             =   240
                  Width           =   1425
               End
               Begin VB.TextBox txtinstranc 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00FFFFFF&
                  BeginProperty Font 
                     Name            =   "Arial"
                     Size            =   8.25
                     Charset         =   178
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   315
                  Left            =   120
                  MaxLength       =   50
                  RightToLeft     =   -1  'True
                  TabIndex        =   195
                  Top             =   1440
                  Width           =   2025
               End
               Begin VB.TextBox txttotal1 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00FFFFFF&
                  BeginProperty Font 
                     Name            =   "Arial"
                     Size            =   8.25
                     Charset         =   178
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   315
                  Left            =   240
                  MaxLength       =   50
                  RightToLeft     =   -1  'True
                  TabIndex        =   194
                  Top             =   600
                  Width           =   1425
               End
               Begin VB.TextBox txtComisin 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00FFFFFF&
                  BeginProperty Font 
                     Name            =   "Arial"
                     Size            =   8.25
                     Charset         =   178
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   315
                  Left            =   120
                  MaxLength       =   50
                  RightToLeft     =   -1  'True
                  TabIndex        =   192
                  Top             =   1080
                  Width           =   2025
               End
               Begin VB.Label Label1 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  BackStyle       =   0  'Transparent
                  Caption         =   "«Ã„«·Ì 2"
                  Height          =   195
                  Index           =   7
                  Left            =   1320
                  RightToLeft     =   -1  'True
                  TabIndex        =   199
                  Top             =   240
                  Width           =   990
               End
               Begin VB.Label Label1 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  BackStyle       =   0  'Transparent
                  Caption         =   "ﬁÌ„… «·«ÌÃ«—"
                  Height          =   195
                  Index           =   10
                  Left            =   2040
                  RightToLeft     =   -1  'True
                  TabIndex        =   197
                  Top             =   1440
                  Width           =   990
               End
               Begin VB.Label Label1 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  BackStyle       =   0  'Transparent
                  Caption         =   "«Ã„«·Ì1"
                  Height          =   195
                  Index           =   9
                  Left            =   1320
                  RightToLeft     =   -1  'True
                  TabIndex        =   196
                  Top             =   600
                  Width           =   990
               End
               Begin VB.Label Label1 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  BackStyle       =   0  'Transparent
                  Caption         =   "ﬁÌ„… «·”⁄Ì"
                  Height          =   195
                  Index           =   8
                  Left            =   2040
                  RightToLeft     =   -1  'True
                  TabIndex        =   193
                  Top             =   1080
                  Width           =   990
               End
            End
            Begin VB.TextBox TXTContNo 
               Height          =   495
               Left            =   600
               TabIndex        =   168
               Text            =   "0"
               Top             =   3360
               Visible         =   0   'False
               Width           =   375
            End
            Begin VB.Frame FraInfo 
               BackColor       =   &H00E2E9E9&
               Caption         =   "„⁄·Ê„«   Â„ﬂ"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   178
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000080&
               Height          =   2955
               Left            =   0
               RightToLeft     =   -1  'True
               TabIndex        =   148
               Top             =   120
               Width           =   6465
               Begin ImpulseAniLabel.ISAniLabel LblLinkInfo 
                  Height          =   225
                  Index           =   0
                  Left            =   2400
                  TabIndex        =   149
                  Top             =   480
                  Width           =   960
                  _ExtentX        =   1693
                  _ExtentY        =   397
                  ActiveUnderline =   -1  'True
                  BackStyle       =   1
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "Arial"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  FontBold        =   -1  'True
                  FontName        =   "Arial"
                  FontSize        =   8.25
                  ForeColor       =   4210688
                  MousePointer    =   99
                  MouseIcon       =   "FrmCashing1.frx":6D40
                  BackColor       =   14871017
                  Alignment       =   1
                  Caption         =   ""
                  ColorHover      =   16711680
                  RightToLeft     =   -1  'True
                  ImageCount      =   0
               End
               Begin ImpulseAniLabel.ISAniLabel LblLinkInfo 
                  Height          =   225
                  Index           =   1
                  Left            =   120
                  TabIndex        =   150
                  Top             =   780
                  Visible         =   0   'False
                  Width           =   960
                  _ExtentX        =   1693
                  _ExtentY        =   397
                  ActiveUnderline =   -1  'True
                  BackStyle       =   1
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "Arial"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  FontBold        =   -1  'True
                  FontName        =   "Arial"
                  FontSize        =   8.25
                  ForeColor       =   4210688
                  MousePointer    =   99
                  MouseIcon       =   "FrmCashing1.frx":6EA2
                  BackColor       =   14871017
                  Alignment       =   1
                  Caption         =   ""
                  ColorHover      =   16711680
                  RightToLeft     =   -1  'True
                  ImageCount      =   0
               End
               Begin ImpulseAniLabel.ISAniLabel LblLinkInfo 
                  Height          =   225
                  Index           =   2
                  Left            =   2400
                  TabIndex        =   151
                  Top             =   1110
                  Width           =   960
                  _ExtentX        =   1693
                  _ExtentY        =   397
                  ActiveUnderline =   -1  'True
                  BackStyle       =   1
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "Arial"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  FontBold        =   -1  'True
                  FontName        =   "Arial"
                  FontSize        =   8.25
                  ForeColor       =   4210688
                  MousePointer    =   99
                  MouseIcon       =   "FrmCashing1.frx":7004
                  BackColor       =   14871017
                  Alignment       =   1
                  Caption         =   ""
                  ColorHover      =   16711680
                  RightToLeft     =   -1  'True
                  ImageCount      =   0
               End
               Begin ImpulseAniLabel.ISAniLabel LblLinkInfo 
                  Height          =   225
                  Index           =   3
                  Left            =   120
                  TabIndex        =   152
                  Top             =   1350
                  Visible         =   0   'False
                  Width           =   960
                  _ExtentX        =   1693
                  _ExtentY        =   397
                  ActiveUnderline =   -1  'True
                  BackStyle       =   1
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "Arial"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  FontBold        =   -1  'True
                  FontName        =   "Arial"
                  FontSize        =   8.25
                  ForeColor       =   4210688
                  MousePointer    =   99
                  MouseIcon       =   "FrmCashing1.frx":7166
                  BackColor       =   14871017
                  Alignment       =   1
                  Caption         =   ""
                  ColorHover      =   16711680
                  RightToLeft     =   -1  'True
                  ImageCount      =   0
               End
               Begin ImpulseAniLabel.ISAniLabel LblLinkInfo 
                  Height          =   225
                  Index           =   4
                  Left            =   2400
                  TabIndex        =   153
                  Top             =   1800
                  Width           =   960
                  _ExtentX        =   1693
                  _ExtentY        =   397
                  ActiveUnderline =   -1  'True
                  BackStyle       =   1
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "Arial"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  FontBold        =   -1  'True
                  FontName        =   "Arial"
                  FontSize        =   8.25
                  ForeColor       =   4210688
                  MousePointer    =   99
                  MouseIcon       =   "FrmCashing1.frx":72C8
                  BackColor       =   14871017
                  Alignment       =   1
                  Caption         =   ""
                  ColorHover      =   16711680
                  RightToLeft     =   -1  'True
                  ImageCount      =   0
               End
               Begin ImpulseAniLabel.ISAniLabel LblLinkInfo 
                  Height          =   225
                  Index           =   5
                  Left            =   120
                  TabIndex        =   154
                  Top             =   1920
                  Visible         =   0   'False
                  Width           =   960
                  _ExtentX        =   1693
                  _ExtentY        =   397
                  ActiveUnderline =   -1  'True
                  BackStyle       =   1
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "Arial"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  FontBold        =   -1  'True
                  FontName        =   "Arial"
                  FontSize        =   8.25
                  ForeColor       =   4210688
                  MousePointer    =   99
                  MouseIcon       =   "FrmCashing1.frx":742A
                  BackColor       =   14871017
                  Alignment       =   1
                  Caption         =   ""
                  ColorHover      =   16711680
                  RightToLeft     =   -1  'True
                  ImageCount      =   0
               End
               Begin ImpulseAniLabel.ISAniLabel LblLinkInfo 
                  Height          =   225
                  Index           =   6
                  Left            =   -120
                  TabIndex        =   155
                  Top             =   600
                  Width           =   1230
                  _ExtentX        =   2170
                  _ExtentY        =   397
                  ActiveUnderline =   -1  'True
                  BackStyle       =   1
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "Arial"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  FontBold        =   -1  'True
                  FontName        =   "Arial"
                  FontSize        =   8.25
                  ForeColor       =   4210688
                  MousePointer    =   99
                  MouseIcon       =   "FrmCashing1.frx":758C
                  BackColor       =   14871017
                  Alignment       =   1
                  Caption         =   ""
                  ColorHover      =   16711680
                  RightToLeft     =   -1  'True
                  ImageCount      =   0
               End
               Begin ImpulseAniLabel.ISAniLabel LblLinkInfo 
                  Height          =   225
                  Index           =   7
                  Left            =   120
                  TabIndex        =   156
                  Top             =   1110
                  Width           =   1230
                  _ExtentX        =   2170
                  _ExtentY        =   397
                  ActiveUnderline =   -1  'True
                  BackStyle       =   1
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "Arial"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  FontBold        =   -1  'True
                  FontName        =   "Arial"
                  FontSize        =   8.25
                  ForeColor       =   4210688
                  MousePointer    =   99
                  MouseIcon       =   "FrmCashing1.frx":76EE
                  BackColor       =   14871017
                  Alignment       =   1
                  Caption         =   ""
                  ColorHover      =   16711680
                  RightToLeft     =   -1  'True
                  ImageCount      =   0
               End
               Begin ImpulseAniLabel.ISAniLabel LblLinkInfo 
                  Height          =   225
                  Index           =   8
                  Left            =   120
                  TabIndex        =   157
                  Top             =   1680
                  Width           =   1230
                  _ExtentX        =   2170
                  _ExtentY        =   397
                  ActiveUnderline =   -1  'True
                  BackStyle       =   1
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "Arial"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  FontBold        =   -1  'True
                  FontName        =   "Arial"
                  FontSize        =   8.25
                  ForeColor       =   4210688
                  MousePointer    =   99
                  MouseIcon       =   "FrmCashing1.frx":7850
                  BackColor       =   14871017
                  Alignment       =   1
                  Caption         =   ""
                  ColorHover      =   16711680
                  RightToLeft     =   -1  'True
                  ImageCount      =   0
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "‘Ìﬂ« "
                  BeginProperty Font 
                     Name            =   "Tahoma"
                     Size            =   8.25
                     Charset         =   178
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H000000C0&
                  Height          =   225
                  Index           =   28
                  Left            =   1110
                  RightToLeft     =   -1  'True
                  TabIndex        =   167
                  Top             =   780
                  Visible         =   0   'False
                  Width           =   675
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "‰ﬁœÌ"
                  BeginProperty Font 
                     Name            =   "Tahoma"
                     Size            =   8.25
                     Charset         =   178
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H000000C0&
                  Height          =   225
                  Index           =   27
                  Left            =   2820
                  RightToLeft     =   -1  'True
                  TabIndex        =   166
                  Top             =   780
                  Visible         =   0   'False
                  Width           =   705
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "‘Ìﬂ« "
                  BeginProperty Font 
                     Name            =   "Tahoma"
                     Size            =   8.25
                     Charset         =   178
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H000000C0&
                  Height          =   225
                  Index           =   26
                  Left            =   1110
                  RightToLeft     =   -1  'True
                  TabIndex        =   165
                  Top             =   1920
                  Visible         =   0   'False
                  Width           =   675
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "‰ﬁœÌ"
                  BeginProperty Font 
                     Name            =   "Tahoma"
                     Size            =   8.25
                     Charset         =   178
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H000000C0&
                  Height          =   225
                  Index           =   25
                  Left            =   2820
                  RightToLeft     =   -1  'True
                  TabIndex        =   164
                  Top             =   1920
                  Visible         =   0   'False
                  Width           =   705
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "‘Ìﬂ« "
                  BeginProperty Font 
                     Name            =   "Tahoma"
                     Size            =   8.25
                     Charset         =   178
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H000000C0&
                  Height          =   225
                  Index           =   24
                  Left            =   1110
                  RightToLeft     =   -1  'True
                  TabIndex        =   163
                  Top             =   1350
                  Visible         =   0   'False
                  Width           =   675
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "≈Ã„«·Ï „ﬁ»Ê÷«  «·ÌÊ„:"
                  BeginProperty Font 
                     Name            =   "Tahoma"
                     Size            =   9.75
                     Charset         =   178
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H000000C0&
                  Height          =   225
                  Index           =   23
                  Left            =   3540
                  RightToLeft     =   -1  'True
                  TabIndex        =   162
                  Top             =   420
                  Width           =   2235
               End
               Begin VB.Label lbl 
                  Alignment       =   2  'Center
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "«·≈”»Ê⁄ «·Õ«·Ï"
                  Height          =   255
                  Index           =   22
                  Left            =   3360
                  RightToLeft     =   -1  'True
                  TabIndex        =   161
                  Top             =   240
                  Visible         =   0   'False
                  Width           =   3495
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "‰ﬁœÌ"
                  BeginProperty Font 
                     Name            =   "Tahoma"
                     Size            =   8.25
                     Charset         =   178
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H000000C0&
                  Height          =   225
                  Index           =   21
                  Left            =   2820
                  RightToLeft     =   -1  'True
                  TabIndex        =   160
                  Top             =   1350
                  Visible         =   0   'False
                  Width           =   705
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "«·„ﬁ»Ê÷«  ›Ï «·‘Â— «·Õ«·Ï :"
                  BeginProperty Font 
                     Name            =   "Tahoma"
                     Size            =   9.75
                     Charset         =   178
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H000000C0&
                  Height          =   225
                  Index           =   20
                  Left            =   3540
                  RightToLeft     =   -1  'True
                  TabIndex        =   159
                  Top             =   1680
                  Width           =   2235
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "«·„ﬁ»Ê÷«  ›Ï «·≈”»Ê⁄ «·Õ«·Ï:"
                  BeginProperty Font 
                     Name            =   "Tahoma"
                     Size            =   9.75
                     Charset         =   178
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H000000C0&
                  Height          =   225
                  Index           =   19
                  Left            =   3540
                  RightToLeft     =   -1  'True
                  TabIndex        =   158
                  Top             =   1110
                  Width           =   2235
               End
            End
         End
         Begin VB.TextBox TxtContractNo 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFC0&
            Height          =   315
            Left            =   5760
            RightToLeft     =   -1  'True
            TabIndex        =   145
            Top             =   960
            Width           =   1515
         End
         Begin VB.TextBox TxtBookNo 
            Alignment       =   1  'Right Justify
            Height          =   315
            Left            =   3240
            RightToLeft     =   -1  'True
            TabIndex        =   135
            Top             =   960
            Width           =   1155
         End
         Begin VB.TextBox TxtManulaNO 
            Alignment       =   1  'Right Justify
            Height          =   315
            Left            =   3240
            RightToLeft     =   -1  'True
            TabIndex        =   129
            Top             =   240
            Width           =   1515
         End
         Begin Dynamic_Byte.NourHijriCal Txt_DateHigri 
            Height          =   315
            Left            =   5640
            TabIndex        =   127
            Top             =   240
            Width           =   1335
            _ExtentX        =   2355
            _ExtentY        =   556
         End
         Begin MSComCtl2.DTPicker XPDtbTrans 
            Height          =   315
            Left            =   6960
            TabIndex        =   0
            Top             =   240
            Width           =   1545
            _ExtentX        =   2725
            _ExtentY        =   556
            _Version        =   393216
            Format          =   166461441
            CurrentDate     =   41640
         End
         Begin VB.Frame Frame3 
            BackColor       =   &H00E2E9E9&
            Caption         =   "„⁄·Ê„«  «·ÕÊ«·Â"
            Enabled         =   0   'False
            Height          =   975
            Left            =   3960
            RightToLeft     =   -1  'True
            TabIndex        =   117
            Top             =   3240
            Visible         =   0   'False
            Width           =   3855
            Begin VB.TextBox Text4 
               Alignment       =   1  'Right Justify
               Height          =   315
               Left            =   120
               RightToLeft     =   -1  'True
               TabIndex        =   14
               Top             =   240
               Width           =   2565
            End
            Begin MSComCtl2.DTPicker DTPicker1 
               Height          =   315
               Left            =   120
               TabIndex        =   15
               Top             =   570
               Width           =   2565
               _ExtentX        =   4524
               _ExtentY        =   556
               _Version        =   393216
               Format          =   166461441
               CurrentDate     =   39614
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "—ﬁ„ «·ÕÊ«·Â"
               Height          =   285
               Index           =   45
               Left            =   2970
               RightToLeft     =   -1  'True
               TabIndex        =   119
               Top             =   240
               Width           =   735
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   " «—ÌŒÂ«"
               Height          =   285
               Index           =   44
               Left            =   2910
               RightToLeft     =   -1  'True
               TabIndex        =   118
               Top             =   570
               Width           =   735
            End
         End
         Begin VB.TextBox TxtCustCode 
            Alignment       =   1  'Right Justify
            Height          =   315
            Left            =   9360
            RightToLeft     =   -1  'True
            TabIndex        =   115
            Text            =   " "
            Top             =   1320
            Width           =   1275
         End
         Begin VB.TextBox Text2 
            Alignment       =   1  'Right Justify
            Height          =   315
            Left            =   9240
            RightToLeft     =   -1  'True
            TabIndex        =   114
            Text            =   " "
            Top             =   600
            Width           =   1395
         End
         Begin VB.Frame Fra 
            BackColor       =   &H00E2E9E9&
            Height          =   1005
            Index           =   0
            Left            =   20550
            RightToLeft     =   -1  'True
            TabIndex        =   66
            Top             =   270
            Width           =   3735
            Begin VB.TextBox TxtTransID 
               Alignment       =   1  'Right Justify
               Height          =   315
               Left            =   60
               RightToLeft     =   -1  'True
               TabIndex        =   69
               Top             =   120
               Visible         =   0   'False
               Width           =   375
            End
            Begin VB.TextBox TxtTransSerial 
               Alignment       =   1  'Right Justify
               Height          =   375
               Left            =   1110
               RightToLeft     =   -1  'True
               TabIndex        =   68
               Top             =   570
               Width           =   1005
            End
            Begin VB.ComboBox CboTrans 
               Height          =   315
               Left            =   120
               RightToLeft     =   -1  'True
               Style           =   2  'Dropdown List
               TabIndex        =   67
               Top             =   240
               Width           =   1995
            End
            Begin ImpulseButton.ISButton CmdSearchTrans 
               Height          =   345
               Left            =   600
               TabIndex        =   70
               Top             =   570
               Width           =   495
               _ExtentX        =   873
               _ExtentY        =   609
               ButtonPositionImage=   1
               Caption         =   "..."
               FontName        =   "Arial"
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   178
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ButtonImage     =   "FrmCashing1.frx":79B2
            End
            Begin ImpulseButton.ISButton CmdOpenTrans 
               Height          =   345
               Left            =   90
               TabIndex        =   71
               Top             =   570
               Visible         =   0   'False
               Width           =   495
               _ExtentX        =   873
               _ExtentY        =   609
               ButtonPositionImage=   1
               Caption         =   "..."
               FontName        =   "Arial"
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   178
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ButtonImage     =   "FrmCashing1.frx":7D4C
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "«œŒ· —ﬁ„ «·›« Ê—…"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   178
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000040&
               Height          =   315
               Index           =   10
               Left            =   2100
               RightToLeft     =   -1  'True
               TabIndex        =   73
               Top             =   630
               Width           =   1305
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "«Œ — ‰Ê⁄ «·›« Ê—…"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   178
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000040&
               Height          =   255
               Index           =   12
               Left            =   2100
               RightToLeft     =   -1  'True
               TabIndex        =   72
               Top             =   300
               Width           =   1305
            End
         End
         Begin VB.ComboBox DCboCashType 
            Height          =   315
            ItemData        =   "FrmCashing1.frx":80E6
            Left            =   8370
            List            =   "FrmCashing1.frx":80E8
            RightToLeft     =   -1  'True
            Style           =   2  'Dropdown List
            TabIndex        =   2
            Top             =   960
            Width           =   2265
         End
         Begin VB.TextBox XPMTxtRemarks 
            Alignment       =   1  'Right Justify
            Height          =   465
            Left            =   3930
            MultiLine       =   -1  'True
            RightToLeft     =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   17
            Top             =   4650
            Width           =   2715
         End
         Begin VB.TextBox XPTxtVal 
            Alignment       =   1  'Right Justify
            Height          =   315
            Left            =   7920
            MaxLength       =   10
            RightToLeft     =   -1  'True
            TabIndex        =   8
            Top             =   2415
            Width           =   2685
         End
         Begin VB.CheckBox ChkTrans 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "„‰ Õ”«» ›« Ê—…"
            Height          =   195
            Left            =   20040
            RightToLeft     =   -1  'True
            TabIndex        =   64
            Top             =   120
            Width           =   1575
         End
         Begin VB.Frame FraNote 
            BackColor       =   &H00E2E9E9&
            Height          =   1965
            Left            =   7860
            RightToLeft     =   -1  'True
            TabIndex        =   58
            Top             =   3120
            Width           =   4155
            Begin VB.TextBox TXTBankName 
               Alignment       =   1  'Right Justify
               Height          =   315
               Left            =   30
               RightToLeft     =   -1  'True
               TabIndex        =   10
               Top             =   480
               Visible         =   0   'False
               Width           =   2685
            End
            Begin VB.TextBox TxtChequeNumber 
               Alignment       =   1  'Right Justify
               Height          =   315
               Left            =   30
               RightToLeft     =   -1  'True
               TabIndex        =   11
               Top             =   840
               Width           =   2685
            End
            Begin MSComCtl2.DTPicker DtpChequeDueDate 
               Height          =   315
               Left            =   30
               TabIndex        =   12
               Top             =   1140
               Width           =   2685
               _ExtentX        =   4736
               _ExtentY        =   556
               _Version        =   393216
               Format          =   166461441
               CurrentDate     =   39614
            End
            Begin MSDataListLib.DataCombo DcboBankName 
               Height          =   315
               Left            =   30
               TabIndex        =   59
               Top             =   480
               Width           =   2685
               _ExtentX        =   4736
               _ExtentY        =   582
               _Version        =   393216
               Text            =   ""
               RightToLeft     =   -1  'True
            End
            Begin MSDataListLib.DataCombo DcboBox 
               Height          =   315
               Left            =   30
               TabIndex        =   9
               Top             =   150
               Width           =   2685
               _ExtentX        =   4736
               _ExtentY        =   582
               _Version        =   393216
               Text            =   ""
               RightToLeft     =   -1  'True
            End
            Begin MSDataListLib.DataCombo DcChequeBox 
               Height          =   315
               Left            =   30
               TabIndex        =   13
               Top             =   1560
               Width           =   2685
               _ExtentX        =   4736
               _ExtentY        =   582
               _Version        =   393216
               Text            =   ""
               RightToLeft     =   -1  'True
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "Õ«›Ÿ… «·‘Ìﬂ« "
               Height          =   285
               Index           =   43
               Left            =   2760
               RightToLeft     =   -1  'True
               TabIndex        =   116
               Top             =   1560
               Width           =   1215
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "«”„ «·Œ“‰…"
               Height          =   285
               Index           =   9
               Left            =   2790
               RightToLeft     =   -1  'True
               TabIndex        =   63
               Top             =   180
               Width           =   1215
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "«”„ «·»‰ﬂ"
               Height          =   285
               Index           =   15
               Left            =   2790
               RightToLeft     =   -1  'True
               TabIndex        =   62
               Top             =   510
               Width           =   1215
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "—ﬁ„ «·‘Ìﬂ"
               Height          =   285
               Index           =   16
               Left            =   2760
               RightToLeft     =   -1  'True
               TabIndex        =   61
               Top             =   810
               Width           =   1215
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   " «—ÌŒ «·≈” Õﬁ«ﬁ"
               Height          =   285
               Index           =   17
               Left            =   2820
               RightToLeft     =   -1  'True
               TabIndex        =   60
               Top             =   1140
               Width           =   1215
            End
         End
         Begin VB.Frame Fra 
            BackColor       =   &H00E2E9E9&
            Caption         =   "«·ﬁÌœ «·„Õ«”»Ì"
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
            Height          =   885
            Index           =   1
            Left            =   3900
            RightToLeft     =   -1  'True
            TabIndex        =   48
            Top             =   6000
            Width           =   8175
            Begin VB.TextBox TxtNoteSerial 
               Alignment       =   1  'Right Justify
               Height          =   315
               Left            =   5040
               RightToLeft     =   -1  'True
               TabIndex        =   49
               Top             =   200
               Width           =   1875
            End
            Begin MSDataListLib.DataCombo DcboDebitSide 
               Height          =   315
               Left            =   90
               TabIndex        =   50
               Top             =   180
               Width           =   3975
               _ExtentX        =   7011
               _ExtentY        =   582
               _Version        =   393216
               Enabled         =   0   'False
               Text            =   ""
               RightToLeft     =   -1  'True
            End
            Begin MSDataListLib.DataCombo DcboCreditSide 
               Height          =   315
               Left            =   90
               TabIndex        =   51
               Top             =   510
               Width           =   3975
               _ExtentX        =   7011
               _ExtentY        =   582
               _Version        =   393216
               Enabled         =   0   'False
               Text            =   ""
               RightToLeft     =   -1  'True
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "ÿ—› „œÌ‰"
               Height          =   285
               Index           =   32
               Left            =   3960
               RightToLeft     =   -1  'True
               TabIndex        =   57
               Top             =   180
               Width           =   885
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "ÿ—› œ«∆‰"
               Height          =   285
               Index           =   31
               Left            =   3960
               RightToLeft     =   -1  'True
               TabIndex        =   56
               Top             =   510
               Width           =   885
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "—ﬁ„ «·ﬁÌœ:"
               Height          =   315
               Index           =   30
               Left            =   6930
               RightToLeft     =   -1  'True
               TabIndex        =   55
               Top             =   210
               Width           =   975
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "—ﬁ„ «·› —… :"
               Height          =   315
               Index           =   29
               Left            =   6930
               RightToLeft     =   -1  'True
               TabIndex        =   54
               Top             =   540
               Width           =   975
            End
            Begin VB.Label LblDevID 
               Alignment       =   2  'Center
               BackColor       =   &H00E2E9E9&
               Height          =   285
               Left            =   3870
               RightToLeft     =   -1  'True
               TabIndex        =   53
               Top             =   210
               Width           =   1485
            End
            Begin VB.Label lbl 
               Alignment       =   2  'Center
               BackColor       =   &H00E2E9E9&
               Height          =   285
               Index           =   33
               Left            =   5190
               RightToLeft     =   -1  'True
               TabIndex        =   52
               Top             =   510
               Width           =   1485
            End
         End
         Begin VB.TextBox XPTxtBillID 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   21840
            RightToLeft     =   -1  'True
            TabIndex        =   46
            Top             =   930
            Visible         =   0   'False
            Width           =   975
         End
         Begin VB.Frame Frame1 
            BackColor       =   &H00C0FFFF&
            Caption         =   "ŒÌ«—« "
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   1215
            Left            =   0
            RightToLeft     =   -1  'True
            TabIndex        =   39
            Top             =   90
            Width           =   3135
            Begin VB.OptionButton Option7 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00C0FFFF&
               Caption         =   "„‘«—Ì⁄ ”«»ﬁ…"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   178
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   195
               Left            =   840
               RightToLeft     =   -1  'True
               TabIndex        =   133
               Top             =   960
               Width           =   2055
            End
            Begin VB.OptionButton Option3 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00C0FFFF&
               Caption         =   "œ›⁄Â „ﬁœ„Â"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   178
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   195
               Left            =   1200
               RightToLeft     =   -1  'True
               TabIndex        =   43
               Top             =   240
               Width           =   1695
            End
            Begin VB.OptionButton Option1 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00C0FFFF&
               Caption         =   "FIFO"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   178
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   195
               Left            =   1560
               RightToLeft     =   -1  'True
               TabIndex        =   42
               Top             =   480
               Width           =   1335
            End
            Begin VB.OptionButton Option2 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00C0FFFF&
               Caption         =   " ÕœÌœ ›Ê« Ì—"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   178
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   195
               Left            =   840
               RightToLeft     =   -1  'True
               TabIndex        =   41
               Top             =   720
               Width           =   2055
            End
            Begin VB.OptionButton Option6 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00FFFFFF&
               Caption         =   " ÕœÌœ „” Œ·’« "
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   178
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   195
               Left            =   1320
               RightToLeft     =   -1  'True
               TabIndex        =   40
               Top             =   1560
               Value           =   -1  'True
               Width           =   2055
            End
            Begin ALLButtonS.ALLButton ALLButton3 
               Height          =   255
               Left            =   120
               TabIndex        =   44
               Top             =   720
               Width           =   615
               _ExtentX        =   1085
               _ExtentY        =   450
               BTYPE           =   3
               TX              =   " ÕœÌœ"
               ENAB            =   0   'False
               BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   178
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               COLTYPE         =   1
               FOCUSR          =   -1  'True
               BCOL            =   15790320
               BCOLO           =   15790320
               FCOL            =   0
               FCOLO           =   0
               MCOL            =   12632256
               MPTR            =   1
               MICON           =   "FrmCashing1.frx":80EA
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
               Left            =   120
               TabIndex        =   45
               Top             =   1320
               Width           =   855
               _ExtentX        =   1508
               _ExtentY        =   450
               BTYPE           =   3
               TX              =   " ÕœÌœ"
               ENAB            =   0   'False
               BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   178
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               COLTYPE         =   1
               FOCUSR          =   -1  'True
               BCOL            =   15790320
               BCOLO           =   15790320
               FCOL            =   0
               FCOLO           =   0
               MCOL            =   12632256
               MPTR            =   1
               MICON           =   "FrmCashing1.frx":8106
               UMCOL           =   -1  'True
               SOFT            =   0   'False
               PICPOS          =   0
               NGREY           =   0   'False
               FX              =   0
               HAND            =   0   'False
               CHECK           =   0   'False
               VALUE           =   -1  'True
            End
         End
         Begin VB.TextBox txtAdv_payment_value 
            Alignment       =   1  'Right Justify
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Left            =   3990
            MaxLength       =   10
            RightToLeft     =   -1  'True
            TabIndex        =   38
            Top             =   1950
            Visible         =   0   'False
            Width           =   1935
         End
         Begin VB.TextBox Text1 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   21960
            RightToLeft     =   -1  'True
            TabIndex        =   37
            Text            =   "Text1"
            Top             =   690
            Visible         =   0   'False
            Width           =   735
         End
         Begin VB.TextBox TxtNoteSerial1 
            Alignment       =   1  'Right Justify
            Height          =   315
            Left            =   9240
            Locked          =   -1  'True
            RightToLeft     =   -1  'True
            TabIndex        =   36
            Top             =   210
            Width           =   1395
         End
         Begin VB.Frame Frame2 
            BackColor       =   &H00E2E9E9&
            Caption         =   "›Ì Õ«·… «·„‘«—Ì⁄"
            Enabled         =   0   'False
            Height          =   615
            Left            =   7320
            RightToLeft     =   -1  'True
            TabIndex        =   35
            Top             =   1650
            Visible         =   0   'False
            Width           =   3375
            Begin VB.OptionButton Option4 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "⁄„Ì· ‰Â«∆Ì"
               Height          =   195
               Left            =   720
               RightToLeft     =   -1  'True
               TabIndex        =   6
               Top             =   120
               Value           =   -1  'True
               Width           =   1215
            End
            Begin VB.OptionButton Option5 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "„ﬁ«Ê· »«ÿ‰"
               Height          =   195
               Left            =   480
               RightToLeft     =   -1  'True
               TabIndex        =   7
               Top             =   360
               Width           =   1455
            End
         End
         Begin VB.TextBox txtperson 
            Alignment       =   1  'Right Justify
            Height          =   315
            Left            =   3960
            RightToLeft     =   -1  'True
            TabIndex        =   16
            Top             =   4290
            Width           =   2685
         End
         Begin vbalIml6.vbalImageList vbalImageList1 
            Left            =   21600
            Top             =   450
            _ExtentX        =   953
            _ExtentY        =   953
         End
         Begin ALLButtonS.ALLButton ALLButton1 
            Height          =   375
            Left            =   21360
            TabIndex        =   47
            Top             =   2610
            Visible         =   0   'False
            Width           =   1695
            _ExtentX        =   2990
            _ExtentY        =   661
            BTYPE           =   3
            TX              =   "«ŸÂ«— «·«ﬁ”«ÿ"
            ENAB            =   -1  'True
            BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   178
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            COLTYPE         =   1
            FOCUSR          =   -1  'True
            BCOL            =   15790320
            BCOLO           =   15790320
            FCOL            =   0
            FCOLO           =   0
            MCOL            =   12632256
            MPTR            =   1
            MICON           =   "FrmCashing1.frx":8122
            UMCOL           =   -1  'True
            SOFT            =   0   'False
            PICPOS          =   0
            NGREY           =   0   'False
            FX              =   0
            HAND            =   0   'False
            CHECK           =   0   'False
            VALUE           =   0   'False
         End
         Begin MSDataListLib.DataCombo DBCboClientName 
            Height          =   315
            Left            =   5640
            TabIndex        =   3
            Top             =   1320
            Width           =   3645
            _ExtentX        =   6429
            _ExtentY        =   582
            _Version        =   393216
            Text            =   ""
            RightToLeft     =   -1  'True
         End
         Begin C1SizerLibCtl.C1Elastic Ele 
            Height          =   540
            Index           =   2
            Left            =   120
            TabIndex        =   65
            TabStop         =   0   'False
            Top             =   7050
            Width           =   7995
            _cx             =   14102
            _cy             =   953
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
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
         End
         Begin ImpulseAniLabel.ISAniLabel LblLink 
            Height          =   315
            Left            =   210
            TabIndex        =   74
            Top             =   1320
            Width           =   4320
            _ExtentX        =   7620
            _ExtentY        =   556
            ActiveUnderline =   -1  'True
            BackStyle       =   1
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            FontBold        =   -1  'True
            FontName        =   "Arial"
            FontSize        =   8.25
            ForeColor       =   4210688
            MousePointer    =   99
            MouseIcon       =   "FrmCashing1.frx":813E
            BackColor       =   14871017
            Alignment       =   1
            Caption         =   ""
            ColorHover      =   16711680
            RightToLeft     =   -1  'True
            ImageCount      =   0
         End
         Begin ALLButtonS.ALLButton ALLButton2 
            Height          =   375
            Left            =   21000
            TabIndex        =   75
            Top             =   2850
            Visible         =   0   'False
            Width           =   1695
            _ExtentX        =   2990
            _ExtentY        =   661
            BTYPE           =   3
            TX              =   "«ŸÂ«— ”‰œ «·„œÌÊ‰Ì…"
            ENAB            =   -1  'True
            BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   178
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            COLTYPE         =   1
            FOCUSR          =   -1  'True
            BCOL            =   15790320
            BCOLO           =   15790320
            FCOL            =   0
            FCOLO           =   0
            MCOL            =   12632256
            MPTR            =   1
            MICON           =   "FrmCashing1.frx":82A0
            UMCOL           =   -1  'True
            SOFT            =   0   'False
            PICPOS          =   0
            NGREY           =   0   'False
            FX              =   0
            HAND            =   0   'False
            CHECK           =   0   'False
            VALUE           =   0   'False
         End
         Begin MSDataListLib.DataCombo DCPROJECT 
            Height          =   315
            Left            =   19560
            TabIndex        =   76
            Top             =   4170
            Visible         =   0   'False
            Width           =   2655
            _ExtentX        =   4683
            _ExtentY        =   582
            _Version        =   393216
            Text            =   ""
            RightToLeft     =   -1  'True
         End
         Begin MSDataListLib.DataCombo Dcbranch 
            Bindings        =   "FrmCashing1.frx":82BC
            Height          =   315
            Left            =   3240
            TabIndex        =   1
            Top             =   600
            Width           =   6015
            _ExtentX        =   10610
            _ExtentY        =   582
            _Version        =   393216
            BackColor       =   16777215
            ListField       =   "account_name"
            BoundColumn     =   "code"
            Text            =   ""
            RightToLeft     =   -1  'True
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   178
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
         Begin MSDataListLib.DataCombo dcEmployee 
            Height          =   315
            Left            =   5640
            TabIndex        =   5
            Top             =   1410
            Visible         =   0   'False
            Width           =   3645
            _ExtentX        =   6429
            _ExtentY        =   582
            _Version        =   393216
            Text            =   ""
            RightToLeft     =   -1  'True
         End
         Begin MSDataListLib.DataCombo DCAccounts 
            Height          =   315
            Left            =   5640
            TabIndex        =   122
            Top             =   1320
            Visible         =   0   'False
            Width           =   3645
            _ExtentX        =   6429
            _ExtentY        =   582
            _Version        =   393216
            Text            =   ""
            RightToLeft     =   -1  'True
         End
         Begin MSDataListLib.DataCombo DcEmp 
            Bindings        =   "FrmCashing1.frx":82D1
            Height          =   315
            Left            =   0
            TabIndex        =   18
            Top             =   2400
            Width           =   2595
            _ExtentX        =   4577
            _ExtentY        =   582
            _Version        =   393216
            BackColor       =   16777215
            ListField       =   "account_name"
            BoundColumn     =   "code"
            Text            =   ""
            RightToLeft     =   -1  'True
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   178
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
         Begin MSDataListLib.DataCombo DCCar 
            Height          =   315
            Left            =   0
            TabIndex        =   19
            Top             =   2760
            Width           =   2595
            _ExtentX        =   4577
            _ExtentY        =   582
            _Version        =   393216
            Text            =   ""
            RightToLeft     =   -1  'True
         End
         Begin MSDataListLib.DataCombo DCDriver 
            Height          =   315
            Left            =   0
            TabIndex        =   20
            Top             =   3120
            Width           =   2595
            _ExtentX        =   4577
            _ExtentY        =   582
            _Version        =   393216
            Text            =   ""
            RightToLeft     =   -1  'True
         End
         Begin ImpulseButton.ISButton SearchCashCustomer 
            Height          =   375
            Left            =   0
            TabIndex        =   169
            TabStop         =   0   'False
            ToolTipText     =   "«÷€ÿ ·«÷«›… ⁄„Ì· ÃœÌœ"
            Top             =   480
            Width           =   510
            _ExtentX        =   900
            _ExtentY        =   661
            ButtonStyle     =   1
            ButtonPositionImage=   4
            Caption         =   ""
            BackColor       =   14871017
            FontName        =   "Arial"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   178
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BackStyle       =   0
            ButtonImage     =   "FrmCashing1.frx":82E6
            ColorButton     =   14871017
            ColorHighlight  =   16777215
            ColorShadow     =   -2147483631
            ColorOutline    =   -2147483631
            DrawFocusRectangle=   0   'False
         End
         Begin ImpulseButton.ISButton ISButton1 
            Height          =   375
            Left            =   5280
            TabIndex        =   170
            TabStop         =   0   'False
            ToolTipText     =   "«÷€ÿ ·«÷«›… ⁄„Ì· ÃœÌœ"
            Top             =   960
            Width           =   510
            _ExtentX        =   900
            _ExtentY        =   661
            ButtonStyle     =   1
            ButtonPositionImage=   4
            Caption         =   ""
            BackColor       =   14871017
            FontName        =   "Arial"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   178
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BackStyle       =   0
            ButtonImage     =   "FrmCashing1.frx":86E3
            ColorButton     =   14871017
            ColorHighlight  =   16777215
            ColorShadow     =   -2147483631
            ColorOutline    =   -2147483631
            DrawFocusRectangle=   0   'False
         End
         Begin MSComCtl2.DTPicker Dtaefilter 
            Height          =   315
            Left            =   0
            TabIndex        =   223
            Top             =   2520
            Visible         =   0   'False
            Width           =   1545
            _ExtentX        =   2725
            _ExtentY        =   556
            _Version        =   393216
            Format          =   163643393
            CurrentDate     =   41640
         End
         Begin MSDataListLib.DataCombo DcCostCenter 
            Bindings        =   "FrmCashing1.frx":8AE0
            Height          =   315
            Left            =   3960
            TabIndex        =   247
            Top             =   2850
            Width           =   2655
            _ExtentX        =   4683
            _ExtentY        =   582
            _Version        =   393216
            BackColor       =   16777215
            ListField       =   "account_name"
            BoundColumn     =   "code"
            Text            =   ""
            RightToLeft     =   -1  'True
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   178
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
         Begin MSDataListLib.DataCombo DcbAccount 
            Height          =   315
            Left            =   3930
            TabIndex        =   362
            Top             =   5160
            Width           =   5055
            _ExtentX        =   8916
            _ExtentY        =   582
            _Version        =   393216
            Text            =   ""
            RightToLeft     =   -1  'True
         End
         Begin VB.TextBox TxtVATValue 
            Alignment       =   1  'Right Justify
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Left            =   3960
            MaxLength       =   10
            RightToLeft     =   -1  'True
            TabIndex        =   382
            Top             =   2400
            Width           =   1935
         End
         Begin MSDataListLib.DataCombo DcboRevenuesTypes 
            Height          =   315
            Left            =   5640
            TabIndex        =   4
            Top             =   1320
            Width           =   3645
            _ExtentX        =   6429
            _ExtentY        =   582
            _Version        =   393216
            Text            =   ""
            RightToLeft     =   -1  'True
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "  «„Ì‰"
            Height          =   285
            Index           =   109
            Left            =   7200
            RightToLeft     =   -1  'True
            TabIndex        =   389
            Top             =   1680
            Width           =   555
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "ﬁÌ„… „÷«›…"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000C0&
            Height          =   285
            Index           =   108
            Left            =   5970
            RightToLeft     =   -1  'True
            TabIndex        =   383
            Top             =   2415
            Visible         =   0   'False
            Width           =   1035
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "«·Õ”«»"
            Height          =   285
            Index           =   104
            Left            =   10680
            RightToLeft     =   -1  'True
            TabIndex        =   364
            Top             =   5160
            Width           =   885
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "„·«ÕŸ« "
            Height          =   285
            Index           =   62
            Left            =   10800
            RightToLeft     =   -1  'True
            TabIndex        =   232
            Top             =   5640
            Width           =   885
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   " «Ã„«·Ì "
            Height          =   405
            Index           =   61
            Left            =   4440
            RightToLeft     =   -1  'True
            TabIndex        =   220
            Top             =   960
            Width           =   555
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "—ﬁ„ «· ’›ÌÂ"
            Height          =   285
            Index           =   60
            Left            =   7440
            RightToLeft     =   -1  'True
            TabIndex        =   218
            Top             =   960
            Width           =   795
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "—ﬁ„ «·⁄ﬁœ"
            Height          =   285
            Index           =   53
            Left            =   7440
            RightToLeft     =   -1  'True
            TabIndex        =   146
            Top             =   960
            Width           =   795
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   " «·œ› —"
            Height          =   285
            Index           =   51
            Left            =   4920
            RightToLeft     =   -1  'True
            TabIndex        =   136
            Top             =   960
            Width           =   555
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            BackStyle       =   0  'Transparent
            Caption         =   "Õœœ «·„⁄œÂ/«·”Ì«—…"
            Height          =   285
            Index           =   50
            Left            =   2640
            RightToLeft     =   -1  'True
            TabIndex        =   131
            Top             =   2760
            Width           =   1215
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            BackStyle       =   0  'Transparent
            Caption         =   "Õœœ «·”«∆ﬁ"
            Height          =   285
            Index           =   49
            Left            =   2640
            RightToLeft     =   -1  'True
            TabIndex        =   130
            Top             =   3120
            Width           =   1215
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "«·—ﬁ„ «·ÌœÊÌ"
            Height          =   285
            Index           =   48
            Left            =   4680
            RightToLeft     =   -1  'True
            TabIndex        =   128
            Top             =   240
            Width           =   915
         End
         Begin VB.Label Label3 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "«·„‰œÊ»"
            Height          =   255
            Left            =   2640
            RightToLeft     =   -1  'True
            TabIndex        =   126
            Top             =   2400
            Width           =   1215
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "ﬁÌ„… «·„ﬁ»Ê÷« "
            Height          =   315
            Index           =   47
            Left            =   4200
            RightToLeft     =   -1  'True
            TabIndex        =   125
            Top             =   1680
            Width           =   1155
         End
         Begin VB.Label Label2 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "«·›—⁄"
            Height          =   255
            Left            =   10650
            RightToLeft     =   -1  'True
            TabIndex        =   113
            Top             =   600
            Width           =   1395
         End
         Begin VB.Label XPTxtCurrent 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Height          =   315
            Left            =   1890
            RightToLeft     =   -1  'True
            TabIndex        =   94
            Top             =   6570
            Width           =   825
         End
         Begin VB.Label XPTxtCount 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Height          =   315
            Left            =   180
            RightToLeft     =   -1  'True
            TabIndex        =   93
            Top             =   6570
            Width           =   615
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   " ⁄œœ «·”Ã·« :"
            Height          =   315
            Index           =   37
            Left            =   840
            RightToLeft     =   -1  'True
            TabIndex        =   92
            Top             =   6570
            Width           =   1005
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   " «·”Ã· «·Õ«·Ì:"
            Height          =   315
            Index           =   7
            Left            =   2760
            RightToLeft     =   -1  'True
            TabIndex        =   91
            Top             =   6570
            Width           =   1065
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "‰Ê⁄ «·„ﬁ»Ê÷« "
            Height          =   285
            Index           =   6
            Left            =   10530
            RightToLeft     =   -1  'True
            TabIndex        =   90
            Top             =   960
            Width           =   1395
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "«· «—ÌŒ"
            Height          =   285
            Index           =   1
            Left            =   8610
            RightToLeft     =   -1  'True
            TabIndex        =   89
            Top             =   285
            Width           =   555
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "ﬁÌ„… «·„ﬁ»Ê÷« "
            Height          =   285
            Index           =   2
            Left            =   10770
            RightToLeft     =   -1  'True
            TabIndex        =   88
            Top             =   2430
            Width           =   1275
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "«”„ «·⁄„Ì· √Ê «·„Ê—œ"
            Height          =   315
            Index           =   3
            Left            =   10650
            RightToLeft     =   -1  'True
            TabIndex        =   87
            Top             =   1290
            Width           =   1395
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "ﬂÊœ «·”‰œ"
            Height          =   285
            Index           =   4
            Left            =   10650
            RightToLeft     =   -1  'True
            TabIndex        =   86
            Top             =   300
            Width           =   1395
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "–·ﬂ „ﬁ«»·"
            Height          =   285
            Index           =   5
            Left            =   6690
            RightToLeft     =   -1  'True
            TabIndex        =   85
            Top             =   4770
            Width           =   885
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "«·—’Ìœ «·Õ«·Ï:"
            Height          =   315
            Index           =   13
            Left            =   4200
            RightToLeft     =   -1  'True
            TabIndex        =   84
            Top             =   1320
            Width           =   1155
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "ÿ—Ìﬁ… «·ﬁ»÷"
            Height          =   315
            Index           =   14
            Left            =   10770
            RightToLeft     =   -1  'True
            TabIndex        =   83
            Top             =   2850
            Width           =   1275
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000080&
            Height          =   435
            Index           =   18
            Left            =   210
            RightToLeft     =   -1  'True
            TabIndex        =   82
            Top             =   1680
            Width           =   4065
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "«·„‘—Ê⁄"
            Height          =   285
            Index           =   34
            Left            =   18480
            RightToLeft     =   -1  'True
            TabIndex        =   81
            Top             =   4410
            Visible         =   0   'False
            Width           =   1455
         End
         Begin VB.Label lblsqlstring 
            Alignment       =   1  'Right Justify
            Height          =   855
            Left            =   20400
            RightToLeft     =   -1  'True
            TabIndex        =   80
            Top             =   2250
            Width           =   2895
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "œ›⁄Â „ﬁœ„Â"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000C0&
            Height          =   285
            Index           =   35
            Left            =   5970
            RightToLeft     =   -1  'True
            TabIndex        =   79
            Top             =   1950
            Visible         =   0   'False
            Width           =   1035
         End
         Begin VB.Label Label8 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "„—ﬂ“ «· ﬂ·›… «·⁄«„"
            Height          =   255
            Left            =   6600
            RightToLeft     =   -1  'True
            TabIndex        =   78
            Top             =   2850
            Width           =   1215
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "„‰ «·„ﬂ—„"
            Height          =   285
            Index           =   36
            Left            =   6720
            RightToLeft     =   -1  'True
            TabIndex        =   77
            Top             =   4290
            Width           =   975
         End
      End
      Begin C1SizerLibCtl.C1Elastic Ele 
         Height          =   7080
         Index           =   0
         Left            =   19455
         TabIndex        =   23
         TabStop         =   0   'False
         Top             =   45
         Width           =   18720
         _cx             =   33020
         _cy             =   12488
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
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
         Begin VSFlex8Ctl.VSFlexGrid Grid1 
            Height          =   2115
            Left            =   3720
            TabIndex        =   100
            Top             =   4080
            Width           =   14835
            _cx             =   26167
            _cy             =   3731
            Appearance      =   1
            BorderStyle     =   1
            Enabled         =   -1  'True
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
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
            Rows            =   2
            Cols            =   15
            FixedRows       =   1
            FixedCols       =   1
            RowHeightMin    =   320
            RowHeightMax    =   0
            ColWidthMin     =   0
            ColWidthMax     =   0
            ExtendLastCol   =   0   'False
            FormatString    =   $"FrmCashing1.frx":8AF5
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
         Begin ALLButtonS.ALLButton CmdRemove 
            Height          =   375
            Left            =   0
            TabIndex        =   111
            Tag             =   "Delete Row"
            Top             =   6240
            Width           =   1215
            _ExtentX        =   2143
            _ExtentY        =   661
            BTYPE           =   3
            TX              =   "Õ–› „” Œ·’"
            ENAB            =   -1  'True
            BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   178
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            COLTYPE         =   2
            FOCUSR          =   -1  'True
            BCOL            =   0
            BCOLO           =   0
            FCOL            =   255
            FCOLO           =   255
            MCOL            =   12632256
            MPTR            =   1
            MICON           =   "FrmCashing1.frx":8D59
            UMCOL           =   -1  'True
            SOFT            =   0   'False
            PICPOS          =   0
            NGREY           =   0   'False
            FX              =   0
            HAND            =   0   'False
            CHECK           =   0   'False
            VALUE           =   0   'False
         End
         Begin VSFlex8Ctl.VSFlexGrid Grid 
            Height          =   2115
            Left            =   3720
            TabIndex        =   112
            Top             =   960
            Width           =   14715
            _cx             =   25956
            _cy             =   3731
            Appearance      =   1
            BorderStyle     =   1
            Enabled         =   -1  'True
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
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
            Rows            =   2
            Cols            =   14
            FixedRows       =   1
            FixedCols       =   1
            RowHeightMin    =   320
            RowHeightMax    =   0
            ColWidthMin     =   0
            ColWidthMax     =   0
            ExtendLastCol   =   0   'False
            FormatString    =   $"FrmCashing1.frx":8D75
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
         Begin VB.Shape Shape3 
            BorderWidth     =   2
            Height          =   495
            Left            =   3840
            Top             =   360
            Width           =   14535
         End
         Begin VB.Label lbl 
            Alignment       =   2  'Center
            BackColor       =   &H00C0FFFF&
            Caption         =   "«·„„” Œ·’«  «· Ì  „ ”œ«œÂ« ··„‘—Ê⁄"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   12
               Charset         =   178
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000080&
            Height          =   420
            Index           =   42
            Left            =   7440
            RightToLeft     =   -1  'True
            TabIndex        =   99
            Top             =   3240
            Width           =   10935
         End
         Begin VB.Label lbl 
            Alignment       =   2  'Center
            BackColor       =   &H00C0FFFF&
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   14.25
               Charset         =   178
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000080&
            Height          =   420
            Index           =   41
            Left            =   3840
            RightToLeft     =   -1  'True
            TabIndex        =   98
            Top             =   3240
            Width           =   3735
         End
         Begin VB.Label lbl 
            Alignment       =   2  'Center
            BackColor       =   &H00C0FFFF&
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   14.25
               Charset         =   178
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000080&
            Height          =   420
            Index           =   38
            Left            =   3960
            RightToLeft     =   -1  'True
            TabIndex        =   95
            Top             =   360
            Width           =   3735
         End
         Begin VB.Label lbl 
            Alignment       =   2  'Center
            BackColor       =   &H00C0FFFF&
            Caption         =   "ﬁ„ » ÕœÌœ «·„” Œ·’«   «·„—«œ ”œ«œÂ« ··„‘—Ê⁄"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   12
               Charset         =   178
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000080&
            Height          =   420
            Index           =   0
            Left            =   7680
            RightToLeft     =   -1  'True
            TabIndex        =   25
            Top             =   360
            Width           =   10695
         End
         Begin VB.Shape Shape2 
            BorderWidth     =   2
            Height          =   495
            Left            =   3720
            Top             =   3240
            Width           =   14775
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Height          =   375
            Index           =   0
            Left            =   4560
            RightToLeft     =   -1  'True
            TabIndex        =   24
            Top             =   840
            Width           =   7575
         End
      End
      Begin C1SizerLibCtl.C1Elastic Ele 
         Height          =   7080
         Index           =   3
         Left            =   19755
         TabIndex        =   137
         TabStop         =   0   'False
         Top             =   45
         Width           =   18720
         _cx             =   33020
         _cy             =   12488
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
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
         Begin ALLButtonS.ALLButton ALLButton5 
            Height          =   375
            Left            =   0
            TabIndex        =   138
            Tag             =   "Delete Row"
            Top             =   6240
            Width           =   1215
            _ExtentX        =   2143
            _ExtentY        =   661
            BTYPE           =   3
            TX              =   "Õ–› „” Œ·’"
            ENAB            =   -1  'True
            BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   178
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            COLTYPE         =   2
            FOCUSR          =   -1  'True
            BCOL            =   0
            BCOLO           =   0
            FCOL            =   255
            FCOLO           =   255
            MCOL            =   12632256
            MPTR            =   1
            MICON           =   "FrmCashing1.frx":8FBC
            UMCOL           =   -1  'True
            SOFT            =   0   'False
            PICPOS          =   0
            NGREY           =   0   'False
            FX              =   0
            HAND            =   0   'False
            CHECK           =   0   'False
            VALUE           =   0   'False
         End
         Begin VSFlex8Ctl.VSFlexGrid Grid3 
            Height          =   1515
            Left            =   0
            TabIndex        =   139
            Top             =   960
            Width           =   18555
            _cx             =   32729
            _cy             =   2672
            Appearance      =   1
            BorderStyle     =   1
            Enabled         =   -1  'True
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
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
            Rows            =   2
            Cols            =   55
            FixedRows       =   1
            FixedCols       =   1
            RowHeightMin    =   320
            RowHeightMax    =   0
            ColWidthMin     =   0
            ColWidthMax     =   0
            ExtendLastCol   =   -1  'True
            FormatString    =   $"FrmCashing1.frx":8FD8
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
         Begin VSFlex8Ctl.VSFlexGrid Grid4 
            Height          =   1635
            Left            =   0
            TabIndex        =   144
            Top             =   3480
            Width           =   18555
            _cx             =   32729
            _cy             =   2884
            Appearance      =   1
            BorderStyle     =   1
            Enabled         =   -1  'True
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
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
            Rows            =   2
            Cols            =   26
            FixedRows       =   1
            FixedCols       =   1
            RowHeightMin    =   320
            RowHeightMax    =   0
            ColWidthMin     =   0
            ColWidthMax     =   0
            ExtendLastCol   =   -1  'True
            FormatString    =   $"FrmCashing1.frx":989A
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
         Begin VSFlex8Ctl.VSFlexGrid VSFlexGrid3 
            Height          =   1275
            Left            =   8730
            TabIndex        =   242
            Top             =   5760
            Width           =   9825
            _cx             =   17330
            _cy             =   2249
            Appearance      =   1
            BorderStyle     =   1
            Enabled         =   -1  'True
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
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
            Cols            =   10
            FixedRows       =   1
            FixedCols       =   1
            RowHeightMin    =   320
            RowHeightMax    =   0
            ColWidthMin     =   0
            ColWidthMax     =   0
            ExtendLastCol   =   -1  'True
            FormatString    =   $"FrmCashing1.frx":9CF6
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
            Begin VB.TextBox Text3 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00FFFFFF&
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   178
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   315
               Left            =   0
               MaxLength       =   50
               RightToLeft     =   -1  'True
               TabIndex        =   243
               Top             =   -360
               Visible         =   0   'False
               Width           =   1065
            End
         End
         Begin VB.Label lblremain 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "0"
            Height          =   315
            Left            =   240
            RightToLeft     =   -1  'True
            TabIndex        =   248
            Top             =   3120
            Width           =   1275
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00C0FFFF&
            Caption         =   " "
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   15.75
               Charset         =   178
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000080&
            Height          =   300
            Index           =   66
            Left            =   10560
            RightToLeft     =   -1  'True
            TabIndex        =   245
            Top             =   5280
            Width           =   4335
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00C0FFFF&
            Caption         =   " «·œ›⁄«  «· Ì  „ ”œ«œÂ«   ·⁄ﬁœ —ﬁ„"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   15.75
               Charset         =   178
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000080&
            Height          =   300
            Index           =   65
            Left            =   14160
            RightToLeft     =   -1  'True
            TabIndex        =   244
            Top             =   5280
            Width           =   4335
         End
         Begin VB.Shape Shape6 
            BorderWidth     =   2
            FillColor       =   &H00C0FFFF&
            FillStyle       =   0  'Solid
            Height          =   495
            Left            =   3600
            Top             =   5160
            Width           =   15015
         End
         Begin VB.Label lbltotal 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "0"
            Height          =   315
            Left            =   120
            RightToLeft     =   -1  'True
            TabIndex        =   206
            Top             =   2640
            Width           =   1275
         End
         Begin VB.Label lblservice 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "0"
            Height          =   315
            Left            =   2160
            RightToLeft     =   -1  'True
            TabIndex        =   205
            Top             =   2520
            Width           =   1275
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "Œœ„« "
            Height          =   315
            Index           =   58
            Left            =   3480
            RightToLeft     =   -1  'True
            TabIndex        =   204
            Top             =   2520
            Width           =   435
         End
         Begin VB.Label lblcomision 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "0"
            Height          =   315
            Left            =   6000
            RightToLeft     =   -1  'True
            TabIndex        =   203
            Top             =   2520
            Width           =   1275
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "”⁄Ì"
            Height          =   315
            Index           =   57
            Left            =   7320
            RightToLeft     =   -1  'True
            TabIndex        =   202
            Top             =   2520
            Width           =   435
         End
         Begin VB.Label lblrent 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "0"
            Height          =   315
            Left            =   7800
            RightToLeft     =   -1  'True
            TabIndex        =   201
            Top             =   2520
            Width           =   1275
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "«ÌÃ«—"
            Height          =   315
            Index           =   56
            Left            =   9120
            RightToLeft     =   -1  'True
            TabIndex        =   200
            Top             =   2520
            Width           =   435
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00C0FFFF&
            Caption         =   " «·œ›⁄«  «· Ì  „ ”œ«œÂ«  ›Ì Â–« «·”‰œ"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   15.75
               Charset         =   178
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000080&
            Height          =   300
            Index           =   52
            Left            =   14160
            RightToLeft     =   -1  'True
            TabIndex        =   140
            Top             =   3000
            Width           =   4335
         End
         Begin VB.Label Label4 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Height          =   375
            Left            =   4560
            RightToLeft     =   -1  'True
            TabIndex        =   143
            Top             =   840
            Width           =   7575
         End
         Begin VB.Shape Shape5 
            BorderWidth     =   2
            FillColor       =   &H00C0FFFF&
            FillStyle       =   0  'Solid
            Height          =   495
            Left            =   3600
            Top             =   2880
            Width           =   15015
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00C0FFFF&
            Caption         =   "ﬁ„ » ÕœÌœ «·œ›⁄«  «·„—«œ ”œ«œÂ« „‰ «·⁄ﬁœ"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   15.75
               Charset         =   178
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000080&
            Height          =   420
            Index           =   55
            Left            =   14280
            RightToLeft     =   -1  'True
            TabIndex        =   142
            Top             =   360
            Width           =   4215
         End
         Begin VB.Label lbl 
            Alignment       =   2  'Center
            BackColor       =   &H00C0FFFF&
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   14.25
               Charset         =   178
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000080&
            Height          =   420
            Index           =   54
            Left            =   3960
            RightToLeft     =   -1  'True
            TabIndex        =   141
            Top             =   360
            Width           =   3735
         End
         Begin VB.Shape Shape4 
            BorderWidth     =   2
            FillColor       =   &H00C0FFFF&
            FillStyle       =   0  'Solid
            Height          =   495
            Left            =   3840
            Top             =   360
            Width           =   14775
         End
      End
   End
   Begin ImpulseButton.ISButton Cmd 
      Height          =   375
      Index           =   0
      Left            =   13620
      TabIndex        =   101
      Top             =   8160
      Width           =   855
      _ExtentX        =   1508
      _ExtentY        =   661
      ButtonPositionImage=   1
      Caption         =   "ÃœÌœ"
      BackColor       =   14871017
      FontName        =   "Arial"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
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
      Index           =   1
      Left            =   12720
      TabIndex        =   102
      Top             =   8160
      Width           =   855
      _ExtentX        =   1508
      _ExtentY        =   661
      ButtonPositionImage=   1
      Caption         =   " ⁄œÌ·"
      BackColor       =   14871017
      FontName        =   "Arial"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
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
      Index           =   2
      Left            =   11835
      TabIndex        =   103
      Top             =   8160
      Width           =   855
      _ExtentX        =   1508
      _ExtentY        =   661
      ButtonPositionImage=   1
      Caption         =   "Õ›Ÿ"
      BackColor       =   14871017
      FontName        =   "Arial"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
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
      Left            =   10935
      TabIndex        =   104
      Top             =   8160
      Width           =   855
      _ExtentX        =   1508
      _ExtentY        =   661
      ButtonPositionImage=   1
      Caption         =   " —«Ã⁄"
      BackColor       =   14871017
      FontName        =   "Arial"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
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
      Left            =   10050
      TabIndex        =   105
      Top             =   8160
      Width           =   855
      _ExtentX        =   1508
      _ExtentY        =   661
      ButtonPositionImage=   1
      Caption         =   "Õ–›"
      BackColor       =   14871017
      FontName        =   "Arial"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
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
      Left            =   4080
      TabIndex        =   106
      Top             =   8160
      Width           =   855
      _ExtentX        =   1508
      _ExtentY        =   661
      ButtonPositionImage=   1
      Caption         =   "Œ—ÊÃ"
      BackColor       =   14871017
      FontName        =   "Arial"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
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
   Begin ImpulseButton.ISButton CmdHelp 
      Height          =   375
      Left            =   4965
      TabIndex        =   107
      Top             =   7560
      Width           =   855
      _ExtentX        =   1508
      _ExtentY        =   661
      ButtonPositionImage=   1
      Caption         =   "„”«⁄œ…"
      BackColor       =   14871017
      FontName        =   "Arial"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
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
      Left            =   9150
      TabIndex        =   108
      Top             =   8160
      Width           =   855
      _ExtentX        =   1508
      _ExtentY        =   661
      ButtonPositionImage=   1
      Caption         =   "»ÕÀ"
      BackColor       =   14871017
      FontName        =   "Arial"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
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
      Left            =   8265
      TabIndex        =   109
      Top             =   8160
      Width           =   855
      _ExtentX        =   1508
      _ExtentY        =   661
      ButtonPositionImage=   1
      Caption         =   "ÿ»«⁄…"
      BackColor       =   14871017
      FontName        =   "Arial"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
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
      Index           =   9
      Left            =   7080
      TabIndex        =   110
      Top             =   8160
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   661
      ButtonPositionImage=   1
      Caption         =   "ÿ»«⁄… «·ﬁÌœ"
      BackColor       =   14871017
      FontName        =   "Arial"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
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
   Begin MSDataListLib.DataCombo DCboUserName 
      Height          =   315
      Left            =   120
      TabIndex        =   120
      Top             =   8160
      Width           =   2340
      _ExtentX        =   4128
      _ExtentY        =   582
      _Version        =   393216
      Enabled         =   0   'False
      Text            =   ""
      RightToLeft     =   -1  'True
   End
   Begin ImpulseButton.ISButton CmdAttach 
      Height          =   375
      Left            =   6000
      TabIndex        =   134
      Top             =   8160
      Width           =   915
      _ExtentX        =   1614
      _ExtentY        =   661
      ButtonPositionImage=   1
      Caption         =   "«·„—›ﬁ« "
      BackColor       =   14871017
      FontName        =   "Arial"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
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
   Begin ImpulseButton.ISButton ISButton2 
      Height          =   375
      Left            =   5400
      TabIndex        =   214
      TabStop         =   0   'False
      ToolTipText     =   "«÷€ÿ ·«÷«›… ⁄„Ì· ÃœÌœ"
      Top             =   1440
      Width           =   510
      _ExtentX        =   900
      _ExtentY        =   661
      ButtonStyle     =   1
      ButtonPositionImage=   4
      Caption         =   ""
      BackColor       =   14871017
      FontName        =   "Arial"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BackStyle       =   0
      ButtonImage     =   "FrmCashing1.frx":9E88
      ColorButton     =   14871017
      ColorHighlight  =   16777215
      ColorShadow     =   -2147483631
      ColorOutline    =   -2147483631
      DrawFocusRectangle=   0   'False
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   "—ﬁ„ «·⁄ﬁœ"
      Height          =   285
      Index           =   59
      Left            =   7440
      RightToLeft     =   -1  'True
      TabIndex        =   215
      Top             =   1440
      Width           =   795
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
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
      Height          =   435
      Index           =   46
      Left            =   14640
      RightToLeft     =   -1  'True
      TabIndex        =   124
      Top             =   8040
      Width           =   3915
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      BackStyle       =   0  'Transparent
      Caption         =   "Õ—— »Ê«”ÿ… : "
      Height          =   315
      Index           =   8
      Left            =   2280
      RightToLeft     =   -1  'True
      TabIndex        =   121
      Top             =   8160
      Width           =   1410
   End
   Begin VB.Shape Shape1 
      BorderWidth     =   2
      Height          =   495
      Left            =   0
      Top             =   5760
      Width           =   8175
   End
   Begin VB.Label lbl 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFFF&
      Caption         =   "ﬁ„ » ÕœÌœ «·„” Œ·’«   «·„—«œ ”œ«œÂ« ··„‘—Ê⁄"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   420
      Index           =   40
      Left            =   3720
      RightToLeft     =   -1  'True
      TabIndex        =   97
      Top             =   5760
      Width           =   4335
   End
   Begin VB.Label lbl 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFFF&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   420
      Index           =   39
      Left            =   0
      RightToLeft     =   -1  'True
      TabIndex        =   96
      Top             =   5760
      Width           =   3735
   End
End
Attribute VB_Name = "FrmCashing1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rs As ADODB.Recordset
 Dim i As Long
Dim Rs1 As ADODB.Recordset
Dim TTP As clstooltip
Dim cSearchDcbo  As clsDCboSearch
Dim Dcombos As ClsDataCombos
Dim Line1 As Double
Dim FlgNew As Boolean
Dim Line2 As Double
Dim Line3 As Double
Dim netVatPayed As Double
Dim fittervat As Double
Dim Line4 As Double
Dim DayNO As Integer
Dim UonitStatus As Integer
Dim departement_name As Integer
Dim numbering_type As Integer
Dim Balance As String
Dim balanceString As String
Public RereivID As Long
Dim CommissionAcc   As String
Dim CommissionAccDue As String
Dim RentAccount As String
Dim lineno As Double
Sub GetUonitStatus()
 Dim RsDetails1 As ADODB.Recordset
 Dim StrSQL As String

       Set RsDetails1 = New ADODB.Recordset
         StrSQL = "SELECT   Status  from  TblAqarDetai where id =" & val(DcbUnitNo.BoundText) & ""
   RsDetails1.Open StrSQL, Cn, adOpenKeyset, adLockOptimistic, adCmdText
   If RsDetails1.RecordCount > 0 Then
 UonitStatus = val(IIf(IsNull(RsDetails1("Status").value), 0, RsDetails1("Status").value))
   End If
   End Sub
   
   
      Sub SaveUoitInformation(Optional Index As Integer)
 Dim RsDetails1 As ADODB.Recordset
 Dim StrSQL, Msg As String
Msg = ""

 Select Case Index
 Case 8
      If SystemOptions.UserInterface = EnglishInterface Then
      Msg = Msg & "Work was guaranteed catch Contract No."
             Msg = Msg & CHR(13) & txtContractNo.Text
             Msg = Msg & CHR(13)
             Msg = Msg & "  catch Contract Date "
             Msg = Msg & CHR(13) & XPDtbTrans.value
             Msg = Msg & CHR(13)
             Msg = Msg & "  catch Contract Date Arabic "
             Msg = Msg & CHR(13) & Txt_DateHigri.value
              Msg = Msg & CHR(13)
             Msg = Msg & " Value "
             Msg = Msg & CHR(13) & XPTxtVal.Text
      Else
      Msg = Msg & " „ ⁄„· ”‰œﬁ»÷ ·⁄ﬁœ —ﬁ„ "
             Msg = Msg & CHR(13) & txtContractNo.Text
             Msg = Msg & CHR(13)
             Msg = Msg & "   «—ÌŒ «·”‰œ „Ì·«œÌ "
             Msg = Msg & CHR(13) & XPDtbTrans.value
             Msg = Msg & CHR(13)
             Msg = Msg & "   «—ÌŒ «·”‰œ ÂÃ—Ì "
             Msg = Msg & CHR(13) & Txt_DateHigri.value
              Msg = Msg & CHR(13)
             Msg = Msg & "    ﬁÌ„… «·„ﬁ»Ê÷«  "
             Msg = Msg & CHR(13) & XPTxtVal.Text
    End If
 Case 9
       If SystemOptions.UserInterface = EnglishInterface Then
      Msg = Msg & "Action has been arrested token support"
            ' Msg = Msg & Chr(13) & TxtContractNo.text
      Else
      Msg = Msg & " „ ⁄„· ”‰œ ﬁ»÷ ·⁄—»Ê‰ "
             Msg = Msg & CHR(13)
             Msg = Msg & "   «—ÌŒ «·⁄—»Ê‰ „Ì·«œÌ "
             Msg = Msg & CHR(13) & XPDtbTrans.value
             Msg = Msg & CHR(13)
             Msg = Msg & "   «—ÌŒ «·⁄—»Ê‰ ÂÃ—Ì "
             Msg = Msg & CHR(13) & Txt_DateHigri.value
             Msg = Msg & CHR(13)
             Msg = Msg & "    ﬁÌ„… «·⁄—»Ê‰ "
             Msg = Msg & XPTxtVal.Text
             Msg = Msg & CHR(13)
             Msg = Msg & "   «”„ «·„” «Ã— "
             Msg = Msg & CHR(13) & txtrenterName.Text
             Msg = Msg & CHR(13)
             Msg = Msg & "    «·«ÌÃ«— "
             Msg = Msg & TxtRent.Text
           '  Msg = Msg &
             Msg = Msg & "    «·”⁄Ì "
             Msg = Msg & Txtcommission.Text
             Msg = Msg & CHR(13)
             Msg = Msg & "    ”⁄Ì „ﬂ » "
             Msg = Msg & TxtCommissionOut.Text
             Msg = Msg & CHR(13)
             Msg = Msg & "      «„Ì‰ "
             Msg = Msg & TxtCommissionOut.Text
             Msg = Msg & CHR(13)
             Msg = Msg & "     „Ì«Â "
             Msg = Msg & TxtWater.Text
    End If
 
 Case 10
    If SystemOptions.UserInterface = EnglishInterface Then
      Msg = Msg & "Work was guaranteed catch filtering No."
             Msg = Msg & CHR(13) & TxtFilterNo.Text
      Else
      Msg = Msg & " „ ⁄„· ”‰œ ﬁ»÷ · ’›ÌÂ —ﬁ„ "
             Msg = Msg & CHR(13) & TxtFilterNo.Text
              Msg = Msg & CHR(13)
             Msg = Msg & "   «—ÌŒ «·”‰œ „Ì·«œÌ "
             Msg = Msg & CHR(13) & XPDtbTrans.value
             Msg = Msg & CHR(13)
             Msg = Msg & "   «—ÌŒ «·”‰œ ÂÃ—Ì "
             Msg = Msg & CHR(13) & Txt_DateHigri.value
              Msg = Msg & CHR(13)
             Msg = Msg & "    ﬁÌ„… «·„ﬁ»Ê÷«  "
             Msg = Msg & CHR(13) & XPTxtVal.Text
             
    End If

 End Select

   Select Case Index
   
   Case 8, 9, 10
        Set RsDetails1 = New ADODB.Recordset
         StrSQL = "SELECT     *  from  TblUnitNoInformation Where (1 = -1)"
   RsDetails1.Open StrSQL, Cn, adOpenKeyset, adLockOptimistic, adCmdText
      RsDetails1.AddNew
      RsDetails1("CusID").value = val(DcbUnitNo.BoundText)
      RsDetails1("BranchID").value = val(Dcbranch.BoundText)
           RsDetails1("UnitNo").value = val(DcbUnitNo.BoundText)
           RsDetails1("UnitStatus").value = UonitStatus
           RsDetails1("Des").value = Msg
           RsDetails1("RecDate").value = XPDtbTrans.value
           RsDetails1("RecDateH").value = Txt_DateHigri.value
           RsDetails1("NoteID").value = val(XPTxtID.Text)
           RsDetails1("ContNo").value = Null
           RsDetails1("FilterNo").value = Null
           RsDetails1("OrderMaint").value = Null
           RsDetails1("LocOrderMaint").value = Null
           RsDetails1.update

   End Select

   End Sub
Sub InserTypeAmount()
 Dim RsDetails1 As ADODB.Recordset
 Dim StrSQL As String
   StrSQL = "Delete  TblAmoutType  where 1 <> 11 "
                Cn.Execute StrSQL, , adExecuteNoRecords
                
       Set RsDetails1 = New ADODB.Recordset
         StrSQL = "SELECT     *  from  TblAmoutType Where (1 = -1)"
   RsDetails1.Open StrSQL, Cn, adOpenKeyset, adLockOptimistic, adCmdText
   RsDetails1.AddNew
     RsDetails1("id").value = 1
           RsDetails1("name").value = "”⁄Ì œ«Œ·Ì"
           RsDetails1.update
          RsDetails1.AddNew
     RsDetails1("id").value = 2
           RsDetails1("name").value = " «ÌÃ«—« "
           RsDetails1.update
            RsDetails1.AddNew
     RsDetails1("id").value = 3
           RsDetails1("name").value = " ”⁄Ì Œ«—ÃÌ"
           RsDetails1.update
            RsDetails1.AddNew
     RsDetails1("id").value = 4
           RsDetails1("name").value = " ”⁄Ì „‘ —ﬂ"
           RsDetails1.update
           RsDetails1.AddNew
     RsDetails1("id").value = 5
           RsDetails1("name").value = " „»·€  ’›ÌÂ"
           RsDetails1.update
                RsDetails1.update
                RsDetails1.AddNew
     RsDetails1("id").value = 6
           RsDetails1("name").value = "  «Ì—«œ«  «Œ—Ï"
           RsDetails1.update
   End Sub
   Sub NoOfDayl(Optional Index As Integer)
 Dim RsDetails1 As ADODB.Recordset
 Dim StrSQL As String
Dim i, k As Integer
       Set RsDetails1 = New ADODB.Recordset
         StrSQL = "SELECT     *  from  TBLSalesRepGroups where (count <> 0 or count IS NOT NULL)"
   RsDetails1.Open StrSQL, Cn, adOpenKeyset, adLockOptimistic, adCmdText
   If RsDetails1.RecordCount > 0 Then
   Select Case val(RsDetails1("DMY").value) & ""
   
   Case 0
   DayNO = val(RsDetails1("Count").value) & ""
   Case 1
   DayNO = val(RsDetails1("Count").value) & "" * 30
   Case 2
   DayNO = val(RsDetails1("Count").value) & "" * 365
   
   End Select
   End If
   End Sub
   Function GetComm(Optional ID As Double, Optional FiledType As Integer = 0) As Double
   Dim rs2 As ADODB.Recordset
   Set rs2 = New ADODB.Recordset
   Dim sql As String
   sql = "select * from TBLSalesRepGroups"
   If ID <> 0 Then
   sql = sql & " Where ID = " & ID & ""
   End If
   rs2.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
   If rs2.RecordCount > 0 Then
   If FiledType = 1 Then
   GetComm = IIf(IsNull(rs2("InternalComm").value), 0, rs2("InternalComm").value)
   Else
   GetComm = IIf(IsNull(rs2("Rent").value), 0, rs2("Rent").value)
   End If
   Else
   GetComm = 0
   End If
   End Function
Private Function RoundC(ByVal X As Currency, Optional ByVal dp As Integer = 2) As Currency
    Dim m As Double
    m = 10 ^ dp
    RoundC = CCur(VBA.Round(CDbl(X) * m, 0) / m)
End Function

Private Sub BuildEligibleKs(ByRef arrKs() As Integer, ByRef cnt As Long, _
                            ByVal ColName As String, _
                            Optional ByVal AddColName As String = "")
    Dim k As Long
    cnt = 0
    Erase arrKs

    For k = 1 To Grid3.rows - 1
        Dim v As Double
        v = val(Grid3.TextMatrix(k, Grid3.ColIndex(ColName)))
        If AddColName <> "" Then
            v = v + val(Grid3.TextMatrix(k, Grid3.ColIndex(AddColName)))
        End If

        If v <> 0 Then
            cnt = cnt + 1
            ReDim Preserve arrKs(1 To cnt)
            arrKs(cnt) = k
        End If
    Next k
End Sub

Sub AqrCommisiion(Optional Index As Integer)
 Dim RsDetails1 As ADODB.Recordset

 Dim StrSQL As String
Dim i, k, dif As Integer

       Set RsDetails1 = New ADODB.Recordset
         StrSQL = "SELECT     *  from  TblAqarCommissions Where (1 = -1)"
   RsDetails1.Open StrSQL, Cn, adOpenKeyset, adLockOptimistic, adCmdText
Select Case Index
Case 8
    If Grid3.rows > 1 Then

        '==== (A) «·œ›⁄«  «·„ƒÂ·… ··⁄„Ê·«  CommissionsPayed
        Dim KsComm() As Integer, CntComm As Long
        BuildEligibleKs KsComm, CntComm, "CommissionsPayed"

        If VSFlexGrid2.rows > 1 And CntComm > 0 Then
            With VSFlexGrid2

                Dim idxK As Long, kRow As Long
               
                Dim totalVA As Currency, baseVA As Currency, lastVA As Currency, distVA As Currency

                For i = .FixedRows To .rows - 1
                    If val(.TextMatrix(i, .ColIndex("id"))) <> 0 Then
                        If val(.TextMatrix(i, .ColIndex("rate"))) <> 0 Or val(.TextMatrix(i, .ColIndex("ValueAmount"))) <> 0 Then

                            totalVA = CCur(val(.TextMatrix(i, .ColIndex("ValueAmount"))))

                            'ﬁ”¯„ ValueAmount ⁄·Ï ⁄œœ «·œ›⁄«  «·„ƒÂ·… (·Ê ValueAmount <> 0)
                            If totalVA <> 0 Then
                                baseVA = RoundC(totalVA / CntComm, 2)
                                lastVA = totalVA - baseVA * (CntComm - 1)
                            Else
                                baseVA = 0
                                lastVA = 0
                            End If

                            For idxK = 1 To CntComm
                                kRow = KsComm(idxK)

                                RsDetails1.AddNew
                                RsDetails1("PymentNo").value = val(Grid3.TextMatrix(kRow, Grid3.ColIndex("InstallNo")))
                                RsDetails1("ContNo").value = val(TxtContNo.Text)
                                RsDetails1("IqarID").value = val(Me.DcbUnitNo.BoundText)
                                RsDetails1("FilterNo").value = Null
                                RsDetails1("NoteID").value = val(XPTxtID.Text)
                                RsDetails1("TypeOper").value = 8
                                RsDetails1("TypeAmount").value = 1
                                RsDetails1("EmpID").value = val(.TextMatrix(i, .ColIndex("id")))

                                '«· Ê“Ì⁄
                                If totalVA <> 0 Then
                                    If idxK < CntComm Then
                                        distVA = baseVA
                                    Else
                                        distVA = lastVA   '¬Œ— œ›⁄…  «Œœ «·›—ﬁ
                                    End If
                                Else
                                    distVA = 0
                                End If
                                RsDetails1("ValueAmount").value = distVA

                                RsDetails1("Amount").value = (val(.TextMatrix(i, .ColIndex("rate"))) / 100) * _
                                                             val(Grid3.TextMatrix(kRow, Grid3.ColIndex("CommissionsPayed"))) * _
                                                             GetComm(val(.TextMatrix(i, .ColIndex("groupid"))), 1) / 100

                                dif = DateDiff("d", XPDtbTrans, Grid3.TextMatrix(kRow, Grid3.ColIndex("Installdate")))
                                dif = Abs(dif)
                                If dif > DayNO Then
                                    RsDetails1("Flage").value = 1
                                Else
                                    RsDetails1("Flage").value = 0
                                End If

                                If VSFlexGrid2.rows > 2 Then RsDetails1("Crosses").value = 1
                                RsDetails1.update
                            Next idxK

                        End If
                    End If
                Next i

            End With
        End If

        '... dÂ‰ﬂ„¯· «·Ã“¡ «· «‰Ì  Õ 

        '==== (B) «·œ›⁄«  «·„ƒÂ·… ··≈ÌÃ«— RentValuePayed + ActRent
        Dim KsRent() As Integer, CntRent As Long
        BuildEligibleKs KsRent, CntRent, "RentValuePayed", "ActRent"

        If VSFlexGrid2.rows > 1 And CntRent > 0 Then
            With VSFlexGrid2

                Dim idxK2 As Long, kRow2 As Long
                Dim i2 As Long
                Dim totalVA2 As Currency, baseVA2 As Currency, lastVA2 As Currency, distVA2 As Currency

                For i2 = .FixedRows To .rows - 1
                    If val(.TextMatrix(i2, .ColIndex("id"))) <> 0 Then
                        If val(.TextMatrix(i2, .ColIndex("rate"))) <> 0 Or val(.TextMatrix(i2, .ColIndex("ValueAmount"))) <> 0 Then

                            totalVA2 = CCur(val(.TextMatrix(i2, .ColIndex("ValueAmount"))))

                            If totalVA2 <> 0 Then
                                baseVA2 = RoundC(totalVA2 / CntRent, 2)
                                lastVA2 = totalVA2 - baseVA2 * (CntRent - 1)
                            Else
                                baseVA2 = 0
                                lastVA2 = 0
                            End If

                            For idxK2 = 1 To CntRent
                                kRow2 = KsRent(idxK2)

                                RsDetails1.AddNew
                                RsDetails1("PymentNo").value = val(Grid3.TextMatrix(kRow2, Grid3.ColIndex("InstallNo")))
                                RsDetails1("ContNo").value = val(TxtContNo.Text)
                                RsDetails1("FilterNo").value = Null
                                RsDetails1("NoteID").value = val(XPTxtID.Text)
                                RsDetails1("IqarID").value = val(Me.DcbUnitNo.BoundText)
                                RsDetails1("TypeOper").value = 8
                                RsDetails1("TypeAmount").value = 2
                                RsDetails1("EmpID").value = val(.TextMatrix(i2, .ColIndex("id")))

                                If totalVA2 <> 0 Then
                                    If idxK2 < CntRent Then
                                        distVA2 = baseVA2
                                    Else
                                        distVA2 = lastVA2
                                    End If
                                Else
                                    distVA2 = 0
                                End If
                                RsDetails1("ValueAmount").value = distVA2

                                '«·„Â„ Â‰«: ·Ê ValueAmount „ÊÃÊœ…° Amount ·«“„ Ì«Œœ «·„Ê“¯⁄ „‘ «·≈Ã„«·Ì
                                If totalVA2 <> 0 Then
                                    RsDetails1("Amount").value = distVA2
                                Else
                                    RsDetails1("Amount").value = (val(.TextMatrix(i2, .ColIndex("rate"))) / 100) * _
                                                                 val(Grid3.TextMatrix(kRow2, Grid3.ColIndex("RentValuePayed"))) * _
                                                                 GetComm(val(.TextMatrix(i2, .ColIndex("groupid"))), 0) / 100
                                End If

                                dif = DateDiff("d", XPDtbTrans, Grid3.TextMatrix(kRow2, Grid3.ColIndex("Installdate")))
                                dif = Abs(dif)
                                If dif > DayNO Then
                                    RsDetails1("Flage").value = 1
                                Else
                                    RsDetails1("Flage").value = 0
                                End If

                                If VSFlexGrid2.rows > 2 Then RsDetails1("Crosses").value = 1
                                RsDetails1.update
                            Next idxK2

                        End If
                    End If
                Next i2

            End With
        End If

    End If
Case 9
If CheckStatusEarnest(1).value = vbUnchecked Then
     
       If VSFlexGrid1.rows > 1 Then
With VSFlexGrid1
If val(txtComisin.Text) <> 0 Then
For i = .FixedRows To .rows - 1
     If val(.TextMatrix(i, .ColIndex("id"))) <> 0 Then
           RsDetails1.AddNew
           RsDetails1("PymentNo").value = Null
           RsDetails1("ContNo").value = Null
           RsDetails1("FilterNo").value = Null
           RsDetails1("NoteID").value = val(XPTxtID.Text)
           RsDetails1("IqarID").value = val(Me.DcbUnitNo.BoundText)
           RsDetails1("TypeOper").value = 9
           RsDetails1("TypeAmount").value = 1
           RsDetails1("EmpID").value = val(.TextMatrix(i, .ColIndex("id")))
           RsDetails1("Amount").value = (val(.TextMatrix(i, .ColIndex("rate"))) / 100) * val(txtComisin.Text) * GetComm(val(.TextMatrix(i, .ColIndex("groupid"))), 0) / 100
        If CheckStatusEarnest(0).value = vbChecked Then
           RsDetails1("Canceel").value = 1
        End If
        If CheckStatusEarnest(1).value = vbChecked Then
            RsDetails1("ReVal").value = 1
            End If
            If VSFlexGrid1.rows > 2 Then
            RsDetails1("Crosses").value = 1
            End If
           RsDetails1.update
         End If
           Next i
           End If
       
         '''\\\\\
        If val(txtinstranc.Text) <> 0 Then
         For i = .FixedRows To .rows - 1
     If val(.TextMatrix(i, .ColIndex("id"))) <> 0 Then
           RsDetails1.AddNew
           RsDetails1("PymentNo").value = Null
           RsDetails1("ContNo").value = Null
           RsDetails1("FilterNo").value = Null
           RsDetails1("NoteID").value = val(XPTxtID.Text)
           RsDetails1("IqarID").value = val(Me.DcbUnitNo.BoundText)
           RsDetails1("TypeOper").value = 9
           RsDetails1("TypeAmount").value = 2
           RsDetails1("EmpID").value = val(.TextMatrix(i, .ColIndex("id")))
           RsDetails1("Amount").value = (val(.TextMatrix(i, .ColIndex("rate"))) / 100) * val(txtinstranc.Text) * GetComm(val(.TextMatrix(i, .ColIndex("groupid"))), 0) / 100
            If CheckStatusEarnest(0).value = vbChecked Then
            RsDetails1("Canceel").value = 1
            End If
             If CheckStatusEarnest(1).value = vbChecked Then
            RsDetails1("ReVal").value = 1
            End If
            If VSFlexGrid1.rows > 2 Then
            RsDetails1("Crosses").value = 1
            End If
           RsDetails1.update
            End If
                Next i
                End If
  End With
  End If
   End If
Case 12
       If VSFlexGrid1.rows > 1 Then
With VSFlexGrid1

For i = .FixedRows To .rows - 1
     If val(.TextMatrix(i, .ColIndex("id"))) <> 0 Then
           RsDetails1.AddNew
           RsDetails1("PymentNo").value = Null
           RsDetails1("ContNo").value = Null
           RsDetails1("FilterNo").value = Null
           RsDetails1("NoteID").value = val(XPTxtID.Text)
           RsDetails1("TypeOper").value = 12
           RsDetails1("TypeAmount").value = 3
           RsDetails1("EmpID").value = val(.TextMatrix(i, .ColIndex("id")))
           RsDetails1("Amount").value = (val(.TextMatrix(i, .ColIndex("rate"))) / 100) * val(Me.XPTxtVal.Text) * GetComm(val(.TextMatrix(i, .ColIndex("groupid"))), 0) / 100
     
           RsDetails1.update
         End If
           Next i
       
       
    
  End With
  End If

Case 0, 1, 2, 3, 4, 5, 6, 7, 10, 11


     ''///
            If VSFlexGrid1.rows > 1 Then
With VSFlexGrid1

For i = .FixedRows To .rows - 1
     If val(.TextMatrix(i, .ColIndex("id"))) <> 0 Then
           RsDetails1.AddNew
           RsDetails1("PymentNo").value = Null
           RsDetails1("ContNo").value = Null
           RsDetails1("FilterNo").value = Null
           RsDetails1("NoteID").value = val(XPTxtID.Text)
           RsDetails1("TypeOper").value = Index
           RsDetails1("IqarID").value = val(Me.DcbUnitNo.BoundText)
          If Index = 10 Then
           
           RsDetails1("FilterNo").value = val(Me.TxtFilterNo.Text)
           RsDetails1("TypeAmount").value = 5
           Else
           RsDetails1("TypeAmount").value = 6
           End If
           RsDetails1("EmpID").value = val(.TextMatrix(i, .ColIndex("id")))
           RsDetails1("Amount").value = (val(.TextMatrix(i, .ColIndex("rate"))) / 100) * val(Me.XPTxtVal.Text) * GetComm(val(.TextMatrix(i, .ColIndex("groupid"))), 0) / 100
           If Index = 10 Then
                dif = DateDiff("d", XPDtbTrans, Dtaefilter)
                dif = Abs(dif)
                If dif > DayNO Then
                    RsDetails1("Flage").value = 1
                Else
                    RsDetails1("Flage").value = 0
           
                End If
         
         End If
         RsDetails1.update
         End If
           Next i
  End With
  End If
     
End Select

End Sub
Private Sub ALLButton1_Click()
    If IsNumeric(Me.DBCboClientName.BoundText) Then
    End If
End Sub

Private Sub ALLButton2_Click()
    If IsNumeric(Me.DBCboClientName.BoundText) Then
    End If
End Sub

Private Sub ALLButton3_Click()
    lblsqlstring.Caption = ""
    FrmPaymentTime1.show
    FrmPaymentTime1.lblcusid = val(DBCboClientName.BoundText)
    FrmPaymentTime1.LblValue = val(XPTxtVal.Text)
End Sub
Sub GetInstalDate(Optional InstalNo As Double = 0)
If Me.TxtModFlg.Text <> "R" Then
Dim Rs8 As ADODB.Recordset
Dim sql As String
If InstalNo <> 0 Then
Set Rs8 = New ADODB.Recordset
sql = " SELECT     ContNo, InstallNo, Installdate, InstalldateH"
sql = sql & " From dbo.TblContractInstallments"
sql = sql & " Where (ContNo = " & val(TxtContNo.Text) & ") And (InstallNo = " & InstalNo & ")"
Rs8.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
If Rs8.RecordCount > 0 Then
ToPriodDate.value = IIf(IsNull(Rs8("Installdate").value), Date, Rs8("Installdate").value)
ToPriodDateH.value = IIf(IsNull(Rs8("InstalldateH").value), "", Rs8("InstalldateH").value)

End If
End If
End If
End Sub
Function CheckmaxInstal(Optional InstalNo As Double = 0) As Boolean
If Me.TxtModFlg.Text <> "R" Then
Dim Rs8 As ADODB.Recordset
Dim sql As String
If InstalNo <> 0 Then
Set Rs8 = New ADODB.Recordset
sql = "SELECT     ContNo, MAX(InstallNo) AS MaxInstallNo"
sql = sql & " From dbo.TblContractInstallments"
sql = sql & " Where (ContNo = " & val(TxtContNo.Text) & ")"
sql = sql & " GROUP BY ContNo"
sql = sql & " Having (Max(InstallNo) = " & InstalNo & ")"
Rs8.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
If Rs8.RecordCount > 0 Then
CheckmaxInstal = True
Else
CheckmaxInstal = False
End If
End If
End If
End Function


Sub GetInstalMaxDate(Optional InstalNo As Double = 0, Optional PeriodsID As Integer, Optional Periods As Integer)
If Me.TxtModFlg.Text <> "R" Then
Dim Rs8 As ADODB.Recordset
Dim sql As String
If InstalNo <> 0 Then
Set Rs8 = New ADODB.Recordset
sql = " SELECT     ContNo, InstallNo, Installdate, InstalldateH"
sql = sql & " From dbo.TblContractInstallments"
sql = sql & " Where (ContNo = " & val(TxtContNo.Text) & ") And (InstallNo = " & InstalNo & ")"
Rs8.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
If Rs8.RecordCount > 0 Then
ToPriodDate.value = IIf(IsNull(Rs8("Installdate").value), Date, Rs8("Installdate").value)
ToPriodDateH.value = IIf(IsNull(Rs8("InstalldateH").value), "", Rs8("InstalldateH").value)
If PeriodsID = 0 Then
ToPriodDate.value = DateAdd("D", Periods, ToPriodDate.value)
ElseIf PeriodsID = 1 Then
ToPriodDate.value = DateAdd("m", Periods, ToPriodDate.value)
ElseIf PeriodsID = 2 Then
ToPriodDate.value = DateAdd("YYYY", Periods, ToPriodDate.value)
End If
ToPriodDateH.value = ToHijriDate(ToPriodDate.value)
End If
End If
End If
End Sub
Sub GetInstalPeriod(Optional ByRef PeriodsID As Integer, Optional ByRef Periods As Integer)
If Me.TxtModFlg.Text <> "R" Then
Dim Rs8 As ADODB.Recordset
Dim sql As String

Set Rs8 = New ADODB.Recordset
sql = " SELECT     PeriodsID, Periods, ContNo"
sql = sql & " FROM         dbo.TblContract"
sql = sql & " Where (ContNo  = " & val(TxtContNo.Text) & ") "
Rs8.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
If Rs8.RecordCount > 0 Then
PeriodsID = IIf(IsNull(Rs8("PeriodsID").value), -1, Rs8("PeriodsID").value)
Periods = IIf(IsNull(Rs8("Periods").value), "", Rs8("Periods").value)
End If
End If
End Sub
Sub GetBeforInstalDate(Optional InstalNo As Double = 0)
If Me.TxtModFlg.Text <> "R" Then
Dim Rs8 As ADODB.Recordset
Dim sql As String
If InstalNo <> 0 Then
Set Rs8 = New ADODB.Recordset
sql = " SELECT     ContNo, InstallNo, Installdate, InstalldateH"
sql = sql & " From dbo.TblContractInstallments"
sql = sql & " Where (ContNo = " & val(TxtContNo.Text) & ") And (InstallNo = " & InstalNo & ")"
Rs8.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
If Rs8.RecordCount > 0 Then
FrmPriodDate.value = IIf(IsNull(Rs8("Installdate").value), Date, Rs8("Installdate").value)
FrmPriodDateH.value = IIf(IsNull(Rs8("InstalldateH").value), "", Rs8("InstalldateH").value)
End If
End If
End If
End Sub
Public Sub FillGridWithDataContract(NoteSerial1 As String, Optional NoteID As Double)

    'On Error GoTo ErrTrap
    Dim My_SQL As String
    Dim i As Integer
    Dim X As Integer
    Dim rs2 As ADODB.Recordset
    Dim ActRent As Double
    Dim ActOldValue As Double
    Dim ActComm As Double
    Dim ActInsu As Double
    Dim ActWater As Double
    Dim ActElec As Double
    Dim ActService As Double
    Dim ActualTotal As Double
    Dim Result As Double
    Dim resultpercentage As Double
    Dim ActVAT As Double
    Dim sql As String

    Grid3.Clear flexClearScrollable, flexClearEverything
    Grid3.rows = 1
     'VSFlexGrid2.Clear flexClearScrollable, flexClearEverything
    'SFlexGrid2.Rows = 1
    '
    Grid4.Clear flexClearScrollable, flexClearEverything
    Grid4.rows = 1

    If DCboCashType.ListIndex <> 8 Then Exit Sub
 
    lbl(38).Caption = DBCboClientName.Text
    lbl(41).Caption = DBCboClientName.Text
    '
 
sql = " SELECT     dbo.TblContract.ComResid,  dbo.TblContractInstallments.des, dbo.TblContractInstallments.OldValueDate, dbo.TblContractInstallments.OldValueDateH, dbo.TblContractInstallments.OldValue, "
sql = sql & "                         dbo.TblContractInstallments.InstallNo, dbo.TblContractInstallments.Installdate, dbo.TblContractInstallments.InstalldateH, dbo.TblContract.ownerid, dbo.TblContractInstallments.RentValue,"
sql = sql & "                         dbo.TblContractInstallments.Commissions, dbo.TblContractInstallments.Insurance, dbo.TblContractInstallments.Water, dbo.TblContractInstallments.Electric, dbo.TblContractInstallments.TelandNet,"
sql = sql & "                         dbo.TblContractInstallments.RentValuePayed, dbo.TblContractInstallments.CommissionsPayed, dbo.TblContractInstallments.InsurancePayed, dbo.TblContractInstallments.WaterPayed,"
sql = sql & "                         dbo.TblContractInstallments.ElectricPayed, dbo.TblContractInstallments.TelandNetPayed, dbo.TblContract.CusID, dbo.TblContractInstallments.installValue, dbo.TblContractInstallments.Status,"
sql = sql & "                         dbo.TblContractInstallments.ContNo, dbo.TblContractInstallments.id, dbo.TblContract.ContDate, dbo.TblAqarDetai.unitno, dbo.TblAqarDetai.unittype, dbo.TblAkarUnit.name AS unitname,"
sql = sql & "                         dbo.TblAkarUnit.namee AS unitnamee, dbo.TblAqar.Aqarid, dbo.TblAqar.aqarNo, dbo.TblAqar.CountryID, dbo.TblAqar.aqarname, dbo.TblAqar.streetname, dbo.TblCustemers.CusName AS owner,"
sql = sql & "                         dbo.TblCustemers.CusNamee AS ownere, dbo.TblCountriesGovernments.GovernmentName AS Country, dbo.TblCountriesGovernmentsCities.CityName AS hey, dbo.TblContract.StrDate, dbo.TblContract.EndDate,"
sql = sql & "                         dbo.TblContract.MeterValue, dbo.TblContract.MeterCount, dbo.TblContract.TotalContract, dbo.TblContract.PayAmini, dbo.TblContract.CommiValue, dbo.TblContract.InsuranceValue,"
sql = sql & "                         dbo.TblContract.Water AS totalWater, dbo.TblContract.Electricity AS totalElectricity, dbo.TblContract.Enternet AS totalEnternet, dbo.TblContract.Phone AS totalPhone, dbo.TblContract.IncresYearValue,"
sql = sql & "                         dbo.TblContract.IncresYearRate, dbo.TblContract.PaymentCount, dbo.TblContract.FristPaymentDate, dbo.TblContract.PeriodsID, dbo.TblContract.Periods, dbo.TblContract.Furnishing, dbo.TblContract.Remarks,"
sql = sql & "                         dbo.TblContract.RecorddateH, dbo.TblContract.FromdateH, dbo.TblContract.TodateH, dbo.TblContract.FirstInstallDateH, dbo.TblContract.Branch_NO, dbo.TblContract.NewOrOpeneing, dbo.TblContract.OthersRules,"
sql = sql & "                         dbo.TblContract.NoteID, dbo.TblContract.NoteSerial, dbo.TblContract.NoteSerial1, dbo.TblContractInstallments.NoteSerial1 AS NoteSerial1Install, dbo.TblContractInstallments.NoteSerial AS NoteSerialInstall,"
sql = sql & "                         dbo.TblAqarDetai.Id AS UntID, dbo.TblContract.Iqar, dbo.TblContract.UnitType AS UnitTypeID, dbo.TblContract.UnitNo AS UnitNoiD, dbo.TblContractInstallments.ServiceArbon,"
sql = sql & "                         dbo.TblContractInstallments.WaterArbon, dbo.TblContractInstallments.ElectricArbon, dbo.TblContractInstallments.NetElectric, dbo.TblContractInstallments.Electric1, dbo.TblContractInstallments.NetWater,"
sql = sql & "                         dbo.TblContractInstallments.Water1, dbo.TblContractInstallments.NetInsurance, dbo.TblContractInstallments.InsuranceArbon, dbo.TblContractInstallments.Insurance1,"
sql = sql & "                         dbo.TblContractInstallments.NetCommissions, dbo.TblContractInstallments.CommissionsArbon, dbo.TblContractInstallments.Commissions1, dbo.TblContractInstallments.RentArbon,dbo.TblContractInstallments.VATArboon,"
sql = sql & "                         dbo.TblContractInstallments.NetRent , dbo.TblContractInstallments.Rent1, dbo.TblContractInstallments.OldValuePayed, dbo.TblContractInstallments.VATPayed, dbo.TblContractInstallments.VATValue,TblContractInstallments.VATValue1Com,TblContractInstallments.VATValue2Com"
sql = sql & " FROM            dbo.TblCountriesGovernmentsCities RIGHT OUTER JOIN"
sql = sql & "                         dbo.TblAqar INNER JOIN"
sql = sql & "                         dbo.TblAqarDetai ON dbo.TblAqar.Aqarid = dbo.TblAqarDetai.Aqarid LEFT OUTER JOIN"
sql = sql & "                         dbo.TblCountriesGovernments ON dbo.TblAqar.cityid = dbo.TblCountriesGovernments.GovernmentID ON dbo.TblCountriesGovernmentsCities.CityID = dbo.TblAqar.heyid LEFT OUTER JOIN"
sql = sql & "                         dbo.TblAkarUnit ON dbo.TblAqarDetai.unittype = dbo.TblAkarUnit.id RIGHT OUTER JOIN"
sql = sql & "                         dbo.TblContractInstallments RIGHT OUTER JOIN"
sql = sql & "                         dbo.TblCustemers RIGHT OUTER JOIN"
sql = sql & "                         dbo.TblContract ON dbo.TblCustemers.CusID = dbo.TblContract.ownerid ON dbo.TblContractInstallments.ContNo = dbo.TblContract.ContNo ON dbo.TblAqarDetai.Id = dbo.TblContract.UnitNo"
sql = sql & "   WHERE     ( (dbo.TblContractInstallments.Status is null  or dbo.TblContractInstallments.Status=0)  and  dbo.TblContract.ContNo =" & val(TxtContNo.Text) & ")"
sql = sql & " order by dbo.TblContractInstallments.InstallNo"
    Set rs2 = New ADODB.Recordset
    rs2.Open sql, Cn, adOpenStatic, adLockReadOnly, adCmdText

    If rs2.RecordCount = 0 Then
 
        Exit Sub
    End If

    i = 0

    With Me.Grid3
        .rows = 1
        .Clear flexClearScrollable
  
        rs2.MoveFirst
DBCboClientName.BoundText = IIf(IsNull(rs2.Fields("CusID").value), "", rs2.Fields("CusID").value)
''//12 05 2015

Me.DcbIqara.BoundText = IIf(IsNull(rs2.Fields("Iqar").value), "", rs2.Fields("Iqar").value)
Me.DcbUnitType.BoundText = IIf(IsNull(rs2.Fields("UnitTypeID").value), "", rs2.Fields("UnitTypeID").value)
Me.DcbUnitNo.BoundText = IIf(IsNull(rs2.Fields("UnitNoiD").value), "", rs2.Fields("UnitNoiD").value)
     If Not IsNull(rs2("ComResid").value) Then
        If (rs2("ComResid").value) = 1 Then
        ComResid(1).value = True
        Else
        ComResid(0).value = True
        End If
        Else
        ComResid(0).value = True
        End If
        
         For X = 1 To rs2.RecordCount
          If Me.TxtModFlg.Text <> "E" Then
            ActualTotal = getinsttPayedTocontract(val(rs2.Fields("id").value), ActRent, ActComm, ActInsu, ActWater, ActElec, ActService, ActOldValue, , , ActVAT)
          Else
          ActualTotal = getinsttPayedTocontract(val(rs2.Fields("id").value), ActRent, ActComm, ActInsu, ActWater, ActElec, ActService, ActOldValue, NoteID, 1, ActVAT)
          End If
            Result = Round(IIf(IsNull(rs2.Fields("installValue").value), 0, (rs2.Fields("installValue").value)) - ActualTotal, 2)
            If rs2.Fields("installValue").value <> 0 Then
            resultpercentage = Round((ActualTotal / val(rs2.Fields("installValue").value)) * 100, 2)
            Else
            resultpercentage = 0
           End If
            If val(rs2.Fields("installValue").value) > ActualTotal Then
                i = i + 1
                .rows = .rows + 1
                .TextMatrix(i, .ColIndex("Ser")) = i
                .TextMatrix(i, .ColIndex("Select")) = -1
                .TextMatrix(i, .ColIndex("id")) = IIf(IsNull(rs2.Fields("id").value), "", rs2.Fields("id").value)
            
                '                             .TextMatrix(I, .ColIndex("bill_id")) = IIf(IsNull(rs2.Fields("bill_id").value), _
                                              "", rs2.Fields("bill_id").value)
            
                .TextMatrix(i, .ColIndex("Installdate")) = IIf(IsNull(rs2.Fields("Installdate").value), "", rs2.Fields("Installdate").value)
                .TextMatrix(i, .ColIndex("Installdateh")) = IIf(IsNull(rs2.Fields("Installdateh").value), "", rs2.Fields("Installdateh").value)
              
             Dim datedifferent As Double
             datedifferent = DateDiff("d", .TextMatrix(i, .ColIndex("Installdate")), XPDtbTrans.value)
             
             If datedifferent <= 30 Then
                 .TextMatrix(i, .ColIndex("CommisionTypesid")) = 1
                  .TextMatrix(i, .ColIndex("CommisionTypes")) = " ”ÊÌﬁ"
             Else
               .TextMatrix(i, .ColIndex("CommisionTypesid")) = 2
                  .TextMatrix(i, .ColIndex("CommisionTypes")) = " Õ’Ì·"
             End If
             
              
               .TextMatrix(i, .ColIndex("VATValue1Com")) = IIf(IsNull(rs2.Fields("VATValue1Com").value), 0, rs2.Fields("VATValue1Com").value)
                .TextMatrix(i, .ColIndex("VATValue2Com")) = IIf(IsNull(rs2.Fields("VATValue2Com").value), 0, rs2.Fields("VATValue2Com").value)
               .TextMatrix(i, .ColIndex("total")) = IIf(IsNull(rs2.Fields("installValue").value), 0, rs2.Fields("installValue").value)
 
              .TextMatrix(i, .ColIndex("OldValueDate")) = IIf(IsNull(rs2.Fields("OldValueDate").value), "", rs2.Fields("OldValueDate").value)
                .TextMatrix(i, .ColIndex("OldValueDateH")) = IIf(IsNull(rs2.Fields("OldValueDateH").value), "", rs2.Fields("OldValueDateH").value)
              .TextMatrix(i, .ColIndex("des")) = IIf(IsNull(rs2.Fields("des").value), "", rs2.Fields("des").value)
               .TextMatrix(i, .ColIndex("OldValue")) = IIf(IsNull(rs2.Fields("OldValue").value), 0, rs2.Fields("OldValue").value)
                
                
               .TextMatrix(i, .ColIndex("InstallNo")) = IIf(IsNull(rs2.Fields("InstallNo").value), "", rs2.Fields("InstallNo").value)
                
                
                .TextMatrix(i, .ColIndex("ActualTotal")) = ActualTotal
                .TextMatrix(i, .ColIndex("ResultPercentage")) = resultpercentage
                .TextMatrix(i, .ColIndex("Result")) = Result
                .TextMatrix(i, .ColIndex("ActRent")) = ActRent
                .TextMatrix(i, .ColIndex("ActOldValue")) = ActOldValue
                .TextMatrix(i, .ColIndex("ActComm")) = ActComm
                .TextMatrix(i, .ColIndex("ActInsu")) = ActInsu
                .TextMatrix(i, .ColIndex("ActWater")) = ActWater
                .TextMatrix(i, .ColIndex("ActElec")) = ActElec
                .TextMatrix(i, .ColIndex("ActService")) = ActService
                .TextMatrix(i, .ColIndex("ActVAT")) = ActVAT
 
     'RentValue,Commissions,Insurance,Water,Electric,TelandNet
     'RentValuePayed,CommissionsPayed,InsurancePayed,WaterPayed,ElectricPayed,TelandNetPayed
     If ActRent = 0 And ActComm = 0 And ActInsu = 0 And ActWater = 0 And ActElec = 0 And ActService = 0 Then
     .TextMatrix(i, .ColIndex("ActVAT")) = (IIf(IsNull(rs2.Fields("VATArboon").value), 0, rs2.Fields("VATArboon").value))
     .TextMatrix(i, .ColIndex("ActRent")) = (IIf(IsNull(rs2.Fields("RentArbon").value), 0, rs2.Fields("RentArbon").value))
     .TextMatrix(i, .ColIndex("ActComm")) = (IIf(IsNull(rs2.Fields("CommissionsArbon").value), 0, rs2.Fields("CommissionsArbon").value))
     .TextMatrix(i, .ColIndex("ActInsu")) = (IIf(IsNull(rs2.Fields("InsuranceArbon").value), 0, rs2.Fields("InsuranceArbon").value))
     .TextMatrix(i, .ColIndex("ActWater")) = IIf(IsNull(rs2.Fields("WaterArbon").value), 0, rs2.Fields("WaterArbon").value)
     .TextMatrix(i, .ColIndex("ActElec")) = (IIf(IsNull(rs2.Fields("ElectricArbon").value), 0, rs2.Fields("ElectricArbon").value))
     .TextMatrix(i, .ColIndex("ActService")) = (IIf(IsNull(rs2.Fields("ServiceArbon").value), 0, rs2.Fields("ServiceArbon").value))
     '.TextMatrix(i, .ColIndex("ActOldValue")) = (IIf(IsNull(rs2.Fields("ActOldValue").value), 0, rs2.Fields("ActOldValue").value))
     
     .TextMatrix(i, .ColIndex("RentArbon")) = (IIf(IsNull(rs2.Fields("RentArbon").value), 0, rs2.Fields("RentArbon").value))
     .TextMatrix(i, .ColIndex("VATArboon")) = (IIf(IsNull(rs2.Fields("VATArboon").value), 0, rs2.Fields("VATArboon").value))
     
     .TextMatrix(i, .ColIndex("CommissionsArbon")) = (IIf(IsNull(rs2.Fields("CommissionsArbon").value), 0, rs2.Fields("CommissionsArbon").value))
     .TextMatrix(i, .ColIndex("InsuranceArbon")) = (IIf(IsNull(rs2.Fields("InsuranceArbon").value), 0, rs2.Fields("InsuranceArbon").value))
     .TextMatrix(i, .ColIndex("WaterArbon")) = IIf(IsNull(rs2.Fields("WaterArbon").value), 0, rs2.Fields("WaterArbon").value)
     .TextMatrix(i, .ColIndex("ElectricArbon")) = (IIf(IsNull(rs2.Fields("ElectricArbon").value), 0, rs2.Fields("ElectricArbon").value))
     .TextMatrix(i, .ColIndex("ServiceArbon")) = (IIf(IsNull(rs2.Fields("ServiceArbon").value), 0, rs2.Fields("ServiceArbon").value))
     End If
     .TextMatrix(i, .ColIndex("RentValue")) = (IIf(IsNull(rs2.Fields("RentValue").value), 0, rs2.Fields("RentValue").value))
     If SystemOptions.NoCreatJLInRentContract = False Then
     .TextMatrix(i, .ColIndex("VATValue")) = (IIf(IsNull(rs2.Fields("VATValue").value), 0, rs2.Fields("VATValue").value))
     Else
     .TextMatrix(i, .ColIndex("VATValue")) = 0
     'salim here 19 08 2019
     .TextMatrix(i, .ColIndex("VATValue")) = (IIf(IsNull(rs2.Fields("VATValue").value), 0, rs2.Fields("VATValue").value))
     
     End If
    .TextMatrix(i, .ColIndex("Commissions")) = (IIf(IsNull(rs2.Fields("Commissions").value), 0, rs2.Fields("Commissions").value))
     
    .TextMatrix(i, .ColIndex("Insurance")) = (IIf(IsNull(rs2.Fields("Insurance").value), 0, rs2.Fields("Insurance").value))
    
    .TextMatrix(i, .ColIndex("Water")) = (IIf(IsNull(rs2.Fields("Water").value), 0, rs2.Fields("Water").value))
    
    .TextMatrix(i, .ColIndex("Electric")) = (IIf(IsNull(rs2.Fields("Electric").value), 0, rs2.Fields("Electric").value))
    
    .TextMatrix(i, .ColIndex("TelandNet")) = (IIf(IsNull(rs2.Fields("TelandNet").value), 0, rs2.Fields("TelandNet").value))
     
    .TextMatrix(i, .ColIndex("RentValuePayed")) = 0
   .TextMatrix(i, .ColIndex("CommissionsPayed")) = 0
   .TextMatrix(i, .ColIndex("InsurancePayed")) = 0
   .TextMatrix(i, .ColIndex("WaterPayed")) = 0
   .TextMatrix(i, .ColIndex("ElectricPayed")) = 0
   .TextMatrix(i, .ColIndex("TelandNetPayed")) = 0
   
   If NoteID <> 0 And Me.TxtModFlg.Text = "E" Then
   getinsttPayedToContNote NoteID, ActRent, ActComm, ActInsu, ActWater, ActElec, ActService, ActOldValue, rs2.Fields("ID").value, ActVAT
   .TextMatrix(i, .ColIndex("RentValuePayed")) = ActRent
   .TextMatrix(i, .ColIndex("CommissionsPayed")) = ActComm
   .TextMatrix(i, .ColIndex("InsurancePayed")) = ActInsu
   .TextMatrix(i, .ColIndex("WaterPayed")) = ActWater
   .TextMatrix(i, .ColIndex("ElectricPayed")) = ActElec
   .TextMatrix(i, .ColIndex("TelandNetPayed")) = ActService
   .TextMatrix(i, .ColIndex("VATPayed")) = ActVAT
   .TextMatrix(i, .ColIndex("OldValuePayed")) = ActOldValue
   Else
   GoTo l
    .TextMatrix(i, .ColIndex("RentValuePayed")) = (IIf(IsNull(rs2.Fields("RentValuePayed").value), 0, rs2.Fields("RentValuePayed").value))
    .TextMatrix(i, .ColIndex("CommissionsPayed")) = (IIf(IsNull(rs2.Fields("CommissionsPayed").value), 0, rs2.Fields("CommissionsPayed").value))
    .TextMatrix(i, .ColIndex("InsurancePayed")) = (IIf(IsNull(rs2.Fields("InsurancePayed").value), 0, rs2.Fields("InsurancePayed").value))
    .TextMatrix(i, .ColIndex("WaterPayed")) = (IIf(IsNull(rs2.Fields("WaterPayed").value), 0, rs2.Fields("WaterPayed").value))
    .TextMatrix(i, .ColIndex("ElectricPayed")) = (IIf(IsNull(rs2.Fields("ElectricPayed").value), 0, rs2.Fields("ElectricPayed").value))
    .TextMatrix(i, .ColIndex("OldValuePayed")) = (IIf(IsNull(rs2.Fields("OldValuePayed").value), 0, rs2.Fields("OldValuePayed").value))
    .TextMatrix(i, .ColIndex("TelandNetPayed")) = (IIf(IsNull(rs2.Fields("TelandNetPayed").value), 0, rs2.Fields("TelandNetPayed").value))
    .TextMatrix(i, .ColIndex("VATPayed")) = (IIf(IsNull(rs2.Fields("VATPayed").value), 0, rs2.Fields("VATPayed").value))
End If
            
l:            End If

            rs2.MoveNext
        Next

        rs2.Close
 
        .RowHeight(-1) = 300
    End With

    If TxtNoteSerial = "" Then

        Exit Sub
    End If
'  rs2("NoteID").value = val(XPTxtID.text)
    sql = "SELECT  * FROM     ContracttBillInstallmentsDone     where NoteID =" & val(XPTxtID.Text)
 
   ' rs2.Open sql, Cn, adOpenStatic, adLockReadOnly, adCmdText


    'If rs2.RecordCount = 0 Then
 
    '    Exit Sub
    'End If
 
      'rs2.Open sql, Cn, adOpenStatic, adLockReadOnly, adCmdText


  sql = "SELECT  * FROM     ContracttBillInstallmentsDone     where NoteID =" & val(XPTxtID.Text)
 rs2.Open sql, Cn, adOpenStatic, adLockReadOnly, adCmdText
    If rs2.RecordCount = 0 Then
 
        Exit Sub
    End If
    With Me.Grid4
        .rows = 1
        .rows = .rows + rs2.RecordCount
        .Clear flexClearScrollable
  
        rs2.MoveFirst

        For i = 1 To .rows - 1
 
            .TextMatrix(i, .ColIndex("Ser")) = i
            .TextMatrix(i, .ColIndex("id")) = IIf(IsNull(rs2.Fields("id").value), "", rs2.Fields("id").value)
        
            .TextMatrix(i, .ColIndex("Installdate")) = IIf(IsNull(rs2.Fields("RecordDate").value), "", rs2.Fields("RecordDate").value)
              .TextMatrix(i, .ColIndex("Installdateh")) = IIf(IsNull(rs2.Fields("RecordDateh").value), "", rs2.Fields("RecordDateh").value)
              
            .TextMatrix(i, .ColIndex("InstallNo")) = IIf(IsNull(rs2.Fields("InstallNo").value), "", rs2.Fields("InstallNo").value)
 
            .TextMatrix(i, .ColIndex("total")) = IIf(IsNull(rs2.Fields("total").value), "", rs2.Fields("total").value)
            
            .TextMatrix(i, .ColIndex("ActualTotal")) = IIf(IsNull(rs2.Fields("value").value), "", rs2.Fields("value").value)
            Result = val(.TextMatrix(i, .ColIndex("total"))) - getinsttPayedTocontract(val(rs2.Fields("istallid").value)) '
            resultpercentage = val(rs2.Fields("value").value) / val(.TextMatrix(i, .ColIndex("total"))) * 100
            If resultpercentage > 100 Then resultpercentage = 100
            .TextMatrix(i, .ColIndex("ResultPercentage")) = Round(resultpercentage, 2)
            If Result < 0 Then Result = 0
            .TextMatrix(i, .ColIndex("Result")) = Result
    .TextMatrix(i, .ColIndex("VATPayed")) = (IIf(IsNull(rs2.Fields("VATPayed").value), 0, rs2.Fields("VATPayed").value))
    .TextMatrix(i, .ColIndex("RentValuePayed")) = (IIf(IsNull(rs2.Fields("RentValuePayed").value), 0, rs2.Fields("RentValuePayed").value))
    .TextMatrix(i, .ColIndex("CommissionsPayed")) = (IIf(IsNull(rs2.Fields("CommissionsPayed").value), 0, rs2.Fields("CommissionsPayed").value))
    .TextMatrix(i, .ColIndex("InsurancePayed")) = (IIf(IsNull(rs2.Fields("InsurancePayed").value), 0, rs2.Fields("InsurancePayed").value))
    .TextMatrix(i, .ColIndex("WaterPayed")) = (IIf(IsNull(rs2.Fields("WaterPayed").value), 0, rs2.Fields("WaterPayed").value))
    .TextMatrix(i, .ColIndex("ElectricPayed")) = (IIf(IsNull(rs2.Fields("ElectricPayed").value), 0, rs2.Fields("ElectricPayed").value))
    .TextMatrix(i, .ColIndex("OldValuePayed")) = (IIf(IsNull(rs2.Fields("OldValuePayed").value), 0, rs2.Fields("OldValuePayed").value))
    .TextMatrix(i, .ColIndex("TelandNetPayed")) = (IIf(IsNull(rs2.Fields("TelandNetPayed").value), 0, rs2.Fields("TelandNetPayed").value))
   


            rs2.MoveNext
        Next






        rs2.Close
 
        .RowHeight(-1) = 300
    End With
ReLineGrid 1

ErrTrap:
End Sub
Sub FillGridWithDatSales(NoteSerial1 As String)
  

    'On Error GoTo ErrTrap
Dim My_SQL As String
    Dim i As Integer
    Dim X As Integer
    Dim rs2 As ADODB.Recordset
 
    Dim ActualTotal As Double
    Dim Result As Double
    Dim resultpercentage As Double
    Dim sql As String

  
      VSFlexGrid2.Clear flexClearScrollable, flexClearEverything
    VSFlexGrid2.rows = 1
        Set rs2 = New ADODB.Recordset
My_SQL = "SELECT     dbo.TblCOntractSales.ContNo, dbo.TblCOntractSales.ID, dbo.TblCOntractSales.rate, dbo.TblCOntractSales.EmpID, dbo.TblEmployee.Emp_Name, "
My_SQL = My_SQL & "                      dbo.TblEmployee.Fullcode, dbo.TblEmployee.Emp_Namee, dbo.TblCOntractSales.idd, dbo.TblCOntractSales.GroupID, dbo.TBLSalesRepGroups.name,"
My_SQL = My_SQL & "                      dbo.TBLSalesRepGroups.NameE"
My_SQL = My_SQL & " FROM         dbo.TblCOntractSales LEFT OUTER JOIN"
My_SQL = My_SQL & "                      dbo.TBLSalesRepGroups ON dbo.TblCOntractSales.GroupID = dbo.TBLSalesRepGroups.id LEFT OUTER JOIN"
My_SQL = My_SQL & "                      dbo.TblEmployee ON dbo.TblCOntractSales.EmpID = dbo.TblEmployee.Emp_ID"
My_SQL = My_SQL & " Where (dbo.TblCOntractSales.ContNo =" & val(TxtContNo.Text) & ")"

    rs2.Open My_SQL, Cn, adOpenKeyset, adLockReadOnly, adCmdText
'    rs1.Open My_SQL, Cn, adOpenStatic, adLockReadOnly, adCmdText
      With Me.VSFlexGrid2
       .rows = 1
        .Clear flexClearScrollable

        If rs2.RecordCount > 0 Then
           .rows = rs2.RecordCount + 1
           rs2.MoveFirst

            For i = 1 To .rows - 1
               .TextMatrix(i, .ColIndex("Ser")) = i
   If SystemOptions.UserInterface = EnglishInterface Then
       .TextMatrix(i, .ColIndex("empname")) = IIf(IsNull(rs2.Fields("Emp_Namee").value), "", rs2.Fields("Emp_Namee").value)
      .TextMatrix(i, .ColIndex("group")) = IIf(IsNull(rs2.Fields("namee").value), "", rs2.Fields("namee").value)
      Else
      .TextMatrix(i, .ColIndex("empname")) = IIf(IsNull(rs2.Fields("Emp_Name").value), "", rs2.Fields("Emp_Name").value)
      .TextMatrix(i, .ColIndex("group")) = IIf(IsNull(rs2.Fields("name").value), "", rs2.Fields("name").value)
 
    End If
    .TextMatrix(i, .ColIndex("groupid")) = val(IIf(IsNull(rs2.Fields("GroupID").value), "", rs2.Fields("GroupID").value))
 .TextMatrix(i, .ColIndex("rate")) = val(IIf(IsNull(rs2.Fields("rate").value), "", rs2.Fields("rate").value))
  .TextMatrix(i, .ColIndex("code")) = IIf(IsNull(rs2.Fields("Fullcode").value), "", rs2.Fields("Fullcode").value)
  .TextMatrix(i, .ColIndex("id")) = IIf(IsNull(rs2.Fields("EmpID").value), "", rs2.Fields("EmpID").value)
  .TextMatrix(i, .ColIndex("idd")) = IIf(IsNull(rs2.Fields("idd").value), "", rs2.Fields("idd").value)
        rs2.MoveNext
            Next i

         
        End If

        .RowHeight(-1) = 300
    End With
    VSFlexGrid2.rows = VSFlexGrid2.rows + 1
 ReLineGrid
End Sub


Public Sub FillGridWithData(project_no As Long, _
                            Optional TxtNoteSerial As String)

    'On Error GoTo ErrTrap

    Dim i As Integer
    Dim X As Integer
    Dim rs2 As ADODB.Recordset
 
    Dim ActualTotal As Double
    Dim Result As Double
    Dim resultpercentage As Double
    Dim sql As String

    Grid.Clear flexClearScrollable, flexClearEverything
    Grid.rows = 1
          
    GRID1.Clear flexClearScrollable, flexClearEverything
    GRID1.rows = 1

    If DCboCashType.ListIndex <> 5 Then Exit Sub
 
    lbl(38).Caption = DBCboClientName.Text
    lbl(41).Caption = DBCboClientName.Text
    sql = "SELECT  * FROM     project_billl     where project_no = " & project_no
    Set rs2 = New ADODB.Recordset
    rs2.Open sql, Cn, adOpenStatic, adLockReadOnly, adCmdText

    If rs2.RecordCount = 0 Then
 
        Exit Sub
    End If

    i = 0

    With Me.Grid
        .rows = 1
        .Clear flexClearScrollable
  
        rs2.MoveFirst

        For X = 1 To rs2.RecordCount
       
            ActualTotal = getBillPayedToproject(val(rs2.Fields("id").value))
            Result = val(rs2.Fields("total").value) - ActualTotal
            resultpercentage = Round((ActualTotal / val(rs2.Fields("total").value)) * 100, 2)
 
            If val(rs2.Fields("total").value) > ActualTotal Then
                i = i + 1
                .rows = .rows + 1
                .TextMatrix(i, .ColIndex("Ser")) = i
                .TextMatrix(i, .ColIndex("id")) = IIf(IsNull(rs2.Fields("id").value), "", rs2.Fields("id").value)
            
                '                             .TextMatrix(I, .ColIndex("bill_id")) = IIf(IsNull(rs2.Fields("bill_id").value), _
                                              "", rs2.Fields("bill_id").value)
            
                .TextMatrix(i, .ColIndex("bill_date")) = IIf(IsNull(rs2.Fields("bill_date").value), "", rs2.Fields("bill_date").value)
                .TextMatrix(i, .ColIndex("project_no")) = IIf(IsNull(rs2.Fields("project_no").value), "", rs2.Fields("project_no").value)
                .TextMatrix(i, .ColIndex("Project_name")) = IIf(IsNull(rs2.Fields("project_name").value), "", rs2.Fields("project_name").value)
            
                .TextMatrix(i, .ColIndex("End_user_name")) = IIf(IsNull(rs2.Fields("End_user_name").value), "", rs2.Fields("End_user_name").value)
            
                .TextMatrix(i, .ColIndex("Sub_user_name")) = IIf(IsNull(rs2.Fields("Sub_user_name").value), "", rs2.Fields("Sub_user_name").value)
            
                .TextMatrix(i, .ColIndex("bill_to")) = IIf(IsNull(rs2.Fields("bill_to").value), "", rs2.Fields("bill_to").value)
 
                .TextMatrix(i, .ColIndex("total")) = IIf(IsNull(rs2.Fields("total").value), "", rs2.Fields("total").value)
            
                .TextMatrix(i, .ColIndex("ActualTotal")) = ActualTotal
                .TextMatrix(i, .ColIndex("ResultPercentage")) = resultpercentage
                .TextMatrix(i, .ColIndex("Result")) = Result

            End If

            rs2.MoveNext
        Next

        rs2.Close
 
        .RowHeight(-1) = 300
    End With

    If TxtNoteSerial = "" Then

        Exit Sub
    End If

    sql = "SELECT  * FROM     ProjectBillBuy     where TxtNoteSerial ='" & TxtNoteSerial & "'"
 
    rs2.Open sql, Cn, adOpenStatic, adLockReadOnly, adCmdText

    If rs2.RecordCount = 0 Then
 
        Exit Sub
    End If
 
    With Me.GRID1
        .rows = 1
        .rows = .rows + rs2.RecordCount
        .Clear flexClearScrollable
  
        rs2.MoveFirst

        For i = 1 To .rows - 1
 
            .TextMatrix(i, .ColIndex("Ser")) = i
            .TextMatrix(i, .ColIndex("id")) = IIf(IsNull(rs2.Fields("id").value), "", rs2.Fields("id").value)
            
            .TextMatrix(i, .ColIndex("bill_id")) = IIf(IsNull(rs2.Fields("bill_id").value), "", rs2.Fields("bill_id").value)
            
            .TextMatrix(i, .ColIndex("bill_date")) = IIf(IsNull(rs2.Fields("RecordDate").value), "", rs2.Fields("RecordDate").value)
            '                                           .TextMatrix(I, .ColIndex("project_no")) = IIf(IsNull(rs2.Fields("project_no").value), _
                                                        "", rs2.Fields("project_no").value)
            '                         .TextMatrix(I, .ColIndex("Project_name")) = IIf(IsNull(rs2.Fields("project_name").value), _
                                      "", rs2.Fields("project_name").value)
            
            .TextMatrix(i, .ColIndex("bill_to")) = IIf(IsNull(rs2.Fields("bill_to").value), "", rs2.Fields("bill_to").value)
 
            .TextMatrix(i, .ColIndex("total")) = IIf(IsNull(rs2.Fields("total").value), "", rs2.Fields("total").value)
            
            .TextMatrix(i, .ColIndex("ActualTotal")) = IIf(IsNull(rs2.Fields("value").value), "", rs2.Fields("value").value)
            Result = val(.TextMatrix(i, .ColIndex("total"))) - val(rs2.Fields("value").value)
            resultpercentage = val(rs2.Fields("value").value) / val(.TextMatrix(i, .ColIndex("total"))) * 100
            .TextMatrix(i, .ColIndex("ResultPercentage")) = resultpercentage
            .TextMatrix(i, .ColIndex("Result")) = Result
      
            rs2.MoveNext
        Next

        rs2.Close
 
        .RowHeight(-1) = 300
    End With

ErrTrap:
End Sub

Private Sub ALLButton4_Click()

    If DCboCashType.ListIndex <> 5 Then
        If SystemOptions.UserInterface = ArabicInterface Then
            MsgBox "Â–… «·⁄„·Ì… „ «Õ… „⁄ ›Ê« Ì— «·„‘«—Ì⁄ ›ﬁÿ", vbInformation
        Else
            MsgBox "This Process For Project Bill Only", vbInformation
    
        End If

        DCboCashType.SetFocus
        Sendkeys "{F4}"
        Exit Sub
    End If

    If val(DBCboClientName.BoundText) = 0 Then
        If SystemOptions.UserInterface = ArabicInterface Then
            MsgBox "«Œ — „‘—Ê⁄ «Ê·«", vbInformation
        Else
            MsgBox "select Project Firstly, vbInformation"
    
        End If

        DBCboClientName.SetFocus
        Sendkeys "{F4}"
        Exit Sub

    End If
 
    FillGridWithData val(Me.DBCboClientName.BoundText), TxtNoteSerial.Text

End Sub

Private Sub CboPayMentType_Change()
DBCboClientName_Change

DcboBox_Change
DcboBankName_Click (0)
DcbAccount_Change
   Me.lbl(9).Enabled = False
        Me.DcboBox.Enabled = False
        Me.lbl(15).Enabled = False
        Me.lbl(16).Enabled = False
        Me.lbl(17).Enabled = False
        Me.DcboBankName.Enabled = False
        Me.TxtChequeNumber.Enabled = False
        Me.DtpChequeDueDate.Enabled = False
        TxtAccount.Enabled = False
        DcbAccount.Enabled = False
    If Me.TxtModFlg.Text = "E" Then
        DcboBankName.Text = ""
        TxtChequeNumber.Text = ""
        Me.DcboBox.Text = ""
        DCChequeBox.Text = ""
        TXTBankName.Text = ""
    End If

    DCChequeBox.Enabled = False

    If SystemOptions.UserInterface = ArabicInterface Then
        lbl(16).Caption = "—ﬁ„ «·‘Ìﬂ"
        lbl(17).Caption = " «—ÌŒ «·«” Õﬁ«ﬁ"
    
    Else
        lbl(16).Caption = "Cheque No"
        lbl(17).Caption = "Due Date"
    End If
    
    If Me.CboPaymentType.ListIndex = 0 Then
       DcbAccount.BoundText = ""
       DcbAccount_Change
        Me.lbl(9).Enabled = True
        Me.DcboBox.Enabled = True
        Me.lbl(15).Enabled = False
        Me.lbl(16).Enabled = False
        Me.lbl(17).Enabled = False
        Me.DcboBankName.Enabled = False
        Me.TxtChequeNumber.Enabled = False
        Me.DtpChequeDueDate.Enabled = False
        Frame3.Enabled = False
    ElseIf Me.CboPaymentType.ListIndex = 1 Then

        If SystemOptions.ChequeBox = True Then
            TXTBankName.Visible = True
            DCChequeBox.Enabled = True
        Else
            TXTBankName.Visible = False
        End If
       DcbAccount.BoundText = ""
       DcbAccount_Change
        Me.lbl(9).Enabled = False
        Me.DcboBox.Enabled = False
        Me.lbl(15).Enabled = True
        Me.lbl(16).Enabled = True
        Me.lbl(17).Enabled = True
        Me.DcboBankName.Enabled = True
        Me.TxtChequeNumber.Enabled = True
        Me.DtpChequeDueDate.Enabled = True
        Frame3.Enabled = False
    ElseIf Me.CboPaymentType.ListIndex = 2 Then
 
        TXTBankName.Visible = False
 
        Me.lbl(9).Enabled = False
        Me.DcboBox.Enabled = False
        Me.lbl(15).Enabled = True
        Me.lbl(16).Enabled = True
        Me.lbl(17).Enabled = True
        Me.DcboBankName.Enabled = True
        Me.TxtChequeNumber.Enabled = True
        Me.DtpChequeDueDate.Enabled = True
        TXTBankName.Visible = False
        Frame3.Enabled = True
       DcbAccount.BoundText = ""
       DcbAccount_Change
        If SystemOptions.UserInterface = ArabicInterface Then
            lbl(16).Caption = "—ﬁ„ «·ÕÊ«·Â"
            lbl(17).Caption = " «—ÌŒÂ«"
    
        Else
            lbl(16).Caption = "Transfer No"
            lbl(17).Caption = "Date"
        End If
 
    ElseIf Me.CboPaymentType.ListIndex = 3 Then
       DcbAccount.BoundText = ""
       DcbAccount_Change
        TXTBankName.Visible = False
 
        Me.lbl(9).Enabled = False
        Me.DcboBox.Enabled = False
        Me.lbl(15).Enabled = True
        Me.lbl(16).Enabled = True
        Me.lbl(17).Enabled = True
        Me.DcboBankName.Enabled = True
        Me.TxtChequeNumber.Enabled = True
        Me.DtpChequeDueDate.Enabled = True
        TXTBankName.Visible = False
        Frame3.Enabled = True

        If SystemOptions.UserInterface = ArabicInterface Then
            lbl(16).Caption = "—ﬁ„ «·‘Ìﬂ"
            lbl(17).Caption = " «—ÌŒÂ"
    
        Else
            lbl(16).Caption = "Chequ No"
            lbl(17).Caption = "Date"
        End If
 ElseIf Me.CboPaymentType.ListIndex = 4 Then
         Me.lbl(9).Enabled = False
        Me.DcboBox.Enabled = False
        Me.lbl(15).Enabled = False
        Me.lbl(16).Enabled = False
        Me.lbl(17).Enabled = False
        Me.DcboBankName.Enabled = False
        Me.TxtChequeNumber.Enabled = False
        Me.DtpChequeDueDate.Enabled = False
        TxtAccount.Enabled = True
        DcbAccount.Enabled = True
    Else

        Me.lbl(9).Enabled = False
        Me.DcboBox.Enabled = False
        Me.lbl(15).Enabled = False
        Me.lbl(16).Enabled = False
        Me.lbl(17).Enabled = False
        Me.DcboBankName.Enabled = False
        Me.TxtChequeNumber.Enabled = False
        Me.DtpChequeDueDate.Enabled = False
        TxtAccount.Enabled = False
        DcbAccount.Enabled = False
    End If

End Sub

Private Sub CboPayMentType_Click()
    CboPayMentType_Change
End Sub

Private Sub CheckStatusEarnest_Click(Index As Integer)
Select Case Index
Case 0
  If CheckStatusEarnest(0).value = vbChecked Then
  CheckStatusEarnest(1).value = vbUnchecked
  End If
 Case 1
     If CheckStatusEarnest(1).value = vbChecked Then
  CheckStatusEarnest(0).value = vbUnchecked
  End If
  
  Case 3
  If CheckStatusEarnest(3).value = vbChecked Then
  CheckStatusEarnest(1).value = vbUnchecked
  CheckStatusEarnest(0).value = vbUnchecked
  End If

End Select

End Sub

Private Sub ChkTrans_Click()
    Me.lbl(10).Enabled = ChkTrans.value
    Me.lbl(12).Enabled = ChkTrans.value
    Me.CboTrans.Enabled = ChkTrans.value
    Me.TxtTransID.Enabled = ChkTrans.value
    Me.TxtTransSerial.Enabled = ChkTrans.value
    Me.CmdSearchTrans.Enabled = ChkTrans.value
    Me.CmdOpenTrans.Enabled = ChkTrans.value
End Sub

Function sand_numbering() As String
    On Error Resume Next
    Dim start_at As Integer
    Dim end_at As Integer
    Dim auto_sanad_no As String
    Dim NO As Integer
    auto_sanad_no = ""
    departement_name = 1
 
    connection_string = Cn.ConnectionString
    numbering.ConnectionString = connection_string
    numbering.CommandType = adCmdText
    numbering.RecordSource = "select * from sanad_numbering where branch_no=" & my_branch & " and departement='" & departement_name & "' and  sanad_no=2"
    numbering.Refresh

    If numbering.Recordset.RecordCount = 0 Then
        numbering_type = 0
    Else
        numbering_type = numbering.Recordset.Fields!numbering_id
        start_at = numbering.Recordset.Fields!start_at
        end_at = numbering.Recordset.Fields!end_at

    End If

    If numbering_type = 1 Then
        detect_no.ConnectionString = connection_string
        detect_no.CommandType = adCmdText
        detect_no.RecordSource = "select max(NoteSerial) as last_sand_no from  Notes where NoteType=4 and numbering_type=" & numbering_type  ' branch_no=" & branch_no & " and departement='" & departement_name & "' and  type='" & "”‰œ ﬁÌœ" & "' and numbering_type=" & numbering_type
        detect_no.Refresh

        If Not IsNull(detect_no.Recordset.Fields!last_sand_no) Then
 
            If end_at = 0 Then end_at = detect_no.Recordset.Fields!last_sand_no + 1
 
            If detect_no.Recordset.Fields!last_sand_no >= end_at Then
                sand_numbering = "error"
                Exit Function
            End If
        End If

    Else

        If numbering_type = 2 Then
 
            detect_no.ConnectionString = connection_string
            detect_no.CommandType = adCmdText
            detect_no.RecordSource = "select max(NoteSerial) as last_sand_no from  Notes where NoteType=4 and numbering_type=" & numbering_type & "and sanad_year=" & mId(Format$(XPDtbTrans.value, "dd/mm/yyyy"), 7, 4) & "and sanad_month=" & mId(Format$(XPDtbTrans.value, "dd/mm/yyyy"), 4, 2)
            'detect_no.RecordSource = "select max(sanad_no) as last_sand_no from  sandat_ked where  branch_no=" & branch_no & " and departement='" & departement_name & "' and  type='" & "”‰œ ﬁÌœ" & "' and numbering_type=" & numbering_type & "and sanad_year=" & Mid(Format$(Now, "dd/mm/yyyy"), 7, 4) & "and sanad_month=" & Mid(Format$(Now, "dd/mm/yyyy"), 4, 2)
            detect_no.Refresh

            If Not IsNull(detect_no.Recordset.Fields!last_sand_no) Then
                NO = mId(detect_no.Recordset.Fields!last_sand_no, 7, Len(detect_no.Recordset.Fields!last_sand_no) - 6)

                If end_at = 0 Then end_at = NO + 1
                If NO >= end_at Then
                    sand_numbering = "error"
                    Exit Function
                End If
            End If

        Else

            If numbering_type = 3 Then
 
                detect_no.ConnectionString = connection_string
                detect_no.CommandType = adCmdText
                detect_no.RecordSource = "select max(NoteSerial) as last_sand_no from  Notes where NoteType=4 and numbering_type=" & numbering_type & "and sanad_year=" & mId(Format$(XPDtbTrans.value, "dd/mm/yyyy"), 7, 4)
                'detect_no.RecordSource = "select max(sanad_no) as last_sand_no from  sandat_ked where  branch_no=" & branch_no & " and departement='" & departement_name & "'  and  type='" & "”‰œ ﬁÌœ" & "' and numbering_type=" & numbering_type & "and sanad_year=" & Mid(Format$(Now, "dd/mm/yyyy"), 7, 4)
                detect_no.Refresh

                If Not IsNull(detect_no.Recordset.Fields!last_sand_no) Then
                    NO = mId(detect_no.Recordset.Fields!last_sand_no, 5, Len(detect_no.Recordset.Fields!last_sand_no) - 4)

                    If end_at = 0 Then end_at = NO + 1
                    If NO >= end_at Then
                        sand_numbering = "error"
                        Exit Function
                    End If
                End If
 
            End If
 
        End If
    End If

    If detect_no.Recordset.RecordCount = 0 Or IsNull(detect_no.Recordset.Fields!last_sand_no) Then

        If numbering_type = 0 Then
            ' auto_sanad_no = 1
        Else

            If numbering_type = 1 Then
                auto_sanad_no = start_at
            Else
                
                If numbering_type = 2 Then
                    auto_sanad_no = mId(Format$(XPDtbTrans.value, "dd/mm/yyyy"), 7, 4) & mId(Format$(XPDtbTrans.value, "dd/mm/yyyy"), 4, 2) & start_at

                Else

                    If numbering_type = 3 Then
                        auto_sanad_no = mId(Format$(XPDtbTrans.value, "dd/mm/yyyy"), 7, 4) & start_at

                    End If
                End If
            End If
        End If

    Else

        If numbering_type = 0 Then
            'auto_sanad_no = x + 1
        Else

            If numbering_type = 1 Then
                auto_sanad_no = detect_no.Recordset.Fields!last_sand_no + 1
            Else
                
                If numbering_type = 2 Then
                    '  If Mid(detect_no.Recordset.Fields!last_sand_no, 1, 6) <> Mid(Format$(Now, "dd/mm/yyyy"), 7, 4) & Mid(Format$(Now, "dd/mm/yyyy"), 4, 2) Then
                    ' no = 1
                    '  auto_sanad_no = Mid(Format$(Now, "dd/mm/yyyy"), 7, 4) & Mid(Format$(Now, "dd/mm/yyyy"), 4, 2) & "1"
                    '  Else
                    NO = mId(detect_no.Recordset.Fields!last_sand_no, 7, Len(detect_no.Recordset.Fields!last_sand_no) - 6)
                    auto_sanad_no = mId(detect_no.Recordset.Fields!last_sand_no, 1, 6) & (NO + 1)
                    '  End If
                      
                Else

                    If numbering_type = 3 Then
                        '    If Mid(detect_no.Recordset.Fields!last_sand_no, 1, 4) <> Mid(Format$(Now, "dd/mm/yyyy"), 7, 4) Then
                        'no = 1
                        '    auto_sanad_no = Mid(Format$(Now, "dd/mm/yyyy"), 7, 4) & "1"
                        '    Else
                        NO = mId(detect_no.Recordset.Fields!last_sand_no, 5, Len(detect_no.Recordset.Fields!last_sand_no) - 4)
                        auto_sanad_no = mId(detect_no.Recordset.Fields!last_sand_no, 1, 4) & (NO + 1)

                        '    End If

                    End If
                End If
            End If
        End If

    End If

    sand_numbering = auto_sanad_no

    'MsgBox auto_sanad_no

End Function

Public Function newrecord()
Cmd_Click (0)
End Function
Private Sub SetDefaults()

    Dim StrTemp As String
    Dim RsTemp As ADODB.Recordset

        If SystemOptions.SysCashDateTakeType = InvDateFromLocalCompuer Then
               XPDtbTrans.value = Date
               XPDtbTrans.Enabled = True
               Txt_DateHigri.Enabled = True
          ElseIf SystemOptions.SysCashDateTakeType = InvDateFromServerComputer Then
    XPDtbTrans.Enabled = False
     Txt_DateHigri.Enabled = False
        StrTemp = "select Getdate() as ServerDate"
        Set RsTemp = New ADODB.Recordset
        RsTemp.Open StrTemp, Cn, adOpenStatic, adLockReadOnly, adCmdText

                    If Not (RsTemp.BOF Or RsTemp.EOF) Then
                                If Not IsNull(RsTemp("ServerDate").value) Then
                                    XPDtbTrans.value = Format(RsTemp("ServerDate").value, "yyyy/M/d")
                                End If
     
                        'XPDtbBill.Value = IIf(IsNull(RsTemp("ServerDate").Value), Date, (RsTemp("ServerDate").Value))
                    End If

                    RsTemp.Close
                    Set RsTemp = Nothing
          End If

    If Not (rs.BOF Or rs.EOF) Then
        rs.MoveLast
'SetDefaults
              If SystemOptions.SysCashDateTakeType = InvDateFromLastInvDate Then
            XPDtbTrans.value = IIf(IsNull(rs("NoteDate").value), Date, (rs("NoteDate").value))
               End If
    End If

 
End Sub
 Public Sub Cmd_Click(Index As Integer)
Dim i As Integer
    Dim cNoteReport As ClsNotesReports
    Dim Msg As String
'     On Error GoTo ErrTrap

    If SystemOptions.usertype <> UserAdminAll Then
 
        Me.Dcbranch.Enabled = True
        ' XPDtbBill.Enabled = False
    End If

    Select Case Index
Case 10
' C1Elastic1.Visible = False
        Case 0

            If SystemOptions.SysRegisterState = DemoRun Then
                If Not rs Is Nothing Then
                    If Not (rs.BOF Or rs.EOF) Then
                        If rs.RecordCount >= 25 Then
                            Msg = "›Ï «·‰”Œ… «· Ã—Ì»Ì… ·«Ì„ﬂ‰  ”ÃÌ· «ﬂÀ— „‰ 25 ⁄„·Ì… ﬁ»÷ «Ê œ›⁄"
                            MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
                            Exit Sub
                        End If
                    End If
                End If
            End If

            If DoPremis(Do_New, Me.Name, True) = False Then
                Exit Sub
            End If

            clear_all Me
            Dim Dcombos As ClsDataCombos

   Set Dcombos = New ClsDataCombos
   Dcombos.GetIqarUnit -2, 1, Me.DcbUnitNo
        CheckStatusEarnest(0).value = xtpUnchecked
        CheckStatusEarnest(3).value = xtpUnchecked
        CheckStatusEarnest(1).value = xtpUnchecked
            Grid.Clear flexClearScrollable, flexClearEverything
            Grid.rows = 1
          
            GRID1.Clear flexClearScrollable, flexClearEverything
            GRID1.rows = 1
            VSFlexGrid1.Clear flexClearScrollable, flexClearEverything
            VSFlexGrid1.rows = 2
             VSFlexGrid2.Clear flexClearScrollable, flexClearEverything
            VSFlexGrid2.rows = 2
               Grid3.Clear flexClearScrollable, flexClearEverything
            Grid3.rows = 1
          
            Grid4.Clear flexClearScrollable, flexClearEverything
            Grid4.rows = 1
            
            TxtModFlg.Text = "N"
            '       XPTxtID.text = CStr(new_id("Notes", "NoteID", "", True))
            ' Me.TxtNoteSerial.text = CStr(new_id("Notes", "NoteSerial", "", True, "NoteType=4"))
            Me.DCboUserName.BoundText = user_id
            
            Text1.Text = setfoxy
            Option1.value = True
            Me.Dcbranch.BoundText = Current_branch
            Txt_DateHigri.value = ToHijriDate(Date)
CboPayMentType_Change
XPTab301.CurrTab = 0
      XPDtbTrans.SetFocus
     ' Option1.value = True
Option4.value = True
cbointervaltype.ListIndex = 0

SetDefaults
    
   
        Case 1
                   If ChekClodePeriod(XPDtbTrans.value) = True Then
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
              
              If ChekPayeArbon(val(Me.XPTxtID)) = True Then
              If SystemOptions.UserInterface = ArabicInterface Then
                    Msg = "·‰ Ì„ﬂ‰ «·”„«Õ » ⁄œÌ· Â–« «·⁄„·Ì…..!!!"
                    Msg = Msg & CHR(13) & "ÕÌÀ «‰Â« ”Ê› Ì‰ Ã ⁄‰Â« Œÿ« ›Ï   ›Ì ⁄ﬁœ «·«ÌÃ«— ·«‰Â  „ ⁄·ÌÂ« Õ—ﬂ«   "
                    Else
                    Msg = "You Can Not Edit this Process..!!!"
                    Msg = Msg & CHR(13) & "Linked contract  "
                   
                    End If
                    
                    MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
                    Exit Sub
                End If
                
            If SystemOptions.ChequeBox = True And CboPaymentType.ListIndex = 1 Then
         
                If ChequeBoxOperations(val(Me.XPTxtID)) = False Then
                    Msg = "·‰ Ì„ﬂ‰ «·”„«Õ » ⁄œÌ· Â–« «·⁄„·Ì…..!!!"
                    Msg = Msg & CHR(13) & "ÕÌÀ «‰Â« ”Ê› Ì‰ Ã ⁄‰Â« Œÿ« ›Ï   Õ«›Ÿ… «·‘Ìﬂ«  ·«‰Â  „ ⁄·ÌÂ« Õ—ﬂ«  «Ìœ«⁄ «Ê  Õ’Ì· "
                    MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
                    Exit Sub
                End If
    
            End If
    
            TxtModFlg.Text = "E"
       '     Me.DCboUserName.BoundText = user_id
            CuurentLogdata
   VSFlexGrid1.rows = VSFlexGrid1.rows + 1
    VSFlexGrid2.rows = VSFlexGrid2.rows + 1
    If Me.DCboCashType.ListIndex = 8 And (Me.TxtModFlg.Text = "E") Then
      FillGridWithDataContract txtContractNo, val(XPTxtID.Text)
    End If
        Case 2
        
              If ChekClodePeriod(XPDtbTrans.value) = True Then
               If SystemOptions.UserInterface = ArabicInterface Then
                MsgBox "Ì—ÃÏ  €ÌÌ—  «—ÌŒ «·Õ—ﬂ… ·«‰ Â–Â «·› —… „€·ﬁ…"
               Else
               MsgBox "Please Change Date Becouse This is Period is Closed"
              End If
              Exit Sub
              End If
              
XPTab301.CurrTab = 0
            If Trim(Dcbranch.BoundText) = "" Then
                If SystemOptions.UserInterface = EnglishInterface Then
                    Msg = "Specify Branch"
                Else
                    Msg = "Õœœ «·›—⁄ "
                End If

                MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
                Dcbranch.SetFocus
                Sendkeys "{F4}"
                Screen.MousePointer = vbDefault
                Exit Sub
            End If
   my_branch = val(Dcbranch.BoundText)
            my_branch = Me.Dcbranch.BoundText
 If val(XPTxtVal.Text) <= 0 Then
 If SystemOptions.UserInterface = ArabicInterface Then
 MsgBox "·«Ì„ﬂ‰ «·Õ›Ÿ Ì—ÃÏ «œŒ«· ﬁÌ„… «·„ﬁ»Ê÷«  «·’ÕÌÕ…"
 Else
 MsgBox "Please Enter Value"
 End If
' XPTxtVal.SetFocus
 Exit Sub
 End If
            If Option2.value = True And lblsqlstring.Caption = "" Then MsgBox "·«»œ „‰  ÕœÌœ ›Ê« Ì—": Exit Sub
       Dim SUM As Double
  If val(DCboCashType.ListIndex) = 9 Then
    SUM = 0
    If VSFlexGrid1.rows > 1 Then
                ' fg2.Rows = fg2.Rows - 1
     With VSFlexGrid1
       For i = .FixedRows To .rows - 1
       
              If .TextMatrix(i, .ColIndex("empname")) <> "" Then
      
            SUM = SUM + val(.TextMatrix(i, .ColIndex("rate")))
    
       End If
           Next i
           If SUM < 100 Or SUM > 100 Then
        '   MsgBox " ·« Ì„ﬂ‰ «·Õ›Ÿ ÌÃ» «‰ ÌﬂÊ‰ „Ã„Ê⁄ «·‰”» Ì”«ÊÌ 100%"
           MsgBox "  ‰»ÌÂ „Ã„Ê⁄ «·‰”» ·« Ì”«ÊÌ 100%"
        'Exit Sub
        End If
    End With
   End If
       
 End If
 
 Dim Account_Code_dynamic As String
 Dim ComVal As Double
'If Me.DCboCashType.ListIndex = 8 And Rd(1).value = False And SystemOptions.NoCreatJLInRentContract = False Then
If Me.DCboCashType.ListIndex = 8 And Rd(1).value = False Then
''////////
'If SystemOptions.DueComm = True Then
 If True = True Then
      With Me.Grid3
 For i = 1 To .rows - 1
        If .cell(flexcpChecked, i, .ColIndex("Select")) = flexChecked And val(.TextMatrix(i, .ColIndex("CommissionsPayed"))) <> 0 Then
ComVal = ComVal + val(.TextMatrix(i, .ColIndex("CommissionsPayed")))
         End If
       Next i
 End With
 If ComVal > 0 Then
          Account_Code_dynamic = get_account_code_branch(153, my_branch)
          If Account_Code_dynamic = "NO branch" Then
          MsgBox "·„ Ì „ «‰‘«¡ «·›—⁄", vbCritical
            GoTo ErrTrap
            Else
           If Account_Code_dynamic = "NO account" Then
              MsgBox "·„ Ì „  ÕœÌœ Õ”«»  «” Õﬁ«ﬁ «·”⁄Ì", vbCritical
             Exit Sub
    
           End If
           End If
                   Account_Code_dynamic = get_account_code_branch(83, my_branch)
          If Account_Code_dynamic = "NO branch" Then
          MsgBox "·„ Ì „ «‰‘«¡ «·›—⁄", vbCritical
            GoTo ErrTrap
            Else
           If Account_Code_dynamic = "NO account" Then
              MsgBox "·„ Ì „  ÕœÌœ Õ”«»  «Ì—«œ «·”⁄Ì", vbCritical
             Exit Sub
    
           End If
           End If
           
 End If

 End If
End If
If SystemOptions.CreateJLEmpCommissions = True Then
             Account_Code_dynamic = get_account_code_branch(161, my_branch)
          If Account_Code_dynamic = "NO branch" Then
          MsgBox "·„ Ì „ «‰‘«¡ «·›—⁄", vbCritical
            GoTo ErrTrap
            Else
           If Account_Code_dynamic = "NO account" Then
              MsgBox "·„ Ì „  ÕœÌœ Õ”«»   ⁄„Ê·«  «·„‰«œÌ»", vbCritical
             Exit Sub
    
           End If
           End If
End If
         Dim IarType As Integer
        
            IarType = AqarCommisionType(val(DcbIqara.BoundText))
If SystemOptions.NoCreatJLInRentContract = True And DCboCashType.ListIndex = 8 Then
 If IarType = 1 Then
' If val(TxtVATValue.Text) > 0 Then
'Account_Code_dynamic = ModAccounts.GetMyAccountCode("TblCustemers", "CusID", val(Txtownerid), "Account_VAT")
'      If Account_Code_dynamic = "" Then
'              MsgBox "·„ Ì „  ÕœÌœ Õ”«»   «·ﬁÌ„… «·„÷«›… ·Â–« «·„«·ﬂ", vbCritical
'             Exit Sub
'    End If
End If
  With Me.Grid3
 For i = 1 To .rows - 1
        If .cell(flexcpChecked, i, .ColIndex("Select")) = flexChecked And val(.TextMatrix(i, .ColIndex("RentValuePayed"))) <> 0 Then
ComVal = ComVal + val(.TextMatrix(i, .ColIndex("RentValuePayed")))
         End If
       Next i
 End With
 
 
 ComVal = ComVal * val(TxtKickbacks) / 100
 If ComVal > 0 Then
 PercentgValueAddedAccount_Transec XPDtbTrans.value, 8, 1, Account_Code_dynamic
'GetValueAddedAccount XPDtbTrans.value, , Account_Code_dynamic, 1, 21
      If Account_Code_dynamic = "" Then
              MsgBox "·„ Ì „  ÕœÌœ Õ”«»   «·ﬁÌ„… «·„÷«›…  ·⁄ﬁœ «·«ÌÃ«—", vbCritical
             Exit Sub
    End If

End If

 ElseIf IarType = 0 Then
     With Me.Grid3
 For i = 1 To .rows - 1
        If .cell(flexcpChecked, i, .ColIndex("Select")) = flexChecked And val(.TextMatrix(i, .ColIndex("RentValuePayed"))) <> 0 Then
ComVal = ComVal + val(.TextMatrix(i, .ColIndex("RentValuePayed")))
         End If
       Next i
 End With

 If ComVal > 0 Then
             Account_Code_dynamic = get_account_code_branch(86, val(Dcbranch.BoundText))
        If Account_Code_dynamic = "NO branch" Then
          MsgBox "·„ Ì „ «‰‘«¡ «·›—⁄", vbCritical
            GoTo ErrTrap
            Else
           If Account_Code_dynamic = "NO account" Then
              MsgBox "·„ Ì „  ÕœÌœ Õ”«»     «Ì—«œ«  «·«ÌÃ«—« ", vbCritical
             Exit Sub
    
           End If
        End If
 End If

 ''///////////
 ComVal = 0
      With Me.Grid3
 For i = 1 To .rows - 1
        If .cell(flexcpChecked, i, .ColIndex("Select")) = flexChecked And val(.TextMatrix(i, .ColIndex("WaterPayed"))) <> 0 Then
ComVal = ComVal + val(.TextMatrix(i, .ColIndex("WaterPayed")))
         End If
       Next i
 End With

 If ComVal > 0 Then
             Account_Code_dynamic = get_account_code_branch(83, val(Dcbranch.BoundText))
        If Account_Code_dynamic = "NO branch" Then
          MsgBox "·„ Ì „ «‰‘«¡ «·›—⁄", vbCritical
            GoTo ErrTrap
            Else
           If Account_Code_dynamic = "NO account" Then
              MsgBox "·„ Ì „  ÕœÌœ Õ”«»     «Ì—«œ«  «·„Ì«Â", vbCritical
             Exit Sub
    
           End If
        End If
 End If

 ''/////////////////
  ComVal = 0
       With Me.Grid3
 For i = 1 To .rows - 1
        If .cell(flexcpChecked, i, .ColIndex("Select")) = flexChecked And val(.TextMatrix(i, .ColIndex("ElectricPayed"))) <> 0 Then
ComVal = ComVal + val(.TextMatrix(i, .ColIndex("ElectricPayed")))
         End If
       Next i
 End With

 If ComVal > 0 Then
 
            Account_Code_dynamic = get_account_code_branch(84, val(Dcbranch.BoundText))
        If Account_Code_dynamic = "NO branch" Then
          MsgBox "·„ Ì „ «‰‘«¡ «·›—⁄", vbCritical
            GoTo ErrTrap
            Else
           If Account_Code_dynamic = "NO account" Then
              MsgBox "·„ Ì „  ÕœÌœ Õ”«»     «Ì—«œ«  «·ﬂÂ—»«¡", vbCritical
             Exit Sub
    
           End If
        End If
 End If

 
  ''/////////////////
   ComVal = 0
       With Me.Grid3
 For i = 1 To .rows - 1
        If .cell(flexcpChecked, i, .ColIndex("Select")) = flexChecked And val(.TextMatrix(i, .ColIndex("TelandNetPayed"))) <> 0 Then
ComVal = ComVal + val(.TextMatrix(i, .ColIndex("TelandNetPayed")))
         End If
       Next i
 End With

 If ComVal > 0 Then
           Account_Code_dynamic = get_account_code_branch(85, val(Dcbranch.BoundText))
        If Account_Code_dynamic = "NO branch" Then
          MsgBox "·„ Ì „ «‰‘«¡ «·›—⁄", vbCritical
            GoTo ErrTrap
            Else
           If Account_Code_dynamic = "NO account" Then
              MsgBox "·„ Ì „  ÕœÌœ Õ”«»     «Ì—«œ«  «·Œœ„« ", vbCritical
             Exit Sub
    
           End If
        End If
 End If

 ''//////////
  ComVal = 0
        With Me.Grid3
 For i = 1 To .rows - 1
        If .cell(flexcpChecked, i, .ColIndex("Select")) = flexChecked And val(.TextMatrix(i, .ColIndex("CommissionsPayed"))) <> 0 Then
ComVal = ComVal + val(.TextMatrix(i, .ColIndex("CommissionsPayed")))
         End If
       Next i
 End With

 If ComVal > 0 Then
 
           Account_Code_dynamic = get_account_code_branch(81, val(Dcbranch.BoundText))
        If Account_Code_dynamic = "NO branch" Then
          MsgBox "·„ Ì „ «‰‘«¡ «·›—⁄", vbCritical
            GoTo ErrTrap
            Else
           If Account_Code_dynamic = "NO account" Then
              MsgBox "·„ Ì „  ÕœÌœ Õ”«»     «Ì—«œ«  «·”⁄Ì Ê«·⁄„Ê·« ", vbCritical
             Exit Sub
    
           End If
        End If
 End If

 
 End If
 'End If
 If val(DCboCashType.ListIndex) = 8 Then
    ComVal = val(TxtVATValue.Text)

 If ComVal > 0 Then
 
         ' Account_Code_dynamic = get_account_code_branch(145, val(dcBranch.BoundText))
        'GetValueAddedAccount XPDtbTrans.value, , Account_Code_dynamic, 1, 21
         PercentgValueAddedAccount_Transec XPDtbTrans.value, 8, 1, Account_Code_dynamic
        If Account_Code_dynamic = "NO branch" Then
          MsgBox "·„ Ì „ «‰‘«¡ «·›—⁄", vbCritical
            GoTo ErrTrap
            Else
           If Account_Code_dynamic = "NO account" Then
              MsgBox "·„ Ì „  ÕœÌœ Õ”«» «·ﬁÌ„… «·„÷«›… „»Ì⁄«  ", vbCritical
             Exit Sub
    
           End If
        End If
 End If
 End If
  If val(DCboCashType.ListIndex) = 8 Then
    SUM = 0
    If VSFlexGrid2.rows > 1 Then
                ' fg2.Rows = fg2.Rows - 1
     With VSFlexGrid2
       For i = .FixedRows To .rows - 1
       
              If .TextMatrix(i, .ColIndex("empname")) <> "" Then
      
            SUM = SUM + val(.TextMatrix(i, .ColIndex("rate")))
    
       End If
           Next i
           If SUM < 100 Or SUM > 100 Then
         '  MsgBox " ·« Ì„ﬂ‰ «·Õ›Ÿ ÌÃ» «‰ ÌﬂÊ‰ „Ã„Ê⁄ «·‰”» Ì”«ÊÌ 100%"
        'Exit Sub
        End If
    End With
   End If
    End If
            'TxtNoteSerial.text = Notes_coding(Val(my_branch), XPDtbTrans.value)
            SaveData
        
        Case 3
            Undo

        Case 4
                           If ChekClodePeriod(XPDtbTrans.value) = True Then
               If SystemOptions.UserInterface = ArabicInterface Then
                MsgBox "Ì—ÃÏ  €ÌÌ—  «—ÌŒ «·Õ—ﬂ… ·«‰ Â–Â «·› —… „€·ﬁ…"
               Else
               MsgBox "Please Change Date Becouse This is Period is Closed"
              End If
              Exit Sub
              End If
            If ChekPayeArbon(val(Me.XPTxtID)) = True Then
                  If SystemOptions.UserInterface = ArabicInterface Then
                    Msg = "·‰ Ì„ﬂ‰ «·”„«Õ »Õ–› Â–« «·⁄„·Ì…..!!!"
                    Msg = Msg & CHR(13) & "ÕÌÀ «‰Â« ”Ê› Ì‰ Ã ⁄‰Â« Œÿ« ›Ï   ›Ì ⁄ﬁœ «·«ÌÃ«— ·«‰Â  „ ⁄·ÌÂ« Õ—ﬂ«   "
                    Else
                    Msg = "You Can Not Delete this Process..!!!"
                    Msg = Msg & CHR(13) & "Linked contract  "
                   
                    End If
                    
                    MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
                    Exit Sub
                End If
                

            If DoPremis(Do_Delete, Me.Name, True) = False Then
                Exit Sub
            End If

            If SystemOptions.ChequeBox = True And CboPaymentType.ListIndex = 1 Then
         
                If ChequeBoxOperations(val(Me.XPTxtID)) = False Then
                    Msg = "·‰ Ì„ﬂ‰ «·”„«Õ »Õ–› Â–« «·⁄„·Ì…..!!!"
                    Msg = Msg & CHR(13) & "ÕÌÀ «‰Â« ”Ê› Ì‰ Ã ⁄‰Â« Œÿ« ›Ï   Õ«›Ÿ… «·‘Ìﬂ«  ·«‰Â  „ ⁄·ÌÂ« Õ—ﬂ«  «Ìœ«⁄ «Ê  Õ’Ì· "
                    MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
                    Exit Sub
                End If
            End If

            Del_Trans

        Case 5

            If DoPremis(Do_Search, Me.Name, True) = False Then
                Exit Sub
            End If

           Load FrmNotesSearch
            FrmNotesSearch.SearchType = 6
            FrmNotesSearch.show vbModal

        Case 6
            Unload Me

        Case 7

            If DoPremis(Do_Print, Me.Name, True) = False Then
                Exit Sub
            End If

            ' If Val(Me.XPTxtID.text) <> 0 Then
            '     Set cNoteReport = New ClsNotesReports
            '     cNoteReport.PrintReceipt Val(Me.XPTxtID.text), WindowTarget
            '     Set cNoteReport = Nothing
            ' End If
            If TxtNoteSerial <> "" Then
            print_report
            SendMessage (2)
                'print_report TxtNoteSerial, Me.TxtNoteSerial1.text, TXTBankName.text, CboPayMentType.text, DcboBox.text, TxtCustCode.text
            End If

        Case 8

            'ViewDataList
        Case 9
    
            If DoPremis(Do_Print, Me.Name, True) = False Then
                Exit Sub
            End If

         '   ShowGL_cc Me.TxtNoteSerial.Text, , 200
            ShowGL_cc TxtNoteSerial.Text, , , XPTxtID.Text
            
            Case 13
            RemoveGridRow
             Case 14
            RemoveGridRow1
            
    End Select

    Exit Sub
ErrTrap:
End Sub
Private Sub RemoveGridRow1()

    With Me.VSFlexGrid1

        If .Row <= 0 Then Exit Sub
        .RemoveItem .Row
    End With

    ReLineGrid
End Sub
Private Sub RemoveGridRow()

    With Me.VSFlexGrid2

        If .Row <= 0 Then Exit Sub
        .RemoveItem .Row
    End With

    ReLineGrid
End Sub
'Public Function print_report(Optional NoteSerial As String, Optional NoteSerial1 As String, Optional BankName As String, Optional PaymentType As String, Optional Box As String, Optional Custcode As String)
'
'    Dim MySQL As String
'    Dim RsData As New ADODB.Recordset
'    Dim xApp As New CRAXDRT.Application
'    Dim xReport As CRAXDRT.Report
'    Dim CViewer As ClsReportViewer
'    Dim StrReportTitle As String
'    Dim StrFileName As String
'    Dim Msg As String
'
'    MySQL = "Select * From payment_voucher  where noteserial='" & NoteSerial & "'"
'
'    If SystemOptions.UserInterface = ArabicInterface Then
'        StrFileName = App.path & "\Reports\" & "Payment_voucher.rpt"
'    Else
'        StrFileName = App.path & "\Reports\" & "Payment_voucherE.rpt"
'    End If
'
'    If Dir(StrFileName) = "" Then
'        'GetMsgs 139, vbExclamation
'        Screen.MousePointer = vbDefault
'        Exit Function
'    End If
'
'    Set RsData = New ADODB.Recordset
'    RsData.Open MySQL, Cn, adOpenStatic, adLockReadOnly, adCmdText
'
'    If RsData.BOF Or RsData.EOF Then
'        'GetMsgs 138, vbExclamation
'        Msg = "·« ÊÃœ »Ì«‰«  ··⁄—÷"
'        MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
'        RsData.Close
'        Set RsData = Nothing
'        Screen.MousePointer = vbDefault
'        Exit Function
'    End If
'
'    Screen.MousePointer = vbArrowHourglass
'    Set xReport = xApp.OpenReport(StrFileName)
'    xReport.Database.SetDataSource RsData

 '   Dim cCompanyInfo As New ClsCompanyInfo
'
'    If SystemOptions.UserInterface = ArabicInterface Then
'        xReport.ParameterFields(1).AddCurrentValue cCompanyInfo.ArabCompanyName 'RPTCompany_Name_Arabic
'        xReport.ParameterFields(5).AddCurrentValue DcboCreditSide.text
'
'        StrReportTitle = "" '& StrAccountName
'
'    Else
'
'        xReport.ParameterFields(1).AddCurrentValue cCompanyInfo.ArabCompanyName ' RPTCompany_Name_Eng
'
'        xReport.ParameterFields(4).AddCurrentValue get_branch_name(val(my_branch))
'        xReport.ParameterFields(5).AddCurrentValue DcboCreditSide.text
'        StrReportTitle = ""
'
'    End If
'
'    xReport.ParameterFields(3).AddCurrentValue user_name
'    '
'    xReport.ParameterFields(6).AddCurrentValue NoteSerial1
'
'    xReport.ParameterFields(7).AddCurrentValue BankName
'    xReport.ParameterFields(8).AddCurrentValue PaymentType
'    xReport.ParameterFields(9).AddCurrentValue Box
'    xReport.ParameterFields(10).AddCurrentValue Custcode
'    xReport.reporttitle = StrReportTitle
'    xReport.EnableParameterPrompting = False
'    xReport.ApplicationName = App.title
'    xReport.ReportAuthor = App.title
'    Set CViewer = New ClsReportViewer
'    CViewer.FireReport xReport, WindowTarget, "", , , , StrFileName
'
'    RsData.Close
'    Set RsData = Nothing
'    Screen.MousePointer = vbDefault
'
'End Function
 
 Function print_report(Optional NoteSerial As String)
    Dim MySQL As String
    Dim RsData As New ADODB.Recordset
    Dim xApp As New CRAXDRT.Application
    Dim xReport As CRAXDRT.Report
    Dim CViewer As ClsReportViewer
    Dim StrReportTitle As String
    Dim StrFileName As String
    Dim Msg As String
  

'newww
MySQL = " SELECT  TblContract.NewNO , dbo.Notes.Note_Value2,Accredit =  case IsNull(TblContract.Accredit,0) When 0 Then '€Ì— „ÊÀﬁ' else '„ÊÀﬁ' end , dbo.Notes.vat,    dbo.Notes.NoteID, dbo.Notes.NoteDate, dbo.Notes.NoteType, dbo.Notes.NoteSerial, dbo.Notes.NoteSerial1, dbo.Notes.Note_Value, dbo.Notes.NoteDateH,"
MySQL = MySQL & "                       dbo.Notes.ContractNo, dbo.Notes.ContNo, dbo.Notes.commission, dbo.Notes.rent, dbo.Notes.Water, dbo.Notes.FilterID, dbo.Notes.FIlterTotal, dbo.Notes.Instrunce,"
MySQL = MySQL & "                       dbo.Notes.comX, dbo.Notes.ComY, dbo.Notes.CommissionOut, dbo.Notes.NoteOrBonID, dbo.Notes.comXold, dbo.Notes.ComYold, dbo.Notes.NoteOrBonValue,"
MySQL = MySQL & "                       dbo.Notes.NoteOrBonSereal, dbo.Notes.Telephone, dbo.Notes.CashingType, dbo.Notes.CusID, "

If DCboCashType.ListIndex = 7 Then
        MySQL = MySQL & "                       ACCOUNTS.Account_Name CusName,Accounts.Account_NameEng CusNamee,"
    Else
        MySQL = MySQL & "                           dbo.TblCustemers.CusName, dbo.TblCustemers.CusNamee,"


End If
MySQL = MySQL & "                       dbo.TblCustemers.Cus_Phone, dbo.TblCustemers.Cus_mobile, dbo.Notes.renterName, dbo.Notes.NoteCashingType, dbo.Notes.BankName, dbo.Notes.DueDate,"
MySQL = MySQL & "                       dbo.Notes.ChqueNum, dbo.Notes.Remark, dbo.Notes.Remark2, dbo.Notes.ToPriodDateH, dbo.Notes.FrmPriodDateH, dbo.Notes.ToPriodDate, dbo.Notes.FrmPriodDate,"
MySQL = MySQL & "                       dbo.Notes.branch_no, dbo.TblBranchesData.branch_name, dbo.TblBranchesData.branch_namee, dbo.TblAqarDetai.Id, dbo.TblAqarDetai.unitno,"
MySQL = MySQL & "                       dbo.TblAqarDetai.unittype, dbo.TblAqarDetai.Aqarid, TblAqar_1.aqarname, TblAkarUnit_2.name, TblAkarUnit_2.namee, dbo.Notes.akarid,"
                      MySQL = MySQL & " TblAqar_1.aqarname AS aqarname2, dbo.Notes.unittype AS unittype2, TblAkarUnit_1.name AS name2, TblAkarUnit_1.namee AS namee2, dbo.Notes.Electricity,"
MySQL = MySQL & "                       dbo.Notes.BankID, dbo.BanksData.BankNamee, dbo.BanksData.BankName AS BankName2,dbo.Notes.Servce,"
', dbo.TblNotesSales.rate, dbo.TblNotesSales.valu,"
'MySQL = MySQL & "                       dbo.TblNotesSales.Type, dbo.TblNotesSales.EmpID, dbo.TblEmployee.Emp_Name, dbo.TblEmployee.Fullcode, dbo.TblEmployee.Emp_Namee, dbo.Notes.Servce,"
 MySQL = MySQL & "                      dbo.Notes.RemaiValue, dbo.ContracttBillInstallmentsDone.WaterPayed, dbo.ContracttBillInstallmentsDone.RentValuePayed,"
MySQL = MySQL & "                       dbo.ContracttBillInstallmentsDone.CommissionsPayed, dbo.ContracttBillInstallmentsDone.InsurancePayed, dbo.ContracttBillInstallmentsDone.ElectricPayed,"
MySQL = MySQL & "                       dbo.ContracttBillInstallmentsDone.TelandNetPayed, dbo.ContracttBillInstallmentsDone.RecordDate, dbo.ContracttBillInstallmentsDone.RecordDateH,"
MySQL = MySQL & "                       dbo.ContracttBillInstallmentsDone.total, dbo.ContracttBillInstallmentsDone.[Value], dbo.ContracttBillInstallmentsDone.InstallNo,"
MySQL = MySQL & "                       dbo.ContracttBillInstallmentsDone.VATPayed, dbo.ContracttBillInstallmentsDone.VATValue, dbo.ContracttBillInstallmentsDone.ActVAT,"
MySQL = MySQL & "                       dbo.ContracttBillInstallmentsDone.Commisionvalue , dbo.ContracttBillInstallmentsDone.OldValuePayed, dbo.ContracttBillInstallmentsDone.PaymentType"
MySQL = MySQL & " FROM         dbo.ContracttBillInstallmentsDone RIGHT OUTER JOIN"
MySQL = MySQL & "                       dbo.Notes ON dbo.ContracttBillInstallmentsDone.NoteID = dbo.Notes.NoteID LEFT OUTER JOIN"

'MySQL = MySQL & "                       dbo.TblNotesSales LEFT OUTER JOIN"
'MySQL = MySQL & "                       dbo.TblEmployee ON dbo.TblNotesSales.EmpID = dbo.TblEmployee.Emp_ID ON dbo.Notes.NoteID = dbo.TblNotesSales.NoteID LEFT OUTER JOIN"
MySQL = MySQL & "                       dbo.BanksData ON dbo.Notes.BankID = dbo.BanksData.BankID LEFT OUTER JOIN"
MySQL = MySQL & "                       dbo.TblAkarUnit TblAkarUnit_1 ON dbo.Notes.unittype = TblAkarUnit_1.id LEFT OUTER JOIN"
MySQL = MySQL & "                       dbo.TblAqar TblAqar_1 ON dbo.Notes.akarid = TblAqar_1.Aqarid LEFT OUTER JOIN"
MySQL = MySQL & "                       dbo.TblAqarDetai LEFT OUTER JOIN"
MySQL = MySQL & "                       dbo.TblAkarUnit TblAkarUnit_2 ON dbo.TblAqarDetai.unittype = TblAkarUnit_2.id LEFT OUTER JOIN"
MySQL = MySQL & "                       dbo.TblAqar TblAqar_2 ON dbo.TblAqarDetai.Aqarid = TblAqar_2.Aqarid ON dbo.Notes.UnitNo = dbo.TblAqarDetai.Id LEFT OUTER JOIN"
MySQL = MySQL & "                       dbo.TblBranchesData ON dbo.Notes.branch_no = dbo.TblBranchesData.branch_id LEFT OUTER JOIN"

MySQL = MySQL & "                       dbo.TblCustemers ON dbo.Notes.CusID = dbo.TblCustemers.CusID"


MySQL = MySQL & "                       LEFT OUTER JOIN"

MySQL = MySQL & "                       dbo.TblContract ON dbo.Notes.ContNo = dbo.TblContract.ContNo"


'

If DCboCashType.ListIndex = 7 Then
    MySQL = MySQL & "    LEFT OUTER JOIN ACCOUNTS  ON notes.AccountsCode = ACCOUNTS.Account_Code"
End If
'Where (dbo.Notes.NoteID = 4441)
MySQL = MySQL & " Where (dbo.Notes.NoteID =" & val(XPTxtID.Text) & ")"

 If SystemOptions.UserInterface = ArabicInterface Then
          StrFileName = App.path & "\REPORTS\REPORTS NEW\" & "Expenses_order10.rpt"
     Else
        StrFileName = App.path & "\REPORTS\REPORTS NEW\" & "Expenses_order10.rpt"
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
        xReport.ParameterFields(4).AddCurrentValue WriteNo(Format(val(XPTxtVal.Text) + val(TxtVATValue.Text), "0.00"), 0, True, ".")
     xReport.ParameterFields(5).AddCurrentValue val(lblremain.Caption)
     
    ' xReport.ParameterFields(6).AddCurrentValue CStr(val(XPTxtVal.Text) + val(TxtVATValue.Text))
    
    xReport.ParameterFields(6).AddCurrentValue "”‰œ " & CboPaymentType.Text
    
     
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


Private Sub ViewDataList()
    Dim FrmView As FrmViewList
    Dim FG As VSFlex8UCtl.VSFlexGrid
    Dim StrSQL As String
    Dim rs2 As ADODB.Recordset
    Dim StrComboList As String
    Dim GrdBack As ClsBackGroundPic
    'Dim cProgress As ClsProgress
    Dim BolFrmLoaded As Boolean
    Set FrmView = New FrmViewList
    Set FG = FrmView.vsfGroup1.VSFlexGrid

    With FG
        .Cols = 18
        .RowHeightMin = 320
        .ExplorerBar = flexExSortShowAndMove
        .TextMatrix(0, 0) = "—ﬁ„ «·⁄„·Ì…"
        .ColKey(0) = "NoteID"
        .TextMatrix(0, 1) = "ﬂÊœ «·⁄„·Ì…"
        .ColKey(1) = "NoteSerial"
        .TextMatrix(0, 2) = "«· «—ÌŒ"
        .ColKey(2) = "NoteDate"
        .TextMatrix(0, 3) = " ‰Ê⁄ «·„ﬁ»Ê÷« "
        .ColKey(3) = "Name"
        .TextMatrix(0, 4) = "ﬁÌ„… «·„ﬁ»Ê÷« "
        .ColKey(4) = "Note_Value"
        .ColFormat(.ColIndex("Note_Value")) = "#,###.##"
        .TextMatrix(0, 5) = "«”„ «·Œ“‰…"
        .ColKey(5) = "BoxName"
        .TextMatrix(0, 6) = "„·«ÕŸ« "
        .ColKey(6) = "Remark"
        .TextMatrix(0, 7) = "Õ—— »Ê«”ÿ…"
        .ColKey(7) = "UserName"
    
        StrSQL = "SELECT NoteID, NoteSerial, NoteDate, Name, Note_Value, BoxName," & "Remark, UserName From ExpensesReport"
        StrSQL = StrSQL + " Order By NoteID"
        Set rs2 = New ADODB.Recordset
        rs2.Open StrSQL, Cn, adOpenKeyset, adLockReadOnly, adCmdText
        'Â‰« Ìﬂ » ﬂÊœ ·⁄„· „⁄œ·  Õ„Ì· «·»Ì«‰« 
        '------------------------------------
        '
        '
        '
        '
    
        '------------------------------------
        Set .DataSource = rs2
        .TextMatrix(0, 0) = "—ﬁ„ «·⁄„·Ì…"
        .ColKey(0) = "NoteID"
        .TextMatrix(0, 1) = "ﬂÊœ «·⁄„·Ì…"
        .ColKey(1) = "NoteSerial"
        .TextMatrix(0, 2) = "«· «—ÌŒ"
        .ColKey(2) = "NoteDate"
        .TextMatrix(0, 3) = "‰Ê⁄ «·„’—Ê›« "
        .ColKey(3) = "Name"
        .TextMatrix(0, 4) = "ﬁÌ„… «·„’—Ê›« "
        .ColKey(4) = "Note_Value"
        .ColFormat(.ColIndex("Note_Value")) = "#,###.##"
        .TextMatrix(0, 5) = "«”„ «·Œ“‰…"
        .ColKey(5) = "BoxName"
        .TextMatrix(0, 6) = "„·«ÕŸ« "
        .ColKey(6) = "Remark"
        .TextMatrix(0, 7) = "Õ—— »Ê«”ÿ…"
        .ColKey(7) = "UserName"
    
        'rs2.Close
        'Set rs2 = Nothing
        .AutoSize 0, .Cols - 1, False
    End With

    Set GrdBack = New ClsBackGroundPic
    FrmView.vsfGroup1.VSFlexGrid.WallPaper = GrdBack.Picture
    FrmView.vsfGroup1.SetRTL = True
    FrmView.vsfGroup1.TotalOnColKey = "Note_Value"
    FrmView.vsfGroup1.sql = StrSQL
    FrmView.vsfGroup1.ShowTreeGroups = True
    FrmView.vsfGroup1.update
    FrmView.SetDblClickRetrun Me, "NoteID"
    FrmView.Caption = "⁄—÷ ‘Ã—Ï ÃœÊ·Ï ·»Ì«‰«  «·„’—Ê›« "
    FrmView.show
End Sub

Private Sub CmdAttach_Click()
    On Error Resume Next
          If DoPremis(Do_Attach, Me.Name, True) = False Then
                Exit Sub
            End If
            
ShowAttachments TxtNoteSerial1, "0612201408"

End Sub

Private Sub CmdHelp_Click()
    SystemOptions.SysHelp.HHTopicID = Me.HelpContextID
    SystemOptions.SysHelp.HHDisplayTopicID Me.hWnd
End Sub

Private Sub CmdRemove_Click()
    Dim X As Integer

    If SystemOptions.UserInterface = EnglishInterface Then
        X = MsgBox("Confirm Delete This line", vbCritical + vbYesNo)
    Else
        X = MsgBox(" √ﬂÌœ Õ–› «·„Õœœ ", vbCritical + vbYesNo)
    End If

    Dim sql As String

    If X = vbNo Then Exit Sub
    sql = "delete from ProjectBillBuy where id=" & val(GRID1.TextMatrix(GRID1.Row, GRID1.ColIndex("id")))
    Cn.Execute sql

    If GRID1.rows > 1 Then
        If GRID1.rows = 2 Then
            Me.GRID1.Clear flexClearScrollable, flexClearEverything
        Else

            If Me.GRID1.rows > 1 Then
                If Me.GRID1.Row <> Me.GRID1.FixedRows - 1 Then
                    Me.GRID1.RemoveItem (Me.GRID1.Row)
                End If
            End If
        End If
    End If

    If DCboCashType.ListIndex = 5 Then
        FillGridWithData val(Me.DBCboClientName.BoundText), TxtNoteSerial.Text
    End If
  
End Sub

Private Sub CmdSearchTrans_Click()
    Dim Msg As String

    If Me.CboTrans.ListIndex = -1 Then
        Msg = "ÌÃ» ≈Œ Ì«— ‰Ê⁄ «·Õ—ﬂ… «·„—«œ «·»ÕÀ ⁄‰Â«..."
        MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
        CboTrans.SetFocus
        Sendkeys "{F4}"
        Exit Sub
    End If

    If Me.CboTrans.ListIndex = 0 Then
        ' ›« Ê—… „»Ì⁄« 
        Load FrmBuySearch
        FrmBuySearch.DealingForm = InvoiceTransaction
        Set FrmBuySearch.ExtraRetrunObject = Me.TxtTransID
        FrmBuySearch.CboPaymentType.ListIndex = 1
        FrmBuySearch.CboPaymentType.Enabled = False
        FrmBuySearch.Caption = "«·»ÕÀ ⁄‰ ⁄„·Ì… »Ì⁄"
        FrmBuySearch.DCboClientsName.BoundText = Me.DBCboClientName.BoundText
        FrmBuySearch.show
    ElseIf Me.CboTrans.ListIndex = 1 Then
        '›« Ê—… „— Ã⁄ „‘ —Ì« 
        Load FrmBuySearch
        FrmBuySearch.DealingForm = Returntransaction
        Set FrmBuySearch.ExtraRetrunObject = Me.TxtTransID
        FrmBuySearch.CboPaymentType.ListIndex = 1
        FrmBuySearch.CboPaymentType.Enabled = False
        FrmBuySearch.Caption = "«·»ÕÀ ⁄‰ „— Ã⁄ «·„‘ —Ì« "
        FrmBuySearch.DCboClientsName.BoundText = Me.DBCboClientName.BoundText
        FrmBuySearch.show vbModal
    ElseIf Me.CboTrans.ListIndex = 2 Then
        '›« Ê—… ’Ì«‰…
        Load FrmMaintanenceSearch
        Set FrmMaintanenceSearch.ExtraRetrunObject = Me.TxtTransID
        FrmMaintanenceSearch.CboPaymentType.ListIndex = 1
        FrmMaintanenceSearch.SearchType = 4
        FrmMaintanenceSearch.CboPaymentType.Enabled = False
        FrmMaintanenceSearch.show vbModal
    End If

End Sub




Private Sub CMDSENDSMS_Click()
'0 manual
'1 save
'2 Print

SendMessage (0) '0 manual
'1 save
'2 Print

End Sub
Function SendMessage(currentOpt As Integer)
            Dim subject As String
            Dim Msg As String
            Dim msgstatus As Boolean
           Dim CompanyName As String
           Dim cOptions As ClsCompanyInfo
           Set cOptions = New ClsCompanyInfo
           Dim companyphone As String
           Dim Opt As Integer
            Dim CurrentMessage As String
            Dim t As String
    CurrentMessage = ComposMessage(Me.Name, 0, "", "", Opt)
  If Opt = currentOpt Then
  
       
 Msg = "  „ ”œ«œ „»·€  " & XPTxtVal & " ”‰œ ﬁ»÷" & TxtNoteSerial1 & " ··ÊÕœ… " & DcbUnitNo.Text & "  ··⁄ﬁ«—   " & DcbIqara.Text
t = sendMessageM("user", "password", Msg, "", GetCustomerNumber(val(DBCboClientName.BoundText)))
DoEvents


 Msg = "  „ ”œ«œ „»·€  " & XPTxtVal & " ”‰œ ﬁ»÷" & TxtNoteSerial1 & " ··ÊÃœ… " & DcbUnitNo.Text & "  ··⁄ﬁ«—   " & DcbIqara.Text
t = sendMessageM("user", "password", Msg, "", GetCustomerNumber(Txtownerid))
DoEvents

MsgBox " „ «·«—”«·"
     
     
     End If
 
End Function


Private Sub ComResid_Click(Index As Integer)
ClculteVAT
End Sub
Sub ClculteVAT()
Dim commisiontype As Integer
commisiontype = AqarCommisionType(val(DcbIqara.BoundText))
If ComResid(1).value = True And val(DCboCashType.ListIndex) = 9 And commisiontype = 1 = 0 Then
TxtVATValue.Text = 0 ' val(XPTxtVal.Text) * 5 / 100

ElseIf val(DCboCashType.ListIndex) <> 8 Then
TxtVATValue.Text = 0
End If
'salimhere
' TxtVATValue = netVatPayed
 
End Sub
 Private Sub DBCboClientName_Change()
    Dim pstate As Integer
    TxtCustCode.Text = ""

    If Me.DCboCashType.ListIndex = 5 And Me.TxtModFlg.Text = "N" Or Me.TxtModFlg.Text = "E" Then
        FillGridWithData (val(Me.DBCboClientName.BoundText)), TxtNoteSerial.Text
       ' Option4.value = True
       
       
                    
'                 pstate = get_project_Account(val(DBCboClientName.BoundText), "pstate")
'If pstate = 1 Then Option7.value = True Else Option7.value = False


    End If

    Dim DefaultSalesPersonId As Integer
    Dim fullcode As String

    GetCustomersDetail val(DBCboClientName.BoundText), DefaultSalesPersonId, fullcode

    'TxtCustCode.text = fullcode

    If DBCboClientName.BoundText = "" Then Exit Sub
 
    If 1 = 1 Then
  'fullcode = ""
  
        'Dim fullcode As String
      If Me.DCboCashType.ListIndex = 0 Or DCboCashType.ListIndex = 10 Or Me.DCboCashType.ListIndex = 11 Or DCboCashType.ListIndex = 12 Then
      
        GetCustomersDetail val(DBCboClientName.BoundText), DefaultSalesPersonId, fullcode
        TxtCustCode.Text = fullcode

        DcEmp.BoundText = DefaultSalesPersonId
        ElseIf Me.DCboCashType.ListIndex = 5 Then
        
       
        GetProjectsDetail val(DBCboClientName.BoundText), DefaultSalesPersonId, fullcode
       TxtCustCode.Text = fullcode

        
        
        
        End If
        
            If Me.TxtModFlg.Text = "N" Or Me.TxtModFlg.Text = "E" Then
            
            DcEmp.BoundText = DefaultSalesPersonId
            End If
            
                                 Dim IarType As Integer
            IarType = AqarCommisionType(val(DcbIqara.BoundText))
        
        
        
        If SystemOptions.CustomerhavethreeAccounts = True Then ' «·⁄„·«¡ ·Â« À·«À Õ”«»« 
        If DCboCashType.ListIndex = 10 And SystemOptions.NoCreatJLInRentContract = True Then
         
         If SystemOptions.NoCreatJLInRentContract = True And DCboCashType.ListIndex = 10 And IarType <> 0 Then
         Me.DcboCreditSide.BoundText = ModAccounts.GetMyAccountCode("TblCustemers", "CusID", val(DBCboClientName.BoundText), "Account_code")
         Else
         If SystemOptions.OpenAccountAqar = False Then
            Me.DcboCreditSide.BoundText = ModAccounts.GetMyAccountCode("TblCustemers", "CusID", val(Txtownerid), "Account_code")
         Else
            Me.DcboCreditSide.BoundText = GetAqarAcountCode(val(DcbIqara.BoundText))
         End If
    End If
      
                         '  Me.DcboCreditSide.BoundText = ModAccounts.GetMyAccountCode("TblCustemers", "CusID", val(Txtownerid), "Account_code")
                   
                            ElseIf CboPaymentType.ListIndex = 0 Then '‰ﬁœÌ
                                               If Option3.value = True Then 'œ›⁄«  „ﬁœ„…
                                                        Me.DcboCreditSide.BoundText = ModAccounts.GetMyAccountCode("TblCustemers", "CusID", val(Me.DBCboClientName.BoundText), "Account_code2")
                                             Else
                                                                 Me.DcboCreditSide.BoundText = ModAccounts.GetMyAccountCode("TblCustemers", "CusID", val(Me.DBCboClientName.BoundText))
                                             End If
                               
                            ElseIf CboPaymentType.ListIndex = 1 Then '‘Ìﬂ
                            
                                                If Option3.value = True Then 'œ›⁄«  „ﬁœ„…
                                                    Me.DcboCreditSide.BoundText = ModAccounts.GetMyAccountCode("TblCustemers", "CusID", val(Me.DBCboClientName.BoundText), "Account_code2")
                                             Else
                                                                 Me.DcboCreditSide.BoundText = ModAccounts.GetMyAccountCode("TblCustemers", "CusID", val(Me.DBCboClientName.BoundText), "Account_code1")
                                             End If
                                     
                             ElseIf CboPaymentType.ListIndex = 2 Then 'ÕÊ«·… '
                                                If Option3.value = True Then 'œ›⁄«  „ﬁœ„…
                                                        Me.DcboCreditSide.BoundText = ModAccounts.GetMyAccountCode("TblCustemers", "CusID", val(Me.DBCboClientName.BoundText), "Account_code2")
                                             Else
                                                                 Me.DcboCreditSide.BoundText = ModAccounts.GetMyAccountCode("TblCustemers", "CusID", val(Me.DBCboClientName.BoundText))
                                             End If
                              ElseIf CboPaymentType.ListIndex = 3 Then '‘Ìﬂ „”œœ '
                                                                    If Option3.value = True Then 'œ›⁄«  „ﬁœ„…
                                                        Me.DcboCreditSide.BoundText = ModAccounts.GetMyAccountCode("TblCustemers", "CusID", val(Me.DBCboClientName.BoundText), "Account_code2")
                                             Else
                                                                 Me.DcboCreditSide.BoundText = ModAccounts.GetMyAccountCode("TblCustemers", "CusID", val(Me.DBCboClientName.BoundText))
                                             End If
                                 ElseIf CboPaymentType.ListIndex = 4 Then
                                           If Option3.value = True Then 'œ›⁄«  „ﬁœ„…
                                                        Me.DcboCreditSide.BoundText = ModAccounts.GetMyAccountCode("TblCustemers", "CusID", val(Me.DBCboClientName.BoundText), "Account_code2")
                                             Else
                                                                 Me.DcboCreditSide.BoundText = ModAccounts.GetMyAccountCode("TblCustemers", "CusID", val(Me.DBCboClientName.BoundText))
                                             End If
                              End If
                             
'
        Else '«·⁄„·«¡ ·Â„ Õ”«» Ê«Õœ ›ﬁÿ
        If DCboCashType.ListIndex = 10 And SystemOptions.NoCreatJLInRentContract = True Then
   
         If SystemOptions.NoCreatJLInRentContract = True And DCboCashType.ListIndex = 10 And IarType <> 0 Then
         Me.DcboCreditSide.BoundText = ModAccounts.GetMyAccountCode("TblCustemers", "CusID", val(DBCboClientName.BoundText), "Account_code")
         Else
          If SystemOptions.OpenAccountAqar = False Then
            Me.DcboCreditSide.BoundText = ModAccounts.GetMyAccountCode("TblCustemers", "CusID", val(Txtownerid), "Account_code")
         Else
            Me.DcboCreditSide.BoundText = GetAqarAcountCode(val(DcbIqara.BoundText))
         End If
         
         End If
                  
        Else
                Me.DcboCreditSide.BoundText = ModAccounts.GetMyAccountCode("TblCustemers", "CusID", val(Me.DBCboClientName.BoundText))
        End If
        End If
        

        If DCboCashType.ListIndex = 5 Then 'Õ«·… «·„‘«—Ì⁄
                                        
       If Option4.value = True Then ' ⁄„Ì· ‰Â«∆Ì
                                        
        If SystemOptions.CustomerhavethreeAccounts = True Then ' «·⁄„·«¡ ·Â« À·«À Õ”«»« 
        
                            If CboPaymentType.ListIndex = 0 Then '‰ﬁœÌ
                                                                    If Option3.value = True Then 'œ›⁄Â „ﬁœ„…
                                                                           Me.DcboCreditSide.BoundText = get_project_customer_account(val(DBCboClientName.BoundText), "Account_Code2") 'œ›⁄«  „ﬁœ„…
                                                                      Else
                                                                           Me.DcboCreditSide.BoundText = get_project_customer_account(val(DBCboClientName.BoundText), "Account_Code") ' Ã«—Ì
                                                                      End If
                               
                            ElseIf CboPaymentType.ListIndex = 1 Then '‘Ìﬂ
                            
                                                                If Option3.value = True Then 'œ›⁄Â „ﬁœ„…
                                                                    Me.DcboCreditSide.BoundText = get_project_customer_account(val(DBCboClientName.BoundText), "Account_Code2") 'œ›⁄«  „ﬁœ„…
                                                                      Else
                                                                          Me.DcboCreditSide.BoundText = get_project_customer_account(val(DBCboClientName.BoundText), "Account_Code1") '  Õ  «· Õ’Ì·
                                                                      End If
                                     
                             ElseIf CboPaymentType.ListIndex = 2 Then 'ÕÊ«·… '
                                               If Option3.value = True Then 'œ›⁄Â „ﬁœ„…
                                                                    Me.DcboCreditSide.BoundText = get_project_customer_account(val(DBCboClientName.BoundText), "Account_Code2") 'œ›⁄«  „ﬁœ„…
                                                                      Else
                                                                          Me.DcboCreditSide.BoundText = get_project_customer_account(val(DBCboClientName.BoundText), "Account_Code") ' Ã«—Ì
                                                                      End If
                              ElseIf CboPaymentType.ListIndex = 3 Then '‘Ìﬂ „”œœ '
                                                      If Option3.value = True Then 'œ›⁄Â „ﬁœ„…
                                                  Me.DcboCreditSide.BoundText = get_project_customer_account(val(DBCboClientName.BoundText), "Account_Code2") 'œ›⁄«  „ﬁœ„…
                                                                      Else
                                                                          Me.DcboCreditSide.BoundText = get_project_customer_account(val(DBCboClientName.BoundText), "Account_Code") ' Ã«—Ì
                                                                      End If
                              End If
                             
'
        Else '«·⁄„·«¡ ·Â„ Õ”«» Ê«Õœ ›ﬁÿ
                Me.DcboCreditSide.BoundText = get_project_customer_account(val(DBCboClientName.BoundText), "Account_Code") ' Ã«—Ì

        End If
                                                
                                                
                                                
          Else '⁄„Ì· «·»«ÿ‰55555555555555555555555555555555555555555
          
                  If SystemOptions.CustomerhavethreeAccounts = True Then ' «·⁄„·«¡ ·Â« À·«À Õ”«»« 
        
                            If CboPaymentType.ListIndex = 0 Then '‰ﬁœÌ
                                                If Option3.value = True Then 'œ›⁄Â „ﬁœ„…
                                                                    Me.DcboCreditSide.BoundText = get_project_customer_account(val(DBCboClientName.BoundText), "Account_Code2", 1) 'œ›⁄«  „ﬁœ„…
                                                                      Else
                                                                          Me.DcboCreditSide.BoundText = get_project_customer_account(val(DBCboClientName.BoundText), "Account_Code", 1) ' Ã«—Ì
                                                                      End If
                               
                            ElseIf CboPaymentType.ListIndex = 1 Then '‘Ìﬂ
                            
                                                                If Option3.value = True Then 'œ›⁄Â „ﬁœ„…
                                                                    Me.DcboCreditSide.BoundText = get_project_customer_account(val(DBCboClientName.BoundText), "Account_Code2", 1) 'œ›⁄«  „ﬁœ„…
                                                                      Else
                                                                          Me.DcboCreditSide.BoundText = get_project_customer_account(val(DBCboClientName.BoundText), "Account_Code1", 1) '  Õ  «· Õ’Ì·
                                                                      End If
                                     
                             ElseIf CboPaymentType.ListIndex = 2 Then 'ÕÊ«·… '
                                               If Option3.value = True Then 'œ›⁄Â „ﬁœ„…
                                                                    Me.DcboCreditSide.BoundText = get_project_customer_account(val(DBCboClientName.BoundText), "Account_Code2", 1) 'œ›⁄«  „ﬁœ„…
                                                                      Else
                                                                          Me.DcboCreditSide.BoundText = get_project_customer_account(val(DBCboClientName.BoundText), "Account_Code", 1) ' Ã«—Ì
                                                                      End If
                              ElseIf CboPaymentType.ListIndex = 3 Then '‘Ìﬂ „”œœ '
                                                      If Option3.value = True Then 'œ›⁄Â „ﬁœ„…
                                                  Me.DcboCreditSide.BoundText = get_project_customer_account(val(DBCboClientName.BoundText), "Account_Code2", 1) 'œ›⁄«  „ﬁœ„…
                                                                      Else
                                                                          Me.DcboCreditSide.BoundText = get_project_customer_account(val(DBCboClientName.BoundText), "Account_Code", 1) ' Ã«—Ì
                                                                      End If
                              End If
                             
'
        Else '«·⁄„·«¡ ·Â„ Õ”«» Ê«Õœ ›ﬁÿ
                Me.DcboCreditSide.BoundText = get_project_customer_account(val(DBCboClientName.BoundText), "Account_Code", 1) ' Ã«—Ì

        End If
        
          
          
          
          End If
       End If
    End If

Dim Account_Code_dynamic As String

      If DCboCashType.ListIndex = 9 Then
             
             
                                 Account_Code_dynamic = get_account_code_branch(95, my_branch)
        
        If Account_Code_dynamic = "NO branch" Then
            If SystemOptions.UserInterface = ArabicInterface Then
                MsgBox "·„ Ì „ «‰‘«¡ «·›—⁄", vbCritical
            Else
                MsgBox "Branch Not Created ", vbCritical
            End If

            GoTo ErrTrap
        ElseIf Account_Code_dynamic = "NO account" Then

            If SystemOptions.UserInterface = ArabicInterface Then
                MsgBox "·„ Ì „  ÕœÌœ Õ”«»  „œ›Ê⁄«  „ﬁœ„… ·ÕÃ“ «·ÊÕœ«  ", vbCritical
            Else
                MsgBox "   Insatllemts Revenu Not Deined in this Branch", vbCritical
            End If

            GoTo ErrTrap
         
   
    
        End If
             '    If Me.TxtModFlg.text = "N" Or Me.TxtModFlg.text = "E" Then
        Me.DcboCreditSide.BoundText = Account_Code_dynamic
             End If
    
    
    
ErrTrap:
End Sub
Private Sub DBCboClientName_KeyUp(KeyCode As Integer, _
                                  Shift As Integer)

 '   If DCboCashType.ListIndex = 0 Then
 '       If KeyCode = vbKeyF3 Then
 '        FrmCustemerSearch.SearchType = 3
 '           FrmCustemerSearch.show vbModal
 '
 '       End If
'
'    ElseIf DCboCashType.ListIndex = 1 Then
'
'        If KeyCode = vbKeyF3 Then
'          FrmCompanySearch.lblSearchtype.Caption = 2
'            FrmCompanySearch.show vbModal
          
'        End If
'
'   ElseIf DCboCashType.ListIndex = 5 Then
'
'        If KeyCode = vbKeyF3 Then
'         FrmProjectSearch.lblSearchtype.Caption = 1
'             FrmProjectSearch.show vbModal
'
'        End If
'
'    End If




    If DCboCashType.ListIndex = 0 Or DCboCashType.ListIndex = 10 Or DCboCashType.ListIndex = 11 Or DCboCashType.ListIndex = 12 Then
        If KeyCode = vbKeyF3 Then
         FrmCustemerSearch.SearchType = 2412
             FrmCustemerSearch.show vbModal
           
        End If

    ElseIf DCboCashType.ListIndex = 1 Then

        If KeyCode = vbKeyF3 Then
          FrmCompanySearch.lblSearchtype.Caption = 1915
            FrmCompanySearch.show vbModal
          
        End If

   ElseIf DCboCashType.ListIndex = 5 Then

        If KeyCode = vbKeyF3 Then
         FrmProjectSearch.lblSearchtype.Caption = 1
             FrmProjectSearch.show vbModal
           
        End If
   ElseIf DCboCashType.ListIndex = 6 Then
    End If
    
End Sub

Private Sub DCAccounts_Change()

    If Me.TxtModFlg.Text = "N" Or Me.TxtModFlg.Text = "E" Then
        '   Me.DcboCreditSide.BoundText = ModAccounts.GetMyAccountCode("TblRevenuesTypes", "RevenuesID", Val(Me.DcboRevenuesTypes.BoundText))
        Me.DcboCreditSide.BoundText = DCAccounts.BoundText
  
    End If

End Sub

Private Sub DCAccounts_KeyUp(KeyCode As Integer, _
                             Shift As Integer)

    If KeyCode = vbKeyF3 Then
    DCAccounts.Text = ""
        Unload Account_search
        Account_search.show
        Account_search.case_id = 1200
            
    End If

End Sub

Private Sub DcbAccount_Change()
DcbAccount_Click (0)
End Sub

Private Sub DcbAccount_Click(Area As Integer)
TxtAccount.Text = getAccountSerial_Code("Account_Serial", "Account_Code", DcbAccount.BoundText)
If 1 = 1 Then
        If DcbAccount.BoundText <> "" Then
            Me.DcboDebitSide.BoundText = DcbAccount.BoundText
        End If
 End If
End Sub

Private Sub DcbIqara_Change()
DcbIqara_Click (0)
End Sub

Private Sub DcbIqara_Click(Area As Integer)
If Me.TxtModFlg.Text <> "R" Then
If val(Me.DcbIqara.BoundText) <> 0 Then
GetAmola val(Me.DcbIqara.BoundText)
End If
End If
End Sub
Sub GetAmola(Optional Aqarid As Variant = 0)
If Aqarid <> 0 Then
Dim Rs9  As ADODB.Recordset
Set Rs9 = New ADODB.Recordset
Dim sql As String
sql = "select * from tblaqar where Aqarid =" & Aqarid & ""
Rs9.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
If Rs9.RecordCount > 0 Then

Txtownerid = IIf(IsNull(Rs9("ownerid").value), 0, Rs9("ownerid").value)

   If Not IsNull(Rs9("TypAmola").value) Then
   
                        If Rs9("TypAmola").value = 1 Then
                        Rd(1).value = True
                         Else
                        Rd(0).value = True
                        End If
   Else
        Rd(0).value = True
   End If
        
      TxtKickbacks.Text = IIf(IsNull(Rs9("AmolaValus").value), 0, Rs9("AmolaValus").value)
End If
End If
End Sub
Sub GetWonerID(Optional Aqarid As Integer = 0)
If Aqarid <> 0 Then
Dim Rs9  As ADODB.Recordset
Set Rs9 = New ADODB.Recordset
Dim sql As String
sql = "select * from tblaqar where Aqarid =" & Aqarid & ""
Rs9.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
If Rs9.RecordCount > 0 Then
Txtownerid.Text = IIf(IsNull(Rs9("ownerid").value), 0, Rs9("ownerid").value)
End If
End If
End Sub

Private Sub DcboBankName_Change()
DcboBankName_Click (0)
End Sub

Private Sub DcboBankName_Click(Area As Integer)

    If DcboBankName.BoundText = "" Then Exit Sub
    Dim RsSavRec As ADODB.Recordset
    Dim My_SQL As String
    Dim Account_Code_dynamic As String

    'If Me.TxtModFlg.Text = "N" Or Me.TxtModFlg.Text = "E" Then
    If 1 = 1 Then
        'Me.DcboDebitSide.BoundText =   "a1a2a4"
        My_SQL = "  select Account_Code from BanksData WHERE BankID=" & DcboBankName.BoundText

        Set RsSavRec = New ADODB.Recordset
        RsSavRec.CursorLocation = adUseClient
        RsSavRec.Open My_SQL, Cn, adOpenStatic, adLockOptimistic, adCmdText
 
        If SystemOptions.ChequeBox = True Then
            Me.DcboDebitSide.BoundText = ""
        Else

            If SystemOptions.banks_Accounts3 = True Then
                Me.DcboDebitSide.BoundText = get_bank_Account(val(Me.DcboBankName.BoundText), "Account_Code1")
            Else
                Me.DcboDebitSide.BoundText = RsSavRec.Fields("Account_Code").value
                     
            End If
        End If

        If CboPaymentType.ListIndex = 2 Or CboPaymentType.ListIndex = 3 Then
                     
            Me.DcboDebitSide.BoundText = RsSavRec.Fields("Account_Code").value
                    
        End If

    End If

End Sub

Private Sub DcboBox_Change()

    'If Me.TxtModFlg.Text = "N" Or Me.TxtModFlg.Text = "E" Then
    If 1 = 1 Then
        Me.DcboDebitSide.BoundText = ModAccounts.GetMyAccountCode("TblBoxesData", "BoxID", val(Me.DcboBox.BoundText))
    End If

End Sub


Private Sub DCboCashType_Change()
    On Error GoTo ErrTrap
    lbl(108).Visible = False
    Frame2.Enabled = False
    Dim StrSQL As String
    Dim intDef As Integer
txtContractNo.Visible = False
lbl(53).Visible = False
Frame11.Visible = False
Frame12.Visible = False
Frame13.Visible = False
XPTxtVal.Enabled = True
txtTotalinsuranceS.Visible = False
lbl(109).Visible = False
TxtVATValue.Visible = False
' C1Elastic1.Visible = False
Me.TXtFilter.Visible = False
Me.TxtFilterNo.Visible = False
lbl(61).Visible = False
lbl(60).Visible = False
ISButton3.Visible = False
ISButton1.Visible = False
Frame9.Visible = False
Frame6.Visible = True
txtContractNo.Visible = False
Me.TXtFilter.Visible = False
Me.TxtFilterNo.Visible = False
Me.XPTxtVal.Enabled = True

    Select Case DCboCashType.ListIndex

        Case 0
            Dcombos.GetCustomersSuppliers 56, Me.DBCboClientName, False
            Me.DBCboClientName.Visible = True
            Me.DcboRevenuesTypes.Visible = False
        
            DCEmployee.Visible = False
            DCAccounts.Visible = False
            ChkTrans.Visible = True
            Fra(0).Visible = True

            If SystemOptions.UserInterface <> EnglishInterface Then
                Me.lbl(3).Caption = "«”„ «·„” √Ã—"
            Else
                Me.lbl(3).Caption = "Renter Name"
            End If
        
            Me.lbl(13).Visible = True
            Me.LblLink.Visible = True

        Case 1
        
         Frame9.Visible = False
            Dcombos.GetCustomersSuppliers 56, Me.DBCboClientName, False
            Me.DBCboClientName.Visible = True
            Me.DcboRevenuesTypes.Visible = False
            DCEmployee.Visible = False
            DCAccounts.Visible = False
            ChkTrans.Visible = True
            Fra(0).Visible = True

            If SystemOptions.UserInterface <> EnglishInterface Then
                Me.lbl(3).Caption = "«”„ «·„” √Ã—"
            Else
                Me.lbl(3).Caption = "Renter Name"
            End If
        
            Me.lbl(13).Visible = True
            Me.LblLink.Visible = True

        Case 2
    
        Frame9.Visible = False
            Dcombos.GetPersons Me.DBCboClientName
            Me.DBCboClientName.Visible = True
            Me.DcboRevenuesTypes.Visible = False
            DCEmployee.Visible = False
            DCAccounts.Visible = False
            ChkTrans.Visible = False
            Fra(0).Visible = False

            If SystemOptions.UserInterface = EnglishInterface Then
                Me.lbl(3).Caption = "name"
            Else
                Me.lbl(3).Caption = "„ﬁ«Ê· «·»«ÿ‰"
            End If
                
            Me.lbl(13).Visible = True
            Me.LblLink.Visible = True

        Case 3
       
        Frame9.Visible = False
            '≈Ì—«œ«  ≈Œ—Ï
            Me.DBCboClientName.Visible = False
            Me.DcboRevenuesTypes.Visible = True
            Me.ChkTrans.Visible = False
            DBCboClientName.Visible = False
            DCEmployee.Visible = False
            DCAccounts.Visible = False
            Fra(0).Visible = False
        
            If SystemOptions.UserInterface <> EnglishInterface Then
                Me.lbl(3).Caption = "‰Ê⁄ «·«Ì—«œ"
            Else
                Me.lbl(3).Caption = "RVN Type"
            End If
                
            Me.lbl(13).Visible = False
            Me.LblLink.Visible = False
        
        Case 4
       
        Frame9.Visible = False
            Dcombos.GetCustomersSuppliers 56, Me.DBCboClientName, True
            Me.DBCboClientName.Visible = True
            Me.DcboRevenuesTypes.Visible = False
            DCEmployee.Visible = False
            DCAccounts.Visible = False
            ChkTrans.Visible = True
            Fra(0).Visible = True

            If SystemOptions.UserInterface <> EnglishInterface Then
                Me.lbl(3).Caption = "«”„ «·„” √Ã—"
            Else
                Me.lbl(3).Caption = "Renter Name"
            End If
        
            Me.lbl(13).Visible = True
            Me.LblLink.Visible = True
        
        Case 5
      
        Frame9.Visible = False
            Dim My_SQL As String
            My_SQL = "  select id,Project_name from projects where not(REVENUE_account is null) order by Project_name" '
            fill_combo Me.DBCboClientName, My_SQL
         
            Me.DBCboClientName.Visible = True
            Me.DcboRevenuesTypes.Visible = False
            DCEmployee.Visible = False
            DCAccounts.Visible = False

            If SystemOptions.UserInterface <> EnglishInterface Then
                Me.lbl(3).Caption = "«”„ «·„‘—Ê⁄"
            Else
                Me.lbl(3).Caption = "project Name"
            End If
        
            Frame2.Enabled = True
        
        Case 6
     
        Frame9.Visible = False
            Dcombos.GetEmployees Me.DCEmployee
            Me.DCEmployee.Visible = True
            Me.DcboRevenuesTypes.Visible = False
            DBCboClientName.Visible = False
            DCAccounts.Visible = False
            ChkTrans.Visible = True

            '   Fra(0).Visible = True
            If SystemOptions.UserInterface <> EnglishInterface Then
                Me.lbl(3).Caption = "«”„ «·„ÊŸ›"
            Else
                Me.lbl(3).Caption = "Employee  Name"
            End If

        Case 7

        Frame9.Visible = False
            Dcombos.GetAccountingCodes Me.DCAccounts, True
            DCAccounts.Visible = True
            Me.DCEmployee.Visible = False
            Me.DcboRevenuesTypes.Visible = False
            DBCboClientName.Visible = False
        
            ChkTrans.Visible = True

            '   Fra(0).Visible = True
            If SystemOptions.UserInterface <> EnglishInterface Then
                Me.lbl(3).Caption = "«”„ «·Õ”«»"
            Else
                Me.lbl(3).Caption = "Accounts Nam  "
            End If
        
            '  Me.lbl(13).Visible = True
            '      Me.LblLink.Visible = True
    
Case 8 '  „‰ ⁄ﬁœ
lbl(108).Visible = True
TxtVATValue.Visible = True
txtContractNo.Visible = False
lbl(53).Caption = "—ﬁ„ «·⁄ﬁœ"
Me.XPTxtVal.Enabled = False
Frame6.Visible = False
txtContractNo.Visible = True
ISButton1.Visible = True
C1Elastic1.Visible = True
            Dcombos.GetCustomersSuppliers 56, Me.DBCboClientName, False
            Me.DBCboClientName.Visible = True
            Me.DcboRevenuesTypes.Visible = False
        Frame9.Visible = True
            DCEmployee.Visible = False
            DCAccounts.Visible = False
            ChkTrans.Visible = True
            Fra(0).Visible = True

            If SystemOptions.UserInterface <> EnglishInterface Then
                Me.lbl(3).Caption = "«”„ «·„” √Ã—"
            Else
                Me.lbl(3).Caption = "Customer Name"
            End If
        
            Me.lbl(13).Visible = True
            Me.LblLink.Visible = True
txtContractNo.Visible = True
lbl(53).Visible = True
''//„‰  ’›ÌÂ
Case 10
TxtBookNo.Visible = False
Me.TXtFilter.Visible = True
Me.TxtFilterNo.Visible = True
ISButton3.Visible = True
ISButton1.Visible = False
txtTotalinsuranceS.Visible = True
lbl(109).Visible = True
Frame9.Visible = False
            Dcombos.GetCustomersSuppliers 56, Me.DBCboClientName, False
            Me.DBCboClientName.Visible = True
            Me.DcboRevenuesTypes.Visible = False
   
            DCEmployee.Visible = False
            DCAccounts.Visible = False
            ChkTrans.Visible = True
            Fra(0).Visible = True

            If SystemOptions.UserInterface <> EnglishInterface Then
                Me.lbl(3).Caption = "«”„ «·„” √Ã—"
            Else
                Me.lbl(3).Caption = "Customer Name"
            End If
        
            Me.lbl(13).Visible = True
            Me.LblLink.Visible = False
txtContractNo.Visible = False
lbl(53).Visible = False
Me.TXtFilter.Visible = True
Me.TxtFilterNo.Visible = True
lbl(61).Visible = True
lbl(60).Visible = True
lbl(108).Visible = True
TxtVATValue.Visible = True

   Case 9
       Dcombos.GetCustomersSuppliers 56, Me.DBCboClientName, False
            Me.DBCboClientName.Visible = True
         lbl(108).Visible = True
         TxtVATValue.Visible = True
         txtinstrunce.Visible = True
         TxtWater.Visible = True
         TxtCommissionOut.Visible = True
         Txtcommission.Visible = True
         TxtRent.Visible = True
         Label1(6).Visible = True
         Label1(5).Visible = True
         Label1(12).Visible = True
         Label1(2).Visible = True
         Label1(3).Visible = True
               cbointervaltype.Visible = True
        TxtInterval.Visible = True
        DcbUnitNo.Visible = True
        Label5.Visible = True
        DcbUnitType.Visible = True
        Frame7.Visible = True
        Label1(4).Visible = True
        Label1(15).Visible = True
        Label1(14).Visible = True
        DcbIqara.Visible = True
        TxtSearch.Visible = True
        
        
        
   Frame6.Visible = True
   Frame9.Visible = False
   C1Elastic1.Visible = True
            Dim Account_Code_dynamic As String
                    Account_Code_dynamic = get_account_code_branch(95, my_branch)
        
        If Account_Code_dynamic = "NO branch" Then
            If SystemOptions.UserInterface = ArabicInterface Then
                MsgBox "·„ Ì „ «‰‘«¡ «·›—⁄", vbCritical
            Else
                MsgBox "Branch Not Created ", vbCritical
            End If

            GoTo ErrTrap
        ElseIf Account_Code_dynamic = "NO account" Then

            If SystemOptions.UserInterface = ArabicInterface Then
                MsgBox "·„ Ì „  ÕœÌœ Õ”«»  „œ›Ê⁄«  „ﬁœ„… ·ÕÃ“ «·ÊÕœ«  ", vbCritical
            Else
                MsgBox "   Insatllemts Revenu Not Deined in this Branch", vbCritical
            End If

            GoTo ErrTrap
         
   
    
        End If
             '    If Me.TxtModFlg.text = "N" Or Me.TxtModFlg.text = "E" Then
        Me.DcboCreditSide.BoundText = Account_Code_dynamic
        Case 12
      '  txtinstrunce.Visible = False
      '  txtWater.Visible = False
      '  TxtCommissionOut.Visible = False
      '  TxtCommission.Visible = False
      '  TxtRent.Visible = False
      '   Label1(6).Visible = False
      '   Label1(5).Visible = False
      '   Label1(12).Visible = False
      '   Label1(2).Visible = False
      '   Label1(3).Visible = False
      '         cbointervaltype.Visible = False
      '  txtinterval.Visible = False
      '  DcbUnitNo.Visible = False
      '  Label5.Visible = False
      '  DcbUnitType.Visible = False
      '  Frame7.Visible = False
      '  Label1(4).Visible = False
      '  Label1(15).Visible = False
      '  Label1(14).Visible = False
      '  DcbIqara.Visible = False
      '  TxtSearch.Visible = False
        C1Elastic1.Visible = True
      Frame6.Visible = True
     Case 13
     ISButton1.Visible = True
     XPTxtVal.Enabled = False
     Frame12.Visible = True
     txtContractNo.Visible = True
      lbl(53).Visible = True
      lbl(53).Caption = "—ﬁ„ «·Õ—ﬂ…"
      If RdTypeTrans(0).value = True Then
     Frame11.Visible = True
     Frame13.Visible = False
     Else
     Frame11.Visible = False
     Frame13.Visible = True
     End If
     txtContractNo.Visible = True
    'End If
    End Select
If val(DCboCashType2.ListIndex) = -1 And val(DCboCashType.ListIndex) >= 7 Then
DCboCashType2.ListIndex = val(DCboCashType.ListIndex) - 7
End If
    cSearchDcbo.Refresh
    Exit Sub
ErrTrap:
End Sub

Private Sub DCboCashType_Click()
    DCboCashType_Change
End Sub

Private Sub DCboCashType2_Change()
If val(DCboCashType2.ListIndex) <> -1 Then
DCboCashType.ListIndex = val(DCboCashType2.ListIndex) + 7
Else
DCboCashType.ListIndex = -1
End If
End Sub

Private Sub DCboCashType2_Click()
DCboCashType2_Change
End Sub

Private Sub DcboCreditSide_Change()

    WriteCustomerBalPublic Me.DcboCreditSide.BoundText, Balance, balanceString
    LblLink.Caption = balanceString
End Sub

Private Sub DcboRevenuesTypes_Change()

    If Me.TxtModFlg.Text = "N" Or Me.TxtModFlg.Text = "E" Then
        Me.DcboCreditSide.BoundText = ModAccounts.GetMyAccountCode("TblRevenuesTypes", "RevenuesID", val(Me.DcboRevenuesTypes.BoundText))


    End If

End Sub

Private Sub Dcbranch_Click(Area As Integer)
If Me.TxtModFlg.Text <> "R" Then
    TxtNoteSerial.Text = ""
    TxtNoteSerial1.Text = ""
 End If
End Sub

Private Sub Dcbranch_GotFocus()
Dcbranch_Click (0)
End Sub

Private Sub DcbUnitNo_Change()
On Error Resume Next
Dim Typed As Integer
If Me.TxtModFlg <> "R" And Me.TxtModFlg <> "" Then
Dim str As String
If val(DCboCashType.ListIndex) = 9 Then
 str = checkDepositeRent(val(DcbUnitNo.BoundText), XPDtbTrans)
If str <> "" Then
MsgBox str, vbInformation
End If
GetIqarUnitData val(DcbUnitNo.BoundText), , , , , , , , , , , , , Typed
ComResid(Typed).value = True
End If
End If
End Sub

Private Sub DcbUnitNo_Click(Area As Integer)
DcbUnitNo_Change
End Sub

Private Sub DcbUnitType_Change()
Dim Dcombos As ClsDataCombos
Dim idd As Long
Dim idd1 As Long
   Set Dcombos = New ClsDataCombos
  ' Dcombos.GetIqarUnit idd, idd1, Me.DcbUnitNo, "R"

If val(DcbIqara.BoundText) > 0 Then
idd = val(DcbIqara.BoundText)

idd1 = val(DcbUnitType.BoundText)
If DCboCashType.ListIndex = 9 Then
If Me.TxtModFlg = "R" Then
Dcombos.GetIqarUnit idd, idd1, Me.DcbUnitNo, "R"
Else
Dcombos.GetIqarUnit idd, idd1, Me.DcbUnitNo
End If
Else
Dcombos.GetIqarUnit idd, idd1, Me.DcbUnitNo, "R"
End If
End If
End Sub

Private Sub DcbUnitType_Click(Area As Integer)
DcbUnitType_Change
End Sub

Private Sub DcChequeBox_Change()

    If Me.TxtModFlg.Text = "N" Or Me.TxtModFlg.Text = "E" Then
        Me.DcboDebitSide.BoundText = ModAccounts.GetMyAccountCodeRefined("TblBoxesData", "BoxID", val(Me.DCChequeBox.BoundText), "Account_Code1")
    End If

End Sub

Private Sub DcCostCenter_KeyUp(KeyCode As Integer, _
                               Shift As Integer)

    If KeyCode = vbKeyF3 Then
        CostCenterSearch.show
        CostCenterSearch.RetrunType = 6
    End If

End Sub



Private Sub dcEmp_Change()
Dcemp_Click (0)
End Sub

Private Sub Dcemp_Click(Area As Integer)
Dim i As Integer
If val(Me.DcEmp.BoundText) <> 0 Then

With VSFlexGrid1
For i = 1 To .rows - 1
If .TextMatrix(i, .ColIndex("id")) = Me.DcEmp.BoundText Then
Exit Sub
End If
Next i
If .rows = 2 Then
.TextMatrix(.rows - 1, .ColIndex("rate")) = 100
End If
.TextMatrix(.rows - 1, .ColIndex("id")) = Me.DcEmp.BoundText
.TextMatrix(.rows - 1, .ColIndex("empname")) = Me.DcEmp.Text
.rows = .rows + 1
End With
With VSFlexGrid2
For i = 1 To .rows - 1
If .TextMatrix(i, .ColIndex("id")) = Me.DcEmp.BoundText Then
Exit Sub
End If
Next i
If .rows = 2 Then
.TextMatrix(.rows - 1, .ColIndex("rate")) = 100
End If
.TextMatrix(.rows - 1, .ColIndex("id")) = Me.DcEmp.BoundText
.TextMatrix(.rows - 1, .ColIndex("empname")) = Me.DcEmp.Text
.rows = .rows + 1
End With
End If

End Sub

Private Sub DcEmployee_Change()

    If Me.TxtModFlg.Text = "N" Or Me.TxtModFlg.Text = "E" Then
        '   Me.DcboCreditSide.BoundText = ModAccounts.GetMyAccountCode("TblRevenuesTypes", "RevenuesID", Val(Me.DcboRevenuesTypes.BoundText))
        Me.DcboCreditSide.BoundText = get_EMPLOYEE_Account(val(DCEmployee.BoundText), "Account_Code")
       ' TxtCustCode.text = val(dcEmployee.BoundText)
        TxtCustCode.Text = getemployeeCode(val(DCEmployee.BoundText))
       
       
    End If

End Sub

Private Sub dcCar_Change()

    GetDriverInformation (val(DCCar.BoundText))

End Sub

Private Sub dcCar_Click(Area As Integer)
    GetDriverInformation (val(DCCar.BoundText))

End Sub
Function CheckStatusofUnit(ID As Integer) As Boolean

    If Me.TxtModFlg.Text = "N" Or Me.TxtModFlg.Text = "E" Then
        Dim sql As String
        Dim Status As Boolean
        Dim i As Integer
        Dim rs2 As New ADODB.Recordset
 Status = True
        sql = " SELECT    StatusEarnest "
        sql = sql & " from dbo.TblAqrEarnest"
        sql = sql & " Where (UnitNo = " & ID & ") "

        rs2.Open sql, Cn, adOpenStatic, adLockReadOnly, adCmdText

        If rs2.RecordCount > 0 Then
        
        For i = 1 To rs.RecordCount
        If rs2("StatusEarnest").value = 0 Then
        Status = False
     End If
        
  Next i
    End If
    CheckStatusofUnit = Status
End If
End Function
Function GetDriverInformation(ID As Double)

    If Me.TxtModFlg.Text = "N" Or Me.TxtModFlg.Text = "E" Then
        Dim sql As String
        Dim rs2 As New ADODB.Recordset
 
        sql = " SELECT    * "
        sql = sql & " from dbo.TblCarsData"
        sql = sql & " Where (id = " & ID & ") "

        rs2.Open sql, Cn, adOpenStatic, adLockReadOnly, adCmdText

        If rs2.RecordCount > 0 Then
            DCDriver.BoundText = IIf(IsNull(rs2("Emp_id").value), 0, rs2("Emp_id").value)
                  
        Else
            DcEmp = 0
               
        End If

    End If

End Function

Private Sub DCEmployee_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF3 Then
        FrmEmployeeSearch.lbltype = 36
       ' Set FrmEmployeeSearch.RetrunFrm = Me

        FrmEmployeeSearch.show
  
    End If
    
End Sub

Private Sub Ele_Click(Index As Integer)
C1Elastic1.Visible = True
End Sub
Public Sub FillGridWithData1(Optional ContNo As Double = 0)
    On Error GoTo ErrTrap
    Dim i As Integer
    Dim rs2 As ADODB.Recordset
    Dim My_SQL As String
    Set rs2 = New ADODB.Recordset
    My_SQL = "SELECT     dbo.Notes.NoteID, dbo.Notes.NoteDate, dbo.Notes.NoteType, dbo.Notes.NoteSerial1, dbo.Notes.ContNo, dbo.Notes.ContractNo, "
    My_SQL = My_SQL & "                   dbo.Notes.Note_Value ,dbo.Notes.Note_Value2 , dbo.Notes.NoteDateH, dbo.Notes.CusID, dbo.TblCustemers.CusName, dbo.TblCustemers.CusNamee, dbo.TblCustemers.fullcode"
    My_SQL = My_SQL & "  FROM         dbo.Notes LEFT OUTER JOIN"
    My_SQL = My_SQL & "                   dbo.TblCustemers ON dbo.Notes.CusID = dbo.TblCustemers.CusID"
    My_SQL = My_SQL & "  WHERE     (dbo.Notes.NoteType = 4) AND (dbo.Notes.ContNo = " & val(ContNo) & ")"
    rs2.Open My_SQL, Cn, adOpenKeyset, adLockReadOnly, adCmdText

    With Me.VSFlexGrid3
        .rows = 2
        .Clear flexClearScrollable

        If rs2.RecordCount > 0 Then
            .rows = rs2.RecordCount + 1
            rs2.MoveFirst

            For i = 1 To .rows - 1
                .TextMatrix(i, .ColIndex("Ser")) = i
                .TextMatrix(i, .ColIndex("id")) = IIf(IsNull(rs2.Fields("NoteID").value), 0, rs2.Fields("NoteID").value)
             .TextMatrix(i, .ColIndex("ContractNo")) = IIf(IsNull(rs2.Fields("ContractNo").value), "", rs2.Fields("ContractNo").value)
             .TextMatrix(i, .ColIndex("ContNo")) = IIf(IsNull(rs2.Fields("ContNo").value), 0, rs2.Fields("ContNo").value)
                .TextMatrix(i, .ColIndex("NoteDate")) = IIf(IsNull(rs2.Fields("NoteDate").value), "", rs2.Fields("NoteDate").value)
                .TextMatrix(i, .ColIndex("NoteSerial1")) = IIf(IsNull(rs2.Fields("NoteSerial1").value), "", rs2.Fields("NoteSerial1").value)
               
                .TextMatrix(i, .ColIndex("NoteDateH")) = IIf(IsNull(rs2.Fields("NoteDateH").value), "", rs2.Fields("NoteDateH").value)
           
                .TextMatrix(i, .ColIndex("NoteSerial1")) = IIf(IsNull(rs2.Fields("NoteSerial1").value), "", rs2.Fields("NoteSerial1").value)
            
                .TextMatrix(i, .ColIndex("Note_Value")) = IIf(IsNull(rs2.Fields("Note_Value2").value), "", rs2.Fields("Note_Value2").value)
            If SystemOptions.UserInterface = ArabicInterface Then
                .TextMatrix(i, .ColIndex("CusName")) = IIf(IsNull(rs2.Fields("CusName").value), "", rs2.Fields("CusName").value)
                Else
                .TextMatrix(i, .ColIndex("CusName")) = IIf(IsNull(rs2.Fields("CusNamee").value), "", rs2.Fields("CusNamee").value)
                End If
            
                rs2.MoveNext
            Next

            rs2.Close
        End If

        .RowHeight(-1) = 300
    End With

ErrTrap:
End Sub
Function ChekPayeArbon(Optional NotID As Double = 0) As Boolean
Dim Rs7 As ADODB.Recordset
Set Rs7 = New ADODB.Recordset
Dim sql As String
sql = "Select * from Notes Where NoteID=" & NotID & " and PayedOrBon=1 and NoteType=4  "
Rs7.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
If Rs7.RecordCount > 0 Then
ChekPayeArbon = True
Else
ChekPayeArbon = False
End If
End Function

Private Sub Form_Load()
 On Error GoTo ErrTrap
    Me.left = (mdifrmmain.Width - Me.Width) / 2
    Me.top = (mdifrmmain.Height - Me.Height) / 2 - 500


    
If SystemOptions.SpecialVersion = True Then
Cmd(9).Visible = False
Fra(1).Visible = False
   End If
   
   
    If SystemOptions.DateOpt = 1 Then
        Txt_DateHigri.Visible = True
    
    End If
        If SystemOptions.TypeContractAutoFromIqar = True Then
       ComResid(0).Enabled = False
       ComResid(1).Enabled = False
    Else
       ComResid(0).Enabled = True
       ComResid(1).Enabled = True
    End If
If SystemOptions.NoCreatJLInRentContract = True Then
TxtVATValue.Visible = True
lbl(108).Visible = True
Else
TxtVATValue.Visible = False
lbl(108).Visible = False
End If
    If mdifrmmain.TransporterMain.Visible = False Then
        lbl(49).Visible = False
        lbl(50).Visible = False
        DCCar.Visible = False
        DCDriver.Visible = False

    End If
    RereivID = 0
If ChekSanNumber(Current_branch, 2) = True Then
TxtNoteSerial1.Enabled = False
Else
TxtNoteSerial1.Enabled = True
End If
'sa If mdifrmmain.MnuProjects.Visible = True Then
'sa XPTab301.TabVisible(1) = True
'sa Else
'sa  XPTab301.TabVisible(1) = False
'sa End If

'sa If mdifrmmain.AssetsMngBase.Visible = True Then
 'sa XPTab301.TabVisible(2) = True
'sa Else
 'sa XPTab301.TabVisible(2) = False
 'sa End If

    ScreenNameArabic = "«·„ﬁ»Ê÷« "
    ScreenNameEnglish = "Cash Receipt Voucher"
    RegisterLogInOut Me.Name, ScreenNameArabic, ScreenNameEnglish, "1", 4
 
    Dim StrSQL As String
    Dim Msg As String
    Set Dcombos = New ClsDataCombos
    Dcombos.GetCostCenter DcCostCenter
    Dcombos.GetSalesRepData Me.DcEmp
    Dcombos.GetCars Me.DCCar
    Dcombos.GetEmployees Me.DCDriver, , True
    Dcombos.GetIqar DcbIqara
    Dcombos.getAkarUnit Me.DcbUnitType
    Dcombos.GetAccountingCodes Me.DcbAccount, True, False
'    StrSQL = "  SELECT code ,account_name FROM markaas_taklefa  WHERE level=3 and NOT(account_no IS NULL)  "
'    fill_combo Me.DcCostCenter, StrSQL

'    Dim Dcombos As ClsDataCombos
'Set Dcombos = New ClsDataCombos


    Set Cmd(0).ButtonImage = mdifrmmain.ImgLstTree.ListImages("New").Picture
    Set Cmd(1).ButtonImage = mdifrmmain.ImgLstTree.ListImages("Edit").Picture
    Set Cmd(2).ButtonImage = mdifrmmain.ImgLstTree.ListImages("save").Picture
    Set Cmd(3).ButtonImage = mdifrmmain.ImgLstTree.ListImages("Undo").Picture
    Set Cmd(4).ButtonImage = mdifrmmain.ImgLstTree.ListImages("Del").Picture
    Set Cmd(5).ButtonImage = mdifrmmain.ImgLstTree.ListImages("Search").Picture
    Set Cmd(6).ButtonImage = mdifrmmain.ImgLstTree.ListImages("Exit").Picture
    Set Cmd(7).ButtonImage = mdifrmmain.ImgLstTree.ListImages("Print").Picture
    Set CmdHelp.ButtonImage = mdifrmmain.ImgLstTree.ListImages("Help").Picture
    Set Cmd(8).ButtonImage = mdifrmmain.ImgLstTree.ListImages("FillData").Picture
    'Resize_Form Me
    AddTip
    DCboCashType.Clear
    DCboCashType.AddItem "„‰ ⁄„Ì·"
    DCboCashType.AddItem "„‰ „Ê—œ"
    DCboCashType.AddItem "„ﬁ«Ê· »«ÿ‰"
    DCboCashType.AddItem "≈Ì—«œ«  ≈Œ—Ï"
    DCboCashType.AddItem "„œ›Ê⁄«  „ﬁœ„Â"
    DCboCashType.AddItem "„‘—Ê⁄"
    DCboCashType.AddItem "„‰ „ÊŸ›"
    DCboCashType.AddItem "„‰ Õ”«»"
   DCboCashType.AddItem "„‰ ⁄ﬁœ"
DCboCashType.AddItem "œ›⁄Â ÕÃ“ "
DCboCashType.AddItem " ’›ÌÂ "
DCboCashType.AddItem "‘ƒÊ‰ ﬁ«‰Ê‰ÌÂ/Â—Ê»"
DCboCashType.AddItem "”⁄Ì Œ«—ÃÌ"
DCboCashType.AddItem "«· ’›Ì«  Ê›Ê« Ì— «·ﬂÂ—»«¡"
With DCboCashType2
.Clear
    .AddItem "„‰ Õ”«»"
    .AddItem "„‰ ⁄ﬁœ"
    .AddItem "œ›⁄Â ÕÃ“ "
    .AddItem " ’›ÌÂ "
    .AddItem "‘ƒÊ‰ ﬁ«‰Ê‰ÌÂ/Â—Ê»"
    .AddItem "”⁄Ì Œ«—ÃÌ"
    .AddItem "«· ’›Ì«  Ê›Ê« Ì— «·ﬂÂ—»«¡"
End With

    With Me.CboPaymentType
        .Clear
        .AddItem "‰ﬁœÌ"
        .AddItem "‘Ìﬂ"
        .AddItem "ÕÊ«·Â »‰ﬂÌÂ"
        .AddItem "  ‘Ìﬂ „Õ’· "
       .AddItem "Õ”«»"
    End With

    Dcombos.GetUsers Me.DCboUserName

'    Dcombos.GetBoxes Me.DcboBox

If SystemOptions.AllowHideAssest = False Then
    Dcombos.GetBoxes Me.DcboBox
    Else
    Dcombos.GetBoxes Me.DcboBox, 0
    End If
    
    Dcombos.GetChequeBox Me.DCChequeBox

    Dcombos.GetBanks Me.DcboBankName
    Dcombos.GetCustomersSuppliers 56, Me.DBCboClientName, False
    Dcombos.GetRevenuesTypes Me.DcboRevenuesTypes
    'Set cSearchDcbo = New clsDCboSearch
    'Set cSearchDcbo.Client = Me.DBCboClientName

    Dcombos.GetAccountingCodes Me.DcboDebitSide
    Dcombos.GetAccountingCodes Me.DcboCreditSide
    Dcombos.GetBranches Me.Dcbranch

    If SystemOptions.usertype <> UserAdminAll Then
        Me.Dcbranch.Enabled = True
    End If

  '  Set rs = New ADODB.Recordset
    'StrSQL = "select * From Notes where NoteType=4 and   displayed is null Order By NoteID"
    StrSQL = "select * From Notes where 1=-1    and   branch_no in(" & Current_branchSql & ")"

   ' If SystemOptions.usertype <> UserAdminAll Or val(Current_branch) <> 0 Then
   '     StrSQL = StrSQL & " AND   branch_no=" & Current_branch
   ' End If
            
   '              If SystemOptions.usertype <> UserAdminAll Then
 '
 '         If SystemOptions.FixedCustomer = 1 Then
 '           StrSQL = StrSQL & " and  UserID = " & user_id
 '            End If
  
       ' Me.dcBranch.Enabled = TRUE
      
      
  '  End If
    
  '  StrSQL = StrSQL & "and  displayed is null Order By NoteID"
    'rs.CursorLocation = adUseClient
    Set rs = New ADODB.Recordset
    rs.Open StrSQL, Cn, adOpenKeyset, adLockOptimistic, adCmdText

    If Not (rs.BOF Or rs.EOF) Then
        rs.MoveLast
    End If

    SetDtpickerDate Me.XPDtbTrans
    SetDtpickerDate Me.DtpChequeDueDate

    With Me.CboTrans
        .Clear
        .AddItem "›« Ê—… „»Ì⁄« "
        .AddItem "„— Ã⁄ „‘ —Ì« "
        .AddItem " ”·Ì„ ’Ì«‰… ·⁄„Ì·"
        .AddItem "Œœ„« "
    End With

    If SystemOptions.UserInterface = EnglishInterface Then
        SetInterface Me
        ChangeLang
    End If

    Msg = "„·ÕÊŸ…:-"
    Msg = Msg & CHR(13) & "≈–« ﬂ«‰  Â–Â «·„ﬁ»Ê÷«   Õ’Ì· ·›« Ê—… „⁄Ì‰…"
    Msg = Msg & "›ÌÃ» ⁄·Ìﬂ «‰  ﬁÊ„ » ÕœÌœ Â–Â «·›« Ê—… "
    Msg = Msg & "Õ Ï Ì „ —»ÿ ⁄„·Ì… «· Õ’Ì· Â–Â „⁄ «·›« Ê—…"
    Me.lbl(11).Caption = Msg
    SetDtpickerDate Me.XPDtbTrans
    ChkTrans.value = Unchecked
    ChkTrans_Click
   XPBtnMove_Click 1
    Me.TxtModFlg.Text = "R"
    WriteInfo
     NoOfDayl
    Dim My_SQL As String
ReLineGrid
    'My_SQL = "  select expanses_account,account_name from projects  where not (account_no is null)"
    My_SQL = "  select id,Project_name from projects where not(REVENUE_account is null)" '
    fill_combo dcproject, My_SQL
 InserTypeAmount
    If OPEN_NEW_SCREEN = True Then
        Cmd_Click (0)
       
    End If

    Exit Sub
ErrTrap:
End Sub

Private Sub Form_Unload(Cancel As Integer)
    On Error GoTo ErrTrap
    RegisterLogInOut Me.Name, ScreenNameArabic, ScreenNameEnglish, 4

    If rs.State = adStateOpen Then
        If Not (rs.EOF Or rs.BOF) Then
            If rs.EditMode <> adEditNone Then
                rs.CancelUpdate
            End If
        End If

        rs.Close
        Set rs = Nothing
    End If

    Set TTP = Nothing
    'Set EmpReport = Nothing
    Exit Sub
ErrTrap:
End Sub

Private Sub FrmPriodDate_Change()
   If Me.TxtModFlg.Text <> "R" Then
     
    FrmPriodDateH.value = ToHijriDate(FrmPriodDate.value)
    
End If
End Sub

Private Sub FrmPriodDateH_LostFocus()
 If Me.TxtModFlg.Text <> "R" Then
             FrmPriodDate.value = ToGregorianDate(FrmPriodDateH.value)
        End If
End Sub

Private Sub Grid3_AfterEdit(ByVal Row As Long, ByVal Col As Long)
Dim StrAccountCode  As String
Dim LngRow As Long
 With Grid3

        Select Case .ColKey(Col)
 
 Case "CommisionTypes"
 StrAccountCode = .ComboData
                LngRow = .FindRow(StrAccountCode, .FixedRows, .ColIndex("CommisionTypesid"), False, True)
                .TextMatrix(Row, .ColIndex("CommisionTypesid")) = StrAccountCode
     

End Select

End With

payed
 ReLineGrid
End Sub
Sub payed()
Dim i As Integer
With Grid3
For i = 1 To .rows - 1
If .cell(flexcpChecked, i, .ColIndex("Payed")) = flexChecked And .cell(flexcpChecked, i, .ColIndex("Select")) = flexChecked Then
.TextMatrix(i, .ColIndex("VATPayed")) = val(.TextMatrix(i, .ColIndex("VATValue"))) - val(.TextMatrix(i, .ColIndex("ActVAT")))
.TextMatrix(i, .ColIndex("RentValuePayed")) = val(.TextMatrix(i, .ColIndex("RentValue"))) - val(.TextMatrix(i, .ColIndex("ActRent")))
.TextMatrix(i, .ColIndex("CommissionsPayed")) = val(.TextMatrix(i, .ColIndex("Commissions"))) - val(.TextMatrix(i, .ColIndex("ActComm")))
.TextMatrix(i, .ColIndex("InsurancePayed")) = val(.TextMatrix(i, .ColIndex("Insurance"))) - val(.TextMatrix(i, .ColIndex("ActInsu")))
.TextMatrix(i, .ColIndex("WaterPayed")) = val(.TextMatrix(i, .ColIndex("Water"))) - val(.TextMatrix(i, .ColIndex("ActWater")))
.TextMatrix(i, .ColIndex("OldValuePayed")) = val(.TextMatrix(i, .ColIndex("OldValue"))) - val(.TextMatrix(i, .ColIndex("ActOldValue")))
.TextMatrix(i, .ColIndex("ElectricPayed")) = val(.TextMatrix(i, .ColIndex("Electric"))) - val(.TextMatrix(i, .ColIndex("ActElec")))
.TextMatrix(i, .ColIndex("TelandNetPayed")) = val(.TextMatrix(i, .ColIndex("TelandNet"))) - val(.TextMatrix(i, .ColIndex("ActService")))
End If
Next i
End With
End Sub
 Function PayVAT() As Boolean
 Dim i As Integer
 PayVAT = True
 With Me.Grid3
 For i = 1 To .rows - 1
        If .cell(flexcpChecked, i, .ColIndex("Select")) = flexChecked And val(.TextMatrix(i, .ColIndex("VATValue"))) <> 0 And Round(val(.TextMatrix(i, .ColIndex("VATValue"))), 2) > (Round(val(.TextMatrix(i, .ColIndex("ActVAT"))), 2) + Round(val(.TextMatrix(i, .ColIndex("VATPayed"))), 2)) Then
        PayVAT = False
                  Exit Function
         End If
       Next i
 End With
 End Function
Function ReLineGrid(Optional idd As Integer = 0)
 Dim i As Integer
 With Me.Grid3
 Dim InstalNo As Double
 InstalNo = 0
 If Me.TxtModFlg.Text <> "R" And TxtModFlg.Text <> "" Then
    lblremain.Caption = 0
 ElseIf Me.TxtModFlg.Text = "R" Then
    Exit Function
 End If
 Dim RamainValue As Double
 RamainValue = 0
 If Grid3.rows > 1 Then
If idd = 0 Then
        For i = .FixedRows To .rows - 1
            If .cell(flexcpChecked, i, .ColIndex("Select")) = flexChecked Then
                  
                  If val(.TextMatrix(i, .ColIndex("InstallNo"))) <> 0 Then
                  InstalNo = val(.TextMatrix(i, .ColIndex("InstallNo")))
                  End If
                  If Me.TxtModFlg.Text <> "R" And TxtModFlg.Text <> "" Then
                  If Round(val(.TextMatrix(i, .ColIndex("VATValue"))), 2) < (Round(val(.TextMatrix(i, .ColIndex("ActVAT"))), 2) + Round(val(.TextMatrix(i, .ColIndex("VATPayed"))), 2)) Then
                  If SystemOptions.UserInterface = ArabicInterface Then
                  MsgBox "·«Ì„ﬂ‰ «‰   Ã«Ê“ «·ﬁÌ„… «·„÷«›… «·«’·Ì…"
                  Else
                  MsgBox "The Value Larger than Toatal VAT"
                  End If
                  .TextMatrix(i, .ColIndex("VATPayed")) = 0
                  Exit Function
                  End If
                  End If
                  
                  If Me.TxtModFlg.Text <> "R" And TxtModFlg.Text <> "" Then
                  If Round(val(.TextMatrix(i, .ColIndex("RentValue"))), 2) < (Round(val(.TextMatrix(i, .ColIndex("ActRent"))), 2) + Round(val(.TextMatrix(i, .ColIndex("RentValuePayed"))), 2)) Then
                  If SystemOptions.UserInterface = ArabicInterface Then
                  MsgBox "·«Ì„ﬂ‰ «‰   Ã«Ê“ ﬁÌ„… «·«ÌÃ«— «·«’·Ì…"
                  Else
                  MsgBox "The Value Larger than Toatal Rent"
                  End If
                  .TextMatrix(i, .ColIndex("RentValuePayed")) = 0
                  Exit Function
                  End If
                  End If
                  If Me.TxtModFlg.Text <> "R" And TxtModFlg.Text <> "" Then
                     If Round(val(.TextMatrix(i, .ColIndex("Commissions"))), 2) < (Round(val(.TextMatrix(i, .ColIndex("ActComm"))), 2) + Round(val(.TextMatrix(i, .ColIndex("CommissionsPayed"))), 2)) Then
                  If SystemOptions.UserInterface = ArabicInterface Then
                  MsgBox "·«Ì„ﬂ‰ «‰   Ã«Ê“ ﬁÌ„… «·”⁄Ì «·«’·Ì…"
                  Else
                  MsgBox "The Value Larger than Toatal Commissions"
                  End If
                  .TextMatrix(i, .ColIndex("CommissionsPayed")) = 0
                  Exit Function
                  End If
                  End If
                  If Me.TxtModFlg.Text <> "R" And TxtModFlg.Text <> "" Then
                  If Round(val(.TextMatrix(i, .ColIndex("Insurance"))), 2) < (Round(val(.TextMatrix(i, .ColIndex("ActInsu"))), 2) + Round(val(.TextMatrix(i, .ColIndex("InsurancePayed"))), 2)) Then
                  If SystemOptions.UserInterface = ArabicInterface Then
                  MsgBox "·«Ì„ﬂ‰ «‰   Ã«Ê“ ﬁÌ„… «· √„Ì‰ «·«’·Ì…"
                  Else
                  MsgBox "The Value Larger than Toatal Insurance"
                  End If
                  .TextMatrix(i, .ColIndex("InsurancePayed")) = 0
                  Exit Function
                  End If
                     If Round(val(.TextMatrix(i, .ColIndex("Water"))), 2) < (Round(val(.TextMatrix(i, .ColIndex("ActWater"))), 2) + Round(val(.TextMatrix(i, .ColIndex("WaterPayed"))), 2)) Then
                  If SystemOptions.UserInterface = ArabicInterface Then
                  MsgBox "·«Ì„ﬂ‰ «‰   Ã«Ê“ ﬁÌ„… «·„Ì«Â «·«’·Ì…"
                  Else
                  MsgBox "The Value Larger than Toatal Water"
                  End If
                  .TextMatrix(i, .ColIndex("WaterPayed")) = 0
                  Exit Function
                  End If
                  End If
                  If Me.TxtModFlg.Text <> "R" And TxtModFlg.Text <> "" Then
                If Round(val(.TextMatrix(i, .ColIndex("Electric"))), 2) < (Round(val(.TextMatrix(i, .ColIndex("ActElec"))), 2) + Round(val(.TextMatrix(i, .ColIndex("ElectricPayed"))), 2)) Then
                  If SystemOptions.UserInterface = ArabicInterface Then
                  MsgBox "·«Ì„ﬂ‰ «‰   Ã«Ê“ ﬁÌ„… «·ﬂÂ—»«¡ «·«’·Ì…"
                  Else
                  MsgBox "The Value Larger than Toatal Electric"
                  End If
                  .TextMatrix(i, .ColIndex("ElectricPayed")) = 0
                  Exit Function
                  End If
                  If Round(val(.TextMatrix(i, .ColIndex("OldValue"))), 2) < (Round(val(.TextMatrix(i, .ColIndex("ActOldValue"))), 2) + Round(val(.TextMatrix(i, .ColIndex("OldValuePayed"))), 2)) Then
                  If SystemOptions.UserInterface = ArabicInterface Then
                  MsgBox "·«Ì„ﬂ‰ «‰   Ã«Ê“ «·—’Ìœ «·”«»ﬁ"
                  Else
                  MsgBox "The Value Larger than  Previous Balance"
                  End If
                  .TextMatrix(i, .ColIndex("OldValuePayed")) = 0
                  Exit Function
                  End If
               If Me.TxtModFlg.Text <> "R" And TxtModFlg.Text <> "" Then
                 If Round(val(.TextMatrix(i, .ColIndex("TelandNet"))), 2) < (Round(val(.TextMatrix(i, .ColIndex("ActService"))), 2) + Round(val(.TextMatrix(i, .ColIndex("TelandNetPayed"))), 2)) Then
                  If SystemOptions.UserInterface = ArabicInterface Then
                  MsgBox "·«Ì„ﬂ‰ «‰   Ã«Ê“ ﬁÌ„… «·Œœ„«  «·«’·Ì…"
                  Else
                  MsgBox "The Value Larger than Toatal Service"
                  End If
                  .TextMatrix(i, .ColIndex("TelandNetPayed")) = 0
                
                  Exit Function
                  End If
                  End If
                  End If
              End If
        
        Next i
        End If
 End If
 Dim InstalNo1 As Double
  For i = .FixedRows To .rows - 1
            If .cell(flexcpChecked, i, .ColIndex("Select")) = flexChecked Then
                  If val(.TextMatrix(i, .ColIndex("InstallNo"))) <> 0 Then
                  InstalNo1 = val(.TextMatrix(i, .ColIndex("InstallNo")))
                  GoTo l
                  End If
                End If
            Next i
End With
l:
Dim PerioDID As Integer
Dim Period As Integer
GetBeforInstalDate InstalNo1
GetInstalPeriod PerioDID, Period
If CheckmaxInstal(InstalNo) = False Then
GetInstalDate InstalNo + 1
Else
GetInstalMaxDate InstalNo, PerioDID, Period
End If




Dim IntCounter As Interlaced
Dim SUM As Double
  Dim totalPayed As Double
    totalPayed = 0
  With Me.Grid3
 Dim bol As Boolean
 bol = False
lblrent.Caption = 0
lblcomision.Caption = 0
'lblremain.Caption = 0
Dim totalPayed1 As Double


totalPayed = 0
Dim Totalss As Double
Totalss = 0
        For i = .FixedRows To .rows - 1

            If .cell(flexcpChecked, i, .ColIndex("Select")) = flexChecked Then
         totalPayed = 0
         totalPayed1 = 0
         
                    totalPayed = totalPayed + val(.TextMatrix(i, .ColIndex("ActRent")))
                    totalPayed = totalPayed + val(.TextMatrix(i, .ColIndex("ActComm")))
                    totalPayed = totalPayed + val(.TextMatrix(i, .ColIndex("ActInsu")))
                    totalPayed = totalPayed + val(.TextMatrix(i, .ColIndex("ActWater")))
                    totalPayed = totalPayed + val(.TextMatrix(i, .ColIndex("ActElec")))
                    totalPayed = totalPayed + val(.TextMatrix(i, .ColIndex("ActService")))
                    totalPayed = totalPayed + val(.TextMatrix(i, .ColIndex("ActOldValue")))
                    totalPayed = totalPayed + val(.TextMatrix(i, .ColIndex("ActVAT")))
                    'xxx
                   totalPayed1 = totalPayed1 + val(.TextMatrix(i, .ColIndex("RentValuePayed")))
                    totalPayed1 = totalPayed1 + val(.TextMatrix(i, .ColIndex("CommissionsPayed")))
                    totalPayed1 = totalPayed1 + val(.TextMatrix(i, .ColIndex("InsurancePayed")))
                    totalPayed1 = totalPayed1 + val(.TextMatrix(i, .ColIndex("WaterPayed")))
                    totalPayed1 = totalPayed1 + val(.TextMatrix(i, .ColIndex("ElectricPayed")))
                    totalPayed1 = totalPayed1 + val(.TextMatrix(i, .ColIndex("TelandNetPayed")))
                    totalPayed1 = totalPayed1 + val(.TextMatrix(i, .ColIndex("OldValuePayed")))
                    totalPayed1 = totalPayed1 + val(.TextMatrix(i, .ColIndex("VATPayed")))

 

        
                   .TextMatrix(i, .ColIndex("ActualTotal")) = totalPayed
             Totalss = Round(val(.TextMatrix(i, .ColIndex("total"))), 2)
            ' If Me.TxtModFlg.Text <> "E" Then
                      .TextMatrix(i, .ColIndex("result")) = Round(Totalss - Round(totalPayed, 2) - Round(totalPayed1, 2), 2)
            '     Else
            '     .TextMatrix(I, .ColIndex("result")) = Round(Totalss - Round(totalPayed1, 2), 2)
            'End If
                      If val(.TextMatrix(i, .ColIndex("total"))) > 0 Then
                    .TextMatrix(i, .ColIndex("resultpercentage")) = Round(totalPayed / val(.TextMatrix(i, .ColIndex("total"))) * 100, 2)
                    End If
                    lblrent.Caption = val(lblrent.Caption) + val(.TextMatrix(i, .ColIndex("RentValuePayed")))
                    lblcomision.Caption = val(lblcomision.Caption) + val(.TextMatrix(i, .ColIndex("CommissionsPayed")))
                    lblservice.Caption = val(lblservice.Caption) + val(.TextMatrix(i, .ColIndex("TelandNetPayed")))
                  '  lblremain.Caption = val(lblremain.Caption) + val(.TextMatrix(i, .ColIndex("Result")))
                  RamainValue = RamainValue + val(.TextMatrix(i, .ColIndex("Result")))
                 Else
                 
        End If

        Next i
If Me.TxtModFlg.Text <> "R" And TxtModFlg.Text <> "" Then
lblremain.Caption = RamainValue
End If
totalPayed = 0
        For i = .FixedRows To .rows - 1

            If .cell(flexcpChecked, i, .ColIndex("Select")) = flexChecked Then
       
                    totalPayed = totalPayed + val(.TextMatrix(i, .ColIndex("RentValuePayed")))
                    totalPayed = totalPayed + val(.TextMatrix(i, .ColIndex("CommissionsPayed")))
                    totalPayed = totalPayed + val(.TextMatrix(i, .ColIndex("InsurancePayed")))
                    totalPayed = totalPayed + val(.TextMatrix(i, .ColIndex("WaterPayed")))
                    totalPayed = totalPayed + val(.TextMatrix(i, .ColIndex("ElectricPayed")))
                    totalPayed = totalPayed + val(.TextMatrix(i, .ColIndex("TelandNetPayed")))
                    totalPayed = totalPayed + val(.TextMatrix(i, .ColIndex("OldValuePayed")))
                    'If SystemOptions.NoCreatJLInRentContract = False Then
                    totalPayed = totalPayed + val(.TextMatrix(i, .ColIndex("VATPayed")))
                    'End If
               '   totalPayed = totalPayed + val(.TextMatrix(i, .ColIndex("RentArbon")))
               '   totalPayed = totalPayed + val(.TextMatrix(i, .ColIndex("CommissionsArbon")))
               '   totalPayed = totalPayed + val(.TextMatrix(i, .ColIndex("InsuranceArbon")))
               '   totalPayed = totalPayed + val(.TextMatrix(i, .ColIndex("WaterArbon")))
               '   totalPayed = totalPayed + val(.TextMatrix(i, .ColIndex("ElectricArbon")))
               '   totalPayed = totalPayed + val(.TextMatrix(i, .ColIndex("ServiceArbon")))
                  
                 Else
                 
        End If

        Next i
        
    End With
 SUM = 0
     IntCounter = 0
    With VSFlexGrid1

        For i = .FixedRows To .rows - 1
bol = True
            If .TextMatrix(i, .ColIndex("empname")) <> "" Then
                IntCounter = IntCounter + 1
                
                .TextMatrix(i, .ColIndex("Ser")) = IntCounter
          '  If val(.TextMatrix(i, .ColIndex("rate"))) > 0 Then
          '  .TextMatrix(i, .ColIndex("rate")) = (val(.TextMatrix(i, .ColIndex("rate"))) / 100) * val(.TextMatrix(i, .ColIndex("rate")))
          '  End If
                SUM = SUM + val(.TextMatrix(i, .ColIndex("rate")))
                If SUM > 100 Then
                .TextMatrix(i, .ColIndex("rate")) = 0
                MsgBox "·«Ì„ﬂ‰ «‰ÌﬂÊ‰ „Ã„Ê⁄ «·‰”» «ﬂ»— „‰ 100%"
                Exit Function
                End If
            End If

        Next i

    End With
     IntCounter = 0
    With VSFlexGrid2
SUM = 0
        For i = .FixedRows To .rows - 1

            If .TextMatrix(i, .ColIndex("empname")) <> "" Then
                IntCounter = IntCounter + 1
                .TextMatrix(i, .ColIndex("Ser")) = IntCounter
             '   .TextMatrix(i, .ColIndex("values")) = (val(LblTotal.Caption) / 100) * val(.TextMatrix(i, .ColIndex("rate")))
                                SUM = SUM + val(.TextMatrix(i, .ColIndex("rate")))
                If SUM > 100 Then
                .TextMatrix(i, .ColIndex("rate")) = 0
                MsgBox "·«Ì„ﬂ‰ «‰ÌﬂÊ‰ „Ã„Ê⁄ «·‰”» «ﬂ»— „‰ 100%"
                Exit Function
                End If
                
              
            End If

        Next i

    End With
  
  
    LblTotal.Caption = val(lblrent.Caption) + val(lblcomision.Caption) + val(lblservice.Caption)
   If val(DCboCashType.ListIndex) = 8 Then
      Me.XPTxtVal.Text = totalPayed
      If ComResid(1).value = True Then
      If DCboCashType2.ListIndex = 0 Then
      
      Exit Function
      End If
   '   TxtVATValue.Text = totalPayed * 5 / 100
   '   TxtVATValue = netVatPayed
        
  ' salimhere
       Dim total_value As Double
        Dim Percetage As Double
   PercentgValueAddedAccount_Transec XPDtbTrans.value, 8, 1, , Percetage
       total_value = 0
    Dim mVATPayed As Double
    total_value = 0
     With Grid3
         For i = .FixedRows To .rows - 1
            If .cell(flexcpChecked, i, .ColIndex("Select")) = flexChecked Then
                        total_value = total_value + Round(.TextMatrix(i, .ColIndex("RentValuePayed")), 2) + Round(.TextMatrix(i, .ColIndex("WaterPayed")), 2) + Round(.TextMatrix(i, .ColIndex("ElectricPayed")), 2) '+ Round(.TextMatrix(i, .ColIndex("TelandNetPayed")), 2)
                        mVATPayed = mVATPayed + val(.TextMatrix(i, .ColIndex("VATPayed")))
             End If
       
               
        Next i
        
        End With

'        If ComResid(1).value = True Then
'            TxtVATValue = mVATPayed ' total_value * Percetage / 100
'        Else
'            TxtVATValue = 0
'        End If
        'salimhere
   
        Else
            TxtVATValue.Text = 0
      End If
      End If
TxtVATValue.Text = 0

End Function

Private Sub Grid3_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    With Grid3
 
    'If .ColKey(Col) <> "Due_DateH" And .ColKey(Col) <> "Status" Then
   If Me.TxtModFlg = "R" Then Exit Sub
   If .ColKey(Col) <> "Select" Then
            If .cell(flexcpChecked, Row, .ColIndex("Select")) = flexUnchecked Then
            ReLineGrid
            Cancel = True
             
             Exit Sub
            
            End If
            
End If
 
   Select Case .ColKey(Col)
          Case "Select"
      If SystemOptions.AllowSkipPayment = False Then
       If Row > 1 Then
        If val(.TextMatrix(Row - 1, .ColIndex("Result"))) > 0 Then
          MsgBox "·«Ì„ﬂ‰  Ã«Ê“ «·œ›⁄… «·”«»ﬁ…"
          .cell(flexcpChecked, Row, .ColIndex("Select")) = flexUnchecked
          Exit Sub
        End If
       End If
      End If
                  If .cell(flexcpChecked, Row, .ColIndex("Select")) = flexUnchecked Then
           ReLineGrid
             
             
             Exit Sub
            
            End If
            
 Case "RentValue"
 Cancel = True
 
  Case "VATValue"
 Cancel = True
 
 Case "Commissions"
 Cancel = True
  Case "Insurance"
 Cancel = True
   Case "Water"
 Cancel = True
    Case "Electric"
 Cancel = True
     Case "TelandNet"
 Cancel = True
   Case "OldValue"
 Cancel = True
    Case "total"
 Cancel = True
     Case "RentValuePayed"
 .ComboList = ""
      Case "CommissionsPayed"
 .ComboList = ""
      Case "InsurancePayed"
 .ComboList = ""
      Case "WaterPayed"
 .ComboList = ""
      Case "ElectricPayed"
 .ComboList = ""
      Case "TelandNetPayed"
 .ComboList = ""
      Case "OldValuePayed"
 .ComboList = ""
       Case "OldValuePayed"
 .ComboList = ""
      Case "ActualTotal"
 Cancel = True
       Case "ActualTotal"
 Cancel = True
       Case "ActOldValue"
 Cancel = True
       Case "ActComm"
 Cancel = True
       Case "ActInsu"
 Cancel = True
       Case "ActWater"
 Cancel = True
        Case "ActElec"
 Cancel = True
         Case "Result"
 Cancel = True
          Case "ResultPercentage"
 Cancel = True
 Case "ActService"
 Cancel = True
 End Select
    End With
End Sub

Private Sub Grid3_StartEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    With Grid3
Dim StrSQL  As String
Dim StrComboList As String
Dim rs2 As New ADODB.Recordset

        Select Case .ColKey(Col)
 Case "CommisionTypes"
 
                StrSQL = "select * from TblCommisionTypes"
                rs2.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

                If SystemOptions.UserInterface = ArabicInterface Then
                    StrComboList = .BuildComboList(rs, "name", "id")
                Else
                    StrComboList = .BuildComboList(rs, "namee", "id")
                End If
       
                If StrComboList <> "" Then
                    StrComboList = "|" & StrComboList
                End If
                 .ComboList = StrComboList

   End Select
   End With
End Sub

Private Sub ImgFavorites_Click()
AddTofaforites Me.Name, Me.Caption, Me.Caption
End Sub

Private Sub ISButton1_Click()
'Load FrmIqarContractSearch vbmodal
If val(DCboCashType.ListIndex) = 13 Then
 FrmIqarWaiverSet.m_RetrunType = 808
Load FrmIqarWaiverSet
FrmIqarWaiverSet.m_RetrunType = 808
FrmIqarWaiverSet.show
Else
FrmIqarContractSearch.m_RetrunType = 5
 FrmIqarContractSearch.show vbModal
End If
End Sub

Private Sub ISButton3_Click()
FrmIqarWaiverSet.m_RetrunType = 5
 FrmIqarWaiverSet.show vbModal
End Sub

Private Sub Label10_Click()
Frame13.Visible = False
End Sub

Private Sub Label40_Click()
Frame11.Visible = False
End Sub

Private Sub lbl_MouseMove(Index As Integer, _
                          Button As Integer, _
                          Shift As Integer, _
                          X As Single, _
                          Y As Single)

    If Index = 18 Then
        If SystemOptions.UserInterface = ArabicInterface Then
            Me.lbl(18).ToolTipText = "ﬁÌ„… „»·€ «·„ﬁ»Ê÷« :" & lbl(18).Caption
        ElseIf SystemOptions.UserInterface = EnglishInterface Then
            Me.lbl(18).ToolTipText = "Notes Recivable Value:" & lbl(18).Caption
        End If
    End If

End Sub

Private Sub LblLink_Click()
  If SystemOptions.SpecialVersion = True Then
        Exit Sub
End If
    
    Dim FirstPeriod As Date
    getFirstPeriodDateInthisYear FirstPeriod
    ShowReport DcboCreditSide.BoundText, DcboCreditSide.Text, FirstPeriod, Date

End Sub

Private Sub LblLink_MouseMove(Button As Integer, _
                              Shift As Integer, _
                              X As Single, _
                              Y As Single)
 
    If SystemOptions.UserInterface = ArabicInterface Then
        LblLink.ToolTipText = "—’Ìœ «·ÿ—› «·œ«∆‰:" & WriteNo(Balance, 0, True)
    ElseIf SystemOptions.UserInterface = EnglishInterface Then
        LblLink.ToolTipText = "Credit Balance:" & WriteNo(Balance, 0, True)
    End If
 
End Sub

Private Sub Option1_Click()

    If Option2.value = True Then
        ALLButton3.Enabled = True
    Else
        ALLButton3.Enabled = False
    End If

    If Option6.value = True Then
        ALLButton4.Enabled = True
    Else
        ALLButton4.Enabled = False
    End If
DBCboClientName_Change
End Sub

Private Sub Option2_Click()

    If Option2.value = True Then
        ALLButton3.Enabled = True
    Else
        ALLButton3.Enabled = False
    End If

    If Option6.value = True Then
        ALLButton4.Enabled = True
    Else
        ALLButton4.Enabled = False
    End If
DBCboClientName_Change
End Sub

Private Sub Option3_Click()

    If Option2.value = True Then
        ALLButton3.Enabled = True
    Else
        ALLButton3.Enabled = False
    End If

    If Option6.value = True Then
        ALLButton4.Enabled = True
    Else
        ALLButton4.Enabled = False
    End If
DBCboClientName_Change
End Sub

Private Sub Option4_Click()

    If DCboCashType.ListIndex <> 5 Then Exit Sub
 DBCboClientName_Change

End Sub

Private Sub Option5_Click()

    If DCboCashType.ListIndex <> 5 Then Exit Sub
 DBCboClientName_Change

End Sub

Private Sub Option6_Click()

    If Option6.value = True Then
        ALLButton4.Enabled = True
    Else
        ALLButton4.Enabled = False
    End If

    If Option6.value = True Then
        ALLButton4.Enabled = True
    Else
        ALLButton4.Enabled = False
    End If

End Sub

Private Sub ToPriodDate_Change()
   If Me.TxtModFlg.Text <> "R" Then
     
    ToPriodDateH.value = ToHijriDate(ToPriodDate.value)
    
End If
End Sub

Private Sub ToPriodDateH_LostFocus()
 If Me.TxtModFlg.Text <> "R" Then
             
             ToPriodDate.value = ToGregorianDate(ToPriodDateH.value)

               
        End If
End Sub

Private Sub TxtAccount_KeyPress(KeyAscii As Integer)
DcbAccount.BoundText = getAccountSerial_Code("Account_Code", "Account_Serial", TxtAccount.Text)
End Sub

Private Sub TxtCommission_Change()
If Me.TxtModFlg <> "R" Then
txtTotal1.Text = val(Txtcommission.Text) + val(TxtWater.Text) + val(txtinstrunce.Text)
If val(XPTxtVal.Text) >= (val(Txtcommission.Text) - val(TxtCommissionOut.Text)) Then
txtComisin.Text = val(Txtcommission.Text) - val(TxtCommissionOut.Text)
Else
txtComisin.Text = val(XPTxtVal.Text)
End If
End If
End Sub

Private Sub TxtCommissionOut_Change()
If Me.TxtModFlg <> "R" Then
txtTotal1.Text = val(Txtcommission.Text) + val(TxtWater.Text) + val(txtinstrunce.Text)
If val(XPTxtVal.Text) >= (val(Txtcommission.Text) - val(TxtCommissionOut.Text)) Then
txtComisin.Text = val(Txtcommission.Text) - val(TxtCommissionOut.Text)
Else
txtComisin.Text = val(XPTxtVal.Text)
End If
End If
End Sub

Function GetIDNo(Optional NoteSerial1 As String, Optional ID As Double) As Boolean
Dim sql As String
Dim Rs3 As ADODB.Recordset
Set Rs3 = New ADODB.Recordset
sql = " SELECT      ID,FlgPayed From dbo.TblOtheExpensAqar where (FlgPayed is null) and NoteSerial1='" & NoteSerial1 & "' and BranchID=" & val(Me.Dcbranch.BoundText) & " "
Rs3.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
If Rs3.RecordCount > 0 Then
ID = IIf(IsNull(Rs3("ID").value), 0, Rs3("ID").value)
   
        
GetIDNo = True
Else
GetIDNo = False
ID = 0
End If
End Function
Sub Calculte()
If Me.TxtModFlg.Text <> "R" Then
TxtTotal23.Text = val(TxtMaintOther3.Text) + val(TxtWindows3.Text) + val(TxtMaintDoors3.Text) + val(TxtMaintenance3.Text) + val(TxtRemainRent3.Text) + val(TxtMaintCondition3.Text) + val(TxtMaintClean3.Text) + val(TxtPaints3.Text) + val(TxtMaintkitchen3.Text) + val(TxtElectricity13.Text)
txtTotal.Text = val(TxtMaintenance.Text) + val(txtRemainRent.Text) + val(TxtMaintCondition.Text) + val(TxtMaintClean.Text)
txtTotal.Text = val(txtTotal.Text) + val(TxtPaints.Text) + val(TxtMaintkitchen.Text) + val(TxtElectricity1.Text)
txtTotal.Text = val(txtTotal.Text) + val(TxtMaintDoors.Text) + val(TxtWindows.Text) + val(TxtMaintOther.Text)
txtNet.Text = val(TxtRemMaintenance.Text) + val(TxtRemRemainRent.Text) + val(TxtRemMaintCondition.Text) + val(TxtRemMaintClean.Text)
txtNet.Text = val(txtNet.Text) + val(TxtRemPaints.Text) + val(TxtRemMaintkitchen.Text) + val(TxtRemElectricity.Text)
txtNet.Text = val(txtNet.Text) + val(TxtRemMaintDoors.Text) + val(TxtRemWindows.Text) + val(TxtRemMaintOther.Text)
If RdTypeTrans(0).value Then
XPTxtVal.Text = val(txtTotal.Text)
Else
XPTxtVal.Text = val(Me.TxtPrice.Text)
End If
End If
End Sub
Function CheckValue() As Boolean
If val(txtTotal.Text) < val(Me.TxtTotal22.Text) + val(Me.TxtTotal23.Text) Then
CheckValue = True
Else
CheckValue = False
End If
End Function
Sub FillFIlBill(Optional ID As Double = 0)
Frame13.Visible = False
Dim sql As String
Dim Rs3 As ADODB.Recordset
Set Rs3 = New ADODB.Recordset
sql = "Select * from TblOtheExpensAqar where id=" & ID & ""
Rs3.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
If Rs3.RecordCount > 0 Then
DBCboClientName.BoundText = IIf(IsNull(Rs3("CusID").value), 0, Rs3("CusID").value)
DcbIqara.BoundText = IIf(IsNull(Rs3("AqarID").value), 0, Rs3("AqarID").value)
DcbUnitType.BoundText = IIf(IsNull(Rs3("UnitTypID").value), 0, Rs3("UnitTypID").value)
DcbUnitNo.BoundText = IIf(IsNull(Rs3("UnitID").value), 0, Rs3("UnitID").value)
If Not IsNull(Rs3("TypID").value) Then
If (Rs3("TypID").value) = 1 Then
RdTypeTrans(1).value = True
Else
RdTypeTrans(0).value = True
End If
End If
If RdTypeTrans(0).value = True Then
     Frame11.Visible = True
txtInsurance.Text = IIf(IsNull(Rs3("Insurance").value), 0, Rs3("Insurance").value)
txtDiscount.Text = IIf(IsNull(Rs3("Discount").value), 0, Rs3("Discount").value)
TxtMaintenance2.Text = IIf(IsNull(Rs3("Maintenance").value), 0, Rs3("Maintenance").value)
TxtRemainRent2.Text = IIf(IsNull(Rs3("RemainRent").value), 0, Rs3("RemainRent").value)
TxtMaintCondition2.Text = IIf(IsNull(Rs3("MaintCondition").value), 0, Rs3("MaintCondition").value)
TxtMaintClean2.Text = IIf(IsNull(Rs3("MaintClean").value), 0, Rs3("MaintClean").value)
TxtPaints2.Text = IIf(IsNull(Rs3("Paints").value), 0, Rs3("Paints").value)
TxtMaintkitchen2.Text = IIf(IsNull(Rs3("MaintKitchen").value), 0, Rs3("MaintKitchen").value)
TxtElectricity12.Text = IIf(IsNull(Rs3("Electricity").value), 0, Rs3("Electricity").value)
TxtMaintDoors2.Text = IIf(IsNull(Rs3("MaintDoors").value), 0, Rs3("MaintDoors").value)
TxtWindows2.Text = IIf(IsNull(Rs3("Windows").value), 0, Rs3("Windows").value)
TxtMaintOther2.Text = IIf(IsNull(Rs3("MaintOther").value), 0, Rs3("MaintOther").value)
TxtTotal22.Text = IIf(IsNull(Rs3("Total").value), 0, Rs3("Total").value)
'TxtTotalAftreIns2.Text = IIf(IsNull(RS3("TotalAfterIns").value), 0, RS3("TotalAfterIns").value)
'TxtNet2.Text = IIf(IsNull(RS3("Net").value), 0, RS3("Net").value)
     Else
     Frame11.Visible = False
     Frame13.Visible = True
     TxtPrice2.Text = IIf(IsNull(Rs3("Valuee").value), 0, Rs3("Valuee").value)
     End If
End If
End Sub
Private Sub TxtContNo_Change()
If Me.DCboCashType.ListIndex = 13 And (Me.TxtModFlg.Text = "N" Or Me.TxtModFlg.Text = "E") Then
If GetIDNo(txtContractNo.Text) = True Then
FillFIlBill val(TxtContNo.Text)
GetTotalPayedElect
Calculte

End If
ElseIf Me.DCboCashType.ListIndex = 8 And (Me.TxtModFlg.Text = "N" Or Me.TxtModFlg.Text = "E") Then
       ' FillGridWithDataContract val(TXTContNo.text)
      FillGridWithDataContract txtContractNo, val(XPTxtID.Text)
 FillGridWithDatSales val(TxtContNo.Text)
    End If
End Sub

Private Sub TxtContractNo_Change()
Dim ID As Double

If val(DCboCashType.ListIndex) = 13 And Me.TxtModFlg.Text <> "R" Then
ClearText

If GetIDNo(txtContractNo.Text, ID) = True Then
TxtContNo.Text = ID
Else
If SystemOptions.UserInterface = ArabicInterface Then
MsgBox " «ﬂœ „‰ —ﬁ„ «·Õ—ﬂ… «Ê ﬁœÌﬂÊ‰  „ ”œ«œÂ« „‰ ﬁ»·"
End If
ClearText
Calculte
Exit Sub
End If
Else
FillGridWithData1 val(TxtContNo.Text)
lbl(66).Caption = Me.txtContractNo.Text
DcbIqara_Click (0)
End If
End Sub
Sub ClearText()
TxtMaintenance2.Text = 0
TxtRemainRent2.Text = 0
TxtMaintCondition2.Text = 0
TxtMaintClean2.Text = 0
TxtPaints2.Text = 0
TxtMaintkitchen2.Text = 0
TxtElectricity12.Text = 0
TxtMaintDoors2.Text = 0
TxtWindows2.Text = 0
TxtMaintOther2.Text = 0
txtNet.Text = 0
txtTotal.Text = 0
TxtTotal23.Text = 0
TxtMaintenance3.Text = 0
TxtRemainRent3.Text = 0
TxtMaintCondition3.Text = 0
TxtMaintClean3.Text = 0
TxtPaints3.Text = 0
TxtMaintkitchen3.Text = 0
TxtElectricity13.Text = 0
TxtMaintDoors3.Text = 0
TxtWindows3.Text = 0
TxtMaintOther3.Text = 0
TxtTotal22.Text = 0

TxtMaintenance.Text = 0
txtRemainRent.Text = 0
TxtMaintCondition.Text = 0
TxtMaintClean.Text = 0
TxtPaints.Text = 0
TxtMaintkitchen.Text = 0
TxtElectricity1.Text = 0
TxtMaintDoors.Text = 0
TxtWindows.Text = 0
TxtMaintOther.Text = 0

TxtRemMaintenance.Text = 0
TxtRemRemainRent.Text = 0
TxtRemMaintCondition.Text = 0
TxtRemMaintClean.Text = 0
TxtRemPaints.Text = 0
TxtRemMaintkitchen.Text = 0
TxtRemElectricity.Text = 0
TxtRemMaintDoors.Text = 0
TxtRemWindows.Text = 0
TxtRemMaintOther.Text = 0

TxtPrice2.Text = 0
TxtPrice3.Text = 0
TxtPrice.Text = 0
TxtRemPrice.Text = 0
End Sub
Private Sub TxtCustCode_KeyPress(KeyAscii As Integer)
    Dim CUSTID As Integer
Dim EmpID As Integer
    If KeyAscii = vbKeyReturn Then
    
            If DCboCashType.ListIndex <> 6 Then
             
                 GetCustomersDetail CUSTID, , TxtCustCode.Text ' , DCboCashType.ListIndex + 1
                 DBCboClientName.BoundText = CUSTID
              
             ElseIf DCboCashType.ListIndex = 6 Then
             
             
                      
                             GetEmployeeIDFromCode TxtCustCode.Text, EmpID
                             Me.DCEmployee.BoundText = EmpID
                   
             
             End If
             
       End If
             
       

End Sub

Private Sub TxtElectricity1_Change()
Calculte
End Sub

Private Sub TxtElectricity1_LostFocus()
If Me.TxtModFlg.Text <> "R" Then
If val(TxtElectricity1.Text) > val(TxtRemElectricity.Text) Then
If SystemOptions.UserInterface = ArabicInterface Then
MsgBox "«·ﬁÌ„… «·„œ›Ê⁄… «ﬂ»— „‰ «·„ »ﬁÌ"
Else
MsgBox "Required value is greater than required"
End If
TxtElectricity1.Text = 0
TxtElectricity1.SetFocus
Exit Sub
End If
End If
End Sub

Function getfitterDeyails(Optional ID As Double) As Double

 
Dim rs2 As ADODB.Recordset
Set rs2 = New ADODB.Recordset
Dim sql As String
sql = " SELECT        dbo.TblFiterWaiver.ID, dbo.TblFiterWaiver.ContNo, dbo.TblContract.ComResid, dbo.TblContract.Iqar, dbo.TblContract.ownerid, dbo.TblContract.UnitType, dbo.TblContract.UnitNo"
sql = sql & " FROM            dbo.TblFiterWaiver INNER JOIN"
sql = sql & "                          dbo.TblContract ON dbo.TblFiterWaiver.ContNo = dbo.TblContract.ContNo"
sql = sql & "  Where (dbo.TblFiterWaiver.ID = " & ID & ")"
Dim ComResid1 As Integer
rs2.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText

If rs2.RecordCount > 0 Then
ComResid1 = IIf(IsNull(rs2("ComResid").value), 0, rs2("ComResid").value)
        
        If ComResid1 = 0 Then
                 ComResid(0).value = True
        Else
                ComResid(1).value = True
        End If

Else
 
End If
 


End Function

Public Sub TxtFilterNo_Change()
getfitterDeyails val(TxtFilterNo)
End Sub

Private Sub txtinstrunce_Change()
If Me.TxtModFlg <> "R" Then
txtTotal1.Text = val(Txtcommission.Text) + val(TxtWater.Text) + val(txtinstrunce.Text)
End If
End Sub

Private Sub TxtMaintClean_Change()
Calculte
End Sub

Private Sub TxtMaintClean_LostFocus()
If Me.TxtModFlg.Text <> "R" Then
If val(TxtMaintClean.Text) > val(TxtRemMaintClean.Text) Then
If SystemOptions.UserInterface = ArabicInterface Then
MsgBox "«·ﬁÌ„… «·„œ›Ê⁄… «ﬂ»— „‰ «·„ »ﬁÌ"
Else
MsgBox "Required value is greater than required"
End If
TxtMaintClean.Text = 0
TxtMaintClean.SetFocus
Exit Sub
End If
End If
End Sub

Private Sub TxtMaintCondition_Change()
Calculte
End Sub

Private Sub TxtMaintCondition_LostFocus()
If Me.TxtModFlg.Text <> "R" Then
If val(TxtMaintCondition.Text) > val(TxtRemMaintCondition.Text) Then
If SystemOptions.UserInterface = ArabicInterface Then
MsgBox "«·ﬁÌ„… «·„œ›Ê⁄… «ﬂ»— „‰ «·„ »ﬁÌ"
Else
MsgBox "Required value is greater than required"
End If
TxtMaintCondition.Text = 0
TxtMaintCondition.SetFocus
Exit Sub
End If
End If
End Sub

Private Sub TxtMaintDoors_Change()
Calculte
End Sub

Private Sub TxtMaintDoors_LostFocus()
If Me.TxtModFlg.Text <> "R" Then
If val(TxtMaintDoors.Text) > val(TxtRemMaintDoors.Text) Then
If SystemOptions.UserInterface = ArabicInterface Then
MsgBox "«·ﬁÌ„… «·„œ›Ê⁄… «ﬂ»— „‰ «·„ »ﬁÌ"
Else
MsgBox "Required value is greater than required"
End If
TxtMaintDoors.Text = 0
TxtMaintDoors.SetFocus
Exit Sub
End If
End If
End Sub

Private Sub TxtMaintenance_Change()
Calculte
End Sub

Private Sub TxtMaintenance_LostFocus()
If Me.TxtModFlg.Text <> "R" Then
If val(TxtMaintenance.Text) > val(TxtRemMaintenance.Text) Then
If SystemOptions.UserInterface = ArabicInterface Then
MsgBox "«·ﬁÌ„… «·„œ›Ê⁄… «ﬂ»— „‰ «·„ »ﬁÌ"
Else
MsgBox "Required value is greater than required"
End If
TxtMaintenance.Text = 0
 TxtMaintenance.SetFocus
Exit Sub
End If
End If
End Sub

Private Sub TxtMaintkitchen_Change()
Calculte
End Sub

Private Sub TxtMaintkitchen_LostFocus()
If Me.TxtModFlg.Text <> "R" Then
If val(TxtMaintkitchen.Text) > val(TxtRemMaintkitchen.Text) Then
If SystemOptions.UserInterface = ArabicInterface Then
MsgBox "«·ﬁÌ„… «·„œ›Ê⁄… «ﬂ»— „‰ «·„ »ﬁÌ"
Else
MsgBox "Required value is greater than required"
End If
TxtMaintkitchen.Text = 0
TxtMaintkitchen.SetFocus
Exit Sub
End If
End If
End Sub

Private Sub TxtMaintOther_Change()
Calculte
End Sub

Private Sub TxtMaintOther_LostFocus()
If Me.TxtModFlg.Text <> "R" Then
If val(TxtMaintOther.Text) > val(TxtRemMaintOther.Text) Then
If SystemOptions.UserInterface = ArabicInterface Then
MsgBox "«·ﬁÌ„… «·„œ›Ê⁄… «ﬂ»— „‰ «·„ »ﬁÌ"
Else
MsgBox "Required value is greater than required"
End If
TxtMaintOther.Text = 0
TxtMaintOther.SetFocus
Exit Sub
End If
End If
End Sub

Private Sub TxtModFlg_Change()
    On Error GoTo ErrTrap

    Select Case Me.TxtModFlg.Text

        Case "R"
    
            If SystemOptions.UserInterface = EnglishInterface Then
                Me.Caption = "Receipts"
            Else
                '        Me.Caption = "«·„ﬁ»Ê÷« "
            End If
'Grid3.Visible = False
            Ele(0).Enabled = False
            Grid.Enabled = False
            GRID1.Enabled = False
          '    Grid3.Enabled = False
            CmdRemove.Enabled = False
            ' Frame1.Enabled = False
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
            XPTxtVal.locked = True
            XPDtbTrans.Enabled = False
            Txt_DateHigri.Enabled = False
            XPMTxtRemarks.locked = True
            DBCboClientName.locked = True
            DCboCashType.locked = True
            Me.CboPaymentType.locked = True
            Me.DcboBox.locked = True
            Me.DcboBankName.locked = True
            Me.TxtChequeNumber.locked = True
            Me.DtpChequeDueDate.Enabled = False

            'If rs.RecordCount < 1 Then
            '    Me.XPBtnMove(0).Enabled = False
            '    Me.XPBtnMove(1).Enabled = False
            '    Me.XPBtnMove(2).Enabled = False
            '    Me.XPBtnMove(3).Enabled = False
            '    Me.Cmd(1).Enabled = False
            '    Me.Cmd(4).Enabled = False
           ' End If

            Fra(0).Enabled = False
            ChkTrans.Enabled = False

        Case "N"
            '        Me.Caption = "«·„ﬁ»Ê÷« ( ÃœÌœ )"
              Grid3.Enabled = True
            Me.Cmd(2).Enabled = True
            Me.Cmd(3).Enabled = True
            Grid.Enabled = True
            GRID1.Enabled = False
            CmdRemove.Enabled = False
    '    Grid3.Visible = True
            Me.Cmd(0).Enabled = False
            Me.Cmd(1).Enabled = False
            Me.Cmd(4).Enabled = False
            Me.Cmd(5).Enabled = False
            Me.Cmd(7).Enabled = False
            '    Me.XPBtnMove(0).Enabled = False
            '    Me.XPBtnMove(1).Enabled = False
            '    Me.XPBtnMove(2).Enabled = False
            '    Me.XPBtnMove(3).Enabled = False
            XPDtbTrans.Enabled = True
            Txt_DateHigri.Enabled = True
            XPTxtVal.locked = False
            XPMTxtRemarks.locked = False
            DBCboClientName.locked = False
            XPDtbTrans.value = Date
            DCboCashType.locked = False
            DCboCashType.ListIndex = 0
        
            Me.CboPaymentType.locked = False
            Me.DcboBox.locked = False
            Me.DcboBankName.locked = False
            Me.TxtChequeNumber.locked = False
            Me.DtpChequeDueDate.Enabled = True
        
            Fra(0).Enabled = True
            ChkTrans.Enabled = True

        Case "E"
            '        Me.Caption = "«·„ﬁ»Ê÷« (  ⁄œÌ· )"
'Grid3.Visible = True
Grid3.Enabled = True
            Grid.Enabled = True
            GRID1.Enabled = True
            
            Grid3.Enabled = True
             
            
            CmdRemove.Enabled = True
        
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
        
            XPTxtVal.locked = False
             If SystemOptions.SysCashDateTakeType = InvDateFromLocalCompuer Then
             XPDtbTrans.Enabled = True
            Txt_DateHigri.Enabled = True
              ElseIf SystemOptions.SysCashDateTakeType = InvDateFromServerComputer Then
         XPDtbTrans.Enabled = False
         Txt_DateHigri.Enabled = False
             ElseIf SystemOptions.SysCashDateTakeType = InvDateFromLastInvDate Then
      XPDtbTrans.Enabled = False
      Txt_DateHigri.Enabled = False

            End If

            
            '        XPCboProfLevel.Locked = False
            '        XPTxtProfMail.Locked = False
            '        XPTxtPhone.Locked = False
            '        XPTxtMobile.Locked = False
            XPMTxtRemarks.locked = False
            DBCboClientName.locked = False
            DCboCashType.locked = False
            Fra(0).Enabled = True
            ChkTrans.Enabled = True
            Me.CboPaymentType.locked = False
            Me.DcboBox.locked = False
            Me.DcboBankName.locked = False
            Me.TxtChequeNumber.locked = False
            Me.DtpChequeDueDate.Enabled = True
    End Select

    Exit Sub
ErrTrap:
End Sub

Private Sub TxtNoteSerial1_Change()
'FillGridWithData1 val(TXTContNo.text)
'If Me.TxtModFlg.text <> "R" Then
If val(Me.DcbIqara.BoundText) <> 0 Then
GetAmola val(Me.DcbIqara.BoundText)
End If
'End If

End Sub

Private Sub TxtPaints_Change()
Calculte
End Sub

Private Sub TxtPaints_LostFocus()
If Me.TxtModFlg.Text <> "R" Then
If val(TxtPaints.Text) > val(TxtRemPaints.Text) Then
If SystemOptions.UserInterface = ArabicInterface Then
MsgBox "«·ﬁÌ„… «·„œ›Ê⁄… «ﬂ»— „‰ «·„ »ﬁÌ"
Else
MsgBox "Required value is greater than required"
End If
TxtPaints.Text = 0
TxtPaints.SetFocus
Exit Sub
End If
End If
End Sub

Private Sub txtPrice_Change()
Calculte
End Sub

Private Sub TxtPrice_LostFocus()
If Me.TxtModFlg.Text <> "R" Then
If val(Me.TxtPrice.Text) > val(TxtRemPrice.Text) Then
If SystemOptions.UserInterface = ArabicInterface Then
MsgBox "«·ﬁÌ„… «·„œ›Ê⁄… «ﬂ»— „‰ «·„ »ﬁÌ"
Else
MsgBox "Required value is greater than required"
End If
TxtPrice.Text = 0
TxtPrice.SetFocus
Exit Sub
End If
End If
End Sub

Private Sub TxtRemainRent_Change()
Calculte
End Sub

Private Sub TxtRemainRent_LostFocus()
If Me.TxtModFlg.Text <> "R" Then
If val(txtRemainRent.Text) > val(TxtRemRemainRent.Text) Then
If SystemOptions.UserInterface = ArabicInterface Then
MsgBox "«·ﬁÌ„… «·„œ›Ê⁄… «ﬂ»— „‰ «·„ »ﬁÌ"
Else
MsgBox "Required value is greater than required"
End If
txtRemainRent.Text = 0
txtRemainRent.SetFocus
Exit Sub
End If
End If
End Sub

Private Sub TxtService_KeyPress(KeyAscii As Integer)
KeyAscii = KeyAscii_Num(KeyAscii, Me.TxtService.Text)
End Sub

Private Sub txttotal1_Change()
If Me.TxtModFlg <> "R" Then
txtTotal2.Text = val(XPTxtVal.Text) - val(txtTotal1.Text)
End If
End Sub

Private Sub txttotal2_Change()
If Me.TxtModFlg <> "R" Then
If val(txtTotal2.Text) > 0 Then
txtinstranc.Text = txtTotal2.Text
Else
txtinstranc.Text = 0
End If
End If
End Sub

Private Sub TxtTransID_Change()

    If Me.TxtModFlg.Text = "N" Or Me.TxtModFlg.Text = "E" Then
        If Me.TxtTransID.Text <> "" Then
            If Me.CboTrans.ListIndex = 0 Or Me.CboTrans.ListIndex = 1 Then
                Me.TxtTransSerial.Text = GetTransIDSerial(1, val(Me.TxtTransID.Text))
            Else
                Me.TxtTransSerial.Text = Me.TxtTransID.Text
            End If
        End If
    End If

End Sub

Private Sub TxtTransSerial_KeyPress(KeyAscii As Integer)
    KeyAscii = KeyAscii_Num(KeyAscii, Me.TxtTransSerial.Text, 1)
End Sub

Private Sub txtWater_Change()
txtTotal1.Text = val(Txtcommission.Text) + val(TxtWater.Text) + val(txtinstrunce.Text)
End Sub

Private Sub ValidityDate_CallbackKeyDown(ByVal KeyCode As Integer, ByVal Shift As Integer, ByVal CallbackField As String, CallbackDate As Date)

End Sub

Private Sub TxtWindows_Change()
Calculte
End Sub

Private Sub TxtWindows_LostFocus()
If Me.TxtModFlg.Text <> "R" Then
If val(TxtWindows.Text) > val(TxtRemWindows.Text) Then
If SystemOptions.UserInterface = ArabicInterface Then
MsgBox "«·ﬁÌ„… «·„œ›Ê⁄… «ﬂ»— „‰ «·„ »ﬁÌ"
Else
MsgBox "Required value is greater than required"
End If
TxtWindows.Text = 0
TxtWindows.SetFocus
Exit Sub
End If
End If
End Sub

Private Sub VSFlexGrid1_AfterEdit(ByVal Row As Long, ByVal Col As Long)
Dim StrAccountCode As String
    Dim Msg As String
    Dim rs2 As New ADODB.Recordset
    Dim StrSQL As String
    Dim ClsAcc As New ClsAccounts
    Dim LngRow As Long
  
    Dim StrAccountType As String
    Dim StrComboList As String
Dim StrAccountCode1 As String
    With VSFlexGrid1
        Select Case .ColKey(Col)
        Case "group"
        StrAccountCode = .ComboData
                LngRow = .FindRow(StrAccountCode, .FixedRows, .ColIndex("groupid"), False, True)
                .TextMatrix(Row, .ColIndex("groupid")) = StrAccountCode
                .TextMatrix(Row, .ColIndex("empname")) = ""
                .TextMatrix(Row, .ColIndex("id")) = ""
                
 Case "empname"
 StrAccountCode = .ComboData
                LngRow = .FindRow(StrAccountCode, .FixedRows, .ColIndex("id"), False, True)
                .TextMatrix(Row, .ColIndex("id")) = StrAccountCode
                '''//
                         
    StrSQL = "SELECT     dbo.TBLSalesRepData.EmpID, dbo.TBLSalesRepData.id, dbo.TblEmployee.Fullcode, dbo.TBLSalesRepData.GroupID, "
    StrSQL = StrSQL & "                 dbo.TBLSalesRepGroups.name ,dbo.TBLSalesRepGroups.NameE "
   
    StrSQL = StrSQL & " FROM         dbo.TBLSalesRepGroups RIGHT OUTER JOIN"
    StrSQL = StrSQL & "                  dbo.TBLSalesRepData ON dbo.TBLSalesRepGroups.id = dbo.TBLSalesRepData.GroupID LEFT OUTER JOIN"
    StrSQL = StrSQL & "                  dbo.TblEmployee ON dbo.TBLSalesRepData.EmpID = dbo.TblEmployee.Emp_ID"

    StrSQL = StrSQL & " where dbo.TBLSalesRepData.EmpID  = " & val(StrAccountCode) & ""
                ''//
                
                 rs2.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText
                 If rs2.RecordCount > 0 Then
                  .TextMatrix(Row, .ColIndex("groupid")) = IIf(IsNull(rs2("GroupID").value), "", rs2("GroupID").value)
                  If SystemOptions.UserInterface = ArabicInterface Then
                  .TextMatrix(Row, .ColIndex("group")) = IIf(IsNull(rs2("name").value), "", rs2("name").value)
                  Else
                  .TextMatrix(Row, .ColIndex("group")) = IIf(IsNull(rs2("namee").value), "", rs2("namee").value)
                  End If
                  
                  .TextMatrix(Row, .ColIndex("idd")) = IIf(IsNull(rs2("id").value), "", rs2("id").value)
                  .TextMatrix(Row, .ColIndex("code")) = IIf(IsNull(rs2("Fullcode").value), "", rs2("Fullcode").value)
                  Else
                   .TextMatrix(Row, .ColIndex("code")) = ""
                   End If
               Case "code"
' StrAccountCode = .ComboData
'                LngRow = .FindRow(StrAccountCode, .FixedRows, .ColIndex("id"), False, True)
'                .TextMatrix(Row, .ColIndex("id")) = StrAccountCode
             If SystemOptions.UserInterface = ArabicInterface Then
    StrSQL = "SELECT     dbo.TBLSalesRepData.EmpID ,  dbo.TblEmployee.emp_name , dbo.TBLSalesRepData.id "
    Else
    StrSQL = "SELECT     dbo.TBLSalesRepData.EmpID ,  dbo.TblEmployee.emp_nameE , dbo.TBLSalesRepData.id "
    End If
    StrSQL = StrSQL & " FROM         dbo.TblEmployee RIGHT OUTER JOIN"
     StrSQL = StrSQL & "                 dbo.TBLSalesRepData ON dbo.TblEmployee.Emp_ID = dbo.TBLSalesRepData.EmpID"
    StrSQL = StrSQL & " where dbo.TblEmployee.Fullcode ='" & .TextMatrix(Row, .ColIndex("code")) & "'"
    
                   rs2.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText
If rs2.RecordCount > 0 Then
                If SystemOptions.UserInterface = ArabicInterface Then
                    .TextMatrix(Row, .ColIndex("empname")) = IIf(IsNull(rs2("emp_name").value), "", rs2("emp_name").value)
                     .TextMatrix(Row, .ColIndex("id")) = IIf(IsNull(rs2("EmpID").value), "", rs2("EmpID").value)
                      .TextMatrix(Row, .ColIndex("idd")) = IIf(IsNull(rs2("id").value), "", rs2("id").value)
                Else
                .TextMatrix(Row, .ColIndex("empname")) = IIf(IsNull(rs2("emp_nameE").value), "", rs2("emp_nameE").value)
                    .TextMatrix(Row, .ColIndex("id")) = IIf(IsNull(rs2("EmpID").value), "", rs2("EmpID").value)
                     .TextMatrix(Row, .ColIndex("idd")) = IIf(IsNull(rs2("id").value), "", rs2("id").value)
                End If
                End If
               ' StrSQL = " select Fullcode from TblEmployee where Emp_ID= " & val(StrAccountCode) & ""
               '  rs2.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText
               '  If rs2.RecordCount > 0 Then
               '   .TextMatrix(Row, .ColIndex("code")) = IIf(IsNull(rs2("Fullcode").value), "", rs2("Fullcode").value)
               '   Else
               '    .TextMatrix(Row, .ColIndex("code")) = ""
               '    End If
                
             
       
    
  End Select
      If Row = .rows - 1 Then
            .rows = .rows + 1
             End If
End With
 ReLineGrid
End Sub

Private Sub VSFlexGrid1_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
With VSFlexGrid1

      
        Select Case .ColKey(Col)
               Case "code"
             .ComboList = ""
                    Case "rate"
             .ComboList = ""
        End Select
    End With
End Sub

Private Sub VSFlexGrid1_StartEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    Dim rs2 As New ADODB.Recordset
    Dim StrSQL  As String
    Dim StrAccountType As String
    Dim StrComboList As String
    Dim Msg As String
    Dim StrAccountCode As String
    Dim StrAccountCode1 As String
    Dim ClsAcc As New ClsAccounts
    Dim LngRow As Long

    With VSFlexGrid1
       Select Case .ColKey(Col)
 
       Case "empname"
             If SystemOptions.UserInterface = ArabicInterface Then
                StrSQL = "SELECT     dbo.TBLSalesRepData.EmpID ,  dbo.TblEmployee.emp_name , dbo.TBLSalesRepData.GroupID"
             Else
                StrSQL = "SELECT     dbo.TBLSalesRepData.EmpID ,  dbo.TblEmployee.emp_nameE , dbo.TBLSalesRepData.GroupID"
             End If
                StrSQL = StrSQL & " FROM         dbo.TblEmployee RIGHT OUTER JOIN"
                StrSQL = StrSQL & " dbo.TBLSalesRepData ON dbo.TblEmployee.Emp_ID = dbo.TBLSalesRepData.EmpID"
    If val(.TextMatrix(Row, .ColIndex("groupid"))) <> 0 Then
    StrSQL = StrSQL & " where dbo.TBLSalesRepData.GroupID=" & val(.TextMatrix(Row, .ColIndex("groupid"))) & ""
    End If
                rs2.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

                If SystemOptions.UserInterface = ArabicInterface Then
                    StrComboList = .BuildComboList(rs2, "emp_name", "EmpID")
                Else
                    StrComboList = .BuildComboList(rs2, "emp_nameE", "EmpID")
                End If
       
                If StrComboList <> "" Then
                    StrComboList = "|" & StrComboList
                End If
                 .ComboList = StrComboList
            Case "group"
             If SystemOptions.UserInterface = ArabicInterface Then
    StrSQL = "SELECT     id ,  name "
    Else
    StrSQL = "SELECT     id , namee"
    End If
    StrSQL = StrSQL & " FROM  TBLSalesRepGroups "
    
    
                rs2.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

                If SystemOptions.UserInterface = ArabicInterface Then
                    StrComboList = .BuildComboList(rs2, "name", "id")
                Else
                    StrComboList = .BuildComboList(rs2, "namee", "id")
                End If
       
                If StrComboList <> "" Then
                    StrComboList = "|" & StrComboList
                End If
                 .ComboList = StrComboList
                 End Select
                 End With
End Sub

Private Sub VSFlexGrid2_AfterEdit(ByVal Row As Long, ByVal Col As Long)
Dim StrAccountCode As String
    Dim Msg As String
    Dim rs2 As New ADODB.Recordset
    Dim StrSQL As String
    Dim ClsAcc As New ClsAccounts
    Dim LngRow As Long
  
    Dim StrAccountType As String
    Dim StrComboList As String
 

Dim StrAccountCode1 As String


    With VSFlexGrid2
               
    
        Select Case .ColKey(Col)
        Case "group"
        StrAccountCode = .ComboData
                LngRow = .FindRow(StrAccountCode, .FixedRows, .ColIndex("groupid"), False, True)
                .TextMatrix(Row, .ColIndex("groupid")) = StrAccountCode
                .TextMatrix(Row, .ColIndex("empname")) = ""
                .TextMatrix(Row, .ColIndex("id")) = ""
                
                
 Case "empname"
 StrAccountCode = .ComboData
                LngRow = .FindRow(StrAccountCode, .FixedRows, .ColIndex("id"), False, True)
                .TextMatrix(Row, .ColIndex("id")) = StrAccountCode
                '''//
                         
    StrSQL = "SELECT     dbo.TBLSalesRepData.EmpID, dbo.TBLSalesRepData.id, dbo.TblEmployee.Fullcode, dbo.TBLSalesRepData.GroupID, "
    StrSQL = StrSQL & "                 dbo.TBLSalesRepGroups.name ,dbo.TBLSalesRepGroups.NameE "
   
    StrSQL = StrSQL & " FROM         dbo.TBLSalesRepGroups RIGHT OUTER JOIN"
    StrSQL = StrSQL & "                  dbo.TBLSalesRepData ON dbo.TBLSalesRepGroups.id = dbo.TBLSalesRepData.GroupID LEFT OUTER JOIN"
    StrSQL = StrSQL & "                  dbo.TblEmployee ON dbo.TBLSalesRepData.EmpID = dbo.TblEmployee.Emp_ID"

    StrSQL = StrSQL & " where dbo.TBLSalesRepData.EmpID  = " & val(StrAccountCode) & ""
                ''//
                
                 rs2.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText
                 If rs2.RecordCount > 0 Then
                  .TextMatrix(Row, .ColIndex("groupid")) = IIf(IsNull(rs2("GroupID").value), "", rs2("GroupID").value)
                  If SystemOptions.UserInterface = ArabicInterface Then
                  .TextMatrix(Row, .ColIndex("group")) = IIf(IsNull(rs2("name").value), "", rs2("name").value)
                  Else
                  .TextMatrix(Row, .ColIndex("group")) = IIf(IsNull(rs2("namee").value), "", rs2("namee").value)
                  End If
                  
                  .TextMatrix(Row, .ColIndex("idd")) = IIf(IsNull(rs2("id").value), "", rs2("id").value)
                  .TextMatrix(Row, .ColIndex("code")) = IIf(IsNull(rs2("Fullcode").value), "", rs2("Fullcode").value)
                  Else
                   .TextMatrix(Row, .ColIndex("code")) = ""
                   End If
               Case "code"
' StrAccountCode = .ComboData
'                LngRow = .FindRow(StrAccountCode, .FixedRows, .ColIndex("id"), False, True)
'                .TextMatrix(Row, .ColIndex("id")) = StrAccountCode
             If SystemOptions.UserInterface = ArabicInterface Then
    StrSQL = "SELECT     dbo.TBLSalesRepData.EmpID ,  dbo.TblEmployee.emp_name , dbo.TBLSalesRepData.id "
    Else
    StrSQL = "SELECT     dbo.TBLSalesRepData.EmpID ,  dbo.TblEmployee.emp_nameE , dbo.TBLSalesRepData.id "
    End If
    StrSQL = StrSQL & " FROM         dbo.TblEmployee RIGHT OUTER JOIN"
     StrSQL = StrSQL & "                 dbo.TBLSalesRepData ON dbo.TblEmployee.Emp_ID = dbo.TBLSalesRepData.EmpID"
    StrSQL = StrSQL & " where dbo.TblEmployee.Fullcode ='" & .TextMatrix(Row, .ColIndex("code")) & "'"
    
                   rs2.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText
If rs2.RecordCount > 0 Then
                If SystemOptions.UserInterface = ArabicInterface Then
                    .TextMatrix(Row, .ColIndex("empname")) = IIf(IsNull(rs2("emp_name").value), "", rs2("emp_name").value)
                     .TextMatrix(Row, .ColIndex("id")) = IIf(IsNull(rs2("EmpID").value), "", rs2("EmpID").value)
                      .TextMatrix(Row, .ColIndex("idd")) = IIf(IsNull(rs2("id").value), "", rs2("id").value)
                Else
                .TextMatrix(Row, .ColIndex("empname")) = IIf(IsNull(rs2("emp_nameE").value), "", rs2("emp_nameE").value)
                    .TextMatrix(Row, .ColIndex("id")) = IIf(IsNull(rs2("EmpID").value), "", rs2("EmpID").value)
                     .TextMatrix(Row, .ColIndex("idd")) = IIf(IsNull(rs2("id").value), "", rs2("id").value)
                End If
                End If


  End Select
      If Row = .rows - 1 Then
            .rows = .rows + 1
             End If
End With
 ReLineGrid
End Sub

Private Sub VSFlexGrid2_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
With VSFlexGrid2

      
        Select Case .ColKey(Col)
      
 
           
               Case "code"
             .ComboList = ""
                    Case "rate"
             .ComboList = ""
        End Select

    End With
End Sub

Private Sub VSFlexGrid2_StartEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
Dim rs2 As New ADODB.Recordset
    Dim StrSQL  As String
    Dim StrAccountType As String
    Dim StrComboList As String
    Dim Msg As String
Dim StrAccountCode As String
Dim StrAccountCode1 As String

    Dim ClsAcc As New ClsAccounts
    Dim LngRow As Long

    With VSFlexGrid2

        Select Case .ColKey(Col)
 
            Case "empname"
             If SystemOptions.UserInterface = ArabicInterface Then
    StrSQL = "SELECT     dbo.TBLSalesRepData.EmpID ,  dbo.TblEmployee.emp_name , dbo.TBLSalesRepData.GroupID"
    Else
    StrSQL = "SELECT     dbo.TBLSalesRepData.EmpID ,  dbo.TblEmployee.emp_nameE , dbo.TBLSalesRepData.GroupID"
    End If
    StrSQL = StrSQL & " FROM         dbo.TblEmployee RIGHT OUTER JOIN"
    StrSQL = StrSQL & " dbo.TBLSalesRepData ON dbo.TblEmployee.Emp_ID = dbo.TBLSalesRepData.EmpID"
    If val(.TextMatrix(Row, .ColIndex("groupid"))) <> 0 Then
    StrSQL = StrSQL & " where dbo.TBLSalesRepData.GroupID=" & val(.TextMatrix(Row, .ColIndex("groupid"))) & ""
    End If
                rs2.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

                If SystemOptions.UserInterface = ArabicInterface Then
                    StrComboList = VSFlexGrid2.BuildComboList(rs2, "emp_name", "EmpID")
                Else
                    StrComboList = VSFlexGrid2.BuildComboList(rs2, "emp_nameE", "EmpID")
                End If
       
                If StrComboList <> "" Then
                    StrComboList = "|" & StrComboList
                End If
                 .ComboList = StrComboList
            Case "group"
             If SystemOptions.UserInterface = ArabicInterface Then
    StrSQL = "SELECT     id ,  name "
    Else
    StrSQL = "SELECT     id , namee"
    End If
    StrSQL = StrSQL & " FROM  TBLSalesRepGroups "
    
    
                rs2.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

                If SystemOptions.UserInterface = ArabicInterface Then
                    StrComboList = VSFlexGrid1.BuildComboList(rs2, "name", "id")
                Else
                    StrComboList = VSFlexGrid1.BuildComboList(rs2, "namee", "id")
                End If
       
                If StrComboList <> "" Then
                    StrComboList = "|" & StrComboList
                End If
                 .ComboList = StrComboList
                 End Select
                 End With
End Sub

Public Sub XPBtnMove_Click(Index As Integer)
Dim StrSQL As String
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
  '  Set rs = New ADODB.Recordset
    'StrSQL = "select * From Notes where NoteType=4 and   displayed is null Order By NoteID"
If RereivID = 0 Then
    StrSQL = "select * From Notes where NoteType=4    "
   Else
   StrSQL = "select * From Notes where NoteType=4 and NoteID=" & RereivID & "   "
   End If
StrSQL = StrSQL & " and CashingType >= 7"
    If SystemOptions.usertype <> UserAdminAll Then
        StrSQL = StrSQL & " AND   branch_no=" & Current_branch
    End If
     RereivID = 0
                 If SystemOptions.usertype <> UserAdminAll Then
 
          If SystemOptions.FixedCustomer = 1 Then
            StrSQL = StrSQL & " and  UserID = " & user_id
             End If
  
        Me.Dcbranch.Enabled = True
      
      
    End If
    
    StrSQL = StrSQL & "and  displayed is null Order By NoteID"
    Set rs = New ADODB.Recordset
    rs.CursorLocation = adUseClient
    rs.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText
    
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
    ReLineGrid

    Exit Sub
ErrTrap:
End Sub

Public Sub Retrive(Optional Lngid As Long = 0)
       Dim rs2 As ADODB.Recordset
    Dim RsTemp As ADODB.Recordset
    Dim StrSQL As String
    Dim RsDev As ADODB.Recordset
    Dim i As Integer
    On Error GoTo ErrTrap

    If rs.RecordCount < 1 Then
        XPTxtCurrent.Caption = 0
        XPTxtCount.Caption = 0
        Exit Sub
    End If

    If rs.EOF Or rs.BOF Then
        Exit Sub
    Else

        If Lngid <> 0 Then
            rs.Find "NoteID=" & Lngid, , adSearchForward, adBookmarkFirst

           ' If rs.EOF Or rs.BOF Then
           '     Exit Sub
           ' End If
        End If
    End If

   ' If Not IsNull(rs("general_cost_center").value) Then
   '     Me.DcCostCenter.BoundText = IIf(rs("general_cost_center").value = "", "", rs("general_cost_center").value)
   ' End If

    Dcbranch.BoundText = IIf(IsNull(rs("branch_no").value), "", rs("branch_no").value)
    Me.DcEmp.BoundText = IIf(IsNull(rs("EmpId")), "", rs("EmpId"))
    DcboCreditSide.BoundText = IIf(IsNull(rs("CreditSide").value), "", rs("CreditSide").value)
    DcboDebitSide.BoundText = IIf(IsNull(rs("DebitSide").value), "", rs("DebitSide").value)
    
    Me.Text1.Text = IIf(IsNull(rs("foxy_no").value), "", rs("foxy_no").value)
    XPTxtID.Text = IIf(IsNull(rs("NoteID").value), "", val(rs("NoteID").value))
    TxtNoteSerial.Text = IIf(IsNull(rs("NoteSerial").value), "", rs("NoteSerial").value)
    TxtNoteSerial1.Text = IIf(IsNull(rs("NoteSerial1").value), "", rs("NoteSerial1").value)
    TxtManulaNO.Text = IIf(IsNull(rs("ManulaNO").value), "", rs("ManulaNO").value)
    TxtBookNo.Text = IIf(IsNull(rs("BookNo").value), "", rs("BookNo").value)
    Me.TxtContNo.Text = IIf(IsNull(rs("ContNo").value), "", rs("ContNo").value)
        Me.txtContractNo.Text = IIf(IsNull(rs("ContractNo").value), "", rs("ContractNo").value)
        If Not IsNull(rs("ComResid").value) Then
        If (rs("ComResid").value) = 1 Then
        ComResid(1).value = True
        Else
        ComResid(0).value = True
        End If
        Else
        ComResid(0).value = True
        End If
        If Not IsNull(rs("TypAmola").value) Then
        If rs("TypAmola").value = 1 Then
        Rd(1).value = True
         Else
        Rd(0).value = True
        End If
        Else
        Rd(0).value = True
        End If
        
      TxtKickbacks.Text = IIf(IsNull(rs("AmolaValus").value), 0, rs("AmolaValus").value)
    Me.oldtxtNoteSerial1.Text = IIf(IsNull(rs("OldNoteSerial1").value), IIf(IsNull(rs("NoteSerial1").value), "", rs("NoteSerial1").value), rs("OldNoteSerial1").value)

    lbl(46).Caption = showLabel(TxtNoteSerial1, oldtxtNoteSerial1)
TxtPrice.Text = IIf(IsNull(rs("Price").value), 0, rs("Price").value)
TxtPrice2.Text = IIf(IsNull(rs("Price2").value), 0, rs("Price2").value)
TxtPrice3.Text = IIf(IsNull(rs("Price3").value), 0, rs("Price3").value)
TxtRemPrice.Text = IIf(IsNull(rs("RemPrice").value), 0, rs("RemPrice").value)
TxtVATValue.Text = IIf(IsNull(rs("VAT").value), 0, rs("VAT").value)
'rs("VAT").value = IIf(TxtVATValue.Text = "", Null, val(TxtVATValue.Text))
lblremain.Caption = IIf(IsNull(rs("RemaiValue").value), 0, rs("RemaiValue").value)
    txtperson.Text = IIf(IsNull(rs("person").value), "", rs("person").value)
Option1.value = False
Option2.value = False
Option3.value = False
Option7.value = False
' C1Elastic1.Visible = False
If IsNull(rs("NCashingType").value) Then

Else
        If rs("NCashingType").value = 1 Then
               Option1.value = True
        ElseIf rs("NCashingType").value = 2 Then
              Option2.value = True
        ElseIf rs("NCashingType").value = 3 Then
             Option3.value = True
           ElseIf rs("NCashingType").value = 7 Then
             Option7.value = True
        End If
End If



 
   
    XPTxtVal.Text = IIf(IsNull(rs("Note_Value").value), "", Trim(rs("Note_Value").value))
    
    Me.txtoldvalue.Text = val(XPTxtVal.Text)
    ''//
    DcbAccount.BoundText = IIf(IsNull(rs("AccountPaym").value), "", rs("AccountPaym").value)
    TxtService.Text = IIf(IsNull(rs("Servce").value), 0, Trim(rs("Servce").value))
           TxtRent.Text = IIf(IsNull(rs("rent").value), "", Trim(rs("rent").value))
          Txtcommission.Text = IIf(IsNull(rs("commission").value), "", Trim(rs("commission").value))
          Me.TxtCommissionOut.Text = IIf(IsNull(rs("CommissionOut").value), "", Trim(rs("CommissionOut").value))
          TxtWater.Text = IIf(IsNull(rs("Water").value), "", Trim(rs("Water").value))
          TxtElectricity.Text = IIf(IsNull(rs("Electricity").value), "", Trim(rs("Electricity").value))
         txtinstrunce.Text = IIf(IsNull(rs("Instrunce").value), "", Trim(rs("Instrunce").value))
         txtComisin.Text = IIf(IsNull(rs("comX").value), "", Trim(rs("comX").value))
          txtinstranc.Text = IIf(IsNull(rs("ComY").value), "", Trim(rs("ComY").value))
          TxtTelphone.Text = IIf(IsNull(rs("Telephone").value), "", Trim(rs("Telephone").value))
          If rs("StatusEarnest").value = 1 Then
          CheckStatusEarnest(0).value = vbChecked
   ElseIf rs("StatusEarnest").value = 2 Then
     CheckStatusEarnest(1).value = vbChecked
     ElseIf rs("StatusEarnest").value = 3 Then
     CheckStatusEarnest(2).value = vbChecked
     ElseIf rs("StatusEarnest").value = 4 Then
     CheckStatusEarnest(3).value = vbChecked
     
     Else
      CheckStatusEarnest(1).value = vbUnchecked
      CheckStatusEarnest(0).value = vbUnchecked
      CheckStatusEarnest(2).value = vbUnchecked
      CheckStatusEarnest(3).value = vbUnchecked
     End If
    ''//
    FrmPriodDate.value = IIf(IsNull(rs("FrmPriodDate").value), Date, rs("FrmPriodDate").value)
    FrmPriodDateH.value = IIf(IsNull(rs("FrmPriodDateH").value), ToHijriDate(FrmPriodDate.value), rs("FrmPriodDateH").value)
    XPDtbTrans.value = IIf(IsNull(rs("ToPriodDate").value), Date, rs("ToPriodDate").value)
    ToPriodDateH.value = IIf(IsNull(rs("ToPriodDateH").value), ToHijriDate(ToPriodDate.value), rs("ToPriodDateH").value)
    TxtRemarks.Text = IIf(IsNull(rs("Remark2").value), "", Trim(rs("Remark2").value))
    ''//
    TXTBankName.Text = IIf(IsNull(rs("BankName").value), "", Trim(rs("BankName").value))
 
    txtAdv_payment_value.Text = IIf(IsNull(rs("Adv_payment_value").value), "", Trim(rs("Adv_payment_value").value))

    XPMTxtRemarks.Text = IIf(IsNull(rs("Remark").value), "", Trim(rs("Remark").value))
    'dcproject.BoundText = IIf(IsNull(Rs("Remark").value), "", Trim(Rs("Remark").value))

    XPDtbTrans.value = IIf(IsNull(rs("NoteDate").value), Date, rs("NoteDate").value)
    Txt_DateHigri.value = IIf(IsNull(rs("NoteDateH").value), ToHijriDate(XPDtbTrans.value), rs("NoteDateH").value)
    Me.DCCar.BoundText = IIf(IsNull(rs("CarId").value), "", rs("CarId").value)
    Me.DCDriver.BoundText = IIf(IsNull(rs("DriverId").value), "", rs("DriverId").value)

    Me.DCboUserName.BoundText = IIf(IsNull(rs("UserID").value), "", rs("UserID").value)
    DBCboClientName.BoundText = IIf(IsNull(rs("CusID").value), 0, rs("CusID").value)

  Me.DcbIqara.BoundText = val(IIf(IsNull(rs.Fields("akarid").value), 0, rs.Fields("akarid").value))
     Me.DcbUnitType.BoundText = val(IIf(IsNull(rs.Fields("UnitType").value), -1, rs.Fields("UnitType").value))
  DcbUnitType_Change
     Me.DcbUnitNo.BoundText = val(IIf(IsNull(rs.Fields("UnitNo").value), -1, rs.Fields("UnitNo").value))

TxtInterval.Text = IIf(IsNull(rs("interval").value), 0, (rs("interval").value))
cbointervaltype.ListIndex = IIf(IsNull(rs("intervaltype").value), 0, (rs("intervaltype").value))
    txtrenterName.Text = IIf(IsNull(rs("renterName").value), "", Trim(rs("renterName").value))
TxtFilterNo.Text = IIf(IsNull(rs("FilterID").value), "", Trim(rs("FilterID").value))
TXtFilter.Text = IIf(IsNull(rs("FIlterTotal").value), "", Trim(rs("FIlterTotal").value))
txtTotalinsuranceS.Text = IIf(IsNull(rs("TotalInsurances").value), "", Trim(rs("TotalInsurances").value))


    '-----------------------------------------------------------------------------
    If IsNull(rs("NoteCashingType").value) Then
        Me.CboPaymentType.ListIndex = 0
        Me.DcboBox.BoundText = IIf(IsNull(rs("BoxID").value), 0, rs("BoxID").value)
    
        'project_Expensen_account
        Me.DcboBankName.BoundText = ""
        Me.TxtChequeNumber.Text = ""
        Me.DCChequeBox.BoundText = ""
    ElseIf rs("NoteCashingType").value = 0 Then
        Me.CboPaymentType.ListIndex = 0
        Me.DcboBox.BoundText = IIf(IsNull(rs("BoxID").value), 0, rs("BoxID").value)
        Me.DcboBankName.BoundText = ""
        Me.TxtChequeNumber.Text = ""
        Me.DCChequeBox.BoundText = ""
    ElseIf rs("NoteCashingType").value = 1 Then
        Me.CboPaymentType.ListIndex = 1
        Me.DcboBox.BoundText = ""
    
        Me.TxtChequeNumber.Text = rs("ChqueNum").value
        Me.DtpChequeDueDate.value = rs("DueDate").value
    
        If SystemOptions.ChequeBox = True Then
            Me.DCChequeBox.BoundText = rs("ChequeBoxID").value
        Else
            Me.DCChequeBox.BoundText = ""
            Me.DcboBankName.BoundText = rs("BankID").value
        End If

    ElseIf rs("NoteCashingType").value = 2 Then

        If SystemOptions.ChequeBox = True Then
            TXTBankName.Visible = True
            'Me.DCChequeBox.BoundText = rs("ChequeBoxID").value
        Else
            TXTBankName.Visible = False
            Me.DCChequeBox.BoundText = ""
            Me.DcboBankName.BoundText = rs("BankID").value
        End If

        Me.CboPaymentType.ListIndex = 2
        Me.DcboBox.BoundText = ""
        Me.DcboBankName.BoundText = rs("BankID").value
        Me.TxtChequeNumber.Text = rs("ChqueNum").value
        Me.DtpChequeDueDate.value = rs("DueDate").value
        Me.DCChequeBox.BoundText = ""
ElseIf rs("NoteCashingType").value = 4 Then
Me.CboPaymentType.ListIndex = 4
Me.DcboBox.BoundText = ""
     Me.DcboBankName.BoundText = ""
        Me.TxtChequeNumber.Text = ""
        Me.DCChequeBox.BoundText = ""
    ElseIf rs("NoteCashingType").value = 3 Then

        If SystemOptions.ChequeBox = True Then
            TXTBankName.Visible = True
            'Me.DCChequeBox.BoundText = rs("ChequeBoxID").value
        Else
            TXTBankName.Visible = False
            Me.DCChequeBox.BoundText = ""
            Me.DcboBankName.BoundText = rs("BankID").value
        End If

        Me.CboPaymentType.ListIndex = 3
        Me.DcboBox.BoundText = ""
        Me.DcboBankName.BoundText = rs("BankID").value
        Me.TxtChequeNumber.Text = rs("ChqueNum").value
        Me.DtpChequeDueDate.value = rs("DueDate").value
        Me.DCChequeBox.BoundText = ""
    
    End If
DCboCashType.ListIndex = IIf(IsNull(rs("CashingType").value), -1, rs("CashingType").value)
    CboPayMentType_Change
    If val(DCboCashType.ListIndex) = 13 Then
    If Not IsNull(rs("TypeTrans").value) Then
    If (rs("TypeTrans").value) = 1 Then
    RdTypeTrans(1).value = True
    Else
    RdTypeTrans(0).value = True
    End If
    Else
    RdTypeTrans(0).value = True
    End If
    txtInsurance.Text = IIf(IsNull(rs("Insurance").value), 0, rs("Insurance").value)
    txtDiscount.Text = IIf(IsNull(rs("Discount").value), 0, rs("Discount").value)
    TxtMaintenance2.Text = IIf(IsNull(rs("Maintenance2").value), 0, rs("Maintenance2").value)
    TxtMaintenance3.Text = IIf(IsNull(rs("Maintenance3").value), 0, rs("Maintenance3").value)
    TxtMaintenance.Text = IIf(IsNull(rs("Maintenance").value), 0, rs("Maintenance").value)
    txtRemainRent.Text = IIf(IsNull(rs("RemainRent").value), 0, rs("RemainRent").value)
    TxtRemainRent2.Text = IIf(IsNull(rs("RemainRent2").value), 0, rs("RemainRent2").value)
    TxtRemainRent3.Text = IIf(IsNull(rs("RemainRent3").value), 0, rs("RemainRent3").value)
    TxtMaintCondition3.Text = IIf(IsNull(rs("MaintCondition3").value), 0, rs("MaintCondition3").value)
    TxtMaintCondition2.Text = IIf(IsNull(rs("MaintCondition2").value), 0, rs("MaintCondition2").value)
    TxtMaintCondition.Text = IIf(IsNull(rs("MaintCondition").value), 0, rs("MaintCondition").value)
    TxtMaintClean3.Text = IIf(IsNull(rs("MaintClean3").value), 0, rs("MaintClean3").value)
    TxtMaintClean2.Text = IIf(IsNull(rs("MaintClean2").value), 0, rs("MaintClean2").value)
    TxtMaintClean.Text = IIf(IsNull(rs("MaintClean").value), 0, rs("MaintClean").value)
    TxtPaints2.Text = IIf(IsNull(rs("Paints2").value), 0, rs("Paints2").value)
    TxtPaints3.Text = IIf(IsNull(rs("Paints3").value), 0, rs("Paints3").value)
    TxtPaints.Text = IIf(IsNull(rs("Paints").value), 0, rs("Paints").value)
    TxtMaintkitchen2.Text = IIf(IsNull(rs("Maintkitchen2").value), 0, rs("Maintkitchen2").value)
    TxtMaintkitchen3.Text = IIf(IsNull(rs("Maintkitchen3").value), 0, rs("Maintkitchen3").value)
    
    TxtMaintkitchen2.Text = IIf(IsNull(rs("Maintkitchen2").value), 0, rs("Maintkitchen2").value)
    TxtMaintkitchen3.Text = IIf(IsNull(rs("Maintkitchen3").value), 0, rs("Maintkitchen3").value)
    TxtMaintkitchen.Text = IIf(IsNull(rs("Maintkitchen").value), 0, rs("Maintkitchen").value)
    TxtElectricity12.Text = IIf(IsNull(rs("Electricity12").value), 0, rs("Electricity12").value)
    TxtElectricity13.Text = IIf(IsNull(rs("Electricity13").value), 0, rs("Electricity13").value)
    TxtElectricity1.Text = IIf(IsNull(rs("Electricity1").value), 0, rs("Electricity1").value)
    TxtMaintDoors2.Text = IIf(IsNull(rs("MaintDoors2").value), 0, rs("MaintDoors2").value)
    TxtMaintDoors3.Text = IIf(IsNull(rs("MaintDoors3").value), 0, rs("MaintDoors3").value)
    TxtMaintDoors.Text = IIf(IsNull(rs("MaintDoors").value), 0, rs("MaintDoors").value)
    TxtWindows2.Text = IIf(IsNull(rs("Windows2").value), 0, rs("Windows2").value)
    TxtWindows3.Text = IIf(IsNull(rs("Windows3").value), 0, rs("Windows3").value)
    TxtWindows.Text = IIf(IsNull(rs("Windows").value), 0, rs("Windows").value)
    TxtMaintOther2.Text = IIf(IsNull(rs("MaintOther2").value), 0, rs("MaintOther2").value)
    TxtMaintOther.Text = IIf(IsNull(rs("MaintOther").value), 0, rs("MaintOther").value)
    TxtMaintOther3.Text = IIf(IsNull(rs("MaintOther3").value), 0, rs("MaintOther3").value)
    TxtTotal22.Text = IIf(IsNull(rs("Total22").value), 0, rs("Total22").value)
    TxtTotal23.Text = IIf(IsNull(rs("Total23").value), 0, rs("Total23").value)
    txtTotal.Text = IIf(IsNull(rs("Total21").value), 0, rs("Total21").value)
    TxtTotalAftreIns.Text = IIf(IsNull(rs("TotalAftreIns").value), 0, rs("TotalAftreIns").value)
    TxtTotalAftreIns2.Text = IIf(IsNull(rs("TotalAftreIns2").value), 0, rs("TotalAftreIns2").value)
    TxtTotalAftreIns3.Text = IIf(IsNull(rs("TotalAftreIns3").value), 0, rs("TotalAftreIns3").value)
    txtNet.Text = IIf(IsNull(rs("Net").value), 0, rs("Net").value)
    txtNet2.Text = IIf(IsNull(rs("Net2").value), 0, rs("Net2").value)
    TxtNet3.Text = IIf(IsNull(rs("Net3").value), 0, rs("Net3").value)
    End If
    DCboCashType.ListIndex = IIf(IsNull(rs("CashingType").value), -1, rs("CashingType").value)
    XPTxtVal.Text = IIf(IsNull(rs("Note_Value2").value), IIf(IsNull(rs("Note_Value").value), 0, (rs("Note_Value").value)), (rs("Note_Value2").value)) - IIf(IsNull(rs("PreVAT").value), 0, (rs("PreVAT").value))
    If val(XPTxtVal.Text) = 0 Or val(XPTxtVal.Text) < val(rs("Note_Value").value & "") Then
        XPTxtVal.Text = val(rs("Note_Value").value & "") - IIf(IsNull(rs("PreVAT").value), 0, (rs("PreVAT").value))
    End If
    '-----------------------------------------------------------------------------
    If Not IsNull(rs("Transaction_ID").value) Then
        Me.ChkTrans.value = vbChecked
        'Me.ChkTrans.Enabled = True
        Set RsTemp = New ADODB.Recordset
        StrSQL = "Select * From Transactions Where Transaction_ID=" & rs("Transaction_ID").value
        RsTemp.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

        If Not (RsTemp.BOF Or RsTemp.EOF) Then
            Me.TxtTransID.Text = RsTemp("Transaction_ID").value
            Me.TxtTransSerial.Text = IIf(IsNull(RsTemp("Transaction_Serial").value), "", RsTemp("Transaction_Serial").value)

            If Not (IsNull(RsTemp("Transaction_Type").value)) Then
                If RsTemp("Transaction_Type").value = 5 Then
                    Me.CboTrans.ListIndex = 1
                ElseIf RsTemp("Transaction_Type").value = 2 Then
                    Me.CboTrans.ListIndex = 0
                End If
            End If
        End If

    ElseIf Not IsNull(rs("MaintananceID").value) Then
        Me.ChkTrans.value = vbChecked
        Me.CboTrans.ListIndex = 2
        Me.TxtTransID.Text = rs("MaintananceID").value
        Me.TxtTransSerial.Text = rs("MaintananceID").value
    ElseIf Not IsNull(rs("RevenuesID").value) Then
        Me.DcboRevenuesTypes.BoundText = rs("RevenuesID").value
        Me.ChkTrans.value = vbUnchecked
        Me.CboTrans.ListIndex = -1
        Me.TxtTransID.Text = ""
        Me.TxtTransSerial.Text = ""
    Else
        Me.ChkTrans.value = vbUnchecked
        Me.CboTrans.ListIndex = -1
        Me.TxtTransID.Text = ""
        Me.TxtTransSerial.Text = ""
    End If

    If DCboCashType.ListIndex = 5 Then
        Dim My_SQL As String
        My_SQL = "  select id,Project_name from projects where not(REVENUE_account is null)" '
        fill_combo Me.DBCboClientName, My_SQL
      
        DBCboClientName.BoundText = IIf(IsNull(rs("project_id").value), "", rs("project_id").value)
        Dim cus_or_sub As Integer
        cus_or_sub = IIf(IsNull(rs("cus_or_sub").value), 0, rs("cus_or_sub").value)

        If cus_or_sub = 0 Then
            Option4.value = True
        Else
            Option5.value = True
        End If

    End If

    If DCboCashType.ListIndex = 6 Then
        DCEmployee.BoundText = IIf(IsNull(rs("EmployeeID").value), "", rs("EmployeeID").value)
    End If
  
    If DCboCashType.ListIndex = 7 Then
        Me.DCAccounts.BoundText = IIf(IsNull(rs("AccountsCode").value), "", rs("AccountsCode").value)
    End If
''//

    Set rs2 = New ADODB.Recordset
My_SQL = "SELECT     dbo.TblNotesSales.NoteID, dbo.TblNotesSales.ID, TblNotesSales.ValueAmount,dbo.TblNotesSales.rate, dbo.TblNotesSales.valu, dbo.TblNotesSales.Type, dbo.TblNotesSales.EmpID, "
My_SQL = My_SQL & "                      dbo.TblEmployee.Emp_Name, dbo.TblEmployee.Fullcode, dbo.TblEmployee.Emp_Namee, dbo.TblNotesSales.idd, dbo.TblNotesSales.GroupID,"
My_SQL = My_SQL & "                      dbo.TBLSalesRepGroups.name , dbo.TBLSalesRepGroups.NameE"
My_SQL = My_SQL & " FROM         dbo.TblNotesSales LEFT OUTER JOIN"
My_SQL = My_SQL & "                      dbo.TBLSalesRepGroups ON dbo.TblNotesSales.GroupID = dbo.TBLSalesRepGroups.id LEFT OUTER JOIN"
My_SQL = My_SQL & "                      dbo.TblEmployee ON dbo.TblNotesSales.EmpID = dbo.TblEmployee.Emp_ID"
My_SQL = My_SQL & " Where (dbo.TblNotesSales.Type = 0) And (dbo.TblNotesSales.NoteID =" & val(Me.XPTxtID.Text) & ")"
'My_SQL = My_SQL & " Where (dbo.TblCOntractSales.ContNo =" & val(Me.TXTContNo.text) & ")"
    rs2.Open My_SQL, Cn, adOpenKeyset, adLockReadOnly, adCmdText
'    rs1.Open My_SQL, Cn, adOpenStatic, adLockReadOnly, adCmdText
      With Me.VSFlexGrid2
       .rows = 1
        .Clear flexClearScrollable

        If rs2.RecordCount > 0 Then
           .rows = rs2.RecordCount + 1
           rs2.MoveFirst

            For i = 1 To .rows - 1
               .TextMatrix(i, .ColIndex("Ser")) = i
   If SystemOptions.UserInterface = EnglishInterface Then
   .TextMatrix(i, .ColIndex("group")) = IIf(IsNull(rs2.Fields("namee").value), "", rs2.Fields("namee").value)
      .TextMatrix(i, .ColIndex("empname")) = IIf(IsNull(rs2.Fields("Emp_Namee").value), "", rs2.Fields("Emp_Namee").value)
      Else
      .TextMatrix(i, .ColIndex("empname")) = IIf(IsNull(rs2.Fields("Emp_Name").value), "", rs2.Fields("Emp_Name").value)
 .TextMatrix(i, .ColIndex("group")) = IIf(IsNull(rs2.Fields("name").value), "", rs2.Fields("name").value)
    End If
     .TextMatrix(i, .ColIndex("values")) = val(IIf(IsNull(rs2.Fields("valu").value), "", rs2.Fields("valu").value))
 .TextMatrix(i, .ColIndex("rate")) = val(IIf(IsNull(rs2.Fields("rate").value), "", rs2.Fields("rate").value))
  .TextMatrix(i, .ColIndex("code")) = IIf(IsNull(rs2.Fields("Fullcode").value), "", rs2.Fields("Fullcode").value)
  .TextMatrix(i, .ColIndex("ValueAmount")) = IIf(IsNull(rs2.Fields("ValueAmount").value), "", rs2.Fields("ValueAmount").value)
  
  .TextMatrix(i, .ColIndex("id")) = IIf(IsNull(rs2.Fields("EmpID").value), "", rs2.Fields("EmpID").value)
   .TextMatrix(i, .ColIndex("idd")) = IIf(IsNull(rs2.Fields("idd").value), "", rs2.Fields("idd").value)
   .TextMatrix(i, .ColIndex("groupid")) = IIf(IsNull(rs2.Fields("GroupID").value), "", rs2.Fields("GroupID").value)
        rs2.MoveNext
            Next i

         
        End If

        .RowHeight(-1) = 300
    End With
    
    
    '''
        Set rs2 = New ADODB.Recordset
My_SQL = "SELECT     dbo.TblNotesSales.NoteID, dbo.TblNotesSales.ID, dbo.TblNotesSales.rate, dbo.TblNotesSales.valu, dbo.TblNotesSales.Type, dbo.TblNotesSales.EmpID, "
My_SQL = My_SQL & "                      dbo.TblEmployee.Emp_Name, dbo.TblEmployee.Fullcode, dbo.TblEmployee.Emp_Namee, dbo.TblNotesSales.idd, dbo.TblNotesSales.GroupID,"
My_SQL = My_SQL & "                      dbo.TBLSalesRepGroups.name , dbo.TBLSalesRepGroups.NameE"
My_SQL = My_SQL & " FROM         dbo.TblNotesSales LEFT OUTER JOIN"
My_SQL = My_SQL & "                      dbo.TBLSalesRepGroups ON dbo.TblNotesSales.GroupID = dbo.TBLSalesRepGroups.id LEFT OUTER JOIN"
My_SQL = My_SQL & "                      dbo.TblEmployee ON dbo.TblNotesSales.EmpID = dbo.TblEmployee.Emp_ID"
My_SQL = My_SQL & " Where (dbo.TblNotesSales.Type = 1) And (dbo.TblNotesSales.NoteID =" & val(Me.XPTxtID.Text) & ")"
'My_SQL = My_SQL & " Where (dbo.TblCOntractSales.ContNo =" & val(Me.TXTContNo.text) & ")"
    rs2.Open My_SQL, Cn, adOpenKeyset, adLockReadOnly, adCmdText
'    rs1.Open My_SQL, Cn, adOpenStatic, adLockReadOnly, adCmdText
      With Me.VSFlexGrid1
       .rows = 1
        .Clear flexClearScrollable

        If rs2.RecordCount > 0 Then
           .rows = rs2.RecordCount + 1
           rs2.MoveFirst

            For i = 1 To .rows - 1
               .TextMatrix(i, .ColIndex("Ser")) = i
   If SystemOptions.UserInterface = EnglishInterface Then
   .TextMatrix(i, .ColIndex("group")) = IIf(IsNull(rs2.Fields("NameE").value), "", rs2.Fields("NameE").value)
      .TextMatrix(i, .ColIndex("empname")) = IIf(IsNull(rs2.Fields("Emp_Namee").value), "", rs2.Fields("Emp_Namee").value)
      Else
      .TextMatrix(i, .ColIndex("group")) = IIf(IsNull(rs2.Fields("name").value), "", rs2.Fields("name").value)
      .TextMatrix(i, .ColIndex("empname")) = IIf(IsNull(rs2.Fields("Emp_Name").value), "", rs2.Fields("Emp_Name").value)
 
    End If
     .TextMatrix(i, .ColIndex("values")) = val(IIf(IsNull(rs2.Fields("valu").value), "", rs2.Fields("valu").value))
 .TextMatrix(i, .ColIndex("rate")) = val(IIf(IsNull(rs2.Fields("rate").value), "", rs2.Fields("rate").value))
  .TextMatrix(i, .ColIndex("code")) = IIf(IsNull(rs2.Fields("Fullcode").value), "", rs2.Fields("Fullcode").value)
  .TextMatrix(i, .ColIndex("id")) = IIf(IsNull(rs2.Fields("EmpID").value), "", rs2.Fields("EmpID").value)
  .TextMatrix(i, .ColIndex("idd")) = IIf(IsNull(rs2.Fields("idd").value), "", rs2.Fields("idd").value)
  .TextMatrix(i, .ColIndex("groupid")) = IIf(IsNull(rs2.Fields("GroupID").value), "", rs2.Fields("GroupID").value)
        rs2.MoveNext
            Next i

         
        End If

        .RowHeight(-1) = 300
    End With
''/

 
    
    '-----------------------------------------------------------------------------
    If DcboDebitSide.BoundText = "" And DcboCreditSide.BoundText = "" Then
    If SystemOptions.SysAppAccoutingType = CompeleteAccounting Then
        StrSQL = "Select * From DOUBLE_ENTREY_VOUCHERS Where Notes_ID=" & val(Me.XPTxtID.Text)
        StrSQL = StrSQL + " Order By DEV_ID_Line_No "
        Set RsDev = New ADODB.Recordset
        RsDev.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

        If Not (RsDev.BOF Or rs.EOF) Then
            Me.LblDevID.Caption = RsDev("Double_Entry_Vouchers_ID").value
            Me.lbl(33).Caption = RsDev("Account_Interval_ID").value
            RsDev.MoveFirst

            For i = 1 To 2 ' RsDev.RecordCount

                If RsDev("Credit_Or_Debit").value = 0 Then
                    Me.DcboDebitSide.BoundText = RsDev("Account_Code").value
                ElseIf RsDev("Credit_Or_Debit").value = 1 Then
                    Me.DcboCreditSide.BoundText = RsDev("Account_Code").value
                End If

                RsDev.MoveNext
            Next i

        End If
    End If
End If
    '-----------------------------------------------------------------------------
    ChkTrans_Click
    '⁄—÷ «·„” Œ·’« 
    'If DCboCashType.ListIndex = 5 Then
    FillGridWithData val(Me.DBCboClientName.BoundText), TxtNoteSerial.Text
    '⁄—÷ «·«ﬁ”« ÿ ·⁄ﬁÊœ  « ··«ÌÃ«—
       FillGridWithDataContract txtContractNo.Text, val(XPTxtID.Text)
       ReLineGrid
    '  End If
    
   FillGridWithData1 val(TxtContNo.Text)
If val(Me.DcbIqara.BoundText) <> 0 Then
GetAmola val(Me.DcbIqara.BoundText)
End If
 ReLineGrid
 If val(DCboCashType.ListIndex) >= 7 Then
 DCboCashType2.ListIndex = val(DCboCashType.ListIndex) - 7
 End If
 DCboCashType_Change
 CalCulteRemainElec
    XPTxtCurrent.Caption = rs.AbsolutePosition
    XPTxtCount.Caption = rs.RecordCount
    
    Exit Sub
ErrTrap:
End Sub
Sub GetTotalPayedElect()
Dim rs2 As ADODB.Recordset
Set rs2 = New ADODB.Recordset
Dim sql As String
sql = " SELECT    SUM(Price) AS SumPrice, SUM(Insurance) AS SumInsurance, SUM(Maintenance) AS SumMaintenance, SUM(MaintCondition) AS SumMaintCondition, SUM(MaintClean) AS SumMaintClean,"
sql = sql & "                       SUM(Paints) AS SumPaints, SUM(Maintkitchen) AS SumMaintkitchen, SUM(Electricity1) AS SumElectricity1, SUM(MaintDoors) AS SumMaintDoors, SUM(Windows)"
sql = sql & "                      AS SumWindows, SUM(MaintOther) AS SumMaintOther, SUM(TotalAftreIns) AS SumTotalAftreIns, SUM(RemainRent) AS SumRemainRent, SUM(Net) AS SumNet"
sql = sql & " From dbo.Notes"
sql = sql & " Where (CashingType = 13) And (ContNo = " & val(TxtContNo.Text) & ")"
sql = sql & " GROUP BY ContNo"
rs2.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
If rs2.RecordCount > 0 Then
TxtMaintenance3.Text = IIf(IsNull(rs2("SumMaintenance").value), 0, rs2("SumMaintenance").value)
TxtRemainRent3.Text = IIf(IsNull(rs2("SumRemainRent").value), 0, rs2("SumRemainRent").value)
TxtMaintCondition3.Text = IIf(IsNull(rs2("SumMaintCondition").value), 0, rs2("SumMaintCondition").value)
TxtMaintClean3.Text = IIf(IsNull(rs2("SumMaintClean").value), 0, rs2("SumMaintClean").value)
TxtPaints3.Text = IIf(IsNull(rs2("SumPaints").value), 0, rs2("SumPaints").value)
TxtMaintkitchen3.Text = IIf(IsNull(rs2("SumMaintkitchen").value), 0, rs2("SumMaintkitchen").value)
TxtElectricity13.Text = IIf(IsNull(rs2("SumElectricity1").value), 0, rs2("SumElectricity1").value)
TxtMaintDoors3.Text = IIf(IsNull(rs2("SumMaintDoors").value), 0, rs2("SumMaintDoors").value)
TxtWindows3.Text = IIf(IsNull(rs2("SumWindows").value), 0, rs2("SumWindows").value)
TxtMaintOther3.Text = IIf(IsNull(rs2("SumMaintOther").value), 0, rs2("SumMaintOther").value)
TxtPrice3.Text = IIf(IsNull(rs2("SumPrice").value), 0, rs2("SumPrice").value)
'TxtTotal23.Text = val(TxtWindows3.Text) + val(TxtMaintOther3.Text) + val(TxtMaintDoors3.Text) + val(TxtElectricity13.Text) + val(TxtMaintkitchen3.Text) + val(TxtPaints3.Text) + val(TxtMaintCondition3.Text) + val(TxtMaintClean3.Text) + val(TxtMaintenance3.Text) + val(TxtRemainRent3.Text)
Else
TxtPrice3.Text = 0
'TxtTotal23.Text = val(TxtInsurance.Text) + val(TxtDiscount.Text)
TxtMaintenance3.Text = 0
TxtRemainRent3.Text = 0
TxtMaintCondition3.Text = 0
TxtMaintClean3.Text = 0
TxtPaints3.Text = 0
TxtMaintkitchen3.Text = 0
TxtElectricity13.Text = 0
TxtMaintDoors3.Text = 0
TxtWindows3.Text = 0
TxtMaintOther3.Text = 0
End If
Dim totalPayed As Double
FlgNew = False
Calculte
 totalPayed = val(txtInsurance.Text) + val(txtDiscount.Text)
If totalPayed > 0 Then
FlgNew = True
If val(TxtMaintenance2.Text) >= totalPayed And totalPayed > 0 Then
TxtMaintenance3.Text = val(TxtMaintenance3.Text) + totalPayed
totalPayed = 0
ElseIf totalPayed > 0 Then
TxtMaintenance3.Text = val(TxtMaintenance3.Text) + val(TxtMaintenance2.Text)
totalPayed = totalPayed - val(TxtMaintenance2.Text)
End If
If val(TxtRemainRent2.Text) >= totalPayed And totalPayed > 0 Then
TxtRemainRent3.Text = val(TxtRemainRent3.Text) + totalPayed
totalPayed = 0
ElseIf totalPayed > 0 Then
TxtRemainRent3.Text = val(TxtRemainRent3.Text) + val(TxtRemainRent2.Text)
totalPayed = totalPayed - val(TxtRemainRent2.Text)
End If
''''
If val(TxtMaintCondition2.Text) >= totalPayed And totalPayed > 0 Then
TxtMaintCondition3.Text = val(TxtMaintCondition3.Text) + totalPayed
totalPayed = 0
ElseIf totalPayed > 0 Then
TxtMaintCondition3.Text = val(TxtMaintCondition3.Text) + val(TxtMaintCondition2.Text)
totalPayed = totalPayed - val(TxtMaintCondition2.Text)
End If
If val(TxtMaintClean2.Text) >= totalPayed And totalPayed > 0 Then
TxtMaintClean3.Text = val(TxtMaintClean3.Text) + totalPayed
totalPayed = 0
ElseIf totalPayed > 0 Then
TxtMaintClean3.Text = val(TxtMaintClean3.Text) + val(TxtMaintClean2.Text)
totalPayed = totalPayed - val(TxtMaintClean2.Text)
End If
''//
If val(TxtPaints2.Text) >= totalPayed And totalPayed > 0 Then
TxtPaints3.Text = val(TxtPaints3.Text) + totalPayed
totalPayed = 0
ElseIf totalPayed > 0 Then
TxtPaints3.Text = val(TxtPaints3.Text) + val(TxtPaints2.Text)
totalPayed = totalPayed - val(TxtPaints2.Text)
End If
''//
If val(TxtMaintkitchen2.Text) >= totalPayed And totalPayed > 0 Then
TxtMaintkitchen3.Text = val(TxtMaintkitchen3.Text) + totalPayed
totalPayed = 0
ElseIf totalPayed > 0 Then
TxtMaintkitchen3.Text = val(TxtMaintkitchen3.Text) + val(TxtMaintkitchen2.Text)
totalPayed = totalPayed - val(TxtMaintkitchen2.Text)
End If
''//
If val(TxtElectricity12.Text) >= totalPayed And totalPayed > 0 Then
TxtElectricity13.Text = val(TxtElectricity13.Text) + totalPayed
totalPayed = 0
ElseIf totalPayed > 0 Then
TxtElectricity13.Text = val(TxtElectricity13.Text) + val(TxtElectricity12.Text)
totalPayed = totalPayed - val(TxtElectricity12.Text)
End If
If val(TxtMaintDoors2.Text) >= totalPayed And totalPayed > 0 Then
TxtMaintDoors3.Text = val(TxtMaintDoors3.Text) + totalPayed
totalPayed = 0
ElseIf totalPayed > 0 Then
TxtMaintDoors3.Text = val(TxtMaintDoors3.Text) + val(TxtMaintDoors2.Text)
totalPayed = totalPayed - val(TxtMaintDoors2.Text)
End If
''//
If val(TxtWindows2.Text) >= totalPayed And totalPayed > 0 Then
TxtWindows3.Text = val(TxtWindows3.Text) + totalPayed
totalPayed = 0
ElseIf totalPayed > 0 Then
TxtWindows3.Text = val(TxtWindows3.Text) + val(TxtWindows2.Text)
totalPayed = totalPayed - val(TxtWindows2.Text)
End If
''//
If val(TxtMaintOther2.Text) >= totalPayed And totalPayed > 0 Then
TxtMaintOther3.Text = val(TxtMaintOther3.Text) + totalPayed
totalPayed = 0
ElseIf totalPayed > 0 Then
TxtMaintOther3.Text = val(TxtMaintOther3.Text) + val(TxtMaintOther2.Text)
totalPayed = totalPayed - val(TxtMaintOther2.Text)
End If
End If
CalCulteRemainElec
End Sub
Sub CalCulteRemainElec()
TxtRemMaintenance.Text = val(TxtMaintenance2.Text) - val(TxtMaintenance3.Text)
TxtRemRemainRent.Text = val(TxtRemainRent2.Text) - val(TxtRemainRent3.Text)
TxtRemMaintCondition.Text = val(TxtMaintCondition2.Text) - val(TxtMaintCondition3.Text)
TxtRemMaintClean.Text = val(TxtMaintClean2.Text) - val(TxtMaintClean3.Text)
TxtRemPaints.Text = val(TxtPaints2.Text) - val(TxtPaints3.Text)
TxtRemMaintkitchen.Text = val(TxtMaintkitchen2.Text) - val(TxtMaintkitchen3.Text)
TxtRemElectricity.Text = val(TxtElectricity12.Text) - val(TxtElectricity13.Text)
TxtRemMaintDoors.Text = val(TxtMaintDoors2.Text) - val(TxtMaintDoors3.Text)
TxtRemWindows.Text = val(TxtWindows2.Text) - val(TxtWindows3.Text)
TxtRemMaintOther.Text = val(TxtMaintOther2.Text) - val(TxtMaintOther3.Text)
TxtRemPrice.Text = val(TxtPrice2.Text) - val(TxtPrice3.Text)
End Sub
Sub GetCommInformation(Optional ByRef EmpID As Double, Optional ByRef Amount As Double)
Dim sql As String
Dim rs2 As ADODB.Recordset
Set rs2 = New ADODB.Recordset
sql = "Select * from TblAqarCommissions where NoteID=" & val(XPTxtID.Text) & " "
rs2.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
If rs2.RecordCount > 0 Then
EmpID = IIf(IsNull(rs2("EmpID").value), 0, rs2("EmpID").value)
Amount = IIf(IsNull(rs2("Amount").value), 0, rs2("Amount").value)
Else
EmpID = 0
Amount = 0
End If
End Sub
 Function OtherOwnerNoreatJlInContract(LngDevID As Long, notes_id As Double) As Double
netVatPayed = 0
If DCboCashType.ListIndex <> 8 Then Exit Function
Dim total_value As Double
Dim cProgress As ClsProgress
Set cProgress = New ClsProgress
    cProgress.ProgressType = Waiting
 Dim foxy_ked_NO As String
 Dim credit_side As String
 Dim My_SQL As String
 Dim Line1 As Double
 
'         lineno = 1
 Dim AccountCode As String
    cProgress.StartProgress
    DoEvents
    LngDevID = new_id("DOUBLE_ENTREY_VOUCHERS", "Double_Entry_Vouchers_ID", "")
    Dim Msgdes As String
    Dim CURRENT_LINE As Double
    Dim depit_side As String
    Dim Msg As String
     Dim i As Integer
Msgdes = "»‰«¡ ⁄·Ï „ﬁ»Ê÷«  «„·«ﬂ «·€Ì— —ﬁ„ " & TxtNoteSerial1.Text & " "
Dim Percetage As Double
Dim commissionvalue As Double
Dim vaTAccount As String
'salimhere
                            ''///  «·ﬁÌ„… «·„÷«›… ··”⁄Ì
                            PercentgValueAddedAccount_Transec XPDtbTrans.value, 8, 1, AccountCode, Percetage
vaTAccount = AccountCode ''///  «·ﬁÌ„… «·„÷«›… ··”⁄Ì
'salimhere
    With Grid3
Msg = XPMTxtRemarks.Text & CHR(13) & Msgdes
depit_side = ModAccounts.GetMyAccountCode("TblCustemers", "CusID", val(DBCboClientName.BoundText), "Account_code")
        For i = .FixedRows To .rows - 1
 
         If .cell(flexcpChecked, i, .ColIndex("Select")) = flexChecked Then
         '—’Ìœ „ »ﬁÌ ⁄·Ì «·„” √Ã—   OldValuePayed
        'salimhere
        If val(.TextMatrix(i, .ColIndex("OldValuePayed"))) > 0 Then
                 total_value = Round((.TextMatrix(i, .ColIndex("OldValuePayed"))), 2)
               Else
               total_value = 0
               End If
               
                If total_value > 0 Then
                lineno = lineno + 1
                 If ModAccounts.AddNewDev(LngDevID, lineno, depit_side, total_value, 0, Msg & " " & "—’Ìœ „ »ﬁÌ ⁄·Ì «·„” √Ã—", val(notes_id), , , , XPDtbTrans.value, user_id, , , , , , , , , CURRENT_LINE, , , , , , , , , val(Dcbranch.BoundText), , , , , , , , , , val(DcbIqara.BoundText), val(DcbUnitType.BoundText), val(DcbUnitNo.BoundText)) = False Then
                        GoTo ErrTrap
                    End If
                    lineno = lineno + 1
                    
                
               
                    If SystemOptions.OpenAccountAqar = False Then
                        AccountCode = ModAccounts.GetMyAccountCode("TblCustemers", "CusID", val(Txtownerid), "Account_code")
                    Else
                        AccountCode = GetAqarAcountCode(val(DcbIqara.BoundText))
                    End If
                  '  AccountCode = ModAccounts.GetMyAccountCode("TblCustemers", "CusID", val(Txtownerid), "Account_code")
                      If ModAccounts.AddNewDev(LngDevID, lineno, AccountCode, total_value, 1, Msg & " " & "—’Ìœ „ »ﬁÌ ⁄·Ì «·„” √Ã—", val(notes_id), , , , XPDtbTrans.value, user_id, , , , , , , , , CURRENT_LINE, , , , , , , , , val(Dcbranch.BoundText), , , , , , , , , , val(DcbIqara.BoundText), val(DcbUnitType.BoundText), val(DcbUnitNo.BoundText)) = False Then
                        GoTo ErrTrap
                    End If
                      lineno = lineno + 1
              
         
                    
              End If
              
       'salimhere
         If val(.TextMatrix(i, .ColIndex("RentValuePayed"))) > 0 Then
                total_value = Round(.TextMatrix(i, .ColIndex("RentValuePayed")), 2)
          Else
          total_value = 0
          End If
             If total_value > 0 Then
             
                    If ModAccounts.AddNewDev(LngDevID, lineno, depit_side, total_value, 0, Msg & " " & "«·«ÌÃ«—", val(notes_id), , , , XPDtbTrans.value, user_id, , , , , , , , , CURRENT_LINE, , , , , , , , , val(Dcbranch.BoundText), , , , , , , , , , val(DcbIqara.BoundText), val(DcbUnitType.BoundText), val(DcbUnitNo.BoundText)) = False Then
                        GoTo ErrTrap
                    End If
                    lineno = lineno + 1
                    If SystemOptions.OpenAccountAqar = False Then
                        AccountCode = ModAccounts.GetMyAccountCode("TblCustemers", "CusID", val(Txtownerid), "Account_code")
                    Else
                        AccountCode = GetAqarAcountCode(val(DcbIqara.BoundText))
                    End If
                  '  AccountCode = ModAccounts.GetMyAccountCode("TblCustemers", "CusID", val(Txtownerid), "Account_code")
                      If ModAccounts.AddNewDev(LngDevID, lineno, AccountCode, total_value, 1, Msg & " " & "«·«ÌÃ«—", val(notes_id), , , , XPDtbTrans.value, user_id, , , , , , , , , CURRENT_LINE, , , , , , , , , val(Dcbranch.BoundText), , , , , , , , , , val(DcbIqara.BoundText), val(DcbUnitType.BoundText), val(DcbUnitNo.BoundText)) = False Then
                        GoTo ErrTrap
                    End If
                      lineno = lineno + 1
              End If
              ''''// «·”⁄Ì
            total_value = Round(.TextMatrix(i, .ColIndex("CommissionsPayed")), 2)
             If total_value > 0 Then
                          If Percetage <> 0 Then
                    total_value = total_value / (Percetage / 100 + 1)
                    End If
                    
                    If ModAccounts.AddNewDev(LngDevID, lineno, depit_side, total_value, 0, Msg & " " & "«·”⁄Ì", val(notes_id), , , , XPDtbTrans.value, user_id, , , , , , , , , CURRENT_LINE, , , , , , , , , val(Dcbranch.BoundText), , , , , , , , , , val(DcbIqara.BoundText), val(DcbUnitType.BoundText), val(DcbUnitNo.BoundText)) = False Then
                        GoTo ErrTrap
                    End If
                    lineno = lineno + 1
       
                     AccountCode = get_account_code_branch(81, my_branch)
     
                    
                         If ModAccounts.AddNewDev(LngDevID, lineno, AccountCode, total_value, 1, Msg & " " & "«·”⁄Ì", val(notes_id), , , , XPDtbTrans.value, user_id, , , , , , , , , CURRENT_LINE, , , , , , , , , val(Dcbranch.BoundText), , , , , , , , , , val(DcbIqara.BoundText), val(DcbUnitType.BoundText), val(DcbUnitNo.BoundText)) = False Then
                        GoTo ErrTrap
                    End If
                      lineno = lineno + 1
              End If
                            
              
              ''///«·„Ì«Â
                total_value = Round(.TextMatrix(i, .ColIndex("WaterPayed")), 2)
             If total_value > 0 Then
                    If ModAccounts.AddNewDev(LngDevID, lineno, depit_side, total_value, 0, Msg & " " & "«·„Ì«Â", val(notes_id), , , , XPDtbTrans.value, user_id, , , , , , , , , CURRENT_LINE, , , , , , , , , val(Dcbranch.BoundText), , , , , , , , , , val(DcbIqara.BoundText), val(DcbUnitType.BoundText), val(DcbUnitNo.BoundText)) = False Then
                        GoTo ErrTrap
                    End If
                    lineno = lineno + 1
                   'salimhere  AccountCode = get_account_code_branch(83, my_branch)
                             If SystemOptions.OpenAccountAqar = False Then
                        AccountCode = ModAccounts.GetMyAccountCode("TblCustemers", "CusID", val(Txtownerid), "Account_code")
                    Else
                        AccountCode = GetAqarAcountCode(val(DcbIqara.BoundText))
                    End If
                    'salimhere
                         If ModAccounts.AddNewDev(LngDevID, lineno, AccountCode, total_value, 1, Msg & " " & "«·„Ì«Â", val(notes_id), , , , XPDtbTrans.value, user_id, , , , , , , , , CURRENT_LINE, , , , , , , , , val(Dcbranch.BoundText), , , , , , , , , , val(DcbIqara.BoundText), val(DcbUnitType.BoundText), val(DcbUnitNo.BoundText)) = False Then
                        GoTo ErrTrap
                    End If
                      lineno = lineno + 1
              End If
              
                          ''///«·ﬂÂ—»«¡
                total_value = Round(.TextMatrix(i, .ColIndex("ElectricPayed")), 2)
             If total_value > 0 Then
                    If ModAccounts.AddNewDev(LngDevID, lineno, depit_side, total_value, 0, Msg & " " & "«·ﬂÂ—»«¡", val(notes_id), , , , XPDtbTrans.value, user_id, , , , , , , , , CURRENT_LINE, , , , , , , , , val(Dcbranch.BoundText), , , , , , , , , , val(DcbIqara.BoundText), val(DcbUnitType.BoundText), val(DcbUnitNo.BoundText)) = False Then
                        GoTo ErrTrap
                    End If
                    lineno = lineno + 1
        ' salimhere            AccountCode = get_account_code_branch(84, my_branch)
                            If SystemOptions.OpenAccountAqar = False Then
                        AccountCode = ModAccounts.GetMyAccountCode("TblCustemers", "CusID", val(Txtownerid), "Account_code")
                    Else
                        AccountCode = GetAqarAcountCode(val(DcbIqara.BoundText))
                    End If
                    'salimhere
                    
                    If ModAccounts.AddNewDev(LngDevID, lineno, AccountCode, total_value, 1, Msg & " " & "«·ﬂÂ—»«¡", val(notes_id), , , , XPDtbTrans.value, user_id, , , , , , , , , CURRENT_LINE, , , , , , , , , val(Dcbranch.BoundText), , , , , , , , , , val(DcbIqara.BoundText), val(DcbUnitType.BoundText), val(DcbUnitNo.BoundText)) = False Then
                        GoTo ErrTrap
                    End If
                      lineno = lineno + 1
              End If
                                     ''///«·Œœ„« 
                total_value = Round(.TextMatrix(i, .ColIndex("TelandNetPayed")), 2)
                
             If total_value > 0 Then
                       If Percetage <> 0 Then
                    total_value = total_value / (Percetage / 100 + 1)
                    End If
                    
                    If ModAccounts.AddNewDev(LngDevID, lineno, depit_side, total_value, 0, Msg & " " & "«·Œœ„« .", val(notes_id), , , , XPDtbTrans.value, user_id, , , , , , , , , CURRENT_LINE, , , , , , , , , val(Dcbranch.BoundText), , , , , , , , , , val(DcbIqara.BoundText), val(DcbUnitType.BoundText), val(DcbUnitNo.BoundText)) = False Then
                        GoTo ErrTrap
                    End If
                    lineno = lineno + 1
               
                     AccountCode = get_account_code_branch(85, my_branch)
                         If ModAccounts.AddNewDev(LngDevID, lineno, AccountCode, total_value, 1, Msg & " " & "«·Œœ„« ...", val(notes_id), , , , XPDtbTrans.value, user_id, , , , , , , , , CURRENT_LINE, , , , , , , , , val(Dcbranch.BoundText), , , , , , , , , , val(DcbIqara.BoundText), val(DcbUnitType.BoundText), val(DcbUnitNo.BoundText)) = False Then
                        GoTo ErrTrap
                    End If
                      lineno = lineno + 1
              End If
              
                                                   ''///«· «„Ì‰
                total_value = Round(.TextMatrix(i, .ColIndex("InsurancePayed")), 2)
             If total_value > 0 Then
                    If ModAccounts.AddNewDev(LngDevID, lineno, depit_side, total_value, 0, Msg & " " & "«· «„Ì‰", val(notes_id), , , , XPDtbTrans.value, user_id, , , , , , , , , CURRENT_LINE, , , , , , , , , val(Dcbranch.BoundText), , , , , , , , , , val(DcbIqara.BoundText), val(DcbUnitType.BoundText), val(DcbUnitNo.BoundText)) = False Then
                        GoTo ErrTrap
                    End If
                    lineno = lineno + 1
                    If SystemOptions.OpenAccountAqar = False Then
                        AccountCode = ModAccounts.GetMyAccountCode("TblCustemers", "CusID", val(Txtownerid), "Account_code")
                    Else
                        AccountCode = GetAqarAcountCode(val(DcbIqara.BoundText))
                    End If
                    ' AccountCode = ModAccounts.GetMyAccountCode("TblCustemers", "CusID", val(Txtownerid), "Account_code")
                         If ModAccounts.AddNewDev(LngDevID, lineno, AccountCode, total_value, 1, Msg & " " & "«· «„Ì‰", val(notes_id), , , , XPDtbTrans.value, user_id, , , , , , , , , CURRENT_LINE, , , , , , , , , val(Dcbranch.BoundText), , , , , , , , , , val(DcbIqara.BoundText), val(DcbUnitType.BoundText), val(DcbUnitNo.BoundText)) = False Then
                        GoTo ErrTrap
                    End If
                      lineno = lineno + 1
              End If
              

                total_value = val(.TextMatrix(i, .ColIndex("CommissionsPayed"))) / (Percetage / 100 + 1)
              commissionvalue = total_value * Percetage / 100
               
             If total_value > 0 And commissionvalue > 0 Then
                    If ModAccounts.AddNewDev(LngDevID, lineno, depit_side, commissionvalue, 0, Msg & " " & "«·ﬁÌ„… «·„÷«›… ··”⁄Ì", val(notes_id), , , , XPDtbTrans.value, user_id, , , , , , , , , CURRENT_LINE, , , , , , , , , val(Dcbranch.BoundText), , , , , , , , , , val(DcbIqara.BoundText), val(DcbUnitType.BoundText), val(DcbUnitNo.BoundText)) = False Then
                        GoTo ErrTrap
                    End If
                    lineno = lineno + 1
                     
                         If ModAccounts.AddNewDev(LngDevID, lineno, vaTAccount, commissionvalue, 1, Msg & " " & "«·ﬁÌ„… «·„÷«›… ··”⁄Ì/", val(notes_id), , , , XPDtbTrans.value, user_id, , , , , , , , , CURRENT_LINE, , , , , , , , , val(Dcbranch.BoundText), , , , , , , , , , val(DcbIqara.BoundText), val(DcbUnitType.BoundText), val(DcbUnitNo.BoundText)) = False Then
                        GoTo ErrTrap
                    End If
                      lineno = lineno + 1
              End If
              
              
            total_value = val(.TextMatrix(i, .ColIndex("TelandNetPayed"))) / (Percetage / 100 + 1)
              commissionvalue = total_value * Percetage / 100
               
             If total_value > 0 And commissionvalue > 0 Then
                    If ModAccounts.AddNewDev(LngDevID, lineno, depit_side, commissionvalue, 0, Msg & " " & "«·ﬁÌ„… «·„÷«›… ··Œœ„« ", val(notes_id), , , , XPDtbTrans.value, user_id, , , , , , , , , CURRENT_LINE, , , , , , , , , val(Dcbranch.BoundText), , , , , , , , , , val(DcbIqara.BoundText), val(DcbUnitType.BoundText), val(DcbUnitNo.BoundText)) = False Then
                        GoTo ErrTrap
                    End If
                    lineno = lineno + 1
                     
                         If ModAccounts.AddNewDev(LngDevID, lineno, vaTAccount, commissionvalue, 1, Msg & " " & "«·ﬁÌ„… «·„÷«›… ··Œœ„« /", val(notes_id), , , , XPDtbTrans.value, user_id, , , , , , , , , CURRENT_LINE, , , , , , , , , val(Dcbranch.BoundText), , , , , , , , , , val(DcbIqara.BoundText), val(DcbUnitType.BoundText), val(DcbUnitNo.BoundText)) = False Then
                        GoTo ErrTrap
                    End If
                      lineno = lineno + 1
              End If
              
              
              ''///«·ﬁÌ„… «·„÷«›…
              'salimhere
              
               total_value = Round(.TextMatrix(i, .ColIndex("RentValuePayed")), 2) + Round(.TextMatrix(i, .ColIndex("WaterPayed")), 2) + Round(.TextMatrix(i, .ColIndex("ElectricPayed")), 2) '+ Round(.TextMatrix(i, .ColIndex("TelandNetPayed")), 2)
           
              If ComResid(1).value = True Then
                 total_value = total_value * Percetage / 100
               Else
               total_value = 0
               End If
            
            'TxtVATValue = netVatPayed
             If total_value > 0 Then
          '   netVatPayed = netVatPayed + total_value
             
                    If ModAccounts.AddNewDev(LngDevID, lineno, depit_side, total_value, 0, Msg & " " & "«·ﬁÌ„… «·„÷«›… ", val(notes_id), , , , XPDtbTrans.value, user_id, , , , , , , , , CURRENT_LINE, , , , , , , , , val(Dcbranch.BoundText), , , , , , , , , , val(DcbIqara.BoundText), val(DcbUnitType.BoundText), val(DcbUnitNo.BoundText)) = False Then
                        GoTo ErrTrap
                    End If
                    lineno = lineno + 1
                     If SystemOptions.OpenAccountAqar = False Then
                        AccountCode = ModAccounts.GetMyAccountCode("TblCustemers", "CusID", val(Txtownerid), "Account_code")
                    Else
                        AccountCode = GetAqarAcountCode(val(DcbIqara.BoundText))
                    End If
                    '  AccountCode = ModAccounts.GetMyAccountCode("TblCustemers", "CusID", val(Txtownerid), "Account_code")
                     If ModAccounts.AddNewDev(LngDevID, lineno, AccountCode, total_value, 1, Msg & " " & "«·ﬁÌ„… «·„÷«›… ", val(notes_id), , , , XPDtbTrans.value, user_id, , , , , , , , , CURRENT_LINE, , , , , , , , , val(Dcbranch.BoundText), , , , , , , , , , val(DcbIqara.BoundText), val(DcbUnitType.BoundText), val(DcbUnitNo.BoundText)) = False Then
                        GoTo ErrTrap
                    End If
                      lineno = lineno + 1
              End If
              
              
                    
    End If
   
        

        Next i

    End With
    
    
    'salimhere  ﬁÌœ «·’‰œÊﬁ / «Ê «·»‰ﬂ
                 total_value = val(XPTxtVal.Text) + val(TxtVATValue)
             If total_value > 0 Then
                    If ModAccounts.AddNewDev(LngDevID, lineno, DcboDebitSide.BoundText, total_value, 0, Msg & " " & "«·’‰œÊﬁ/«·»‰ﬂ", val(notes_id), , , , XPDtbTrans.value, user_id, , , , , , , , , CURRENT_LINE, , , , , , , , , val(Dcbranch.BoundText), , , , , , , , , , val(DcbIqara.BoundText), val(DcbUnitType.BoundText), val(DcbUnitNo.BoundText)) = False Then
                        GoTo ErrTrap
                    End If
                    lineno = lineno + 1
                     
                         If ModAccounts.AddNewDev(LngDevID, lineno, depit_side, total_value, 1, Msg & " " & "«·„” «Ã—", val(notes_id), , , , XPDtbTrans.value, user_id, , , , , , , , , CURRENT_LINE, , , , , , , , , val(Dcbranch.BoundText), , , , , , , , , , val(DcbIqara.BoundText), val(DcbUnitType.BoundText), val(DcbUnitNo.BoundText)) = False Then
                        GoTo ErrTrap
                    End If
                      lineno = lineno + 1
              End If

    DoEvents
    cProgress.FinishProgress
    cProgress.StopProgess
    Set cProgress = Nothing
    
ErrTrap:
End Function
'
'
'
'
 Function MyOwnerNoreatJlInContract(LngDevID As Long, notes_id As Double) As Double

If DCboCashType.ListIndex <> 8 Then Exit Function
Dim total_value As Double

 Dim foxy_ked_NO As String
 Dim credit_side As String
 Dim My_SQL As String
 Dim Line1 As Double
 
         
 Dim AccountCode As String
    
    DoEvents
    LngDevID = new_id("DOUBLE_ENTREY_VOUCHERS", "Double_Entry_Vouchers_ID", "")
    Dim Msgdes As String
    Dim CURRENT_LINE As Double
    Dim depit_side As String
    Dim Msg As String
     Dim i As Integer
Msgdes = "»‰«¡ ⁄·Ï „ﬁ»Ê÷«  «„·«ﬂÌ —ﬁ„ " & TxtNoteSerial1.Text & " "
Dim Percetage As Double
Dim commissionvalue As Double
Dim AccountCodeVat As String
Dim vaTAccount As String
                            ''///  «·ﬁÌ„… «·„÷«›… ··”⁄Ì
                            PercentgValueAddedAccount_Transec XPDtbTrans.value, 8, 1, AccountCodeVat, Percetage
                                         
vaTAccount = AccountCodeVat ''///  ?????? ??????? ?????



Msg = XPMTxtRemarks.Text & CHR(13) & Msgdes
depit_side = ModAccounts.GetMyAccountCode("TblCustemers", "CusID", val(DBCboClientName.BoundText), "Account_code")
total_value = val(XPTxtVal)
     'salimhere  ﬁÌœ «·’‰œÊﬁ / «Ê «·»‰ﬂ
                 total_value = val(XPTxtVal.Text) + val(TxtVATValue)
             If total_value > 0 Then
             lineno = lineno + 1
                    If ModAccounts.AddNewDev(LngDevID, lineno, DcboDebitSide.BoundText, total_value, 0, Msg & " " & "«·’‰œÊﬁ/«·»‰ﬂ", val(notes_id), , , , XPDtbTrans.value, user_id, , , , , , , , , CURRENT_LINE, , , , , , , , , val(Dcbranch.BoundText), , , , , , , , , , 0, val(DcbUnitType.BoundText), val(DcbUnitNo.BoundText)) = False Then
                        GoTo ErrTrap
                    End If
                    lineno = lineno + 1
                     
                         If ModAccounts.AddNewDev(LngDevID, lineno, depit_side, total_value, 1, Msg & " " & "«·„” «Ã—", val(notes_id), , , , XPDtbTrans.value, user_id, , , , , , , , , CURRENT_LINE, , , , , , , , , val(Dcbranch.BoundText), , , , , , , , , , val(DcbIqara.BoundText), val(DcbUnitType.BoundText), val(DcbUnitNo.BoundText)) = False Then
                        GoTo ErrTrap
                    End If
                      lineno = lineno + 1
              End If
              
'              If val(TxtVATValue) <> 0 Then
'                lineno = lineno + 1
'                    If ModAccounts.AddNewDev(LngDevID, lineno, AccountCodeVat, val(TxtVATValue), 0, Msg & " " & "«·ﬁÌ„… «·„÷«›…", val(notes_id), , , , XPDtbTrans.value, user_id, , , , , , , , , CURRENT_LINE, , , , , , , , , val(dcBranch.BoundText), , , , , , , , , , 0, val(DcbUnitType.BoundText), val(DcbUnitNo.BoundText)) = False Then
'                        GoTo ErrTrap
'                    End If
'                    lineno = lineno + 1
'
'                        If ModAccounts.AddNewDev(LngDevID, lineno, depit_side, val(TxtVATValue), 1, Msg & " " & "«·ﬁÌ„… «·„÷«›… ··„” «Ã—", val(notes_id), , , , XPDtbTrans.value, user_id, , , , , , , , , CURRENT_LINE, , , , , , , , , val(dcBranch.BoundText), , , , , , , , , , val(DcbIqara.BoundText), val(DcbUnitType.BoundText), val(DcbUnitNo.BoundText)) = False Then
'                        GoTo ErrTrap
'                    End If
'
'              End If
 
 
    DoEvents

    
ErrTrap:
End Function


 Function MyOwnerNoreatJlInContractOld(LngDevID As Long, notes_id As Double) As Double

If DCboCashType.ListIndex <> 8 Then Exit Function
Dim total_value As Double
Dim cProgress As ClsProgress
Set cProgress = New ClsProgress
    cProgress.ProgressType = Waiting
 Dim foxy_ked_NO As String
 Dim credit_side As String
 Dim My_SQL As String
 Dim Line1 As Double

 Dim AccountCode As String
    cProgress.StartProgress
    DoEvents
    LngDevID = new_id("DOUBLE_ENTREY_VOUCHERS", "Double_Entry_Vouchers_ID", "")
    Dim Msgdes As String
    Dim CURRENT_LINE As Double
    Dim depit_side As String
    Dim Msg As String
     Dim i As Integer
Msgdes = "»‰«¡ ⁄·Ï „ﬁ»Ê÷«  «„·«ﬂÌ —ﬁ„ " & TxtNoteSerial1.Text & " "
Dim Percetage As Double
Dim commissionvalue As Double
Dim AccountCodeVat As String
Dim vaTAccount As String
                            ''///  «·ﬁÌ„… «·„÷«›… ··”⁄Ì
                            PercentgValueAddedAccount_Transec XPDtbTrans.value, 8, 1, AccountCodeVat, Percetage
                                         
vaTAccount = AccountCodeVat ''///  ?????? ??????? ?????



    With Grid3
Msg = XPMTxtRemarks.Text & CHR(13) & Msgdes
depit_side = ModAccounts.GetMyAccountCode("TblCustemers", "CusID", val(DBCboClientName.BoundText), "Account_code")
        For i = .FixedRows To .rows - 1 'salimher  -2
 
         If .cell(flexcpChecked, i, .ColIndex("Select")) = flexChecked Then
             commissionvalue = val(.TextMatrix(i, .ColIndex("CommissionsPayed"))) * Percetage / 100
              commissionvalue = Round(commissionvalue, 2)
                total_value = Round(.TextMatrix(i, .ColIndex("RentValuePayed")), 2)
               If total_value > 0 Then
                    If ModAccounts.AddNewDev(LngDevID, lineno, depit_side, total_value, 0, Msg & " " & "«·«ÌÃ«—", val(notes_id), , , , XPDtbTrans.value, user_id, , , , , , , , , CURRENT_LINE, , , , , , , , , val(Dcbranch.BoundText), , , , , , , , , , 0, val(DcbUnitType.BoundText), val(DcbUnitNo.BoundText)) = False Then
                        GoTo ErrTrap
                    End If
                    lineno = lineno + 1
              End If
              '''//////ﬂÂ—»«¡
              total_value = Round(.TextMatrix(i, .ColIndex("ElectricPayed")), 2)
             If total_value > 0 Then
                    If ModAccounts.AddNewDev(LngDevID, lineno, depit_side, total_value, 0, Msg & " " & "«·ﬂÂ—»«¡", val(notes_id), , , , XPDtbTrans.value, user_id, , , , , , , , , CURRENT_LINE, , , , , , , , , val(Dcbranch.BoundText), , , , , , , , , , 0, val(DcbUnitType.BoundText), val(DcbUnitNo.BoundText)) = False Then
                        GoTo ErrTrap
                    End If
                    lineno = lineno + 1
             End If
             '////„Ì«Â
              total_value = Round(.TextMatrix(i, .ColIndex("WaterPayed")), 2)
             If total_value > 0 Then
                    If ModAccounts.AddNewDev(LngDevID, lineno, depit_side, total_value, 0, Msg & " " & "«·„Ì«Â", val(notes_id), , , , XPDtbTrans.value, user_id, , , , , , , , , CURRENT_LINE, , , , , , , , , val(Dcbranch.BoundText), , , , , , , , , , 0, val(DcbUnitType.BoundText), val(DcbUnitNo.BoundText)) = False Then
                        GoTo ErrTrap
                    End If
                    lineno = lineno + 1
             End If
             ''///////////Œœ„« 
                           total_value = Round(.TextMatrix(i, .ColIndex("TelandNetPayed")), 2)
             If total_value > 0 Then
                    If ModAccounts.AddNewDev(LngDevID, lineno, depit_side, total_value, 0, Msg & " " & "Œœ„« ", val(notes_id), , , , XPDtbTrans.value, user_id, , , , , , , , , CURRENT_LINE, , , , , , , , , val(Dcbranch.BoundText), , , , , , , , , , 0, val(DcbUnitType.BoundText), val(DcbUnitNo.BoundText)) = False Then
                        GoTo ErrTrap
                    End If
                    lineno = lineno + 1
                    
                    
                            total_value = val(.TextMatrix(i, .ColIndex("TelandNetPayed"))) / (Percetage / 100 + 1)
              commissionvalue = total_value * Percetage / 100
                  
                  
                         If ModAccounts.AddNewDev(LngDevID, lineno, vaTAccount, commissionvalue, 1, Msg & " " & " ﬁ „ Œœ„«  ", val(notes_id), , , , XPDtbTrans.value, user_id, , , , , , , , , CURRENT_LINE, , , , , , , , , val(Dcbranch.BoundText), , , , , , , , , , 0, val(DcbUnitType.BoundText), val(DcbUnitNo.BoundText)) = False Then
                        GoTo ErrTrap
                    End If
                      lineno = lineno + 1
                      
             End If
                          ''///////////”⁄Ì
                           total_value = Round(.TextMatrix(i, .ColIndex("CommissionsPayed")), 2)
         
        'salimhere
 
             
             
        
             If total_value > 0 And commissionvalue > 0 Then
                    If ModAccounts.AddNewDev(LngDevID, lineno, depit_side, total_value, 0, Msg & " " & "    ”⁄Ì", val(notes_id), , , , XPDtbTrans.value, user_id, , , , , , , , , CURRENT_LINE, , , , , , , , , val(Dcbranch.BoundText), , , , , , , , , , 0, val(DcbUnitType.BoundText), val(DcbUnitNo.BoundText)) = False Then
                        GoTo ErrTrap
                    End If
                    lineno = lineno + 1
                     
                        total_value = val(.TextMatrix(i, .ColIndex("CommissionsPayed"))) / (Percetage / 100 + 1)
              commissionvalue = total_value * Percetage / 100
                  
                  
                         If ModAccounts.AddNewDev(LngDevID, lineno, vaTAccount, commissionvalue, 1, Msg & " " & " ﬁ „ ”⁄Ì ", val(notes_id), , , , XPDtbTrans.value, user_id, , , , , , , , , CURRENT_LINE, , , , , , , , , val(Dcbranch.BoundText), , , , , , , , , , 0, val(DcbUnitType.BoundText), val(DcbUnitNo.BoundText)) = False Then
                        GoTo ErrTrap
                    End If
                      lineno = lineno + 1
              End If

         
         '    If total_value > 0 Then
         '           If ModAccounts.AddNewDev(LngDevID, lineno, depit_side, total_value, 0, Msg & " " & "”⁄Ì", val(notes_id), , , , XPDtbTrans.value, user_id, , , , , , , , , CURRENT_LINE, , , , , , , , , val(Dcbranch.BoundText), , , , , , , , , , val(DcbIqara.BoundText), val(DcbUnitType.BoundText), val(DcbUnitNo.BoundText)) = False Then
         '               GoTo ErrTrap
         '           End If
         '           lineno = lineno + 1
         '    End If
             
             
             
             ''///ﬁÌ„… „÷«›…
                     total_value = Round(.TextMatrix(i, .ColIndex("RentValuePayed")), 2) + Round(.TextMatrix(i, .ColIndex("WaterPayed")), 2) + Round(.TextMatrix(i, .ColIndex("ElectricPayed")), 2) '+ Round(.TextMatrix(i, .ColIndex("TelandNetPayed")), 2)
              If ComResid(1).value = True Then
                 total_value = total_value * Percetage / 100
               Else
               total_value = 0
               End If
                                    
             If total_value > 0 Then
                    If ModAccounts.AddNewDev(LngDevID, lineno, depit_side, total_value, 0, Msg & " " & "«·ﬁÌ„… «·„÷«›…", val(notes_id), , , , XPDtbTrans.value, user_id, , , , , , , , , CURRENT_LINE, , , , , , , , , val(Dcbranch.BoundText), , , , , , , , , , 0, val(DcbUnitType.BoundText), val(DcbUnitNo.BoundText)) = False Then
                        GoTo ErrTrap
                    End If
                    lineno = lineno + 1
             End If

          '//////////////////////////////«·œ«∆‰
               
              '''//////////////////«·«ÌÃ«—
            total_value = Round(.TextMatrix(i, .ColIndex("RentValuePayed")), 2)
            AccountCode = get_account_code_branch(86, my_branch)
               If total_value > 0 Then
                 If ModAccounts.AddNewDev(LngDevID, lineno, AccountCode, total_value, 1, Msg & " " & "«·«ÌÃ«—", val(notes_id), , , , XPDtbTrans.value, user_id, , , , , , , , , CURRENT_LINE, , , , , , , , , val(Dcbranch.BoundText), , , , , , , , , , 0, val(DcbUnitType.BoundText), val(DcbUnitNo.BoundText)) = False Then
                        GoTo ErrTrap
                    End If
                      lineno = lineno + 1
              End If
               ''///«·ﬂÂ—»«¡
                total_value = Round(.TextMatrix(i, .ColIndex("ElectricPayed")), 2)
             If total_value > 0 Then
                     AccountCode = get_account_code_branch(84, my_branch)
                         If ModAccounts.AddNewDev(LngDevID, lineno, AccountCode, total_value, 1, Msg & " " & "«·ﬂÂ—»«¡", val(notes_id), , , , XPDtbTrans.value, user_id, , , , , , , , , CURRENT_LINE, , , , , , , , , val(Dcbranch.BoundText), , , , , , , , , , 0, val(DcbUnitType.BoundText), val(DcbUnitNo.BoundText)) = False Then
                        GoTo ErrTrap
                    End If
                      lineno = lineno + 1
              End If
                            ''///«·„Ì«Â
                total_value = Round(.TextMatrix(i, .ColIndex("WaterPayed")), 2)
             If total_value > 0 Then
                     AccountCode = get_account_code_branch(83, my_branch)
                         If ModAccounts.AddNewDev(LngDevID, lineno, AccountCode, total_value, 1, Msg & " " & "«·„Ì«Â", val(notes_id), , , , XPDtbTrans.value, user_id, , , , , , , , , CURRENT_LINE, , , , , , , , , val(Dcbranch.BoundText), , , , , , , , , , 0, val(DcbUnitType.BoundText), val(DcbUnitNo.BoundText)) = False Then
                        GoTo ErrTrap
                    End If
                      lineno = lineno + 1
              End If
                   total_value = Round(.TextMatrix(i, .ColIndex("TelandNetPayed")), 2)
                   ''///«·Œœ„« 
                     If Percetage <> 0 Then
                    total_value = total_value / (Percetage / 100 + 1)
                    End If
               
             If total_value > 0 Then
                     AccountCode = get_account_code_branch(85, my_branch)
                         If ModAccounts.AddNewDev(LngDevID, lineno, AccountCode, total_value, 1, Msg & " " & "«·Œœ„« ", val(notes_id), , , , XPDtbTrans.value, user_id, , , , , , , , , CURRENT_LINE, , , , , , , , , val(Dcbranch.BoundText), , , , , , , , , , 0, val(DcbUnitType.BoundText), val(DcbUnitNo.BoundText)) = False Then
                        GoTo ErrTrap
                    End If
                      lineno = lineno + 1
              End If
              
              ''''// «·”⁄Ì
            total_value = Round(.TextMatrix(i, .ColIndex("CommissionsPayed")), 2)
                If Percetage <> 0 Then
                    total_value = total_value / (Percetage / 100 + 1)
                    End If
             If total_value > 0 Then
                     AccountCode = get_account_code_branch(81, my_branch)
                         If ModAccounts.AddNewDev(LngDevID, lineno, AccountCode, total_value, 1, Msg & " " & "«·”⁄Ì", val(notes_id), , , , XPDtbTrans.value, user_id, , , , , , , , , CURRENT_LINE, , , , , , , , , val(Dcbranch.BoundText), , , , , , , , , , 0, val(DcbUnitType.BoundText), val(DcbUnitNo.BoundText)) = False Then
                        GoTo ErrTrap
                    End If
                      lineno = lineno + 1
              End If
                        ''///«·ﬁÌ„… «·„÷«›…
               total_value = Round(.TextMatrix(i, .ColIndex("RentValuePayed")), 2) + Round(.TextMatrix(i, .ColIndex("WaterPayed")), 2) + Round(.TextMatrix(i, .ColIndex("ElectricPayed")), 2) '+ Round(.TextMatrix(i, .ColIndex("TelandNetPayed")), 2)
              If ComResid(1).value = True Then
                 total_value = total_value * Percetage / 100
               Else
               total_value = 0
               End If
 
             If total_value > 0 Then
                    '  AccountCode = ModAccounts.GetMyAccountCode("TblCustemers", "CusID", val(Txtownerid), "Account_code1")
                    If SystemOptions.OpenAccountAqar = False Then
                        AccountCode = ModAccounts.GetMyAccountCode("TblCustemers", "CusID", val(Txtownerid), "Account_code")
                    Else
                        AccountCode = GetAqarAcountCode(val(DcbIqara.BoundText))
                    End If
                         If ModAccounts.AddNewDev(LngDevID, lineno, AccountCodeVat, total_value, 1, Msg & " " & "«·ﬁÌ„… «·„÷«›… ", val(notes_id), , , , XPDtbTrans.value, user_id, , , , , , , , , CURRENT_LINE, , , , , , , , , val(Dcbranch.BoundText), , , , , , , , , , 0, val(DcbUnitType.BoundText), val(DcbUnitNo.BoundText)) = False Then
                        GoTo ErrTrap
                    End If
                      lineno = lineno + 1
              End If
              
              
'////////////////////«·ﬁÌ„… «·„÷«›… ··”⁄Ì
    total_value = Round(.TextMatrix(i, .ColIndex("CommissionsPayed")), 2)
                If Percetage <> 0 Then
                    total_value = total_value / (Percetage / 100 + 1)
                    End If
                  commissionvalue = total_value * Percetage / 100
             If commissionvalue > 0 Then
                     commissionvalue = 0 'salimhere
                   If ModAccounts.AddNewDev(LngDevID, lineno, AccountCodeVat, commissionvalue, 1, Msg & " " & "«·ﬁÌ„… «·„÷«›… ··”⁄Ì", val(notes_id), , , , XPDtbTrans.value, user_id, , , , , , , , , CURRENT_LINE, , , , , , , , , val(Dcbranch.BoundText), , , , , , , , , , 0, val(DcbUnitType.BoundText), val(DcbUnitNo.BoundText)) = False Then
                        GoTo ErrTrap
                    End If
                      lineno = lineno + 1
              End If
                  total_value = Round(.TextMatrix(i, .ColIndex("TelandNetPayed")), 2)
                If Percetage <> 0 Then
                    total_value = total_value / (Percetage / 100 + 1)
                    End If
                  commissionvalue = total_value * Percetage / 100
              
                         If commissionvalue > 0 Then
                    commissionvalue = 0 'salimhere
                   If ModAccounts.AddNewDev(LngDevID, lineno, AccountCodeVat, commissionvalue, 1, Msg & " " & "«·ﬁÌ„… «·„÷«›… ··Œœ„« ", val(notes_id), , , , XPDtbTrans.value, user_id, , , , , , , , , CURRENT_LINE, , , , , , , , , val(Dcbranch.BoundText), , , , , , , , , , 0, val(DcbUnitType.BoundText), val(DcbUnitNo.BoundText)) = False Then
                        GoTo ErrTrap
                    End If
                      lineno = lineno + 1
              End If

              
              
              
              
              
       'salimHere*************************************
                                                          ''///«· «„Ì‰
                total_value = Round(.TextMatrix(i, .ColIndex("InsurancePayed")), 2)
             If total_value > 0 Then
                    If ModAccounts.AddNewDev(LngDevID, lineno, depit_side, total_value, 0, Msg & " " & "«· «„Ì‰", val(notes_id), , , , XPDtbTrans.value, user_id, , , , , , , , , CURRENT_LINE, , , , , , , , , val(Dcbranch.BoundText), , , , , , , , , , 0, val(DcbUnitType.BoundText), val(DcbUnitNo.BoundText)) = False Then
                        GoTo ErrTrap
                    End If
                    lineno = lineno + 1
          
          AccountCode = get_account_code_branch(82, my_branch)
          
                    ' AccountCode = ModAccounts.GetMyAccountCode("TblCustemers", "CusID", val(Txtownerid), "Account_code")
                         If ModAccounts.AddNewDev(LngDevID, lineno, AccountCode, total_value, 1, Msg & " " & "«· «„Ì‰", val(notes_id), , , , XPDtbTrans.value, user_id, , , , , , , , , CURRENT_LINE, , , , , , , , , val(Dcbranch.BoundText), , , , , , , , , , 0, val(DcbUnitType.BoundText), val(DcbUnitNo.BoundText)) = False Then
                        GoTo ErrTrap
                    End If
                      lineno = lineno + 1
              End If
              
                  '—’Ìœ „ »ﬁÌ ⁄·Ì «·„” √Ã—   OldValuePayed
        'salimhere
        If val(.TextMatrix(i, .ColIndex("OldValuePayed"))) > 0 Then
                 total_value = Round((.TextMatrix(i, .ColIndex("OldValuePayed"))), 2)
               Else
               total_value = 0
               End If
               
                If total_value > 0 Then
                 If ModAccounts.AddNewDev(LngDevID, lineno, depit_side, total_value, 0, Msg & " " & "—’Ìœ „ »ﬁÌ ⁄·Ì «·„” √Ã—", val(notes_id), , , , XPDtbTrans.value, user_id, , , , , , , , , CURRENT_LINE, , , , , , , , , val(Dcbranch.BoundText), , , , , , , , , , 0, val(DcbUnitType.BoundText), val(DcbUnitNo.BoundText)) = False Then
                        GoTo ErrTrap
                    End If
                    lineno = lineno + 1
                    
                
               
             AccountCode = get_account_code_branch(86, my_branch)
             
                  '  AccountCode = ModAccounts.GetMyAccountCode("TblCustemers", "CusID", val(Txtownerid), "Account_code")
                      If ModAccounts.AddNewDev(LngDevID, lineno, AccountCode, total_value, 1, Msg & " " & "—’Ìœ „ »ﬁÌ ⁄·Ì «·„” √Ã—", val(notes_id), , , , XPDtbTrans.value, user_id, , , , , , , , , CURRENT_LINE, , , , , , , , , val(Dcbranch.BoundText), , , , , , , , , , 0, val(DcbUnitType.BoundText), val(DcbUnitNo.BoundText)) = False Then
                        GoTo ErrTrap
                    End If
                      lineno = lineno + 1
              
         
                    
              End If
              
     

'salimhere******************************************************
                    
    End If
   
        

        Next i

    End With
     'salimhere  ﬁÌœ «·’‰œÊﬁ / «Ê «·»‰ﬂ
                 total_value = val(XPTxtVal.Text) + val(TxtVATValue)
             If total_value > 0 Then
                    If ModAccounts.AddNewDev(LngDevID, lineno, DcboDebitSide.BoundText, total_value, 0, Msg & " " & "«·’‰œÊﬁ/«·»‰ﬂ", val(notes_id), , , , XPDtbTrans.value, user_id, , , , , , , , , CURRENT_LINE, , , , , , , , , val(Dcbranch.BoundText), , , , , , , , , , 0, val(DcbUnitType.BoundText), val(DcbUnitNo.BoundText)) = False Then
                        GoTo ErrTrap
                    End If
                    lineno = lineno + 1
                     
                         If ModAccounts.AddNewDev(LngDevID, lineno, depit_side, total_value, 1, Msg & " " & "«·„” «Ã—", val(notes_id), , , , XPDtbTrans.value, user_id, , , , , , , , , CURRENT_LINE, , , , , , , , , val(Dcbranch.BoundText), , , , , , , , , , val(DcbIqara.BoundText), val(DcbUnitType.BoundText), val(DcbUnitNo.BoundText)) = False Then
                        GoTo ErrTrap
                    End If
                      lineno = lineno + 1
              End If

 
    DoEvents
    cProgress.FinishProgress
    cProgress.StopProgess
    Set cProgress = Nothing
    
ErrTrap:
End Function

 Function OtherOwnerNoreatJlInContractFiter(LngDevID As Long, notes_id As Double) As Double
fittervat = 0
If DCboCashType.ListIndex <> 10 Then Exit Function
Dim Percetage As Double
Dim commissionvalue As Double
Dim total_value As Double
Dim cProgress As ClsProgress
Set cProgress = New ClsProgress
    cProgress.ProgressType = Waiting
 Dim foxy_ked_NO As String
 Dim credit_side As String
 Dim My_SQL As String
 Dim Line1 As Double
 

         lineno = lineno + 1
 Dim AccountCode As String
    cProgress.StartProgress
    DoEvents
    LngDevID = new_id("DOUBLE_ENTREY_VOUCHERS", "Double_Entry_Vouchers_ID", "")
    Dim Msgdes As String
    Dim CURRENT_LINE As Double
    Dim depit_side As String
    Dim Msg As String
     Dim i As Integer
     Dim AccountCodeDept As String
Msgdes = "»‰«¡ ⁄·Ï „ﬁ»Ê÷«   ’›Ì… «„·«ﬂ «·€Ì— —ﬁ„ " & TxtNoteSerial1.Text & " "
TxtVATValue.Text = 0

Dim AccountCodeVat As String
Msg = XPMTxtRemarks.Text & CHR(13) & Msgdes
AccountCodeDept = ModAccounts.GetMyAccountCode("TblCustemers", "CusID", val(DBCboClientName.BoundText), "Account_code")
                  If SystemOptions.OpenAccountAqar = False Then
                        AccountCode = ModAccounts.GetMyAccountCode("TblCustemers", "CusID", val(Txtownerid), "Account_code")
                    Else
                        AccountCode = GetAqarAcountCode(val(DcbIqara.BoundText))
                    End If
'AccountCode = ModAccounts.GetMyAccountCode("TblCustemers", "CusID", val(Txtownerid), "Account_code")
                total_value = GetValueFiter(val(TxtFilterNo.Text), "RemainRent") 'val(XPTxtVal.Text)
                If ComResid(1).value = True Then
                     PercentgValueAddedAccount_Transec XPDtbTrans.value, 8, 1, AccountCodeVat, Percetage
                commissionvalue = total_value * Percetage / 100
              commissionvalue = Round(commissionvalue, 2)
               Else
               commissionvalue = 0
               End If
                'salimhere
               fittervat = val(fittervat) + commissionvalue
             If total_value > 0 Then
                    If ModAccounts.AddNewDev(LngDevID, lineno, AccountCodeDept, total_value + commissionvalue, 0, Msg & " " & "«·«ÌÃ«— ", val(notes_id), , , , XPDtbTrans.value, user_id, , , , , , , , , CURRENT_LINE, , , , , , , , , val(Dcbranch.BoundText), , , , , , , , , , val(DcbIqara.BoundText), val(DcbUnitType.BoundText), val(DcbUnitNo.BoundText)) = False Then
                        GoTo ErrTrap
                    End If
                    lineno = lineno + 1
                   
                   
                         If ModAccounts.AddNewDev(LngDevID, lineno, AccountCode, total_value, 1, Msg & " " & "  «·«ÌÃ«—", val(notes_id), , , , XPDtbTrans.value, user_id, , , , , , , , , CURRENT_LINE, , , , , , , , , val(Dcbranch.BoundText), , , , , , , , , , val(DcbIqara.BoundText), val(DcbUnitType.BoundText), val(DcbUnitNo.BoundText)) = False Then
                        GoTo ErrTrap
                    End If
                      lineno = lineno + 1
                    If commissionvalue > 0 Then
                    
                         If ModAccounts.AddNewDev(LngDevID, lineno, AccountCodeVat, commissionvalue, 1, Msg & " " & " «·ﬁÌ„… «·„÷«›…··«ÌÃ«— ", val(notes_id), , , , XPDtbTrans.value, user_id, , , , , , , , , CURRENT_LINE, , , , , , , , , val(Dcbranch.BoundText), , , , , , , , , , val(DcbIqara.BoundText), val(DcbUnitType.BoundText), val(DcbUnitNo.BoundText)) = False Then
                        GoTo ErrTrap
                    End If
                      lineno = lineno + 1
                    End If
              End If
           ''//////—’Ìœ ”«»ﬁ
                           total_value = GetValueFiter(val(TxtFilterNo.Text), "OldRent") 'val(XPTxtVal.Text)
                If ComResid(1).value = True Then
                     PercentgValueAddedAccount_Transec XPDtbTrans.value, 8, 1, AccountCodeVat, Percetage
                commissionvalue = total_value * Percetage / 100
              commissionvalue = Round(commissionvalue, 2)
               Else
               commissionvalue = 0
               End If
             'salimhere
               commissionvalue = 0
            'salimhere
            
             If total_value > 0 Then
                    If ModAccounts.AddNewDev(LngDevID, lineno, AccountCodeDept, total_value + commissionvalue, 0, Msg & " " & "—’Ìœ ”«»ﬁ ", val(notes_id), , , , XPDtbTrans.value, user_id, , , , , , , , , CURRENT_LINE, , , , , , , , , val(Dcbranch.BoundText), , , , , , , , , , val(DcbIqara.BoundText), val(DcbUnitType.BoundText), val(DcbUnitNo.BoundText)) = False Then
                        GoTo ErrTrap
                    End If
                    lineno = lineno + 1
                   
                   
                         If ModAccounts.AddNewDev(LngDevID, lineno, AccountCode, total_value, 1, Msg & " " & "  —’Ìœ ”«»ﬁ ", val(notes_id), , , , XPDtbTrans.value, user_id, , , , , , , , , CURRENT_LINE, , , , , , , , , val(Dcbranch.BoundText), , , , , , , , , , val(DcbIqara.BoundText), val(DcbUnitType.BoundText), val(DcbUnitNo.BoundText)) = False Then
                        GoTo ErrTrap
                    End If
                      lineno = lineno + 1
                    If commissionvalue > 0 Then
                    
                         If ModAccounts.AddNewDev(LngDevID, lineno, AccountCodeVat, commissionvalue, 1, Msg & " " & " «·ﬁÌ„… «·„÷«›…—’Ìœ ”«»ﬁ   ", val(notes_id), , , , XPDtbTrans.value, user_id, , , , , , , , , CURRENT_LINE, , , , , , , , , val(Dcbranch.BoundText), , , , , , , , , , val(DcbIqara.BoundText), val(DcbUnitType.BoundText), val(DcbUnitNo.BoundText)) = False Then
                        GoTo ErrTrap
                    End If
                      lineno = lineno + 1
                    End If
              End If
              ''//////////«ÌÃ«— «Ì«„ “Ì«œ…
                    total_value = GetValueFiterHeader(val(TxtFilterNo.Text), "DaysValueIncrease") 'val(XPTxtVal.Text)
                If ComResid(1).value = True Then
                     PercentgValueAddedAccount_Transec XPDtbTrans.value, 8, 1, AccountCodeVat, Percetage
                commissionvalue = total_value * Percetage / 100
              commissionvalue = Round(commissionvalue, 2)
               Else
               commissionvalue = 0
               End If
              'salimhere
              fittervat = val(fittervat) + commissionvalue
             If total_value > 0 Then
                    If ModAccounts.AddNewDev(LngDevID, lineno, AccountCodeDept, total_value + commissionvalue, 0, Msg & " " & "«·«ÌÃ«— «Ì«„ “Ì«œ…", val(notes_id), , , , XPDtbTrans.value, user_id, , , , , , , , , CURRENT_LINE, , , , , , , , , val(Dcbranch.BoundText), , , , , , , , , , val(DcbIqara.BoundText), val(DcbUnitType.BoundText), val(DcbUnitNo.BoundText)) = False Then
                        GoTo ErrTrap
                    End If
                    lineno = lineno + 1
                   
                   
                         If ModAccounts.AddNewDev(LngDevID, lineno, AccountCode, total_value, 1, Msg & " " & "  «·«ÌÃ«— «Ì«„ “Ì«œ…", val(notes_id), , , , XPDtbTrans.value, user_id, , , , , , , , , CURRENT_LINE, , , , , , , , , val(Dcbranch.BoundText), , , , , , , , , , val(DcbIqara.BoundText), val(DcbUnitType.BoundText), val(DcbUnitNo.BoundText)) = False Then
                        GoTo ErrTrap
                    End If
                      lineno = lineno + 1
                    If commissionvalue > 0 Then
                    
                         If ModAccounts.AddNewDev(LngDevID, lineno, AccountCodeVat, commissionvalue, 1, Msg & " " & " «·ﬁÌ„… «·„÷«›…··«ÌÃ«— «·«ÌÃ«— «Ì«„ “Ì«œ…", val(notes_id), , , , XPDtbTrans.value, user_id, , , , , , , , , CURRENT_LINE, , , , , , , , , val(Dcbranch.BoundText), , , , , , , , , , val(DcbIqara.BoundText), val(DcbUnitType.BoundText), val(DcbUnitNo.BoundText)) = False Then
                        GoTo ErrTrap
                    End If
                      lineno = lineno + 1
                    End If
              End If
              
                ''//////„Ì«Â
                          total_value = GetValueFiter(val(TxtFilterNo.Text), "RemainWater") 'val(XPTxtVal.Text)
               If total_value > 0 Then
                If ComResid(1).value = True Then
                     PercentgValueAddedAccount_Transec XPDtbTrans.value, 8, 1, AccountCodeVat, Percetage
                     commissionvalue = total_value * Percetage / 100
                     commissionvalue = Round(commissionvalue, 2)
                 Else
                     commissionvalue = 0
                End If
             fittervat = val(fittervat) + commissionvalue
                    If ModAccounts.AddNewDev(LngDevID, lineno, AccountCodeDept, total_value + commissionvalue, 0, Msg & " " & "„Ì«Â ", val(notes_id), , , , XPDtbTrans.value, user_id, , , , , , , , , CURRENT_LINE, , , , , , , , , val(Dcbranch.BoundText), , , , , , , , , , val(DcbIqara.BoundText), val(DcbUnitType.BoundText), val(DcbUnitNo.BoundText)) = False Then
                        GoTo ErrTrap
                    End If
                    lineno = lineno + 1
                   
                    
                         If ModAccounts.AddNewDev(LngDevID, lineno, AccountCode, total_value, 1, Msg & " " & " „Ì«Â ", val(notes_id), , , , XPDtbTrans.value, user_id, , , , , , , , , CURRENT_LINE, , , , , , , , , val(Dcbranch.BoundText), , , , , , , , , , val(DcbIqara.BoundText), val(DcbUnitType.BoundText), val(DcbUnitNo.BoundText)) = False Then
                        GoTo ErrTrap
                    End If
                      lineno = lineno + 1
                    If commissionvalue > 0 Then
                    
                         If ModAccounts.AddNewDev(LngDevID, lineno, AccountCodeVat, commissionvalue, 1, Msg & " " & " «·ﬁÌ„… «·„÷«›…··„Ì«Â ", val(notes_id), , , , XPDtbTrans.value, user_id, , , , , , , , , CURRENT_LINE, , , , , , , , , val(Dcbranch.BoundText), , , , , , , , , , val(DcbIqara.BoundText), val(DcbUnitType.BoundText), val(DcbUnitNo.BoundText)) = False Then
                        GoTo ErrTrap
                    End If
                      lineno = lineno + 1
                    End If
              End If
              ''//////ﬂÂ—»«¡
              
                     total_value = GetValueFiter(val(TxtFilterNo.Text), "BillPrice") 'val(XPTxtVal.Text)
                  If total_value > 0 Then
                     PercentgValueAddedAccount_Transec XPDtbTrans.value, 8, 1, AccountCodeVat, Percetage
                commissionvalue = total_value * Percetage / 100
                  commissionvalue = Round(commissionvalue, 2)
     
        fittervat = val(fittervat) + commissionvalue
        
                    If ModAccounts.AddNewDev(LngDevID, lineno, AccountCodeDept, total_value + commissionvalue, 0, Msg & " " & "ﬂÂ—»«¡ ", val(notes_id), , , , XPDtbTrans.value, user_id, , , , , , , , , CURRENT_LINE, , , , , , , , , val(Dcbranch.BoundText), , , , , , , , , , val(DcbIqara.BoundText), val(DcbUnitType.BoundText), val(DcbUnitNo.BoundText)) = False Then
                        GoTo ErrTrap
                    End If
                    lineno = lineno + 1
                   
                    
                         If ModAccounts.AddNewDev(LngDevID, lineno, AccountCode, total_value, 1, Msg & " " & "  ﬂÂ—»«¡", val(notes_id), , , , XPDtbTrans.value, user_id, , , , , , , , , CURRENT_LINE, , , , , , , , , val(Dcbranch.BoundText), , , , , , , , , , val(DcbIqara.BoundText), val(DcbUnitType.BoundText), val(DcbUnitNo.BoundText)) = False Then
                        GoTo ErrTrap
                    End If
                      lineno = lineno + 1
                      'salimhere
               If commissionvalue > 0 Then
                    
                          If ModAccounts.AddNewDev(LngDevID, lineno, AccountCodeVat, commissionvalue, 1, Msg & " " & " «·ﬁÌ„… «·„÷«›… ··ﬂÂ—»«¡ ", val(notes_id), , , , XPDtbTrans.value, user_id, , , , , , , , , CURRENT_LINE, , , , , , , , , val(Dcbranch.BoundText), , , , , , , , , , val(DcbIqara.BoundText), val(DcbUnitType.BoundText), val(DcbUnitNo.BoundText)) = False Then
                      GoTo ErrTrap
                    End If
                     lineno = lineno + 1
                    End If
              
              
              End If
                         ''//////Œœ„« 
              
                     total_value = GetValueFiter(val(TxtFilterNo.Text), "RemainService") 'val(XPTxtVal.Text)
                  If total_value > 0 Then
                     PercentgValueAddedAccount_Transec XPDtbTrans.value, 8, 1, AccountCodeVat, Percetage
     
            
                    If ModAccounts.AddNewDev(LngDevID, lineno, AccountCodeDept, total_value, 0, Msg & " " & " Œœ„« ", val(notes_id), , , , XPDtbTrans.value, user_id, , , , , , , , , CURRENT_LINE, , , , , , , , , val(Dcbranch.BoundText), , , , , , , , , , val(DcbIqara.BoundText), val(DcbUnitType.BoundText), val(DcbUnitNo.BoundText)) = False Then
                        GoTo ErrTrap
                    End If
                    lineno = lineno + 1
                    'salimhere
                   If Percetage > 0 Then 'salimhere
                    total_value = total_value / (Percetage / 100 + 1)
                    commissionvalue = total_value * Percetage / 100
                    Else
                    commissionvalue = 0
                    End If
                 
                         If ModAccounts.AddNewDev(LngDevID, lineno, AccountCode, total_value, 1, Msg & " " & "  Œœ„« ", val(notes_id), , , , XPDtbTrans.value, user_id, , , , , , , , , CURRENT_LINE, , , , , , , , , val(Dcbranch.BoundText), , , , , , , , , , val(DcbIqara.BoundText), val(DcbUnitType.BoundText), val(DcbUnitNo.BoundText)) = False Then
                        GoTo ErrTrap
                    End If
                      lineno = lineno + 1
                    If commissionvalue > 0 Then
                    
                         If ModAccounts.AddNewDev(LngDevID, lineno, AccountCodeVat, commissionvalue, 1, Msg & " " & " «·ﬁÌ„… «·„÷«›…«·Œœ„«  ", val(notes_id), , , , XPDtbTrans.value, user_id, , , , , , , , , CURRENT_LINE, , , , , , , , , val(Dcbranch.BoundText), , , , , , , , , , val(DcbIqara.BoundText), val(DcbUnitType.BoundText), val(DcbUnitNo.BoundText)) = False Then
                        GoTo ErrTrap
                    End If
                      lineno = lineno + 1
                    End If
              End If
              
              
                                     ''//////”⁄Ì
              
                     total_value = GetValueFiter(val(TxtFilterNo.Text), "RemainCommissions") 'val(XPTxtVal.Text)
                  If total_value > 0 Then
                     PercentgValueAddedAccount_Transec XPDtbTrans.value, 8, 1, AccountCodeVat, Percetage
            
                    If ModAccounts.AddNewDev(LngDevID, lineno, AccountCodeDept, total_value, 0, Msg & " " & " ”⁄Ì", val(notes_id), , , , XPDtbTrans.value, user_id, , , , , , , , , CURRENT_LINE, , , , , , , , , val(Dcbranch.BoundText), , , , , , , , , , val(DcbIqara.BoundText), val(DcbUnitType.BoundText), val(DcbUnitNo.BoundText)) = False Then
                        GoTo ErrTrap
                    End If
                    lineno = lineno + 1
                    'salimhere
                   If Percetage > 0 Then
                    total_value = total_value / (Percetage / 100 + 1)
                    commissionvalue = total_value * Percetage / 100
                    Else
                    commissionvalue = 0
                    End If
                     
                         If ModAccounts.AddNewDev(LngDevID, lineno, AccountCode, total_value, 1, Msg & " " & "  ”⁄Ì", val(notes_id), , , , XPDtbTrans.value, user_id, , , , , , , , , CURRENT_LINE, , , , , , , , , val(Dcbranch.BoundText), , , , , , , , , , val(DcbIqara.BoundText), val(DcbUnitType.BoundText), val(DcbUnitNo.BoundText)) = False Then
                        GoTo ErrTrap
                    End If
                      lineno = lineno + 1
                    If commissionvalue > 0 Then
                    
                         If ModAccounts.AddNewDev(LngDevID, lineno, AccountCodeVat, commissionvalue, 1, Msg & " " & " «·ﬁÌ„… «·„÷«›…··”⁄Ì ", val(notes_id), , , , XPDtbTrans.value, user_id, , , , , , , , , CURRENT_LINE, , , , , , , , , val(Dcbranch.BoundText), , , , , , , , , , val(DcbIqara.BoundText), val(DcbUnitType.BoundText), val(DcbUnitNo.BoundText)) = False Then
                        GoTo ErrTrap
                    End If
                      lineno = lineno + 1
                    End If
              End If
              
   
                            
                         ''''// «· «„Ì‰ «·”«»ﬁ
              
            total_value = val(txtTotalinsuranceS.Text) - GetValueFiter(val(TxtFilterNo.Text), "insurance")
            total_value = Abs(total_value)
               
          ''''// «· «„Ì‰ «·”«»ﬁ  Salim here
     '     total_value = 0
    If total_value > 0 Then
            'salimhere
        '            If ModAccounts.AddNewDev(LngDevID, lineno, AccountCodeDept, total_value, 0, Msg & " " & " «„Ì‰ ”«»ﬁ", val(notes_id), , , , XPDtbTrans.value, user_id, , , , , , , , , CURRENT_LINE, , , , , , , , , val(Dcbranch.BoundText), , , , , , , , , , val(DcbIqara.BoundText), val(DcbUnitType.BoundText), val(DcbUnitNo.BoundText)) = False Then
        '                GoTo ErrTrap
        '            End If
        '            lineno = lineno + 1
        '
        '                 If ModAccounts.AddNewDev(LngDevID, lineno, AccountCode, total_value, 1, Msg & " " & " «„Ì‰ ”«»ﬁ", val(notes_id), , , , XPDtbTrans.value, user_id, , , , , , , , , CURRENT_LINE, , , , , , , , , val(Dcbranch.BoundText), , , , , , , , , , val(DcbIqara.BoundText), val(DcbUnitType.BoundText), val(DcbUnitNo.BoundText)) = False Then
        '                GoTo ErrTrap
        '            End If
        '              lineno = lineno + 1
                      
                             If ModAccounts.AddNewDev(LngDevID, lineno, AccountCode, total_value, 0, Msg & " " & " «„Ì‰ ”«»ﬁ", val(notes_id), , , , XPDtbTrans.value, user_id, , , , , , , , , CURRENT_LINE, , , , , , , , , val(Dcbranch.BoundText), , , , , , , , , , val(DcbIqara.BoundText), val(DcbUnitType.BoundText), val(DcbUnitNo.BoundText)) = False Then
                        GoTo ErrTrap
                    End If
                    lineno = lineno + 1
                 
                         If ModAccounts.AddNewDev(LngDevID, lineno, AccountCodeDept, total_value, 1, Msg & " " & " «„Ì‰”«»ﬁ", val(notes_id), , , , XPDtbTrans.value, user_id, , , , , , , , , CURRENT_LINE, , , , , , , , , val(Dcbranch.BoundText), , , , , , , , , , val(DcbIqara.BoundText), val(DcbUnitType.BoundText), val(DcbUnitNo.BoundText)) = False Then
                        GoTo ErrTrap
                    End If
                      lineno = lineno + 1
                      
              End If
                                               
                                       ''''// «· «„Ì‰
              
            total_value = GetValueFiter(val(TxtFilterNo.Text), "insurance")
            total_value = Abs(total_value)
            'total_value = 0
             If total_value > 0 Then
                'salimhere
                    'If ModAccounts.AddNewDev(LngDevID, lineno, AccountCodeDept, total_value, 0, Msg & " " & " «„Ì‰", val(notes_id), , , , XPDtbTrans.value, user_id, , , , , , , , , CURRENT_LINE, , , , , , , , , val(Dcbranch.BoundText), , , , , , , , , , val(DcbIqara.BoundText), val(DcbUnitType.BoundText), val(DcbUnitNo.BoundText)) = False Then
                    '    GoTo ErrTrap
                    'End If
                    'lineno = lineno + 1
                    
                    '     If ModAccounts.AddNewDev(LngDevID, lineno, AccountCode, total_value, 1, Msg & " " & " «„Ì‰", val(notes_id), , , , XPDtbTrans.value, user_id, , , , , , , , , CURRENT_LINE, , , , , , , , , val(Dcbranch.BoundText), , , , , , , , , , val(DcbIqara.BoundText), val(DcbUnitType.BoundText), val(DcbUnitNo.BoundText)) = False Then
                    '    GoTo ErrTrap
                    'End If
                    '  lineno = lineno + 1
                      
                             If ModAccounts.AddNewDev(LngDevID, lineno, AccountCode, total_value, 0, Msg & " " & " «„Ì‰", val(notes_id), , , , XPDtbTrans.value, user_id, , , , , , , , , CURRENT_LINE, , , , , , , , , val(Dcbranch.BoundText), , , , , , , , , , val(DcbIqara.BoundText), val(DcbUnitType.BoundText), val(DcbUnitNo.BoundText)) = False Then
                        GoTo ErrTrap
                    End If
                    lineno = lineno + 1
                 
                         If ModAccounts.AddNewDev(LngDevID, lineno, AccountCodeDept, total_value, 1, Msg & " " & " «„Ì‰", val(notes_id), , , , XPDtbTrans.value, user_id, , , , , , , , , CURRENT_LINE, , , , , , , , , val(Dcbranch.BoundText), , , , , , , , , , val(DcbIqara.BoundText), val(DcbUnitType.BoundText), val(DcbUnitNo.BoundText)) = False Then
                        GoTo ErrTrap
                    End If
                      lineno = lineno + 1
                      
              End If
              
               
             'salimhere
                total_value = val(XPTxtVal.Text) + val(fittervat)
             If total_value > 0 Then
                    If ModAccounts.AddNewDev(LngDevID, lineno, DcboDebitSide.BoundText, total_value, 0, Msg & " " & "«·’‰œÊﬁ/«·»‰ﬂ", val(notes_id), , , , XPDtbTrans.value, user_id, , , , , , , , , CURRENT_LINE, , , , , , , , , val(Dcbranch.BoundText), , , , , , , , , , val(DcbIqara.BoundText), val(DcbUnitType.BoundText), val(DcbUnitNo.BoundText)) = False Then
                        GoTo ErrTrap
                    End If
                    lineno = lineno + 1
                     
                         If ModAccounts.AddNewDev(LngDevID, lineno, AccountCodeDept, total_value, 1, Msg & " " & "«·„” «Ã—", val(notes_id), , , , XPDtbTrans.value, user_id, , , , , , , , , CURRENT_LINE, , , , , , , , , val(Dcbranch.BoundText), , , , , , , , , , val(DcbIqara.BoundText), val(DcbUnitType.BoundText), val(DcbUnitNo.BoundText)) = False Then
                        GoTo ErrTrap
                    End If
                      lineno = lineno + 1
              End If
              

        
    DoEvents
    cProgress.FinishProgress
    cProgress.StopProgess
    Set cProgress = Nothing
    
ErrTrap:
End Function
 Function MyOwnerNoreatJlInContractFiter(LngDevID As Long, notes_id As Double) As Double
fittervat = 0
If DCboCashType.ListIndex <> 10 Then Exit Function
Dim Percetage As Double
Dim commissionvalue As Double
Dim total_value As Double
Dim cProgress As ClsProgress
Set cProgress = New ClsProgress
    cProgress.ProgressType = Waiting
 Dim foxy_ked_NO As String
 Dim credit_side As String
 Dim My_SQL As String
 Dim Line1 As Double

 Dim AccountCodeDept As String
         lineno = lineno + 1
 Dim AccountCode As String
    cProgress.StartProgress
    DoEvents
    LngDevID = new_id("DOUBLE_ENTREY_VOUCHERS", "Double_Entry_Vouchers_ID", "")
    Dim Msgdes As String
    Dim CURRENT_LINE As Double
    Dim depit_side As String
    Dim Msg As String
     Dim i As Integer
     
     Dim AccountCodeVat As String
Msgdes = "»‰«¡ ⁄·Ï „ﬁ»Ê÷«   ’›Ì… «„·«ﬂÌ  —ﬁ„ " & TxtNoteSerial1.Text & " "
PercentgValueAddedAccount_Transec XPDtbTrans.value, 8, 1, AccountCodeVat, Percetage

Msg = XPMTxtRemarks.Text & CHR(13) & Msgdes
 AccountCodeDept = ModAccounts.GetMyAccountCode("TblCustemers", "CusID", val(DBCboClientName.BoundText), "Account_code")
                total_value = GetValueFiter(val(TxtFilterNo.Text), "RemainRent")
                If total_value > 0 Then
              If ComResid(1).value = True Then
                     commissionvalue = total_value * Percetage / 100
                     commissionvalue = Round(commissionvalue, 2)
              Else
              commissionvalue = 0
              End If
            fittervat = val(fittervat) + commissionvalue
                    If ModAccounts.AddNewDev(LngDevID, lineno, AccountCodeDept, total_value + commissionvalue, 0, Msg & " " & "«·„” «Ã— ", val(notes_id), , , , XPDtbTrans.value, user_id, , , , , , , , , CURRENT_LINE, , , , , , , , , val(Dcbranch.BoundText), , , , , , , , , , 0, val(DcbUnitType.BoundText), val(DcbUnitNo.BoundText)) = False Then
                        GoTo ErrTrap
                    End If
                    lineno = lineno + 1
                   
                    AccountCode = get_account_code_branch(86, my_branch)
                         If ModAccounts.AddNewDev(LngDevID, lineno, AccountCode, total_value, 1, Msg & " " & " «Ì—«œ«  «·«ÌÃ«—", val(notes_id), , , , XPDtbTrans.value, user_id, , , , , , , , , CURRENT_LINE, , , , , , , , , val(Dcbranch.BoundText), , , , , , , , , , 0, val(DcbUnitType.BoundText), val(DcbUnitNo.BoundText)) = False Then
                        GoTo ErrTrap
                    End If
                      lineno = lineno + 1
                    If commissionvalue > 0 Then
                    
                         If ModAccounts.AddNewDev(LngDevID, lineno, AccountCodeVat, commissionvalue, 1, Msg & " " & " «·ﬁÌ„… «·„÷«›… ", val(notes_id), , , , XPDtbTrans.value, user_id, , , , , , , , , CURRENT_LINE, , , , , , , , , val(Dcbranch.BoundText), , , , , , , , , , 0, val(DcbUnitType.BoundText), val(DcbUnitNo.BoundText)) = False Then
                        GoTo ErrTrap
                    End If
                      lineno = lineno + 1
                    End If
              End If
           '////—’Ìœ ”«»ﬁ
                          total_value = GetValueFiter(val(TxtFilterNo.Text), "OldRent")
                If total_value > 0 Then
              If ComResid(1).value = True Then
                     commissionvalue = total_value * Percetage / 100
                     commissionvalue = Round(commissionvalue, 2)
              Else
              commissionvalue = 0
              End If
            'salim here
            commissionvalue = 0
            
                    If ModAccounts.AddNewDev(LngDevID, lineno, AccountCodeDept, total_value + commissionvalue, 0, Msg & " " & "—’Ìœ ”«»ﬁ ", val(notes_id), , , , XPDtbTrans.value, user_id, , , , , , , , , CURRENT_LINE, , , , , , , , , val(Dcbranch.BoundText), , , , , , , , , , 0, val(DcbUnitType.BoundText), val(DcbUnitNo.BoundText)) = False Then
                        GoTo ErrTrap
                    End If
                    lineno = lineno + 1
                   
                    AccountCode = get_account_code_branch(86, my_branch)
                         If ModAccounts.AddNewDev(LngDevID, lineno, AccountCode, total_value, 1, Msg & " " & "  —’Ìœ ”«»ﬁ", val(notes_id), , , , XPDtbTrans.value, user_id, , , , , , , , , CURRENT_LINE, , , , , , , , , val(Dcbranch.BoundText), , , , , , , , , , 0, val(DcbUnitType.BoundText), val(DcbUnitNo.BoundText)) = False Then
                        GoTo ErrTrap
                    End If
                      lineno = lineno + 1
                    If commissionvalue > 0 Then
                    
                         If ModAccounts.AddNewDev(LngDevID, lineno, AccountCodeVat, commissionvalue, 1, Msg & " " & " «·ﬁÌ„… «·„÷«›… —’Ìœ ”«»ﬁ", val(notes_id), , , , XPDtbTrans.value, user_id, , , , , , , , , CURRENT_LINE, , , , , , , , , val(Dcbranch.BoundText), , , , , , , , , , 0, val(DcbUnitType.BoundText), val(DcbUnitNo.BoundText)) = False Then
                        GoTo ErrTrap
                    End If
                      lineno = lineno + 1
                    End If
              End If
              '////«·«ÌÃ«— «Ì«„ “Ì«œ…
                              total_value = GetValueFiterHeader(val(TxtFilterNo.Text), "DaysValueIncrease")
                If total_value > 0 Then
              If ComResid(1).value = True Then
                     commissionvalue = total_value * Percetage / 100
                     commissionvalue = Round(commissionvalue, 2)
              Else
              commissionvalue = 0
              End If
            
            fittervat = val(fittervat) + commissionvalue
                    If ModAccounts.AddNewDev(LngDevID, lineno, AccountCodeDept, total_value + commissionvalue, 0, Msg & " " & "«ÌÃ«— «Ì«„ “Ì«œ… ", val(notes_id), , , , XPDtbTrans.value, user_id, , , , , , , , , CURRENT_LINE, , , , , , , , , val(Dcbranch.BoundText), , , , , , , , , , 0, val(DcbUnitType.BoundText), val(DcbUnitNo.BoundText)) = False Then
                        GoTo ErrTrap
                    End If
                    lineno = lineno + 1
                   
                    AccountCode = get_account_code_branch(86, my_branch)
                         If ModAccounts.AddNewDev(LngDevID, lineno, AccountCode, total_value, 1, Msg & " " & "  «Ã«— «Ì„ “Ì«œ…", val(notes_id), , , , XPDtbTrans.value, user_id, , , , , , , , , CURRENT_LINE, , , , , , , , , val(Dcbranch.BoundText), , , , , , , , , , 0, val(DcbUnitType.BoundText), val(DcbUnitNo.BoundText)) = False Then
                        GoTo ErrTrap
                    End If
                      lineno = lineno + 1
                    If commissionvalue > 0 Then
                    
                         If ModAccounts.AddNewDev(LngDevID, lineno, AccountCodeVat, commissionvalue, 1, Msg & " " & " «·ﬁÌ„… «·„÷«›… «ÌÃ«— «Ì«„ “Ì«œ…", val(notes_id), , , , XPDtbTrans.value, user_id, , , , , , , , , CURRENT_LINE, , , , , , , , , val(Dcbranch.BoundText), , , , , , , , , , 0, val(DcbUnitType.BoundText), val(DcbUnitNo.BoundText)) = False Then
                        GoTo ErrTrap
                    End If
                      lineno = lineno + 1
                    End If
              End If
              
              ''////„Ì«Â
                              total_value = GetValueFiter(val(TxtFilterNo.Text), "RemainWater")
                If total_value > 0 Then

             If ComResid(1).value = True Then
                     
                     commissionvalue = total_value * Percetage / 100
                     commissionvalue = Round(commissionvalue, 2)
              Else
              commissionvalue = 0
              End If
              fittervat = val(fittervat) + commissionvalue
                    If ModAccounts.AddNewDev(LngDevID, lineno, AccountCodeDept, total_value + commissionvalue, 0, Msg & " " & "«·„” «Ã— ", val(notes_id), , , , XPDtbTrans.value, user_id, , , , , , , , , CURRENT_LINE, , , , , , , , , val(Dcbranch.BoundText), , , , , , , , , , 0, val(DcbUnitType.BoundText), val(DcbUnitNo.BoundText)) = False Then
                        GoTo ErrTrap
                    End If
                    lineno = lineno + 1

                    AccountCode = get_account_code_branch(83, my_branch)
                         If ModAccounts.AddNewDev(LngDevID, lineno, AccountCode, total_value, 1, Msg & " " & " «Ì—«œ«  «·„Ì«Â", val(notes_id), , , , XPDtbTrans.value, user_id, , , , , , , , , CURRENT_LINE, , , , , , , , , val(Dcbranch.BoundText), , , , , , , , , , 0, val(DcbUnitType.BoundText), val(DcbUnitNo.BoundText)) = False Then
                        GoTo ErrTrap
                    End If
                      lineno = lineno + 1
                    If commissionvalue > 0 Then
                    
                         If ModAccounts.AddNewDev(LngDevID, lineno, AccountCodeVat, commissionvalue, 1, Msg & " " & " «·ﬁÌ„… «·„÷«›… ··„Ì«Â ", val(notes_id), , , , XPDtbTrans.value, user_id, , , , , , , , , CURRENT_LINE, , , , , , , , , val(Dcbranch.BoundText), , , , , , , , , , 0, val(DcbUnitType.BoundText), val(DcbUnitNo.BoundText)) = False Then
                        GoTo ErrTrap
                    End If
                      lineno = lineno + 1
                    End If
              End If
              
              ''////«·ﬂÂ—»«¡
                  total_value = GetValueFiter(val(TxtFilterNo.Text), "BillPrice")
                If total_value > 0 Then
                     commissionvalue = total_value * Percetage / 100
                     commissionvalue = Round(commissionvalue, 2)
fittervat = val(fittervat) + commissionvalue
                    If ModAccounts.AddNewDev(LngDevID, lineno, AccountCodeDept, total_value + commissionvalue, 0, Msg & " " & "«·„” «Ã— ", val(notes_id), , , , XPDtbTrans.value, user_id, , , , , , , , , CURRENT_LINE, , , , , , , , , val(Dcbranch.BoundText), , , , , , , , , , 0, val(DcbUnitType.BoundText), val(DcbUnitNo.BoundText)) = False Then
                        GoTo ErrTrap
                    End If
                    lineno = lineno + 1
                   
                    AccountCode = get_account_code_branch(84, my_branch)
                         If ModAccounts.AddNewDev(LngDevID, lineno, AccountCode, total_value, 1, Msg & " " & " «Ì—«œ«  «·ﬂÂ—»«¡", val(notes_id), , , , XPDtbTrans.value, user_id, , , , , , , , , CURRENT_LINE, , , , , , , , , val(Dcbranch.BoundText), , , , , , , , , , 0, val(DcbUnitType.BoundText), val(DcbUnitNo.BoundText)) = False Then
                        GoTo ErrTrap
                    End If
                      lineno = lineno + 1
                    If commissionvalue > 0 Then
                    
                         If ModAccounts.AddNewDev(LngDevID, lineno, AccountCodeVat, commissionvalue, 1, Msg & " " & " «·ﬁÌ„… «·„÷«›… ··ﬂÂ—»«¡ ", val(notes_id), , , , XPDtbTrans.value, user_id, , , , , , , , , CURRENT_LINE, , , , , , , , , val(Dcbranch.BoundText), , , , , , , , , , 0, val(DcbUnitType.BoundText), val(DcbUnitNo.BoundText)) = False Then
                        GoTo ErrTrap
                    End If
                      lineno = lineno + 1
                    End If
              End If
              
              
                            ''////Œœ„« 
                              total_value = GetValueFiter(val(TxtFilterNo.Text), "RemainService")
                If total_value > 0 Then

            
                    If ModAccounts.AddNewDev(LngDevID, lineno, AccountCodeDept, total_value, 0, Msg & " " & "«·„” «Ã— ", val(notes_id), , , , XPDtbTrans.value, user_id, , , , , , , , , CURRENT_LINE, , , , , , , , , val(Dcbranch.BoundText), , , , , , , , , , 0, val(DcbUnitType.BoundText), val(DcbUnitNo.BoundText)) = False Then
                        GoTo ErrTrap
                    End If
                    lineno = lineno + 1
            If True = True Then 'salimhere
                     total_value = total_value / (Percetage / 100 + 1)
                     commissionvalue = total_value * Percetage / 100
                     commissionvalue = Round(commissionvalue, 2)
              Else
              commissionvalue = 0
              End If
              
                    AccountCode = get_account_code_branch(85, my_branch)
                         If ModAccounts.AddNewDev(LngDevID, lineno, AccountCode, total_value, 1, Msg & " " & " «Ì—«œ«  Œœ„« ", val(notes_id), , , , XPDtbTrans.value, user_id, , , , , , , , , CURRENT_LINE, , , , , , , , , val(Dcbranch.BoundText), , , , , , , , , , 0, val(DcbUnitType.BoundText), val(DcbUnitNo.BoundText)) = False Then
                        GoTo ErrTrap
                    End If
                      lineno = lineno + 1
                    If commissionvalue > 0 Then
                    
                         If ModAccounts.AddNewDev(LngDevID, lineno, AccountCodeVat, commissionvalue, 1, Msg & " " & " «·ﬁÌ„… «·„÷«›… Œœ„«  ", val(notes_id), , , , XPDtbTrans.value, user_id, , , , , , , , , CURRENT_LINE, , , , , , , , , val(Dcbranch.BoundText), , , , , , , , , , 0, val(DcbUnitType.BoundText), val(DcbUnitNo.BoundText)) = False Then
                        GoTo ErrTrap
                    End If
                      lineno = lineno + 1
                    End If
              End If
                                          ''////”⁄Ì
                              total_value = GetValueFiter(val(TxtFilterNo.Text), "RemainCommissions")
                If total_value > 0 Then
            
                    If ModAccounts.AddNewDev(LngDevID, lineno, AccountCodeDept, total_value, 0, Msg & " " & "«·„” «Ã— ", val(notes_id), , , , XPDtbTrans.value, user_id, , , , , , , , , CURRENT_LINE, , , , , , , , , val(Dcbranch.BoundText), , , , , , , , , , 0, val(DcbUnitType.BoundText), val(DcbUnitNo.BoundText)) = False Then
                        GoTo ErrTrap
                    End If
                    lineno = lineno + 1
           If True = True Then 'salimhere
                     total_value = total_value / (Percetage / 100 + 1)
                     commissionvalue = total_value * Percetage / 100
                     commissionvalue = Round(commissionvalue, 2)
              Else
              commissionvalue = 0
              End If
                    AccountCode = get_account_code_branch(81, my_branch)
                         If ModAccounts.AddNewDev(LngDevID, lineno, AccountCode, total_value, 1, Msg & " " & " «Ì—«œ«  «·”⁄Ì", val(notes_id), , , , XPDtbTrans.value, user_id, , , , , , , , , CURRENT_LINE, , , , , , , , , val(Dcbranch.BoundText), , , , , , , , , , 0, val(DcbUnitType.BoundText), val(DcbUnitNo.BoundText)) = False Then
                        GoTo ErrTrap
                    End If
                      lineno = lineno + 1
                    If commissionvalue > 0 Then
                    
                         If ModAccounts.AddNewDev(LngDevID, lineno, AccountCodeVat, commissionvalue, 1, Msg & " " & " «·ﬁÌ„… «·„÷«›… ··”⁄Ì ", val(notes_id), , , , XPDtbTrans.value, user_id, , , , , , , , , CURRENT_LINE, , , , , , , , , val(Dcbranch.BoundText), , , , , , , , , , 0, val(DcbUnitType.BoundText), val(DcbUnitNo.BoundText)) = False Then
                        GoTo ErrTrap
                    End If
                      lineno = lineno + 1
                    End If
              End If
              ''''// «· «„Ì‰
              
            total_value = val(txtTotalinsuranceS.Text) - GetValueFiter(val(TxtFilterNo.Text), "insurance")
            total_value = Abs(total_value)
             If total_value > 0 Then
             AccountCode = get_account_code_branch(82, my_branch)
             
                        If ModAccounts.AddNewDev(LngDevID, lineno, AccountCode, total_value, 0, Msg & " " & " «„Ì‰ „” —œ...”«»ﬁ ", val(notes_id), , , , XPDtbTrans.value, user_id, , , , , , , , , CURRENT_LINE, , , , , , , , , val(Dcbranch.BoundText), , , , , , , , , , 0, val(DcbUnitType.BoundText), val(DcbUnitNo.BoundText)) = False Then
                        GoTo ErrTrap
                    End If
                      lineno = lineno + 1
                      
                     
                    If ModAccounts.AddNewDev(LngDevID, lineno, AccountCodeDept, total_value, 1, Msg & " " & " «„Ì‰ „” —œ.”«»ﬁ", val(notes_id), , , , XPDtbTrans.value, user_id, , , , , , , , , CURRENT_LINE, , , , , , , , , val(Dcbranch.BoundText), , , , , , , , , , 0, val(DcbUnitType.BoundText), val(DcbUnitNo.BoundText)) = False Then
                        GoTo ErrTrap
                    End If
                    lineno = lineno + 1
                     
              
              End If
                            
                        ''''// «· «„Ì‰
              
            total_value = GetValueFiter(val(TxtFilterNo.Text), "insurance")
             If total_value > 0 Then
             AccountCode = get_account_code_branch(82, my_branch)
                              If ModAccounts.AddNewDev(LngDevID, lineno, AccountCode, total_value, 0, Msg & " " & " «„Ì‰ „” —œ", val(notes_id), , , , XPDtbTrans.value, user_id, , , , , , , , , CURRENT_LINE, , , , , , , , , val(Dcbranch.BoundText), , , , , , , , , , 0, val(DcbUnitType.BoundText), val(DcbUnitNo.BoundText)) = False Then
                        GoTo ErrTrap
                    End If
                      lineno = lineno + 1
                      
                    
                    If ModAccounts.AddNewDev(LngDevID, lineno, AccountCodeDept, total_value, 1, Msg & " " & " «„Ì‰ „” —œ", val(notes_id), , , , XPDtbTrans.value, user_id, , , , , , , , , CURRENT_LINE, , , , , , , , , val(Dcbranch.BoundText), , , , , , , , , , 0, val(DcbUnitType.BoundText), val(DcbUnitNo.BoundText)) = False Then
                        GoTo ErrTrap
                    End If
                    lineno = lineno + 1
                     
        
                      
           '         If ModAccounts.AddNewDev(LngDevID, lineno, AccountCode, total_value, 0, Msg & " " & " «„Ì‰ „” —œ", val(notes_id), , , , XPDtbTrans.value, user_id, , , , , , , , , CURRENT_LINE, , , , , , , , , val(Dcbranch.BoundText), , , , , , , , , , 0, val(DcbUnitType.BoundText), val(DcbUnitNo.BoundText)) = False Then
           '             GoTo ErrTrap
           '         End If
           '         lineno = lineno + 1
           '
           '              If ModAccounts.AddNewDev(LngDevID, lineno, AccountCodeDept, total_value, 1, Msg & " " & " «„Ì‰ „” —œ", val(notes_id), , , , XPDtbTrans.value, user_id, , , , , , , , , CURRENT_LINE, , , , , , , , , val(Dcbranch.BoundText), , , , , , , , , , 0, val(DcbUnitType.BoundText), val(DcbUnitNo.BoundText)) = False Then
           '             GoTo ErrTrap
           '         End If
           '           lineno = lineno + 1
                      
              End If
              
             'salimhere
                total_value = val(XPTxtVal.Text) + val(fittervat)
             If total_value > 0 Then
                    If ModAccounts.AddNewDev(LngDevID, lineno, DcboDebitSide.BoundText, total_value, 0, Msg & " " & "«·’‰œÊﬁ/«·»‰ﬂ", val(notes_id), , , , XPDtbTrans.value, user_id, , , , , , , , , CURRENT_LINE, , , , , , , , , val(Dcbranch.BoundText), , , , , , , , , , 0, val(DcbUnitType.BoundText), val(DcbUnitNo.BoundText)) = False Then
                        GoTo ErrTrap
                    End If
                    lineno = lineno + 1
                     
                         If ModAccounts.AddNewDev(LngDevID, lineno, AccountCodeDept, total_value, 1, Msg & " " & "«·„” «Ã—", val(notes_id), , , , XPDtbTrans.value, user_id, , , , , , , , , CURRENT_LINE, , , , , , , , , val(Dcbranch.BoundText), , , , , , , , , , val(DcbIqara.BoundText), val(DcbUnitType.BoundText), val(DcbUnitNo.BoundText)) = False Then
                        GoTo ErrTrap
                    End If
                      lineno = lineno + 1
              End If
              

        
    DoEvents
    cProgress.FinishProgress
    cProgress.StopProgess
    Set cProgress = Nothing
    
ErrTrap:
End Function

' Function MyOwnerNoreatJlInContractFiter2222(LngDevID As Long, notes_id As Double) As Double
'
'If DCboCashType.ListIndex <> 10 Then Exit Function
'Dim Percetage As Double
'Dim commissionvalue As Double
'Dim total_value As Double
'Dim cProgress As ClsProgress
'Set cProgress = New ClsProgress
'    cProgress.ProgressType = Waiting
' Dim foxy_ked_NO As String
' Dim credit_side As String
' Dim My_SQL As String
' Dim Line1 As Double
' Dim lineno As Double
' Dim AccountCodeDept As String
'         lineno = 1
' Dim AccountCode As String
'    cProgress.StartProgress
'    DoEvents
'    LngDevID = new_id("DOUBLE_ENTREY_VOUCHERS", "Double_Entry_Vouchers_ID", "")
'    Dim Msgdes As String
'    Dim CURRENT_LINE As Double
'    Dim depit_side As String
'    Dim Msg As String
'     Dim i As Integer
'Msgdes = "»‰«¡ ⁄·Ï „ﬁ»Ê÷«   ’›Ì… «„·«ﬂÌ  —ﬁ„ " & TxtNoteSerial1.Text & " "
'Dim AccountCodeVat As String
'Msg = XPMTxtRemarks.Text & Chr(13) & Msgdes
' AccountCodeDept = ModAccounts.GetMyAccountCode("TblCustemers", "CusID", val(DBCboClientName.BoundText), "Account_code")
'                total_value = val(TxtTotalInsurances.Text) + val(XPTxtVal.Text)
'                     PercentgValueAddedAccount_Transec XPDtbTrans.value, 8, 1, AccountCodeVat, Percetage
'                commissionvalue = total_value * Percetage / 100
'              commissionvalue = Round(commissionvalue, 2)
'
'             If total_value > 0 Then
'                    If ModAccounts.AddNewDev(LngDevID, lineno, AccountCodeDept, total_value + commissionvalue, 0, Msg & " " & "«·„” «Ã— ", val(notes_id), , , , XPDtbTrans.value, user_id, , , , , , , , , CURRENT_LINE, , , , , , , , , val(Dcbranch.BoundText)) = False Then
'                        GoTo ErrTrap
'                    End If
'                    lineno = lineno + 1
'
'                    AccountCode = get_account_code_branch(86, my_branch)
'                         If ModAccounts.AddNewDev(LngDevID, lineno, AccountCode, total_value, 1, Msg & " " & " «Ì—«œ«  «·«ÌÃ«—", val(notes_id), , , , XPDtbTrans.value, user_id, , , , , , , , , CURRENT_LINE, , , , , , , , , val(Dcbranch.BoundText)) = False Then
'                        GoTo ErrTrap
'                    End If
'                      lineno = lineno + 1
'                    If commissionvalue > 0 Then
'
'                         If ModAccounts.AddNewDev(LngDevID, lineno, AccountCodeVat, commissionvalue, 1, Msg & " " & " «·ﬁÌ„… «·„÷«›… ", val(notes_id), , , , XPDtbTrans.value, user_id, , , , , , , , , CURRENT_LINE, , , , , , , , , val(Dcbranch.BoundText)) = False Then
'                        GoTo ErrTrap
'                    End If
'                      lineno = lineno + 1
'                    End If
'              End If
'              ''''// «· «„Ì‰
'
'            total_value = val(TxtTotalInsurances.Text)
'             If total_value > 0 Then
'                    AccountCode = get_account_code_branch(82, my_branch)
'                    If ModAccounts.AddNewDev(LngDevID, lineno, AccountCode, total_value, 0, Msg & " " & " «„Ì‰ „” —œ", val(notes_id), , , , XPDtbTrans.value, user_id, , , , , , , , , CURRENT_LINE, , , , , , , , , val(Dcbranch.BoundText)) = False Then
'                        GoTo ErrTrap
'                    End If
'                    lineno = lineno + 1
'
'                         If ModAccounts.AddNewDev(LngDevID, lineno, AccountCodeDept, total_value, 1, Msg & " " & " «„Ì‰ „” —œ", val(notes_id), , , , XPDtbTrans.value, user_id, , , , , , , , , CURRENT_LINE, , , , , , , , , val(Dcbranch.BoundText)) = False Then
'                        GoTo ErrTrap
'                    End If
'                      lineno = lineno + 1
'              End If
'
'
'
'                total_value = val(XPTxtVal.Text) + commissionvalue
'             If total_value > 0 Then
'                    If ModAccounts.AddNewDev(LngDevID, lineno, DcboDebitSide.BoundText, total_value, 0, Msg & " " & "«·’‰œÊﬁ/«·»‰ﬂ", val(notes_id), , , , XPDtbTrans.value, user_id, , , , , , , , , CURRENT_LINE, , , , , , , , , val(Dcbranch.BoundText)) = False Then
'                        GoTo ErrTrap
'                    End If
'                    lineno = lineno + 1
'
'                         If ModAccounts.AddNewDev(LngDevID, lineno, AccountCodeDept, total_value, 1, Msg & " " & "«·„” «Ã—", val(notes_id), , , , XPDtbTrans.value, user_id, , , , , , , , , CURRENT_LINE, , , , , , , , , val(Dcbranch.BoundText)) = False Then
'                        GoTo ErrTrap
'                    End If
'                      lineno = lineno + 1
'              End If
'
'
'
'    DoEvents
'    cProgress.FinishProgress
'    cProgress.StopProgess
'    Set cProgress = Nothing
'
'ErrTrap:
'End Function
Private Sub SaveData()
    Dim Msg As String
    Dim RsTemp As New ADODB.Recordset
    Dim RsDetails1  As New ADODB.Recordset
  Dim xx As Integer
  xx = 1
  Dim i As Integer
    Dim StrSQL As String
    Dim StrTemp As String
    Dim LngDevID As Long
    Dim RsDev As ADODB.Recordset
Set RsDev = New ADODB.Recordset
    Dim BeginTrans As Boolean
'  On Error GoTo ErrTrap
lineno = 1
If SystemOptions.IsElecWaterCont Then
    Dim Commissions  As Double, CommissionsPayed As Double, Insurance As Double, InsurancePayed As Double, Water As Double, WaterPayed As Double, Electric As Double, ElectricPayed As Double
    Dim TelandNet  As Double, TelandNetPayed As Double
    With Me.Grid3

        

 
 
        For i = 1 To .rows - 1
            Commissions = val(.TextMatrix(i, .ColIndex("Commissions")))
            CommissionsPayed = val(.TextMatrix(i, .ColIndex("CommissionsPayed")))
            Insurance = val(.TextMatrix(i, .ColIndex("Insurance")))
            InsurancePayed = val(.TextMatrix(i, .ColIndex("InsurancePayed")))
            Water = val(.TextMatrix(i, .ColIndex("Water")))
            WaterPayed = val(.TextMatrix(i, .ColIndex("WaterPayed")))
            Electric = val(.TextMatrix(i, .ColIndex("Electric")))
            ElectricPayed = val(.TextMatrix(i, .ColIndex("ElectricPayed")))
            TelandNet = val(.TextMatrix(i, .ColIndex("TelandNet")))
            TelandNetPayed = val(.TextMatrix(i, .ColIndex("TelandNetPayed")))
            If .cell(flexcpChecked, i, .ColIndex("Select")) = flexChecked Then
                If val(.TextMatrix(i, .ColIndex("ActRent"))) = 0 Then
                        If CommissionsPayed = 0 And InsurancePayed = 0 And WaterPayed = 0 And ElectricPayed = 0 And TelandNetPayed = 0 Then
                        
                        If Commissions <> 0 And CommissionsPayed = 0 Then
                                MsgBox "·« Ì„ﬂ‰ «·Õ›Ÿ ﬁ»· ”œ«œ «·”⁄Ì"
                                Exit Sub
                        End If
                        If Insurance <> 0 And InsurancePayed = 0 Then
                                MsgBox "·« Ì„ﬂ‰ «·Õ›Ÿ ﬁ»· ”œ«œ «· √„Ì‰"
                                Exit Sub
                        End If
                        If Water <> 0 And WaterPayed = 0 Then
                                MsgBox "·« Ì„ﬂ‰ «·Õ›Ÿ ﬁ»· ”œ«œ «·„Ì«Â"
                                Exit Sub
                        End If
                        If Electric <> 0 And ElectricPayed = 0 Then
                                MsgBox "·« Ì„ﬂ‰ «·Õ›Ÿ ﬁ»· ”œ«œ «·ﬂÂ—»«¡"
                                Exit Sub
                        End If

                        If TelandNet <> 0 And TelandNetPayed = 0 Then
                                MsgBox "·« Ì„ﬂ‰ «·Õ›Ÿ ﬁ»· ”œ«œ «·Œœ„« "
                                Exit Sub
                        End If

                       End If

                Else
                   If (Commissions <> CommissionsPayed Or _
                   Insurance <> InsurancePayed Or _
                   Water <> WaterPayed Or _
                   Electric <> ElectricPayed Or _
                   TelandNet <> TelandNetPayed) Then
                   
                       MsgBox "·« Ì„ﬂ‰ «·Õ›Ÿ ﬁ»· ”œ«œ «·ﬂÂ—»«¡ Ê»«ﬁÏ «·Œœ„« "
                       Exit Sub
                End If
                End If
             End If
           Next i
    End With
 End If

Dim suma As Double
suma = 0
   my_branch = val(Dcbranch.BoundText)
  If Me.DCboCashType.ListIndex = 8 And PayVAT() = False Then
    If SystemOptions.CanPartialpayment = False Then
        MsgBox "Ì—ÃÏ œ›⁄ «·ﬁÌ„… «·„÷«›… ﬂ«„·…"
        Exit Sub
    End If
    
  End If
If Me.DCboCashType.ListIndex = 8 And Rd(1).value = True Then
  
  RentAccount = get_account_code_branch(123, my_branch)
 
        If RentAccount = "NO branch" Then
                            If SystemOptions.UserInterface = ArabicInterface Then
                                MsgBox "·„ Ì „ «‰‘«¡ «·›—⁄", vbCritical
                            Else
                                MsgBox "Branch Not Created ", vbCritical
                            End If
                
                            GoTo ErrTrap
        ElseIf RentAccount = "NO account" Then

            If SystemOptions.UserInterface = ArabicInterface Then
                MsgBox "·„ Ì „  ÕœÌœ Õ”«»         «ÌÃ«—«  „” Õﬁ… ··€Ì— ", vbCritical
            Else
                MsgBox "   Insatllemts Revenu Not Deined in this Branch", vbCritical
            End If

            GoTo ErrTrap
         
        End If
           
 
    


CommissionAcc = get_account_code_branch(81, my_branch)
 
        If CommissionAcc = "NO branch" Then
                            If SystemOptions.UserInterface = ArabicInterface Then
                                MsgBox "·„ Ì „ «‰‘«¡ «·›—⁄", vbCritical
                            Else
                                MsgBox "Branch Not Created ", vbCritical
                            End If
                
                            GoTo ErrTrap
        ElseIf CommissionAcc = "NO account" Then

            If SystemOptions.UserInterface = ArabicInterface Then
                MsgBox "·„ Ì „  ÕœÌœ Õ”«» «Ì—«œ«  ⁄„Ê·«  «„·«ﬂ «·€Ì— ”⁄Ï/⁄„Ê·«  ", vbCritical
            Else
                MsgBox "   Insatllemts Revenu Not Deined in this Branch", vbCritical
            End If

            GoTo ErrTrap
         
        End If
           
 If val(TxtKickbacks.Text) > 0 Then
    CommissionAccDue = get_account_code_branch(125, my_branch)
 
        If CommissionAccDue = "NO branch" Then
                            If SystemOptions.UserInterface = ArabicInterface Then
                                MsgBox "·„ Ì „ «‰‘«¡ «·›—⁄", vbCritical
                            Else
                                MsgBox "Branch Not Created ", vbCritical
                            End If
                
                            GoTo ErrTrap
        ElseIf CommissionAccDue = "NO account" Then

            If SystemOptions.UserInterface = ArabicInterface Then
                MsgBox "·„ Ì „  ÕœÌœ Õ”«»   ⁄„Ê·«  „” Õﬁ…  „‰ «„·«ﬂ «·€Ì—   ", vbCritical
            Else
                MsgBox "   Insatllemts Revenu Not Deined in this Branch", vbCritical
            End If

            GoTo ErrTrap
         
        End If
           
End If


  End If
  
  
    If Me.TxtModFlg.Text <> "R" Then
    If val(DCboCashType.ListIndex) = 8 Then
' With Grid3
'     If Grid3.Rows > 1 Then
''
 '       For i = .FixedRows To .Rows - 1
 '       suma = 0
 '           If .Cell(flexcpChecked, i, .ColIndex("Select")) = flexChecked Then
 '                 suma = suma + val(.TextMatrix(i, .ColIndex("RentValuePayed"))) + val(.TextMatrix(i, .ColIndex("CommissionsPayed")))
 '                 suma = suma + val(.TextMatrix(i, .ColIndex("WaterPayed"))) + val(.TextMatrix(i, .ColIndex("ElectricPayed")))
 '                 suma = suma + val(.TextMatrix(i, .ColIndex("TelandNetPayed")))
 ''               If suma = 0 Then
  '     MsgBox "·«Ì„ﬂ‰ «·Õ›Ÿ Ì—ÃÏ «œŒ«· ﬁÌ„… «·œ›⁄«  ›Ì «·”ÿ—" & i
  '     Exit Sub
  '     End If
  '
  '      End If
  '      Next i
      
 'End If
'End With
End If
        If DCboCashType.ListIndex = -1 Then
            Msg = "ÌÃ»  ÕœÌœ ‰Ê⁄ «·„ﬁ»Ê÷«  "
            MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
            DCboCashType.SetFocus
            Sendkeys "{F4}"
            Exit Sub
        End If


        If Me.DCboCashType.ListIndex = 3 Then
            If val(Me.DcboRevenuesTypes.BoundText) = 0 Then
                Msg = "ÌÃ»  ÕœÌœ ‰Ê⁄ «·≈Ì—«œ«  «·√Œ—Ï...!!!"
                MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title

                If Me.DcboRevenuesTypes.Visible = True Then
                    DcboRevenuesTypes.SetFocus
                    Sendkeys "{F4}"
                End If

                Exit Sub
            End If
        End If

        If Me.DCboCashType.ListIndex = 0 Or Me.DCboCashType.ListIndex = 1 Or Me.DCboCashType.ListIndex = 2 Then
            If DBCboClientName.Text = "" Then
                Msg = "ÌÃ» «Œ Ì«— «”„ «·⁄„Ì· √Ê «·„Ê—œ"
                MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
                DBCboClientName.SetFocus
                Sendkeys "{F4}"
                Exit Sub
            End If
        End If
    
           If Me.DCboCashType.ListIndex = 8 Then
            If txtContractNo.Text = "" Then
                Msg = "ÌÃ» « œŒ«· —ﬁ„ «·⁄ﬁœ"
                MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
                txtContractNo.SetFocus
                Sendkeys "{F4}"
                Exit Sub
            End If
        End If
         
           If Me.DCboCashType.ListIndex = 13 Then
            If txtContractNo.Text = "" Then
                Msg = "ÌÃ» « œŒ«· —ﬁ„ «· ’›Ì«  Ê›Ê« Ì— «·ﬂÂ—»«¡ "
                MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
                txtContractNo.SetFocus
                Sendkeys "{F4}"
                Exit Sub
            End If
        End If
        If Me.DCboCashType.ListIndex = 5 Then
            If DBCboClientName.Text = "" Then
                Msg = "ÌÃ» «Œ Ì«— «”„ ««·„‘—Ê⁄"
                MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
                DBCboClientName.SetFocus
                Sendkeys "{F4}"
                Exit Sub
            End If
        End If
    
        If Me.DCboCashType.ListIndex = 6 Then
            If DCEmployee.BoundText = "" Then
                Msg = "ÌÃ» «Œ Ì«— «”„ «·„ÊŸ›"
                MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
                DCEmployee.SetFocus
                Sendkeys "{F4}"
                Exit Sub
            End If
        End If
    
        If Me.DCboCashType.ListIndex = 7 Then
            If Me.DCAccounts.BoundText = "" Then
                If SystemOptions.UserInterface = ArabicInterface Then
                    Msg = "ÌÃ» «Œ Ì«— «”„ «·Õ”«»"
                Else
                    Msg = "Select Account Firstly"
                End If

                MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
                DCAccounts.SetFocus
                Sendkeys "{F4}"
                Exit Sub
            End If
        End If
    
            If Me.DCboCashType.ListIndex = 8 Then
            If Me.TxtContNo.Text = "" Then
                If SystemOptions.UserInterface = ArabicInterface Then
                    Msg = "ÌÃ»      «Œ Ì«— ⁄ﬁœ "
                Else
                    Msg = "Select Contract Firstly"
                End If

                MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
                TxtContNo.SetFocus
                Sendkeys "{F4}"
                Exit Sub
            End If
         If xx = 1 Then
          With VSFlexGrid2
    If val(.rows) >= 2 Then
    If val(.TextMatrix(1, .ColIndex("id"))) = 0 Then
    If SystemOptions.UserInterface = ArabicInterface Then
    MsgBox "Ì—ÃÏ «Œ Ì«— «·„‰œÊ»"
    Else
    MsgBox "Select salesperson"
    End If
    Exit Sub
    End If
       Else
          If SystemOptions.UserInterface = ArabicInterface Then
    MsgBox "Ì—ÃÏ «Œ Ì«— «·„‰œÊ»"
    Else
    MsgBox "Select salesperson"
    End If
       .rows = .rows + 1
    Exit Sub
    
    End If
    
    End With
        
    
         End If
        End If
'       If Me.DCboCashType.ListIndex <> 8 Then
''               If xx = 1 Then
'           With VSFlexGrid1
'     If val(.Rows) >= 2 Then
'     If val(.TextMatrix(1, .ColIndex("id"))) = 0 Then
'    If SystemOptions.UserInterface = ArabicInterface Then
'     MsgBox "Ì—ÃÏ «Œ Ì«— «·„‰œÊ»"
'     Else
'     MsgBox "Select salesperson"
'     End If
'     Exit Sub
'    End If
'        Else
'           If SystemOptions.UserInterface = ArabicInterface Then
'     MsgBox "Ì—ÃÏ «Œ Ì«— «·„‰œÊ»"
'     Else
'     MsgBox "Select salesperson"
'     End If
'        .Rows = .Rows + 1
'    Exit Sub
    
'     End If
    
'    End With
'     End If
       End If
    
           If Me.DCboCashType.ListIndex = 9 Then
         
            If Me.DcbIqara.BoundText = "" Then
                If SystemOptions.UserInterface = ArabicInterface Then
                    Msg = "ÌÃ» «Œ Ì«— «”„ «·⁄ﬁ«—"
                Else
                    Msg = "Select entity Firstly"
                End If

                MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
                DcbIqara.SetFocus
                Sendkeys "{F4}"
                Exit Sub
            End If
             If Me.DCboCashType.ListIndex = 10 Then
            If Me.TxtFilterNo.Text = "" Then
                If SystemOptions.UserInterface = ArabicInterface Then
                    Msg = "ÌÃ» «Œ Ì«—  —ﬁ„ «· ’›ÌÂ"
                Else
                    Msg = "Select entity Firstly"
                End If

                MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
                Me.TxtFilterNo.SetFocus
                Sendkeys "{F4}"
                Exit Sub
            End If
            End If
            
                     If Me.DcbUnitType.BoundText = "" Then
                If SystemOptions.UserInterface = ArabicInterface Then
                    Msg = "ÌÃ» «Œ Ì«—    ‰Ê⁄ «·ÊÕœ…"
                Else
                    Msg = "Select unit type Firstly"
                End If

                MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
                DcbUnitType.SetFocus
                Sendkeys "{F4}"
                Exit Sub
            End If
            
                      If Me.DcbUnitNo.BoundText = "" Then
                If SystemOptions.UserInterface = ArabicInterface Then
                    Msg = "ÌÃ» «Œ Ì«—    —ﬁ„ «·ÊÕœ…   "
                Else
                    Msg = "Select unit no Firstly"
                End If

                MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
                DcbUnitNo.SetFocus
                Sendkeys "{F4}"
                Exit Sub
            End If
            
            
           If Me.TxtInterval.Text = "" Then
                If SystemOptions.UserInterface = ArabicInterface Then
                    Msg = "ÌÃ»   ÕœÌœ «·„œ…   "
                Else
                    Msg = "Select Account Firstly"
                End If

                MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
                TxtInterval.SetFocus
                Sendkeys "{F4}"
                Exit Sub
            End If
            
        End If
        
        If XPTxtVal.Text = "" Then
            Msg = "ÌÃ» «œŒ«· ﬁÌ„… «·„ﬁ»Ê÷«  "
            MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
            '        XPTxtVal.SetFocus
            Exit Sub
        End If

        If Not IsNumeric(XPTxtVal.Text) Then
            Msg = "ﬁÌ„… «·„ﬁ»Ê÷«  ÌÃ» √‰  ﬂÊ‰ ﬁÌ„… —ﬁ„Ì…"
            MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
            XPTxtVal.SetFocus
            SelectText XPTxtVal
            Exit Sub
        End If

        If Me.ChkTrans.value = vbChecked Then
            If Me.CboTrans.ListIndex = -1 Then
                Msg = "»—Ã«¡ ≈Œ Ì«— ‰Ê⁄ «·›« Ê—…..!!!"
                MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
                CboTrans.SetFocus
                Sendkeys "{F4}"
                Exit Sub
            End If

            If Trim(Me.TxtTransSerial.Text) = "" Then
                Msg = "»—Ã«¡ ≈œŒ«· —ﬁ„ «·›« Ê—…..!!!"
                MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
                TxtTransSerial.SetFocus
                Exit Sub
            Else

                If Me.CboTrans.ListIndex = 0 Then
                    StrTemp = GetTransIDSerial(0, , Trim(Me.TxtTransSerial.Text), 2)

                    If CheckDebitTrans(val(StrTemp)) = False Then
                        Exit Sub
                    End If

                ElseIf Me.CboTrans.ListIndex = 1 Then
                    StrTemp = GetTransIDSerial(0, , Trim(Me.TxtTransSerial.Text), 5)

                    If CheckDebitTrans(val(StrTemp)) = False Then
                        Exit Sub
                    End If

                ElseIf Me.CboTrans.ListIndex = 2 Then

                    If CheckDebitMaintaince(val(Me.TxtTransSerial.Text)) = False Then
                        Exit Sub
                    End If

                ElseIf Me.CboTrans.ListIndex = 3 Then
                    Msg = "⁄›Ê« .. Ã«—Ï  ÿÊÌ— «·»—‰«„Ã .. ·⁄„· «·„ﬁ»Ê÷«  „‰ «·Œœ„« "
                    MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
                    Exit Sub
                End If
            End If
        End If

        If Me.CboPaymentType.ListIndex = -1 Then
            Msg = "ÌÃ» ≈Œ Ì«— ÿ—Ìﬁ… «·œ›⁄...!!"
            MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
            CboPaymentType.SetFocus
            Sendkeys "{F4}"
            Exit Sub
        End If
        If Me.CboPaymentType.ListIndex = 4 Then
        If DcbAccount.BoundText = "" Or DcbAccount.Text = "" Then
        If SystemOptions.UserInterface = ArabicInterface Then
        MsgBox "Ì—ÃÏ «Œ Ì«— «·Õ”«»"
        Else
        MsgBox "Please Select Account"
        End If
        DcbAccount.SetFocus
        Exit Sub
        End If
        End If
        If Me.CboPaymentType.ListIndex = 0 Then
            If Me.DcboBox.BoundText = "" Then
                Msg = "ÌÃ» ≈Œ Ì«— «·Œ“‰…...!!"
                MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
                DcboBox.SetFocus
                Sendkeys "{F4}"
                Exit Sub
            End If

        ElseIf Me.CboPaymentType.ListIndex = 1 Then
      
            '  If DateDiff("d", Me.DtpChequeDueDate.value, Date) > 0 Then
            '      Msg = " «—ÌŒ ≈” Õﬁ«ﬁ «·‘Ìﬂ €Ì— ’ÕÌÕ...!!"
            '      MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
            '      DtpChequeDueDate.SetFocus
            '      SendKeys "{F4}"
            '      Exit Sub
            '  End If
            If SystemOptions.ChequeBox = True Then
         
                If DCChequeBox.BoundText = "" Then
                           
                    If SystemOptions.UserInterface = ArabicInterface Then
                        Msg = "Õœœ Õ«›Ÿ… «·‘Ìﬂ«  ...!!"
                    Else
                        Msg = "Select Cheque Box ...!!"
                    End If

                    MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
                    DCChequeBox.SetFocus
                    Sendkeys "{F4}"
                    Exit Sub
                   
                End If
    
                If TXTBankName.Text = "" Then
                           
                    If SystemOptions.UserInterface = ArabicInterface Then
                        Msg = "«ﬂ » «”„ »‰ﬂ «·‘Ìﬂ    « ...!!"
                    Else
                        Msg = " Enter Bank Name For Cheque  ...!!"
                    End If

                    MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
                    TXTBankName.SetFocus
                    Sendkeys "{F4}"
                    Exit Sub
                    
                End If
        
                If Trim$(Me.TxtChequeNumber.Text) = "" Then
                    Msg = "ÌÃ» ﬂ «»… —ﬁ„ «·‘Ìﬂ...!!"
                    MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
                    TxtChequeNumber.SetFocus
                    Exit Sub
                End If

            Else
       
                If Me.DcboBankName.BoundText = "" Then
                    Msg = "ÌÃ» ≈Œ Ì«— «·»‰ﬂ...!!"
                    MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
                    DcboBankName.SetFocus
                    Sendkeys "{F4}"
                    Exit Sub
                End If

                If Trim$(Me.TxtChequeNumber.Text) = "" Then
                    Msg = "ÌÃ» ﬂ «»… —ﬁ„ «·‘Ìﬂ...!!"
                    MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
                    TxtChequeNumber.SetFocus
                    Exit Sub
                End If
            End If
    
        ElseIf Me.CboPaymentType.ListIndex = 2 Then

            If Me.DcboBankName.BoundText = "" Then
                Msg = "ÌÃ» ≈Œ Ì«— «·»‰ﬂ...!!"
                MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
                DcboBankName.SetFocus
                Sendkeys "{F4}"
                Exit Sub
            End If

            If Trim$(Me.TxtChequeNumber.Text) = "" Then
                Msg = "ÌÃ» ﬂ «»… —ﬁ„ «·ÕÊ«·Â...!!"
                MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
                TxtChequeNumber.SetFocus
                Exit Sub
            End If
     
        ElseIf Me.CboPaymentType.ListIndex = 3 Then

            If Me.DcboBankName.BoundText = "" Then
                Msg = "ÌÃ» ≈Œ Ì«— «·»‰ﬂ...!!"
                MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
                DcboBankName.SetFocus
                Sendkeys "{F4}"
                Exit Sub
            End If

            If Trim$(Me.TxtChequeNumber.Text) = "" Then
                Msg = "ÌÃ» ﬂ «»… —ﬁ„ «·‘Ìﬂ...!!"
                MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
                TxtChequeNumber.SetFocus
                Exit Sub
            End If
     
        End If

        Dim notes_result As String
        Dim Vchr_result As String
my_branch = val(Me.Dcbranch.BoundText)

        If TxtNoteSerial1.Text = "" Then
            Vchr_result = Voucher_coding(val(my_branch), XPDtbTrans.value, 2, 4)

            If Vchr_result = "error" Then
                MsgBox " ·« Ì„ﬂ‰ «÷«›… ”‰œ ﬁ»÷ ÃœÌœ ·«‰ﬂ  ⁄œÌ  «·Õœ «·–Ì ﬁ„  » ÕœÌœ… „‰ «·”‰œ«   ": Exit Sub
            Else
                
                If Vchr_result = "" Then
                    MsgBox " ·«»œ „‰ ﬂ «»… —ﬁ„ «·”‰œ ÌœÊÌ« ﬂ„« Õœœ   ": Exit Sub
                Else
                    ' txtNoteSerial1.text = Voucher_coding(val(my_branch), XPDtbTrans.value, 2, 4)
                End If
            End If
        End If
    
        If TxtNoteSerial.Text = "" Then
            notes_result = Notes_coding(val(my_branch), XPDtbTrans.value)

            If notes_result = "error" Then
                MsgBox " ·« Ì„ﬂ‰ «÷«›… ﬁÌÊœ ·Â–… «·⁄„·Ì… ·«‰ﬂ  ⁄œÌ  «·Õœ «·«ﬁ’Ì ··ﬁÌÊœ ﬂ„« Õœœ  ›Ì  —ﬁÌ„ «·”‰œ«  ": Exit Sub
            Else
                       
                If notes_result = "" Then
                    MsgBox " ·«Ì„ﬂ‰ «‰‘«¡ «·ﬁÌœ ·Â–« «·„” ‰œ ·«‰ﬂ Õœœ   —ﬁÌ„ ﬁÌÊœ ÌœÊÌ  ": Exit Sub
                Else
                    '     TxtNoteSerial.text = Notes_coding(val(my_branch), XPDtbTrans.value)
                End If
            End If
        End If
    
        Cn.BeginTrans
        BeginTrans = True

        If TxtModFlg.Text = "N" Then
            XPTxtID.Text = CStr(new_id("Notes", "NoteID", "", True))
            'Me.TxtNoteSerial.text = CStr(new_id("Notes", "NoteSerial", "", True, "NoteType=4"))
            rs.AddNew
       
            rs("NoteID").value = val(XPTxtID.Text)
            Me.oldtxtNoteSerial1.Text = Trim$(Me.TxtNoteSerial1.Text)
         
        ElseIf TxtModFlg.Text = "E" Then
        StrSQL = "Delete From TblUnitNoInformation Where NoteID=" & val(XPTxtID.Text)
            Cn.Execute StrSQL, , adExecuteNoRecords
            
            StrSQL = "Delete From DOUBLE_ENTREY_VOUCHERS Where Notes_ID=" & val(XPTxtID.Text)
            Cn.Execute StrSQL, , adExecuteNoRecords
        
        StrSQL = "Delete From DOUBLE_ENTREY_VOUCHERS Where Notes_ID=" & val(XPTxtID.Text)
            Cn.Execute StrSQL, , adExecuteNoRecords
            
            StrSQL = "Delete  TblNotesSales  where NoteID =" & val(XPTxtID.Text)
            Cn.Execute StrSQL, , adExecuteNoRecords
            
StrSQL = " delete   TblAqrEarnest where    NoteID=" & val(XPTxtID.Text)
Cn.Execute StrSQL

   StrSQL = " delete   notes where NoteType= 2000   and  NoteSerial='" & TxtNoteSerial.Text & "'"
  
Cn.Execute StrSQL
StrSQL = " delete   TblAqarCommissions where    NoteID=" & val(XPTxtID.Text)
Cn.Execute StrSQL
   StrSQL = "Delete From ContracttBillInstallmentsDone Where NoteID =" & val(Me.XPTxtID)
                Cn.Execute StrSQL, , adExecuteNoRecords
 
         End If


             If DCboCashType.ListIndex = 5 Then
                '«·„‘«—Ì⁄
                Dim pstate As Integer
          
             '   account_codeLegal = get_project_Account(val(DBCboClientName.BoundText), "legal")
                     pstate = val(get_project_Account(val(DBCboClientName.BoundText), "pstate"))

        If pstate = 1 Then Option7.value = True Else Option7.value = False


      End If
      AqrCommisiion (val(Me.DCboCashType.ListIndex))
        rs("branch_no").value = val(Me.Dcbranch.BoundText)
        rs("EmpId").value = IIf(Me.DcEmp.BoundText = "", Null, (Me.DcEmp.BoundText))
        rs("foxy_no").value = val(Text1.Text)
        rs("general_cost_center").value = IIf(Me.DcCostCenter.BoundText = "", "", Me.DcCostCenter.BoundText)
    
        rs("CarId").value = IIf(Me.DCCar.BoundText = "", Null, (Me.DCCar.BoundText))
        rs("DriverId").value = IIf(Me.DCDriver.BoundText = "", Null, (Me.DCDriver.BoundText))
    
        If TxtNoteSerial1.Text = "" Then
            TxtNoteSerial1.Text = Voucher_coding(val(my_branch), XPDtbTrans.value, 2, 4)
        End If
      rs("CreditSide").value = IIf(Trim(DcboCreditSide.BoundText) = "", Null, (DcboCreditSide.BoundText))
       rs("DebitSide").value = IIf(Trim(DcboDebitSide.BoundText) = "", Null, (DcboDebitSide.BoundText))
        If TxtNoteSerial.Text = "" Then
            TxtNoteSerial.Text = Notes_coding(val(my_branch), XPDtbTrans.value)
        End If
             If CheckStatusEarnest(0).value = vbChecked Then
         rs("StatusEarnest").value = 1
    
         ElseIf CheckStatusEarnest(1).value = vbChecked Then
         rs("StatusEarnest").value = 2
         ElseIf CheckStatusEarnest(2).value = vbChecked Then
         rs("StatusEarnest").value = 3
        ElseIf CheckStatusEarnest(3).value = vbChecked Then
         rs("StatusEarnest").value = 4

         Else
          rs("StatusEarnest").value = 0
        End If
        If ComResid(1).value = True Then
        rs("ComResid").value = 1
        Else
        rs("ComResid").value = 0
        End If
          rs("VAT").value = IIf(TxtVATValue.Text = "", Null, val(TxtVATValue.Text))
          If DCboCashType2.ListIndex = 3 Then ' ’›Ì…
         rs("VAT").value = fittervat
          End If
          
  If Rd(1).value = True Then
  rs.Fields("TypAmola").value = 1
  rs("AmolaValus").value = IIf(TxtKickbacks.Text = "", 0, val(TxtKickbacks.Text))
  End If
        rs("Servce").value = IIf(TxtService.Text = "", 0, val(TxtService.Text))
        rs("commission").value = IIf(Txtcommission.Text = "", 0, Trim(Txtcommission.Text))
        rs("CommissionOut").value = IIf(Me.TxtCommissionOut.Text = "", 0, Trim(TxtCommissionOut.Text))
        rs("rent").value = IIf(TxtRent.Text = "", 0, Trim(TxtRent.Text))
        rs("Water").value = IIf(TxtWater.Text = "", 0, Trim(TxtWater.Text))
        rs("Electricity").value = IIf(Me.TxtElectricity.Text = "", 0, Trim(TxtElectricity.Text))
        rs("Instrunce").value = IIf(txtinstrunce.Text = "", 0, Trim(txtinstrunce.Text))
        rs("comX").value = IIf(txtComisin.Text = "", 0, Trim(txtComisin.Text))
        rs("ComY").value = IIf(txtinstranc.Text = "", 0, Trim(txtinstranc.Text))
        rs("Telephone").value = IIf(TxtTelphone.Text = "", "", Trim(TxtTelphone.Text))
        
    If Option1.value = True Then
       rs("NCashingType").value = 1
   ElseIf Option2.value = True Then
        rs("NCashingType").value = 2
   ElseIf Option3.value = True Then
        rs("NCashingType").value = 3
       ElseIf Option7.value = True Then
        rs("NCashingType").value = 7
        
    Else
    
         rs("NCashingType").value = 0
   End If
  If val(DCboCashType.ListIndex) = 13 Then
  If RdTypeTrans(1).value = True Then
  rs("TypeTrans").value = 1
    If val(TxtPrice3.Text) + val(TxtPrice.Text) = val(TxtPrice2.Text) Then
  Cn.Execute "Update TblOtheExpensAqar set FlgPayed=1 where ID=" & val(Me.TxtContNo.Text) & "  "
  End If
 ' Cn.Execute "Update TblOtheExpensAqar set FlgPayed=1 where ID=" & val(Me.TXTContNo.Text) & "  "
  Else
  rs("TypeTrans").value = 0
    If val(TxtTotal23.Text) + val(txtTotal.Text) = val(TxtTotal22.Text) Then
  Cn.Execute "Update TblOtheExpensAqar set FlgPayed=1 where ID=" & val(Me.TxtContNo.Text) & "  "
  End If
  End If

       rs("Price").value = IIf(Trim(TxtPrice.Text) = "", Null, val(TxtPrice.Text))
       rs("Price2").value = IIf(Trim(TxtPrice2.Text) = "", Null, val(TxtPrice2.Text))
       rs("Price3").value = IIf(Trim(TxtPrice3.Text) = "", Null, val(TxtPrice3.Text))
       rs("RemPrice").value = IIf(Trim(TxtRemPrice.Text) = "", Null, val(TxtRemPrice.Text))
     
       
       rs("Insurance").value = IIf(Trim(txtInsurance.Text) = "", Null, val(txtInsurance.Text))
       rs("Discount").value = IIf(Trim(txtDiscount.Text) = "", Null, val(txtDiscount.Text))
       rs("Maintenance2").value = IIf(Trim(TxtMaintenance2.Text) = "", Null, val(TxtMaintenance2.Text))
       rs("Maintenance3").value = IIf(Trim(TxtMaintenance3.Text) = "", Null, val(TxtMaintenance3.Text))
       rs("Maintenance").value = IIf(Trim(TxtMaintenance.Text) = "", Null, val(TxtMaintenance.Text))
       rs("RemainRent2").value = IIf(Trim(TxtRemainRent2.Text) = "", Null, val(TxtRemainRent2.Text))
       rs("RemainRent3").value = IIf(Trim(TxtRemainRent3.Text) = "", Null, val(TxtRemainRent3.Text))
       rs("RemainRent").value = IIf(Trim(txtRemainRent.Text) = "", Null, val(txtRemainRent.Text))
       rs("MaintCondition3").value = IIf(Trim(TxtMaintCondition3.Text) = "", Null, val(TxtMaintCondition3.Text))
       rs("MaintCondition2").value = IIf(Trim(TxtMaintCondition2.Text) = "", Null, val(TxtMaintCondition2.Text))
       rs("MaintCondition").value = IIf(Trim(TxtMaintCondition.Text) = "", Null, val(TxtMaintCondition.Text))
       rs("MaintClean3").value = IIf(Trim(TxtMaintClean3.Text) = "", Null, val(TxtMaintClean3.Text))
       rs("MaintClean2").value = IIf(Trim(TxtMaintClean2.Text) = "", Null, val(TxtMaintClean2.Text))
       rs("MaintClean").value = IIf(Trim(TxtMaintClean.Text) = "", Null, val(TxtMaintClean.Text))
       rs("Paints2").value = IIf(Trim(TxtPaints2.Text) = "", Null, val(TxtPaints2.Text))
       rs("Paints3").value = IIf(Trim(TxtPaints3.Text) = "", Null, val(TxtPaints3.Text))
       rs("Paints").value = IIf(Trim(TxtPaints.Text) = "", Null, val(TxtPaints.Text))
       rs("Maintkitchen2").value = IIf(Trim(TxtMaintkitchen2.Text) = "", Null, val(TxtMaintkitchen2.Text))
       rs("Maintkitchen3").value = IIf(Trim(TxtMaintkitchen3.Text) = "", Null, val(TxtMaintkitchen3.Text))
       rs("Maintkitchen").value = IIf(Trim(TxtMaintkitchen.Text) = "", Null, val(TxtMaintkitchen.Text))
       rs("Electricity12").value = IIf(Trim(TxtElectricity12.Text) = "", Null, val(TxtElectricity12.Text))
       rs("Electricity13").value = IIf(Trim(TxtElectricity13.Text) = "", Null, val(TxtElectricity13.Text))
       rs("Electricity1").value = IIf(Trim(TxtElectricity1.Text) = "", Null, val(TxtElectricity1.Text))
       rs("MaintDoors2").value = IIf(Trim(TxtMaintDoors2.Text) = "", Null, val(TxtMaintDoors2.Text))
       rs("MaintDoors3").value = IIf(Trim(TxtMaintDoors3.Text) = "", Null, val(TxtMaintDoors3.Text))
       rs("MaintDoors").value = IIf(Trim(TxtMaintDoors.Text) = "", Null, val(TxtMaintDoors.Text))
       rs("Windows2").value = IIf(Trim(TxtWindows2.Text) = "", Null, val(TxtWindows2.Text))
       rs("Windows3").value = IIf(Trim(TxtWindows3.Text) = "", Null, val(TxtWindows3.Text))
       rs("Windows").value = IIf(Trim(TxtWindows.Text) = "", Null, val(TxtWindows.Text))
       rs("MaintOther2").value = IIf(Trim(TxtMaintOther2.Text) = "", Null, val(TxtMaintOther2.Text))
       rs("MaintOther").value = IIf(Trim(TxtMaintOther.Text) = "", Null, val(TxtMaintOther.Text))
       rs("MaintOther3").value = IIf(Trim(TxtMaintOther3.Text) = "", Null, val(TxtMaintOther3.Text))
       rs("Total22").value = IIf(Trim(TxtTotal22.Text) = "", Null, val(TxtTotal22.Text))
       rs("Total23").value = IIf(Trim(TxtTotal23.Text) = "", Null, val(TxtTotal23.Text))
       rs("Total21").value = IIf(Trim(txtTotal.Text) = "", Null, val(txtTotal.Text))
       rs("TotalAftreIns").value = IIf(Trim(TxtTotalAftreIns.Text) = "", Null, val(TxtTotalAftreIns.Text))
       rs("TotalAftreIns2").value = IIf(Trim(TxtTotalAftreIns2.Text) = "", Null, val(TxtTotalAftreIns2.Text))
       rs("TotalAftreIns3").value = IIf(Trim(TxtTotalAftreIns3.Text) = "", Null, val(TxtTotalAftreIns3.Text))
       rs("Net").value = IIf(Trim(txtNet.Text) = "", Null, val(txtNet.Text))
       rs("Net2").value = IIf(Trim(txtNet2.Text) = "", Null, val(txtNet2.Text))
       rs("Net3").value = IIf(Trim(TxtNet3.Text) = "", Null, val(TxtNet3.Text))
  
  End If
       rs("AccountPaym").value = IIf(Trim(DcbAccount.BoundText) = "", Null, DcbAccount.BoundText)
        rs("ManulaNO").value = IIf(Trim(Me.TxtManulaNO.Text) = "", Null, Trim(Me.TxtManulaNO.Text))
        rs("BookNo").value = IIf(Trim(Me.TxtBookNo.Text) = "", Null, Trim(Me.TxtBookNo.Text))
        
        rs("RemaiValue").value = IIf(Trim(Me.lblremain.Caption) = "", Null, val(Me.lblremain.Caption))
        rs("NoteSerial").value = IIf(Trim(Me.TxtNoteSerial.Text) = "", Null, Trim(Me.TxtNoteSerial.Text))
        rs("NoteSerial1").value = IIf(Trim(Me.TxtNoteSerial1.Text) = "", Null, Trim(Me.TxtNoteSerial1.Text))
        rs("OldNoteSerial1").value = Trim$(Me.oldtxtNoteSerial1.Text) '
    
        rs("person").value = IIf(txtperson.Text = "", "", Trim(txtperson.Text))
        rs("Note_Value").value = IIf(XPTxtVal.Text = "", Null, val(XPTxtVal.Text))
        rs("Adv_payment_value").value = IIf(txtAdv_payment_value.Text = "", Null, val(txtAdv_payment_value.Text))
    ''//

        
    ''//
    
        '    Rs("Remark").value = IIf(dcproject.BoundText = "", "", Trim(dcproject.BoundText))
        rs("Remark").value = IIf(XPMTxtRemarks.Text = "", "", Trim(XPMTxtRemarks.Text))
        rs("BankName").value = IIf(TXTBankName.Text = "", "", Trim(TXTBankName.Text))

        rs("NoteType").value = 4
           rs("NoteDate").value = XPDtbTrans.value
        'rs("NoteDate").value = Format$(Date, "dd-mm-yyyy")
        rs("NoteDateH").value = Me.Txt_DateHigri.value
        ''//
        rs("FrmPriodDate").value = Me.FrmPriodDate.value
        rs("FrmPriodDateH").value = Me.FrmPriodDateH.value
        rs("ToPriodDate").value = Me.ToPriodDate.value
        rs("ToPriodDateH").value = Me.ToPriodDateH.value
        rs("Remark2").value = Me.TxtRemarks.Text
        ''//

        Select Case DCboCashType.ListIndex

            Case 0, 1

                If Me.ChkTrans.value = vbChecked Then
                    If Me.CboTrans.ListIndex = 0 Or Me.CboTrans.ListIndex = 1 Then
                        rs("Transaction_ID").value = val(Me.TxtTransID.Text)
                        rs("MaintananceID").value = Null
                    ElseIf Me.CboTrans.ListIndex = 2 Then
                        rs("Transaction_ID").value = Null
                        rs("MaintananceID").value = val(Me.TxtTransID.Text)
                    End If

                Else
                    rs("Transaction_ID").value = Null
                    rs("MaintananceID").value = Null
                End If

                rs("RevenuesID").value = Null

            Case 2
                rs("Transaction_ID").value = Null
                rs("MaintananceID").value = Null
                rs("RevenuesID").value = Null

            Case 3
                rs("RevenuesID").value = val(Me.DcboRevenuesTypes.BoundText)
                rs("Transaction_ID").value = Null
                rs("MaintananceID").value = Null

            Case 4
                '       Set rs1 = New ADODB.Recordset
                '       StrSQL = "select * From Transactions"
                '       rs1.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText
                '        XPTxtBillID.text = CStr(new_id("Transactions", "Transaction_ID", "", True))
                '       rs1.AddNew
                '       rs1("Transaction_ID").value = Val(XPTxtBillID.text)
                '       rs1("Transaction_Date").value = XPDtbTrans.value
                '       rs1("Transaction_Type").value = 23
                '       rs1.update
                '
                '        Rs("Transaction_ID").value = Val(XPTxtBillID.text)
                '
        End Select
     

        rs("CashingType").value = val(DCboCashType.ListIndex)
    
        If Me.DCboCashType.ListIndex = 0 Or Me.DCboCashType.ListIndex = 1 Or Me.DCboCashType.ListIndex = 2 Or Me.DCboCashType.ListIndex = 4 Or Me.DCboCashType.ListIndex = 8 Or Me.DCboCashType.ListIndex = 9 Or Me.DCboCashType.ListIndex = 10 Or Me.DCboCashType.ListIndex = 11 Or Me.DCboCashType.ListIndex = 12 Or Me.DCboCashType.ListIndex = 13 Then
            rs("CusID").value = IIf(DBCboClientName.Text = "", Null, DBCboClientName.BoundText)
     
        ElseIf Me.DCboCashType.ListIndex = 5 Then
            Dim X As Double
                    If IsNull(rs("note_count").value) Then
                         rs("note_count").value = CStr(new_id("Notes", "note_count", " ", True, " project_id=" & val(DBCboClientName.BoundText) & ""))
                    End If
            
            If Option4.value = True Then
                X = get_project_customer_id(DBCboClientName.BoundText, "End_user_Account")
            Else
                X = get_project_customer_id(DBCboClientName.BoundText, "sub_contractor_Account")
            End If

            rs("CusID").value = X
     
        Else
            rs("CusID").value = Null
        End If

        '--------------------------------------------------------------------------
        'ÿ—Ìﬁ… «·œ›⁄ «·‰ﬁœÏ «Ê «·‘Ìﬂ
        If Me.CboPaymentType.ListIndex = 0 Then
            rs("NoteCashingType").value = 0
            rs("BoxID").value = IIf(DcboBox.BoundText = "", Null, DcboBox.BoundText)
            rs("BankID").value = Null
            rs("ChqueNum").value = Null
            rs("DueDate").value = Null
        
        ElseIf Me.CboPaymentType.ListIndex = 1 Then
            rs("NoteCashingType").value = 1
            rs("BoxID").value = Null

            If SystemOptions.ChequeBox = False Then
        
                rs("BankID").value = val(Me.DcboBankName.BoundText)
            Else
                rs("BankID").value = Null
            End If
        
            rs("ChqueNum").value = Trim$(Me.TxtChequeNumber.Text)
            rs("DueDate").value = Me.DtpChequeDueDate.value

            If SystemOptions.ChequeBox = True Then
                rs("ChequeBoxID").value = IIf(DCChequeBox.BoundText = "", Null, DCChequeBox.BoundText)
            Else
                rs("ChequeBoxID").value = Null
                
            End If
                
        ElseIf Me.CboPaymentType.ListIndex = 2 Then
            rs("NoteCashingType").value = 2
            rs("BoxID").value = Null
            rs("BankID").value = val(Me.DcboBankName.BoundText)
            rs("ChqueNum").value = Trim$(Me.TxtChequeNumber.Text)
            rs("DueDate").value = Me.DtpChequeDueDate.value
            rs("ChequeBoxID").value = Null
                
        ElseIf Me.CboPaymentType.ListIndex = 3 Then
            rs("NoteCashingType").value = 3
            rs("BoxID").value = Null
            rs("BankID").value = val(Me.DcboBankName.BoundText)
            rs("ChqueNum").value = Trim$(Me.TxtChequeNumber.Text)
            rs("DueDate").value = Me.DtpChequeDueDate.value
            rs("ChequeBoxID").value = Null
          ElseIf Me.CboPaymentType.ListIndex = 4 Then
            rs("NoteCashingType").value = 4
            rs("BoxID").value = Null
                
        End If

        '--------------------------------------------------------------------------
        rs("UserID").value = user_id
        rs("numbering_type").value = sand_numbering_type(0) '”‰œ «·ﬁÌœ
        rs("numbering_type1").value = sand_numbering_type(2) '”‰œ «·ﬁ»÷
    
        If DCboCashType.ListIndex = 5 Then
            rs("project_id").value = IIf(DBCboClientName.BoundText = "", 0, DBCboClientName.BoundText)
        End If
    
        If DCboCashType.ListIndex = 6 Then
            rs("EmployeeID").value = IIf(DCEmployee.BoundText = "", 0, DCEmployee.BoundText)
        End If
    
        If DCboCashType.ListIndex = 7 Then
            rs("AccountsCode").value = IIf(Me.DCAccounts.BoundText = "", Null, DCAccounts.BoundText)
        End If
    
       If DCboCashType.ListIndex = 8 Or DCboCashType.ListIndex = 13 Then
            rs("ContractNo").value = IIf(txtContractNo.Text = "", Null, txtContractNo.Text)
            rs("ContNo").value = IIf(TxtContNo.Text = "", Null, TxtContNo.Text)
            Else
             rs("ContractNo").value = Null
             rs("ContNo").value = Null
        End If
         If DCboCashType.ListIndex = 10 Then
            rs("FilterID").value = IIf(TxtFilterNo.Text = "", Null, TxtFilterNo.Text)
            rs("FIlterTotal").value = IIf(TXtFilter.Text = "", Null, TXtFilter.Text)
            rs("TotalInsurances").value = IIf(txtTotalinsuranceS.Text = "", Null, txtTotalinsuranceS.Text)
            Else
             rs("FilterID").value = Null
             rs("FIlterTotal").value = Null
             rs("TotalInsurances").value = Null
        End If
        
        
     rs("akarid").value = IIf(val(Me.DcbIqara.BoundText) <> 0, val(DcbIqara.BoundText), Null)
       rs.Fields("UnitType").value = IIf(Me.DcbUnitType.BoundText <> "", val(DcbUnitType.BoundText), Null)
     rs.Fields("UnitNo").value = IIf(Me.DcbUnitNo.BoundText <> "", val(DcbUnitNo.BoundText), Null)
   
     If DCboCashType.ListIndex = 9 Then
  
     
     
     rs("interval").value = IIf(TxtInterval.Text = "", Null, val(TxtInterval.Text))
     rs("intervaltype").value = val(cbointervaltype.ListIndex)
     rs("renterName").value = IIf(txtrenterName.Text = "", Null, txtrenterName.Text)
              If cbointervaltype.ListIndex = 0 Then
              rs("allowdate").value = DateAdd("d", val(TxtInterval), XPDtbTrans.value)
              ElseIf cbointervaltype.ListIndex = 1 Then
              rs("allowdate").value = DateAdd("M", val(TxtInterval), XPDtbTrans.value)
              
            ElseIf cbointervaltype.ListIndex = 2 Then
              rs("allowdate").value = DateAdd("YYYY", val(TxtInterval), XPDtbTrans.value)
             
             End If
                  rs("allowdateH").value = ToHijriDate(rs("allowdate").value)
         
            Else
         ' rs("akarid").value = Null
    ' rs.Fields("UnitType").value = Null
     'rs.Fields("UnitNo").value = Null
     rs("interval").value = Null
     rs("intervaltype").value = Null
     rs("renterName").value = Null
          
        End If

        rs("sanad_year").value = year(XPDtbTrans.value)
        rs("sanad_month").value = Month(XPDtbTrans.value)
    
        If DCboCashType.ListIndex = 5 Then
            rs("note_value_by_characters").value = WriteNo(val(Me.XPTxtVal.Text) * 2, 0, True)
        Else
            rs("note_value_by_characters").value = Trim$(Me.lbl(18).Caption)
        End If

        If Option4.value = True Then
            rs("cus_or_sub").value = 0 '⁄„Ì· ‰Â«∆Ì
        Else
            rs("cus_or_sub").value = 1 '⁄„Ì· »«ÿ‰
        End If
    
   

        saveChequeBoxContents (XPTxtID.Text)
        rs.update
           Dim IarType As Integer
            IarType = AqarCommisionType(val(DcbIqara.BoundText))
        '==========================================================================
If SystemOptions.NoCreatJLInRentContract = True And DCboCashType.ListIndex = 8 Then
            If IarType <> 0 Then
            OtherOwnerNoreatJlInContract 1, val(XPTxtID.Text)
            Else
            MyOwnerNoreatJlInContract 1, val(XPTxtID.Text)
            End If
GoTo x22
End If
If SystemOptions.NoCreatJLInRentContract = True And DCboCashType.ListIndex = 10 Then
If IarType <> 0 Then
OtherOwnerNoreatJlInContractFiter 1, val(XPTxtID.Text)
Else
MyOwnerNoreatJlInContractFiter 1, val(XPTxtID.Text)
End If
GoTo x22
End If


Dim newdes As String
Dim ComVal As Double
Dim commissionvalue As Double
Dim RemainValue As Double
Dim RntVal As Double
 
'lineno = 1
  rs.update
If SystemOptions.SysAppAccoutingType = CompeleteAccounting Then
SaveJL LngDevID, lineno, newdes, ComVal, RntVal, RemainValue, commissionvalue, IarType
  
 ''///⁄„Ê·«  «·„‰«œÌ»

 '//////
             If DCboCashType.ListIndex = 5 And (Option1.value = True Or Option2.value = True) Then
                '«·„‘«—Ì⁄

                
                Dim account_codeLegal As String
                Dim account_codeREVENUE_account As String
               ' Dim pstate As Integer
                account_codeLegal = get_project_Account(val(DBCboClientName.BoundText), "legal")
                account_codeREVENUE_account = get_project_Account(val(DBCboClientName.BoundText), "REVENUE_account")
                pstate = val(get_project_Account(val(DBCboClientName.BoundText), "pstate"))
                If SystemOptions.Revenueowed = False Then
GoTo ll
                End If
                
If pstate = 1 Then Option7.value = True: GoTo ll

                If account_codeLegal = "" Or account_codeREVENUE_account = "" Then GoTo ll
       
                RsDev.AddNew
                RsDev("branch_id").value = val(Me.Dcbranch.BoundText)
                RsDev("Double_Entry_Vouchers_ID").value = LngDevID
                RsDev("DEV_ID_Line_No").value = 3
                RsDev("DEV_ID_Line_No1").value = Line3
            
                RsDev("Account_Code").value = account_codeLegal
                RsDev("Value").value = val(Me.XPTxtVal.Text)
                RsDev("Credit_Or_Debit").value = 0
                RsDev("Double_Entry_Vouchers_Description").value = XPMTxtRemarks.Text
                'RsDev("Double_Entry_Vouchers_Description").value = dcproject.BoundText
            
                RsDev("Notes_ID").value = val(XPTxtID.Text)
                RsDev("RecordDate").value = Me.XPDtbTrans.value
                RsDev("RecordDateH").value = ToHijriDate(XPDtbTrans.value)
                RsDev("UserID").value = Me.DCboUserName.BoundText
                RsDev("Account_Interval_ID").value = SystemOptions.SysCurrentAccountIntervalID

                If DCboCashType.ListIndex = 5 Then
                    RsDev("project_id").value = IIf(DBCboClientName.BoundText = "", 0, DBCboClientName.BoundText)
                End If

                RsDev.update
                '«·ÿ—› «·œ«∆‰
                RsDev.AddNew
                lineno = lineno + 1
                RsDev("branch_id").value = val(Me.Dcbranch.BoundText)
                RsDev("Double_Entry_Vouchers_ID").value = LngDevID
                RsDev("DEV_ID_Line_No").value = lineno
                RsDev("DEV_ID_Line_No1").value = Line4
                RsDev("Account_Code").value = account_codeREVENUE_account
                RsDev("Value").value = val(Me.XPTxtVal.Text)
                RsDev("Credit_Or_Debit").value = 1
                RsDev("Double_Entry_Vouchers_Description").value = XPMTxtRemarks.Text
                ' RsDev("Double_Entry_Vouchers_Description").value = dcproject.BoundText
                RsDev("Notes_ID").value = val(XPTxtID.Text)
                RsDev("RecordDate").value = Me.XPDtbTrans.value
                RsDev("RecordDateH").value = ToHijriDate(XPDtbTrans.value)
                RsDev("UserID").value = Me.DCboUserName.BoundText
                RsDev("Account_Interval_ID").value = SystemOptions.SysCurrentAccountIntervalID

                If DCboCashType.ListIndex = 5 Then
                    RsDev("project_id").value = IIf(DBCboClientName.BoundText = "", 0, DBCboClientName.BoundText)
                End If
    
                RsDev.update
            If SystemOptions.DueComm = True Then
            lineno = lineno + 1
      RsDev.AddNew
            RsDev("branch_id").value = val(Me.Dcbranch.BoundText)
            RsDev("Double_Entry_Vouchers_ID").value = LngDevID
            RsDev("DEV_ID_Line_No").value = lineno
            RsDev("DEV_ID_Line_No1").value = Line1
 If SystemOptions.Create2account4Supp = True Then
 
  RentAccount = get_account_code_branch(153, val(Dcbranch.BoundText))
   End If

            RsDev("Account_Code").value = RentAccount
            RsDev("Value").value = (ComVal)
            RsDev("Credit_Or_Debit").value = 0
            RsDev("Double_Entry_Vouchers_Description").value = XPMTxtRemarks.Text & CHR(13) & newdes
            RsDev("Notes_ID").value = val(XPTxtID.Text)
            RsDev("RecordDate").value = Me.XPDtbTrans.value
            RsDev("RecordDateH").value = ToHijriDate(XPDtbTrans.value)
            RsDev("UserID").value = Me.DCboUserName.BoundText
            RsDev("Account_Interval_ID").value = SystemOptions.SysCurrentAccountIntervalID
             
            RsDev.update
 
lineno = lineno + 1
 RsDev.AddNew
        RsDev("branch_id").value = val(Me.Dcbranch.BoundText)
            RsDev("Double_Entry_Vouchers_ID").value = LngDevID
            RsDev("DEV_ID_Line_No").value = lineno
            RsDev("DEV_ID_Line_No1").value = Line1
            RsDev("Account_Code").value = CommissionAcc
            RsDev("Value").value = ComVal
            RsDev("Credit_Or_Debit").value = 1
            RsDev("Double_Entry_Vouchers_Description").value = XPMTxtRemarks.Text & CHR(13) & newdes
            RsDev("Notes_ID").value = val(XPTxtID.Text)
            RsDev("RecordDate").value = Me.XPDtbTrans.value
            RsDev("RecordDateH").value = ToHijriDate(XPDtbTrans.value)
            RsDev("UserID").value = Me.DCboUserName.BoundText
            RsDev("Account_Interval_ID").value = SystemOptions.SysCurrentAccountIntervalID
            RsDev.update
 End If
 '''''''

 ''''''''''''
ll:
            End If
x22:

            LblDevID.Caption = LngDevID
            lbl(33).Caption = SystemOptions.SysCurrentAccountIntervalID
        End If
   If SystemOptions.CreateJLEmpCommissions = True Then
   lineno = lineno + 1
   LngDevID = new_id("DOUBLE_ENTREY_VOUCHERS", "Double_Entry_Vouchers_ID", "")
JlContEmpComm LngDevID, lineno, newdes
End If
        '==========================================================================
            
              If DCboCashType.ListIndex = 5 Then
            saveprojectBillPayment TxtNoteSerial.Text, val(XPTxtVal.Text)
        End If
      
        
        rs.update
                      Save3
            CuurentLogdata
             save_General_cost_center Me.DcCostCenter.BoundText, Me.DcCostCenter.Text, "„ﬁ»Ê÷« ", Me.XPDtbTrans.value
        save_cost_center
          If val(DCboCashType.ListIndex) <> 8 Then
           updateNotesValueAndNobytext val(XPTxtID.Text), Format(XPTxtVal.Text, "###.00")
           Else
         '  updateNotesValueAndNobytext val(XPTxtID.text), Format(XPTxtVal.text, "###.00")
       End If
       
        Cn.CommitTrans
        rs.Resync adAffectCurrent
        BeginTrans = False
        XPTxtCurrent.Caption = rs.AbsolutePosition
        XPTxtCount.Caption = rs.RecordCount

        'Õ›Ÿ «·„” Œ·’« 
      
    
        If DCboCashType.ListIndex = 5 Then
            FillGridWithData val(Me.DBCboClientName.BoundText), TxtNoteSerial.Text
        End If
    
    
    
       'Õ›Ÿ «·«ﬁ”«ÿ ·⁄ﬁÊœ «·«ÌÃ«—

    ''//

        
        If Me.ChkTrans.value = vbUnchecked Then
            Me.CboTrans.ListIndex = -1
            Me.TxtTransSerial.Text = ""
            Me.TxtTransID.Text = ""
        End If
    SendMessage (1)
        
'rs.Resync
        Select Case Me.TxtModFlg.Text

            Case "N"
                Msg = "  „ Õ›Ÿ »Ì«‰«  Â–Â «·⁄„·Ì… " & CHR(13)
                Msg = Msg + "Â·  —€» ›Ì ≈÷«›… »Ì«‰«  √Œ—Ì"

                If MsgBox(Msg, vbYesNo + vbQuestion + vbMsgBoxRight + vbMsgBoxRtlReading + vbDefaultButton2, App.Title) = vbYes Then
                    Cmd_Click (0)
                    Exit Sub
                End If
     
            Case "E"
                MsgBox " „ Õ›Ÿ Â–Â «· ⁄œÌ·« ", vbInformation + vbOKOnly + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
                lbl(46).Caption = showLabel(TxtNoteSerial1, oldtxtNoteSerial1)
           
        End Select
       Me.TxtModFlg.Text = "R"
                Retrive val(XPTxtID.Text)
        '   If Me.DcCostCenter.BoundText <> "" Then
       
        '   End If
       
    
 
        TxtModFlg.Text = "R"
    

    WriteCustomerBalPublic Me.DcboCreditSide.BoundText, Balance, balanceString
    LblLink.Caption = balanceString
    WriteInfo

 
   
 

    TxtModFlg.Text = "R"
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
        MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
        Exit Sub
    End If
    Msg = "⁄›Ê«...ÕœÀ Œÿ√ „« √À‰«¡ Õ›Ÿ Â–Â «·»Ì«‰«  " & CHR(13)
    MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
End Sub
Sub SaveJL(Optional ByRef LngDevID As Long, Optional ByRef lineno As Double, Optional ByRef newdes As String, Optional ByRef ComVal As Double, Optional ByRef RntVal As Double, Optional ByRef RemainValue As Double, Optional ByRef commissionvalue As Double, Optional IarType As Integer)
Dim RsDev As ADODB.Recordset
Dim StrSQL As String
Dim i As Integer
      Line1 = setfoxy_Line
        Line2 = setfoxy_Line
        Line3 = setfoxy_Line
        Line4 = setfoxy_Line
        ClculteVAT

        ' ”ÃÌ· ﬁÌÊœ
        
            LngDevID = new_id("DOUBLE_ENTREY_VOUCHERS", "Double_Entry_Vouchers_ID", "", True)
            Set RsDev = New ADODB.Recordset
        '    RsDev.Open "DOUBLE_ENTREY_VOUCHERS", Cn, adOpenStatic, adLockOptimistic, adCmdTable
                      StrSQL = "SELECT     * from dbo.DOUBLE_ENTREY_VOUCHERS Where (1 = -1)"
   RsDev.Open StrSQL, Cn, adOpenKeyset, adLockOptimistic, adCmdText
         
         ''////////////
         'If DCboCashType.ListIndex = 9 And SystemOptions.NoCreatJLInRentContract = True Then
         If DCboCashType.ListIndex = 9 Then
         ClculteVAT
         
      '      RsDev.AddNew
      '      RsDev("branch_id").value = val(Me.Dcbranch.BoundText)
      '      RsDev("Double_Entry_Vouchers_ID").value = LngDevID
      '      RsDev("DEV_ID_Line_No").value = lineno
      '      RsDev("DEV_ID_Line_No1").value = Line1
      '      RsDev("Account_Code").value = ModAccounts.GetMyAccountCode("TblCustemers", "CusID", val(DBCboClientName.BoundText), "Account_code")
      '      RsDev("Value").value = val(Me.XPTxtVal.Text)
'            RsDev("Credit_Or_Debit").value = 0
      '
      '       newdes = "  ⁄—»Ê‰ ÕÃ“  «·ÊÕœ…   " & DcbUnitType.Text & "  »—ﬁ„   " & DcbUnitNo.Text & "  ··„” √Ã— " & txtrenterName
      '
      '      RsDev("Double_Entry_Vouchers_Description").value = XPMTxtRemarks.Text & CHR(13) & newdes
      '      RsDev("Notes_ID").value = val(XPTxtID.Text)
      '      RsDev("RecordDate").value = Me.XPDtbTrans.value
      '      RsDev("RecordDateH").value = ToHijriDate(XPDtbTrans.value)
      '      RsDev("UserID").value = Me.DCboUserName.BoundText
      '      RsDev("Account_Interval_ID").value = SystemOptions.SysCurrentAccountIntervalID
'
'
'            RsDev.update
            '«·ÿ—› «·œ«∆‰
     
'              RsDev.AddNew
'            lineno = lineno + 1
'            RsDev("branch_id").value = val(Me.Dcbranch.BoundText)
'            RsDev("Double_Entry_Vouchers_ID").value = LngDevID
'            RsDev("DEV_ID_Line_No").value = lineno
'            RsDev("DEV_ID_Line_No1").value = Line2
'
'            RsDev("Account_Code").value = ModAccounts.GetMyAccountCode("TblCustemers", "CusID", val(Txtownerid), "Account_code")
'            RsDev("Value").value = val(Me.XPTxtVal.Text)
'            RsDev("Credit_Or_Debit").value = 1
'            RsDev("Double_Entry_Vouchers_Description").value = XPMTxtRemarks.Text & CHR(13) & newdes
'            RsDev("Notes_ID").value = val(XPTxtID.Text)
'            RsDev("RecordDate").value = Me.XPDtbTrans.value
'            RsDev("RecordDateH").value = ToHijriDate(XPDtbTrans.value)
'            RsDev("UserID").value = Me.DCboUserName.BoundText
'            RsDev("Account_Interval_ID").value = SystemOptions.SysCurrentAccountIntervalID
'            RsDev.update
            End If
            '«·ÿ—› «·„œÌ‰
       ''//////////////////
            RsDev.AddNew
            lineno = lineno + 1
            RsDev("branch_id").value = val(Me.Dcbranch.BoundText)
            RsDev("Double_Entry_Vouchers_ID").value = LngDevID
            RsDev("Aqarid").value = val(DcbIqara.BoundText)
            RsDev("DEV_ID_Line_No").value = lineno
            RsDev("DEV_ID_Line_No1").value = Line1 + 1
            RsDev("Account_Code").value = Me.DcboDebitSide.BoundText
            RsDev("Value").value = val(Me.XPTxtVal.Text) + val(TxtVATValue.Text)
            RsDev("Credit_Or_Debit").value = 0
               If DCboCashType.ListIndex = 9 Then
            
             newdes = "  ⁄—»Ê‰ ÕÃ“  «·⁄ﬁ«—    " & CHR(13) & "  «·⁄ﬁ«—" & DcbIqara.Text & CHR(13) & " «·‰Ê⁄ :" & DcbUnitType.Text & CHR(13) & "  »—ﬁ„   " & DcbUnitNo.Text & CHR(13) & "  ··„” √Ã— " & Me.DBCboClientName.Text
            End If
            RsDev("Double_Entry_Vouchers_Description").value = XPMTxtRemarks.Text & CHR(13) & newdes
            RsDev("Notes_ID").value = val(XPTxtID.Text)
            RsDev("RecordDate").value = Me.XPDtbTrans.value
            RsDev("RecordDateH").value = ToHijriDate(XPDtbTrans.value)
            RsDev("UserID").value = Me.DCboUserName.BoundText
            RsDev("Account_Interval_ID").value = SystemOptions.SysCurrentAccountIntervalID
           ' RsDev("Notes_ID").value = val(XPTxtID.text)

            RsDev.update '1
            '«·ÿ—› «·œ«∆‰
     
            'If SystemOptions.NoCreatJLInRentContract = True And DCboCashType.ListIndex = 8 Then
            If DCboCashType.ListIndex = 8 Then
            
            If IarType <> 0 Then
                RsDev.AddNew
                lineno = lineno + 1
                RsDev("branch_id").value = val(Me.Dcbranch.BoundText)
                RsDev("Double_Entry_Vouchers_ID").value = LngDevID
                RsDev("DEV_ID_Line_No").value = lineno
                RsDev("DEV_ID_Line_No1").value = Line2
                RsDev("Aqarid").value = val(DcbIqara.BoundText)
                If SystemOptions.OpenAccountAqar = False Then
                    RsDev("Account_Code").value = ModAccounts.GetMyAccountCode("TblCustemers", "CusID", val(Txtownerid), "Account_code")
                Else
                    RsDev("Account_Code").value = GetAqarAcountCode(val(DcbIqara.BoundText))
                End If
                RsDev("Value").value = val(Me.XPTxtVal.Text)
                RsDev("Credit_Or_Debit").value = 1
                RsDev("Double_Entry_Vouchers_Description").value = XPMTxtRemarks.Text & CHR(13) & newdes
                RsDev("Notes_ID").value = val(XPTxtID.Text)
                RsDev("RecordDate").value = Me.XPDtbTrans.value
                RsDev("RecordDateH").value = ToHijriDate(XPDtbTrans.value)
                RsDev("UserID").value = Me.DCboUserName.BoundText
                RsDev("Account_Interval_ID").value = SystemOptions.SysCurrentAccountIntervalID
                RsDev.update
            ''////////////////
                If val(Me.TxtVATValue.Text) <> 0 Then
                      RsDev.AddNew
                      lineno = lineno + 1
                    RsDev("branch_id").value = val(Me.Dcbranch.BoundText)
                    RsDev("Double_Entry_Vouchers_ID").value = LngDevID
                    RsDev("DEV_ID_Line_No").value = lineno
                    RsDev("DEV_ID_Line_No1").value = Line2
                    RsDev("Aqarid").value = val(DcbIqara.BoundText)
                    RsDev("Account_Code").value = ModAccounts.GetMyAccountCode("TblCustemers", "CusID", val(Txtownerid), "Account_VAT")
                    RsDev("Value").value = val(Me.TxtVATValue.Text)
                    RsDev("Credit_Or_Debit").value = 1
                    RsDev("Double_Entry_Vouchers_Description").value = XPMTxtRemarks.Text & CHR(13) & newdes
                    RsDev("Notes_ID").value = val(XPTxtID.Text)
                    RsDev("RecordDate").value = Me.XPDtbTrans.value
                    RsDev("RecordDateH").value = ToHijriDate(XPDtbTrans.value)
                    RsDev("UserID").value = Me.DCboUserName.BoundText
                    RsDev("Account_Interval_ID").value = SystemOptions.SysCurrentAccountIntervalID
                    RsDev.update
                End If
            Else
           JLContract LngDevID, lineno, newdes
            End If
           Else
           If Trim(Me.DcboCreditSide.BoundText) <> "" Then
           LngDevID = LngDevID + 1
            RsDev.AddNew
            lineno = lineno + 1
            RsDev("branch_id").value = val(Me.Dcbranch.BoundText)
            RsDev("Aqarid").value = val(DcbIqara.BoundText)
            RsDev("Double_Entry_Vouchers_ID").value = LngDevID
            RsDev("DEV_ID_Line_No").value = lineno
            RsDev("DEV_ID_Line_No1").value = Line2
         
        
        If IarType = 1 And DCboCashType.ListIndex = 9 Then '⁄—»Ê‰
                            If SystemOptions.OpenAccountAqar = False Then
                            RsDev("Account_Code").value = Me.DcboCreditSide.BoundText '  ModAccounts.GetMyAccountCode("TblCustemers", "CusID", val(Txtownerid), "Account_code")
                            Else
                            RsDev("Account_Code").value = GetAqarAcountCode(val(DcbIqara.BoundText))
                            End If
'            RsDev("Account_Code").value = Me.DcboCreditSide.BoundText
        Else
            RsDev("Account_Code").value = Me.DcboCreditSide.BoundText
           End If
            
            
            RsDev("Value").value = val(Me.XPTxtVal.Text)
            RsDev("Aqarid").value = val(DcbIqara.BoundText)
            RsDev("Credit_Or_Debit").value = 1
             If DCboCashType.ListIndex = 9 Then
            
             newdes = "  ⁄—»Ê‰ ÕÃ“  «·ÊÕœ…   " & " «·⁄ﬁ«— " & DcbIqara.Text & "  ‰Ê⁄ «·ÊÕœ… " & DcbUnitType.Text & "  »—ﬁ„   " & DcbUnitNo.Text & "  ··„” √Ã— " & txtrenterName
             newdes = "  ⁄—»Ê‰ ÕÃ“  «·⁄ﬁ«—    " & CHR(13) & "  «·⁄ﬁ«—" & DcbIqara.Text & CHR(13) & " «·‰Ê⁄ :" & DcbUnitType.Text & CHR(13) & "  »—ﬁ„   " & DcbUnitNo.Text & CHR(13) & "  ··„” √Ã— " & Me.DBCboClientName.Text
             
            End If
            RsDev("Double_Entry_Vouchers_Description").value = XPMTxtRemarks.Text & CHR(13) & newdes
            RsDev("Notes_ID").value = val(XPTxtID.Text)
            RsDev("RecordDate").value = Me.XPDtbTrans.value
            RsDev("RecordDateH").value = ToHijriDate(XPDtbTrans.value)
            RsDev("UserID").value = Me.DCboUserName.BoundText
            RsDev("Account_Interval_ID").value = SystemOptions.SysCurrentAccountIntervalID
            RsDev("CarId").value = IIf(Me.DCCar.BoundText = "", Null, (Me.DCCar.BoundText))
    
            RsDev.update '2
            End If
End If

 
 lineno = lineno + 1
'salim here 03 10 2019 If Me.DCboCashType.ListIndex = 8 And Rd(1).value = True Then
  If Me.DCboCashType.ListIndex = 8 And IarType = 1 Then
  RentAccount = ""
RentAccount = GetMyAccountCode("TblCustemers", "CusID", val(Txtownerid), "Account_code1")
               
  ComVal = 0
  RntVal = 0
  Dim mVATValue1Com As Double
  Dim mVATValue2Com As Double
  If 1 = 1 Then
'If SystemOptions.DueComm = True Then

  If True = True Then
  With Me.Grid3
 For i = 1 To .rows - 1
        If .cell(flexcpChecked, i, .ColIndex("Select")) = flexChecked And val(.TextMatrix(i, .ColIndex("RentValuePayed"))) + val(.TextMatrix(i, .ColIndex("RentArbon"))) <> 0 Then
RntVal = RntVal + val(.TextMatrix(i, .ColIndex("RentValuePayed"))) + val(.TextMatrix(i, .ColIndex("RentArbon"))) ' + val(.TextMatrix(i, .ColIndex("VATPayed")))
            mVATValue1Com = mVATValue1Com + val(.TextMatrix(i, .ColIndex("VATValue1Com")))
            mVATValue2Com = mVATValue2Com + val(.TextMatrix(i, .ColIndex("VATValue2Com")))
         End If
       Next i
 End With
 
 End If
 mVATValue1Com = mVATValue1Com + mVATValue2Com
 commissionvalue = Round(RntVal * val(TxtKickbacks) / 100, 2)
'commissionvalue = Round(commissionvalue, 2)
'RemainValue = val(Me.XPTxtVal.Text) - commissionvalue
'If SystemOptions.DueComm = True Then
If True = True Then
  With Me.Grid3
 For i = 1 To .rows - 1
        If .cell(flexcpChecked, i, .ColIndex("Select")) = flexChecked And val(.TextMatrix(i, .ColIndex("CommissionsPayed"))) + val(.TextMatrix(i, .ColIndex("CommissionsArbon"))) <> 0 Then
ComVal = ComVal + val(.TextMatrix(i, .ColIndex("CommissionsPayed"))) + val(.TextMatrix(i, .ColIndex("CommissionsArbon")))   '”⁄Ì
         End If
       Next i
 End With
 
 End If
 'RemainValue = val(Me.XPTxtVal.Text) - ComVal
'salim here
 RemainValue = RntVal - ComVal
RemainValue = Round(val(RntVal) - commissionvalue, 2)

 Else
' RemainValue = Round(val(Me.XPTxtVal.Text) - commissionvalue, 2)
RemainValue = Round(val(RntVal) - commissionvalue, 2)

  
 
 
 End If
 newdes = "   ﬁ»÷ „‰ «„·«ﬂ «·€Ì— ·   " & DcbUnitType.Text & "  »—ﬁ„   " & DcbUnitNo.Text & " ··⁄ﬁ«— " & DcbIqara.Text
 'If SystemOptions.NoCreatJLInRentContract = False And DCboCashType.ListIndex = 8 Then
 If DCboCashType.ListIndex = 8 Then
  RsDev.AddNew
            RsDev("branch_id").value = val(Me.Dcbranch.BoundText)
            RsDev("Aqarid").value = val(DcbIqara.BoundText)
            RsDev("Double_Entry_Vouchers_ID").value = LngDevID
            RsDev("DEV_ID_Line_No").value = lineno
            RsDev("DEV_ID_Line_No1").value = Line1


         If SystemOptions.OpenAccountAqar = False Then
             RentAccount = ModAccounts.GetMyAccountCode("TblCustemers", "CusID", val(Txtownerid), "Account_code")
             
                                     If SystemOptions.Create2account4Supp = True Then
                        
                         RentAccount = ModAccounts.GetMyAccountCode("TblCustemers", "CusID", val(Txtownerid), "Account_code1")
                          End If
   
   
             Else
              RentAccount = GetAqarAcountCode(val(DcbIqara.BoundText))
             End If
             
            RsDev("Account_Code").value = RentAccount
            'Wael   „ «Ìﬁ«›Â ·«‰‰« ‰ﬁ’‰« «·⁄„Ê·… „‰ «·ﬁÌœ RsDev("Value").value = RntVal ' RemainValue ' (RntVal)
            
             If RemainValue = 0 Then
                RemainValue = val(Me.XPTxtVal.Text) + val(TxtVATValue.Text)
             End If
            RsDev("Value").value = RemainValue ' (RntVal)
            RsDev("Credit_Or_Debit").value = 0
             
            RsDev("Double_Entry_Vouchers_Description").value = XPMTxtRemarks.Text & CHR(13) & newdes & "«ÌÃ«— "
            RsDev("Notes_ID").value = val(XPTxtID.Text)
            RsDev("RecordDate").value = Me.XPDtbTrans.value
            RsDev("RecordDateH").value = ToHijriDate(XPDtbTrans.value)
            RsDev("UserID").value = Me.DCboUserName.BoundText
            RsDev("Account_Interval_ID").value = SystemOptions.SysCurrentAccountIntervalID
             
            RsDev.update '3
 
lineno = lineno + 1
 RsDev.AddNew
        RsDev("branch_id").value = val(Me.Dcbranch.BoundText)
            RsDev("Double_Entry_Vouchers_ID").value = LngDevID
            RsDev("Aqarid").value = val(DcbIqara.BoundText)
            RsDev("DEV_ID_Line_No").value = lineno
            RsDev("DEV_ID_Line_No1").value = Line1
         '   RsDev("Account_Code").value = ModAccounts.GetMyAccountCode("TblCustemers", "CusID", val(Txtownerid), "Account_code")
            
                   If SystemOptions.OpenAccountAqar = False Then
                RsDev("Account_Code").value = ModAccounts.GetMyAccountCode("TblCustemers", "CusID", val(Txtownerid), "Account_code")
             Else
                 RsDev("Account_Code").value = GetAqarAcountCode(val(DcbIqara.BoundText))
             End If
             
            'Wael   „ «Ìﬁ«›Â ·«‰‰« ‰ﬁ’‰« «·⁄„Ê·… „‰ «·ﬁÌœ RsDev("Value").value = RntVal ' RemainValue ' (RntVal)
                         If RemainValue = 0 Then
                RemainValue = val(Me.XPTxtVal.Text) + val(TxtVATValue.Text)
             End If

            RsDev("Value").value = RemainValue ' (RntVal)
            RsDev("Credit_Or_Debit").value = 1
            
           '  newdes = "   ﬁ»÷ „‰ «„·«ﬂ «·€Ì— ·   " & DcbUnitType.text & "  »—ﬁ„   " & DcbUnitNo.text & "  ··„” √Ã— " & txtrenterName
            
            RsDev("Double_Entry_Vouchers_Description").value = XPMTxtRemarks.Text & CHR(13) & newdes & "«ÌÃ«— "
            RsDev("Notes_ID").value = val(XPTxtID.Text)
            RsDev("RecordDate").value = Me.XPDtbTrans.value
            RsDev("RecordDateH").value = ToHijriDate(XPDtbTrans.value)
            RsDev("UserID").value = Me.DCboUserName.BoundText
            RsDev("Account_Interval_ID").value = SystemOptions.SysCurrentAccountIntervalID
             
            RsDev.update '4
            End If
            
    ''////«·”⁄Ì «·„” Õﬁ
  'If SystemOptions.DueComm = True And ComVal > 0 And SystemOptions.NoCreatJLInRentContract = False Then
  If SystemOptions.DueComm = True And ComVal > 0 Then
  lineno = lineno + 1
      RsDev.AddNew
            RsDev("branch_id").value = val(Me.Dcbranch.BoundText)
            RsDev("Double_Entry_Vouchers_ID").value = LngDevID
            RsDev("DEV_ID_Line_No").value = lineno
            RsDev("DEV_ID_Line_No1").value = Line1
            RsDev("Aqarid").value = val(DcbIqara.BoundText)
 If SystemOptions.Create2account4Supp = True Then
 
  RentAccount = get_account_code_branch(153, val(Dcbranch.BoundText))
   End If

            RsDev("Account_Code").value = RentAccount
            RsDev("Value").value = (ComVal)
            RsDev("Credit_Or_Debit").value = 0
            RsDev("Double_Entry_Vouchers_Description").value = XPMTxtRemarks.Text & CHR(13) & newdes & " ”⁄Ì  "
            RsDev("Notes_ID").value = val(XPTxtID.Text)
            RsDev("RecordDate").value = Me.XPDtbTrans.value
            RsDev("RecordDateH").value = ToHijriDate(XPDtbTrans.value)
            RsDev("UserID").value = Me.DCboUserName.BoundText
            RsDev("Account_Interval_ID").value = SystemOptions.SysCurrentAccountIntervalID
             
            RsDev.update
 
lineno = lineno + 1
 RsDev.AddNew
        RsDev("branch_id").value = val(Me.Dcbranch.BoundText)
            RsDev("Double_Entry_Vouchers_ID").value = LngDevID
            RsDev("DEV_ID_Line_No").value = lineno
            RsDev("DEV_ID_Line_No1").value = Line1
            RsDev("Account_Code").value = CommissionAcc
            RsDev("Value").value = ComVal
            RsDev("Credit_Or_Debit").value = 1
            RsDev("Double_Entry_Vouchers_Description").value = XPMTxtRemarks.Text & CHR(13) & newdes & " ”⁄Ì "
            RsDev("Notes_ID").value = val(XPTxtID.Text)
            RsDev("RecordDate").value = Me.XPDtbTrans.value
            RsDev("RecordDateH").value = ToHijriDate(XPDtbTrans.value)
            RsDev("UserID").value = Me.DCboUserName.BoundText
            RsDev("Account_Interval_ID").value = SystemOptions.SysCurrentAccountIntervalID
            RsDev("Aqarid").value = val(DcbIqara.BoundText)
            RsDev.update
            
            
            
            

            
            
'            StrTempAccountCode = GetMyAccountCode("TblCustemers", "CusID", CLng(ownerid), "Account_Code1")
'            If ModAccounts.AddNewDev(Line1, lineno, StrTempAccountCode, mVATValue1Com, 0, StrTempDes & "       «·ﬁÌ„… «·„÷«›… ··⁄„Ê·… Õ”«» «·„«·ﬂ  ", general_noteid, , , , NoteDate, UserID, , , , , , , , , , , , , , , , , , BranchID) = False Then
'                GoTo ErrTrap
'
'            End If
'            lineno = lineno + 1
'            Dim account As String
'            PercentgValueAddedAccount_Transec XPDtbTrans.value, 21, 1, account
'            If ModAccounts.AddNewDev(Line1, lineno, account, mVATValue1Com, 1, StrTempDes & "       «·ﬁÌ„… «·„÷«›… ··⁄„Ê·… Õ”«» «·„«·ﬂ ", general_noteid, , , , NoteDat, UserID, , , , , , , , , , , , , , , , , , BranchID) = False Then
'                GoTo ErrTrap
'
'            End If
        End If
                    
                    
 End If
'//////////////
' ﬁÌœ «·⁄„Ê·…
'a
'Â‰«  „ «Ìﬁ«› Õ”«» ⁄„Ê·… «·„«·ﬂ
'  Ê·Ìœ Ê„«Ãœ  „ «⁄«œÂ «·ﬂÊ œ
'If commissionvalue > 0 And SystemOptions.NoCreatJLInRentContract = False Then
If commissionvalue > 0 Then
If 1 = 1 And SystemOptions.CommissionDue = True Then
lineno = lineno + 1
 RsDev.AddNew
        RsDev("branch_id").value = val(Me.Dcbranch.BoundText)
            RsDev("Double_Entry_Vouchers_ID").value = LngDevID
            RsDev("DEV_ID_Line_No").value = lineno
            RsDev("DEV_ID_Line_No1").value = Line1
            RsDev("Aqarid").value = val(DcbIqara.BoundText)
            If SystemOptions.CommissionOnOwner = True And SystemOptions.CommissionDue = False Then
          
                         If SystemOptions.OpenAccountAqar = False Then
                        RsDev("Account_Code").value = ModAccounts.GetMyAccountCode("TblCustemers", "CusID", val(Txtownerid), "Account_code")
                        Else
                         RsDev("Account_Code").value = GetAqarAcountCode(val(DcbIqara.BoundText))
                        End If
             Else
              RsDev("Account_Code").value = CommissionAccDue
            End If
            
            RsDev("Value").value = commissionvalue
            RsDev("Credit_Or_Debit").value = 0
           newdes = "   ﬁ»÷  ⁄„Ê·Â „‰ «„·«ﬂ «·€Ì— ·‹‹‹‹‹   " & DcbUnitType.Text & "  »—ﬁ„   " & DcbUnitNo.Text & "  ··„” √Ã— " & txtrenterName
            RsDev("Double_Entry_Vouchers_Description").value = XPMTxtRemarks.Text & CHR(13) & newdes
            RsDev("Notes_ID").value = val(XPTxtID.Text)
            RsDev("RecordDate").value = Me.XPDtbTrans.value
            RsDev("RecordDateH").value = ToHijriDate(XPDtbTrans.value)
            RsDev("UserID").value = Me.DCboUserName.BoundText
            RsDev("Account_Interval_ID").value = SystemOptions.SysCurrentAccountIntervalID
             
            RsDev.update '5
  



lineno = lineno + 1
 RsDev.AddNew
        RsDev("branch_id").value = val(Me.Dcbranch.BoundText)
            RsDev("Double_Entry_Vouchers_ID").value = LngDevID
            RsDev("DEV_ID_Line_No").value = lineno
            RsDev("DEV_ID_Line_No1").value = Line1
            RsDev("Aqarid").value = val(DcbIqara.BoundText)
            Dim dummyCommissionAcc As String
            dummyCommissionAcc = get_account_code_branch(207, my_branch)
            
            If dummyCommissionAcc <> "" And dummyCommissionAcc <> "NO account" Then
            CommissionAcc = dummyCommissionAcc 'Õ”«» «·⁄„Ê·Â „‰›’· ⁄‰ Õ”«»  «·”⁄Ì
            End If
            
            RsDev("Account_Code").value = CommissionAcc
            
            
            RsDev("Value").value = commissionvalue
            RsDev("Credit_Or_Debit").value = 1
           newdes = "   ﬁ»÷ ⁄„Ê·Â „‰ «„·«ﬂ «·€Ì— ·   " & DcbUnitType.Text & "  »—ﬁ„   " & DcbUnitNo.Text & "  ··„” √Ã— " & txtrenterName
            RsDev("Double_Entry_Vouchers_Description").value = XPMTxtRemarks.Text & CHR(13) & newdes
            RsDev("Notes_ID").value = val(XPTxtID.Text)
            RsDev("RecordDate").value = Me.XPDtbTrans.value
            RsDev("RecordDateH").value = ToHijriDate(XPDtbTrans.value)
            RsDev("UserID").value = Me.DCboUserName.BoundText
            RsDev("Account_Interval_ID").value = SystemOptions.SysCurrentAccountIntervalID
             
            RsDev.update '6
            
            
                        
 If mVATValue1Com > 0 Then
           lineno = lineno + 1
           
           Dim account As String
            PercentgValueAddedAccount_Transec XPDtbTrans.value, 21, 1, account
            RsDev.AddNew
        RsDev("branch_id").value = val(Me.Dcbranch.BoundText)
            RsDev("Double_Entry_Vouchers_ID").value = LngDevID
            RsDev("DEV_ID_Line_No").value = lineno
            RsDev("DEV_ID_Line_No1").value = Line1
'            RsDev("Account_Code").value = account
                               If SystemOptions.OpenAccountAqar = False Then
                RsDev("Account_Code").value = ModAccounts.GetMyAccountCode("TblCustemers", "CusID", val(Txtownerid), "Account_code")
             Else
                 RsDev("Account_Code").value = GetAqarAcountCode(val(DcbIqara.BoundText))
             End If

            RsDev("Value").value = mVATValue1Com
            RsDev("Credit_Or_Debit").value = 0
            RsDev("Double_Entry_Vouchers_Description").value = XPMTxtRemarks.Text & CHR(13) & newdes & " ÷—Ì»… «·⁄„Ê·… "
            RsDev("Notes_ID").value = val(XPTxtID.Text)
            RsDev("RecordDate").value = Me.XPDtbTrans.value
            RsDev("RecordDateH").value = ToHijriDate(XPDtbTrans.value)
            RsDev("UserID").value = Me.DCboUserName.BoundText
            RsDev("Account_Interval_ID").value = SystemOptions.SysCurrentAccountIntervalID
            RsDev("Aqarid").value = val(DcbIqara.BoundText)
            RsDev.update
            
            
7
           lineno = lineno + 1
           
           
            PercentgValueAddedAccount_Transec XPDtbTrans.value, 21, 1, account
             Dim ownerid As Double
            GetIqarCode , , DcbIqara.BoundText, , ownerid
            If SystemOptions.Create2account4Supp = True Then
                account = GetMyAccountCode("TblCustemers", "CusID", CLng(ownerid), "Account_Code1")
            End If

            RsDev.AddNew
        RsDev("branch_id").value = val(Me.Dcbranch.BoundText)
            RsDev("Double_Entry_Vouchers_ID").value = LngDevID
            RsDev("DEV_ID_Line_No").value = lineno
            RsDev("DEV_ID_Line_No1").value = Line1
            RsDev("Account_Code").value = account
            RsDev("Value").value = mVATValue1Com
            RsDev("Credit_Or_Debit").value = 1
            RsDev("Double_Entry_Vouchers_Description").value = XPMTxtRemarks.Text & CHR(13) & newdes & " ÷—Ì»… «·⁄„Ê·… "
            RsDev("Notes_ID").value = val(XPTxtID.Text)
            RsDev("RecordDate").value = Me.XPDtbTrans.value
            RsDev("RecordDateH").value = ToHijriDate(XPDtbTrans.value)
            RsDev("UserID").value = Me.DCboUserName.BoundText
            RsDev("Account_Interval_ID").value = SystemOptions.SysCurrentAccountIntervalID
            RsDev("Aqarid").value = val(DcbIqara.BoundText)
            RsDev.update
'mVATValue1Com
End If
'salimhere**************************************************************************  Œ›Ì÷ «·«” Õﬁ«ﬁ »«·⁄„Ê·Â Ê«·ﬁÌ„Â «·„÷«›… ⁄·ÌÂ«
 Dim OwnerAccount As String
  Dim OwnerDueAccount As String
Dim PERCNTAGE As Double
Dim RentAccountX As String
 OwnerAccount = ModAccounts.GetMyAccountCode("TblCustemers", "CusID", val(Txtownerid), "Account_code")
OwnerDueAccount = GetMyAccountCode("TblCustemers", "CusID", val(Txtownerid), "Account_code1")

PercentgValueAddedAccount_Transec XPDtbTrans.value, 51, 1, RentAccountX, PERCNTAGE
       
'If PERCNTAGE = 0 Then
'       PercentgValueAddedAccount_Transec XPDtbTrans.value, 21, 1, RentAccountX, PERCNTAGE
'End If
If SystemOptions.DueComm = True And commissionvalue > 0 Then
lineno = lineno + 1
 RsDev.AddNew
        RsDev("branch_id").value = val(Me.Dcbranch.BoundText)
            RsDev("Double_Entry_Vouchers_ID").value = LngDevID
            RsDev("DEV_ID_Line_No").value = lineno
            RsDev("DEV_ID_Line_No1").value = Line1
            RsDev("Aqarid").value = val(DcbIqara.BoundText)
            RsDev("Account_Code").value = OwnerAccount '«·„«·ﬂ ;
            
            
            RsDev("Value").value = commissionvalue * (1 + PERCNTAGE / 100) '1.05
            

       
            RsDev("Credit_Or_Debit").value = 0
           newdes = "   ﬁ»÷ ⁄„Ê·Â „‰ «„·«ﬂ «·€Ì— ·   " & DcbUnitType.Text & "  »—ﬁ„   " & DcbUnitNo.Text & "  ··„” √Ã— " & txtrenterName '& "«·⁄„Ê·Â „⁄ ﬁ „"
            RsDev("Double_Entry_Vouchers_Description").value = XPMTxtRemarks.Text & CHR(13) & newdes
            RsDev("Notes_ID").value = val(XPTxtID.Text)
            RsDev("RecordDate").value = Me.XPDtbTrans.value
            RsDev("RecordDateH").value = ToHijriDate(XPDtbTrans.value)
            RsDev("UserID").value = Me.DCboUserName.BoundText
            RsDev("Account_Interval_ID").value = SystemOptions.SysCurrentAccountIntervalID
             
            RsDev.update '7
            
            
            lineno = lineno + 1
 RsDev.AddNew
        RsDev("branch_id").value = val(Me.Dcbranch.BoundText)
            RsDev("Double_Entry_Vouchers_ID").value = LngDevID
            RsDev("DEV_ID_Line_No").value = lineno
            RsDev("DEV_ID_Line_No1").value = Line1
            RsDev("Aqarid").value = val(DcbIqara.BoundText)
                        dummyCommissionAcc = get_account_code_branch(207, my_branch)
            
            If dummyCommissionAcc <> "" And dummyCommissionAcc <> "NO account" Then
            CommissionAcc = dummyCommissionAcc 'Õ”«» «·⁄„Ê·Â „‰›’· ⁄‰ Õ”«»  «·”⁄Ì
            End If
       If SystemOptions.CommissionOnOwner = True Then
       ' «·⁄„Ê·Â ··„«·ﬂ ÌÕ”» ÌŒ›÷ «·⁄„Ê·Â „‰ «” Õﬁ«ﬁ «·„«·ﬂ
      RsDev("Account_Code").value = OwnerDueAccount '  „” Õﬁ  «·„«·ﬂ  ;
      RsDev("Value").value = commissionvalue * (1 + PERCNTAGE / 100) '1.05
       Else
            RsDev("Account_Code").value = OwnerDueAccount ' CommissionAcc '  „” Õﬁ  «·„«·ﬂ  ;
     RsDev("Value").value = commissionvalue * (1 + PERCNTAGE / 100) '1.05
       End If
            
            
            RsDev("Credit_Or_Debit").value = 1
           newdes = "   ﬁ»÷ ⁄„Ê·Â „‰ «„·«ﬂ «·€Ì— ·   zzz " & DcbUnitType.Text & "  »—ﬁ„   " & DcbUnitNo.Text & "  ··„” √Ã— " & txtrenterName & "«·⁄„Ê·Â  "
            RsDev("Double_Entry_Vouchers_Description").value = XPMTxtRemarks.Text & CHR(13) & newdes
            RsDev("Notes_ID").value = val(XPTxtID.Text)
            RsDev("RecordDate").value = Me.XPDtbTrans.value
            RsDev("RecordDateH").value = ToHijriDate(XPDtbTrans.value)
            RsDev("UserID").value = Me.DCboUserName.BoundText
            RsDev("Account_Interval_ID").value = SystemOptions.SysCurrentAccountIntervalID
             
            RsDev.update '8

If SystemOptions.CommissionOnOwner = False Then '«·⁄„Ê·Â ··„«·ﬂ  ·«ÌÕ”» ÷—Ì»Â
If commissionvalue * PERCNTAGE / 100 > 0 Then
lineno = lineno + 1

 RsDev.AddNew
        RsDev("branch_id").value = val(Me.Dcbranch.BoundText)
            RsDev("Double_Entry_Vouchers_ID").value = LngDevID
            RsDev("DEV_ID_Line_No").value = lineno
            RsDev("DEV_ID_Line_No1").value = Line1
            GetValueAddedAccount XPDtbTrans.value, , RentAccount, 1, 21
            PercentgValueAddedAccount_Transec XPDtbTrans.value, 21, 1, RentAccount
            RsDev("Account_Code").value = RentAccount
            RsDev("Value").value = commissionvalue * PERCNTAGE / 100 ' 0.05
            RsDev("Credit_Or_Debit").value = 1
           newdes = "  zzz  ﬁ»÷ „‰ «„·«ﬂ «·€Ì— ·   " & DcbUnitType.Text & "  »—ﬁ„   " & DcbUnitNo.Text & "  ··„” √Ã— " & txtrenterName & "  ﬁ „÷«›… "
            RsDev("Double_Entry_Vouchers_Description").value = XPMTxtRemarks.Text & CHR(13) & newdes
            RsDev("Notes_ID").value = val(XPTxtID.Text)
            RsDev("RecordDate").value = Me.XPDtbTrans.value
            RsDev("RecordDateH").value = ToHijriDate(XPDtbTrans.value)
            RsDev("UserID").value = Me.DCboUserName.BoundText
            RsDev("Account_Interval_ID").value = SystemOptions.SysCurrentAccountIntervalID
            RsDev("Aqarid").value = val(DcbIqara.BoundText)
            RsDev.update

End If
End If
'salimhere**************************************************************************


End If










End If






    End If
    ''////////////////
    Dim CommissionVAT As Double
    'Â‰«  „ «Ìﬁ«› Õ”«»  «·ﬁÌ„Â «·„÷«›… ⁄·Ì ⁄„Ê·… «·„«·ﬂ
   CommissionVAT = val(commissionvalue) * 5 / 100
   'If SystemOptions.NoCreatJLInRentContract = True And DCboCashType.ListIndex = 8 And IarType = 1 And CommissionVAT > 0 Then
    If False = True And DCboCashType.ListIndex = 8 And IarType = 1 And CommissionVAT > 0 Then
    If commissionvalue > 0 Then
lineno = lineno + 1
 RsDev.AddNew
        RsDev("branch_id").value = val(Me.Dcbranch.BoundText)
            RsDev("Double_Entry_Vouchers_ID").value = LngDevID
            RsDev("DEV_ID_Line_No").value = lineno
            RsDev("DEV_ID_Line_No1").value = Line1
            RsDev("Account_Code").value = CommissionAcc
            RsDev("Value").value = CommissionVAT
            RsDev("Aqarid").value = val(DcbIqara.BoundText)
            RsDev("Credit_Or_Debit").value = 0
            newdes = "   ﬁ»÷ „‰ «„·«ﬂ «·€Ì— ·   " & DcbUnitType.Text & "  »—ﬁ„   " & DcbUnitNo.Text & "  ··„” √Ã— " & txtrenterName
            RsDev("Double_Entry_Vouchers_Description").value = XPMTxtRemarks.Text & CHR(13) & newdes
            RsDev("Notes_ID").value = val(XPTxtID.Text)
            RsDev("RecordDate").value = Me.XPDtbTrans.value
            RsDev("RecordDateH").value = ToHijriDate(XPDtbTrans.value)
            RsDev("UserID").value = Me.DCboUserName.BoundText
            RsDev("Account_Interval_ID").value = SystemOptions.SysCurrentAccountIntervalID
             
            RsDev.update

lineno = lineno + 1
 RsDev.AddNew
        RsDev("branch_id").value = val(Me.Dcbranch.BoundText)
            RsDev("Double_Entry_Vouchers_ID").value = LngDevID
            RsDev("DEV_ID_Line_No").value = lineno
            RsDev("DEV_ID_Line_No1").value = Line1
            GetValueAddedAccount XPDtbTrans.value, , RentAccount, 1, 21
            PercentgValueAddedAccount_Transec XPDtbTrans.value, 8, 1, RentAccount
            RsDev("Account_Code").value = RentAccount
            RsDev("Value").value = CommissionVAT
            RsDev("Credit_Or_Debit").value = 1
           newdes = "   ﬁ»÷ „‰ «„·«ﬂ «·€Ì— ·   " & DcbUnitType.Text & "  »—ﬁ„   " & DcbUnitNo.Text & "  ··„” √Ã— " & txtrenterName
            RsDev("Double_Entry_Vouchers_Description").value = XPMTxtRemarks.Text & CHR(13) & newdes
            RsDev("Notes_ID").value = val(XPTxtID.Text)
            RsDev("RecordDate").value = Me.XPDtbTrans.value
            RsDev("RecordDateH").value = ToHijriDate(XPDtbTrans.value)
            RsDev("UserID").value = Me.DCboUserName.BoundText
            RsDev("Account_Interval_ID").value = SystemOptions.SysCurrentAccountIntervalID
            RsDev("Aqarid").value = val(DcbIqara.BoundText)
            RsDev.update
End If

    End If
    ''///////////////////////
    
    If DCboCashType.ListIndex = 9 Then
    If val(TxtVATValue.Text) > 0 Then

lineno = lineno + 1
 RsDev.AddNew
        RsDev("branch_id").value = val(Me.Dcbranch.BoundText)
            RsDev("Double_Entry_Vouchers_ID").value = LngDevID
            RsDev("DEV_ID_Line_No").value = lineno
            RsDev("DEV_ID_Line_No1").value = Line1
           ' GetValueAddedAccount XPDtbTrans.value, , RentAccount, 1, 21
            PercentgValueAddedAccount_Transec XPDtbTrans.value, 8, 1, RentAccount
            RsDev("Account_Code").value = RentAccount
            RsDev("Value").value = val(TxtVATValue.Text)
            RsDev("Credit_Or_Debit").value = 1
           newdes = "   ﬁ»÷ „‰ «„·«ﬂ «·€Ì— ·   " & DcbUnitType.Text & "  »—ﬁ„   " & DcbUnitNo.Text & "  ··„” √Ã— " & txtrenterName
            RsDev("Double_Entry_Vouchers_Description").value = XPMTxtRemarks.Text & CHR(13) & newdes
            RsDev("Notes_ID").value = val(XPTxtID.Text)
            RsDev("RecordDate").value = Me.XPDtbTrans.value
            RsDev("RecordDateH").value = ToHijriDate(XPDtbTrans.value)
            RsDev("UserID").value = Me.DCboUserName.BoundText
            RsDev("Account_Interval_ID").value = SystemOptions.SysCurrentAccountIntervalID
             RsDev("Aqarid").value = val(DcbIqara.BoundText)
            RsDev.update
End If

    End If
 ''///////
 '    If Me.DCboCashType.ListIndex = 8 And Rd(1).value = False Then
     If Me.DCboCashType.ListIndex = 8 Then '⁄ﬁœ
     
 '//„Ì«Â
  ComVal = 0
    'If SystemOptions.DueWater = True And SystemOptions.NoCreatJLInRentContract = False Then
    If SystemOptions.DueWater = True Then
     With Me.Grid3
 For i = 1 To .rows - 1
        If .cell(flexcpChecked, i, .ColIndex("Select")) = flexChecked And val(.TextMatrix(i, .ColIndex("WaterPayed"))) + val(.TextMatrix(i, .ColIndex("WaterArbon"))) <> 0 Then
ComVal = ComVal + val(.TextMatrix(i, .ColIndex("WaterPayed"))) + val(.TextMatrix(i, .ColIndex("WaterArbon")))
         End If
       Next i
 End With

 If ComVal > 0 Then
 lineno = lineno + 1
      RsDev.AddNew
            RsDev("branch_id").value = val(Me.Dcbranch.BoundText)
            RsDev("Double_Entry_Vouchers_ID").value = LngDevID
            RsDev("DEV_ID_Line_No").value = lineno
            RsDev("DEV_ID_Line_No1").value = Line1
            If Rd(1).value = False Then '«„·«ﬂÌ
            
                 RentAccount = get_account_code_branch(154, val(Dcbranch.BoundText))
            
            Else
            
                  RentAccount = get_account_code_branch(123, val(Dcbranch.BoundText))
            
                      If SystemOptions.ServicesOnOwner = True Then
                               If SystemOptions.Create2account4Supp = True Then
                                  RentAccount = GetMyAccountCode("TblCustemers", "CusID", CLng(Txtownerid.Text), "Account_Code1")
                             End If
            
                        End If

            End If
            
            
            RsDev("Account_Code").value = RentAccount
            RsDev("Value").value = (ComVal)
            RsDev("Credit_Or_Debit").value = 0
            RsDev("Double_Entry_Vouchers_Description").value = XPMTxtRemarks.Text & CHR(13) & newdes & "   „Ì«Â   "
            RsDev("Notes_ID").value = val(XPTxtID.Text)
            RsDev("RecordDate").value = Me.XPDtbTrans.value
            RsDev("RecordDateH").value = ToHijriDate(XPDtbTrans.value)
            RsDev("UserID").value = Me.DCboUserName.BoundText
            RsDev("Account_Interval_ID").value = SystemOptions.SysCurrentAccountIntervalID
            RsDev("Aqarid").value = val(DcbIqara.BoundText)
            RsDev.update
 
lineno = lineno + 1
            RsDev.AddNew
            RentAccount = get_account_code_branch(83, val(Dcbranch.BoundText))
            RsDev("branch_id").value = val(Me.Dcbranch.BoundText)
            RsDev("Double_Entry_Vouchers_ID").value = LngDevID
            RsDev("DEV_ID_Line_No").value = lineno
            RsDev("DEV_ID_Line_No1").value = Line1
            
               If Rd(1).value = False Then '«„·«ﬂÌ
            
                 RentAccount = get_account_code_branch(83, val(Dcbranch.BoundText))
            
            Else
            
                        If SystemOptions.ServicesOnOwner = True Then
                                   
                                  RentAccount = GetMyAccountCode("TblCustemers", "CusID", CLng(Txtownerid.Text), "Account_Code")
                       End If
                    

              End If
            
             RsDev("Account_Code").value = RentAccount
  
            RsDev("Value").value = ComVal
            RsDev("Credit_Or_Debit").value = 1
            
            RsDev("Double_Entry_Vouchers_Description").value = XPMTxtRemarks.Text & CHR(13) & newdes & "   „Ì«Â   "
            RsDev("Notes_ID").value = val(XPTxtID.Text)
            RsDev("RecordDate").value = Me.XPDtbTrans.value
            RsDev("RecordDateH").value = ToHijriDate(XPDtbTrans.value)
            RsDev("UserID").value = Me.DCboUserName.BoundText
            RsDev("Account_Interval_ID").value = SystemOptions.SysCurrentAccountIntervalID
            RsDev("Aqarid").value = val(DcbIqara.BoundText)
            RsDev.update
 End If
 End If
 ''ﬂÂ»—«¡
 ComVal = 0
     'If SystemOptions.DueElectr = True And SystemOptions.NoCreatJLInRentContract = False Then
     If SystemOptions.DueElectr = True Then
     With Me.Grid3
 For i = 1 To .rows - 1
        If .cell(flexcpChecked, i, .ColIndex("Select")) = flexChecked And val(.TextMatrix(i, .ColIndex("ElectricPayed"))) + val(.TextMatrix(i, .ColIndex("ElectricArbon"))) <> 0 Then
ComVal = ComVal + val(.TextMatrix(i, .ColIndex("ElectricPayed"))) + val(.TextMatrix(i, .ColIndex("ElectricArbon")))
         End If
       Next i
 End With
 Print

 If ComVal > 0 Then
 lineno = lineno + 1
      RsDev.AddNew
            RsDev("branch_id").value = val(Me.Dcbranch.BoundText)
            RsDev("Double_Entry_Vouchers_ID").value = LngDevID
            RsDev("DEV_ID_Line_No").value = lineno
            RsDev("DEV_ID_Line_No1").value = Line1
            RentAccount = get_account_code_branch(155, val(Dcbranch.BoundText))
            
            If Rd(1).value = False Then '«„·«ﬂÌ
            
                 RentAccount = get_account_code_branch(155, val(Dcbranch.BoundText))
            
            Else
            
                  RentAccount = get_account_code_branch(123, val(Dcbranch.BoundText))
            
                      If SystemOptions.ServicesOnOwner = True Then
                               If SystemOptions.Create2account4Supp = True Then
                                  RentAccount = GetMyAccountCode("TblCustemers", "CusID", CLng(Txtownerid.Text), "Account_Code1")
                             End If
            
                        End If

            End If
            
            
            RsDev("Account_Code").value = RentAccount
            RsDev("Value").value = (ComVal)
            RsDev("Credit_Or_Debit").value = 0
            RsDev("Double_Entry_Vouchers_Description").value = XPMTxtRemarks.Text & CHR(13) & newdes & "   ﬂÂ—»«¡   "
            RsDev("Notes_ID").value = val(XPTxtID.Text)
            RsDev("RecordDate").value = Me.XPDtbTrans.value
            RsDev("RecordDateH").value = ToHijriDate(XPDtbTrans.value)
            RsDev("UserID").value = Me.DCboUserName.BoundText
            RsDev("Account_Interval_ID").value = SystemOptions.SysCurrentAccountIntervalID
            RsDev("Aqarid").value = val(DcbIqara.BoundText)
            RsDev.update
 
lineno = lineno + 1
 RsDev.AddNew
 RentAccount = get_account_code_branch(84, val(Dcbranch.BoundText))
           RsDev("branch_id").value = val(Me.Dcbranch.BoundText)
            RsDev("Double_Entry_Vouchers_ID").value = LngDevID
            RsDev("DEV_ID_Line_No").value = lineno
            RsDev("DEV_ID_Line_No1").value = Line1
            
                           If Rd(1).value = False Then '«„·«ﬂÌ
            
                 RentAccount = get_account_code_branch(84, val(Dcbranch.BoundText))
            
            Else
            
                        If SystemOptions.ServicesOnOwner = True Then
                                   
                                  RentAccount = GetMyAccountCode("TblCustemers", "CusID", CLng(Txtownerid.Text), "Account_Code")
                       End If
                    

              End If
              RsDev("Account_Code").value = RentAccount
            RsDev("Value").value = ComVal
            RsDev("Credit_Or_Debit").value = 1
            RsDev("Double_Entry_Vouchers_Description").value = XPMTxtRemarks.Text & CHR(13) & newdes & "   ﬂÂ—»«¡   "
            RsDev("Notes_ID").value = val(XPTxtID.Text)
            RsDev("RecordDate").value = Me.XPDtbTrans.value
            RsDev("RecordDateH").value = ToHijriDate(XPDtbTrans.value)
            RsDev("UserID").value = Me.DCboUserName.BoundText
            RsDev("Account_Interval_ID").value = SystemOptions.SysCurrentAccountIntervalID
            RsDev("Aqarid").value = val(DcbIqara.BoundText)
            RsDev.update
 End If
 End If
 ''//////Œœ„« 
  ComVal = 0
  
      'If SystemOptions.DueService = True And SystemOptions.NoCreatJLInRentContract = False Then
      If SystemOptions.DueService = True Then
     With Me.Grid3
 For i = 1 To .rows - 1
        If .cell(flexcpChecked, i, .ColIndex("Select")) = flexChecked And val(.TextMatrix(i, .ColIndex("TelandNetPayed"))) + val(.TextMatrix(i, .ColIndex("ServiceArbon"))) <> 0 Then
ComVal = ComVal + val(.TextMatrix(i, .ColIndex("TelandNetPayed"))) + val(.TextMatrix(i, .ColIndex("ServiceArbon")))
         End If
       Next i
 End With
 
 If ComVal > 0 Then
      RsDev.AddNew
      lineno = lineno + 1
            RsDev("branch_id").value = val(Me.Dcbranch.BoundText)
            RsDev("Double_Entry_Vouchers_ID").value = LngDevID
            RsDev("DEV_ID_Line_No").value = lineno
            RsDev("DEV_ID_Line_No1").value = Line1
            RentAccount = get_account_code_branch(156, val(Dcbranch.BoundText))
                        If Rd(1).value = False Then '«„·«ﬂÌ
            
                 RentAccount = get_account_code_branch(156, val(Dcbranch.BoundText))
            
            Else
            
                  RentAccount = get_account_code_branch(123, val(Dcbranch.BoundText))
            
                      If SystemOptions.ServicesOnOwner = True Then
                               If SystemOptions.Create2account4Supp = True Then
                                  RentAccount = GetMyAccountCode("TblCustemers", "CusID", CLng(Txtownerid.Text), "Account_Code1")
                             End If
            
                        End If

            End If
            
            RsDev("Account_Code").value = RentAccount
            RsDev("Value").value = (ComVal)
            RsDev("Credit_Or_Debit").value = 0
            RsDev("Double_Entry_Vouchers_Description").value = XPMTxtRemarks.Text & CHR(13) & newdes & "   Œœ„«    "
            RsDev("Notes_ID").value = val(XPTxtID.Text)
            RsDev("RecordDate").value = Me.XPDtbTrans.value
            RsDev("RecordDateH").value = ToHijriDate(XPDtbTrans.value)
            RsDev("UserID").value = Me.DCboUserName.BoundText
            RsDev("Account_Interval_ID").value = SystemOptions.SysCurrentAccountIntervalID
            RsDev("Aqarid").value = val(DcbIqara.BoundText)
            RsDev.update
 
lineno = lineno + 1
 RsDev.AddNew
 RentAccount = get_account_code_branch(85, val(Dcbranch.BoundText))
           RsDev("branch_id").value = val(Me.Dcbranch.BoundText)
            RsDev("Double_Entry_Vouchers_ID").value = LngDevID
            RsDev("DEV_ID_Line_No").value = lineno
            RsDev("DEV_ID_Line_No1").value = Line1
                           If Rd(1).value = False Then '«„·«ﬂÌ
            
                 RentAccount = get_account_code_branch(85, val(Dcbranch.BoundText))
            
            Else
            
                        If SystemOptions.ServicesOnOwner = True Then
                                   
                                  RentAccount = GetMyAccountCode("TblCustemers", "CusID", CLng(Txtownerid.Text), "Account_Code")
                       End If
                    

              End If
              
            RsDev("Account_Code").value = RentAccount
            RsDev("Value").value = ComVal
            RsDev("Credit_Or_Debit").value = 1
            RsDev("Double_Entry_Vouchers_Description").value = XPMTxtRemarks.Text & CHR(13) & newdes & "   Œœ„«    "
            RsDev("Notes_ID").value = val(XPTxtID.Text)
            RsDev("RecordDate").value = Me.XPDtbTrans.value
            RsDev("RecordDateH").value = ToHijriDate(XPDtbTrans.value)
            RsDev("UserID").value = Me.DCboUserName.BoundText
            RsDev("Account_Interval_ID").value = SystemOptions.SysCurrentAccountIntervalID
            RsDev("Aqarid").value = val(DcbIqara.BoundText)
            RsDev.update
 End If
 End If
 ''”⁄Ì
 ComVal = 0 'dd
 'salim here sbe3y
   'If SystemOptions.DueComm = True And SystemOptions.NoCreatJLInRentContract = False And IarType = 0 Then ' ”⁄” «„·«ﬂÌ
   If SystemOptions.DueComm = True And IarType = 0 Then ' ”⁄” «„·«ﬂÌ
     With Me.Grid3
 For i = 1 To .rows - 1
        If .cell(flexcpChecked, i, .ColIndex("Select")) = flexChecked And val(.TextMatrix(i, .ColIndex("CommissionsPayed"))) + val(.TextMatrix(i, .ColIndex("CommissionsArbon"))) <> 0 Then
ComVal = ComVal + val(.TextMatrix(i, .ColIndex("CommissionsPayed"))) + val(.TextMatrix(i, .ColIndex("CommissionsArbon")))
         End If
       Next i
 End With
 If ComVal > 0 Then
      RsDev.AddNew
      lineno = lineno + 1
            RsDev("branch_id").value = val(Me.Dcbranch.BoundText)
            RsDev("Double_Entry_Vouchers_ID").value = LngDevID
            RsDev("DEV_ID_Line_No").value = lineno
            RsDev("DEV_ID_Line_No1").value = Line1
            RsDev("Aqarid").value = val(DcbIqara.BoundText)
            RentAccount = get_account_code_branch(153, val(Dcbranch.BoundText))
                        If Rd(1).value = False Then '«„·«ﬂÌ
            
                 RentAccount = get_account_code_branch(153, val(Dcbranch.BoundText))
            
            Else
            
                  RentAccount = get_account_code_branch(123, val(Dcbranch.BoundText))
            
                      If SystemOptions.ServicesOnOwner = True Then
                               If SystemOptions.Create2account4Supp = True Then
                                  RentAccount = GetMyAccountCode("TblCustemers", "CusID", CLng(Txtownerid.Text), "Account_Code1")
                             End If
            
                        End If

            End If
            
            RsDev("Account_Code").value = RentAccount
            RsDev("Value").value = (ComVal)
            RsDev("Credit_Or_Debit").value = 0
            RsDev("Double_Entry_Vouchers_Description").value = XPMTxtRemarks.Text & CHR(13) & newdes & "  ø ”⁄Ì   "
            RsDev("Notes_ID").value = val(XPTxtID.Text)
            RsDev("RecordDate").value = Me.XPDtbTrans.value
            RsDev("RecordDateH").value = ToHijriDate(XPDtbTrans.value)
            RsDev("UserID").value = Me.DCboUserName.BoundText
            RsDev("Account_Interval_ID").value = SystemOptions.SysCurrentAccountIntervalID
             
            RsDev.update
 
lineno = lineno + 1
 RsDev.AddNew
 CommissionAcc = get_account_code_branch(81, my_branch)
 
        RsDev("branch_id").value = val(Me.Dcbranch.BoundText)
            RsDev("Double_Entry_Vouchers_ID").value = LngDevID
            RsDev("DEV_ID_Line_No").value = lineno
            RsDev("DEV_ID_Line_No1").value = Line1
            
            RsDev("Account_Code").value = CommissionAcc
            RsDev("Value").value = ComVal
            RsDev("Credit_Or_Debit").value = 1
            RsDev("Double_Entry_Vouchers_Description").value = XPMTxtRemarks.Text & CHR(13) & newdes & "  ø ”⁄Ì   "
            RsDev("Notes_ID").value = val(XPTxtID.Text)
            RsDev("RecordDate").value = Me.XPDtbTrans.value
            RsDev("RecordDateH").value = ToHijriDate(XPDtbTrans.value)
            RsDev("UserID").value = Me.DCboUserName.BoundText
            RsDev("Account_Interval_ID").value = SystemOptions.SysCurrentAccountIntervalID
            RsDev("Aqarid").value = val(DcbIqara.BoundText)
            RsDev.update
 End If
 End If
 End If


 ComVal = 0
 'salim here sbe3y  «·‹«„Ì‰ ··„«·ﬂ
   'If SystemOptions.InsuranceOnOwner = True And SystemOptions.NoCreatJLInRentContract = False And IarType = 1 Then
   If SystemOptions.InsuranceOnOwner = True And IarType = 1 Then
     With Me.Grid3
 For i = 1 To .rows - 1
        If .cell(flexcpChecked, i, .ColIndex("Select")) = flexChecked And val(.TextMatrix(i, .ColIndex("InsurancePayed"))) + val(.TextMatrix(i, .ColIndex("InsuranceArbon"))) <> 0 Then
ComVal = ComVal + val(.TextMatrix(i, .ColIndex("InsurancePayed"))) + val(.TextMatrix(i, .ColIndex("InsuranceArbon")))
         End If
       Next i
 End With
 If ComVal > 0 Then
lineno = lineno + 1
RsDev.AddNew
            RsDev("branch_id").value = val(Me.Dcbranch.BoundText)
            RsDev("Aqarid").value = val(DcbIqara.BoundText)
            RsDev("Double_Entry_Vouchers_ID").value = LngDevID
            RsDev("DEV_ID_Line_No").value = lineno
            RsDev("DEV_ID_Line_No1").value = Line1


         If SystemOptions.OpenAccountAqar = False Then
             RentAccount = ModAccounts.GetMyAccountCode("TblCustemers", "CusID", val(Txtownerid), "Account_code")
             
                                     If SystemOptions.Create2account4Supp = True Then
                        
                         RentAccount = ModAccounts.GetMyAccountCode("TblCustemers", "CusID", val(Txtownerid), "Account_code1")
                          End If
   
   
             Else
              RentAccount = GetAqarAcountCode(val(DcbIqara.BoundText))
             End If
             
            RsDev("Account_Code").value = RentAccount
            RsDev("Value").value = ComVal  ' RemainValue ' (RntVal)
            RsDev("Credit_Or_Debit").value = 0
             
            RsDev("Double_Entry_Vouchers_Description").value = XPMTxtRemarks.Text & CHR(13) & newdes & " √„Ì‰ „” —œ "
            RsDev("Notes_ID").value = val(XPTxtID.Text)
            RsDev("RecordDate").value = Me.XPDtbTrans.value
            RsDev("RecordDateH").value = ToHijriDate(XPDtbTrans.value)
            RsDev("UserID").value = Me.DCboUserName.BoundText
            RsDev("Account_Interval_ID").value = SystemOptions.SysCurrentAccountIntervalID
             
            RsDev.update '3
 
lineno = lineno + 1
 RsDev.AddNew
        RsDev("branch_id").value = val(Me.Dcbranch.BoundText)
            RsDev("Double_Entry_Vouchers_ID").value = LngDevID
            RsDev("Aqarid").value = val(DcbIqara.BoundText)
            RsDev("DEV_ID_Line_No").value = lineno
            RsDev("DEV_ID_Line_No1").value = Line1
         '   RsDev("Account_Code").value = ModAccounts.GetMyAccountCode("TblCustemers", "CusID", val(Txtownerid), "Account_code")
            
                   If SystemOptions.OpenAccountAqar = False Then
                RsDev("Account_Code").value = ModAccounts.GetMyAccountCode("TblCustemers", "CusID", val(Txtownerid), "Account_code")
             Else
                 RsDev("Account_Code").value = GetAqarAcountCode(val(DcbIqara.BoundText))
             End If
             
            RsDev("Value").value = ComVal 'RemainValue
            RsDev("Credit_Or_Debit").value = 1
            
           '  newdes = "   ﬁ»÷ „‰ «„·«ﬂ «·€Ì— ·   " & DcbUnitType.text & "  »—ﬁ„   " & DcbUnitNo.text & "  ··„” √Ã— " & txtrenterName
            
            RsDev("Double_Entry_Vouchers_Description").value = XPMTxtRemarks.Text & CHR(13) & newdes & " √„Ì‰ „” —œ "
            RsDev("Notes_ID").value = val(XPTxtID.Text)
            RsDev("RecordDate").value = Me.XPDtbTrans.value
            RsDev("RecordDateH").value = ToHijriDate(XPDtbTrans.value)
            RsDev("UserID").value = Me.DCboUserName.BoundText
            RsDev("Account_Interval_ID").value = SystemOptions.SysCurrentAccountIntervalID
             
            RsDev.update '4
            End If
 



End If

End Sub
Sub Save3()
    save2
        If DCboCashType.ListIndex = 8 Then
               FillGridWithDataContract txtContractNo.Text, val(XPTxtID.Text)
        End If
           GetUonitStatus
 SaveUoitInformation val(DCboCashType.ListIndex)
End Sub
Sub JlContEmpComm(Optional LngDevID As Long, Optional ByRef lineno As Double, Optional newdes As String)
 Dim empID2 As Double
 Dim EmpID22 As String
 Dim ComVal As Double
 Dim i As Integer
 Dim StrSQL As String
 'GetCommInformation EmpID2, ComVal
 Dim sql As String
Dim rs2 As ADODB.Recordset
Set rs2 = New ADODB.Recordset
Dim RsDev As ADODB.Recordset
   Set RsDev = New ADODB.Recordset
   StrSQL = "SELECT     * from dbo.DOUBLE_ENTREY_VOUCHERS Where (1 = -1)"
   RsDev.Open StrSQL, Cn, adOpenKeyset, adLockOptimistic, adCmdText
'sql = "Select * from TblAqarCommissions where NoteID=" & val(XPTxtID.Text) & " "
sql = " SELECT     SUM(Amount) AS Amount, EmpID"
sql = sql & " From dbo.TblAqarCommissions"
sql = sql & " Where (NoteID = " & val(XPTxtID.Text) & ")"
sql = sql & " GROUP BY EmpID"

rs2.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
If rs2.RecordCount > 0 Then
For i = 1 To rs2.RecordCount
empID2 = IIf(IsNull(rs2("EmpID").value), 0, rs2("EmpID").value)
ComVal = IIf(IsNull(rs2("Amount").value), 0, rs2("Amount").value)
 If ComVal > 0 And empID2 <> 0 Then
 EmpID22 = empID2
 If LngDevID = 0 Then
       LngDevID = new_id("DOUBLE_ENTREY_VOUCHERS", "Double_Entry_Vouchers_ID", "")
 End If
 Line1 = Line1 + 1
            RsDev.AddNew
            lineno = lineno + 1
            RsDev("branch_id").value = val(Me.Dcbranch.BoundText)
            RsDev("Double_Entry_Vouchers_ID").value = LngDevID
            RsDev("DEV_ID_Line_No").value = lineno
            RsDev("DEV_ID_Line_No1").value = Line1
            RentAccount = get_account_code_branch(161, val(Dcbranch.BoundText))
            RsDev("Account_Code").value = RentAccount
            RsDev("Value").value = (ComVal)
            RsDev("Credit_Or_Debit").value = 0
            RsDev("Double_Entry_Vouchers_Description").value = XPMTxtRemarks.Text & CHR(13) & newdes & " Õ”«» ⁄„Ê·«  «·„‰«œÌ»  "
            RsDev("Notes_ID").value = val(XPTxtID.Text)
            RsDev("RecordDate").value = Me.XPDtbTrans.value
            RsDev("RecordDateH").value = ToHijriDate(XPDtbTrans.value)
            RsDev("UserID").value = val(Me.DCboUserName.BoundText)
            RsDev.update
            
            Line1 = Line1 + 1
            lineno = lineno + 1
            RsDev.AddNew
            CommissionAcc = get_EMPLOYEE_Account(EmpID22, "Account_code1")
            RsDev("branch_id").value = val(Me.Dcbranch.BoundText)
            RsDev("Double_Entry_Vouchers_ID").value = LngDevID
            RsDev("DEV_ID_Line_No").value = lineno + 1
            RsDev("DEV_ID_Line_No1").value = Line1
            RsDev("Account_Code").value = CommissionAcc
            RsDev("Value").value = ComVal
            RsDev("Credit_Or_Debit").value = 1
            RsDev("Double_Entry_Vouchers_Description").value = XPMTxtRemarks.Text & CHR(13) & newdes
            RsDev("Notes_ID").value = val(XPTxtID.Text)
            RsDev("RecordDate").value = Me.XPDtbTrans.value
            RsDev("RecordDateH").value = ToHijriDate(XPDtbTrans.value)
            RsDev("UserID").value = Me.DCboUserName.BoundText
            RsDev("Account_Interval_ID").value = SystemOptions.SysCurrentAccountIntervalID
            RsDev.update
 End If
 rs2.MoveNext
 Next i
 End If
End Sub
'125070063
Sub JLContract(Optional LngDevID As Long, Optional ByRef lineno As Double, Optional newdes As String)
Dim i As Integer
Dim StrSQL As String
Dim ComVal As Double
Dim RsDev As ADODB.Recordset
   Set RsDev = New ADODB.Recordset
   StrSQL = "SELECT     * from dbo.DOUBLE_ENTREY_VOUCHERS Where (1 = -1)"
   RsDev.Open StrSQL, Cn, adOpenKeyset, adLockOptimistic, adCmdText
     With Me.Grid3
 For i = 1 To .rows - 1
        If .cell(flexcpChecked, i, .ColIndex("Select")) = flexChecked And val(.TextMatrix(i, .ColIndex("RentValuePayed"))) <> 0 Then
ComVal = ComVal + val(.TextMatrix(i, .ColIndex("RentValuePayed")))
         End If
       Next i
 End With

 If ComVal > 0 Then
 lineno = lineno + 1
 Line1 = Line1 + 1
      RsDev.AddNew
            RsDev("branch_id").value = val(Me.Dcbranch.BoundText)
            RsDev("Double_Entry_Vouchers_ID").value = LngDevID
            RsDev("DEV_ID_Line_No").value = lineno
            RsDev("DEV_ID_Line_No1").value = Line1
            RsDev("Aqarid").value = val(DcbIqara.BoundText)
            'RentAccount = get_account_code_branch(86, val(dcBranch.BoundText))
            RentAccount = ModAccounts.GetMyAccountCode("TblCustemers", "CusID", val(Txtownerid), "Account_code")
            RsDev("Account_Code").value = RentAccount
            RsDev("Value").value = (ComVal)
            RsDev("Credit_Or_Debit").value = 1
            RsDev("Double_Entry_Vouchers_Description").value = XPMTxtRemarks.Text & CHR(13) & newdes
            RsDev("Notes_ID").value = val(XPTxtID.Text)
            RsDev("RecordDate").value = Me.XPDtbTrans.value
            RsDev("RecordDateH").value = ToHijriDate(XPDtbTrans.value)
            RsDev("UserID").value = Me.DCboUserName.BoundText
            RsDev("Account_Interval_ID").value = SystemOptions.SysCurrentAccountIntervalID
            RsDev.update
 
 End If
 ''///////////
 ComVal = 0
      With Me.Grid3
 For i = 1 To .rows - 1
        If .cell(flexcpChecked, i, .ColIndex("Select")) = flexChecked And val(.TextMatrix(i, .ColIndex("WaterPayed"))) <> 0 Then
ComVal = ComVal + val(.TextMatrix(i, .ColIndex("WaterPayed")))
         End If
       Next i
 End With

 If ComVal > 0 Then
 lineno = lineno + 1
      RsDev.AddNew
            RsDev("branch_id").value = val(Me.Dcbranch.BoundText)
            RsDev("Aqarid").value = val(DcbIqara.BoundText)
            RsDev("Double_Entry_Vouchers_ID").value = LngDevID
            RsDev("DEV_ID_Line_No").value = lineno
            RsDev("DEV_ID_Line_No1").value = Line1
            RentAccount = get_account_code_branch(83, val(Dcbranch.BoundText))
            RsDev("Account_Code").value = RentAccount
            RsDev("Value").value = (ComVal)
            RsDev("Credit_Or_Debit").value = 1
            RsDev("Double_Entry_Vouchers_Description").value = XPMTxtRemarks.Text & CHR(13) & newdes
            RsDev("Notes_ID").value = val(XPTxtID.Text)
            RsDev("RecordDate").value = Me.XPDtbTrans.value
            RsDev("RecordDateH").value = ToHijriDate(XPDtbTrans.value)
            RsDev("UserID").value = Me.DCboUserName.BoundText
            RsDev("Account_Interval_ID").value = SystemOptions.SysCurrentAccountIntervalID
            RsDev.update
 
 End If
 ''/////////////////
  ComVal = 0
       With Me.Grid3
 For i = 1 To .rows - 1
        If .cell(flexcpChecked, i, .ColIndex("Select")) = flexChecked And val(.TextMatrix(i, .ColIndex("ElectricPayed"))) <> 0 Then
ComVal = ComVal + val(.TextMatrix(i, .ColIndex("ElectricPayed")))
         End If
       Next i
 End With

 If ComVal > 0 Then
 lineno = lineno + 1
      RsDev.AddNew
            RsDev("branch_id").value = val(Me.Dcbranch.BoundText)
            RsDev("Double_Entry_Vouchers_ID").value = LngDevID
            RsDev("DEV_ID_Line_No").value = lineno
            RsDev("DEV_ID_Line_No1").value = Line1
            RentAccount = get_account_code_branch(84, val(Dcbranch.BoundText))
            RsDev("Account_Code").value = RentAccount
            RsDev("Value").value = (ComVal)
            RsDev("Credit_Or_Debit").value = 1
            RsDev("Double_Entry_Vouchers_Description").value = XPMTxtRemarks.Text & CHR(13) & newdes
            RsDev("Notes_ID").value = val(XPTxtID.Text)
            RsDev("RecordDate").value = Me.XPDtbTrans.value
            RsDev("RecordDateH").value = ToHijriDate(XPDtbTrans.value)
            RsDev("UserID").value = Me.DCboUserName.BoundText
            RsDev("Account_Interval_ID").value = SystemOptions.SysCurrentAccountIntervalID
            RsDev("Aqarid").value = val(DcbIqara.BoundText)
            RsDev.update
 
 End If
 
  ''/////////////////
   ComVal = 0
       With Me.Grid3
 For i = 1 To .rows - 1
        If .cell(flexcpChecked, i, .ColIndex("Select")) = flexChecked And val(.TextMatrix(i, .ColIndex("TelandNetPayed"))) <> 0 Then
ComVal = ComVal + val(.TextMatrix(i, .ColIndex("TelandNetPayed")))
         End If
       Next i
 End With

 If ComVal > 0 Then
 lineno = lineno + 1
      RsDev.AddNew
            RsDev("branch_id").value = val(Me.Dcbranch.BoundText)
            RsDev("Double_Entry_Vouchers_ID").value = LngDevID
            RsDev("DEV_ID_Line_No").value = lineno
            RsDev("DEV_ID_Line_No1").value = Line1
            RentAccount = get_account_code_branch(85, val(Dcbranch.BoundText))
            RsDev("Account_Code").value = RentAccount
            RsDev("Value").value = (ComVal)
            RsDev("Credit_Or_Debit").value = 1
            RsDev("Double_Entry_Vouchers_Description").value = XPMTxtRemarks.Text & CHR(13) & newdes
            RsDev("Notes_ID").value = val(XPTxtID.Text)
            RsDev("RecordDate").value = Me.XPDtbTrans.value
            RsDev("RecordDateH").value = ToHijriDate(XPDtbTrans.value)
            RsDev("UserID").value = Me.DCboUserName.BoundText
            RsDev("Account_Interval_ID").value = SystemOptions.SysCurrentAccountIntervalID
            RsDev("Aqarid").value = val(DcbIqara.BoundText)
            RsDev.update
 
 End If
 ''//////////
  ComVal = 0
        With Me.Grid3
 For i = 1 To .rows - 1
        If .cell(flexcpChecked, i, .ColIndex("Select")) = flexChecked And val(.TextMatrix(i, .ColIndex("CommissionsPayed"))) <> 0 Then
ComVal = ComVal + val(.TextMatrix(i, .ColIndex("CommissionsPayed")))
         End If
       Next i
 End With

 If ComVal > 0 Then
 lineno = lineno + 1
      RsDev.AddNew
            RsDev("branch_id").value = val(Me.Dcbranch.BoundText)
            RsDev("Double_Entry_Vouchers_ID").value = LngDevID
            RsDev("DEV_ID_Line_No").value = lineno
            RsDev("DEV_ID_Line_No1").value = Line1
            RentAccount = get_account_code_branch(81, val(Dcbranch.BoundText))
            RsDev("Account_Code").value = RentAccount
            RsDev("Value").value = (ComVal)
            RsDev("Credit_Or_Debit").value = 1
            RsDev("Double_Entry_Vouchers_Description").value = XPMTxtRemarks.Text & CHR(13) & newdes
            RsDev("Notes_ID").value = val(XPTxtID.Text)
            RsDev("RecordDate").value = Me.XPDtbTrans.value
            RsDev("RecordDateH").value = ToHijriDate(XPDtbTrans.value)
            RsDev("UserID").value = Me.DCboUserName.BoundText
            RsDev("Account_Interval_ID").value = SystemOptions.SysCurrentAccountIntervalID
            RsDev("Aqarid").value = val(DcbIqara.BoundText)
            RsDev.update
 
 End If
   ComVal = val(TxtVATValue.Text)

 If ComVal > 0 Then
 lineno = lineno + 1
      RsDev.AddNew
            RsDev("branch_id").value = val(Me.Dcbranch.BoundText)
            RsDev("Double_Entry_Vouchers_ID").value = LngDevID
            RsDev("DEV_ID_Line_No").value = lineno
            RsDev("DEV_ID_Line_No1").value = Line1
            RsDev("Aqarid").value = val(DcbIqara.BoundText)
            GetValueAddedAccount XPDtbTrans.value, , RentAccount, 1, 21
            'RentAccount = get_account_code_branch(145, val(Dcbranch.BoundText))
            PercentgValueAddedAccount_Transec XPDtbTrans.value, 8, 1, RentAccount
            RsDev("Account_Code").value = RentAccount
            RsDev("Value").value = (ComVal)
            RsDev("Credit_Or_Debit").value = 1
            RsDev("Double_Entry_Vouchers_Description").value = XPMTxtRemarks.Text & CHR(13) & newdes
            RsDev("Notes_ID").value = val(XPTxtID.Text)
            RsDev("RecordDate").value = Me.XPDtbTrans.value
            RsDev("RecordDateH").value = ToHijriDate(XPDtbTrans.value)
            RsDev("UserID").value = Me.DCboUserName.BoundText
            RsDev("Account_Interval_ID").value = SystemOptions.SysCurrentAccountIntervalID
            RsDev.update
 
 End If
End Sub


Sub save2()
Dim RsDetails1 As ADODB.Recordset
Dim StrSQL As String
Dim i As Integer
        If DCboCashType.ListIndex = 8 Then
           
 saveContractInstallments val(Me.XPTxtID), XPDtbTrans.value, Txt_DateHigri.value, val(XPTxtVal.Text), val(TxtContNo.Text)
        End If
    ''///
       Set RsDetails1 = New ADODB.Recordset
         StrSQL = "SELECT     *  from  TblNotesSales Where (1 = -1)"
   RsDetails1.Open StrSQL, Cn, adOpenKeyset, adLockOptimistic, adCmdText
      ' RsDetails1.Open "TblCardAuthorizationReformDetails", Cn, adOpenKeyset, adLockOptimistic, adCmdTable
    If val(DCboCashType.ListIndex) = 8 Then
           
      
If VSFlexGrid2.rows > 1 Then
                ' fg2.Rows = fg2.Rows - 1
     With VSFlexGrid2
   
       For i = .FixedRows To .rows - 1
       
              If .TextMatrix(i, .ColIndex("empname")) <> "" Then
           RsDetails1.AddNew
           RsDetails1("Type").value = 0
          RsDetails1("NoteID").value = val(XPTxtID.Text)
          RsDetails1("valu").value = val(.TextMatrix(i, .ColIndex("values")))
            RsDetails1("rate").value = val(.TextMatrix(i, .ColIndex("rate")))
            RsDetails1("ValueAmount").value = val(.TextMatrix(i, .ColIndex("ValueAmount")))
             RsDetails1("EmpID").value = val(.TextMatrix(i, .ColIndex("id")))
             RsDetails1("idd").value = val(.TextMatrix(i, .ColIndex("idd")))
             RsDetails1("GroupID").value = val(.TextMatrix(i, .ColIndex("groupid")))
         RsDetails1.update
     
       End If
           Next i
        
    End With
     
    End If
   End If
    Dim sql As String
    ''\\\«·⁄—»Ê‰
     
   If val(DCboCashType.ListIndex) = 9 Then
                                    
          
               Set RsDetails1 = New ADODB.Recordset
         StrSQL = "SELECT     *  from  TblAqrEarnest Where (1 = -1)"
   RsDetails1.Open StrSQL, Cn, adOpenKeyset, adLockOptimistic, adCmdText
   
           RsDetails1.AddNew
           RsDetails1("CoustomerName").value = IIf(txtrenterName.Text = "", "", txtrenterName.Text)
          RsDetails1("Telephone").value = IIf(TxtTelphone.Text = "", "", TxtTelphone.Text)
          RsDetails1("RecordDate").value = XPDtbTrans.value
            RsDetails1("RecordDateH").value = Txt_DateHigri.value
            
               RsDetails1("UnitNo").value = IIf(DcbUnitNo.BoundText = "", 0, DcbUnitNo.BoundText)
             RsDetails1("Earnest").value = IIf(XPTxtVal.Text = "", 0, XPTxtVal.Text)
             '''\\\
             
             If cbointervaltype.ListIndex = 0 Then
              RsDetails1("ValidityDate").value = DateAdd("d", val(TxtInterval), XPDtbTrans.value)
              ElseIf cbointervaltype.ListIndex = 1 Then
              RsDetails1("ValidityDate").value = DateAdd("M", val(TxtInterval), XPDtbTrans.value)
              
            ElseIf cbointervaltype.ListIndex = 2 Then
              RsDetails1("ValidityDate").value = DateAdd("YYYY", val(TxtInterval), XPDtbTrans.value)
             
             End If
            RsDetails1("ValidityDateH").value = ToHijriDate(rs("allowdate").value)
            RsDetails1("Earnest").value = IIf(XPTxtVal.Text = "", 0, XPTxtVal.Text)
            RsDetails1("NoteID").value = IIf(XPTxtID.Text = "", 0, XPTxtID.Text)
     If CheckStatusEarnest(0).value = vbChecked Then
         RsDetails1("StatusEarnest").value = 1
         
         
         ElseIf CheckStatusEarnest(1).value = vbChecked Then
         RsDetails1("StatusEarnest").value = 2
     ElseIf CheckStatusEarnest(3).value = vbChecked Then
         RsDetails1("StatusEarnest").value = 4
         Else
          RsDetails1("StatusEarnest").value = 0
          End If
     
         RsDetails1.update
     
 End If
If val(DCboCashType.ListIndex) = 9 Then
 If CheckStatusEarnest(0).value = vbChecked Then
        
       If CheckStatusofUnit(val(Me.DcbUnitNo.BoundText)) = True Then
       
        sql = "update TblAqarDetai set   Status =0   where  Id =" & val(Me.DcbUnitNo.BoundText) & " "
        Cn.Execute sql
       End If
         ElseIf CheckStatusEarnest(1).value = vbChecked Then
       
            If CheckStatusofUnit(val(Me.DcbUnitNo.BoundText)) = True Then
        sql = "update TblAqarDetai set   Status =0  where  Id =" & val(Me.DcbUnitNo.BoundText) & " "
        Cn.Execute sql
       End If
        ElseIf CheckStatusEarnest(2).value = vbChecked Then
       
            If CheckStatusofUnit(val(Me.DcbUnitNo.BoundText)) = True Then
        sql = "update TblAqarDetai set   Status =0  where  Id =" & val(Me.DcbUnitNo.BoundText) & " "
        Cn.Execute sql
        
        
       End If
       
       
        ElseIf CheckStatusEarnest(3).value = vbChecked Then
       
            If CheckStatusofUnit(val(Me.DcbUnitNo.BoundText)) = True Then
        sql = "update TblAqarDetai set   Status =0  where  Id =" & val(Me.DcbUnitNo.BoundText) & " "
        Cn.Execute sql
        
        
       End If
         Else
          
           sql = "update TblAqarDetai set   Status =2  where  Id =" & val(Me.DcbUnitNo.BoundText) & " "
           Cn.Execute sql
           
          End If
     End If
    ''//
    
           Set RsDetails1 = New ADODB.Recordset
         StrSQL = "SELECT     *  from  TblNotesSales Where (1 = -1)"
   RsDetails1.Open StrSQL, Cn, adOpenKeyset, adLockOptimistic, adCmdText
      ' RsDetails1.Open "TblCardAuthorizationReformDetails", Cn, adOpenKeyset, adLockOptimistic, adCmdTable
        If val(DCboCashType.ListIndex) <> 8 Then
    
    
If VSFlexGrid1.rows > 1 Then
                ' fg2.Rows = fg2.Rows - 1
     With VSFlexGrid1
     
       For i = .FixedRows To .rows - 1
       
              If .TextMatrix(i, .ColIndex("empname")) <> "" Then
       
           VSFlexGrid2.Clear flexClearScrollable, flexClearEverything
            VSFlexGrid2.rows = 1
       
           RsDetails1.AddNew
           RsDetails1("Type").value = 1
          RsDetails1("NoteID").value = val(XPTxtID.Text)
          RsDetails1("valu").value = val(.TextMatrix(i, .ColIndex("values")))
            RsDetails1("rate").value = val(.TextMatrix(i, .ColIndex("rate")))
             RsDetails1("EmpID").value = val(.TextMatrix(i, .ColIndex("id")))
              RsDetails1("idd").value = val(.TextMatrix(i, .ColIndex("idd")))
               RsDetails1("GroupID").value = val(.TextMatrix(i, .ColIndex("groupid")))
         RsDetails1.update
     
       End If
           Next i
        
    End With
     
    End If
    End If
End Sub
Function saveChequeBoxContents(NoteID As Double)

    Dim i As Integer
    Dim rs2 As New ADODB.Recordset
 
    Dim StrSQL As String

    StrSQL = "Delete  TblChecqueBoxContent  where NoteID =" & NoteID
    Cn.Execute StrSQL, , adExecuteNoRecords

    If val(DCChequeBox.BoundText) = 0 Then Exit Function
 
  '  rs2.Open "TblChecqueBoxContent", Cn, adOpenStatic, adLockOptimistic, adCmdTable
    StrSQL = "SELECT     * from dbo.TblChecqueBoxContent Where (1 = -1)"
   rs2.Open StrSQL, Cn, adOpenKeyset, adLockOptimistic, adCmdText
   
   
    rs2.AddNew
    rs2("noteid").value = NoteID
    rs2("ChequeBoxID").value = val(DCChequeBox.BoundText)
            
    rs2("RecordDate").value = XPDtbTrans.value
    rs2("DueDate").value = DtpChequeDueDate.value
    rs2("BankName").value = TXTBankName.Text
    rs2("ChequeNo").value = TxtChequeNumber.Text
    rs2("ChequeValue").value = val(XPTxtVal.Text)
    
    rs2("Remarks").value = DcboCreditSide.Text
    rs2("Deposited").value = 0
    rs2("Collected").value = 0
    rs2("CreditAccount").value = (DcboCreditSide.BoundText)
    
            If DCboCashType.ListIndex = 0 Then
                        rs2("customeraccount").value = ModAccounts.GetMyAccountCode("TblCustemers", "CusID", val(Me.DBCboClientName.BoundText), "Account_code")
                        rs2("customeraccount1").value = ModAccounts.GetMyAccountCode("TblCustemers", "CusID", val(Me.DBCboClientName.BoundText), "Account_code1")
                        rs2("customeraccount2").value = ModAccounts.GetMyAccountCode("TblCustemers", "CusID", val(Me.DBCboClientName.BoundText), "Account_code2")
                        
             ElseIf DCboCashType.ListIndex = 5 Then
                       rs2("customeraccount").value = get_project_customer_account(val(DBCboClientName.BoundText), "Account_Code")
                        rs2("customeraccount1").value = get_project_customer_account(val(DBCboClientName.BoundText), "Account_Code1")
                        rs2("customeraccount2").value = get_project_customer_account(val(DBCboClientName.BoundText), "Account_Code2")
                        
              
              
            End If
    
    rs2.update
  
    rs2.Close
End Function

Function save_cost_center()

    'on error resume next
    If Not IsNumeric(Text1.Text) Then Exit Function
    Dim i As Integer
    Dim rs2 As ADODB.Recordset
    Set rs2 = New ADODB.Recordset
    Dim sql_str As String

    'Rs.Open "", Cn, adOpenStatic, adLockOptimistic, adCmdTable

    sql_str = "select * from marakes_taklefa_temp where kedno=" & Text1.Text
    rs2.Open sql_str, Cn, adOpenStatic, adLockOptimistic, adCmdText

    For i = 1 To rs2.RecordCount
        rs2("ok").value = 1
        rs2("NoteDate").value = XPDtbTrans.value
        rs2("NoteSerial").value = TxtNoteSerial.Text
        rs2("Remark").value = "”‰œ „ﬁ»Ê÷«     —ﬁ„ " & TxtNoteSerial1.Text & "    " & Me.TxtCustCode
 
        rs2.update
        rs2.MoveNext
    Next i

End Function

Public Function save_General_cost_center(cost_center_id As String, _
                                         cost_center, _
                                         opr_type As String, _
                                         record_date As Date) 'As String, value As Double, depit_or_credit As Boolean, opr_id As Double, opr_type As String, account_no As String, account_name As String, line_no As Double, recorddate As Date)
    Dim i As Integer
    Dim rs2 As New ADODB.Recordset
 
    Dim StrSQL As String

    StrSQL = "Delete  marakes_taklefa_temp  where general_des=1 AND  kedno =" & val(Text1.Text)
    Cn.Execute StrSQL, , adExecuteNoRecords
    
    If Me.DcCostCenter.BoundText = "" Then
        Exit Function
    End If
 
    'rs2.Open "marakes_taklefa_temp", Cn, adOpenStatic, adLockOptimistic, adCmdTable
  StrSQL = "SELECT   *  from dbo.marakes_taklefa_temp Where (1 = -1)"
   rs2.Open StrSQL, Cn, adOpenKeyset, adLockOptimistic, adCmdText
 
    'ÿ—› „œÌ‰
    '       rs2.AddNew
    '       rs2("cost_center_id").value = cost_center_id
    '       rs2("cost_center").value = cost_center
    '       rs2("value").value = XPTxtVal.text
    '       rs2("depit_or_credit").value = "„œÌ‰"
    '       rs2("opr_id").value = Me.Text1.text
    '       rs2("kedno").value = Me.Text1.text
    '
    '       rs2("opr_type").value = opr_type
    '       rs2("account_name").value = DcboDebitSide.text
    '       rs2("account_no").value = DcboDebitSide.BoundText
    '       rs2("line_no").value = Line1
    '       rs2("record_date").value = record_date
    '       rs2.update
    'ÿ—› œ«∆‰
    rs2.AddNew
    rs2("general_des").value = 1
    rs2("cost_center_id").value = cost_center_id
    rs2("cost_center").value = cost_center
    rs2("value").value = XPTxtVal.Text
    rs2("depit_or_credit").value = "œ«∆‰"
    rs2("opr_id").value = Me.Text1.Text
    rs2("kedno").value = Me.Text1.Text

    rs2("opr_type").value = opr_type
    rs2("account_name").value = DcboCreditSide.Text
    rs2("account_no").value = DcboCreditSide.BoundText
    rs2("line_no").value = Line2
    rs2("record_date").value = record_date
    rs2.update
 
    rs2.Close
End Function

Function change_adv_payment_value(note_id As Double, value As Double)
    Dim Rs3 As ADODB.Recordset
    Set Rs3 = New ADODB.Recordset
    Dim sql As String
    Dim i As Integer

    sql = "SELECT * from notes   where  NoteID=" & note_id
 
    Rs3.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
  
    If Rs3.RecordCount = 0 Then Exit Function
' Rs3("Adv_payment_value").value = value
'    Rs3.update
  
End Function

Function Distribute_to_bills(SQL1 As String, CusID As Double)
    Dim Rs3 As ADODB.Recordset
    Set Rs3 = New ADODB.Recordset
    Dim sql As String
    Dim i As Integer

    sql = "SELECT CompanyCreditValues.*  FROM dbo.CompanyCreditValues() CompanyCreditValues  where  requiredvalue>0 and " & SQL1
 
    Rs3.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
  
    If Rs3.RecordCount = 0 Then Exit Function
    Dim total_value As Double
    Dim current_value As Double
    total_value = val(XPTxtVal.Text)
  
    For i = 1 To Rs3.RecordCount

        If total_value > Rs3("requiredvalue") Then
            current_value = Rs3("requiredvalue")
            total_value = total_value - current_value
        
        ElseIf total_value <= Rs3("requiredvalue") Then
            current_value = total_value
            total_value = 0
        ElseIf total_value = 0 Then
            Exit Function
        End If
  
        Add_new_notes Me.XPDtbTrans, 2000, current_value, Rs3("transactionsid").value, CusID, val(DcboBox.BoundText), 1, val(DCboUserName.BoundText)
        Rs3.MoveNext
    Next i

    txtAdv_payment_value.Text = total_value
    change_adv_payment_value XPTxtID.Text, total_value

    ' If IsNull(Rs3("UserName").value) Then FIFO_FUNCTION = "": Exit Function
  
    ' If Not IsNull(Rs3("UserName").value) Then get_user_name = Rs3("UserName").value: Exit Function
    Rs3.Close
 
End Function

Function FIFO_FUNCTION(CusID As Double)
    Dim Rs3 As ADODB.Recordset
    Set Rs3 = New ADODB.Recordset
    Dim sql As String
    Dim i As Integer

   sql = " delete   notes where NoteType= 2000   and  NoteSerial='" & TxtNoteSerial.Text & "'"
 'Cn.Execute sql
Cn.Execute sql


    sql = "SELECT CompanyCreditValues.*  FROM dbo.CompanyCreditValues() CompanyCreditValues  where   (cusid=" & CusID & " and requiredvalue>0  AND TRANSACTION_TYPE=21 )  order by duedate"
 
    Rs3.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
  
    If Rs3.RecordCount = 0 Then Exit Function
    Dim total_value As Double
    Dim current_value As Double
    total_value = val(Me.XPTxtVal.Text)
  
    For i = 1 To Rs3.RecordCount

        If total_value > Rs3("requiredvalue") Then
            current_value = Rs3("requiredvalue")
            total_value = total_value - current_value
        
        ElseIf total_value <= Rs3("requiredvalue") Then
            current_value = total_value
            total_value = 0
        ElseIf total_value = 0 Then
            Exit Function
        End If
  
        Add_new_notes Me.XPDtbTrans, 2000, current_value, Rs3("transactionsid").value, CusID, val(DcboBox.BoundText), 1, val(DCboUserName.BoundText)
        Rs3.MoveNext
    Next i

    ' If IsNull(Rs3("UserName").value) Then FIFO_FUNCTION = "": Exit Function
    txtAdv_payment_value.Text = total_value
  '  change_adv_payment_value XPTxtID.text, total_value
    ' If Not IsNull(Rs3("UserName").value) Then get_user_name = Rs3("UserName").value: Exit Function
    Rs3.Close

End Function

Function Add_new_notes(NoteDate As Date, NoteType As Integer, Note_Value As Double, Transaction_ID As Integer, CusID As Double, BoxID As Integer, displayed As Integer, UserID As Integer)
    Dim RsDev As New ADODB.Recordset
   ' RsDev.Open "notes", Cn, adOpenStatic, adLockOptimistic, adCmdTable
      Dim StrSQL  As String
       StrSQL = "SELECT     dbo.Notes.* from dbo.Notes Where (NoteID = -1)"
   RsDev.Open StrSQL, Cn, adOpenKeyset, adLockOptimistic, adCmdText
   
   
   
    '
    Dim sql As String
    

    RsDev.AddNew
      
    RsDev("NoteID").value = CStr(new_id("Notes", "NoteID", "", True))
    RsDev("NoteSerial").value = TxtNoteSerial.Text ' CStr(new_id("Notes", "NoteSerial", "", True, "NoteType=2000"))
              
    RsDev("NoteDate").value = NoteDate
    RsDev("NoteType").value = NoteType
           
    RsDev("Note_Value").value = Note_Value
    RsDev("Transaction_ID").value = Transaction_ID
    RsDev("CusID").value = CusID
    If BoxID <> 0 Then
    RsDev("BoxID").value = BoxID
    Else
    RsDev("BoxID").value = GetFirstBox
    End If
    RsDev("UserID").value = UserID
    RsDev("displayed").value = 0
           
    RsDev.update

End Function

Private Sub Undo()
    On Error GoTo ErrTrap

    Select Case TxtModFlg.Text

        Case "N"
            clear_all Me
            Me.TxtModFlg.Text = "R"
            XPBtnMove_Click (1)

        Case "E"
            rs.Find "NoteID='" & val(XPTxtID.Text) & "'", , adSearchForward, adBookmarkFirst

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
    On Error GoTo ErrTrap

  If XPTxtID.Text <> "" Then
'        If Me.CboPayMentType.ListIndex = 0 Then
'            If CheckBoxAccount(Me.DcboBox.BoundText, val(Me.XPTxtVal.text), Date, False) = False Then
'                Msg = "·‰ Ì„ﬂ‰ «·”„«Õ »Õ–› Â–« «·⁄„·Ì…..!!!"
'                Msg = Msg & Chr(13) & "ÕÌÀ «‰Â« ”Ê› Ì‰ Ã ⁄‰Â« Œÿ« ›Ï Õ”«»«  «·Œ“‰…"
'                MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
'                Exit Sub
'            End If
'        End If
    
        '      If Me.DCChequeBox.BoundText <> "" Then
        '      If ChequeBoxOperations(Val(Me.XPTxtID)) = False Then
        '          Msg = "·‰ Ì„ﬂ‰ «·”„«Õ »Õ–› Â–« «·⁄„·Ì…..!!!"
        '          Msg = Msg & Chr(13) & "ÕÌÀ «‰Â« ”Ê› Ì‰ Ã ⁄‰Â« Œÿ« ›Ï   Õ«›Ÿ… «·‘Ìﬂ«  ·«‰Â  „ ⁄·ÌÂ« Õ—ﬂ«  «Ìœ«⁄ «Ê  Õ’Ì· "
        '          MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
        '          Exit Sub
        '      End If
        '  End If
    
        Msg = "”Ì „ Õ–› »Ì«‰«  «·⁄„·Ì… —ﬁ„ " & CHR(13)
        Msg = Msg + (TxtNoteSerial.Text) & CHR(13)
        Msg = Msg + " Â·  —€» ›Ì Õ–› Â–Â «·»Ì«‰« ø"

        If MsgBox(Msg, vbYesNo + vbQuestion + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title) = vbYes Then
            If Not rs.RecordCount < 1 Then
                CuurentLogdata ("D")
                rs.delete
                Dim StrSQL As String
                StrSQL = "Delete From TblUnitNoInformation Where NoteID=" & val(XPTxtID.Text)
            Cn.Execute StrSQL, , adExecuteNoRecords
             '   StrSQL = "Delete From notes  Where  (NoteType=2000 OR NoteType=4 ) AND  NoteSerial=" & val(TxtNoteSerial.Text)
             '   Cn.Execute StrSQL, , adExecuteNoRecords
        StrSQL = " delete   TblAqarCommissions where    NoteID=" & val(XPTxtID.Text)
Cn.Execute StrSQL
                StrSQL = "Delete  marakes_taklefa_temp  where kedno =" & val(Text1.Text)
                Cn.Execute StrSQL, , adExecuteNoRecords
       
                StrSQL = "Delete From ReciveDetails Where NoteSerial1='" & val(TxtNoteSerial1.Text) & "'"
                Cn.Execute StrSQL, , adExecuteNoRecords
                StrSQL = "Delete From TblAqrEarnest Where NoteID =" & val(Text1.Text)
                Cn.Execute StrSQL, , adExecuteNoRecords
    
                StrSQL = "Delete From ProjectBillBuy Where TxtNoteSerial='" & TxtNoteSerial.Text & "'"
                Cn.Execute StrSQL, , adExecuteNoRecords
         If val(DCboCashType.ListIndex) = 13 Then
         Cn.Execute "Update TblOtheExpensAqar set FlgPayed=null where ID=" & val(Me.TxtContNo.Text) & "  "
         End If
    
                StrSQL = "Delete From ContracttBillInstallmentsDone Where NoteID =" & val(Me.XPTxtID)
                Cn.Execute StrSQL, , adExecuteNoRecords
                
                StrSQL = "Delete  TblChecqueBoxContent  where NoteID =" & val(Me.XPTxtID)
                Cn.Execute StrSQL, , adExecuteNoRecords
    
    StrSQL = "Delete  TblNotesSales  where NoteID =" & val(Me.XPTxtID)
                Cn.Execute StrSQL, , adExecuteNoRecords
        
   '  StrSQL = " delete   notes where NoteType= 2000   and  NoteSerial='" & TxtNoteSerial.Text & "'"
 'Cn.Execute sql
Cn.Execute StrSQL


                rs.MoveFirst

                If rs.RecordCount < 1 Then
                    clear_all Me
                
                    TxtModFlg_Change
                    XPTxtCurrent.Caption = 0
                    XPTxtCount.Caption = 0
                Else
                    clear_all Me
                    Retrive
                End If

                '--------
                WriteInfo
                '-------
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
    Msg = "⁄›Ê«...ÕœÀ Œÿ√ „« √À‰«¡ Õ–› Â–Â «·»Ì«‰«  " & CHR(13)
    MsgBox Msg, vbMsgBoxRight + vbMsgBoxRtlReading + vbExclamation, App.Title
    rs.CancelUpdate
End Sub

Private Sub ChangeLang()
    lbl(43).Caption = "Cheque Box"
    lbl(50).Caption = "Car"
    lbl(49).Caption = "Driver"
Option7.Caption = "Old Projects"
lbl(48).Caption = "Manual No."
CmdAttach.Caption = "Attachments"
lbl(51).Caption = "Book No."

    Dim XPic As IPictureDisp
    Set XPic = Me.XPBtnMove(1).ButtonImage
    Set Me.XPBtnMove(1).ButtonImage = Me.XPBtnMove(2).ButtonImage
    Set Me.XPBtnMove(2).ButtonImage = XPic
    Set XPic = Me.XPBtnMove(0).ButtonImage
    Set Me.XPBtnMove(0).ButtonImage = Me.XPBtnMove(3).ButtonImage
    Set Me.XPBtnMove(3).ButtonImage = XPic
    lbl(35).Caption = "Adv. Payment"
    Frame1.Caption = "Options"
    Option3.Caption = "Adv. Payment"
    Option2.Caption = "Select Invoice"
    ALLButton3.Caption = "Select"
    lbl(22).Caption = "Current Week"
    Label8.Caption = "General C.C."
    lbl(36).Caption = "From"
 
    Cmd(9).Caption = "GL Print"
 Label3.Caption = "Sales Person."
    Label2.Caption = "Branch"
    lbl(47).Caption = "Value"

    Frame2.Caption = "Project"
    Option4.Caption = "End User"
    Option5.Caption = "Sub-contractor"

    LblLink.Visible = False
    lbl(18).Visible = False
    ALLButton1.Caption = "Installment view"
    ALLButton2.Caption = "debt Voucher"
    Me.Caption = "Cash Receipt Voucher"
    Me.XPTab301.TabCaption(0) = "Receipts"
    Me.XPTab301.TabCaption(1) = "Invoices"
    lbl(37).Caption = "Total Rec."""
    lbl(0).Caption = "Select bills"
    lbl(42).Caption = "Payed  bills"
    CmdRemove.Caption = "Remove Row"

    With Grid

        .TextMatrix(0, .ColIndex("Ser")) = "Ser"
        .TextMatrix(0, .ColIndex("Select")) = "Select"
        .TextMatrix(0, .ColIndex("id")) = "Invoice No."
        .TextMatrix(0, .ColIndex("bill_date")) = "Invoice Date"
        .TextMatrix(0, .ColIndex("total")) = "Invoice Total"
        .TextMatrix(0, .ColIndex("ActualTotal")) = "Payed Totalt"
        .TextMatrix(0, .ColIndex("result")) = "Not Payed"
        .TextMatrix(0, .ColIndex("resultpercentage")) = "Not Payed%"
 
    End With

    With GRID1
        .TextMatrix(0, .ColIndex("Ser")) = "Ser"
        .TextMatrix(0, .ColIndex("bill_id")) = "Invoice Id"
        .TextMatrix(0, .ColIndex("Select")) = "Select"
        .TextMatrix(0, .ColIndex("id")) = "Invoice No."
        .TextMatrix(0, .ColIndex("bill_date")) = "Invoice Date"
        .TextMatrix(0, .ColIndex("total")) = "Invoice Total"
        .TextMatrix(0, .ColIndex("ActualTotal")) = "Payed Totalt"
        .TextMatrix(0, .ColIndex("result")) = "Not Payed"
        .TextMatrix(0, .ColIndex("resultpercentage")) = "Not Payed%"
 
    End With


    Ele(1).Caption = Me.Caption
    lbl(4).Caption = "Opr Code"
    lbl(1).Caption = "Date"
    'lbl(0).Caption = "Type"
    lbl(3).Caption = "Name"
    lbl(2).Caption = "Value"
    lbl(14).Caption = "Cash/Cheque"
    lbl(9).Caption = "Box Name"
    lbl(15).Caption = "Bank Name"
    lbl(16).Caption = "Cheque #"
    lbl(17).Caption = "Cheque Name"
    lbl(5).Caption = "Note"
    ChkTrans.Caption = "From bill"
    lbl(12).Caption = "Bill type"
    lbl(10).Caption = "Bill #"
    lbl(13).Caption = "Current Balance"
    FraInfo.Caption = "Information"
    lbl(22).Caption = "Current Week"

    lbl(23).Caption = "Today Receipts "
    lbl(27).Caption = "Cash"
    lbl(28).Caption = "Cheque"

    lbl(19).Caption = "Week Receipts "

    lbl(21).Caption = "Cash"
    lbl(24).Caption = "Cheque"

    lbl(20).Caption = "Month Receipts "

    lbl(25).Caption = "Cash"
    lbl(26).Caption = "Cheque"
    Fra(1).Caption = "GL"

    lbl(30).Caption = "GL#"
    lbl(29).Caption = "Interval"

    lbl(32).Caption = "Depit"
    lbl(31).Caption = "Credit"
    Cmd(8).Caption = "Table view"
    lbl(8).Caption = "By"
    lbl(7).Caption = "Current "
    lbl(6).Caption = "Records Count "

    Cmd(0).Caption = "New"
    Cmd(1).Caption = "Edit"
    Cmd(2).Caption = "Save"
    Cmd(3).Caption = "Undo"
    Cmd(4).Caption = "Delete"
    Cmd(5).Caption = "Search"
    Cmd(7).Caption = "Print"
    Cmd(6).Caption = "Exit"
    CmdHelp.Caption = "Help"
    DCboCashType.Clear
    DCboCashType.AddItem "To Customer"
    DCboCashType.AddItem "To Vendor"
    DCboCashType.AddItem "Sub-contractor"
    DCboCashType.AddItem "Another Revenues"
    DCboCashType.AddItem "Advanced Payment"
    DCboCashType.AddItem "Projects"
    DCboCashType.AddItem "From Employee"
    DCboCashType.AddItem "From  Account"
DCboCashType.AddItem "From  Contract"
    With Me.CboPaymentType
        .Clear
        .AddItem "Cash"
        .AddItem "Cheque"
        .AddItem "Bank Transfer"
        .AddItem "Coll. Cheque"
        .AddItem "Account"
    End With

    With Me.CboTrans
        .Clear
        .AddItem "Sales invoice"
        .AddItem "Returned purchases"
        .AddItem "Delivery of maintenance for a client"
        .AddItem "Services"
    End With
 
End Sub

Function CuurentLogdata(Optional Currentmode As String)
    LogTextA = "    ‘«‘… " & ScreenNameArabic & CHR(13) & " —ﬁ„ «·”‰œ " & TxtNoteSerial1.Text & CHR(13) & "   «· «—ÌŒ " & XPDtbTrans & CHR(13) & "   ‰Ê⁄ «·„ﬁ»Ê÷«  " & DCboCashType & CHR(13) & "   «·›—⁄  " & Dcbranch & CHR(13) & "   «·«”„  " & DBCboClientName & CHR(13) & "   ﬁÌ„Â «·„ﬁ»Ê÷«   " & XPTxtVal & CHR(13) & "   ÿ—Ìﬁ… «·ﬁ»÷ " & CboPaymentType & CHR(13) & "   «·Œ“Ì‰…  " & DcboBox & CHR(13) & "   «·»‰ﬂ  " & DcboBankName & CHR(13) & "   —ﬁ„ «·‘Ìﬂ  " & TxtChequeNumber & CHR(13) & "    «—ÌŒ «·«” Õﬁ«ﬁ  " & DtpChequeDueDate & CHR(13) & "     »‰«¡ ⁄·Ï   " & XPMTxtRemarks & CHR(13) & "   —ﬁ„ «·ﬁÌœ   " & TxtNoteSerial & CHR(13) & "   —ﬁ„ «·ﬁÌœ   " & TxtNoteSerial & CHR(13) & "ÿ—› „œÌ‰  " & DcboDebitSide & CHR(13) & " ÿ—› œ«∆‰ " & DcboCreditSide & CHR(13) & " «·„‰œÊ» " & DcEmp
                        
    LogTexte = "    Screen  " & ScreenNameEnglish & CHR(13) & " Vchr. NO.  " & TxtNoteSerial1.Text & CHR(13) & "   Date " & XPDtbTrans & CHR(13) & "  Payment Type " & DCboCashType & CHR(13) & "   Branch  " & Dcbranch & CHR(13) & "   Name  " & DBCboClientName & CHR(13) & "  Value" & XPTxtVal & CHR(13) & "   Cash/   Cheque " & CboPaymentType & CHR(13) & "   Box  " & DcboBox & CHR(13) & "   Bank  " & DcboBankName & CHR(13) & "   Cheque No" & TxtChequeNumber & CHR(13) & "  Due Date  " & DtpChequeDueDate & CHR(13) & " Ge NO.  " & TxtNoteSerial & CHR(13) & "Debit " & DcboDebitSide & CHR(13) & "Credit " & DcboCreditSide & CHR(13) & " UserName " & DCboUserName & CHR(13) & " Sales Person " & DcEmp
                     
    If Currentmode <> "D" Then
        AddToLogFile CInt(user_id), 4, Date, Time, LogTextA, LogTexte, Me.Name, Me.TxtModFlg, , , val(TxtNoteSerial), TxtNoteSerial1
    Else
        AddToLogFile CInt(user_id), 4, Date, Time, LogTextA, LogTexte, Me.Name, "D", , , val(TxtNoteSerial), TxtNoteSerial1
    End If
    
End Function

Private Sub Form_KeyDown(KeyCode As Integer, _
                         Shift As Integer)
    On Error GoTo ErrTrap

    If KeyCode = vbKeyReturn Then
        If Me.TxtModFlg.Text = "R" Then
            Cmd_Click (0)
        Else
            Sendkeys "{TAB}"
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
            'Cmd_Click (6)
        End If
    End If

    Exit Sub
ErrTrap:
End Sub

Private Sub AddTip()
    Dim Wrap As String
    On Error GoTo ErrTrap
    Wrap = CHR(13) + CHR(10)
    Set TTP = New clstooltip

    With TTP
        .Create Me.hWnd, "«·„ﬁ»Ê÷« ", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
        .AddControl Cmd(0), "ÃœÌœ ..." & Wrap & "· ”ÃÌ· »Ì«‰«  ⁄„·Ì… ÃœÌœ…" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
    End With

    With TTP
        .Create Me.hWnd, "«·„ﬁ»Ê÷« ", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
        .AddControl Cmd(1), " ⁄œÌ· ..." & Wrap & "· ⁄œÌ· »Ì«‰«  Â–Â «·⁄„·Ì…" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
    End With

    With TTP
        .Create Me.hWnd, "«·„ﬁ»Ê÷« ", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
        .AddControl Cmd(2), "Õ›Ÿ ..." & Wrap & "·Õ›Ÿ »Ì«‰«  «·⁄„·Ì… «·ÃœÌœ…" & Wrap & "·Õ›Ÿ «· ⁄œÌ·« " & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
    End With

    With TTP
        .Create Me.hWnd, "«·„ﬁ»Ê÷« ", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
        .AddControl Cmd(3), " —«Ã⁄ ..." & Wrap & "·· —«Ã⁄ ⁄‰ ⁄„·Ì… «·«÷«›…" & Wrap & "··· —«Ã⁄ ⁄‰ ⁄„·Ì… «· ⁄œÌ·" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
    End With

    With TTP
        .Create Me.hWnd, "«·„ﬁ»Ê÷« ", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
        .AddControl Cmd(4), "Õ–› ..." & Wrap & "·Õ–› »Ì«‰«  «·⁄„·Ì… «·Õ«·Ì…" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
    End With

    With TTP
        .Create Me.hWnd, "«·„ﬁ»Ê÷« ", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
        .AddControl Cmd(6), "Œ—ÊÃ ..." & Wrap & "·«€·«ﬁ Â–Â «·‰«›–…" & Wrap, True
    End With

    With TTP
        .Create Me.hWnd, "«·„ﬁ»Ê÷« ", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
        .AddControl XPBtnMove(1), "«·√Ê· ..." & Wrap & "··«‰ ﬁ«· «·Ï √Ê· ”Ã·" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
    End With

    With TTP
        .Create Me.hWnd, "«·„ﬁ»Ê÷« ", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
        .AddControl XPBtnMove(0), "«·”«»ﬁ ..." & Wrap & "··«‰ ﬁ«· «·Ï «·”Ã· «·”«»ﬁ" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
    End With

    With TTP
        .Create Me.hWnd, "«·„ﬁ»Ê÷« ", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
        .AddControl XPBtnMove(3), "«· «·Ì ..." & Wrap & "··«‰ ﬁ«· «·Ï «·”Ã· «· «·Ì" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
    End With

    With TTP
        .Create Me.hWnd, "«·„ﬁ»Ê÷« ", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
        .AddControl XPBtnMove(2), "«·√ŒÌ— ..." & Wrap & "··«‰ ﬁ«· «·Ï ¬Œ— ”Ã·" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
    End With

    With TTP
        .Create Me.hWnd, "«·„ﬁ»Ê÷« ", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
        .AddControl CmdHelp, "„”«⁄œ… ..." & Wrap & "·· ⁄—› ⁄·Ï ÊŸÌ›… Â–Â «·‰«›–…" & Wrap & "ÊﬂÌ›Ì… «· ⁄«„· „⁄Â«" & Wrap & "≈÷€ÿ Â‰«" & Wrap, True
    End With

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

Private Sub XPDtbTrans_Change()
If Me.TxtModFlg.Text <> "R" Then
    If Trim(TxtNoteSerial1.Text) <> "" Then
        oldtxtNoteSerial1.Text = TxtNoteSerial1.Text
    End If

    TxtNoteSerial.Text = ""
    TxtNoteSerial1.Text = ""
    
    If Me.TxtModFlg.Text <> "R" Then
     
    Txt_DateHigri.value = ToHijriDate(XPDtbTrans.value)
       TxtContNo_Change
End If
End If
End Sub

Private Sub Txt_DateHigri_LostFocus()
      If Me.TxtModFlg.Text <> "R" Then
             XPDtbTrans.value = ToGregorianDate(Txt_DateHigri.value)
        End If
End Sub

Private Sub XPTxtVal_Change()
    'Me.lbl(18).Caption = WriteNo(Me.XPTxtVal.text, 0, True)
    'txtAdv_payment_value.text = Format(Val(XPTxtVal.text), "#,###." & String(Abs(SystemOptions.SysDefCurrencyForamt), "0"))
    If Me.TxtModFlg <> "R" And val(DCboCashType.ListIndex) = 9 Then
txtTotal2.Text = val(XPTxtVal.Text) - val(txtTotal1.Text)
ClculteVAT
If val(XPTxtVal.Text) >= (val(Txtcommission.Text) - val(TxtCommissionOut.Text)) Then
txtComisin.Text = val(Txtcommission.Text) - val(TxtCommissionOut.Text)
Else
txtComisin.Text = val(XPTxtVal.Text)
End If
End If
    If SystemOptions.UserInterface = ArabicInterface Then
        Me.lbl(18).Caption = WriteNo(Format(Me.XPTxtVal.Text, "0.00"), 0, True, ".", , 0)

    Else
 
        Me.lbl(18).Caption = WriteNo(Format(Me.XPTxtVal.Text, "0.00"), 0, True, ".", , 1)

    End If

    'If TxtModFlg.text = "N" Or TxtModFlg.text = "E" And Option3.value = True Then
    If TxtModFlg.Text = "N" Or TxtModFlg.Text = "E" Then
        txtAdv_payment_value.Text = XPTxtVal.Text
    End If

End Sub

Private Sub XPTxtVal_KeyPress(KeyAscii As Integer)
    KeyAscii = KeyAscii_Num(KeyAscii, XPTxtVal.Text, 0)
End Sub

Private Function CheckDebitTrans(LngTransID As Long) As Boolean
    Dim Msg As String
    Dim RsTemp As ADODB.Recordset
    Dim DblCreditNoteValue As Double
    Dim LngDebitNoteID As Long
    Dim StrSQL As String

    CheckDebitTrans = False

    If LngTransID = 0 Then
        Msg = "⁄›Ê« .. ·« ÊÃœ ›« Ê—… »Â–« «·„”·”· „”Ã·… ›Ï «·»—‰«„Ã..!!!"
        Msg = Msg & CHR(13) & "»—Ã«¡ «· «ﬂœ „‰ «·»Ì«‰«  «·„œŒ·…..!!"
        MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
        TxtTransSerial.SetFocus
        Exit Function
    ElseIf LngTransID <> 0 Then
        Set RsTemp = New ADODB.Recordset
        StrSQL = "Select CusID,PaymentType From Transactions where Transaction_ID=" & LngTransID & ""
        RsTemp.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText
    
        If Not (RsTemp.BOF Or RsTemp.EOF) Then
            If RsTemp("PaymentType").value = 0 Then
                Msg = "⁄›Ê« «·›« Ê—… —ﬁ„ " & Trim(Me.TxtTransSerial.Text)
                Msg = Msg & CHR(13) & "›« Ê—… ‰ﬁœÌ… ...Ê·«Ì„ﬂ‰  Õ’Ì· ·Â« „ﬁ»Ê÷« "
                MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
                TxtTransSerial.SetFocus
                Exit Function
            End If

            If Me.DBCboClientName.BoundText <> IIf(IsNull(RsTemp("CusID").value), "", RsTemp("CusID").value) Then
                Msg = "⁄›Ê« «·›« Ê—… —ﬁ„ " & Trim(Me.TxtTransSerial.Text)
                Msg = Msg & CHR(13) & "·Ì”  „”Ã·… „⁄ «·⁄„Ì· " & Me.DBCboClientName.Text
                MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
                TxtTransSerial.SetFocus
                Exit Function
            End If

            If LngTransID <> val(Me.TxtTransID.Text) Then
                Me.TxtTransID.Text = LngTransID
            End If
        
            DblCreditNoteValue = 0
            StrSQL = "SELECT Transactions.Transaction_ID, Transactions.Transaction_Serial," & "Transactions.Transaction_Type, Transactions.PaymentType, " & "Notes.Note_Value, Notes.NoteID "
            StrSQL = StrSQL + " FROM Transactions INNER JOIN Notes ON Transactions.Transaction_ID =" & "Notes.Transaction_ID WHERE (Notes.NoteType=1) AND Transactions.Transaction_ID= " & LngTransID & ""
            Set RsTemp = New ADODB.Recordset
            RsTemp.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText
        
            If Not (RsTemp.BOF Or RsTemp.EOF) Then
                LngDebitNoteID = RsTemp("NoteID").value
                DblCreditNoteValue = IIf(IsNull(RsTemp("Note_Value").value), 0, RsTemp("Note_Value").value)
                '«· «ﬂœ „‰ «‰ Â–Â «·›« Ê—… ·Ì”  ·Â« √ﬁ”«ÿ
                'ÕÌÀ «‰ «·√ﬁ”«ÿ ·«Ì„ﬂ‰  Õ’Ì·Â« „‰ Â‰«
                StrSQL = "Select * From InstallMent Where NoteID=" & LngDebitNoteID & ""
                Set RsTemp = New ADODB.Recordset
                RsTemp.Open StrSQL, Cn, adOpenStatic, adLockReadOnly

                If Not (RsTemp.BOF Or RsTemp.EOF) Then
                    If RsTemp.RecordCount > 0 Then
                        Msg = "⁄›Ê« .. «·ﬁÌ„… «·√Ã·… „‰ «·›« Ê—… ﬁœ  „  ﬁ”ÌÿÂ«..!!"
                        Msg = Msg & CHR(13) & "Ê·«Ì„ﬂ‰  Õ’Ì· «·√ﬁ”«ÿ „‰ ‘«‘… «·„ﬁ»Ê÷« "
                        Msg = Msg & CHR(13) & "≈” Œœ„ ‘«‘…  Õ’Ì· «·√ﬁ”«ÿ »œ·« „‰Â«"
                        MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
                        Exit Function
                    End If
                End If

            Else
                'LngDebitNoteID
                Msg = "·«ÌÊÃœ «Ê—«ﬁ „«·Ì… √Ã·… ⁄·Ï Â–Â «·›« Ê—…..!!"
                MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
                Exit Function
            End If

            If DblCreditNoteValue < val(Me.XPTxtVal.Text) Then
                Msg = "⁄›Ê« ..."
                Msg = Msg & CHR(13) & "«·ﬁÌ„… «·√Ã·… „‰ «·›« Ê—… .. «’€— „‰ «·ﬁÌ„…"
                Msg = Msg & CHR(13) & "«·„—«œ  ”ÃÌ·Â« «·√‰..»—Ã«¡ „—«Ã⁄… «·ﬁÌ„… «·„”Ã·….!"
                Msg = Msg & CHR(13) & "„·ÕÊŸ…:-"
                Msg = Msg & CHR(13) & "«·ﬁÌ„… «·√Ã·… „‰ «·›« Ê—… ÂÏ : " & DblCreditNoteValue
                MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
                Me.XPTxtVal.SetFocus
                Exit Function
            End If

            Set RsTemp = New ADODB.Recordset
            StrSQL = "SELECT Transactions.Transaction_ID, Transactions.Transaction_Serial," & "Transactions.Transaction_Type, Transactions.PaymentType," & "Sum(Notes.Note_Value) AS SumNote_Value "
            StrSQL = StrSQL + " FROM Transactions INNER JOIN Notes ON Transactions.Transaction_ID =" & "Notes.Transaction_ID " & " Where ((Notes.NoteType = 4 OR Notes.NoteType = 9) And Transactions.Transaction_ID = " & LngTransID & ")"

            If Me.TxtModFlg.Text = "E" Then
                StrSQL = StrSQL + " And Notes.NoteID <>" & Me.XPTxtID.Text & ""
            End If

            StrSQL = StrSQL + " GROUP BY Transactions.Transaction_ID, Transactions.Transaction_Serial," & "Transactions.Transaction_Type, Transactions.PaymentType "
            RsTemp.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText
        
            If Not (RsTemp.BOF Or RsTemp.EOF) Then
                If DblCreditNoteValue = RsTemp("SumNote_Value").value Then
                    Msg = "⁄›Ê« ...!!!!!" & CHR(13)
                    Msg = Msg & "·ﬁœ  „  Õ’Ì· „ﬁ»Ê÷«  √Ê (⁄„· Œ’Ê„«  „”„ÊÕ…) ·Â–Â «·›« Ê—… »„« Ì”«ÊÏ «·ﬁÌ„… «·√Ã·… „‰Â«"
                    Msg = Msg & CHR(13) & "Ê·«Ì„ﬂ‰  Õ’Ì· «Ì… „ﬁ»Ê÷«  ≈÷«›Ì… ⁄·ÌÂ«."
                    MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
                    Exit Function
                ElseIf RsTemp("SumNote_Value").value + val(Me.XPTxtVal.Text) > DblCreditNoteValue Then
                    Msg = "⁄›Ê« ..."
                    Msg = Msg & CHR(13) & "·ﬁœ  „  Õ’Ì· „ﬁ»Ê÷«  √Ê (⁄„· Œ’Ê„«  „”„ÊÕ…) „”»ﬁ« ·Â–Â «·›« Ê—…"
                    Msg = Msg & CHR(13) & "Ê»≈÷«›… «·ﬁÌ„… «·Õ«·Ì… ”Ê›   ŒÿÏ «·ﬁÌ„… «·√Ã·… „‰ «·›« Ê—…"
                    Msg = Msg & CHR(13) & "»—Ã«¡ „—«Ã⁄… «·ﬁÌ„… «·„”Ã·…...."
                    Msg = Msg & CHR(13) & "„·ÕÊŸ…:-"
                    Msg = Msg & CHR(13) & "«·ﬁÌ„… «·√Ã·… „‰ «·›« Ê—… ÂÏ : " & DblCreditNoteValue
                    Msg = Msg & CHR(13) & "ﬁÌ„… «·„ﬁ»Ê÷«  «·”«»ﬁ… ·Â–Â «·›« Ê—… : " & RsTemp("SumNote_Value").value
                    MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
                    Exit Function
                End If
            End If

        Else
            Msg = "⁄›Ê« «·›« Ê—… —ﬁ„ " & Trim(Me.TxtTransSerial.Text)
            Msg = Msg & CHR(13) & "·Ì”  „”Ã·… „⁄ «·⁄„Ì· " & Me.DBCboClientName.Text
            MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
            TxtTransSerial.SetFocus
            Exit Function
        End If
    End If

    CheckDebitTrans = True
    Exit Function
ErrTrap:
End Function

Private Function CheckDebitMaintaince(LngTransID As Long) As Boolean
    Dim Msg As String
    Dim RsTemp As ADODB.Recordset
    Dim DblCreditNoteValue As Double
    Dim LngDebitNoteID As Long
    Dim StrSQL As String

    CheckDebitMaintaince = False

    If LngTransID = 0 Then
        Msg = "⁄›Ê« .. ·« ÊÃœ ›« Ê—… »Â–« «·„”·”· „”Ã·… ›Ï «·»—‰«„Ã..!!!"
        Msg = Msg & CHR(13) & "»—Ã«¡ «· «ﬂœ „‰ «·»Ì«‰«  «·„œŒ·…..!!"
        MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
        TxtTransSerial.SetFocus
        Exit Function
    ElseIf LngTransID <> 0 Then
        Set RsTemp = New ADODB.Recordset
        StrSQL = "Select CusID,PaymentType From TblMaintenece where MaintananceID=" & LngTransID & ""
        RsTemp.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText
    
        If Not (RsTemp.BOF Or RsTemp.EOF) Then
            If RsTemp("PaymentType").value = 0 Then
                Msg = "⁄›Ê« «·›« Ê—… —ﬁ„ " & Trim(Me.TxtTransSerial.Text)
                Msg = Msg & CHR(13) & "›« Ê—… ‰ﬁœÌ… ...Ê·«Ì„ﬂ‰  Õ’Ì· ·Â« „ﬁ»Ê÷« "
                MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
                TxtTransSerial.SetFocus
                Exit Function
            End If

            If Me.DBCboClientName.BoundText <> IIf(IsNull(RsTemp("CusID").value), "", RsTemp("CusID").value) Then
                Msg = "⁄›Ê« «·›« Ê—… —ﬁ„ " & Trim(Me.TxtTransSerial.Text)
                Msg = Msg & CHR(13) & "·Ì”  „”Ã·… „⁄ «·⁄„Ì· " & Me.DBCboClientName.Text
                MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
                TxtTransSerial.SetFocus
                Exit Function
            End If

            If LngTransID <> val(Me.TxtTransID.Text) Then
                Me.TxtTransID.Text = LngTransID
            End If
        
            DblCreditNoteValue = 0
            StrSQL = "SELECT Notes.Note_Value, Notes.NoteID, TblMaintenece.MaintananceID," & "TblMaintenece.PaymentType, TblMaintenece.MType "
            StrSQL = StrSQL + " FROM TblMaintenece INNER JOIN Notes ON " & "TblMaintenece.MaintananceID = Notes.MaintananceID " & " WHERE (((Notes.NoteType)=1)) AND TblMaintenece.MaintananceID=" & LngTransID & ""
            Set RsTemp = New ADODB.Recordset
            RsTemp.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

            If Not (RsTemp.BOF Or RsTemp.EOF) Then
                LngDebitNoteID = RsTemp("NoteID").value
                DblCreditNoteValue = IIf(IsNull(RsTemp("Note_Value").value), 0, RsTemp("Note_Value").value)
                '«· «ﬂœ „‰ «‰ Â–Â «·›« Ê—… ·Ì”  ·Â« √ﬁ”«ÿ
                'ÕÌÀ «‰ «·√ﬁ”«ÿ ·«Ì„ﬂ‰  Õ’Ì·Â« „‰ Â‰«
                StrSQL = "Select * From InstallMent Where NoteID=" & LngDebitNoteID & ""
                Set RsTemp = New ADODB.Recordset
                RsTemp.Open StrSQL, Cn, adOpenStatic, adLockReadOnly

                If Not (RsTemp.BOF Or RsTemp.EOF) Then
                    If RsTemp.RecordCount > 0 Then
                        Msg = "⁄›Ê« .. «·ﬁÌ„… «·√Ã·… „‰ «·›« Ê—… ﬁœ  „  ﬁ”ÌÿÂ«..!!"
                        Msg = Msg & CHR(13) & "Ê·«Ì„ﬂ‰  Õ’Ì· «·√ﬁ”«ÿ „‰ ‘«‘… «·„ﬁ»Ê÷« "
                        Msg = Msg & CHR(13) & "≈” Œœ„ ‘«‘…  Õ’Ì· «·√ﬁ”«ÿ »œ·« „‰Â«"
                        MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
                        Exit Function
                    End If
                End If

            Else
                'LngDebitNoteID
                Msg = "·«ÌÊÃœ «Ê—«ﬁ „«·Ì… √Ã·… ⁄·Ï Â–Â «·›« Ê—…..!!"
                MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
                Exit Function
            End If

            If DblCreditNoteValue < val(Me.XPTxtVal.Text) Then
                Msg = "⁄›Ê« ..."
                Msg = Msg & CHR(13) & "«·ﬁÌ„… «·√Ã·… „‰ «·›« Ê—… .. «’€— „‰ «·ﬁÌ„…"
                Msg = Msg & CHR(13) & "«·„—«œ  ”ÃÌ·Â« «·√‰..»—Ã«¡ „—«Ã⁄… «·ﬁÌ„… «·„”Ã·….!"
                Msg = Msg & CHR(13) & "„·ÕÊŸ…:-"
                Msg = Msg & CHR(13) & "«·ﬁÌ„… «·√Ã·… „‰ «·›« Ê—… ÂÏ : " & DblCreditNoteValue
                MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
                Me.XPTxtVal.SetFocus
                Exit Function
            End If

            Set RsTemp = New ADODB.Recordset
        
            StrSQL = "SELECT  TblMaintenece.MaintananceID," & "TblMaintenece.MType, TblMaintenece.PaymentType," & "Sum(Notes.Note_Value) AS SumNote_Value "
            StrSQL = StrSQL + " FROM TblMaintenece INNER JOIN Notes ON TblMaintenece.MaintananceID =" & "Notes.MaintananceID " & " Where ((Notes.NoteType = 4) And TblMaintenece.MaintananceID = " & LngTransID & ")"

            If Me.TxtModFlg.Text = "E" Then
                StrSQL = StrSQL + " And Notes.NoteID <>" & Me.XPTxtID.Text & ""
            End If

            StrSQL = StrSQL + " GROUP BY TblMaintenece.MaintananceID," & "TblMaintenece.MType, TblMaintenece.PaymentType"
        
            RsTemp.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText
        
            If Not (RsTemp.BOF Or RsTemp.EOF) Then
                If DblCreditNoteValue = RsTemp("SumNote_Value").value Then
                    Msg = "⁄›Ê« ...!!!!!"
                    Msg = Msg & "·ﬁœ  „  Õ’Ì· „ﬁ»Ê÷«  ·Â–Â «·›« Ê—… »„« Ì”«ÊÏ «·ﬁÌ„… «·√Ã·… „‰Â«"
                    Msg = Msg & CHR(13) & "Ê·«Ì„ﬂ‰  Õ’Ì· «Ì… „ﬁ»Ê÷«  ≈÷«›Ì… ⁄·ÌÂ«."
                    MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
                    Exit Function
                ElseIf RsTemp("SumNote_Value").value + val(Me.XPTxtVal.Text) > DblCreditNoteValue Then
                    Msg = "⁄›Ê« ..."
                    Msg = Msg & CHR(13) & "·ﬁœ  „  Õ’Ì· „ﬁ»Ê÷«  „”»ﬁ« ·Â–Â «·›« Ê—…"
                    Msg = Msg & CHR(13) & "Ê»≈÷«›… «·ﬁÌ„… «·Õ«·Ì… ”Ê›   ŒÿÏ «·ﬁÌ„… «·√Ã·… „‰ «·›« Ê—…"
                    Msg = Msg & CHR(13) & "»—Ã«¡ „—«Ã⁄… «·ﬁÌ„… «·„”Ã·…...."
                    Msg = Msg & CHR(13) & "„·ÕÊŸ…:-"
                    Msg = Msg & CHR(13) & "«·ﬁÌ„… «·√Ã·… „‰ «·›« Ê—… ÂÏ : " & DblCreditNoteValue
                    Msg = Msg & CHR(13) & "ﬁÌ„… «·„ﬁ»Ê÷«  «·”«»ﬁ… ·Â–Â «·›« Ê—… : " & RsTemp("SumNote_Value").value
                    MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
                    Exit Function
                End If
            End If

        Else
            Msg = "⁄›Ê« «·›« Ê—… —ﬁ„ " & Trim(Me.TxtTransSerial.Text)
            Msg = Msg & CHR(13) & "·Ì”  „”Ã·… „⁄ «·⁄„Ì· " & Me.DBCboClientName.Text
            MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
            TxtTransSerial.SetFocus
            Exit Function
        End If
    End If

    CheckDebitMaintaince = True
    Exit Function
ErrTrap:
End Function


Private Sub WriteInfo()
    Dim rs2 As ADODB.Recordset
    Dim StrSQL As String
    Dim StartWeekDate As Date
    Dim EndWeekDate As Date
    Dim StrTemp As String
    Dim i As Integer

    StartWeekDate = GetWeekStartEND(Date, 0)
    EndWeekDate = DateAdd("d", 7, StartWeekDate)

    If SystemOptions.UserInterface = ArabicInterface Then
        StrTemp = "«·≈”»Ê⁄ «·Õ«·Ï „‰ " & DisplayDate(StartWeekDate)
        StrTemp = StrTemp & " ≈·Ï " & DisplayDate(EndWeekDate)
    Else
        StrTemp = "«Current Week From " & DisplayDate(StartWeekDate)
        StrTemp = StrTemp & " To " & DisplayDate(EndWeekDate)

    End If

    Me.lbl(22).Caption = StrTemp

    For i = LblLinkInfo.LBound To LblLinkInfo.UBound
        LblLinkInfo(i).Caption = "0"
    Next i

    '------------------------------------------------------------------------------
    '„ﬁ»Ê÷«  «·ÌÊ„
    StrSQL = " SELECT     SUM(Note_Value2) AS SumX, NoteCashingType"
    StrSQL = StrSQL + " From Notes "
    StrSQL = StrSQL + " Where (NoteType = 4) "
    StrSQL = StrSQL + " AND NoteDate=" & SQLDate(Date, True)
    StrSQL = StrSQL + " GROUP BY NoteCashingType"
    StrSQL = StrSQL + " Order BY NoteCashingType"
    Set rs2 = New ADODB.Recordset
    rs2.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

    If Not (rs2.BOF Or rs2.EOF) Then
        rs2.MoveFirst

        For i = 0 To rs2.RecordCount - 1

            If rs2("NoteCashingType").value = 0 Then
                Me.LblLinkInfo(0).Caption = IIf(IsNull(rs2("SumX").value), 0, rs2("SumX").value)
            ElseIf rs2("NoteCashingType").value = 1 Then
                Me.LblLinkInfo(1).Caption = IIf(IsNull(rs2("SumX").value), 0, rs2("SumX").value)
            End If

            rs2.MoveNext
        Next

        Me.LblLinkInfo(6).Caption = val(Me.LblLinkInfo(0).Caption) + val(Me.LblLinkInfo(1).Caption)
    Else
        Me.LblLinkInfo(0).Caption = 0
        Me.LblLinkInfo(1).Caption = 0
        Me.LblLinkInfo(6).Caption = 0
    End If

    '------------------------------------------------------------------------------
    '„ﬁ»Ê÷«  «·√”»Ê⁄ «·Õ«·Ï
    StrSQL = " SELECT     SUM(Note_Value2) AS SumX, NoteCashingType"
    StrSQL = StrSQL + " From Notes "
    StrSQL = StrSQL + " Where (NoteType = 4) "
    StrSQL = StrSQL + " AND NoteDate >=" & SQLDate(StartWeekDate, True)
    StrSQL = StrSQL + " AND NoteDate <=" & SQLDate(EndWeekDate, True)
    StrSQL = StrSQL + " GROUP BY NoteCashingType"
    StrSQL = StrSQL + " Order BY NoteCashingType"
    Set rs2 = New ADODB.Recordset
    rs2.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

    If Not (rs2.BOF Or rs2.EOF) Then
        rs2.MoveFirst

        For i = 0 To rs2.RecordCount - 1

            If rs2("NoteCashingType").value = 0 Then
                Me.LblLinkInfo(2).Caption = IIf(IsNull(rs2("SumX").value), 0, rs2("SumX").value)
            ElseIf rs2("NoteCashingType").value = 1 Then
                Me.LblLinkInfo(3).Caption = IIf(IsNull(rs2("SumX").value), 0, rs2("SumX").value)
            End If

            rs2.MoveNext
        Next

        Me.LblLinkInfo(7).Caption = val(Me.LblLinkInfo(2).Caption) + val(Me.LblLinkInfo(3).Caption)
    Else
        Me.LblLinkInfo(0).Caption = 0
        Me.LblLinkInfo(1).Caption = 0
        Me.LblLinkInfo(7).Caption = 0
    End If

    '------------------------------------------------------------------------------
    '„ﬁ»Ê÷«  «·‘Â— «·Õ«·Ï
    StrSQL = " SELECT     SUM(Note_Value2) AS SumX, NoteCashingType"
    StrSQL = StrSQL + " From Notes "
    StrSQL = StrSQL + " Where (NoteType = 4) "
    StrSQL = StrSQL + " AND Month(NoteDate)=" & Month(Date) & ""
    StrSQL = StrSQL + " AND Year(NoteDate)=" & year(Date) & ""
    StrSQL = StrSQL + " GROUP BY NoteCashingType"
    StrSQL = StrSQL + " Order BY NoteCashingType"
    Set rs2 = New ADODB.Recordset
    rs2.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

    If Not (rs2.BOF Or rs2.EOF) Then
        rs2.MoveFirst

        For i = 0 To rs2.RecordCount - 1

            If rs2("NoteCashingType").value = 0 Then
                Me.LblLinkInfo(4).Caption = IIf(IsNull(rs2("SumX").value), 0, rs2("SumX").value)
            ElseIf rs2("NoteCashingType").value = 1 Then
                Me.LblLinkInfo(5).Caption = IIf(IsNull(rs2("SumX").value), 0, rs2("SumX").value)
            End If

            rs2.MoveNext
        Next
        Me.LblLinkInfo(8).Caption = val(Me.LblLinkInfo(4).Caption) + val(Me.LblLinkInfo(5).Caption)
    Else
        Me.LblLinkInfo(4).Caption = 0
        Me.LblLinkInfo(5).Caption = 0
        Me.LblLinkInfo(8).Caption = 0
    End If
End Sub
