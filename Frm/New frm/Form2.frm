VERSION 5.00
Object = "{C115893A-A3BF-43AF-B28D-69DB846077F3}#1.0#0"; "vsflex8u.ocx"
Object = "{0AFE7BE0-11B7-4A3E-978D-D4501E9A57FE}#1.0#0"; "c1sizer.ocx"
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{28D9BF84-BC20-11D2-94CF-004005455FAA}#1.1#0"; "ImpulseAniLabel.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{FE5DCFAD-BC1D-11D2-94CF-004005455FAA}#1.4#0"; "ImpulseButton.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{396F7AC0-A0DD-11D3-93EC-00C0DFE7442A}#1.0#0"; "vbalIml6.ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{49003D3A-66CD-11D7-A449-E937BE2D9041}#1.0#0"; "ALLBUTTONS.ocx"
Begin VB.Form FrmCashing 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "«·„ﬁ»Ê÷« "
   ClientHeight    =   8625
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   12705
   Icon            =   "Form2.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   8625
   ScaleWidth      =   12705
   Begin VB.TextBox txtoldvalue 
      Alignment       =   1  'Right Justify
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
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
      Top             =   8760
      Visible         =   0   'False
      Width           =   2685
   End
   Begin C1SizerLibCtl.C1Elastic Ele 
      Height          =   585
      Index           =   1
      Left            =   150
      TabIndex        =   28
      TabStop         =   0   'False
      Top             =   30
      Width           =   12705
      _cx             =   22410
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
      Begin VB.TextBox XPTxtID1 
         Height          =   495
         Left            =   7920
         TabIndex        =   198
         Top             =   -600
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.TextBox oldtxtNoteSerial1 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   7800
         RightToLeft     =   -1  'True
         TabIndex        =   125
         Top             =   840
         Visible         =   0   'False
         Width           =   1425
      End
      Begin VB.TextBox TxtModFlg 
         Alignment       =   1  'Right Justify
         BackColor       =   &H000000FF&
         Height          =   345
         Left            =   4950
         RightToLeft     =   -1  'True
         TabIndex        =   30
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
         TabIndex        =   29
         TabStop         =   0   'False
         Top             =   90
         Visible         =   0   'False
         Width           =   495
      End
      Begin ImpulseButton.ISButton XPBtnMove 
         Height          =   375
         Index           =   0
         Left            =   1125
         TabIndex        =   31
         Top             =   90
         Width           =   495
         _ExtentX        =   873
         _ExtentY        =   661
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
         ButtonImage     =   "Form2.frx":000C
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
         TabIndex        =   32
         Top             =   90
         Width           =   495
         _ExtentX        =   873
         _ExtentY        =   661
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
         ButtonImage     =   "Form2.frx":03A6
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
         TabIndex        =   33
         Top             =   90
         Width           =   495
         _ExtentX        =   873
         _ExtentY        =   661
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
         ButtonImage     =   "Form2.frx":0740
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
         TabIndex        =   34
         Top             =   90
         Width           =   495
         _ExtentX        =   873
         _ExtentY        =   661
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
         ButtonImage     =   "Form2.frx":0ADA
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
         TabIndex        =   35
         Top             =   60
         Visible         =   0   'False
         Width           =   1545
         _ExtentX        =   2725
         _ExtentY        =   661
         ButtonStyle     =   1
         ButtonPositionImage=   1
         Caption         =   "«·⁄—÷ «·ÃœÊ·Ï"
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
      Begin MSAdodcLib.Adodc numbering 
         Height          =   585
         Left            =   1680
         Top             =   0
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
            Name            =   "MS Sans Serif"
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
         Top             =   -480
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
            Name            =   "MS Sans Serif"
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
         Left            =   5880
         Picture         =   "Form2.frx":0E74
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
         TabIndex        =   36
         Top             =   60
         Visible         =   0   'False
         Width           =   705
      End
   End
   Begin C1SizerLibCtl.C1Tab XPTab301 
      Height          =   7080
      Left            =   0
      TabIndex        =   23
      Top             =   480
      Width           =   12690
      _cx             =   22384
      _cy             =   12488
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
      Caption         =   "«·„ﬁ»Ê÷« |«Œ Ì«—  „” Œ·’«  «·„‘«—Ì⁄|«·œ›⁄« |Õ«·…«·«⁄ „«œ"
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
      Picture(0)      =   "Form2.frx":4ADC
      Flags(2)        =   2
      Begin C1SizerLibCtl.C1Elastic Ele 
         Height          =   6615
         Index           =   12
         Left            =   45
         TabIndex        =   24
         TabStop         =   0   'False
         Top             =   45
         Width           =   12600
         _cx             =   22225
         _cy             =   11668
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
         Begin VB.Frame FramePay 
            BackColor       =   &H00E0E0E0&
            Caption         =   "«·„»·€ «·„œ›Ê⁄"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   6855
            Left            =   -90
            RightToLeft     =   -1  'True
            TabIndex        =   243
            Top             =   -60
            Visible         =   0   'False
            Width           =   12735
            Begin VB.Frame Frame13 
               BackColor       =   &H00FFFFFF&
               Height          =   5055
               Left            =   120
               TabIndex        =   263
               Top             =   480
               Width           =   5535
               Begin ImpulseButton.ISButton CmdNos 
                  Height          =   975
                  Index           =   0
                  Left            =   4320
                  TabIndex        =   264
                  Top             =   3970
                  Width           =   1095
                  _ExtentX        =   1931
                  _ExtentY        =   1720
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
                  ButtonImage     =   "Form2.frx":4E76
               End
               Begin ImpulseButton.ISButton CmdNos 
                  Height          =   975
                  Index           =   1
                  Left            =   2160
                  TabIndex        =   265
                  Top             =   3000
                  Width           =   1095
                  _ExtentX        =   1931
                  _ExtentY        =   1720
                  Caption         =   ""
                  BackColor       =   16777215
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "MS Sans Serif"
                     Size            =   8.25
                     Charset         =   178
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ButtonImage     =   "Form2.frx":5636
                  ColorButton     =   16777215
               End
               Begin ImpulseButton.ISButton CmdNos 
                  Height          =   975
                  Index           =   2
                  Left            =   3240
                  TabIndex        =   266
                  Top             =   3000
                  Width           =   1095
                  _ExtentX        =   1931
                  _ExtentY        =   1720
                  Caption         =   ""
                  BackColor       =   16777215
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "MS Sans Serif"
                     Size            =   8.25
                     Charset         =   178
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ButtonImage     =   "Form2.frx":5C38
                  ColorButton     =   16777215
               End
               Begin ImpulseButton.ISButton CmdNos 
                  Height          =   975
                  Index           =   3
                  Left            =   4320
                  TabIndex        =   267
                  Top             =   3000
                  Width           =   1095
                  _ExtentX        =   1931
                  _ExtentY        =   1720
                  Caption         =   ""
                  BackColor       =   16777215
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "MS Sans Serif"
                     Size            =   8.25
                     Charset         =   178
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ButtonImage     =   "Form2.frx":641F
                  ColorButton     =   16777215
               End
               Begin ImpulseButton.ISButton CmdNos 
                  Height          =   975
                  Index           =   4
                  Left            =   2160
                  TabIndex        =   268
                  Top             =   2040
                  Width           =   1095
                  _ExtentX        =   1931
                  _ExtentY        =   1720
                  Caption         =   ""
                  BackColor       =   16777215
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "MS Sans Serif"
                     Size            =   8.25
                     Charset         =   178
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ButtonImage     =   "Form2.frx":6C34
                  ColorButton     =   16777215
               End
               Begin ImpulseButton.ISButton CmdNos 
                  Height          =   975
                  Index           =   5
                  Left            =   3240
                  TabIndex        =   269
                  Top             =   2040
                  Width           =   1095
                  _ExtentX        =   1931
                  _ExtentY        =   1720
                  Caption         =   ""
                  BackColor       =   16777215
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "MS Sans Serif"
                     Size            =   8.25
                     Charset         =   178
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ButtonImage     =   "Form2.frx":73BF
                  ColorButton     =   16777215
               End
               Begin ImpulseButton.ISButton CmdNos 
                  Height          =   975
                  Index           =   6
                  Left            =   4320
                  TabIndex        =   270
                  Top             =   2040
                  Width           =   1095
                  _ExtentX        =   1931
                  _ExtentY        =   1720
                  Caption         =   ""
                  BackColor       =   16777215
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "MS Sans Serif"
                     Size            =   8.25
                     Charset         =   178
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ButtonImage     =   "Form2.frx":7B7E
                  ColorButton     =   16777215
               End
               Begin ImpulseButton.ISButton CmdNos 
                  Height          =   975
                  Index           =   7
                  Left            =   2160
                  TabIndex        =   271
                  Top             =   1080
                  Width           =   1095
                  _ExtentX        =   1931
                  _ExtentY        =   1720
                  Caption         =   ""
                  BackColor       =   16777215
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "MS Sans Serif"
                     Size            =   8.25
                     Charset         =   178
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ButtonImage     =   "Form2.frx":8318
                  ColorButton     =   16777215
               End
               Begin ImpulseButton.ISButton CmdNos 
                  Height          =   975
                  Index           =   8
                  Left            =   3240
                  TabIndex        =   272
                  Top             =   1080
                  Width           =   1095
                  _ExtentX        =   1931
                  _ExtentY        =   1720
                  Caption         =   ""
                  BackColor       =   16777215
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "MS Sans Serif"
                     Size            =   8.25
                     Charset         =   178
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ButtonImage     =   "Form2.frx":8A1B
                  ColorButton     =   16777215
               End
               Begin ImpulseButton.ISButton CmdNos 
                  Height          =   975
                  Index           =   9
                  Left            =   4320
                  TabIndex        =   273
                  Top             =   1080
                  Width           =   1095
                  _ExtentX        =   1931
                  _ExtentY        =   1720
                  Caption         =   ""
                  BackColor       =   16777215
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "MS Sans Serif"
                     Size            =   8.25
                     Charset         =   178
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ButtonImage     =   "Form2.frx":9236
                  ColorButton     =   16777215
               End
               Begin ImpulseButton.ISButton CmdNos 
                  Height          =   975
                  Index           =   10
                  Left            =   3240
                  TabIndex        =   274
                  Top             =   3970
                  Width           =   1095
                  _ExtentX        =   1931
                  _ExtentY        =   1720
                  Caption         =   ""
                  BackColor       =   16777215
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "MS Sans Serif"
                     Size            =   8.25
                     Charset         =   178
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ButtonImage     =   "Form2.frx":99C5
                  ColorButton     =   16777215
               End
               Begin ImpulseButton.ISButton CmdNos 
                  Height          =   975
                  Index           =   11
                  Left            =   2160
                  TabIndex        =   275
                  Top             =   3970
                  Width           =   1095
                  _ExtentX        =   1931
                  _ExtentY        =   1720
                  Caption         =   ""
                  BackColor       =   16777215
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "MS Sans Serif"
                     Size            =   8.25
                     Charset         =   178
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ButtonImage     =   "Form2.frx":A50C
                  ColorButton     =   16777215
               End
               Begin ImpulseButton.ISButton CmdNos 
                  Height          =   975
                  Index           =   12
                  Left            =   120
                  TabIndex        =   276
                  Top             =   1080
                  Width           =   1095
                  _ExtentX        =   1931
                  _ExtentY        =   1720
                  Caption         =   ""
                  BackColor       =   16777215
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "MS Sans Serif"
                     Size            =   8.25
                     Charset         =   178
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ButtonImage     =   "Form2.frx":A9FE
                  ColorButton     =   16777215
               End
               Begin ImpulseButton.ISButton CmdNos 
                  Height          =   975
                  Index           =   13
                  Left            =   1200
                  TabIndex        =   277
                  Top             =   1080
                  Width           =   1095
                  _ExtentX        =   1931
                  _ExtentY        =   1720
                  Caption         =   ""
                  BackColor       =   16777215
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "MS Sans Serif"
                     Size            =   8.25
                     Charset         =   178
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ButtonImage     =   "Form2.frx":B265
                  ColorButton     =   16777215
               End
               Begin ImpulseButton.ISButton CmdNos 
                  Height          =   2895
                  Index           =   14
                  Left            =   120
                  TabIndex        =   278
                  Top             =   2040
                  Width           =   2055
                  _ExtentX        =   3625
                  _ExtentY        =   5106
                  Caption         =   ""
                  BackColor       =   16777215
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "MS Sans Serif"
                     Size            =   8.25
                     Charset         =   178
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ButtonImage     =   "Form2.frx":B976
                  ButtonImageDisabled=   "Form2.frx":CD24
                  ColorButton     =   16777215
               End
               Begin VB.Label LBLPayVal 
                  Alignment       =   2  'Center
                  BackStyle       =   0  'Transparent
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   24
                     Charset         =   178
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H00FFFFFF&
                  Height          =   615
                  Left            =   1200
                  TabIndex        =   279
                  Top             =   360
                  Width           =   3375
               End
               Begin VB.Image Image13 
                  Height          =   1035
                  Left            =   120
                  Picture         =   "Form2.frx":D0BF
                  Stretch         =   -1  'True
                  Top             =   120
                  Width           =   5295
               End
            End
            Begin VB.CommandButton CmdValue 
               Caption         =   "1500"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   24
                  Charset         =   178
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   615
               Index           =   9
               Left            =   3000
               TabIndex        =   262
               Top             =   7320
               Width           =   1215
            End
            Begin VB.CommandButton CmdValue 
               Caption         =   "2000"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   24
                  Charset         =   178
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   615
               Index           =   8
               Left            =   4200
               TabIndex        =   261
               Top             =   7320
               Width           =   1335
            End
            Begin VB.Frame Frame11 
               BackColor       =   &H00E0E0E0&
               Height          =   1935
               Left            =   5760
               RightToLeft     =   -1  'True
               TabIndex        =   253
               Top             =   4440
               Width           =   7080
               Begin VB.CommandButton Command1 
                  Caption         =   "⁄—÷ «·ﬂ·"
                  Height          =   375
                  Left            =   5280
                  TabIndex        =   257
                  Top             =   360
                  Width           =   1335
               End
               Begin VB.TextBox TxtNetValue2 
                  Alignment       =   2  'Center
                  BackColor       =   &H00000000&
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   18
                     Charset         =   178
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H0000FF00&
                  Height          =   600
                  Left            =   0
                  Locked          =   -1  'True
                  TabIndex        =   256
                  Top             =   240
                  Width           =   2460
               End
               Begin VB.TextBox TxtPayedValue2 
                  Alignment       =   2  'Center
                  BackColor       =   &H00000000&
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   18
                     Charset         =   178
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H0000FF00&
                  Height          =   555
                  Left            =   0
                  Locked          =   -1  'True
                  TabIndex        =   255
                  Top             =   840
                  Width           =   2445
               End
               Begin VB.TextBox TxtRemainValue2 
                  Alignment       =   2  'Center
                  BackColor       =   &H00000000&
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   18
                     Charset         =   178
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H0000FF00&
                  Height          =   555
                  Left            =   0
                  Locked          =   -1  'True
                  TabIndex        =   254
                  Top             =   1320
                  Width           =   2445
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  BackStyle       =   0  'Transparent
                  Caption         =   "«·«Ã„«·Ì"
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   13.5
                     Charset         =   178
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H00000000&
                  Height          =   435
                  Index           =   101
                  Left            =   3600
                  TabIndex        =   260
                  Top             =   240
                  Width           =   855
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  BackStyle       =   0  'Transparent
                  Caption         =   "«·„œ›Ê⁄"
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   13.5
                     Charset         =   178
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H00000000&
                  Height          =   435
                  Index           =   100
                  Left            =   3600
                  TabIndex        =   259
                  Top             =   840
                  Width           =   855
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  BackStyle       =   0  'Transparent
                  Caption         =   "«·„ »ﬁÌ"
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   13.5
                     Charset         =   178
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H00000000&
                  Height          =   435
                  Index           =   99
                  Left            =   3600
                  TabIndex        =   258
                  Top             =   1440
                  Width           =   855
               End
            End
            Begin VB.CommandButton CmdValue 
               Caption         =   "1000"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   24
                  Charset         =   178
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   615
               Index           =   7
               Left            =   1560
               TabIndex        =   252
               Top             =   7320
               Width           =   1455
            End
            Begin VB.CommandButton CmdValue 
               Caption         =   "500"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   24
                  Charset         =   178
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   615
               Index           =   6
               Left            =   240
               TabIndex        =   251
               Top             =   7320
               Width           =   1335
            End
            Begin VB.CommandButton CmdValue 
               Caption         =   "200"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   24
                  Charset         =   178
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   615
               Index           =   5
               Left            =   4200
               TabIndex        =   250
               Top             =   6720
               Width           =   1335
            End
            Begin VB.CommandButton CmdValue 
               Caption         =   "100"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   24
                  Charset         =   178
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   615
               Index           =   4
               Left            =   3000
               TabIndex        =   249
               Top             =   6720
               Width           =   1215
            End
            Begin VB.CommandButton CmdValue 
               Caption         =   "50"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   24
                  Charset         =   178
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   615
               Index           =   3
               Left            =   1560
               TabIndex        =   248
               Top             =   6720
               Width           =   1455
            End
            Begin VB.CommandButton CmdValue 
               Caption         =   "20"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   24
                  Charset         =   178
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   615
               Index           =   2
               Left            =   240
               TabIndex        =   247
               Top             =   6720
               Width           =   1335
            End
            Begin VB.CommandButton CmdValue 
               Caption         =   "10"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   24
                  Charset         =   178
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   615
               Index           =   1
               Left            =   8640
               TabIndex        =   246
               Top             =   2640
               Visible         =   0   'False
               Width           =   1215
            End
            Begin VB.CommandButton CmdValue 
               Caption         =   "5"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   24
                  Charset         =   178
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   615
               Index           =   0
               Left            =   5760
               TabIndex        =   245
               Top             =   120
               Visible         =   0   'False
               Width           =   1215
            End
            Begin ImpulseButton.ISButton CMDPAy 
               Height          =   1695
               Left            =   240
               TabIndex        =   244
               Top             =   5205
               Width           =   5295
               _ExtentX        =   9340
               _ExtentY        =   2990
               Caption         =   "”œ«œ"
               ForeColor       =   16777215
               FontSize        =   24
               FontBold        =   -1  'True
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "MS Sans Serif"
                  Size            =   24
                  Charset         =   178
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ButtonImage     =   "Form2.frx":D475
               ColorHoverText  =   16777215
               ColorToggledText=   16777215
               ColorToggledHoverText=   16777215
               AlignmentIgnoreImage=   -1  'True
            End
            Begin VSFlex8UCtl.VSFlexGrid Grid22 
               Height          =   3885
               Left            =   5760
               TabIndex        =   280
               Top             =   600
               Width           =   6885
               _cx             =   12144
               _cy             =   6853
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
               BackColor       =   -2147483640
               ForeColor       =   65280
               BackColorFixed  =   14871017
               ForeColorFixed  =   -2147483630
               BackColorSel    =   -2147483641
               ForeColorSel    =   -2147483634
               BackColorBkg    =   -2147483636
               BackColorAlternate=   -2147483640
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
               Rows            =   6
               Cols            =   12
               FixedRows       =   1
               FixedCols       =   1
               RowHeightMin    =   650
               RowHeightMax    =   0
               ColWidthMin     =   0
               ColWidthMax     =   0
               ExtendLastCol   =   -1  'True
               FormatString    =   $"Form2.frx":D9EF
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
            Begin VB.Label lblexit 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "X"
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
               Index           =   90
               Left            =   9120
               TabIndex        =   282
               Top             =   240
               Width           =   570
            End
            Begin VB.Label Label20 
               Alignment       =   2  'Center
               BackStyle       =   0  'Transparent
               Caption         =   "X"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   13.5
                  Charset         =   178
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H000000FF&
               Height          =   495
               Left            =   11040
               TabIndex        =   281
               Top             =   540
               Visible         =   0   'False
               Width           =   1335
            End
         End
         Begin VB.Frame Frame12 
            BackColor       =   &H0080FFFF&
            Caption         =   "»Ì«‰«  ›Ê« Ì— «·„»Ì⁄« "
            Height          =   6015
            Left            =   -120
            RightToLeft     =   -1  'True
            TabIndex        =   236
            Top             =   120
            Visible         =   0   'False
            Width           =   12735
            Begin VB.CommandButton Command10 
               BackColor       =   &H8000000B&
               Caption         =   "«·€«¡ «·”œ«œ"
               Height          =   315
               Left            =   9120
               RightToLeft     =   -1  'True
               TabIndex        =   238
               Top             =   240
               Width           =   1695
            End
            Begin VB.CheckBox Check1 
               Alignment       =   1  'Right Justify
               BackColor       =   &H0080FFFF&
               Caption         =   " ÕœÌœ «·ﬂ·"
               Height          =   195
               Left            =   10800
               RightToLeft     =   -1  'True
               TabIndex        =   237
               Top             =   300
               Width           =   1200
            End
            Begin VSFlex8Ctl.VSFlexGrid VSFlexGrid1 
               Height          =   4860
               Left            =   0
               TabIndex        =   239
               Top             =   480
               Width           =   12480
               _cx             =   22013
               _cy             =   8572
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
               SelectionMode   =   1
               GridLines       =   1
               GridLinesFixed  =   2
               GridLineWidth   =   1
               Rows            =   1
               Cols            =   20
               FixedRows       =   1
               FixedCols       =   1
               RowHeightMin    =   0
               RowHeightMax    =   0
               ColWidthMin     =   0
               ColWidthMax     =   0
               ExtendLastCol   =   -1  'True
               FormatString    =   $"Form2.frx":DBD5
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
               ExplorerBar     =   1
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
            Begin VB.Label Label29 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "X"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   178
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H000000FF&
               Height          =   375
               Left            =   12360
               RightToLeft     =   -1  'True
               TabIndex        =   242
               Top             =   240
               Width           =   135
            End
            Begin VB.Label Label28 
               Alignment       =   1  'Right Justify
               BackColor       =   &H8000000B&
               BackStyle       =   0  'Transparent
               Height          =   255
               Left            =   240
               RightToLeft     =   -1  'True
               TabIndex        =   241
               ToolTipText     =   "Ì „  Õ„Ì· Â–« «·„’—Ê› ⁄·Ï «·⁄„Ê·«  «·»‰ﬂÌ…"
               Top             =   5640
               Width           =   8775
            End
            Begin VB.Label Label27 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00FFFFFF&
               BackStyle       =   0  'Transparent
               Caption         =   "≈Ã„«·Ì «·›Ê« Ì—"
               Height          =   255
               Left            =   9600
               RightToLeft     =   -1  'True
               TabIndex        =   240
               ToolTipText     =   "Ì „  Õ„Ì· Â–« «·„’—Ê› ⁄·Ï «·⁄„Ê·«  «·»‰ﬂÌ…"
               Top             =   5640
               Width           =   1575
            End
         End
         Begin VB.TextBox txtContainerNo 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00C0E0FF&
            Height          =   315
            Left            =   5880
            RightToLeft     =   -1  'True
            TabIndex        =   235
            Top             =   960
            Visible         =   0   'False
            Width           =   1545
         End
         Begin VB.TextBox TxtVATValue 
            Alignment       =   1  'Right Justify
            Height          =   315
            Left            =   15360
            Locked          =   -1  'True
            MaxLength       =   15
            RightToLeft     =   -1  'True
            TabIndex        =   222
            Top             =   6120
            Visible         =   0   'False
            Width           =   1005
         End
         Begin VB.TextBox TxtValueTemp 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFC0&
            Enabled         =   0   'False
            Height          =   315
            Left            =   0
            RightToLeft     =   -1  'True
            TabIndex        =   197
            Top             =   1440
            Visible         =   0   'False
            Width           =   1515
         End
         Begin VB.CommandButton Command9 
            Caption         =   "⁄—÷ ›Ê« Ì— «·„»Ì⁄« "
            Height          =   405
            Left            =   7530
            RightToLeft     =   -1  'True
            TabIndex        =   196
            Top             =   1710
            Visible         =   0   'False
            Width           =   975
         End
         Begin VB.ComboBox CboStatus 
            Height          =   315
            ItemData        =   "Form2.frx":DF0E
            Left            =   3240
            List            =   "Form2.frx":DF10
            RightToLeft     =   -1  'True
            Style           =   2  'Dropdown List
            TabIndex        =   194
            Top             =   600
            Width           =   1515
         End
         Begin VB.TextBox TxtEmployeeID 
            Alignment       =   1  'Right Justify
            Height          =   315
            Left            =   2160
            TabIndex        =   184
            Top             =   2160
            Width           =   660
         End
         Begin C1SizerLibCtl.C1Elastic C1Elastic1 
            Height          =   1215
            Left            =   13080
            TabIndex        =   151
            TabStop         =   0   'False
            Top             =   1320
            Visible         =   0   'False
            Width           =   6375
            _cx             =   11245
            _cy             =   2143
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
            Begin VB.TextBox txtrenterName 
               Alignment       =   1  'Right Justify
               Height          =   315
               Left            =   120
               MaxLength       =   10
               RightToLeft     =   -1  'True
               TabIndex        =   162
               Top             =   840
               Width           =   4965
            End
            Begin VB.ComboBox cbointervaltype 
               Height          =   315
               ItemData        =   "Form2.frx":DF12
               Left            =   120
               List            =   "Form2.frx":DF1F
               TabIndex        =   161
               Top             =   480
               Width           =   855
            End
            Begin VB.TextBox txtinterval 
               Height          =   285
               Left            =   1080
               TabIndex        =   160
               Top             =   480
               Width           =   615
            End
            Begin VB.TextBox TxtSearch 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00FFFFFF&
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
               MaxLength       =   50
               RightToLeft     =   -1  'True
               TabIndex        =   152
               Top             =   120
               Width           =   1065
            End
            Begin MSDataListLib.DataCombo DcbIqara 
               Height          =   315
               Left            =   120
               TabIndex        =   153
               Tag             =   "⁄›Ê« Ì—ÃÏ «Œ Ì«—√”„ «·⁄ﬁ«—"
               Top             =   120
               Width           =   3915
               _ExtentX        =   6906
               _ExtentY        =   556
               _Version        =   393216
               Text            =   ""
               RightToLeft     =   -1  'True
            End
            Begin MSDataListLib.DataCombo DcbUnitNo 
               Height          =   315
               Left            =   2160
               TabIndex        =   154
               Tag             =   "⁄›Ê« Ì—ÃÏ «œŒ«· √”„ «·ÕÌ"
               Top             =   480
               Width           =   975
               _ExtentX        =   1720
               _ExtentY        =   556
               _Version        =   393216
               Text            =   ""
               RightToLeft     =   -1  'True
            End
            Begin MSDataListLib.DataCombo DcbUnitType 
               Height          =   315
               Left            =   4080
               TabIndex        =   155
               Tag             =   "⁄›Ê« Ì—ÃÏ «œŒ«· √”„ «·ÕÌ"
               Top             =   480
               Width           =   1065
               _ExtentX        =   1879
               _ExtentY        =   556
               _Version        =   393216
               Text            =   ""
               RightToLeft     =   -1  'True
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
               TabIndex        =   163
               Top             =   840
               Width           =   990
            End
            Begin VB.Label Label5 
               Caption         =   "«·„œ…"
               Height          =   255
               Left            =   1800
               TabIndex        =   159
               Top             =   480
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
               TabIndex        =   158
               Top             =   120
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
               TabIndex        =   157
               Top             =   480
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
               TabIndex        =   156
               Top             =   480
               Width           =   870
            End
         End
         Begin VB.Frame Frame4 
            Height          =   5775
            Left            =   12720
            TabIndex        =   147
            Top             =   240
            Width           =   6255
            Begin VB.TextBox TXTContNo 
               Height          =   495
               Left            =   600
               TabIndex        =   148
               Text            =   "0"
               Top             =   3360
               Visible         =   0   'False
               Width           =   375
            End
         End
         Begin VB.TextBox TxtBookNo 
            Alignment       =   1  'Right Justify
            Height          =   315
            Left            =   3240
            RightToLeft     =   -1  'True
            TabIndex        =   135
            Top             =   960
            Width           =   1515
         End
         Begin VB.TextBox TxtManulaNO 
            Alignment       =   1  'Right Justify
            Height          =   315
            Left            =   3240
            RightToLeft     =   -1  'True
            TabIndex        =   1
            Top             =   210
            Width           =   1515
         End
         Begin MSComCtl2.DTPicker XPDtbTrans 
            Height          =   315
            Left            =   5670
            TabIndex        =   0
            Top             =   240
            Width           =   1305
            _ExtentX        =   2302
            _ExtentY        =   556
            _Version        =   393216
            Format          =   275054593
            CurrentDate     =   41640
         End
         Begin VB.TextBox TxtCustCode 
            Alignment       =   1  'Right Justify
            Height          =   315
            Left            =   9360
            RightToLeft     =   -1  'True
            TabIndex        =   117
            Text            =   " "
            Top             =   1320
            Width           =   1515
         End
         Begin VB.TextBox Text2 
            Alignment       =   1  'Right Justify
            Height          =   315
            Left            =   9480
            RightToLeft     =   -1  'True
            TabIndex        =   116
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
            TabIndex        =   68
            Top             =   270
            Width           =   3735
            Begin VB.TextBox TxtTransID 
               Alignment       =   1  'Right Justify
               Height          =   315
               Left            =   60
               RightToLeft     =   -1  'True
               TabIndex        =   71
               Top             =   120
               Visible         =   0   'False
               Width           =   375
            End
            Begin VB.TextBox TxtTransSerial 
               Alignment       =   1  'Right Justify
               Height          =   375
               Left            =   1110
               RightToLeft     =   -1  'True
               TabIndex        =   70
               Top             =   570
               Width           =   1005
            End
            Begin VB.ComboBox CboTrans 
               Height          =   315
               Left            =   120
               RightToLeft     =   -1  'True
               Style           =   2  'Dropdown List
               TabIndex        =   69
               Top             =   240
               Width           =   1995
            End
            Begin ImpulseButton.ISButton CmdSearchTrans 
               Height          =   345
               Left            =   600
               TabIndex        =   72
               Top             =   570
               Width           =   495
               _ExtentX        =   873
               _ExtentY        =   609
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
               ButtonImage     =   "Form2.frx":DF32
            End
            Begin ImpulseButton.ISButton CmdOpenTrans 
               Height          =   345
               Left            =   90
               TabIndex        =   73
               Top             =   570
               Visible         =   0   'False
               Width           =   495
               _ExtentX        =   873
               _ExtentY        =   609
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
               ButtonImage     =   "Form2.frx":E2CC
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
               TabIndex        =   75
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
               TabIndex        =   74
               Top             =   300
               Width           =   1305
            End
         End
         Begin VB.ComboBox DCboCashType 
            Height          =   315
            Left            =   8640
            RightToLeft     =   -1  'True
            Style           =   2  'Dropdown List
            TabIndex        =   7
            Top             =   960
            Width           =   2265
         End
         Begin VB.TextBox XPMTxtRemarks 
            Alignment       =   1  'Right Justify
            Height          =   585
            Left            =   3930
            MultiLine       =   -1  'True
            RightToLeft     =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   20
            Top             =   4650
            Width           =   2715
         End
         Begin VB.TextBox XPTxtVal 
            Alignment       =   1  'Right Justify
            Height          =   315
            Left            =   9960
            MaxLength       =   15
            RightToLeft     =   -1  'True
            TabIndex        =   5
            Top             =   2160
            Width           =   1365
         End
         Begin VB.CheckBox ChkTrans 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "„‰ Õ”«» ›« Ê—…"
            Height          =   195
            Left            =   20040
            RightToLeft     =   -1  'True
            TabIndex        =   66
            Top             =   120
            Width           =   1575
         End
         Begin VB.ComboBox CboPaymentType 
            Height          =   315
            Left            =   8040
            RightToLeft     =   -1  'True
            Style           =   2  'Dropdown List
            TabIndex        =   6
            Top             =   2550
            Width           =   3285
         End
         Begin VB.Frame FraNote 
            BackColor       =   &H00E2E9E9&
            Height          =   2445
            Left            =   7920
            RightToLeft     =   -1  'True
            TabIndex        =   60
            Top             =   2760
            Width           =   4635
            Begin VB.TextBox TxtAccount 
               Alignment       =   1  'Right Justify
               Height          =   315
               Left            =   2790
               RightToLeft     =   -1  'True
               TabIndex        =   221
               Top             =   1920
               Width           =   585
            End
            Begin VB.TextBox TXTBankName 
               Alignment       =   1  'Right Justify
               Height          =   315
               Left            =   120
               RightToLeft     =   -1  'True
               TabIndex        =   12
               Top             =   480
               Visible         =   0   'False
               Width           =   3255
            End
            Begin VB.TextBox TxtChequeNumber 
               Alignment       =   1  'Right Justify
               Height          =   315
               Left            =   120
               RightToLeft     =   -1  'True
               TabIndex        =   13
               Top             =   810
               Width           =   3255
            End
            Begin MSComCtl2.DTPicker DtpChequeDueDate 
               Height          =   315
               Left            =   120
               TabIndex        =   14
               Top             =   1140
               Width           =   3255
               _ExtentX        =   5741
               _ExtentY        =   556
               _Version        =   393216
               Format          =   275054593
               CurrentDate     =   39614
            End
            Begin MSDataListLib.DataCombo DcboBankName 
               Height          =   315
               Left            =   120
               TabIndex        =   61
               Top             =   480
               Width           =   3255
               _ExtentX        =   5741
               _ExtentY        =   556
               _Version        =   393216
               Text            =   ""
               RightToLeft     =   -1  'True
            End
            Begin MSDataListLib.DataCombo DcboBox 
               Height          =   315
               Left            =   120
               TabIndex        =   11
               Top             =   120
               Width           =   3255
               _ExtentX        =   5741
               _ExtentY        =   556
               _Version        =   393216
               Text            =   ""
               RightToLeft     =   -1  'True
            End
            Begin MSDataListLib.DataCombo DcChequeBox 
               Height          =   315
               Left            =   120
               TabIndex        =   15
               Top             =   1560
               Width           =   3255
               _ExtentX        =   5741
               _ExtentY        =   556
               _Version        =   393216
               Text            =   ""
               RightToLeft     =   -1  'True
            End
            Begin MSDataListLib.DataCombo DcbAccount 
               Height          =   315
               Left            =   120
               TabIndex        =   219
               Top             =   1920
               Width           =   2655
               _ExtentX        =   4683
               _ExtentY        =   556
               _Version        =   393216
               Text            =   ""
               RightToLeft     =   -1  'True
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               BackStyle       =   0  'Transparent
               Caption         =   "«·Õ”«»"
               Height          =   285
               Index           =   64
               Left            =   3180
               RightToLeft     =   -1  'True
               TabIndex        =   220
               Top             =   1920
               Width           =   1215
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "Õ«›Ÿ… «·‘Ìﬂ« "
               Height          =   285
               Index           =   43
               Left            =   3300
               RightToLeft     =   -1  'True
               TabIndex        =   118
               Top             =   1560
               Width           =   1215
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "«”„ «·Œ“‰…"
               Height          =   285
               Index           =   9
               Left            =   3180
               RightToLeft     =   -1  'True
               TabIndex        =   65
               Top             =   180
               Width           =   1215
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "«”„ «·»‰ﬂ"
               Height          =   285
               Index           =   15
               Left            =   3180
               RightToLeft     =   -1  'True
               TabIndex        =   64
               Top             =   510
               Width           =   1215
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "—ﬁ„ «·‘Ìﬂ"
               Height          =   285
               Index           =   16
               Left            =   3180
               RightToLeft     =   -1  'True
               TabIndex        =   63
               Top             =   810
               Width           =   1215
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   " «—ÌŒ «·≈” Õﬁ«ﬁ"
               Height          =   285
               Index           =   17
               Left            =   3300
               RightToLeft     =   -1  'True
               TabIndex        =   62
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
            TabIndex        =   50
            Top             =   5640
            Width           =   8655
            Begin VB.TextBox TxtNoteSerial 
               Alignment       =   1  'Right Justify
               Enabled         =   0   'False
               Height          =   315
               Left            =   5280
               RightToLeft     =   -1  'True
               TabIndex        =   51
               Top             =   200
               Width           =   1875
            End
            Begin MSDataListLib.DataCombo DcboDebitSide 
               Height          =   315
               Left            =   90
               TabIndex        =   52
               Top             =   180
               Width           =   3975
               _ExtentX        =   7011
               _ExtentY        =   556
               _Version        =   393216
               Enabled         =   0   'False
               Text            =   ""
               RightToLeft     =   -1  'True
            End
            Begin MSDataListLib.DataCombo DcboCreditSide 
               Height          =   315
               Left            =   90
               TabIndex        =   53
               Top             =   510
               Width           =   3975
               _ExtentX        =   7011
               _ExtentY        =   556
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
               TabIndex        =   59
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
               TabIndex        =   58
               Top             =   510
               Width           =   885
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "—ﬁ„ «·ﬁÌœ:"
               Height          =   315
               Index           =   30
               Left            =   7530
               RightToLeft     =   -1  'True
               TabIndex        =   57
               Top             =   210
               Width           =   975
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "—ﬁ„ «·› —… :"
               Height          =   315
               Index           =   29
               Left            =   7530
               RightToLeft     =   -1  'True
               TabIndex        =   56
               Top             =   540
               Width           =   975
            End
            Begin VB.Label LblDevID 
               Alignment       =   2  'Center
               BackColor       =   &H00E2E9E9&
               Height          =   285
               Left            =   3870
               RightToLeft     =   -1  'True
               TabIndex        =   55
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
               TabIndex        =   54
               Top             =   510
               Width           =   1485
            End
         End
         Begin VB.TextBox XPTxtBillID 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   21840
            RightToLeft     =   -1  'True
            TabIndex        =   48
            Top             =   930
            Visible         =   0   'False
            Width           =   975
         End
         Begin VB.Frame Frame1 
            BackColor       =   &H00C0FFFF&
            Caption         =   "ŒÌ«—« "
            BeginProperty Font 
               Name            =   "MS Sans Serif"
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
            TabIndex        =   41
            Top             =   90
            Width           =   3135
            Begin VB.OptionButton Option7 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00C0FFFF&
               Caption         =   "„‘«—Ì⁄ ”«»ﬁ…"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   178
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   195
               Left            =   600
               RightToLeft     =   -1  'True
               TabIndex        =   133
               Top             =   960
               Width           =   2295
            End
            Begin VB.OptionButton Option3 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00C0FFFF&
               Caption         =   "œ›⁄Â „ﬁœ„Â"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
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
               TabIndex        =   45
               Top             =   240
               Width           =   1695
            End
            Begin VB.OptionButton Option1 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00C0FFFF&
               Caption         =   "FIFO"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
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
               TabIndex        =   44
               Top             =   480
               Width           =   1335
            End
            Begin VB.OptionButton Option2 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00C0FFFF&
               Caption         =   " ÕœÌœ ›Ê« Ì—"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
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
               TabIndex        =   43
               Top             =   720
               Width           =   2055
            End
            Begin VB.OptionButton Option6 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00FFFFFF&
               Caption         =   " ÕœÌœ „” Œ·’« "
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
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
               TabIndex        =   42
               Top             =   1560
               Value           =   -1  'True
               Width           =   2055
            End
            Begin ALLButtonS.ALLButton ALLButton3 
               Height          =   255
               Left            =   120
               TabIndex        =   46
               Top             =   720
               Width           =   615
               _ExtentX        =   1085
               _ExtentY        =   450
               BTYPE           =   3
               TX              =   " ÕœÌœ"
               ENAB            =   0   'False
               BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "MS Sans Serif"
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
               MICON           =   "Form2.frx":E666
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
               TabIndex        =   47
               Top             =   1320
               Width           =   855
               _ExtentX        =   1508
               _ExtentY        =   450
               BTYPE           =   3
               TX              =   " ÕœÌœ"
               ENAB            =   0   'False
               BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "MS Sans Serif"
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
               MICON           =   "Form2.frx":E682
               UMCOL           =   -1  'True
               SOFT            =   0   'False
               PICPOS          =   0
               NGREY           =   0   'False
               FX              =   0
               HAND            =   0   'False
               CHECK           =   0   'False
               VALUE           =   -1  'True
            End
            Begin ALLButtonS.ALLButton ALLButton6 
               Height          =   255
               Left            =   120
               TabIndex        =   199
               Top             =   480
               Width           =   615
               _ExtentX        =   1085
               _ExtentY        =   450
               BTYPE           =   3
               TX              =   "⁄—÷"
               ENAB            =   0   'False
               BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "MS Sans Serif"
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
               MICON           =   "Form2.frx":E69E
               UMCOL           =   -1  'True
               SOFT            =   0   'False
               PICPOS          =   0
               NGREY           =   0   'False
               FX              =   0
               HAND            =   0   'False
               CHECK           =   0   'False
               VALUE           =   0   'False
            End
            Begin Dynamic_Byte.NourHijriCal Txt_DateHigri 
               Height          =   255
               Left            =   0
               TabIndex        =   227
               Top             =   0
               Visible         =   0   'False
               Width           =   1485
               _ExtentX        =   2619
               _ExtentY        =   450
            End
         End
         Begin VB.TextBox txtAdv_payment_value 
            Alignment       =   1  'Right Justify
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   3960
            MaxLength       =   10
            RightToLeft     =   -1  'True
            TabIndex        =   40
            Top             =   2535
            Visible         =   0   'False
            Width           =   2685
         End
         Begin VB.TextBox Text1 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   21960
            RightToLeft     =   -1  'True
            TabIndex        =   39
            Text            =   "Text1"
            Top             =   690
            Visible         =   0   'False
            Width           =   735
         End
         Begin VB.TextBox TxtNoteSerial1 
            Alignment       =   1  'Right Justify
            Height          =   315
            Left            =   9480
            RightToLeft     =   -1  'True
            TabIndex        =   38
            Top             =   210
            Width           =   1395
         End
         Begin VB.Frame Frame2 
            BackColor       =   &H00E2E9E9&
            Caption         =   "›Ì Õ«·… «·„‘«—Ì⁄"
            Enabled         =   0   'False
            Height          =   495
            Left            =   8520
            RightToLeft     =   -1  'True
            TabIndex        =   37
            Top             =   1650
            Visible         =   0   'False
            Width           =   3975
            Begin VB.OptionButton Option4 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "⁄„Ì· ‰Â«∆Ì"
               Height          =   195
               Left            =   2520
               RightToLeft     =   -1  'True
               TabIndex        =   9
               Top             =   240
               Value           =   -1  'True
               Width           =   1215
            End
            Begin VB.OptionButton Option5 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "„ﬁ«Ê· »«ÿ‰"
               Height          =   195
               Left            =   840
               RightToLeft     =   -1  'True
               TabIndex        =   10
               Top             =   240
               Width           =   1455
            End
         End
         Begin VB.TextBox txtperson 
            Alignment       =   1  'Right Justify
            Height          =   315
            Left            =   3960
            RightToLeft     =   -1  'True
            TabIndex        =   19
            Top             =   4170
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
            TabIndex        =   49
            Top             =   2610
            Visible         =   0   'False
            Width           =   1695
            _ExtentX        =   2990
            _ExtentY        =   661
            BTYPE           =   3
            TX              =   "«ŸÂ«— «·«ﬁ”«ÿ"
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
            COLTYPE         =   1
            FOCUSR          =   -1  'True
            BCOL            =   15790320
            BCOLO           =   15790320
            FCOL            =   0
            FCOLO           =   0
            MCOL            =   12632256
            MPTR            =   1
            MICON           =   "Form2.frx":E6BA
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
            _ExtentY        =   556
            _Version        =   393216
            Text            =   ""
            RightToLeft     =   -1  'True
         End
         Begin C1SizerLibCtl.C1Elastic Ele 
            Height          =   540
            Index           =   2
            Left            =   120
            TabIndex        =   67
            TabStop         =   0   'False
            Top             =   7050
            Width           =   7995
            _cx             =   14102
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
            Left            =   240
            TabIndex        =   76
            Top             =   1320
            Width           =   4320
            _ExtentX        =   7620
            _ExtentY        =   556
            ActiveUnderline =   -1  'True
            BackStyle       =   1
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            FontBold        =   -1  'True
            FontName        =   "MS Sans Serif"
            FontSize        =   8.25
            ForeColor       =   4210688
            MousePointer    =   99
            MouseIcon       =   "Form2.frx":E6D6
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
            TabIndex        =   77
            Top             =   2850
            Visible         =   0   'False
            Width           =   1695
            _ExtentX        =   2990
            _ExtentY        =   661
            BTYPE           =   3
            TX              =   "«ŸÂ«— ”‰œ «·„œÌÊ‰Ì…"
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
            COLTYPE         =   1
            FOCUSR          =   -1  'True
            BCOL            =   15790320
            BCOLO           =   15790320
            FCOL            =   0
            FCOLO           =   0
            MCOL            =   12632256
            MPTR            =   1
            MICON           =   "Form2.frx":E838
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
            TabIndex        =   78
            Top             =   4170
            Visible         =   0   'False
            Width           =   2655
            _ExtentX        =   4683
            _ExtentY        =   556
            _Version        =   393216
            Text            =   ""
            RightToLeft     =   -1  'True
         End
         Begin MSDataListLib.DataCombo DcCostCenter 
            Bindings        =   "Form2.frx":E854
            Height          =   315
            Left            =   3960
            TabIndex        =   16
            Top             =   2850
            Width           =   2655
            _ExtentX        =   4683
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
         Begin MSDataListLib.DataCombo Dcbranch 
            Bindings        =   "Form2.frx":E869
            Height          =   315
            Left            =   5640
            TabIndex        =   2
            Top             =   600
            Width           =   3615
            _ExtentX        =   6376
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
         Begin MSDataListLib.DataCombo dcEmployee 
            Height          =   315
            Left            =   5820
            TabIndex        =   8
            Top             =   1350
            Visible         =   0   'False
            Width           =   3405
            _ExtentX        =   6006
            _ExtentY        =   556
            _Version        =   393216
            Text            =   ""
            RightToLeft     =   -1  'True
         End
         Begin MSDataListLib.DataCombo DCAccounts 
            Height          =   315
            Left            =   5640
            TabIndex        =   124
            Top             =   1350
            Visible         =   0   'False
            Width           =   3525
            _ExtentX        =   6218
            _ExtentY        =   556
            _Version        =   393216
            Text            =   ""
            RightToLeft     =   -1  'True
         End
         Begin MSDataListLib.DataCombo DcEmp 
            Bindings        =   "Form2.frx":E87E
            Height          =   315
            Left            =   0
            TabIndex        =   4
            Top             =   2160
            Width           =   2115
            _ExtentX        =   3731
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
         Begin MSDataListLib.DataCombo DCCar 
            Height          =   315
            Left            =   0
            TabIndex        =   21
            Top             =   2760
            Width           =   2835
            _ExtentX        =   5001
            _ExtentY        =   556
            _Version        =   393216
            Text            =   ""
            RightToLeft     =   -1  'True
         End
         Begin MSDataListLib.DataCombo DCDriver 
            Height          =   315
            Left            =   0
            TabIndex        =   22
            Top             =   3120
            Width           =   2835
            _ExtentX        =   5001
            _ExtentY        =   556
            _Version        =   393216
            Text            =   ""
            RightToLeft     =   -1  'True
         End
         Begin ImpulseButton.ISButton SearchCashCustomer 
            Height          =   375
            Left            =   6840
            TabIndex        =   149
            TabStop         =   0   'False
            ToolTipText     =   "«÷€ÿ ·«÷«›… ⁄„Ì· ÃœÌœ"
            Top             =   -1920
            Width           =   510
            _ExtentX        =   900
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
            ButtonImage     =   "Form2.frx":E893
            ColorButton     =   14871017
            ColorHighlight  =   16777215
            ColorShadow     =   -2147483631
            ColorOutline    =   -2147483631
            DrawFocusRectangle=   0   'False
         End
         Begin ImpulseButton.ISButton ISButton1 
            Height          =   375
            Left            =   5400
            TabIndex        =   150
            TabStop         =   0   'False
            ToolTipText     =   "«÷€ÿ ·«÷«›… ⁄„Ì· ÃœÌœ"
            Top             =   960
            Visible         =   0   'False
            Width           =   510
            _ExtentX        =   900
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
            ButtonImage     =   "Form2.frx":EC90
            ColorButton     =   14871017
            ColorHighlight  =   16777215
            ColorShadow     =   -2147483631
            ColorOutline    =   -2147483631
            DrawFocusRectangle=   0   'False
         End
         Begin MSDataListLib.DataCombo DcboRevenuesTypes 
            Height          =   315
            Left            =   5640
            TabIndex        =   87
            Top             =   1350
            Width           =   3525
            _ExtentX        =   6218
            _ExtentY        =   556
            _Version        =   393216
            Text            =   ""
            RightToLeft     =   -1  'True
         End
         Begin VB.Frame Frame5 
            BackColor       =   &H00E2E9E9&
            BorderStyle     =   0  'None
            Caption         =   "Œ’„ ⁄„Ê·…"
            Height          =   615
            Left            =   3720
            RightToLeft     =   -1  'True
            TabIndex        =   186
            Top             =   5160
            Width           =   8775
            Begin VB.TextBox Commdiscountvalue1 
               Alignment       =   1  'Right Justify
               Height          =   315
               Left            =   1200
               RightToLeft     =   -1  'True
               TabIndex        =   192
               Text            =   " "
               Top             =   240
               Visible         =   0   'False
               Width           =   675
            End
            Begin VB.TextBox Commdiscountvalue 
               Alignment       =   1  'Right Justify
               Height          =   315
               Left            =   5280
               RightToLeft     =   -1  'True
               TabIndex        =   188
               Text            =   " "
               Top             =   120
               Width           =   915
            End
            Begin VB.ComboBox commdiscounttype 
               Height          =   315
               Left            =   6360
               RightToLeft     =   -1  'True
               Style           =   2  'Dropdown List
               TabIndex        =   187
               Top             =   120
               Width           =   1185
            End
            Begin MSDataListLib.DataCombo CommdiscountAccount 
               Height          =   315
               Left            =   240
               TabIndex        =   191
               Top             =   120
               Width           =   3735
               _ExtentX        =   6588
               _ExtentY        =   556
               _Version        =   393216
               Text            =   ""
               RightToLeft     =   -1  'True
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               BackStyle       =   0  'Transparent
               Caption         =   "Õ”«» «·⁄„Ê·…"
               Height          =   285
               Index           =   57
               Left            =   3960
               RightToLeft     =   -1  'True
               TabIndex        =   190
               Top             =   120
               Width           =   1155
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               BackStyle       =   0  'Transparent
               Caption         =   "Œ’„ ⁄„Ê·…"
               Height          =   285
               Index           =   56
               Left            =   7560
               RightToLeft     =   -1  'True
               TabIndex        =   189
               Top             =   120
               Width           =   1155
            End
         End
         Begin MSDataListLib.DataCombo DCPreFix 
            Height          =   315
            Left            =   7980
            TabIndex        =   195
            Top             =   240
            Width           =   1275
            _ExtentX        =   2249
            _ExtentY        =   556
            _Version        =   393216
            BackColor       =   16777215
            Text            =   ""
            RightToLeft     =   -1  'True
         End
         Begin VB.TextBox TxtVAt2 
            Alignment       =   1  'Right Justify
            Height          =   315
            Left            =   8040
            RightToLeft     =   -1  'True
            TabIndex        =   224
            Top             =   2160
            Width           =   795
         End
         Begin VB.TextBox txtTotal 
            Alignment       =   1  'Right Justify
            Height          =   315
            Left            =   3960
            RightToLeft     =   -1  'True
            TabIndex        =   226
            Top             =   2160
            Width           =   2685
         End
         Begin VB.TextBox txtTradingContractID 
            Alignment       =   2  'Center
            Height          =   315
            Left            =   5280
            TabIndex        =   229
            TabStop         =   0   'False
            Top             =   1800
            Width           =   1245
         End
         Begin VB.TextBox TxtBillTransID 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00C0E0FF&
            Enabled         =   0   'False
            Height          =   315
            Left            =   5880
            RightToLeft     =   -1  'True
            TabIndex        =   232
            Top             =   240
            Width           =   1515
         End
         Begin VB.TextBox TxtBillTransNo 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00C0E0FF&
            Height          =   315
            Left            =   5640
            RightToLeft     =   -1  'True
            TabIndex        =   230
            Top             =   960
            Width           =   1755
         End
         Begin VB.TextBox TxtContractNo 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFC0&
            Enabled         =   0   'False
            Height          =   315
            Left            =   5880
            RightToLeft     =   -1  'True
            TabIndex        =   145
            Top             =   960
            Width           =   1515
         End
         Begin VB.TextBox TxtBillMaintID 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00C0E0FF&
            Enabled         =   0   'False
            Height          =   315
            Left            =   6360
            RightToLeft     =   -1  'True
            TabIndex        =   233
            Top             =   120
            Visible         =   0   'False
            Width           =   1515
         End
         Begin VB.TextBox TxtBillMaintNo 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00C0E0FF&
            Height          =   315
            Left            =   5640
            RightToLeft     =   -1  'True
            TabIndex        =   234
            Top             =   960
            Width           =   1755
         End
         Begin VB.Frame Frame3 
            BackColor       =   &H00E2E9E9&
            Caption         =   "„⁄·Ê„«  «·ÕÊ«·Â"
            Enabled         =   0   'False
            Height          =   975
            Left            =   3960
            RightToLeft     =   -1  'True
            TabIndex        =   119
            Top             =   3120
            Visible         =   0   'False
            Width           =   3855
            Begin VB.TextBox Text4 
               Alignment       =   1  'Right Justify
               Height          =   315
               Left            =   120
               RightToLeft     =   -1  'True
               TabIndex        =   17
               Top             =   240
               Width           =   2565
            End
            Begin MSComCtl2.DTPicker DTPicker1 
               Height          =   315
               Left            =   120
               TabIndex        =   18
               Top             =   570
               Width           =   2565
               _ExtentX        =   4524
               _ExtentY        =   556
               _Version        =   393216
               Format          =   266469377
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
               TabIndex        =   121
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
               TabIndex        =   120
               Top             =   570
               Width           =   735
            End
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
            Height          =   2235
            Left            =   0
            RightToLeft     =   -1  'True
            TabIndex        =   164
            Top             =   3480
            Width           =   3825
            Begin ImpulseAniLabel.ISAniLabel LblLinkInfo 
               Height          =   225
               Index           =   0
               Left            =   1830
               TabIndex        =   165
               Top             =   780
               Width           =   960
               _ExtentX        =   1693
               _ExtentY        =   397
               ActiveUnderline =   -1  'True
               BackStyle       =   1
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   178
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               FontBold        =   -1  'True
               FontName        =   "MS Sans Serif"
               FontSize        =   8.25
               ForeColor       =   4210688
               MousePointer    =   99
               MouseIcon       =   "Form2.frx":F08D
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
               TabIndex        =   166
               Top             =   780
               Width           =   960
               _ExtentX        =   1693
               _ExtentY        =   397
               ActiveUnderline =   -1  'True
               BackStyle       =   1
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   178
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               FontBold        =   -1  'True
               FontName        =   "MS Sans Serif"
               FontSize        =   8.25
               ForeColor       =   4210688
               MousePointer    =   99
               MouseIcon       =   "Form2.frx":F1EF
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
               Left            =   1830
               TabIndex        =   167
               Top             =   1320
               Width           =   960
               _ExtentX        =   1693
               _ExtentY        =   397
               ActiveUnderline =   -1  'True
               BackStyle       =   1
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   178
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               FontBold        =   -1  'True
               FontName        =   "MS Sans Serif"
               FontSize        =   8.25
               ForeColor       =   4210688
               MousePointer    =   99
               MouseIcon       =   "Form2.frx":F351
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
               TabIndex        =   168
               Top             =   1350
               Width           =   960
               _ExtentX        =   1693
               _ExtentY        =   397
               ActiveUnderline =   -1  'True
               BackStyle       =   1
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   178
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               FontBold        =   -1  'True
               FontName        =   "MS Sans Serif"
               FontSize        =   8.25
               ForeColor       =   4210688
               MousePointer    =   99
               MouseIcon       =   "Form2.frx":F4B3
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
               Left            =   1830
               TabIndex        =   169
               Top             =   1920
               Width           =   960
               _ExtentX        =   1693
               _ExtentY        =   397
               ActiveUnderline =   -1  'True
               BackStyle       =   1
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   178
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               FontBold        =   -1  'True
               FontName        =   "MS Sans Serif"
               FontSize        =   8.25
               ForeColor       =   4210688
               MousePointer    =   99
               MouseIcon       =   "Form2.frx":F615
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
               TabIndex        =   170
               Top             =   1920
               Width           =   960
               _ExtentX        =   1693
               _ExtentY        =   397
               ActiveUnderline =   -1  'True
               BackStyle       =   1
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   178
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               FontBold        =   -1  'True
               FontName        =   "MS Sans Serif"
               FontSize        =   8.25
               ForeColor       =   4210688
               MousePointer    =   99
               MouseIcon       =   "Form2.frx":F777
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
               Left            =   120
               TabIndex        =   171
               Top             =   420
               Width           =   1230
               _ExtentX        =   2170
               _ExtentY        =   397
               ActiveUnderline =   -1  'True
               BackStyle       =   1
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   178
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               FontBold        =   -1  'True
               FontName        =   "MS Sans Serif"
               FontSize        =   8.25
               ForeColor       =   4210688
               MousePointer    =   99
               MouseIcon       =   "Form2.frx":F8D9
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
               TabIndex        =   172
               Top             =   1080
               Width           =   1230
               _ExtentX        =   2170
               _ExtentY        =   397
               ActiveUnderline =   -1  'True
               BackStyle       =   1
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   178
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               FontBold        =   -1  'True
               FontName        =   "MS Sans Serif"
               FontSize        =   8.25
               ForeColor       =   4210688
               MousePointer    =   99
               MouseIcon       =   "Form2.frx":FA3B
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
               TabIndex        =   173
               Top             =   1680
               Width           =   1230
               _ExtentX        =   2170
               _ExtentY        =   397
               ActiveUnderline =   -1  'True
               BackStyle       =   1
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   178
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               FontBold        =   -1  'True
               FontName        =   "MS Sans Serif"
               FontSize        =   8.25
               ForeColor       =   4210688
               MousePointer    =   99
               MouseIcon       =   "Form2.frx":FB9D
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
               Left            =   1260
               RightToLeft     =   -1  'True
               TabIndex        =   183
               Top             =   1110
               Width           =   2235
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
               Left            =   1260
               RightToLeft     =   -1  'True
               TabIndex        =   182
               Top             =   1680
               Width           =   2235
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
               TabIndex        =   181
               Top             =   1350
               Width           =   705
            End
            Begin VB.Label lbl 
               Alignment       =   2  'Center
               BackColor       =   &H00E2E9E9&
               Caption         =   "«·≈”»Ê⁄ «·Õ«·Ï"
               Height          =   255
               Index           =   22
               Left            =   1320
               RightToLeft     =   -1  'True
               TabIndex        =   180
               Top             =   240
               Visible         =   0   'False
               Width           =   3495
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
               Left            =   1380
               RightToLeft     =   -1  'True
               TabIndex        =   179
               Top             =   420
               Width           =   2235
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
               TabIndex        =   178
               Top             =   1350
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
               TabIndex        =   177
               Top             =   1920
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
               TabIndex        =   176
               Top             =   1920
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
               TabIndex        =   175
               Top             =   780
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
               Index           =   28
               Left            =   1110
               RightToLeft     =   -1  'True
               TabIndex        =   174
               Top             =   780
               Width           =   675
            End
         End
         Begin VB.Frame Frame20 
            Height          =   2055
            Left            =   0
            TabIndex        =   200
            Top             =   1800
            Width           =   4455
            Begin VB.TextBox TxtPaymentValue 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00FFFFC0&
               Height          =   315
               Left            =   240
               RightToLeft     =   -1  'True
               TabIndex        =   204
               Top             =   840
               Width           =   2115
            End
            Begin VB.CommandButton Command5 
               BackColor       =   &H000000FF&
               Caption         =   "X"
               Height          =   255
               Left            =   3840
               Style           =   1  'Graphical
               TabIndex        =   210
               Top             =   120
               Width           =   375
            End
            Begin VB.TextBox TxtPercentageValue 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00FFFFC0&
               Enabled         =   0   'False
               Height          =   315
               Left            =   240
               RightToLeft     =   -1  'True
               TabIndex        =   208
               Top             =   1560
               Width           =   2115
            End
            Begin VB.TextBox TxtPercentage 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00FFFFC0&
               Enabled         =   0   'False
               Height          =   315
               Left            =   360
               RightToLeft     =   -1  'True
               TabIndex        =   206
               Top             =   1200
               Width           =   1995
            End
            Begin VB.TextBox TxtCurrentBalance 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00FFFFC0&
               Enabled         =   0   'False
               Height          =   315
               Left            =   240
               RightToLeft     =   -1  'True
               TabIndex        =   202
               Top             =   480
               Width           =   2115
            End
            Begin VB.Label Label64 
               Alignment       =   2  'Center
               Caption         =   "”Ì«”…  ⁄ÃÌ· «·œ›⁄"
               ForeColor       =   &H00000000&
               Height          =   255
               Left            =   720
               TabIndex        =   201
               Top             =   120
               Width           =   3255
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "%"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   12
                  Charset         =   178
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   285
               Index           =   63
               Left            =   120
               RightToLeft     =   -1  'True
               TabIndex        =   211
               Top             =   1200
               Width           =   195
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "ﬁÌ„… «·Œ’„"
               Height          =   285
               Index           =   62
               Left            =   2640
               RightToLeft     =   -1  'True
               TabIndex        =   209
               Top             =   1560
               Width           =   1635
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "‰”»… «·Œ’„"
               Height          =   285
               Index           =   61
               Left            =   2640
               RightToLeft     =   -1  'True
               TabIndex        =   207
               Top             =   1200
               Width           =   1635
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "„»·€ «·”œ«œ"
               Height          =   285
               Index           =   60
               Left            =   2640
               RightToLeft     =   -1  'True
               TabIndex        =   205
               Top             =   840
               Width           =   1635
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "«·—’Ìœ «·Õ«·Ì"
               Height          =   285
               Index           =   59
               Left            =   2640
               RightToLeft     =   -1  'True
               TabIndex        =   203
               Top             =   480
               Width           =   1635
            End
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "—ﬁ„ «·›« Ê—…"
            Height          =   285
            Index           =   67
            Left            =   7440
            RightToLeft     =   -1  'True
            TabIndex        =   231
            Top             =   960
            Width           =   1185
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "—ﬁ„ «·« ›«ﬁÌ…"
            Height          =   285
            Index           =   95
            Left            =   6600
            TabIndex        =   228
            Top             =   1800
            Width           =   855
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "«·«Ã„«·Ì"
            Height          =   285
            Index           =   66
            Left            =   6600
            RightToLeft     =   -1  'True
            TabIndex        =   225
            Top             =   2160
            Width           =   1035
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "«·ﬁÌ„… «·„÷«›…"
            Height          =   285
            Index           =   65
            Left            =   8610
            RightToLeft     =   -1  'True
            TabIndex        =   223
            Top             =   2175
            Width           =   1275
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "–·ﬂ „ﬁ«»·"
            Height          =   285
            Index           =   5
            Left            =   6840
            RightToLeft     =   -1  'True
            TabIndex        =   218
            Top             =   4680
            Width           =   885
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "Õ«·Â «·”‰œ"
            Height          =   285
            Index           =   58
            Left            =   4680
            RightToLeft     =   -1  'True
            TabIndex        =   193
            Top             =   600
            Width           =   915
         End
         Begin VB.Label lblinvoices 
            Height          =   375
            Left            =   120
            TabIndex        =   185
            Top             =   2040
            Visible         =   0   'False
            Width           =   1215
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
            TabIndex        =   129
            Top             =   210
            Width           =   915
         End
         Begin VB.Label Label3 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "«·„‰œÊ»"
            Height          =   255
            Left            =   2640
            RightToLeft     =   -1  'True
            TabIndex        =   128
            Top             =   2160
            Width           =   1215
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "ﬁÌ„… «·„ﬁ»Ê÷« "
            Height          =   315
            Index           =   47
            Left            =   4320
            RightToLeft     =   -1  'True
            TabIndex        =   127
            Top             =   1560
            Width           =   1155
         End
         Begin VB.Label Label2 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "«·›—⁄"
            Height          =   255
            Left            =   11130
            RightToLeft     =   -1  'True
            TabIndex        =   115
            Top             =   600
            Width           =   1395
         End
         Begin VB.Label XPTxtCurrent 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Height          =   315
            Left            =   1890
            RightToLeft     =   -1  'True
            TabIndex        =   96
            Top             =   5730
            Width           =   825
         End
         Begin VB.Label XPTxtCount 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Height          =   315
            Left            =   180
            RightToLeft     =   -1  'True
            TabIndex        =   95
            Top             =   5730
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
            TabIndex        =   94
            Top             =   5730
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
            TabIndex        =   93
            Top             =   5730
            Width           =   1065
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "‰Ê⁄ «·„ﬁ»Ê÷« "
            Height          =   285
            Index           =   6
            Left            =   11130
            RightToLeft     =   -1  'True
            TabIndex        =   92
            Top             =   990
            Width           =   1395
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "«· «—ÌŒ"
            Height          =   285
            Index           =   1
            Left            =   7050
            RightToLeft     =   -1  'True
            TabIndex        =   91
            Top             =   210
            Width           =   555
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "ﬁÌ„… «·„ﬁ»Ê÷« "
            Height          =   285
            Index           =   2
            Left            =   11250
            RightToLeft     =   -1  'True
            TabIndex        =   90
            Top             =   2190
            Width           =   1275
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "«”„ «·⁄„Ì· √Ê «·„Ê—œ"
            Height          =   315
            Index           =   3
            Left            =   11130
            RightToLeft     =   -1  'True
            TabIndex        =   89
            Top             =   1290
            Width           =   1395
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "ﬂÊœ «·”‰œ"
            Height          =   285
            Index           =   4
            Left            =   11130
            RightToLeft     =   -1  'True
            TabIndex        =   88
            Top             =   300
            Width           =   1395
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "«·—’Ìœ «·Õ«·Ï:"
            Height          =   315
            Index           =   13
            Left            =   4320
            RightToLeft     =   -1  'True
            TabIndex        =   86
            Top             =   1290
            Width           =   1155
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "ÿ—Ìﬁ… «·ﬁ»÷"
            Height          =   315
            Index           =   14
            Left            =   11250
            RightToLeft     =   -1  'True
            TabIndex        =   85
            Top             =   2520
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
            TabIndex        =   84
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
            TabIndex        =   83
            Top             =   4410
            Visible         =   0   'False
            Width           =   1455
         End
         Begin VB.Label lblsqlstring 
            Alignment       =   1  'Right Justify
            Height          =   855
            Left            =   20400
            RightToLeft     =   -1  'True
            TabIndex        =   82
            Top             =   2250
            Width           =   2895
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "œ›⁄Â „ﬁœ„Â"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
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
            Left            =   6690
            RightToLeft     =   -1  'True
            TabIndex        =   81
            Top             =   2550
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
            TabIndex        =   80
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
            TabIndex        =   79
            Top             =   4170
            Width           =   975
         End
      End
      Begin C1SizerLibCtl.C1Elastic Ele 
         Height          =   6615
         Index           =   0
         Left            =   13335
         TabIndex        =   25
         TabStop         =   0   'False
         Top             =   45
         Width           =   12600
         _cx             =   22225
         _cy             =   11668
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
         Begin VB.TextBox txtPayTotalProjVat 
            Height          =   285
            Left            =   1200
            TabIndex        =   286
            Top             =   3180
            Width           =   975
         End
         Begin VB.TextBox txtPayTotalProj 
            Height          =   285
            Left            =   2190
            TabIndex        =   285
            Top             =   3180
            Width           =   975
         End
         Begin VSFlex8Ctl.VSFlexGrid Grid1 
            Height          =   2115
            Left            =   120
            TabIndex        =   102
            Top             =   4080
            Width           =   12375
            _cx             =   21828
            _cy             =   3731
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
            Rows            =   2
            Cols            =   19
            FixedRows       =   1
            FixedCols       =   1
            RowHeightMin    =   320
            RowHeightMax    =   0
            ColWidthMin     =   0
            ColWidthMax     =   0
            ExtendLastCol   =   -1  'True
            FormatString    =   $"Form2.frx":FCFF
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
            TabIndex        =   113
            Tag             =   "Delete Row"
            Top             =   6240
            Width           =   1215
            _ExtentX        =   2143
            _ExtentY        =   661
            BTYPE           =   3
            TX              =   "Õ–› „” Œ·’"
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
            BCOL            =   0
            BCOLO           =   0
            FCOL            =   255
            FCOLO           =   255
            MCOL            =   12632256
            MPTR            =   1
            MICON           =   "Form2.frx":10002
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
            Left            =   120
            TabIndex        =   114
            Top             =   960
            Width           =   12345
            _cx             =   21775
            _cy             =   3731
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
            Rows            =   2
            Cols            =   20
            FixedRows       =   1
            FixedCols       =   1
            RowHeightMin    =   320
            RowHeightMax    =   0
            ColWidthMin     =   0
            ColWidthMax     =   0
            ExtendLastCol   =   -1  'True
            FormatString    =   $"Form2.frx":1001E
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
            TabIndex        =   101
            Top             =   3240
            Width           =   6375
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
            TabIndex        =   100
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
            TabIndex        =   97
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
            TabIndex        =   27
            Top             =   360
            Width           =   5535
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
            TabIndex        =   26
            Top             =   840
            Width           =   7575
         End
      End
      Begin C1SizerLibCtl.C1Elastic Ele 
         Height          =   6615
         Index           =   3
         Left            =   13635
         TabIndex        =   137
         TabStop         =   0   'False
         Top             =   45
         Width           =   12600
         _cx             =   22225
         _cy             =   11668
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
            BCOL            =   0
            BCOLO           =   0
            FCOL            =   255
            FCOLO           =   255
            MCOL            =   12632256
            MPTR            =   1
            MICON           =   "Form2.frx":10365
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
            Height          =   1875
            Left            =   120
            TabIndex        =   139
            Top             =   960
            Width           =   12315
            _cx             =   21722
            _cy             =   3307
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
            Rows            =   2
            Cols            =   34
            FixedRows       =   1
            FixedCols       =   1
            RowHeightMin    =   320
            RowHeightMax    =   0
            ColWidthMin     =   0
            ColWidthMax     =   0
            ExtendLastCol   =   -1  'True
            FormatString    =   $"Form2.frx":10381
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
            Height          =   2115
            Left            =   120
            TabIndex        =   144
            Top             =   3840
            Width           =   12315
            _cx             =   21722
            _cy             =   3731
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
            Rows            =   2
            Cols            =   22
            FixedRows       =   1
            FixedCols       =   1
            RowHeightMin    =   320
            RowHeightMax    =   0
            ColWidthMin     =   0
            ColWidthMax     =   0
            ExtendLastCol   =   -1  'True
            FormatString    =   $"Form2.frx":108CD
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
            Top             =   3360
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
            Top             =   3240
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
      Begin C1SizerLibCtl.C1Elastic C1Elastic6 
         Height          =   6615
         Left            =   13935
         TabIndex        =   212
         TabStop         =   0   'False
         Top             =   45
         Width           =   12600
         _cx             =   22225
         _cy             =   11668
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
         Begin VB.PictureBox Picture1 
            Height          =   3720
            Left            =   390
            ScaleHeight     =   3660
            ScaleWidth      =   3405
            TabIndex        =   284
            Top             =   2850
            Visible         =   0   'False
            Width           =   3465
         End
         Begin VSFlex8UCtl.VSFlexGrid GRID2 
            Height          =   5175
            Left            =   120
            TabIndex        =   213
            Tag             =   "1"
            Top             =   360
            Width           =   12420
            _cx             =   21907
            _cy             =   9128
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
            AllowUserResizing=   0
            SelectionMode   =   0
            GridLines       =   1
            GridLinesFixed  =   2
            GridLineWidth   =   1
            Rows            =   3
            Cols            =   8
            FixedRows       =   1
            FixedCols       =   1
            RowHeightMin    =   300
            RowHeightMax    =   0
            ColWidthMin     =   0
            ColWidthMax     =   0
            ExtendLastCol   =   -1  'True
            FormatString    =   $"Form2.frx":10C93
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
         Begin VB.Label Label1100 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFFC0&
            Caption         =   "„ÿ·Ê» «⁄ „«œ… Õ«·Ì«"
            Height          =   255
            Left            =   6510
            RightToLeft     =   -1  'True
            TabIndex        =   215
            Top             =   6120
            Visible         =   0   'False
            Width           =   3315
         End
         Begin VB.Label Label24 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFFC0&
            Caption         =   "„ÿ·Ê» «⁄ „«œ… Õ«·Ì«"
            Height          =   255
            Left            =   6510
            RightToLeft     =   -1  'True
            TabIndex        =   214
            Top             =   5760
            Width           =   3315
         End
      End
   End
   Begin ImpulseButton.ISButton Cmd 
      Height          =   375
      Index           =   0
      Left            =   11820
      TabIndex        =   103
      Top             =   7560
      Width           =   855
      _ExtentX        =   1508
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
      Index           =   1
      Left            =   10920
      TabIndex        =   104
      Top             =   7560
      Width           =   855
      _ExtentX        =   1508
      _ExtentY        =   661
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
      Height          =   375
      Index           =   2
      Left            =   10035
      TabIndex        =   105
      Top             =   7560
      Width           =   855
      _ExtentX        =   1508
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
      Left            =   9135
      TabIndex        =   106
      Top             =   7560
      Width           =   855
      _ExtentX        =   1508
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
      Left            =   8250
      TabIndex        =   107
      Top             =   7560
      Width           =   855
      _ExtentX        =   1508
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
      Index           =   6
      Left            =   1320
      TabIndex        =   108
      Top             =   7560
      Width           =   855
      _ExtentX        =   1508
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
   Begin ImpulseButton.ISButton CmdHelp 
      Height          =   375
      Left            =   2205
      TabIndex        =   109
      Top             =   7560
      Width           =   855
      _ExtentX        =   1508
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
      Index           =   5
      Left            =   7350
      TabIndex        =   110
      Top             =   7560
      Width           =   855
      _ExtentX        =   1508
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
      Index           =   7
      Left            =   6465
      TabIndex        =   111
      Top             =   7560
      Width           =   855
      _ExtentX        =   1508
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
   Begin ImpulseButton.ISButton Cmd 
      Height          =   375
      Index           =   9
      Left            =   4080
      TabIndex        =   112
      Top             =   7560
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   661
      ButtonPositionImage=   1
      Caption         =   "ÿ»«⁄… «·ﬁÌœ"
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
   Begin MSDataListLib.DataCombo DCboUserName 
      Height          =   315
      Left            =   120
      TabIndex        =   122
      Top             =   8040
      Width           =   2220
      _ExtentX        =   3916
      _ExtentY        =   556
      _Version        =   393216
      Enabled         =   0   'False
      Text            =   ""
      RightToLeft     =   -1  'True
   End
   Begin ImpulseButton.ISButton CmdAttach 
      Height          =   375
      Left            =   3120
      TabIndex        =   134
      Top             =   7560
      Width           =   915
      _ExtentX        =   1614
      _ExtentY        =   661
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
   Begin ImpulseButton.ISButton Accredit 
      Height          =   375
      Left            =   0
      TabIndex        =   216
      Top             =   7560
      Width           =   1200
      _ExtentX        =   2117
      _ExtentY        =   661
      ButtonPositionImage=   1
      Caption         =   "«—”«· ··«⁄ „«œ"
      BackColor       =   -2147483635
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColorButton     =   -2147483635
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
      Index           =   10
      Left            =   5400
      TabIndex        =   217
      Top             =   7560
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   661
      ButtonPositionImage=   1
      Caption         =   "ÿ»«⁄… 2"
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
      Height          =   435
      Index           =   11
      Left            =   4440
      TabIndex        =   283
      TabStop         =   0   'False
      Top             =   7980
      Width           =   1890
      _ExtentX        =   3334
      _ExtentY        =   767
      ButtonStyle     =   2
      Caption         =   "SMS"
      BackColor       =   14871017
      ForeColor       =   192
      FontBold        =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BackStyle       =   0
      ColorButton     =   14871017
      ColorHighlight  =   16777215
      ColorHoverText  =   192
      ColorShadow     =   -2147483631
      ColorOutline    =   -2147483631
      ButtonToggles   =   1
      DrawFocusRectangle=   0   'False
      ColorToggledText=   192
      ColorToggledHoverText=   192
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
      Left            =   7890
      RightToLeft     =   -1  'True
      TabIndex        =   126
      Top             =   7920
      Width           =   3915
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      BackStyle       =   0  'Transparent
      Caption         =   "Õ—— »Ê«”ÿ… : "
      Height          =   315
      Index           =   8
      Left            =   2040
      RightToLeft     =   -1  'True
      TabIndex        =   123
      Top             =   8040
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
      TabIndex        =   99
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
      TabIndex        =   98
      Top             =   5760
      Width           =   3735
   End
End
Attribute VB_Name = "FrmCashing"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rs As ADODB.Recordset
Dim Rs1 As ADODB.Recordset
Dim TTP As clstooltip
Dim cSearchDcbo  As clsDCboSearch
Dim Dcombos As ClsDataCombos
Dim Line1 As Double
Dim Line2 As Double
Dim Line3 As Double
Dim Line4 As Double
Dim FlgBillBuy As Boolean
Dim ScreenNameArabic As String
Dim ScreenNameEnglish As String
Dim departement_name As Integer
Dim numbering_type As Integer
Dim Balance As String
Dim balanceString As String
Dim commvalue As Double
Dim OtherInformation As New ClsGLOther
Dim dstore As Integer
Dim mClick  As Boolean
            Dim dBox As Integer
            Dim usertype As Integer
            Dim EmpID As Integer
            Dim userbranchid As Integer
        Dim isChkPaymentType As Boolean
        Dim isFormFirstRun As Boolean
Function GLByProjectInvoice(LngDevID As Double, lineno As Double, Line2 As Double)
Dim i As Integer
Dim StrSQL, newdes As String
    Dim RsDev1 As New ADODB.Recordset
                      StrSQL = "SELECT     * from dbo.DOUBLE_ENTREY_VOUCHERS Where (1 = -1)"
   RsDev1.Open StrSQL, Cn, adOpenKeyset, adLockOptimistic, adCmdText
   
 With Grid
     
    For i = .FixedRows To .rows - 1
'        If .Cell(flexcpChecked, i, .ColIndex("payed")) = flexChecked Then
If 1 = 1 Then
         If val(.TextMatrix(i, .ColIndex("ActualTotal"))) > 0 Then
            RsDev1.AddNew
            RsDev1("branch_id").value = val(Me.dcBranch.BoundText)
            RsDev1("Double_Entry_Vouchers_ID").value = LngDevID
            RsDev1("DEV_ID_Line_No").value = lineno
            RsDev1("DEV_ID_Line_No1").value = Line2
            RsDev1("Account_Code").value = Me.DcboCreditSide.BoundText
            RsDev1("NextAccount_Code").value = Me.DcboDebitSide.BoundText
            RsDev1("Value").value = val(.TextMatrix(i, .ColIndex("ActualTotal")))
            RsDev1("Credit_Or_Debit").value = 1
            If SystemOptions.PaymentIntoAccouStat = True And val(DCboCashType.ListIndex) = 5 Then
            
            RsDev1("project_id").value = val(.TextMatrix(i, .ColIndex("project_no")))
            RsDev1("projectid").value = val(.TextMatrix(i, .ColIndex("project_no")))
            End If
 
            RsDev1("Double_Entry_Vouchers_Description").value = XPMTxtRemarks.text & CHR(13) & (.TextMatrix(i, .ColIndex("Project_name"))) & CHR(13) & lblinvoices.Caption
               
               
            ' RsDev1("Double_Entry_Vouchers_Description").value = dcproject.BoundText
            RsDev1("Notes_ID").value = val(XPTxtID.text)
            RsDev1("RecordDate").value = Me.XPDtbTrans.value
            RsDev1("RecordDateH").value = ToHijriDate(XPDtbTrans.value)
            RsDev1("UserID").value = Me.DCboUserName.BoundText
            RsDev1("Account_Interval_ID").value = SystemOptions.SysCurrentAccountIntervalID
             
                 RsDev1("project_id").value = val(.TextMatrix(i, .ColIndex("project_no")))
             
             
           '   If Posted = 1 Then
           ' RsDev1("Posted").value = 1
           ' Else
           ' RsDev1("Posted").value = Null
           ' End If
            RsDev1.update
            lineno = lineno + 1
        End If
        End If
    Next i
    
  End With
End Function
Sub RetriveBillBuy(Optional CuID As Double = 0)
Dim sql As String
Dim Rs8 As ADODB.Recordset
Dim i As Integer
Set Rs8 = New ADODB.Recordset
With VSFlexGrid1
.Clear flexClearScrollable, flexClearEverything
.rows = 1
End With
sql = "Select * from ("
sql = sql & "        SELECT     TOP 100 PERCENT dbo.Transactions.Transaction_ID,'›« Ê—…' as TransTypeName,dbo.Transactions.Transaction_Date, dbo.Transactions.Transaction_Type, dbo.Transactions.NoteSerial1, "
sql = sql & "                      dbo.Transactions.ManualNO, dbo.Transactions.BranchId, dbo.TblBranchesData.branch_name, dbo.TblBranchesData.branch_namee, dbo.Transactions.CusID,"
sql = sql & "                      dbo.TblCustemers.CusName, dbo.TblCustemers.CusNamee, dbo.TblCustemers.Fullcode, dbo.Transactions.TotalPayed, dbo.Transactions.OldContID,"
sql = sql & "                      dbo.transactions.OldValue , dbo.transactions.dueDate, dbo.transactions.Vat, dbo.transactions.Transaction_NetValue"
sql = sql & " FROM         dbo.Transactions LEFT OUTER JOIN"
sql = sql & "                      dbo.TblCustemers ON dbo.Transactions.CusID = dbo.TblCustemers.CusID LEFT OUTER JOIN"
sql = sql & "                      dbo.TblBranchesData ON dbo.Transactions.BranchId = dbo.TblBranchesData.branch_id"
sql = sql & "  WHERE     (dbo.Transactions.PaymentType = 1) AND (dbo.Transactions.Transaction_Type = 21 OR"
sql = sql & "                       dbo.Transactions.Transaction_Type = 2 or dbo.Transactions.Transaction_Type = 71) AND (dbo.Transactions.TotalPayed IS NULL OR"
sql = sql & "                       dbo.Transactions.TotalPayed = 0) AND (dbo.Transactions.CusID = " & CuID & ")"

sql = sql & "  Union All"
sql = sql & "  SELECT d.Id             Transaction_ID,'  ›« Ê—… ⁄„·«¡' as TransTypeName,"
sql = sql & "         d.recordDate     Transaction_Date,"
sql = sql & "         Transaction_Type = 9999,"
sql = sql & "         Cast (d.NoteSerial1 as NVARCHAR(255)) NoteSerial1,"
sql = sql & "         ManualNo = '',"
sql = sql & "         BranchId = 0,"
sql = sql & "         branch_name = '',"
sql = sql & "         branch_namee = '',"
sql = sql & "         d.CusID,"
sql = sql & "         dbo.TblCustemers.CusName,"
sql = sql & "         dbo.TblCustemers.CusNamee,"
sql = sql & "         dbo.TblCustemers.Fullcode,"
sql = sql & "         d.TotalPayed,"
sql = sql & "         OldContID = 0,"
sql = sql & "         d.TotalValue     OldValue,"
sql = sql & "         dueDate = GETDATE(),"
sql = sql & "         d.VAT,"
sql = sql & "         d.TotalValue +  IsNull(d.VAT,0) Transaction_NetValue"
sql = sql & "  FROM   TblTravDueK   AS d"
sql = sql & "         LEFT OUTER JOIN dbo.TblCustemers"
sql = sql & "              ON  d.CusID = dbo.TblCustemers.CusID"

sql = sql & "  WHERE     "
sql = sql & "           (d.TotalPayed IS NULL OR"
sql = sql & "                       d.TotalPayed = 0) AND (d.CusID = " & CuID & ")"

sql = sql & " )T  ORDER BY DueDate ,NoteSerial1"
Rs8.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
If Rs8.RecordCount > 0 Then
VSFlexGrid1.Enabled = True


        VSFlexGrid1.Enabled = True
With VSFlexGrid1
.Clear flexClearScrollable, flexClearEverything
.rows = 1
    .rows = .rows + Rs8.RecordCount
.rows = .FixedRows + Rs8.RecordCount
Rs8.MoveFirst
For i = .FixedRows To Rs8.RecordCount
.TextMatrix(i, .ColIndex("Ser")) = i
.TextMatrix(i, .ColIndex("branch_no")) = IIf(IsNull(Rs8("BranchId").value), 0, Rs8("BranchId").value)
If SystemOptions.UserInterface = ArabicInterface Then
.TextMatrix(i, .ColIndex("branch_name")) = IIf(IsNull(Rs8("branch_name").value), 0, Rs8("branch_name").value)
Else
.TextMatrix(i, .ColIndex("branch_name")) = IIf(IsNull(Rs8("branch_namee").value), 0, Rs8("branch_namee").value)
End If

.TextMatrix(i, .ColIndex("DueDate")) = IIf(IsNull(Rs8("DueDate").value), "", Rs8("DueDate").value)
.TextMatrix(i, .ColIndex("NoteID")) = IIf(IsNull(Rs8("Transaction_ID").value), 0, Rs8("Transaction_ID").value)
.TextMatrix(i, .ColIndex("NoteDate")) = IIf(IsNull(Rs8("Transaction_Date").value), "", Rs8("Transaction_Date").value)
.TextMatrix(i, .ColIndex("NoteSerial1")) = IIf(IsNull(Rs8("NoteSerial1").value), "", Rs8("NoteSerial1").value)
.TextMatrix(i, .ColIndex("too")) = IIf(IsNull(Rs8("ManualNO").value), "", Rs8("ManualNO").value)
.TextMatrix(i, .ColIndex("Note_Value")) = val(IIf(IsNull(Rs8("Transaction_NetValue").value), IIf(IsNull(Rs8("OldValue").value), 0, Rs8("OldValue").value), Rs8("Transaction_NetValue").value))
.TextMatrix(i, .ColIndex("TransTypeName")) = IIf(IsNull(Rs8("TransTypeName").value), "", Rs8("TransTypeName").value)
.TextMatrix(i, .ColIndex("Transaction_Type")) = IIf(IsNull(Rs8("Transaction_Type").value), "", Rs8("Transaction_Type").value)

If val(.TextMatrix(i, .ColIndex("NoteID"))) <> 0 Then
.TextMatrix(i, .ColIndex("PayedValue")) = GeteBillBuy(val(.TextMatrix(i, .ColIndex("NoteID"))))
Else
.TextMatrix(i, .ColIndex("PayedValue")) = 0
End If
If .TextMatrix(i, .ColIndex("PayedValue")) < 0 Then
.TextMatrix(i, .ColIndex("PayedValue")) = val(.TextMatrix(i, .ColIndex("PayedValue"))) * -1
End If
.TextMatrix(i, .ColIndex("RemainingValue")) = val(.TextMatrix(i, .ColIndex("Note_Value"))) - val(.TextMatrix(i, .ColIndex("PayedValue")))
Rs8.MoveNext
Next i
End With
End If
End Sub
Function GeteBillBuy(Optional Transaction_ID As Double = 0) As Double
Dim sql As String
Dim Rs8 As ADODB.Recordset
Set Rs8 = New ADODB.Recordset
sql = " SELECT   SUM(PayedValue) AS Smatiobn"
sql = sql & " From dbo.TblBillBuyPayment2"
sql = sql & " Where (Transaction_ID = " & Transaction_ID & ")"
sql = sql & " GROUP BY Transaction_ID"

'salim hereeeeeee
sql = "select sum (PayedValue)  as Smatiobn"
sql = sql & " From"
sql = sql & " ("
sql = sql & "  SELECT"
 sql = sql & " PayedValue=Case"
sql = sql & "  When   transtype=1    Then   abs(PayedValue)"
sql = sql & "  Else PayedValue"
sql = sql & " End"
sql = sql & " From dbo.TblBillBuyPayment2"
'sql = sql & " Where (Transaction_ID = 33426)"
sql = sql & " Where (Transaction_ID = " & Transaction_ID & ")"
sql = sql & " )z"


Rs8.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
If Rs8.RecordCount > 0 Then
GeteBillBuy = IIf(IsNull(Rs8("Smatiobn").value), 0, Rs8("Smatiobn").value)
Else
GeteBillBuy = 0
End If
End Function

Private Sub Accredit_Click()
    Dim BeginTrans As Boolean
If val(XPTxtID.text) = 0 Then
     If SystemOptions.UserInterface = ArabicInterface Then
     MsgBox "«Õ›Ÿ «·”‰œ «Ê·«", vbCritical
     Else
     MsgBox "Save Doc First", vbCritical
     End If
      
      Exit Sub
      End If
 
 
    SendTopost Me.Name, "Notes", "NoteID", 0, val(dcBranch.BoundText), val(XPTxtID.text), TxtNoteSerial1.text
  rs.Resync
 If SystemOptions.UserInterface = ArabicInterface Then
    Accredit.Caption = " „ «·«—”«· ··«⁄ „«œ"
Else
Accredit.Caption = "Sent To approval "
End If
    Retrive (val(Me.XPTxtID.text))
End Sub

Private Sub ALLButton1_Click()

    If IsNumeric(Me.DBCboClientName.BoundText) Then
        'INSTALLMENT_DATA1.show
        'INSTALLMENT_DATA1.Adodc1.CommandType = adCmdText
        'INSTALLMENT_DATA1.Adodc1.RecordSource = "select *  FROM INSTALLMENT_DETAILS where payed=0 and cust_id =" & Me.DBCboClientName.BoundText
        'INSTALLMENT_DATA1.Adodc1.Refresh
 
        'INSTALLMENT_DATA1.id.text = Me.DBCboClientName.BoundText
        'INSTALLMENT_DATA1.lblcustid = Me.DBCboClientName.BoundText
        'INSTALLMENT_DATA1.TxtName.text = Me.DBCboClientName.text
    End If

End Sub

Private Sub ALLButton6_Click()
Frame12.Visible = True
Frame12.Enabled = True
VSFlexGrid1.Enabled = True
End Sub

Private Sub Check1_Click()
    Dim i As Integer

    If Check1.value = vbChecked Then

        With Me.VSFlexGrid1
 
            For i = .FixedRows To .rows - 1
        
                .TextMatrix(i, .ColIndex("payed")) = True
            Next i

        End With

    Else

        With Me.VSFlexGrid1

            For i = .FixedRows To .rows - 1
        
                .TextMatrix(i, .ColIndex("payed")) = False
            Next i

        End With

    End If
    RelineBuy
End Sub
Public Sub RetriveBillBuyData(Optional Lngid As Long = 0)
    Dim RsDetails As ADODB.Recordset
    Dim rs As ADODB.Recordset
    Dim i As Integer
    Dim StrSQL As String


   ' On Error GoTo ErrTrap
    Set RsDetails = New ADODB.Recordset
  StrSQL = "   SELECT     dbo.TblBranchesData.branch_name, dbo.TblBranchesData.branch_namee, dbo.TblNotesBillBuyPayment2.*"
  StrSQL = StrSQL & "  FROM         dbo.TblNotesBillBuyPayment2 LEFT OUTER JOIN"
  StrSQL = StrSQL & "                     dbo.TblBranchesData ON dbo.TblNotesBillBuyPayment2.branch_no = dbo.TblBranchesData.branch_id"
  StrSQL = StrSQL & "  Where (dbo.TblNotesBillBuyPayment2.NoteID1 = " & val(XPTxtID.text) & ")"
    
    RsDetails.Open StrSQL, Cn, adOpenKeyset, adLockOptimistic, adCmdText
    With VSFlexGrid1
    .Clear flexClearScrollable, flexClearEverything
    .rows = .FixedRows

    If Not (RsDetails.BOF Or RsDetails.EOF) Then
      '  Fra(2).Visible = True
      '               lbl(47).Visible = True
      '  TxtAdvance.Visible = True
        RsDetails.MoveFirst
        .rows = .FixedRows + RsDetails.RecordCount

        For i = .FixedRows To RsDetails.RecordCount
        .TextMatrix(i, .ColIndex("Ser")) = i

            .TextMatrix(i, .ColIndex("branch_no")) = IIf(IsNull(RsDetails("branch_no").value), 0, RsDetails("branch_no").value)
            If SystemOptions.UserInterface = ArabicInterface Then
            .TextMatrix(i, .ColIndex("branch_name")) = IIf(IsNull(RsDetails("branch_name").value), "", RsDetails("branch_name").value)
            Else
            .TextMatrix(i, .ColIndex("branch_name")) = IIf(IsNull(RsDetails("branch_namee").value), 0, RsDetails("branch_namee").value)
            End If
            .TextMatrix(i, .ColIndex("NoteID")) = IIf(IsNull(RsDetails("NoteID").value), 0, RsDetails("NoteID").value)
            .TextMatrix(i, .ColIndex("NoteSerial1")) = IIf(IsNull(RsDetails("NoteSerial1").value), 0, RsDetails("NoteSerial1").value)
            .TextMatrix(i, .ColIndex("Note_Value")) = IIf(IsNull(RsDetails("Note_Value").value), 0, RsDetails("Note_Value").value)
            .TextMatrix(i, .ColIndex("PayedValue")) = IIf(IsNull(RsDetails("PayedValue").value), 0, RsDetails("PayedValue").value)
            .TextMatrix(i, .ColIndex("TransPayedValue")) = IIf(IsNull(RsDetails("TransPayedValue").value), 0, RsDetails("TransPayedValue").value)
            .TextMatrix(i, .ColIndex("too")) = IIf(IsNull(RsDetails("too").value), "", RsDetails("too").value)
            .TextMatrix(i, .ColIndex("NetValue")) = IIf(IsNull(RsDetails("NetValue").value), 0, RsDetails("NetValue").value)
            .TextMatrix(i, .ColIndex("RemainingValue")) = IIf(IsNull(RsDetails("RemainingValue").value), 0, RsDetails("RemainingValue").value)
            
           ' .TextMatrix(i, .ColIndex("PartValue")) = Round(RsDetails("PartValue").value, 2)
             .TextMatrix(i, .ColIndex("DueDate")) = IIf(IsNull(((RsDetails("DueDate").value))), " ", ((RsDetails("DueDate").value)))
           ' .TextMatrix(i, .ColIndex("NoteDate")) = DisplayDate(CDate(RsDetails("NoteDate").value))
            .TextMatrix(i, .ColIndex("NoteDate")) = IIf(IsNull(((RsDetails("NoteDate").value))), "", ((RsDetails("NoteDate").value)))
            .cell(flexcpChecked, i, .ColIndex("payed")) = flexChecked
            RsDetails.MoveNext
        Next i
        

    End If
End With
    RsDetails.Close
    Set RsDetails = Nothing
    Set rs = Nothing
    Exit Sub
ErrTrap:
End Sub
Private Sub ALLButton2_Click()

    If IsNumeric(Me.DBCboClientName.BoundText) Then
        'sanad_dean.show
        'sanad_dean.LblID = DBCboClientName.BoundText
        'sanad_dean.LblName = DBCboClientName.text
        'sanad_dean.lblaccountcode.Caption = txtaccount.text
        'sanad_dean.Adodc1.CommandType = adCmdText
        'sanad_dean.Adodc1.RecordSource = "select*  FROM sanad_dean where cust_id=" & DBCboClientName.BoundText
        'sanad_dean.Adodc1.Refresh
        'sanad_dean.ALLButton1.Visible = False
        'sanad_dean.ALLButton1.Visible = False

        'sanad_dean.Adodc2.CommandType = adCmdText
        'sanad_dean.Adodc2.RecordSource = "select *  FROM member_child where cust_id=" & DBCboClientName.BoundText
        'sanad_dean.Adodc2.Refresh
    End If

End Sub

Private Sub ALLButton3_Click()
 '   lblsqlstring.Caption = ""
 '   FrmPaymentTime1.show
 '   FrmPaymentTime1.lblcusid = val(DBCboClientName.BoundText)
 '   FrmPaymentTime1.LblValue = val(XPTxtVal.Text)
 If val(DCboCashType.ListIndex) = 0 Then
If Me.Option2.value = True Then
BillCustomer
Else
Frame12.Visible = False
End If
End If
End Sub

Public Sub FillGridWithDataContract(NoteSerial1 As String)

    'On Error GoTo ErrTrap

    Dim i As Integer
    Dim X As Integer
    Dim rs As ADODB.Recordset
 
    Dim ActualTotal As Double
    Dim Result As Double
    Dim resultpercentage As Double
    Dim sql As String

    Grid3.Clear flexClearScrollable, flexClearEverything
    Grid3.rows = 1
          
    Grid4.Clear flexClearScrollable, flexClearEverything
    Grid4.rows = 1

    If DCboCashType.ListIndex <> 8 Then Exit Sub
 
    lbl(38).Caption = DBCboClientName.text
    lbl(41).Caption = DBCboClientName.text
    '
     
    sql = "SELECT    dbo.TblContractInstallments.DES , dbo.TblContractInstallments.OldValueDate ,dbo.TblContractInstallments.OldValueDateH  ,  dbo.TblContractInstallments.OldValue ,  dbo.TblContractInstallments.InstallNo, dbo.TblContractInstallments.Installdate, dbo.TblContractInstallments.InstalldateH, dbo.TblContract.ownerid, "
sql = sql & " dbo.TblContractInstallments.RentValue , dbo.TblContractInstallments.Commissions, dbo.TblContractInstallments.Insurance, dbo.TblContractInstallments.Water, dbo.TblContractInstallments.Electric, dbo.TblContractInstallments.TelandNet, dbo.TblContractInstallments.RentValuePayed, dbo.TblContractInstallments.CommissionsPayed, dbo.TblContractInstallments.InsurancePayed, WaterPayed, dbo.TblContractInstallments.ElectricPayed, dbo.TblContractInstallments.TelandNetPayed,"
sql = sql & "  dbo.TblContract.CusID, dbo.TblContractInstallments.installValue, dbo.TblContractInstallments.Status, dbo.TblContractInstallments.ContNo,"
sql = sql & "   dbo.TblContractInstallments.id, dbo.TblContract.ContDate, dbo.TblAqarDetai.unitno, dbo.TblAqarDetai.unittype, dbo.TblAkarUnit.name AS unitname,"
sql = sql & "   dbo.TblAkarUnit.namee AS unitnamee, dbo.TblAqar.Aqarid, dbo.TblAqar.aqarNo, dbo.TblAqar.CountryID, dbo.TblAqar.aqarname, dbo.TblAqar.streetname,"
sql = sql & "     dbo.TblCustemers.CusName AS owner, dbo.TblCustemers.CusNamee AS ownere, dbo.TblCountriesGovernments.GovernmentName AS Country,"
sql = sql & "      dbo.TblCountriesGovernmentsCities.CityName AS hey, dbo.TblContract.StrDate, dbo.TblContract.EndDate, dbo.TblContract.MeterValue, dbo.TblContract.MeterCount,"
sql = sql & "     dbo.TblContract.TotalContract, dbo.TblContract.PayAmini, dbo.TblContract.CommiValue, dbo.TblContract.InsuranceValue, dbo.TblContract.Water AS totalWater,"
sql = sql & "      dbo.TblContract.Electricity AS totalElectricity , dbo.TblContract.Enternet AS totalEnternet, dbo.TblContract.Phone AS totalPhone , dbo.TblContract.IncresYearValue, dbo.TblContract.IncresYearRate,"
sql = sql & "      dbo.TblContract.PaymentCount, dbo.TblContract.FristPaymentDate, dbo.TblContract.PeriodsID, dbo.TblContract.Periods, dbo.TblContract.Furnishing,"
sql = sql & "       dbo.TblContract.Remarks, dbo.TblContract.RecorddateH, dbo.TblContract.FromdateH, dbo.TblContract.TodateH, dbo.TblContract.FirstInstallDateH,"
sql = sql & "      dbo.TblContract.Branch_NO, dbo.TblContract.NewOrOpeneing, dbo.TblContract.OthersRules, dbo.TblContract.NoteID, dbo.TblContract.NoteSerial,"
 sql = sql & "                       dbo.TblContract.NoteSerial1, dbo.TblContractInstallments.NoteSerial1 AS NoteSerial1Install, dbo.TblContractInstallments.NoteSerial AS NoteSerialInstall"
sql = sql & "  FROM         dbo.TblCountriesGovernmentsCities INNER JOIN"
sql = sql & "   dbo.TblAqar INNER JOIN"
sql = sql & "     dbo.TblAqarDetai ON dbo.TblAqar.Aqarid = dbo.TblAqarDetai.Aqarid INNER JOIN"
sql = sql & "     dbo.TblCountriesGovernments ON dbo.TblAqar.cityid = dbo.TblCountriesGovernments.GovernmentID ON"
sql = sql & "     dbo.TblCountriesGovernmentsCities.CityID = dbo.TblAqar.heyid LEFT OUTER JOIN"
sql = sql & "      dbo.TblAkarUnit ON dbo.TblAqarDetai.unittype = dbo.TblAkarUnit.id RIGHT OUTER JOIN"
sql = sql & "      dbo.TblCustemers INNER JOIN"
sql = sql & "      dbo.TblContractInstallments INNER JOIN"
sql = sql & "      dbo.TblContract ON dbo.TblContractInstallments.ContNo = dbo.TblContract.ContNo ON dbo.TblCustemers.CusID = dbo.TblContract.ownerid ON"
sql = sql & "    dbo.TblAqarDetai.id = dbo.TblContract.unitno"
sql = sql & "   WHERE     ( (dbo.TblContractInstallments.Status is null  or dbo.TblContractInstallments.Status=0)  and  dbo.TblContract.ContNo =" & val(TxtContNo.text) & ")"
    Set rs = New ADODB.Recordset
    rs.Open sql, Cn, adOpenStatic, adLockReadOnly, adCmdText

    If rs.RecordCount = 0 Then
 
        Exit Sub
    End If

    i = 0

    With Me.Grid3
        .rows = 1
        .Clear flexClearScrollable
  
        rs.MoveFirst
DBCboClientName.BoundText = IIf(IsNull(rs.Fields("CusID").value), "", rs.Fields("CusID").value)
        For X = 1 To rs.RecordCount
       
            ActualTotal = getinsttPayedTocontract(val(rs.Fields("id").value))
            Result = val(rs.Fields("installValue").value) - ActualTotal
            resultpercentage = Round((ActualTotal / val(rs.Fields("installValue").value)) * 100, 2)
 
            If val(rs.Fields("installValue").value) > ActualTotal Then
                i = i + 1
                .rows = .rows + 1
                .TextMatrix(i, .ColIndex("Ser")) = i
                .TextMatrix(i, .ColIndex("id")) = IIf(IsNull(rs.Fields("id").value), "", rs.Fields("id").value)
            
                '                             .TextMatrix(I, .ColIndex("bill_id")) = IIf(IsNull(rs.Fields("bill_id").value), _
                                              "", rs.Fields("bill_id").value)
            
                .TextMatrix(i, .ColIndex("Installdate")) = IIf(IsNull(rs.Fields("Installdate").value), "", rs.Fields("Installdate").value)
                .TextMatrix(i, .ColIndex("Installdateh")) = IIf(IsNull(rs.Fields("Installdateh").value), "", rs.Fields("Installdateh").value)
              
             Dim datedifferent As Integer
             datedifferent = DateDiff("d", .TextMatrix(i, .ColIndex("Installdate")), XPDtbTrans.value)
             
             If datedifferent <= 30 Then
                 .TextMatrix(i, .ColIndex("CommisionTypesid")) = 1
                  .TextMatrix(i, .ColIndex("CommisionTypes")) = " ”ÊÌﬁ"
             Else
               .TextMatrix(i, .ColIndex("CommisionTypesid")) = 2
                  .TextMatrix(i, .ColIndex("CommisionTypes")) = " Õ’Ì·"
             End If
             
              
               .TextMatrix(i, .ColIndex("total")) = IIf(IsNull(rs.Fields("installValue").value), 0, rs.Fields("installValue").value)
               
               
               
 
              .TextMatrix(i, .ColIndex("OldValueDate")) = IIf(IsNull(rs.Fields("OldValueDate").value), "", rs.Fields("OldValueDate").value)
                .TextMatrix(i, .ColIndex("OldValueDateH")) = IIf(IsNull(rs.Fields("OldValueDateH").value), "", rs.Fields("OldValueDateH").value)
              .TextMatrix(i, .ColIndex("des")) = IIf(IsNull(rs.Fields("des").value), "", rs.Fields("des").value)
               .TextMatrix(i, .ColIndex("OldValue")) = IIf(IsNull(rs.Fields("OldValue").value), 0, rs.Fields("OldValue").value)
                
                
               .TextMatrix(i, .ColIndex("InstallNo")) = IIf(IsNull(rs.Fields("InstallNo").value), "", rs.Fields("InstallNo").value)
                
                
                .TextMatrix(i, .ColIndex("ActualTotal")) = ActualTotal
                .TextMatrix(i, .ColIndex("ResultPercentage")) = resultpercentage
                .TextMatrix(i, .ColIndex("Result")) = Result
    
 
     'RentValue,Commissions,Insurance,Water,Electric,TelandNet
     'RentValuePayed,CommissionsPayed,InsurancePayed,WaterPayed,ElectricPayed,TelandNetPayed
     .TextMatrix(i, .ColIndex("RentValue")) = (IIf(IsNull(rs.Fields("RentValue").value), 0, rs.Fields("RentValue").value))
    .TextMatrix(i, .ColIndex("Commissions")) = (IIf(IsNull(rs.Fields("Commissions").value), 0, rs.Fields("Commissions").value))
    .TextMatrix(i, .ColIndex("Insurance")) = (IIf(IsNull(rs.Fields("Insurance").value), 0, rs.Fields("Insurance").value))
    .TextMatrix(i, .ColIndex("Water")) = (IIf(IsNull(rs.Fields("Water").value), 0, rs.Fields("Water").value))
    .TextMatrix(i, .ColIndex("Electric")) = (IIf(IsNull(rs.Fields("Electric").value), 0, rs.Fields("Electric").value))
    .TextMatrix(i, .ColIndex("TelandNet")) = (IIf(IsNull(rs.Fields("TelandNet").value), 0, rs.Fields("TelandNet").value))
     
    
    .TextMatrix(i, .ColIndex("RentValuePayed")) = (IIf(IsNull(rs.Fields("RentValuePayed").value), 0, rs.Fields("RentValuePayed").value))
    .TextMatrix(i, .ColIndex("CommissionsPayed")) = (IIf(IsNull(rs.Fields("CommissionsPayed").value), 0, rs.Fields("CommissionsPayed").value))
    .TextMatrix(i, .ColIndex("InsurancePayed")) = (IIf(IsNull(rs.Fields("InsurancePayed").value), 0, rs.Fields("InsurancePayed").value))
    .TextMatrix(i, .ColIndex("WaterPayed")) = (IIf(IsNull(rs.Fields("WaterPayed").value), 0, rs.Fields("WaterPayed").value))
    .TextMatrix(i, .ColIndex("ElectricPayed")) = (IIf(IsNull(rs.Fields("ElectricPayed").value), 0, rs.Fields("ElectricPayed").value))
    .TextMatrix(i, .ColIndex("TelandNetPayed")) = (IIf(IsNull(rs.Fields("TelandNetPayed").value), 0, rs.Fields("TelandNetPayed").value))

            End If

            rs.MoveNext
        Next

        rs.Close
 
        .RowHeight(-1) = 300
    End With

    If TxtNoteSerial = "" Then

        Exit Sub
    End If
'  rs("NoteID").value = val(XPTxtID.text)
    sql = "SELECT  * FROM     ContracttBillInstallmentsDone     where NoteID =" & val(XPTxtID.text)
 
   ' rs.Open sql, Cn, adOpenStatic, adLockReadOnly, adCmdText


    'If rs.RecordCount = 0 Then
 
    '    Exit Sub
    'End If
 
      'rs.Open sql, Cn, adOpenStatic, adLockReadOnly, adCmdText

  sql = "SELECT  * FROM     ContracttBillInstallmentsDone     where NoteID =" & val(XPTxtID.text)
 rs.Open sql, Cn, adOpenStatic, adLockReadOnly, adCmdText
    If rs.RecordCount = 0 Then
 
        Exit Sub
    End If
    With Me.Grid4
        .rows = 1
        .rows = .rows + rs.RecordCount
        .Clear flexClearScrollable
  
        rs.MoveFirst

        For i = 1 To .rows - 1
 
            .TextMatrix(i, .ColIndex("Ser")) = i
            .TextMatrix(i, .ColIndex("id")) = IIf(IsNull(rs.Fields("id").value), "", rs.Fields("id").value)
        
            .TextMatrix(i, .ColIndex("Installdate")) = IIf(IsNull(rs.Fields("RecordDate").value), "", rs.Fields("RecordDate").value)
              .TextMatrix(i, .ColIndex("Installdateh")) = IIf(IsNull(rs.Fields("RecordDateh").value), "", rs.Fields("RecordDateh").value)
              
            .TextMatrix(i, .ColIndex("InstallNo")) = IIf(IsNull(rs.Fields("InstallNo").value), "", rs.Fields("InstallNo").value)
 
            .TextMatrix(i, .ColIndex("total")) = IIf(IsNull(rs.Fields("total").value), "", rs.Fields("total").value)
            
            .TextMatrix(i, .ColIndex("ActualTotal")) = IIf(IsNull(rs.Fields("value").value), "", rs.Fields("value").value)
            Result = val(.TextMatrix(i, .ColIndex("total"))) - getinsttPayedTocontract(val(rs.Fields("istallid").value)) '
            resultpercentage = val(rs.Fields("value").value) / val(.TextMatrix(i, .ColIndex("total"))) * 100
            If resultpercentage > 100 Then resultpercentage = 100
            .TextMatrix(i, .ColIndex("ResultPercentage")) = Round(resultpercentage, 2)
            If Result < 0 Then Result = 0
            .TextMatrix(i, .ColIndex("Result")) = Result
          .TextMatrix(i, .ColIndex("RentValuePayed")) = (IIf(IsNull(rs.Fields("RentValuePayed").value), 0, rs.Fields("RentValuePayed").value))
    .TextMatrix(i, .ColIndex("CommissionsPayed")) = (IIf(IsNull(rs.Fields("CommissionsPayed").value), 0, rs.Fields("CommissionsPayed").value))
    .TextMatrix(i, .ColIndex("InsurancePayed")) = (IIf(IsNull(rs.Fields("InsurancePayed").value), 0, rs.Fields("InsurancePayed").value))
    .TextMatrix(i, .ColIndex("WaterPayed")) = (IIf(IsNull(rs.Fields("WaterPayed").value), 0, rs.Fields("WaterPayed").value))
    .TextMatrix(i, .ColIndex("ElectricPayed")) = (IIf(IsNull(rs.Fields("ElectricPayed").value), 0, rs.Fields("ElectricPayed").value))
    .TextMatrix(i, .ColIndex("TelandNetPayed")) = (IIf(IsNull(rs.Fields("TelandNetPayed").value), 0, rs.Fields("TelandNetPayed").value))


            rs.MoveNext
        Next

        rs.Close
 
        .RowHeight(-1) = 300
    End With

ErrTrap:
End Sub


Public Sub FillGridWithData(project_no As Integer, _
                            Optional TxtNoteSerial As String)

    'On Error GoTo ErrTrap

    Dim i As Integer
    Dim X As Integer
    Dim rs As ADODB.Recordset
 
    Dim ActualTotal As Double
    Dim TotalPayedFULL As Double
    Dim Result As Double
    Dim resultpercentage As Double
    Dim sql As String

    Grid.Clear flexClearScrollable, flexClearEverything
    Grid.rows = 1
          
    GRID1.Clear flexClearScrollable, flexClearEverything
    GRID1.rows = 1

    If DCboCashType.ListIndex = 5 Or DCboCashType.ListIndex = 11 Then
 
        lbl(38).Caption = DBCboClientName.text
        lbl(41).Caption = DBCboClientName.text
        sql = "SELECT project_billl.Total Total2, project_billl.FATYou,project_billl.PerforValue as performanceDiscount, project_billl.FATValue FATValue2,project_billl.Id Id2, * FROM     project_billl  LEFT OUTER JOIN projects AS p ON p.id = project_billl.project_no    "
        If DCboCashType.ListIndex = 11 Then
            sql = sql & "  where bill_to=0 and (sub_contractor_id = " & project_no
            sql = sql & "  Or End_user_id =" & project_no & ")"
            
        Else
            sql = sql & "  where project_no = " & project_no & " and bill_to=0"
        End If
        Set rs = New ADODB.Recordset
        rs.Open sql, Cn, adOpenStatic, adLockReadOnly, adCmdText
    
        If rs.RecordCount = 0 Then
     
            Exit Sub
        End If
    
        i = 0
    
        With Me.Grid
            .rows = 1
            .Clear flexClearScrollable
      
            rs.MoveFirst
            Dim xx As Double
            For X = 1 To rs.RecordCount
            
                ActualTotal = getBillPayedToproject(IIf(IsNull(rs.Fields("Id2").value), "", rs.Fields("Id2").value), Me.TxtNoteSerial)
                TotalPayedFULL = getBillPayedToproject(IIf(IsNull(rs.Fields("Id2").value), "", rs.Fields("Id2").value))
                Result = IIf(IsNull(rs.Fields("TotalValue").value), 0, rs("Total2").value) - TotalPayedFULL ' + IIf(IsNull(rs.Fields("FATValue2").value), 0, rs("FATValue2").value) - TotalPayedFULL
                
                xx = val(IIf(IsNull(rs.Fields("TotalValue").value), 0, rs("TotalValue").value))   ' + IIf(IsNull(rs.Fields("FATValue2").value), 0, rs("FATValue2").value)) * 100
                
                If xx <> 0 Then
                resultpercentage = ActualTotal / xx
                Else
                    resultpercentage = 0
                End If
                
     
                If val(rs.Fields("TotalValue").value & "") > ActualTotal Then
                    i = i + 1
                    .rows = .rows + 1
                    .TextMatrix(i, .ColIndex("Ser")) = i
                    .TextMatrix(i, .ColIndex("id")) = IIf(IsNull(rs.Fields("Id2").value), "", rs.Fields("Id2").value)
                    .TextMatrix(i, .ColIndex("ManualNO")) = IIf(IsNull(rs.Fields("ManualNO").value), "", rs.Fields("ManualNO").value)
                    .TextMatrix(i, .ColIndex("FATYou")) = IIf(IsNull(rs.Fields("FATYou").value), "", rs.Fields("FATYou").value)
                    .TextMatrix(i, .ColIndex("performanceDiscount")) = IIf(IsNull(rs.Fields("performanceDiscount").value), "", rs.Fields("performanceDiscount").value)
                    
                
                    .TextMatrix(i, .ColIndex("NoteSerial1")) = IIf(IsNull(rs.Fields("NoteSerial1").value), "", rs.Fields("NoteSerial1").value)
                
                    .TextMatrix(i, .ColIndex("bill_date")) = IIf(IsNull(rs.Fields("bill_date").value), "", rs.Fields("bill_date").value)
                    .TextMatrix(i, .ColIndex("project_no")) = IIf(IsNull(rs.Fields("project_no").value), "", rs.Fields("project_no").value)
                    .TextMatrix(i, .ColIndex("Project_name")) = IIf(IsNull(rs.Fields("project_name").value), "", rs.Fields("project_name").value)
                
                    .TextMatrix(i, .ColIndex("End_user_name")) = IIf(IsNull(rs.Fields("End_user_name").value), "", rs.Fields("End_user_name").value)
                
                    .TextMatrix(i, .ColIndex("Sub_user_name")) = IIf(IsNull(rs.Fields("Sub_user_name").value), "", rs.Fields("Sub_user_name").value)
                
                    .TextMatrix(i, .ColIndex("bill_to")) = IIf(IsNull(rs.Fields("bill_to").value), "", rs.Fields("bill_to").value)
     
                    .TextMatrix(i, .ColIndex("total")) = IIf(IsNull(rs.Fields("TotalValue").value), 0, rs.Fields("TotalValue").value) '- val(rs!FATValue & "") ' + IIf(IsNull(rs.Fields("FATValue2").Value), 0, rs.Fields("FATValue2").Value)
                
                .TextMatrix(i, .ColIndex("TotalPayedFULL")) = TotalPayedFULL
                    .TextMatrix(i, .ColIndex("ActualTotal")) = ActualTotal
                    .TextMatrix(i, .ColIndex("ResultPercentage")) = Round(ActualTotal / .TextMatrix(i, .ColIndex("total")) * 100, 2)
                    .TextMatrix(i, .ColIndex("Result")) = TotalPayedFULL
    
                End If
    
                rs.MoveNext
            Next
    
            rs.Close
     
            .RowHeight(-1) = 300
        End With
    
        If TxtNoteSerial = "" Then
    
            Exit Sub
        End If
    
       sql = "SELECT  * FROM     ProjectBillBuy     where TxtNoteSerial ='" & TxtNoteSerial & "'"
     sql = "SELECT     *,ProjectBillBuy.Value Value2,p.project_name,p.Id project_no, ProjectBillBuy.NoteSerial1,project_billl.FATYou, project_billl.ManualNo AS ManualNO ,ProjectBillBuy.Total Total2,project_billl.FATValue FATValue2"
    sql = sql & " FROM         dbo.ProjectBillBuy LEFT OUTER JOIN"
    sql = sql & " dbo.project_billl ON dbo.ProjectBillBuy.Bill_id = dbo.project_billl.id"
    sql = sql & " LEFT OUTER JOIN projects AS p ON p.id = ProjectBillBuy.Bill_id    "
    sql = sql & "  WHERE     (dbo.ProjectBillBuy.TxtNoteSerial = '" & TxtNoteSerial & "')"
        rs.Open sql, Cn, adOpenStatic, adLockReadOnly, adCmdText

        If rs.RecordCount = 0 Then
     
            Exit Sub
        End If

        With Me.GRID1
            .rows = 1
            .rows = .rows + rs.RecordCount
            .Clear flexClearScrollable
      
            rs.MoveFirst
    'Total   ﬁÌ„… «·„” Œ·’
    'ActualTotal «·„”œœ
            For i = 1 To .rows - 1
     
                .TextMatrix(i, .ColIndex("Ser")) = i
                .TextMatrix(i, .ColIndex("id")) = IIf(IsNull(rs.Fields("id").value), "", rs.Fields("id").value)
                .TextMatrix(i, .ColIndex("FATYou")) = IIf(IsNull(rs.Fields("FATYou").value), "", rs.Fields("FATYou").value)
                .TextMatrix(i, .ColIndex("bill_id")) = IIf(IsNull(rs.Fields("bill_id").value), "", rs.Fields("bill_id").value)
                .TextMatrix(i, .ColIndex("ManualNO")) = IIf(IsNull(rs.Fields("ManualNO").value), "", rs.Fields("ManualNO").value)
                .TextMatrix(i, .ColIndex("bill_date")) = IIf(IsNull(rs.Fields("RecordDate").value), "", rs.Fields("RecordDate").value)
                .TextMatrix(i, .ColIndex("NoteSerial1")) = IIf(IsNull(rs.Fields("NoteSerial1").value), "", rs.Fields("NoteSerial1").value)
                .TextMatrix(i, .ColIndex("project_no")) = IIf(IsNull(rs.Fields("project_no").value), "", rs.Fields("project_no").value)
                .TextMatrix(i, .ColIndex("Project_name")) = IIf(IsNull(rs.Fields("project_name").value), "", rs.Fields("project_name").value)
                
                '                                           .TextMatrix(I, .ColIndex("project_no")) = IIf(IsNull(rs.Fields("project_no").value), _
                                                            "", rs.Fields("project_no").value)
                '                         .TextMatrix(I, .ColIndex("Project_name")) = IIf(IsNull(rs.Fields("project_name").value), _
                                          "", rs.Fields("project_name").value)
                
                .TextMatrix(i, .ColIndex("bill_to")) = IIf(IsNull(rs.Fields("bill_to").value), "", rs.Fields("bill_to").value)
     
                .TextMatrix(i, .ColIndex("total")) = IIf(IsNull(rs.Fields("Total2").value), "", rs.Fields("Total2").value)
                
                .TextMatrix(i, .ColIndex("ActualTotal")) = IIf(IsNull(rs.Fields("value2").value), "", rs.Fields("value2").value)
                Result = val(.TextMatrix(i, .ColIndex("total"))) - val(rs.Fields("value2").value)
                If val(.TextMatrix(i, .ColIndex("Total"))) <> 0 Then
                    resultpercentage = Round(val(rs.Fields("value2").value) / val(.TextMatrix(i, .ColIndex("Total"))) * 100, 2) 'grid1
                Else
                
                    resultpercentage = 0
                End If
                
                .TextMatrix(i, .ColIndex("ResultPercentage")) = resultpercentage
                .TextMatrix(i, .ColIndex("Result")) = Result
          
                rs.MoveNext
            Next
    
            rs.Close
     
            .RowHeight(-1) = 300
        End With
    End If
ErrTrap:
End Sub


Function CalculateVATProje(finalAmount As Double, performanceDiscount As Double, vatPercentage As Double) As Double
    Dim VATValue As Double
    VATValue = ((finalAmount + performanceDiscount) * vatPercentage) / (100 + vatPercentage)
    CalculateVATProje = VATValue
End Function

Sub test()
    Dim finalAmount As Double
    Dim performanceDiscount As Double
    Dim vatPercentage As Double
    Dim VATValue As Double

    '  ⁄ÌÌ‰ «·ﬁÌ„
    finalAmount = 105000      ' «·„»·€ «·‰Â«∆Ì
    performanceDiscount = 10000 ' Œ’„ Õ”‰ «·√œ«¡
    vatPercentage = 15         ' ‰”»… «·÷—Ì»… %

    ' Õ”«» «·÷—Ì»…
    VATValue = CalculateVATProje(finalAmount, performanceDiscount, vatPercentage)

    ' ÿ»«⁄… «·‰ ÌÃ…
    MsgBox "ﬁÌ„… «·÷—Ì»… ÂÌ: " & Format(VATValue, "0.00")
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
 
    FillGridWithData val(Me.DBCboClientName.BoundText), TxtNoteSerial.text

End Sub

Private Sub CboPayMentType_Change()
'DBCboClientName_Change
FramePay.Visible = False
    If Me.TxtModFlg.text = "E" Then
        DcboBankName.text = ""
        TxtChequeNumber.text = ""
        Me.DcboBox.text = ""
        DCChequeBox.text = ""
        TXTBankName.text = ""
    End If

    DCChequeBox.Enabled = False

    If SystemOptions.UserInterface = ArabicInterface Then
        lbl(16).Caption = "—ﬁ„ «·‘Ìﬂ"
        lbl(17).Caption = " «—ÌŒ «·«” Õﬁ«ﬁ"
    
    Else
        lbl(16).Caption = "Cheque No"
        lbl(17).Caption = "Due Date"
    End If
    
    If Me.CboPayMentType.ListIndex = 0 Then
        TxtAccount.Enabled = False
        DcbAccount.Enabled = False
        TxtAccount.text = ""
        DcbAccount.BoundText = ""
        Me.lbl(9).Enabled = True
        Me.DcboBox.Enabled = True
        Me.lbl(15).Enabled = False
        Me.lbl(16).Enabled = False
        Me.lbl(17).Enabled = False
        Me.DcboBankName.Enabled = False
        Me.TxtChequeNumber.Enabled = False
        Me.DtpChequeDueDate.Enabled = False
        Frame3.Enabled = False
'                    GetUserData user_id, usertype, userbranchid, dstore, dBox, , EmpID
'DcboBox.BoundText = dBox


    ElseIf Me.CboPayMentType.ListIndex = 1 Then

        If SystemOptions.ChequeBox = True Then
            TXTBankName.Visible = True
            DCChequeBox.Enabled = True
        Else
            TXTBankName.Visible = False
        End If
        TxtAccount.Enabled = False
        DcbAccount.Enabled = False
        TxtAccount.text = ""
        DcbAccount.BoundText = ""
        Me.lbl(9).Enabled = False
        Me.DcboBox.Enabled = False
        Me.lbl(15).Enabled = True
        Me.lbl(16).Enabled = True
        Me.lbl(17).Enabled = True
        Me.DcboBankName.Enabled = True
        Me.TxtChequeNumber.Enabled = True
        Me.DtpChequeDueDate.Enabled = True
        Frame3.Enabled = False
    ElseIf Me.CboPayMentType.ListIndex = 2 Then
 
        TXTBankName.Visible = False
        TxtAccount.Enabled = False
        DcbAccount.Enabled = False
        TxtAccount.text = ""
        DcbAccount.BoundText = ""
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
            lbl(16).Caption = "—ﬁ„ «·ÕÊ«·Â"
            lbl(17).Caption = " «—ÌŒÂ«"
    
        Else
            lbl(16).Caption = "Transfer No"
            lbl(17).Caption = "Date"
        End If
 
    ElseIf Me.CboPayMentType.ListIndex = 3 Then
 
        TXTBankName.Visible = False
        TxtAccount.Enabled = False
        DcbAccount.Enabled = False
        TxtAccount.text = ""
        DcbAccount.BoundText = ""
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
       ElseIf Me.CboPayMentType.ListIndex = 4 Then
 
        TXTBankName.Visible = False
        TxtAccount.Enabled = True
        DcbAccount.Enabled = True
        Me.lbl(9).Enabled = False
        Me.DcboBox.Enabled = False
        Me.lbl(15).Enabled = False
        Me.lbl(16).Enabled = False
        Me.lbl(17).Enabled = False
        Me.DcboBankName.Enabled = False
        Me.TxtChequeNumber.Enabled = False
        Me.DtpChequeDueDate.Enabled = False
        TXTBankName.Visible = False
        Frame3.Enabled = False
     ElseIf Me.CboPayMentType.ListIndex = 5 Then
     Me.DcboBox.BoundText = dBox
        TxtAccount.Enabled = False
        DcbAccount.Enabled = False
        TxtAccount.text = ""
        DcbAccount.BoundText = ""
     
         
     
    If SystemOptions.CanEditOnlyPayMethod And (Me.TxtModFlg = "E" Or Me.TxtModFlg = "R") Then
      '  Label20.Enabled = False
      '  lblexit(90).Enabled = False
      '  Ele(12).Enabled = False
      '  XPTab301.Enabled = False
    Else
      '   Label20.Enabled = True
      '   lblexit(90).Enabled = True
      '   XPTab301.Enabled = True
    End If
    FramePay.Visible = True


  '   If Me.TxtModFlg.Text <> "R" And Me.TxtModFlg.Text <> "" Then
     If Me.TxtModFlg.text = "N" Then
     If val(XPTxtVal.text) > 0 Then
        If SystemOptions.CanEditOnlyPayMethod And (Me.TxtModFlg = "E" Or Me.TxtModFlg = "R") Then
            Label20.Enabled = False
            lblexit(90).Enabled = False
            Ele(12).Enabled = False
            XPTab301.Enabled = False
        Else
            Label20.Enabled = True
            lblexit(90).Enabled = True
            XPTab301.Enabled = True
    
        End If
         
 FramePay.Visible = True


     FillGridWithData222
     LBLPayVal.Caption = 0
LBLPayVal.Caption = val(XPTxtVal.text) + val(txtVat2)
TxtNetValue2.text = val(LBLPayVal.Caption)
    With Grid22
          .TextMatrix(.row, .ColIndex("Value")) = 0
    End With
     ReLineGrid2
     End If
     Else
     If SystemOptions.CanEditOnlyPayMethod And (Me.TxtModFlg = "E" Or Me.TxtModFlg = "R") Then
        Label20.Enabled = False
        lblexit(90).Enabled = False
        Ele(12).Enabled = False
        XPTab301.Enabled = False
    Else
         Label20.Enabled = True
         lblexit(90).Enabled = True
         XPTab301.Enabled = True
    End If
         
 FramePay.Visible = True


      
    If FillGridWithDataPayment() = True Then
     LBLPayVal.Caption = val(XPTxtVal.text) + val(txtVat2)
     TxtNetValue2.text = val(LBLPayVal.Caption)
     ReLineGrid2
     Else
     '''
    If val(XPTxtVal.text) > 0 Then
        If SystemOptions.CanEditOnlyPayMethod And (Me.TxtModFlg = "E" Or Me.TxtModFlg = "R") Then
            Label20.Enabled = False
            lblexit(90).Enabled = False
            Ele(12).Enabled = False
            XPTab301.Enabled = False
        Else
         Label20.Enabled = True
         lblexit(90).Enabled = True
         XPTab301.Enabled = True
    End If
 FramePay.Visible = True

     FillGridWithData222
     LBLPayVal.Caption = 0
LBLPayVal.Caption = val(XPTxtVal.text) + val(txtVat2)
TxtNetValue2.text = val(LBLPayVal.Caption)
    With Grid22
          .TextMatrix(.row, .ColIndex("Value")) = 0
    End With
     ReLineGrid2
     End If
     End If
     
     End If
          Me.lbl(9).Enabled = True
        Me.DcboBox.Enabled = True
        Me.lbl(15).Enabled = False
        Me.lbl(16).Enabled = False
        Me.lbl(17).Enabled = False
        Me.DcboBankName.Enabled = False
        Me.TxtChequeNumber.Enabled = False
        Me.DtpChequeDueDate.Enabled = False
        Frame3.Enabled = False
'    End If
        Else

        Me.lbl(9).Enabled = False
        Me.DcboBox.Enabled = False
        Me.lbl(15).Enabled = False
        Me.lbl(16).Enabled = False
        Me.lbl(17).Enabled = False
        Me.DcboBankName.Enabled = False
        Me.TxtChequeNumber.Enabled = False
        Me.DtpChequeDueDate.Enabled = False
    End If

End Sub

Private Sub CboPayMentType_Click()

If DCboCashType.ListIndex = 7 Then
Else
DBCboClientName_Change
End If

    CboPayMentType_Change
End Sub

Private Sub CboStatus_Click()
If Me.TxtModFlg = "N" Or Me.TxtModFlg = "E" Then


DBCboClientName_Change
If CboPayMentType.ListIndex = 0 Or CboPayMentType.ListIndex = 5 Then
DcboBox_Change
Else
DcChequeBox_Change
DcboBankName_Click 0
End If


If CboStatus.ListIndex <> 0 And CboStatus.ListIndex <> 4 Then
   Me.DcboDebitSide.BoundText = ""
   Me.DcboCreditSide.BoundText = ""
   
End If

End If
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

Public Sub Cmd_Click(index As Integer)

    Dim cNoteReport As ClsNotesReports
    Dim Msg         As String
    '  On Error GoTo ErrTrap

    Select Case index
        Case 11 ' SMS
            If TxtModFlg.text = "N" Or TxtModFlg.text = "E" Then
                Msg = " «Õ›Ÿ «Ê·« "
                MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title

            Else
               ' Msg = "«·”«œÂ " & DBCboClientName.text & "   „ «’œ«— ”‰œ ﬁ»÷ " & TxtNoteSerial1.text & " »ﬁÌ„… " & lbl(18) & " » «—ÌŒ " & XPDtbTrans.value & " ⁄·„« »«‰ —’Ìœﬂ„ «œÌ‰«   " & LblLink.Caption & ""
                
                Msg = StringDotFormat("«·”«œÂ{3}  „ «’œ«— ”‰œ ﬁ»÷ {2} »ﬁÌ„…{1} » «—ÌŒ {0} ⁄·„« »«‰ —’Ìœﬂ„ ·œÌ‰« {4}", _
                   Format(XPDtbTrans.value, "yyyy-MM-dd"), _
                   lbl(18), _
                   TxtNoteSerial1.text, _
                   DBCboClientName.text, _
                   LblLink)
                
                SendSMSToClient DBCboClientName.BoundText, Msg
            End If
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
            If SystemOptions.CanEditOnlyPayMethod Then
                If Not isFormFirstRun Then
                    chkPaymentPermission False
                    GetDefaultEnabled True
                End If
            End If
            clear_all Me
        
            Grid.Clear flexClearScrollable, flexClearEverything
            Grid.rows = 1
          
            GRID1.Clear flexClearScrollable, flexClearEverything
            GRID1.rows = 1
            VSFlexGrid1.Clear flexClearScrollable, flexClearEverything
            VSFlexGrid1.rows = 1
            Grid2.Clear flexClearScrollable, flexClearEverything
            Grid2.rows = 1
            Grid3.Clear flexClearScrollable, flexClearEverything
            Grid3.rows = 1
          
            Grid4.Clear flexClearScrollable, flexClearEverything
            Grid4.rows = 1
            
            TxtModFlg.text = "N"
            '       XPTxtID.text = CStr(new_id("Notes", "NoteID", "", True))
            ' Me.TxtNoteSerial.text = CStr(new_id("Notes", "NoteSerial", "", True, "NoteType=4"))
            Me.DCboUserName.BoundText = user_id
            lbl(18).Caption = ""
            Text1.text = setfoxy
            Option1.value = True
            Me.dcBranch.BoundText = Current_branch
            Txt_DateHigri.value = ToHijriDate(Date)
            commdiscounttype.ListIndex = 0
            Accredit.Caption = ""
            XPTab301.CurrTab = 0
            CboPayMentType.ListIndex = 2
            '  XPDtbTrans.SetFocus
            ' Option1.value = True
            Option4.value = True
            cbointervaltype.ListIndex = 0
            CboStatus.ListIndex = 0
            
            'GetBranchData branch_id, dstore, dBox
            If SystemOptions.usertype <> UserAdminAll Then
 
                '  Me.Dcbranch.Enabled = True
                ' XPDtbBill.Enabled = False
                Me.dcBranch.BoundText = Current_branch
            End If
                 
            GetUserData user_id, usertype, userbranchid, dstore, dBox, , EmpID
            DcEmp.BoundText = EmpID
            CboPayMentType.ListIndex = 1
            DcboBox.BoundText = dBox
            Ele(12).Enabled = True
            Frame12.Enabled = True
            Ele(0).Enabled = True
            XPTab301.Enabled = True
            
            'DBCboClientName.BoundText = ""
            'TxtEmployeeID.text = ""
            'XPDtbTrans.SetFocus

            '3 1 2 7
            If SystemOptions.ChasingStatus = 3 Then
                Option3.value = True
            ElseIf SystemOptions.ChasingStatus = 1 Then
                Option1.value = True
            ElseIf SystemOptions.ChasingStatus = 2 Then
                Option2.value = True
            ElseIf SystemOptions.ChasingStatus = 7 Then
                Option7.value = True
            End If

            If SystemOptions.EnableCustomerAging = False Then
                'Option2.value = True
            End If

        Case 1
            If ScreenAproved(val(XPTxtID.text), Me.Name) = True Then
                If SystemOptions.UserInterface = ArabicInterface Then
                    MsgBox "·«Ì„ﬂ‰ «· ⁄œÌ·.Â–Â «·Õ—ﬂ… „— »ÿ… »«·«⁄ „«œ« "
                Else
                    MsgBox "Can not edit.This process associated with approvals"
                End If
                Exit Sub
            End If
 
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
              
            If SystemOptions.ChequeBox = True And CboPayMentType.ListIndex = 1 Then
         
                If ChequeBoxOperations(val(Me.XPTxtID)) = False Then
                    Msg = "·‰ Ì„ﬂ‰ «·”„«Õ » ⁄œÌ· Â–« «·⁄„·Ì…..!!!"
                    Msg = Msg & CHR(13) & "ÕÌÀ «‰Â« ”Ê› Ì‰ Ã ⁄‰Â« Œÿ« ›Ï   Õ«›Ÿ… «·‘Ìﬂ«  ·«‰Â  „ ⁄·ÌÂ« Õ—ﬂ«  «Ìœ«⁄ «Ê  Õ’Ì· "
                    MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
                    Exit Sub
                End If
    
            End If
            If SystemOptions.AllowEditCashingLinkProj = False Then
                If CheckProjectBill(val(XPTxtID.text)) = True Then
                    If SystemOptions.UserInterface = ArabicInterface Then
                        MsgBox " ·«Ì„ﬂ‰  ⁄œÌ· Â–Â «·Õ—ﬂ… „— »ÿ… »›Ê« Ì— «·„‘«—Ì⁄ "
                    Else
                        MsgBox "You can not edit.This is process Link to Projects Bill"
                    End If
               
                    If SystemOptions.CanEditOnlyPayMethod And (Me.TxtModFlg = "E" Or Me.TxtModFlg = "R") Then
   
                        Ele(12).Enabled = False
                        Frame12.Enabled = False
                        Ele(0).Enabled = False
                        XPTab301.Enabled = False
        
                    End If
                    Exit Sub
                End If
            End If
            TxtModFlg.text = "E"
            '      Me.DCboUserName.BoundText = user_id
            CuurentLogdata
            If SystemOptions.CanEditOnlyPayMethod And (Me.TxtModFlg = "E" Or Me.TxtModFlg = "R") Then
            
                chkPaymentPermission
            End If
            If Me.CboPayMentType.ListIndex = 5 Then
 
                FramePay.Visible = True
            End If
            'XPDtbTrans.SetFocus
        Case 2
            Dim NoteSerial1 As String
            If SystemOptions.CantRepetttransferNoforCashing = True And CboPayMentType.ListIndex Then
 
                If ChekTransferNo(TxtChequeNumber.text, val(DcboBankName.BoundText), val(Me.XPTxtID), NoteSerial1) = True Then
                    MsgBox "—ﬁ„ «·ÕÊ«·Â „” Œœ„ „‰ ﬁ»· ·”‰œ «·ﬁ»÷ —ﬁ„ " & NoteSerial1
                    Exit Sub
            
                End If
            End If
 
            If ChekClodePeriod(XPDtbTrans.value) = True Then
                If SystemOptions.UserInterface = ArabicInterface Then
                    MsgBox "Ì—ÃÏ  €ÌÌ—  «—ÌŒ «·Õ—ﬂ… ·«‰ Â–Â «·› —… „€·ﬁ…"
                Else
                    MsgBox "Please Change Date Becouse This is Period is Closed"
                End If
                Exit Sub
            End If
            XPTab301.CurrTab = 0
            If Trim(dcBranch.BoundText) = "" Then
                If SystemOptions.UserInterface = EnglishInterface Then
                    Msg = "Specify Branch"
                Else
                    Msg = "Õœœ «·›—⁄ "
                End If

                MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
                dcBranch.SetFocus
                Sendkeys "{F4}"
                Screen.MousePointer = vbDefault
                Exit Sub
            End If

            If val(CboPayMentType.ListIndex) = 5 Then
                If val(TxtRemainValue2.text) <> 0 Then
                    If SystemOptions.UserInterface = ArabicInterface Then
                        MsgBox "«·ﬁÌ„… «·„œŒ·… €Ì— ’ÕÌÕ…"
                    Else
                        MsgBox "The  value is incorrect"
                    End If
                    Exit Sub
                End If
                FramePay.Visible = False
            End If
            If Me.Option1.value = True Then
                If val(DCboCashType.ListIndex) = 0 Then
                    If val(XPTxtVal.text) <= 0 Then
                        If SystemOptions.UserInterface = ArabicInterface Then
                            MsgBox "Ì—ÃÏ «œŒ«· «·ﬁÌ„…"
                        Else
                            MsgBox "Please Enter Value "
                        End If
                        'XPTxtVal.SetFocus
                        Exit Sub
                    Else
                        BillCustomer 1
                        If AutoCalculate = False Then
                            If SystemOptions.UserInterface = ArabicInterface Then
                                MsgBox "·«ÌÊÃœ ›Ê« Ì— «Ê «‰ «·ﬁÌ„… «·„œŒ·… «ﬂ»— „‰ «·„” Õﬁ "
                            Else
                                MsgBox "Not Found Bills"
                            End If
                            Exit Sub
                        End If
                        'Frame12.Visible = True
                    End If
                End If
            End If

            If val(TxtVATValue.text) > 0 Then
                If GetValueAddedAccount(XPDtbTrans.value, , , 1, 23) = False Then
                    If SystemOptions.UserInterface = ArabicInterface Then
                        MsgBox "·„ Ì „  ÕœÌœ Õ”«» «·ﬁÌ„… «·„÷«›… ··„⁄«„·«  «·„«·Ì…"
                    Else
                        MsgBox "Value added account not specified"
                    End If
                    Exit Sub
                End If
            End If
            my_branch = Me.dcBranch.BoundText
 
            '   If Option2.value = True And lblsqlstring.Caption = "" Then MsgBox "·«»œ „‰  ÕœÌœ ›Ê« Ì—": Exit Sub
 
            'TxtNoteSerial.text = Notes_coding(Val(my_branch), XPDtbTrans.value)

            If SystemOptions.DealingWithPrepayAccount = True Then
        
                Dim Account_Code_dynamic82 As String
                If SystemOptions.CustomerhavethreeAccounts = False And Option3.value = True And val(DCboCashType.ListIndex) = 0 Then
                    Account_Code_dynamic82 = get_account_code_branch(158, my_branch)
                    If Account_Code_dynamic82 = "NO account" Then
                        If SystemOptions.UserInterface = ArabicInterface Then
                            MsgBox "·„ Ì „  ÕœÌœ Õ”«» œ›⁄«  „ﬁœ„… ··⁄„·«¡   ", vbCritical
                        Else
                            MsgBox "Please Select  Account", vbCritical
                        End If

                        GoTo ErrTrap
                    End If
                End If
     
            End If
            SaveData
        
        Case 3
            Undo

        Case 4
        
            If ScreenAproved(val(XPTxtID.text), Me.Name) = True Then
                If SystemOptions.UserInterface = ArabicInterface Then
                    MsgBox "·«Ì„ﬂ‰ «·Õ–›.Â–Â «·Õ—ﬂ… „— »ÿ… »«·«⁄ „«œ« "
                Else
                    MsgBox "Can not delete.This process associated with approvals"
                End If
                Exit Sub
            End If

            If ChekClodePeriod(XPDtbTrans.value) = True Then
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
            If CheckProjectBill(val(XPTxtID.text)) = True Then
                If SystemOptions.UserInterface = ArabicInterface Then
                    MsgBox " ·«Ì„ﬂ‰ Õ–› Â–Â «·Õ—ﬂ… „— »ÿ… »›Ê« Ì— «·„‘«—Ì⁄ "
                Else
                    MsgBox "You can not delete.This is process Link to Projects Bill"
                End If
                Exit Sub
            End If
            If SystemOptions.ChequeBox = True And CboPayMentType.ListIndex = 1 Then
         
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
            FrmNotesSearch.SearchType = 4
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
            If TxtNoteSerial1 <> "" Then
                print_report TxtNoteSerial, Me.TxtNoteSerial1.text, TXTBankName.text, CboPayMentType.text, DcboBox.text, TxtCustCode.text
            End If

        Case 8

            'ViewDataList
        Case 9
    
            If DoPremis(Do_Print, Me.Name, True) = False Then
                Exit Sub
            End If

            ShowGL_cc Me.TxtNoteSerial.text, , 200, val(XPTxtID.text)
        Case 10

            If DoPremis(Do_Print, Me.Name, True) = False Then
                Exit Sub
            End If

            ' If Val(Me.XPTxtID.text) <> 0 Then
            '     Set cNoteReport = New ClsNotesReports
            '     cNoteReport.PrintReceipt Val(Me.XPTxtID.text), WindowTarget
            '     Set cNoteReport = Nothing
            ' End If
            If TxtNoteSerial <> "" Then
                print_report2 TxtNoteSerial, Me.TxtNoteSerial1.text, TXTBankName.text, CboPayMentType.text, DcboBox.text, TxtCustCode.text
            End If
    End Select

    Exit Sub
ErrTrap:
End Sub
Public Function print_report2(Optional NoteSerial As String, Optional NoteSerial1 As String, Optional BankName As String, Optional PaymentType As String, Optional Box As String, Optional Custcode As String)
    
    Dim MySQL As String
    Dim RsData As New ADODB.Recordset
    Dim xApp As New CRAXDRT.Application
    Dim xReport As CRAXDRT.Report
    Dim CViewer As ClsReportViewer
    Dim StrReportTitle As String
    Dim StrFileName As String
    Dim Msg As String
    
       SaveQRCode "Notes", "NoteID", val(XPTxtID.text), TxtNoteSerial1.text, (XPDtbTrans.value), _
        (txtTotal.text), Picture1, 0, (txtVat2.text), (txtTotal.text)
    
   

   
 
    
    MySQL = " SELECT        dbo.Notes.Note_Value, dbo.Notes.BankID, dbo.Notes.ChqueNum, dbo.BanksData.BankName, dbo.Notes.NoteType, dbo.Notes.BoxID, dbo.TblBoxesData.BoxName, dbo.Notes.CusID, dbo.TblCustemers.CusName,"
MySQL = MySQL & "                         dbo.TblCustemers.CusNamee, dbo.Notes.Remark, dbo.Notes.NoteSerial, dbo.Notes.NoteDate, dbo.Notes.note_value_by_characters, dbo.Notes.NoteID, dbo.Notes.general_des_notes, dbo.Notes.person,"
MySQL = MySQL & "                         dbo.notes.ManulaNO,Notes.QrCodeData,Notes.QrCodeDataPath,Notes.QrCodeImage,TblCustemers.VATNO,TblCustemers.Mobile1,TblCustemers.Cus_mobile,TblCustemers.Address,TblCustemers.E_mail"
MySQL = MySQL & "                         ,TblCustemers.Cus_Phone,TblCustemers.Cus_mobile,TblCustemers.E_mail,TblCustemers.CustGID as CustGID,TblCustemers.VATNO,"
MySQL = MySQL & "                                 TblBranchesData.Company_Arabic_Name,TblBranchesData.Company_Name_Eng,TblBranchesData.StreetName ,"
MySQL = MySQL & "                                 TblBranchesData.CitySubdivisionName,TblBranchesData.VATNO VATNOCompany,TblBranchesData.company_comment as VATRegNoCompany  "
MySQL = MySQL & " FROM            dbo.Notes LEFT OUTER JOIN"
MySQL = MySQL & "                         dbo.TblBoxesData ON dbo.Notes.BoxID = dbo.TblBoxesData.BoxID LEFT OUTER JOIN"
MySQL = MySQL & "                         dbo.TblCustemers ON dbo.Notes.CusID = dbo.TblCustemers.CusID LEFT OUTER JOIN"
MySQL = MySQL & "                         dbo.BanksData ON dbo.Notes.BankID = dbo.BanksData.BankID"
MySQL = MySQL & "                                 Left Outer join TblBranchesData ON dbo.Notes.branch_no = dbo.TblBranchesData.branch_id "
MySQL = MySQL & " Where (dbo.notes.NoteType = 4)  and NoteID=" & val(XPTxtID.text)
    
    'MySQL = "Select * From payment_voucher  where NoteID=" & val(XPTxtID.text)

    If SystemOptions.UserInterface = ArabicInterface Then
    '    StrFileName = App.path & "\Reports\" & "Payment_voucher.rpt"
        StrFileName = App.path & "\Special\" & SystemOptions.Reportpath & "\Payment_voucher2.rpt"
    Else
     '   StrFileName = App.path & "\Reports\" & "Payment_voucherE.rpt"
        StrFileName = App.path & "\Special\" & SystemOptions.Reportpath & "\Payment_voucher2.rpt"
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
        xReport.ParameterFields(5).AddCurrentValue DcboCreditSide.text
   
        StrReportTitle = "" '& StrAccountName
 
    Else
 
        xReport.ParameterFields(1).AddCurrentValue cCompanyInfo.ArabCompanyName ' RPTCompany_Name_Eng
    
        xReport.ParameterFields(4).AddCurrentValue get_branch_name(val(my_branch))
        xReport.ParameterFields(5).AddCurrentValue DcboCreditSide.text
        StrReportTitle = ""
 
    End If
Dim i As Integer
Dim str As String
With VSFlexGrid1
str = ""
For i = 1 To .rows - 1
If (.TextMatrix(i, .ColIndex("NoteSerial1"))) <> "" And .cell(flexcpChecked, i, .ColIndex("payed")) = flexChecked Then
str = str & .TextMatrix(i, .ColIndex("NoteSerial1"))
If i <> (.rows - 1) Then
str = str & ","
End If
End If
Next i
End With
    xReport.ParameterFields(3).AddCurrentValue user_name
    '
    xReport.ParameterFields(6).AddCurrentValue NoteSerial1

    xReport.ParameterFields(7).AddCurrentValue BankName
    xReport.ParameterFields(8).AddCurrentValue PaymentType
    xReport.ParameterFields(9).AddCurrentValue Box
    xReport.ParameterFields(10).AddCurrentValue Custcode
    xReport.ParameterFields(11).AddCurrentValue str
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
Public Function print_report(Optional NoteSerial As String, Optional NoteSerial1 As String, Optional BankName As String, Optional PaymentType As String, Optional Box As String, Optional Custcode As String)
    
    Dim MySQL As String
    Dim RsData As New ADODB.Recordset
    Dim xApp As New CRAXDRT.Application
    Dim xReport As CRAXDRT.Report
    Dim CViewer As ClsReportViewer
    Dim StrReportTitle As String
    Dim StrFileName As String
    Dim Msg As String


ClaCul
Cn.Execute "Update Notes set note_value_by_characters = N'" & Trim(lbl(18).Caption) & "'   where NoteID=" & val(XPTxtID.text) & " and (dbo.Notes.NoteType = 4) "
   ' MySQL = "Select * From payment_voucher  where NoteID=" & val(XPTxtID.Text)
MySQL = "SELECT BillMaintNo, Notes.paydes,    dbo.Notes.Note_Value, dbo.Notes.BankID, dbo.Notes.ChqueNum, dbo.BanksData.BankName, dbo.Notes.NoteType, dbo.Notes.BoxID, dbo.TblBoxesData.BoxName, "
MySQL = MySQL & "                        dbo.Notes.CusID, dbo.TblCustemers.CusName, dbo.TblCustemers.CusNamee, dbo.Notes.Remark, dbo.Notes.NoteSerial, dbo.Notes.NoteDate,"
MySQL = MySQL & "                                 dbo.Notes.note_value_by_characters,Notes.NoteCashingType, dbo.Notes.NoteID, dbo.Notes.general_des_notes, dbo.Notes.person,TblCustemers.VATNO,TblCustemers.CustGID, dbo.TblCustemers.Fullcode, dbo.Notes.PreVAT,"
MySQL = MySQL & "                                 dbo.Notes.Vat , dbo.Notes.NoteSerial1, dbo.Notes.ManulaNO, dbo.Notes.ManualNO,TblCustemers.Address,"
MySQL = MySQL & "                                 TblBranchesData.Company_Arabic_Name,TblBranchesData.Company_Name_Eng,TblBranchesData.StreetName ,"
MySQL = MySQL & "                                 TblBranchesData.CitySubdivisionName,TblBranchesData.VATNO VATNOCompany,TblBranchesData.company_comment as VATRegNoCompany  "
MySQL = MySQL & "           FROM         dbo.Notes LEFT OUTER JOIN"
MySQL = MySQL & "                                 dbo.TblBoxesData ON dbo.Notes.BoxID = dbo.TblBoxesData.BoxID LEFT OUTER JOIN"
MySQL = MySQL & "                                 dbo.TblCustemers ON dbo.Notes.CusID = dbo.TblCustemers.CusID LEFT OUTER JOIN"
MySQL = MySQL & "                                 dbo.BanksData ON dbo.Notes.BankID = dbo.BanksData.BankID"

MySQL = MySQL & "                                 Left Outer join TblBranchesData ON dbo.Notes.branch_no = dbo.TblBranchesData.branch_id "
MySQL = MySQL & "           Where (dbo.Notes.NoteType = 4)"
MySQL = MySQL & "           and NoteID=" & val(XPTxtID.text)


    If SystemOptions.UserInterface = ArabicInterface Then
    '    StrFileName = App.path & "\Reports\" & "Payment_voucher.rpt"
        StrFileName = App.path & "\Special\" & SystemOptions.Reportpath & "\Payment_voucher.rpt"
    Else
     '   StrFileName = App.path & "\Reports\" & "Payment_voucherE.rpt"
        StrFileName = App.path & "\Special\" & SystemOptions.Reportpath & "\Payment_voucher.rpt"
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
        xReport.ParameterFields(5).AddCurrentValue DcboCreditSide.text
   
        StrReportTitle = "" '& StrAccountName
 
    Else
 
        xReport.ParameterFields(1).AddCurrentValue cCompanyInfo.ArabCompanyName ' RPTCompany_Name_Eng
    
        xReport.ParameterFields(4).AddCurrentValue get_branch_name(val(my_branch))
        xReport.ParameterFields(5).AddCurrentValue DcboCreditSide.text
        StrReportTitle = ""
 
    End If
Dim i As Integer
Dim str As String
With VSFlexGrid1
str = ""
For i = 1 To .rows - 1
If (.TextMatrix(i, .ColIndex("NoteSerial1"))) <> "" And .cell(flexcpChecked, i, .ColIndex("payed")) = flexChecked Then
str = str & .TextMatrix(i, .ColIndex("NoteSerial1"))
If i <> (.rows - 1) Then
str = str & ","
End If
End If
Next i
End With
    xReport.ParameterFields(3).AddCurrentValue user_name
    '
    xReport.ParameterFields(6).AddCurrentValue NoteSerial1

    xReport.ParameterFields(7).AddCurrentValue BankName
    xReport.ParameterFields(8).AddCurrentValue PaymentType
    xReport.ParameterFields(9).AddCurrentValue Box
    xReport.ParameterFields(10).AddCurrentValue Custcode
    xReport.ParameterFields(11).AddCurrentValue str
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
    Dim rs As ADODB.Recordset
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
        Set rs = New ADODB.Recordset
        rs.Open StrSQL, Cn, adOpenKeyset, adLockReadOnly, adCmdText
        'Â‰« Ìﬂ » ﬂÊœ ·⁄„· „⁄œ·  Õ„Ì· «·»Ì«‰« 
        '------------------------------------
        '
        '
        '
        '
    
        '------------------------------------
        Set .DataSource = rs
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
    
        'Rs.Close
        'Set Rs = Nothing
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
    sql = "delete from ProjectBillBuy where id=" & val(GRID1.TextMatrix(GRID1.row, GRID1.ColIndex("id")))
    Cn.Execute sql

    If GRID1.rows > 1 Then
        If GRID1.rows = 2 Then
            Me.GRID1.Clear flexClearScrollable, flexClearEverything
        Else

            If Me.GRID1.rows > 1 Then
                If Me.GRID1.row <> Me.GRID1.FixedRows - 1 Then
                    Me.GRID1.RemoveItem (Me.GRID1.row)
                End If
            End If
        End If
    End If

    If DCboCashType.ListIndex = 5 Or DCboCashType.ListIndex = 11 Then
        FillGridWithData val(Me.DBCboClientName.BoundText), TxtNoteSerial.text
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
        FrmBuySearch.CboPayMentType.ListIndex = 1
        FrmBuySearch.CboPayMentType.Enabled = False
        FrmBuySearch.Caption = "«·»ÕÀ ⁄‰ ⁄„·Ì… »Ì⁄"
        FrmBuySearch.DCboClientsName.BoundText = Me.DBCboClientName.BoundText
        FrmBuySearch.show
    ElseIf Me.CboTrans.ListIndex = 1 Then
        '›« Ê—… „— Ã⁄ „‘ —Ì« 
        Load FrmBuySearch
        FrmBuySearch.DealingForm = Returntransaction
        Set FrmBuySearch.ExtraRetrunObject = Me.TxtTransID
        FrmBuySearch.CboPayMentType.ListIndex = 1
        FrmBuySearch.CboPayMentType.Enabled = False
        FrmBuySearch.Caption = "«·»ÕÀ ⁄‰ „— Ã⁄ «·„‘ —Ì« "
        FrmBuySearch.DCboClientsName.BoundText = Me.DBCboClientName.BoundText
        FrmBuySearch.show vbModal
    ElseIf Me.CboTrans.ListIndex = 2 Then
        '›« Ê—… ’Ì«‰…
        Load FrmMaintanenceSearch
        Set FrmMaintanenceSearch.ExtraRetrunObject = Me.TxtTransID
        FrmMaintanenceSearch.CboPayMentType.ListIndex = 1
        FrmMaintanenceSearch.SearchType = 4
        FrmMaintanenceSearch.CboPayMentType.Enabled = False
        FrmMaintanenceSearch.show vbModal
    End If

End Sub





Private Sub Command1_Click()
FillGridWithData222
End Sub

Private Sub Command10_Click()
Dim i As Integer
Dim StrSQL As String
If Me.TxtModFlg.text = "E" Then
DeleteBillBuy
VSFlexGrid1.Enabled = True
        Check1.Enabled = True
      StrSQL = "Delete From TblNotesBillBuyPayment2 Where NoteID1=" & val(Me.XPTxtID.text) & " and TransType is null"
    Cn.Execute StrSQL, , adExecuteNoRecords
        StrSQL = "Delete From TblBillBuyPayment2 Where TypTrans IS NULL and  NoteID=" & val(Me.XPTxtID.text)
    Cn.Execute StrSQL, , adExecuteNoRecords
XPTxtVal.text = 0
            VSFlexGrid1.Clear flexClearScrollable, flexClearEverything
VSFlexGrid1.rows = 1

FlgBillBuy = True
If SystemOptions.UserInterface = ArabicInterface Then
MsgBox " „ «·€«¡ «·”œ«œ"
Else
MsgBox "Done"
End If
    With Me.VSFlexGrid1

            For i = .FixedRows To .rows - 1
        
                .TextMatrix(i, .ColIndex("payed")) = False
            Next i


        End With
End If
End Sub
Sub DeleteBillBuy()
Dim i As Integer
Dim StrSQL As String
With VSFlexGrid1
 For i = .FixedRows To .rows - 1
 If val(.TextMatrix(i, .ColIndex("NoteID"))) <> 0 Then
        If val(.TextMatrix(i, .ColIndex("Transaction_Type"))) <> 9999 Then
            StrSQL = "Update Transactions Set  TotalPayed=0 Where Transaction_ID=" & val(.TextMatrix(i, .ColIndex("NoteID"))) & ""
                Cn.Execute StrSQL, , adExecuteNoRecords
        Else
             StrSQL = "Update TblTravDueK Set  TotalPayed=0 Where ID=" & val(.TextMatrix(i, .ColIndex("NoteID"))) & ""
            Cn.Execute StrSQL, , adExecuteNoRecords
        End If
     End If
     Next i
 End With
End Sub
Sub RelineBu22()
    Dim IntCounter As Integer
    Dim Sm As Double
    Sm = 0
    IntCounter = 0
    Dim i As Integer
    
    With Me.VSFlexGrid1
        For i = .FixedRows To .rows - 1
                If .cell(flexcpChecked, i, .ColIndex("payed")) = flexChecked Then
    
      If val(.TextMatrix(i, .ColIndex("TransPayedValue"))) < 0 Then
         If SystemOptions.UserInterface = ArabicInterface Then
              MsgBox "·«Ì„ﬂ‰ «‰  ﬂÊ‰ ﬁÌ„… «·œ›⁄ …»«·”«·»"
              Else
              MsgBox "Can't enter Negative Number "
              End If
              .TextMatrix(i, .ColIndex("TransPayedValue")) = 0
     Exit Sub
      End If
      
End If
Next i
End With

    
    With Me.VSFlexGrid1
        For i = .FixedRows To .rows - 1
                If .cell(flexcpChecked, i, .ColIndex("payed")) = flexChecked Then
                IntCounter = IntCounter + 1
                .TextMatrix(i, .ColIndex("Ser")) = IntCounter
              If val(.TextMatrix(i, .ColIndex("TransPayedValue"))) > val(.TextMatrix(i, .ColIndex("RemainingValue"))) And val(.TextMatrix(i, .ColIndex("TransPayedValue"))) <> 0 Then
              If SystemOptions.UserInterface = ArabicInterface Then
              MsgBox "·«Ì„ﬂ‰ «‰  ﬂÊ‰ ﬁÌ„… «·œ›⁄… «ﬂ»— „‰ «·„ »ﬁÌ"
              Else
              MsgBox "Can Not PaymentValue Larger Than Total Value "
              End If
              .TextMatrix(i, .ColIndex("TransPayedValue")) = 0
              Exit Sub
              End If
           Sm = Sm + val(.TextMatrix(i, .ColIndex("TransPayedValue")))
           End If
           Next i
  
    End With
   XPTxtVal.text = Sm
   XPTxtVal.Enabled = False
End Sub
Sub BillCustomer(Optional Ind As Integer = 0)
If val(DCboCashType.ListIndex) = 0 Then
Dim Msg As String
If val(DBCboClientName.BoundText) = 0 Then
If SystemOptions.UserInterface = ArabicInterface Then
MsgBox "Ì—ÃÏ ≈Œ Ì«— «·⁄„Ì· «Ê·«"
Else
MsgBox "Please Select Customer"
End If
DBCboClientName.SetFocus
Exit Sub
Else
If Ind = 0 Then
Frame12.Visible = True
End If
If Me.TxtModFlg.text <> "R" Then
If Ind = 0 Then
XPTxtVal.text = 0
End If
If Me.TxtModFlg.text = "N" Then
RetriveBillBuy val(DBCboClientName.BoundText)
End If

If Me.TxtModFlg.text = "E" And (FlgBillBuy = True Or VSFlexGrid1.rows = 1) Then
RetriveBillBuy val(DBCboClientName.BoundText)
End If
End If
End If
End If
End Sub
Function AutoCalculate() As Boolean
Dim i As Integer
Dim NetValu As Double
Dim TempValu As Double
Dim RemainValu As Double
NetValu = val(XPTxtVal.text)
With VSFlexGrid1
For i = 1 To .rows - 1
RemainValu = val(.TextMatrix(i, .ColIndex("RemainingValue")))
If NetValu >= RemainValu Then
TempValu = RemainValu
NetValu = NetValu - TempValu
Else
TempValu = NetValu
NetValu = 0
End If
If TempValu > 0 Then
  .TextMatrix(i, .ColIndex("TransPayedValue")) = TempValu
  .cell(flexcpChecked, i, .ColIndex("payed")) = flexChecked
   End If
Next i
End With
If NetValu <> 0 Then
AutoCalculate = False
Else
AutoCalculate = True
End If
End Function

Private Sub Command5_Click()
Frame20.Visible = False
End Sub

Private Sub CommdiscountAccount_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF3 Then
    DCAccounts.text = ""
        Unload Account_search
        Account_search.show
        Account_search.case_id = 260815
            
    End If
    
    
End Sub

Private Sub commdiscounttype_Change()
calcnet
If commdiscounttype.ListIndex = 0 Then
Commdiscountvalue.text = 0
CommdiscountAccount.Enabled = False
CommdiscountAccount.BoundText = ""
Else
CommdiscountAccount.Enabled = True

If Me.TxtModFlg = "N" Or Me.TxtModFlg = "E" Then

         CommdiscountAccount.BoundText = get_account_code_branch(52, 0)
        
 
        
End If
End If
End Sub

Private Sub commdiscounttype_Click()
commdiscounttype_Change
End Sub

Private Sub Commdiscountvalue_Change()
calcnet
End Sub
Function calcnet()

If commdiscounttype.ListIndex = 1 Then
commvalue = val(Commdiscountvalue.text)
ElseIf commdiscounttype.ListIndex = 2 Then

commvalue = val(Commdiscountvalue.text) * val(XPTxtVal.text) / 100
Else
commvalue = 0
End If

Commdiscountvalue1.text = commvalue

End Function
Private Sub DBCboClientName_Change()
    Dim pstate As Integer
    TxtCustCode.text = ""

    If (Me.DCboCashType.ListIndex = 5 Or Me.DCboCashType.ListIndex = 11) And Me.TxtModFlg.text = "N" Or Me.TxtModFlg.text = "E" Then
        FillGridWithData val(Me.DBCboClientName.BoundText), TxtNoteSerial.text
     End If

    Dim DefaultSalesPersonId As Integer
    Dim Fullcode As String

    GetCustomersDetail val(DBCboClientName.BoundText), DefaultSalesPersonId, Fullcode

 
    If DBCboClientName.BoundText = "" Then Exit Sub
 
    If 1 = 1 Then
  Fullcode = ""
  
       If SystemOptions.AllowAcceleratepayment = True And CheckCustomer(val(DBCboClientName.BoundText)) = True Then
            Frame20.Visible = True
       End If
        'Dim fullcode As String
      If Me.DCboCashType.ListIndex = 0 Then
       
       
     '  If Me.TxtModFlg = "N" Or Me.TxtModFlg = "E" Then
        GetCustomersDetail val(DBCboClientName.BoundText), DefaultSalesPersonId, Fullcode
        TxtCustCode.text = Fullcode
  
If Me.TxtModFlg.text <> "R" Then
        DcEmp.BoundText = DefaultSalesPersonId
    End If
      '  End If
        
        
        ElseIf (Me.DCboCashType.ListIndex = 5 Or Me.DCboCashType.ListIndex = 11) Then
        
       
        GetProjectsDetail val(DBCboClientName.BoundText), DefaultSalesPersonId, Fullcode
       TxtCustCode.text = Fullcode

        
        
        
        End If
        
            If Me.TxtModFlg.text = "N" Or Me.TxtModFlg.text = "E" Then
            
            DcEmp.BoundText = DefaultSalesPersonId
            End If
            
        
        
        
         
        If SystemOptions.CustomerhavethreeAccounts = True Then ' «·⁄„·«¡ ·Â« À·«À Õ”«»« 
        
                            If CboPayMentType.ListIndex = 0 Or CboPayMentType.ListIndex = 5 Then '‰ﬁœÌ
                                               If Option3.value = True Then 'œ›⁄«  „ﬁœ„…
                                                        Me.DcboCreditSide.BoundText = ModAccounts.GetMyAccountCode("TblCustemers", "CusID", val(Me.DBCboClientName.BoundText), "Account_code2")
                                                        
                                                        If Me.DcboCreditSide.BoundText = "" Then
                                                        Me.DcboCreditSide.BoundText = ModAccounts.GetMyAccountCode("TblCustemers", "CusID", val(Me.DBCboClientName.BoundText))
                                                        End If
                                             Else
                                                                 Me.DcboCreditSide.BoundText = ModAccounts.GetMyAccountCode("TblCustemers", "CusID", val(Me.DBCboClientName.BoundText))
                                             End If
                               
                            ElseIf CboPayMentType.ListIndex = 1 Then '‘Ìﬂ
                            
                                                If Option3.value = True Then 'œ›⁄«  „ﬁœ„…
                                                    Me.DcboCreditSide.BoundText = ModAccounts.GetMyAccountCode("TblCustemers", "CusID", val(Me.DBCboClientName.BoundText), "Account_code1") '‘Ìﬂ«   Õ  «· Õ’Ì·
                                                         
                                                         
                                                         If Me.DcboCreditSide.BoundText = "" Then
                                                        Me.DcboCreditSide.BoundText = ModAccounts.GetMyAccountCode("TblCustemers", "CusID", val(Me.DBCboClientName.BoundText))
                                                        End If
                                                        
                                             Else
                                                                 Me.DcboCreditSide.BoundText = ModAccounts.GetMyAccountCode("TblCustemers", "CusID", val(Me.DBCboClientName.BoundText), "Account_code1")
                                             End If
                                     
                             ElseIf CboPayMentType.ListIndex = 2 Then 'ÕÊ«·… '
                                                If Option3.value = True Then 'œ›⁄«  „ﬁœ„…
                                                        Me.DcboCreditSide.BoundText = ModAccounts.GetMyAccountCode("TblCustemers", "CusID", val(Me.DBCboClientName.BoundText), "Account_code2")
                                                        
                                                             If Me.DcboCreditSide.BoundText = "" Then
                                                        Me.DcboCreditSide.BoundText = ModAccounts.GetMyAccountCode("TblCustemers", "CusID", val(Me.DBCboClientName.BoundText))
                                                        End If
                                                        
                                             Else
                                                                 Me.DcboCreditSide.BoundText = ModAccounts.GetMyAccountCode("TblCustemers", "CusID", val(Me.DBCboClientName.BoundText))
                                             End If
                              ElseIf CboPayMentType.ListIndex = 3 Then '‘Ìﬂ „”œœ '
                                                                    If Option3.value = True Then 'œ›⁄«  „ﬁœ„…
                                                        Me.DcboCreditSide.BoundText = ModAccounts.GetMyAccountCode("TblCustemers", "CusID", val(Me.DBCboClientName.BoundText), "Account_code2")
                                                             If Me.DcboCreditSide.BoundText = "" Then
                                                        Me.DcboCreditSide.BoundText = ModAccounts.GetMyAccountCode("TblCustemers", "CusID", val(Me.DBCboClientName.BoundText))
                                                        End If
                                                        
                                             Else
                                                                 Me.DcboCreditSide.BoundText = ModAccounts.GetMyAccountCode("TblCustemers", "CusID", val(Me.DBCboClientName.BoundText))
                                             End If
                              End If
                             
'
        Else
        '«·⁄„·«¡ ·Â„ Õ”«» Ê«Õœ ›ﬁÿ
        


        If Option3.value = True And SystemOptions.DealingWithPrepayAccount = True Then
     
       
                Me.DcboCreditSide.BoundText = get_account_code_branch(158, val(dcBranch.BoundText))
'
         Else
         Me.DcboCreditSide.BoundText = ModAccounts.GetMyAccountCode("TblCustemers", "CusID", val(Me.DBCboClientName.BoundText))
        End If

        End If
        

        If DCboCashType.ListIndex = 5 Then 'Õ«·… «·„‘«—Ì⁄
                                        
       If Option4.value = True Then ' ⁄„Ì· ‰Â«∆Ì
                                        
        If SystemOptions.CustomerhavethreeAccounts = True Then ' «·⁄„·«¡ ·Â« À·«À Õ”«»« 
        
                            If CboPayMentType.ListIndex = 0 Or CboPayMentType.ListIndex = 5 Then '‰ﬁœÌ
                                                                    If Option3.value = True Then 'œ›⁄Â „ﬁœ„…
                                                                           Me.DcboCreditSide.BoundText = get_project_customer_account(val(DBCboClientName.BoundText), "Account_Code2") 'œ›⁄«  „ﬁœ„…
                                                                      Else
                                                                           Me.DcboCreditSide.BoundText = get_project_customer_account(val(DBCboClientName.BoundText), "Account_Code") ' Ã«—Ì
                                                                      End If
                               
                            ElseIf CboPayMentType.ListIndex = 1 Then '‘Ìﬂ
                            
                                                                If Option3.value = True Then 'œ›⁄Â „ﬁœ„…
                                                                    Me.DcboCreditSide.BoundText = get_project_customer_account(val(DBCboClientName.BoundText), "Account_Code1") 'œ›⁄«  „ﬁœ„…
                                                                      Else
                                                                          Me.DcboCreditSide.BoundText = get_project_customer_account(val(DBCboClientName.BoundText), "Account_Code1") '  Õ  «· Õ’Ì·
                                                                      End If
                                     
                             ElseIf CboPayMentType.ListIndex = 2 Then 'ÕÊ«·… '
                                               If Option3.value = True Then 'œ›⁄Â „ﬁœ„…
                                                                    Me.DcboCreditSide.BoundText = get_project_customer_account(val(DBCboClientName.BoundText), "Account_Code2") 'œ›⁄«  „ﬁœ„…
                                                                      Else
                                                                          Me.DcboCreditSide.BoundText = get_project_customer_account(val(DBCboClientName.BoundText), "Account_Code") ' Ã«—Ì
                                                                      End If
                              ElseIf CboPayMentType.ListIndex = 3 Then '‘Ìﬂ „”œœ '
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
        
                            If CboPayMentType.ListIndex = 0 Or CboPayMentType.ListIndex = 5 Then '‰ﬁœÌ
                                                If Option3.value = True Then 'œ›⁄Â „ﬁœ„…
                                                                    Me.DcboCreditSide.BoundText = get_project_customer_account(val(DBCboClientName.BoundText), "Account_Code2", 1) 'œ›⁄«  „ﬁœ„…
                                                                      Else
                                                                          Me.DcboCreditSide.BoundText = get_project_customer_account(val(DBCboClientName.BoundText), "Account_Code", 1) ' Ã«—Ì
                                                                      End If
                               
                            ElseIf CboPayMentType.ListIndex = 1 Then '‘Ìﬂ
                            
                                                                If Option3.value = True Then 'œ›⁄Â „ﬁœ„…
                                                                    Me.DcboCreditSide.BoundText = get_project_customer_account(val(DBCboClientName.BoundText), "Account_Code2", 1) 'œ›⁄«  „ﬁœ„…
                                                                      Else
                                                                          Me.DcboCreditSide.BoundText = get_project_customer_account(val(DBCboClientName.BoundText), "Account_Code1", 1) '  Õ  «· Õ’Ì·
                                                                      End If
                                     
                             ElseIf CboPayMentType.ListIndex = 2 Then 'ÕÊ«·… '
                                               If Option3.value = True Then 'œ›⁄Â „ﬁœ„…
                                                                    Me.DcboCreditSide.BoundText = get_project_customer_account(val(DBCboClientName.BoundText), "Account_Code2", 1) 'œ›⁄«  „ﬁœ„…
                                                                      Else
                                                                          Me.DcboCreditSide.BoundText = get_project_customer_account(val(DBCboClientName.BoundText), "Account_Code", 1) ' Ã«—Ì
                                                                      End If
                              ElseIf CboPayMentType.ListIndex = 3 Then '‘Ìﬂ „”œœ '
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
                                        
                          '(((((((((((((((((((((((((((((((((((((((
                                        
            
                        '(((((((((((((((((((((((((((((((((((((((((((((((((((((((((((((((
       End If
    End If

End Sub
 
 
Private Sub DBCboClientName_Click(Area As Integer)
    'WriteCustomerBalPublic
End Sub

Private Sub DBCboClientName_KeyUp(KeyCode As Integer, _
                                  Shift As Integer)

    If DCboCashType.ListIndex = 0 Then
        If KeyCode = vbKeyF3 Then
         FrmCustemerSearch.SearchType = 3
            FrmCustemerSearch.show vbModal
           
        End If

    ElseIf DCboCashType.ListIndex = 1 Then

        If KeyCode = vbKeyF3 Then

         FrmCompanySearch.lblSearchtype.Caption = 2
         FrmCompanySearch.show vbModal
       
        End If

   ElseIf DCboCashType.ListIndex = 5 Then

        If KeyCode = vbKeyF3 Then
         FrmProjectSearch.lblSearchtype.Caption = 1
             FrmProjectSearch.show vbModal
           
        End If
  
   ElseIf Me.CboPayMentType.ListIndex = 4 Then
   
  
   ElseIf DCboCashType.ListIndex = 6 Then

 
  ElseIf Me.DCboCashType.ListIndex = 7 Then
        If KeyCode = vbKeyF3 Then
            Account_search.show
            Account_search.case_id = 260817
        End If

  
    End If

End Sub

Private Sub DCAccounts_Change()

    If Me.TxtModFlg.text = "N" Or Me.TxtModFlg.text = "E" Then
        '   Me.DcboCreditSide.BoundText = ModAccounts.GetMyAccountCode("TblRevenuesTypes", "RevenuesID", Val(Me.DcboRevenuesTypes.BoundText))
        Me.DcboCreditSide.BoundText = DCAccounts.BoundText
       If DCboCashType.ListIndex = 7 Then
            TxtCustCode.text = getAccountSerial_Code("Account_Serial", "Account_Code", DCAccounts.BoundText)
        End If
        'If Me.TxtModFlg.Text <> "R" Then
      
 
  
  
    End If

End Sub

Private Sub DCAccounts_KeyUp(KeyCode As Integer, _
                             Shift As Integer)

    If KeyCode = vbKeyF3 Then
    DCAccounts.text = ""
     '   Unload Account_search
        Account_search.show
        Account_search.case_id = 260817
            
    End If

End Sub

Private Sub DcbAccount_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF3 Then
        Account_search.show
        Account_search.case_id = 260816
    End If

End Sub

Private Sub DcbAccount_Validate(Cancel As Boolean)
'TxtAccount.Text = DcbAccount.BoundText
End Sub

Private Sub DcboBankName_Click(Area As Integer)

    If DcboBankName.BoundText = "" Then Exit Sub
    Dim RsSavRec As ADODB.Recordset
    Dim My_SQL As String
    Dim Account_Code_dynamic As String

    If Me.TxtModFlg.text = "N" Or Me.TxtModFlg.text = "E" Then
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

        If CboPayMentType.ListIndex = 2 Or CboPayMentType.ListIndex = 3 Then
                     
            Me.DcboDebitSide.BoundText = RsSavRec.Fields("Account_Code").value
                    
        End If

    End If

End Sub

Private Sub DcboBankName_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyF3 Then
                    FrmExpensesSearch.Indx = 1
                    FrmExpensesSearch.show
                    FrmExpensesSearch.Indx = 1
                    FrmExpensesSearch.RetrunType = 2510
    End If
End Sub

Private Sub DcboBox_Change()

    If Me.TxtModFlg.text = "N" Or Me.TxtModFlg.text = "E" Then
        Me.DcboDebitSide.BoundText = ModAccounts.GetMyAccountCode("TblBoxesData", "BoxID", val(Me.DcboBox.BoundText))
    End If

End Sub

Public Sub DcboBox_KeyUp(KeyCode As Integer, Shift As Integer)
 If KeyCode = vbKeyF3 Then
                    FrmExpensesSearch.Indx = 2
                    FrmExpensesSearch.show
                    FrmExpensesSearch.Indx = 2
                    FrmExpensesSearch.RetrunType = 2509
    End If
End Sub

Private Sub DCboCashType_Change()

    On Error GoTo ErrTrap
    Frame2.Enabled = False
    Dim StrSQL As String
    Dim intDef As Integer
TxtContractNo.Visible = False
lbl(53).Visible = False
C1Elastic1.Visible = False
Command9.Visible = False
Frame20.Visible = False
 
 TxtVATValue.Visible = False
 txtVat2.Visible = False

 lbl(65).Visible = False
 TxtBillTransNo.Visible = False
 lbl(67).Visible = False
 TxtBillTransID.Visible = False
 TxtBillMaintNo.Visible = False
 txtContainerNo.Visible = False
    Select Case DCboCashType.ListIndex

        Case 0, 11
            Dcombos.GetCustomersSuppliers 1, Me.DBCboClientName, False
            Me.DBCboClientName.Visible = True
            Me.DcboRevenuesTypes.Visible = False
            TxtVATValue.Visible = True
            txtVat2.Visible = True
            lbl(65).Visible = True
            DCEmployee.Visible = False
            DCAccounts.Visible = False
            ChkTrans.Visible = True
            Fra(0).Visible = True
          ' Command9.Visible = True

            If SystemOptions.UserInterface <> EnglishInterface Then
                Me.lbl(3).Caption = "«”„ «·⁄„Ì·"
            Else
                Me.lbl(3).Caption = "Customer Name"
            End If
        
            Me.lbl(13).Visible = True
            Me.LblLink.Visible = True

        Case 1
            Dcombos.GetCustomersSuppliers 2, Me.DBCboClientName, False
            Me.DBCboClientName.Visible = True
            Me.DcboRevenuesTypes.Visible = False
            DCEmployee.Visible = False
            DCAccounts.Visible = False
            ChkTrans.Visible = True
            Fra(0).Visible = True

            If SystemOptions.UserInterface <> EnglishInterface Then
                Me.lbl(3).Caption = "«”„ «·„Ê—œ"
            Else
                Me.lbl(3).Caption = "Vendor Name"
            End If
            
            TxtVATValue.Visible = True
            txtVat2.Visible = True
            lbl(65).Visible = True
            
        
            Me.lbl(13).Visible = True
            Me.LblLink.Visible = True

        Case 2
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
                
            TxtVATValue.Visible = True
            txtVat2.Visible = True
            lbl(65).Visible = True
            
            Me.lbl(13).Visible = True
            Me.LblLink.Visible = True

        Case 3
            '≈Ì—«œ«  ≈Œ—Ï
            Me.DBCboClientName.Visible = False
            Me.DcboRevenuesTypes.Visible = True
            Me.ChkTrans.Visible = False
            DBCboClientName.Visible = False
            DCEmployee.Visible = False
            DCAccounts.Visible = False
            Fra(0).Visible = False
                TxtVATValue.Visible = True
                txtVat2.Visible = True
            lbl(65).Visible = True
        
            If SystemOptions.UserInterface <> EnglishInterface Then
                Me.lbl(3).Caption = "‰Ê⁄ «·«Ì—«œ"
            Else
                Me.lbl(3).Caption = "RVN Type"
            End If
                
            Me.lbl(13).Visible = False
            Me.LblLink.Visible = False
        
        Case 4
            Dcombos.GetCustomersSuppliers 1, Me.DBCboClientName, True
            Me.DBCboClientName.Visible = True
            Me.DcboRevenuesTypes.Visible = False
            DCEmployee.Visible = False
            DCAccounts.Visible = False
            ChkTrans.Visible = True
            Fra(0).Visible = True

            If SystemOptions.UserInterface <> EnglishInterface Then
                Me.lbl(3).Caption = "«”„ «·⁄„Ì·"
            Else
                Me.lbl(3).Caption = "Customer Name"
            End If
        
            Me.lbl(13).Visible = True
            Me.LblLink.Visible = True
        
        Case 5
            Dim My_SQL As String
            If SystemOptions.UserInterface = ArabicInterface Then
            My_SQL = "  select id,Project_name from projects where not(REVENUE_account is null) and Not (Project_name is null) and Project_name <>N'""' "
            StrSQL = StrSQL & "  AND      branch_no in(" & Current_branchSql & ")"
           StrSQL = StrSQL & "  order by Project_name"
            Else
            My_SQL = "  select id,Project_nameE from projects where not(REVENUE_account is null) and Not (Project_nameE is null) and Project_nameE <>N'""' "
             StrSQL = StrSQL & "  AND      branch_no in(" & Current_branchSql & ")"
            StrSQL = StrSQL & " order by Project_nameE"
            End If
            fill_combo Me.DBCboClientName, My_SQL
             TxtVATValue.Visible = True
             txtVat2.Visible = True
            lbl(65).Visible = True
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
         
            Dcombos.GetAccountingCodes Me.DCAccounts, True
            DCAccounts.Visible = True
            Me.DCEmployee.Visible = False
            Me.DcboRevenuesTypes.Visible = False
            DBCboClientName.Visible = False
            TxtVATValue.Visible = True
            txtVat2.Visible = True
            lbl(65).Visible = True
            ChkTrans.Visible = True

            '   Fra(0).Visible = True
            If SystemOptions.UserInterface <> EnglishInterface Then
                Me.lbl(3).Caption = "«”„ «·Õ”«»"
            Else
                Me.lbl(3).Caption = "Accounts Nam  "
            End If
        
            '  Me.lbl(13).Visible = True
            '      Me.LblLink.Visible = True
   Case 12
           txtContainerNo.Visible = True

            Dcombos.GetCustomersSuppliers 1, Me.DBCboClientName, False
            Me.DBCboClientName.Visible = True
            lbl(67).Visible = True
            
            If SystemOptions.UserInterface <> EnglishInterface Then
                Me.lbl(3).Caption = "«”„ «·⁄„Ì·"
                lbl(67).Caption = "—ﬁ„ ⁄ﬁœ «·Õ«ÊÌ…"
                
            Else
                Me.lbl(3).Caption = "Customer Name"
                lbl(67).Caption = "Container No"
            End If
        
            Me.lbl(13).Visible = True
            Me.LblLink.Visible = True

Case 88 '  „‰ ⁄ﬁœ
            Dcombos.GetCustomersSuppliers 1, Me.DBCboClientName, False
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
            Me.LblLink.Visible = True
TxtContractNo.Visible = True
lbl(53).Visible = True
   Case 99
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
    'End If
           Case 8
            Dcombos.GetCustomersSuppliers 1, Me.DBCboClientName, False
            Me.DBCboClientName.Visible = True
            Me.DcboRevenuesTypes.Visible = False
            lbl(65).Visible = True
            DCEmployee.Visible = False
            DCAccounts.Visible = False
          ' Command9.Visible = True
            If SystemOptions.UserInterface <> EnglishInterface Then
                Me.lbl(3).Caption = "«”„ «·⁄„Ì·"
            Else
                Me.lbl(3).Caption = "Customer Name"
            End If
        
            Me.lbl(13).Visible = True
            Me.LblLink.Visible = True
            TxtBillTransNo.Visible = True
            lbl(67).Visible = True
            XPTxtVal.Enabled = False
            DBCboClientName.Enabled = False
            TxtCustCode.Enabled = False
          '  TxtBillTransID.Visible = True
            Case 9
            Dcombos.GetCustomersSuppliers 1, Me.DBCboClientName, False
            Me.DBCboClientName.Visible = True
            Me.DcboRevenuesTypes.Visible = False
            lbl(65).Visible = True
            DCEmployee.Visible = False
            DCAccounts.Visible = False
            If SystemOptions.UserInterface <> EnglishInterface Then
                Me.lbl(3).Caption = "«”„ «·⁄„Ì·"
            Else
                Me.lbl(3).Caption = "Customer Name"
            End If
        
            Me.lbl(13).Visible = True
            Me.LblLink.Visible = True
            TxtBillMaintNo.Visible = True
            lbl(67).Visible = True
            XPTxtVal.Enabled = False
            DBCboClientName.Enabled = False
            TxtCustCode.Enabled = False
          '  TxtBillTransID.Visible = True
            Case 10
            Dcombos.GetCustomersSuppliers 1, Me.DBCboClientName, False
            Me.DBCboClientName.Visible = True
            Me.DcboRevenuesTypes.Visible = False
          '  DBCboClientName.Enabled = True
            
            lbl(65).Visible = True
            DCEmployee.Visible = False
            DCAccounts.Visible = False
            If SystemOptions.UserInterface <> EnglishInterface Then
                Me.lbl(3).Caption = "«”„ «·⁄„Ì·"
            Else
                Me.lbl(3).Caption = "Customer Name"
            End If
            
            Me.lbl(13).Visible = True
            Me.LblLink.Visible = True
            TxtBillMaintNo.Visible = True
            lbl(67).Visible = True
            lbl(67).Caption = "—ﬁ„ «–‰ «·«’·«Õ"
            
            DBCboClientName.Enabled = True
            TxtCustCode.Enabled = False
            TxtVATValue.Visible = True
            txtVat2.Visible = True
            XPTxtVal.Enabled = True
    End Select
CalCulteVAT
    cSearchDcbo.Refresh
    Exit Sub
ErrTrap:
End Sub

Private Sub DCboCashType_Click()
    DCboCashType_Change
End Sub

Private Sub DcboCreditSide_Change()

    WriteCustomerBalPublic Me.DcboCreditSide.BoundText, Balance, balanceString
    LblLink.Caption = balanceString
   If Me.TxtModFlg.text <> "R" Then
    TxtCurrentBalance.text = Balance
   End If
End Sub

Private Sub DcboRevenuesTypes_Change()

    If Me.TxtModFlg.text = "N" Or Me.TxtModFlg.text = "E" Then
        Me.DcboCreditSide.BoundText = ModAccounts.GetMyAccountCode("TblRevenuesTypes", "RevenuesID", val(Me.DcboRevenuesTypes.BoundText))
    End If

End Sub

Private Sub Dcbranch_Click(Area As Integer)
    TxtNoteSerial.text = ""
    TxtNoteSerial1.text = ""
End Sub

Private Sub DcbUnitNo_Change()
If Me.TxtModFlg <> "R" And Me.TxtModFlg <> "" Then
Dim str As String
str = checkDepositeRent(val(DcbUnitNo.BoundText), XPDtbTrans)
If str <> "" Then
MsgBox str, vbInformation
End If

End If
End Sub

Private Sub DcbUnitType_Change()
Dim Dcombos As ClsDataCombos
Dim idd As Long
Dim idd1 As Long
   Set Dcombos = New ClsDataCombos

If val(DcbIqara.BoundText) > 0 Then
idd = val(DcbIqara.BoundText)

idd1 = val(DcbUnitType.BoundText)
If Me.TxtModFlg = "R" Then
Dcombos.GetIqarUnit idd, idd1, Me.DcbUnitNo, "R"
Else
Dcombos.GetIqarUnit idd, idd1, Me.DcbUnitNo
End If
End If
End Sub

Private Sub DcbUnitType_Click(Area As Integer)
DcbUnitType_Change
End Sub

Private Sub DcChequeBox_Change()

    If Me.TxtModFlg.text = "N" Or Me.TxtModFlg.text = "E" Then
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

         If val(Me.DcEmp.BoundText) = 0 Then Exit Sub
           Me.TxtEmployeeID.text = get_EMPLOYEE_Data(val(Me.DcEmp.BoundText), "Fullcode")

End Sub

Private Sub DcEmployee_Change()
 
    If Me.TxtModFlg.text = "N" Or Me.TxtModFlg.text = "E" Then
        '   Me.DcboCreditSide.BoundText = ModAccounts.GetMyAccountCode("TblRevenuesTypes", "RevenuesID", Val(Me.DcboRevenuesTypes.BoundText))
        Me.DcboCreditSide.BoundText = get_EMPLOYEE_Account(val(DCEmployee.BoundText), "Account_Code")
        
        TxtCustCode.text = getemployeeCode(val(DCEmployee.BoundText))
       
       
       ' TxtCustCode.text = val(dcEmployee.BoundText)
    
    
     
    End If

End Sub

Private Sub dcCar_Change()

    GetDriverInformation (val(DCCar.BoundText))

End Sub

Private Sub dcCar_Click(Area As Integer)
    GetDriverInformation (val(DCCar.BoundText))

End Sub

Function GetDriverInformation(ID As Integer)

    If Me.TxtModFlg.text = "N" Or Me.TxtModFlg.text = "E" Then
        Dim sql As String
        Dim rs As New ADODB.Recordset
 
        sql = " SELECT    * "
        sql = sql & " from dbo.TblCarsData"
        sql = sql & " Where (id = " & ID & ") "

        rs.Open sql, Cn, adOpenStatic, adLockReadOnly, adCmdText

        If rs.RecordCount > 0 Then
            DCDriver.BoundText = IIf(IsNull(rs("Emp_id").value), 0, rs("Emp_id").value)
                  
        Else
            DcEmp = 0
               
        End If

    End If

End Function

Private Sub DCEmployee_KeyUp(KeyCode As Integer, Shift As Integer)


    If KeyCode = vbKeyF3 Then
        FrmEmployeeSearch.lbltype = 35
       ' Set FrmEmployeeSearch.RetrunFrm = Me

        FrmEmployeeSearch.show
  
    End If
    
    
End Sub

Private Sub DCPreFix_Change()
If Me.TxtModFlg.text = "N" Or Me.TxtModFlg.text = "E" Then
   TxtNoteSerial.text = ""
    TxtNoteSerial1.text = ""
    End If
    
End Sub

Private Sub DCPreFix_Click(Area As Integer)
   TxtNoteSerial.text = ""
    TxtNoteSerial1.text = ""
End Sub

Private Sub Form_Load()
    On Error GoTo ErrTrap
    Me.left = (mdifrmmain.Width - Me.Width) / 2
    Me.top = (mdifrmmain.Height - Me.Height) / 2 - 500

    If SystemOptions.DateOpt = 1 Then
        Txt_DateHigri.Visible = True
    
    End If

    If mdifrmmain.TransporterMain.Visible = False Then
        lbl(49).Visible = False
        lbl(50).Visible = False
        DCCar.Visible = False
        DCDriver.Visible = False

    End If

  If mdifrmmain.MnuProjects.Visible = True Then
  XPTab301.TabVisible(1) = True
  Else
  XPTab301.TabVisible(1) = False
  End If

  If mdifrmmain.AssetsMngBase.Visible = True Then
  XPTab301.TabVisible(2) = True
  Else
  XPTab301.TabVisible(2) = False
  End If
  If SystemOptions.PreFixCanNotEdit = True Then
  DCPreFix.Enabled = False
  Else
  DCPreFix.Enabled = True
  End If

    ScreenNameArabic = "«·„ﬁ»Ê÷« "
    ScreenNameEnglish = "Cash Receipt Voucher"
    RegisterLogInOut Me.Name, ScreenNameArabic, ScreenNameEnglish, "1", 4
 
    Dim StrSQL As String
    Dim Msg As String
    Set Dcombos = New ClsDataCombos
'    StrSQL = "  SELECT code ,account_name FROM markaas_taklefa  WHERE level=3 and NOT(account_no IS NULL)  "
'    fill_combo Me.DcCostCenter, StrSQL

'    Dim Dcombos As ClsDataCombos
'Set Dcombos = New ClsDataCombos
    Dcombos.GetCostCenter DcCostCenter
    Dcombos.GetSalesRepData Me.DcEmp
    Dcombos.GetCars Me.DCCar
    Dcombos.GetEmployees Me.DCDriver, , True
    Dcombos.GetIqar DcbIqara
    Dcombos.getAkarUnit Me.DcbUnitType
    Dcombos.GetPrefix2 Me.DCPreFix, 2, 0
    Dcombos.GetAccountingCodes Me.DcbAccount, True, False
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
    DCboCashType.AddItem "„‰ ⁄„Ì·"
    DCboCashType.AddItem "„‰ „Ê—œ"
    DCboCashType.AddItem "„ﬁ«Ê· »«ÿ‰"
    DCboCashType.AddItem "≈Ì—«œ«  ≈Œ—Ï"
    DCboCashType.AddItem "„œ›Ê⁄«  „ﬁœ„Â"
    DCboCashType.AddItem "„‘—Ê⁄"
    DCboCashType.AddItem "„‰ „ÊŸ›"
    DCboCashType.AddItem "„‰ Õ”«»"
    DCboCashType.AddItem "„‰ ›Ê« Ì— «·‰ﬁ·Ì« "
    DCboCashType.AddItem "„‰ ›« Ê—… «·’Ì«‰…"
    DCboCashType.AddItem "»‰«¡« ⁄·Ï ﬂ«—  ’Ì«‰…"
    DCboCashType.AddItem "„‰ ⁄œ… „” Œ·’« "
    DCboCashType.AddItem "»‰«¡« ⁄·Ï ⁄ﬁœ Õ«ÊÌ« "


 
    With Me.CboPayMentType
        .Clear
        .AddItem "‰ﬁœÌ"
        .AddItem "‘Ìﬂ"
        .AddItem "ÕÊ«·Â »‰ﬂÌÂ"
        .AddItem "  ‘Ìﬂ „Õ’· "
        .AddItem "Õ”«»"
     If SystemOptions.AllowAccountMultyPayed = True Then
        .AddItem "„ ⁄œœ"
     End If
     
    End With

    With Me.commdiscounttype
        .Clear
        .AddItem "»·«"
        .AddItem "ﬁÌ„…"
        .AddItem "‰”»…"
        
    End With
    
  With CboStatus
  .Clear
  .AddItem "„›⁄·"
  .AddItem "·œÏ «·⁄„Ì·"
  .AddItem "„·€Ì"
  .AddItem "„›ﬁÊœ"
  
  End With
    Dcombos.GetUsers Me.DCboUserName
If SystemOptions.AllowHideAssest = False Then
    Dcombos.GetBoxes Me.DcboBox
    Else
    Dcombos.GetBoxes Me.DcboBox, 0
    End If
    Dcombos.GetChequeBox Me.DCChequeBox

    Dcombos.GetBanks Me.DcboBankName
    Dcombos.GetCustomersSuppliers 0, Me.DBCboClientName, False
    Dcombos.GetRevenuesTypes Me.DcboRevenuesTypes
    'Set cSearchDcbo = New clsDCboSearch
    'Set cSearchDcbo.Client = Me.DBCboClientName

    Dcombos.GetAccountingCodes Me.DcboDebitSide
    Dcombos.GetAccountingCodes Me.CommdiscountAccount
    
    Dcombos.GetAccountingCodes Me.DcboCreditSide
    Dcombos.GetBranches Me.dcBranch
       If SystemOptions.BranchCanNotEdit = True Then
        dcBranch.Enabled = False
        Text2.Enabled = False
       Else
       Text2.Enabled = True
       dcBranch.Enabled = True
      End If
    If SystemOptions.usertype <> UserAdminAll Then
        If SystemOptions.BranchCanNotEdit = True Then
        dcBranch.Enabled = False
        Text2.Enabled = False
       Else
       Text2.Enabled = True
       dcBranch.Enabled = True
      End If
    End If

    Set rs = New ADODB.Recordset
    'StrSQL = "select * From Notes where NoteType=4 and   displayed is null Order By NoteID"
    StrSQL = "select * From Notes where NoteType=4    AND branch_no in(" & Current_branchSql & ")"
     StrSQL = StrSQL & " and CashingType<=12 and akarid is Null"
'StrSQL = StrSQL & " and CashingType<=11 and akarid is Null"

    'If SystemOptions.usertype <> UserAdminAll Then
    '    StrSQL = StrSQL & " AND   branch_no=" & Current_branch
    'End If
            
    If SystemOptions.usertype <> UserAdmin Then
 
                            If SystemOptions.FixedCustomer = 1 Then
                              StrSQL = StrSQL & " and  UserID = " & user_id
                               End If
                    
    If SystemOptions.BranchCanNotEdit = True Then
        dcBranch.Enabled = False
        Text2.Enabled = False
       Else
       Text2.Enabled = True
       dcBranch.Enabled = True
      End If
      
      
    End If
    
    StrSQL = StrSQL & " and  displayed is null Order By NoteID"

    rs.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText

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
    
    If SystemOptions.CanEditOnlyPayMethod Then
        isFormFirstRun = True
        Cmd_Click (0)
        
        GetDefaultEnabled False
    End If
    XPBtnMove_Click 2
    Me.TxtModFlg.text = "R"
    WriteInfo
      
    Dim My_SQL As String

    'My_SQL = "  select expanses_account,account_name from projects  where not (account_no is null)"
    My_SQL = "  select id,Project_name from projects where not(REVENUE_account is null)" '
    fill_combo DCproject, My_SQL
'XPDtbTrans.SetFocus
    
    If OPEN_NEW_SCREEN = True Then
        Cmd_Click (0)
    End If
                If SystemOptions.CanEditOnlyPayMethod And (Me.TxtModFlg = "E" Or Me.TxtModFlg = "R") Then
   
        Ele(12).Enabled = False
        Frame12.Enabled = False
        Ele(0).Enabled = False
             XPTab301.Enabled = False
    End If
    isFormFirstRun = True
    
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

 

Private Sub Grid_AfterEdit(ByVal row As Long, ByVal Col As Long)
RelineGridBill
End Sub
Sub RelineGridBill()
    Dim IntCounter As Integer
    Dim ActualTotal As Double
    Dim totalPayed As Double, VATValue As Double
    IntCounter = 0
    Dim i As Integer
    txtPayTotalProj = 0
    txtPayTotalProjVat = 0
    Dim TotalPayedFULL  As Double
    With Me.Grid
    '.TextMatrix(i, .ColIndex("TotalPayedFULL")) = 0
        For i = .FixedRows To .rows - 1
        If .cell(flexcpChecked, i, .ColIndex("Select")) = flexChecked Then
                 
                ActualTotal = ActualTotal + val(.TextMatrix(i, .ColIndex("ActualTotal")))
                    TotalPayedFULL = getBillPayedToproject(val(Me.DBCboClientName.BoundText))
                    totalPayed = val(.TextMatrix(i, .ColIndex("ActualTotal")))
                    .TextMatrix(i, .ColIndex("TotalPayedFULL")) = totalPayed + TotalPayedFULL
                      .TextMatrix(i, .ColIndex("result")) = val(.TextMatrix(i, .ColIndex("total"))) - totalPayed
                    .TextMatrix(i, .ColIndex("resultpercentage")) = Round(totalPayed / val(.TextMatrix(i, .ColIndex("total"))) * 100.2)
                    VATValue = CalculateVATProje(val(.TextMatrix(i, .ColIndex("total"))), val(.TextMatrix(i, .ColIndex("performanceDiscount"))), val(.TextMatrix(i, .ColIndex("FATYou"))))
                    VATValue = CalculatePaidVAT(val(.TextMatrix(i, .ColIndex("ActualTotal"))), VATValue, val(.TextMatrix(i, .ColIndex("total"))))
                    .TextMatrix(i, .ColIndex("ActualTotalVat")) = VATValue
                    txtPayTotalProj = val(txtPayTotalProj) + val(.TextMatrix(i, .ColIndex("ActualTotal")))
                    txtPayTotalProjVat = val(txtPayTotalProjVat) + val(.TextMatrix(i, .ColIndex("ActualTotalVat")))
           End If
           Next i
  
    End With
XPTxtVal = ActualTotal
End Sub

Private Sub Grid_StartEdit(ByVal row As Long, ByVal Col As Long, Cancel As Boolean)
   If Col = 1 Or Col = 13 Then
   
   Else
   
   Cancel = True
   End If

End Sub
Function CalculatePaidVAT(paidAmount As Double, TotalVat As Double, totalAmountWithVAT As Double) As Double
    Dim paidVATShare As Double
    paidVATShare = (paidAmount * TotalVat) / totalAmountWithVAT
    CalculatePaidVAT = paidVATShare
End Function
Private Sub Grid3_AfterEdit(ByVal row As Long, ByVal Col As Long)
Dim StrAccountCode  As String
Dim LngRow As Long
 With Grid3

        Select Case .ColKey(Col)
 
 Case "CommisionTypes"
 StrAccountCode = .ComboData
                LngRow = .FindRow(StrAccountCode, .FixedRows, .ColIndex("CommisionTypesid"), False, True)
                .TextMatrix(row, .ColIndex("CommisionTypesid")) = StrAccountCode
     

End Select

End With


ReLineGrid
End Sub
Function ReLineGrid()
If Me.TxtModFlg <> "R" Then
  Dim totalPayed As Double, VATValue As Double
    totalPayed = 0
  With Me.Grid3
 Dim i As Integer
totalPayed = 0
        For i = .FixedRows To .rows - 1

            If .cell(flexcpChecked, i, .ColIndex("Select")) = flexChecked Then
            
                    totalPayed = totalPayed + val(.TextMatrix(i, .ColIndex("RentValuePayed")))
                    totalPayed = totalPayed + val(.TextMatrix(i, .ColIndex("CommissionsPayed")))
                    totalPayed = totalPayed + val(.TextMatrix(i, .ColIndex("InsurancePayed")))
                    totalPayed = totalPayed + val(.TextMatrix(i, .ColIndex("WaterPayed")))
                    totalPayed = totalPayed + val(.TextMatrix(i, .ColIndex("ElectricPayed")))
                    totalPayed = totalPayed + val(.TextMatrix(i, .ColIndex("TelandNetPayed")))
                   .TextMatrix(i, .ColIndex("ActualTotal")) = totalPayed
             
                    
                      .TextMatrix(i, .ColIndex("result")) = val(.TextMatrix(i, .ColIndex("total"))) - totalPayed
                    .TextMatrix(i, .ColIndex("resultpercentage")) = Round(totalPayed / val(.TextMatrix(i, .ColIndex("total"))) * 100.2)
                 Else
                 
        End If

        Next i

    End With
      Me.XPTxtVal.text = totalPayed
End If

End Function

Private Sub Grid3_BeforeEdit(ByVal row As Long, ByVal Col As Long, Cancel As Boolean)
    With Grid3
 
    'If .ColKey(Col) <> "Due_DateH" And .ColKey(Col) <> "Status" Then
   If Me.TxtModFlg = "R" Then Exit Sub
If .ColKey(Col) <> "Select" Then
   If .cell(flexcpChecked, row, .ColIndex("Select")) = flexUnchecked Then Cancel = True: Exit Sub
End If

 

         If .ColKey(Col) <> "Select" And .ColKey(Col) <> "RentValuePayed" And .ColKey(Col) <> "CommissionsPayed" _
          And .ColKey(Col) <> "InsurancePayed" And .ColKey(Col) <> "WaterPayed" And .ColKey(Col) <> "ElectricPayed" _
          And .ColKey(Col) <> "TelandNetPayed" And .ColKey(Col) <> "CommisionTypes" Then
   
        
        Cancel = True
        
        End If
 
        
    End With
End Sub

Private Sub Grid3_StartEdit(ByVal row As Long, ByVal Col As Long, Cancel As Boolean)
    With Grid3
Dim StrSQL  As String
Dim StrComboList As String
Dim rs As New ADODB.Recordset

        Select Case .ColKey(Col)
 Case "CommisionTypes"
 
                StrSQL = "select * from TblCommisionTypes"
                rs.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

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
 
Load FrmIqarContractSearch
FrmIqarContractSearch.show
FrmIqarContractSearch.m_RetrunType = 1
 
End Sub

Function Showcashing(StartDate As Date, EndDate As Date, Optional NoteCashingType = -1, Optional brnchid As Integer)
    Dim StrSQL As String
    Dim Msg As String
    Dim BolBegine As Boolean
    Dim StrDesReport As String

    On Error GoTo ErrTrap
   
  StrDesReport = " ﬁ«—Ì— «·„ﬁ»Ê÷«  "
 
 
   
  StrSQL = "SELECT     TOP 100 PERCENT dbo.Notes.NoteID, dbo.Notes.NoteDate, dbo.Notes.Note_Value, dbo.Notes.CusID, dbo.TblCustemers.CusName, dbo.Notes.UserID, "
StrSQL = StrSQL + "                      dbo.TblUsers.UserName, dbo.Notes.CashingType, dbo.Notes.Remark, dbo.Transactions.Transaction_ID, dbo.Transactions.Transaction_Serial,"
StrSQL = StrSQL + "                      dbo.TransactionTypes.TransactionTypeName, dbo.TblBoxesData.BoxID, dbo.TblBoxesData.BoxName, dbo.Transactions.Transaction_Type, dbo.Notes.RevenuesID,"
StrSQL = StrSQL + "                      dbo.TblRevenuesTypes.RevenuesName, dbo.Notes.NoteSerial, dbo.Notes.NoteSerial1, dbo.Notes.AccountsCode, dbo.ACCOUNTS.Account_Code,"
StrSQL = StrSQL + "                      dbo.ACCOUNTS.Account_Name, dbo.TblCustemers.CusNamee, dbo.TblCustemers.Fullcode, dbo.TblEmployee.Emp_Name,"
StrSQL = StrSQL + "                      dbo.TblEmployee.Fullcode AS EmployeeFullcode, dbo.TblEmployee.Emp_Namee, dbo.Notes.EmpId"
StrSQL = StrSQL + " FROM         dbo.TblRevenuesTypes RIGHT OUTER JOIN"
StrSQL = StrSQL + "                      dbo.ACCOUNTS RIGHT OUTER JOIN"
StrSQL = StrSQL + " dbo.TblEmployee RIGHT OUTER JOIN"
StrSQL = StrSQL + "                      dbo.TblUsers INNER JOIN"
StrSQL = StrSQL + "                      dbo.Notes ON dbo.TblUsers.UserID = dbo.Notes.UserID ON dbo.TblEmployee.Emp_ID = dbo.Notes.EmpId ON"
StrSQL = StrSQL + "                      dbo.ACCOUNTS.Account_Code = dbo.Notes.AccountsCode ON dbo.TblRevenuesTypes.RevenuesID = dbo.Notes.RevenuesID LEFT OUTER JOIN"
StrSQL = StrSQL + "                      dbo.TblBoxesData ON dbo.Notes.BoxID = dbo.TblBoxesData.BoxID LEFT OUTER JOIN"
StrSQL = StrSQL + "                      dbo.TblCustemers ON dbo.Notes.CusID = dbo.TblCustemers.CusID LEFT OUTER JOIN"
 StrSQL = StrSQL + "                     dbo.Transactions LEFT OUTER JOIN"
StrSQL = StrSQL + "                      dbo.TransactionTypes ON dbo.Transactions.Transaction_Type = dbo.TransactionTypes.Transaction_Type ON"
 StrSQL = StrSQL + "                     dbo.Notes.Transaction_ID = dbo.Transactions.Transaction_ID"
 
                      
    StrSQL = StrSQL + " WHERE     (dbo.Notes.NoteType = 4)"
 
    
    StrSQL = StrSQL + "    and Notes.NOTEID <> 0"
    BolBegine = True

 If NoteCashingType <> -1 Then
 StrSQL = StrSQL + " and     (dbo.Notes.NoteCashingType = " & NoteCashingType & ")"
 End If
 

     
        StrDesReport = StrDesReport & CHR(13) & " «—ÌŒ «·Õ—ﬂ«  Ì»œ« „‰:" & DisplayDate(StartDate)
StrSQL = StrSQL + " AND NoteDate>='" & SQLDate(StartDate) & "'"

 StrDesReport = StrDesReport & CHR(13) & " «—ÌŒ «·Õ—ﬂ«  Ì‰ ÂÏ Õ Ï:" & DisplayDate(EndDate)
StrSQL = StrSQL + " and NoteDate<='" & SQLDate(EndDate) & "'"

  
    StrSQL = StrSQL + " Order by Notes.NoteSerial1"
    Dim Reports As New ClsRepoerts
    Reports.CashingReports StrSQL, WindowTarget, StrDesReport, 1
    Exit Function
ErrTrap:

End Function

Private Sub Label29_Click()
Frame12.Visible = False
End Sub

Function fillapprovData()
Dim Num As Integer
 Dim RsDetails As New ADODB.Recordset
 Dim StrSQL As String
 
 
 StrSQL = "SELECT     TOP 100 PERCENT dbo.ApprovalData.Currcursor, dbo.ApprovalData.ScreenName, dbo.ApprovalData.levelo, dbo.ApprovalData.EmpID, dbo.ApprovalData.levelorder, "
StrSQL = StrSQL + " dbo.ApprovalData.currorder, dbo.ApprovalData.Transaction_ID, dbo.ApprovalData.NoteID, dbo.ApprovalData.ApprovDate, dbo.ApprovalData.Remarks,"
StrSQL = StrSQL + " dbo.TbLLevels.name , dbo.TbLLevels.namee, dbo.TblUsers.UserID, dbo.TblUsers.UserName"
StrSQL = StrSQL + " FROM         dbo.ApprovalData left JOIN"
StrSQL = StrSQL + " dbo.TbLLevels ON dbo.ApprovalData.levelo = dbo.TbLLevels.LevelID INNER JOIN"
StrSQL = StrSQL + " dbo.TblUsers ON dbo.ApprovalData.EmpID = dbo.TblUsers.UserID"
StrSQL = StrSQL + " WHERE     (dbo.ApprovalData.Transaction_ID = " & val(Me.XPTxtID.text) & ") AND (dbo.ApprovalData.ScreenName = N'" & Me.Name & "')"
StrSQL = StrSQL + " ORDER BY dbo.ApprovalData.levelorder"

    RsDetails.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText
If RsDetails.RecordCount > 0 Then
 If SystemOptions.UserInterface = ArabicInterface Then
    Accredit.Caption = " „ «·«—”«· ··«⁄ „«œ"
Else
Accredit.Caption = "Sent To approval "
End If
Accredit.Enabled = False
Else
Accredit.Enabled = True
 If SystemOptions.UserInterface = ArabicInterface Then
    Accredit.Caption = " «·«—”«· ··«⁄ „«œ"
Else
Accredit.Caption = "Sent To approval "
End If
End If
 If Not (RsDetails.EOF Or RsDetails.BOF) Then
        Grid2.rows = RsDetails.RecordCount + 1
 

        For Num = 1 To RsDetails.RecordCount
        
       Grid2.TextMatrix(Num, Grid2.ColIndex("Currcursor")) = IIf(IsNull(RsDetails("Currcursor")), "", RsDetails("Currcursor"))
    If Grid2.TextMatrix(Num, Grid2.ColIndex("Currcursor")) = "1" Then
   Grid2.cell(flexcpBackColor, Num, 1, Num, 7) = &HFFFFC0
   Else
    Grid2.cell(flexcpBackColor, Num, 1, Num, 7) = vbWhite
    End If
    
        Grid2.TextMatrix(Num, Grid2.ColIndex("Approved")) = IIf(IsNull(RsDetails("ApprovDate")), "", flexChecked)
           If SystemOptions.UserInterface = ArabicInterface Then
            Grid2.TextMatrix(Num, Grid2.ColIndex("levelName")) = IIf(IsNull(RsDetails("Name")), "", Trim(RsDetails("Name").value))
          Else
             Grid2.TextMatrix(Num, Grid2.ColIndex("levelName")) = IIf(IsNull(RsDetails("Namee")), "", Trim(RsDetails("Namee").value))
          End If
            If SystemOptions.UserInterface = ArabicInterface Then
            Grid2.TextMatrix(Num, Grid2.ColIndex("EmpName")) = IIf(IsNull(RsDetails("UserName")), "", (RsDetails("UserName").value))
            Else
            Grid2.TextMatrix(Num, Grid2.ColIndex("EmpName")) = IIf(IsNull(RsDetails("UserName")), "", (RsDetails("UserName").value))
            End If
            Grid2.TextMatrix(Num, Grid2.ColIndex("ApprovDate")) = IIf(IsNull(RsDetails("ApprovDate")), "", (RsDetails("ApprovDate").value))
          Grid2.TextMatrix(Num, Grid2.ColIndex("REMARKS")) = IIf(IsNull(RsDetails("REMARKS")), "", (RsDetails("REMARKS").value))
 
 
RsDetails.MoveNext
If Num = RsDetails.RecordCount Then

        If Grid2.TextMatrix(Num, Grid2.ColIndex("Approved")) <> "" Then
                                If SystemOptions.UserInterface = ArabicInterface Then
                                      Label24.Caption = " „ «·«⁄ „«œ ··„” ‰œ »«·ﬂ«„·"
                                 Else
                                       Label24.Caption = "Approved"
                                 End If
                            Label24.backcolor = &H80FF80
        Else
                             If SystemOptions.UserInterface = ArabicInterface Then
                                     Label24.Caption = "„ÿ·Ê» «⁄ „«œ… Õ«·Ì«"
                            Else
                                     Label24.Caption = "Currently required Approve"
                            End If
                 Label24.backcolor = &HFFFFC0
        End If

End If

        Next Num
Else
 Grid2.rows = 1
    End If
RsDetails.Close

End Function
Private Sub lbl_MouseMove(index As Integer, _
                          Button As Integer, _
                          Shift As Integer, _
                          X As Single, _
                          Y As Single)

    If index = 18 Then
        If SystemOptions.UserInterface = ArabicInterface Then
            Me.lbl(18).ToolTipText = "ﬁÌ„… „»·€ «·„ﬁ»Ê÷« :" & lbl(18).Caption
        ElseIf SystemOptions.UserInterface = EnglishInterface Then
            Me.lbl(18).ToolTipText = "Notes Recivable Value:" & lbl(18).Caption
        End If
    End If

End Sub

Private Sub LblLink_Click()
 
    Dim FirstPeriod As Date
    getFirstPeriodDateInthisYear FirstPeriod
    ShowReport DcboCreditSide.BoundText, DcboCreditSide.text, FirstPeriod, Date

End Sub
Function GetPercentage(Optional TypeTr As Integer = 0, Optional Perid As Integer = 0) As Double
Dim sql As String
Dim Rs4 As ADODB.Recordset
Set Rs4 = New ADODB.Recordset
sql = "select * from TblAcceleratePaymentDet where TransType=0 and FromValue <=" & Perid & " and ToValue >=" & Perid & ""
Rs4.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
If Rs4.RecordCount > 0 Then
If TypeTr = 0 Then
GetPercentage = IIf(IsNull(Rs4("Percentage").value), 0, Rs4("Percentage").value)
Else
GetPercentage = IIf(IsNull(Rs4("PercentageAll").value), 0, Rs4("PercentageAll").value)
End If
Else
GetPercentage = 0
End If
End Function

Function CheckCustomer(Optional CusID As Double = 0) As Boolean
Dim sql As String
Dim Rs4 As ADODB.Recordset
Set Rs4 = New ADODB.Recordset
sql = "select * from TblAcceleratePaymentDet where CusID=" & CusID & " "
Rs4.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
If Rs4.RecordCount > 0 Then
CheckCustomer = True
Else
CheckCustomer = False
End If
End Function
Function CheckProjectBill(Optional Transaction_ID As Double = 0) As Boolean
Dim sql As String
Dim Rs4 As ADODB.Recordset
Set Rs4 = New ADODB.Recordset
sql = "select * from TblProjePayPrePayed where Transaction_ID=" & Transaction_ID & " "
Rs4.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
If Rs4.RecordCount > 0 Then
CheckProjectBill = True
Else
CheckProjectBill = False
End If
End Function

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

Private Sub LblLinkInfo_Click(index As Integer)
    Dim StartWeekDate As Date
    Dim EndWeekDate As Date
    Dim StartMonthDate As Date
    Dim EndMonthDate As Date
    
    
       StartWeekDate = GetWeekStartEND(Date, 0)
    EndWeekDate = DateAdd("d", 7, StartWeekDate)
    StartMonthDate = dhFirstDayInMonth(Date)
    EndMonthDate = dhLastDayInMonth(Date)
    
     
Select Case index

Case 0
Showcashing Date, Date, 0 'tody cash
Case 1
Showcashing Date, Date, 1 'tody Cheque
Case 2 'week cash
Showcashing StartWeekDate, EndWeekDate, 0

Case 3 'week Cheque
Showcashing StartWeekDate, EndWeekDate, 1
Case 4 'month cash
Showcashing StartMonthDate, EndMonthDate, 0

Case 5 'month cheque
Showcashing StartMonthDate, EndMonthDate, 1

Case 6 ' tody all
Showcashing Date, Date

Case 7 'week all
 

Showcashing StartWeekDate, EndWeekDate
Case 8 'month all
Showcashing StartMonthDate, EndMonthDate
End Select
End Sub

Private Sub Option1_Click()
ALLButton6.Enabled = False
   If Me.TxtModFlg.text = "R" Or Me.TxtModFlg.text = "" Then
   If Option1.value = True Then
        ALLButton6.Enabled = True
    Else
        ALLButton6.Enabled = False
    End If
    End If
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
CalCulteVAT
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
CalCulteVAT
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
CalCulteVAT
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

Private Sub Option7_Click()
CalCulteVAT
End Sub

Private Sub TxtAccount_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF3 Then
        Account_search.show
        Account_search.case_id = 260816
    End If

End Sub

Private Sub TxtBillMaintNo_Change()
If Me.TxtModFlg.text <> "R" And val(DCboCashType.ListIndex) = 9 Or val(DCboCashType.ListIndex) = 10 Then
TxtBillMaintID.text = GetBillMaintID(TxtBillMaintNo.text)
GetInformationBillMaint val(TxtBillMaintID.text)
End If
End Sub

Private Sub TxtBillTransNo_Change()
If Me.TxtModFlg.text <> "R" And val(DCboCashType.ListIndex) = 8 Then
TxtBillTransID.text = GetBillTransID(TxtBillTransNo.text)
GetInformationBillTrans val(TxtBillTransID.text)
End If
End Sub
Function GetBillMaintID(Optional Fullcode As String) As Double
Dim sql As String
Dim rs2 As ADODB.Recordset
Set rs2 = New ADODB.Recordset

 
If DCboCashType.ListIndex = 10 Then
    sql = "SELECT ID FROM TblCardAuthorizationReform WHERE WorkOrder=" & val(Fullcode)
Else
    sql = "SELECT ID FROM TblCarBillMentains WHERE NoteSerial1='" & Fullcode & "' "
    
End If


rs2.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
If rs2.RecordCount > 0 Then
GetBillMaintID = IIf(IsNull(rs2("ID").value), 0, rs2("ID").value)
Else
GetBillMaintID = 0
End If
End Function

Function GetBillTransID(Optional Fullcode As String) As Double
Dim sql As String
Dim rs2 As ADODB.Recordset
Set rs2 = New ADODB.Recordset
sql = "SELECT ID FROM TblTravDueK WHERE NoteSerial1='" & Fullcode & "' "
rs2.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
If rs2.RecordCount > 0 Then
GetBillTransID = IIf(IsNull(rs2("ID").value), 0, rs2("ID").value)
Else
GetBillTransID = 0
End If
End Function
Sub GetInformationBillMaint(Optional ID As Double)
Dim sql As String
Dim rs2 As ADODB.Recordset
Set rs2 = New ADODB.Recordset
Dim mType As Integer
mType = DCboCashType.ListIndex

If DCboCashType.ListIndex = 10 Then
    sql = " select CusID  , Clientname, AmountAccept as TotalValue,* from TblCardAuthorizationReform where ID=" & ID & ""
Else
    sql = " select * from TblCarBillMentains where ID=" & ID & ""
End If



rs2.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
If rs2.RecordCount > 0 Then
    DBCboClientName.BoundText = IIf(IsNull(rs2("CusID").value), 0, rs2("CusID").value)
    XPTxtVal.text = IIf(IsNull(rs2("TotalValue").value), 0, rs2("TotalValue").value)
CalCulteVAT 1
Else
XPTxtVal.text = 0
DBCboClientName.BoundText = 0
End If
End Sub
Sub GetInformationBillTrans(Optional ID As Double)
Dim sql As String
Dim rs2 As ADODB.Recordset
Set rs2 = New ADODB.Recordset
sql = " select * from TblTravDueK where ID=" & ID & ""
rs2.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
If rs2.RecordCount > 0 Then
DBCboClientName.BoundText = IIf(IsNull(rs2("CusID").value), 0, rs2("CusID").value)
XPTxtVal.text = IIf(IsNull(rs2("Total").value), IIf(IsNull(rs2("VAT").value), 0, rs2("VAT").value) + IIf(IsNull(rs2("TotalValue").value), 0, rs2("TotalValue").value), rs2("Total").value)
Else
XPTxtVal.text = 0
DBCboClientName.BoundText = 0
End If
End Sub
'Private Sub TxtContNo_Change()
'  If Me.DCboCashType.ListIndex = 8 And (Me.TxtModFlg.Text = "N" Or Me.TxtModFlg.Text = "E") Then
'        FillGridWithDataContract val(TXTContNo.Text)
'
'    End If
'End Sub
Sub GetContainerData(Optional ID As Double)
Dim sql As String
Dim rs2 As ADODB.Recordset
Set rs2 = New ADODB.Recordset

sql = " select * from ContainerContracts where ID=" & ID & " And "
sql = "SELECT CustID,    Net = ISNULL(Net,0) - IsNull((SELECT SUM(Note_Value) FROM Notes"
sql = sql & " Where IsNull(CashingType, 0) = 12"
sql = sql & " And NoteId <>  " & val(XPTxtID.text)
sql = sql & " and Notes.ContainerNo = " & ID & " ),0)  FROM ContainerContracts"
sql = sql & " Where ID = " & ID


rs2.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
If rs2.RecordCount > 0 Then
    DBCboClientName.BoundText = IIf(IsNull(rs2("CustID").value), 0, rs2("CustID").value)
    XPTxtVal.text = rs2!net & ""
Else
XPTxtVal.text = 0
DBCboClientName.BoundText = 0
End If
End Sub
'Private Sub TxtContNo_Change()
'  If Me.DCboCashType.ListIndex = 8 And (Me.TxtModFlg.Text = "N" Or Me.TxtModFlg.Text = "E") Then
'        FillGridWithDataContract val(TXTContNo.Text)
'
'    End If
'End Sub

Private Sub txtContainerNo_Change()
If Me.TxtModFlg.text <> "R" And val(DCboCashType.ListIndex) = 12 Then
'TxtBillTransID.Text = GetBillTransID(TxtBillTransNo.Text)
    GetContainerData val(txtContainerNo.text)
End If
End Sub

Private Sub TxtCustCode_KeyPress(KeyAscii As Integer)
    
   Dim CUSTID As Integer
 Dim EmpID As Integer
 Dim ID As Double
    If DCboCashType.ListIndex = 0 Or DCboCashType.ListIndex = 11 Then

    If KeyAscii = vbKeyReturn Then
        GetCustomersDetail CUSTID, , TxtCustCode.text, DCboCashType.ListIndex + 1
        DBCboClientName.BoundText = CUSTID
    End If
ElseIf DCboCashType.ListIndex = 5 Then
    If KeyAscii = vbKeyReturn Then
    If Text1.text <> "" Then
GetCodeIDProject ID, TxtCustCode.text
DBCboClientName.BoundText = ID
    End If
    End If
    
    


ElseIf DCboCashType.ListIndex = 6 Then
    If KeyAscii = vbKeyReturn Then
        GetEmployeeIDFromCode TxtCustCode.text, EmpID
        Me.DCEmployee.BoundText = EmpID
    End If
    
ElseIf DCboCashType.ListIndex = 7 Then
    If KeyAscii = vbKeyReturn Then
    
        DCAccounts.BoundText = getAccountSerial_Code("Account_Code", "Account_Serial", TxtCustCode.text)
    End If
    
 End If
 
End Sub

Private Sub TxtCustCode_KeyUp(KeyCode As Integer, Shift As Integer)
    If Me.DCboCashType.ListIndex = 7 Then
        If KeyCode = vbKeyF3 Then
            Account_search.show
            Account_search.case_id = 260817
        End If
    End If
End Sub

Private Sub TxtEmployeeID_KeyPress(KeyAscii As Integer)
 Dim EmpID As Integer

    If KeyAscii = vbKeyReturn Then
        GetEmployeeIDFromCode TxtEmployeeID.text, EmpID
        DcEmp.BoundText = EmpID
    End If

End Sub

Private Sub TxtModFlg_Change()
    On Error GoTo ErrTrap
 VSFlexGrid1.Enabled = True
    Select Case Me.TxtModFlg.text

        Case "R"
       '  VSFlexGrid1.Enabled = False
        txtperson.Enabled = False
    dcBranch.Enabled = False
    Frame5.Enabled = False
    TxtManulaNO.Enabled = False
    Frame2.Enabled = False
    DBCboClientName.Enabled = False
    DCChequeBox.Enabled = False
    
    DcEmp.Enabled = False
    DcCostCenter.Enabled = False
    TxtBookNo.Enabled = False
    DCCar.Enabled = False
    DCDriver.Enabled = False
   ' Frame1.Enabled = False
    dcBranch.Enabled = False
    Txt_DateHigri.Enabled = False
    
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
            XPMTxtRemarks.locked = True
            DBCboClientName.locked = True
            DCboCashType.locked = True
            Me.CboPayMentType.locked = True
            Me.DcboBox.locked = True
            Me.DcboBankName.locked = True
            Me.TxtChequeNumber.locked = True
            Me.DtpChequeDueDate.Enabled = False

            If rs.RecordCount < 1 Then
                Me.XPBtnMove(0).Enabled = False
                Me.XPBtnMove(1).Enabled = False
                Me.XPBtnMove(2).Enabled = False
                Me.XPBtnMove(3).Enabled = False
                Me.Cmd(1).Enabled = False
                Me.Cmd(4).Enabled = False
            End If

            Fra(0).Enabled = False
            ChkTrans.Enabled = False

        Case "N"
        txtperson.Enabled = True
            dcBranch.Enabled = True
    Frame5.Enabled = True
    TxtManulaNO.Enabled = True
    Frame2.Enabled = True
    DBCboClientName.Enabled = True
    DCChequeBox.Enabled = True
    
    DcEmp.Enabled = True
    DcCostCenter.Enabled = True
    TxtBookNo.Enabled = True
    DCCar.Enabled = True
    DCDriver.Enabled = True
    Frame1.Enabled = True
    dcBranch.Enabled = True
    Txt_DateHigri.Enabled = True
    
    
            dcBranch.Enabled = True
    Frame5.Enabled = True
    
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
            If SystemOptions.DateCanNotEdit = True Then
            XPDtbTrans.Enabled = False
            Else
            XPDtbTrans.Enabled = True
            End If
            XPTxtVal.Enabled = True
            XPTxtVal.locked = False
            XPMTxtRemarks.locked = False
            DBCboClientName.locked = False
            XPDtbTrans.value = Date
            DCboCashType.locked = False
            DCboCashType.ListIndex = 0
        
            Me.CboPayMentType.locked = False
            Me.DcboBox.locked = False
            Me.DcboBankName.locked = False
            Me.TxtChequeNumber.locked = False
            Me.DtpChequeDueDate.Enabled = True
        
            Fra(0).Enabled = True
            ChkTrans.Enabled = True

        Case "E"
        txtperson.Enabled = True
       If SystemOptions.BranchCanNotEdit = True Then
        dcBranch.Enabled = False
        Text2.Enabled = False
       Else
       Text2.Enabled = True
       dcBranch.Enabled = True
      End If
    Frame5.Enabled = True
    TxtManulaNO.Enabled = True
    Frame2.Enabled = True
    DBCboClientName.Enabled = True
    DCChequeBox.Enabled = True
    
    DcEmp.Enabled = True
    DcCostCenter.Enabled = True
    TxtBookNo.Enabled = True
    DCCar.Enabled = True
    DCDriver.Enabled = True
    Frame1.Enabled = True
    If SystemOptions.BranchCanNotEdit = True Then
        dcBranch.Enabled = False
        Text2.Enabled = False
       Else
       Text2.Enabled = True
       dcBranch.Enabled = True
      End If
    Txt_DateHigri.Enabled = True
    
            '        Me.Caption = "«·„ﬁ»Ê÷« (  ⁄œÌ· )"
'Grid3.Visible = True
Grid3.Enabled = True
            Grid.Enabled = True
            GRID1.Enabled = True
            
            Grid3.Enabled = True
             
    If SystemOptions.BranchCanNotEdit = True Then
        dcBranch.Enabled = False
        Text2.Enabled = False
       Else
       Text2.Enabled = True
       dcBranch.Enabled = True
      End If
    Frame5.Enabled = True
    
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
          If SystemOptions.DateCanNotEdit = True Then
            XPDtbTrans.Enabled = False
            Else
            XPDtbTrans.Enabled = True
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
            Me.CboPayMentType.locked = False
            Me.DcboBox.locked = False
            Me.DcboBankName.locked = False
            Me.TxtChequeNumber.locked = False
            Me.DtpChequeDueDate.Enabled = True
    End Select

    Exit Sub
ErrTrap:
End Sub

Private Sub CmdNos_Click(index As Integer)
  If index <= 9 Then
LBLPayVal.Caption = LBLPayVal.Caption & index

ElseIf index = 10 Then
LBLPayVal.Caption = LBLPayVal.Caption & "00"

ElseIf index = 11 Then
LBLPayVal.Caption = LBLPayVal.Caption & "."

ElseIf index = 12 Then 'ar
If Len(LBLPayVal.Caption) > 1 Then
LBLPayVal.Caption = mId(LBLPayVal.Caption, 1, Len(LBLPayVal.Caption) - 1)
Else
LBLPayVal.Caption = ""
End If
ElseIf index = 13 Then 'ar
 LBLPayVal.Caption = ""

TxtPayedValue2.text = ""
cleargrid

ElseIf index = 14 Then
TxtPayedValue2.text = val(LBLPayVal)

 
        With Grid22
          .TextMatrix(.row, .ColIndex("Value")) = TxtPayedValue2.text
          End With
    ReLineGrid2
     
 TxtRemainValue2.text = val(Me.TxtPayedValue2.text) - val(Me.TxtNetValue2.text)
 

End If

 ReLineGrid2
 
End Sub
Public Sub FillGridWithData222()

    On Error GoTo ErrTrap

    Dim i As Integer
    Dim rs As ADODB.Recordset
    Dim My_SQL As String

    Set rs = New ADODB.Recordset
    My_SQL = "SELECT     dbo.TblPaymentType.PaymentID, dbo.TblPaymentType.PaymentName, dbo.TblPaymentType.BankId, dbo.TblPaymentType.Accountsus, "
My_SQL = My_SQL & "  dbo.TblPaymentType.Accountcom, dbo.TblPaymentType.commision, dbo.TblPaymentType.PaymentNamee, dbo.BanksData.Account_Code AS bankAccount_Code ,dbo.TblPaymentType.MaxValue"
My_SQL = My_SQL & " FROM         dbo.TblPaymentType LEFT OUTER JOIN"
My_SQL = My_SQL & " dbo.BanksData ON dbo.TblPaymentType.BankId = dbo.BanksData.BankID"
My_SQL = My_SQL & " where (dbo.TblPaymentType.TypTran=2 or dbo.TblPaymentType.TypTran is null)  "
If SystemOptions.LinkUsersWithPayment = True Then
My_SQL = My_SQL & " and dbo.TblPaymentType.PaymentID in (SELECT     PaynetID"
My_SQL = My_SQL & " From dbo.TblPaymentUser"
My_SQL = My_SQL & " Where (UserID = " & user_id & "))"
End If
My_SQL = My_SQL & " order by PaymentID"
    rs.Open My_SQL, Cn, adOpenKeyset, adLockReadOnly, adCmdText

    With Me.Grid22
        .rows = 2
        .Clear flexClearScrollable

        If rs.RecordCount > 0 Then
            .rows = rs.RecordCount + 2
            rs.MoveFirst
      If SystemOptions.UserInterface = ArabicInterface Then
                .TextMatrix(1, .ColIndex("PaymentName")) = " ‰ﬁœÌ"
               Else
               .TextMatrix(1, .ColIndex("PaymentName")) = " Cash"
               End If
               
                .TextMatrix(1, .ColIndex("PaymentID")) = 0
           
           
            For i = 2 To .rows - 1
                .TextMatrix(i, .ColIndex("Ser")) = i
             If SystemOptions.UserInterface = ArabicInterface Then
                .TextMatrix(i, .ColIndex("PaymentName")) = IIf(IsNull(rs.Fields("PaymentName").value), "", rs.Fields("PaymentName").value)
               Else
               .TextMatrix(i, .ColIndex("PaymentName")) = IIf(IsNull(rs.Fields("PaymentNamee").value), "", rs.Fields("PaymentNamee").value)
               End If
               
                .TextMatrix(i, .ColIndex("PaymentID")) = IIf(IsNull(rs.Fields("PaymentID").value), "", rs.Fields("PaymentID").value)
           
                .TextMatrix(i, .ColIndex("BankId")) = IIf(IsNull(rs.Fields("BankId").value), "", rs.Fields("BankId").value)
            
            .TextMatrix(i, .ColIndex("Accountsus")) = IIf(IsNull(rs.Fields("Accountsus").value), "", rs.Fields("Accountsus").value)
            .TextMatrix(i, .ColIndex("Accountcom")) = IIf(IsNull(rs.Fields("Accountcom").value), "", rs.Fields("Accountcom").value)
            .TextMatrix(i, .ColIndex("MaxValue")) = IIf(IsNull(rs.Fields("MaxValue").value), 0, rs.Fields("MaxValue").value)
            .TextMatrix(i, .ColIndex("commision")) = IIf(IsNull(rs.Fields("commision").value), "", rs.Fields("commision").value)
           .TextMatrix(i, .ColIndex("bankAccount_Code")) = IIf(IsNull(rs.Fields("bankAccount_Code").value), "", rs.Fields("bankAccount_Code").value)
            
                rs.MoveNext
            Next

            rs.Close
        End If

  '      .RowHeight(-1) = 300
    End With

ErrTrap:
End Sub
 Function FillGridWithDataPayment() As Boolean
    On Error GoTo ErrTrap
    Dim i As Integer
    Dim rs As ADODB.Recordset
    Dim My_SQL As String

    Set rs = New ADODB.Recordset
    My_SQL = " SELECT     TOP 100 PERCENT dbo.TblPaymentType.PaymentName, dbo.TblPaymentType.BankId, dbo.TblPaymentType.Accountsus, dbo.TblPaymentType.Accountcom, "
    My_SQL = My_SQL & "                   dbo.TblPaymentType.commision, dbo.TblPaymentType.PaymentNamee, dbo.BanksData.Account_Code AS bankAccount_Code, dbo.TblMultuPayment.[Value],"
    My_SQL = My_SQL & "                   dbo.TblMultuPayment.CardNo, dbo.TblMultuPayment.PaymentID , dbo.TblMultuPayment.NoteID,dbo.TblMultuPayment.MaxValue"
    My_SQL = My_SQL & "      FROM         dbo.TblPaymentType RIGHT OUTER JOIN"
    My_SQL = My_SQL & "                   dbo.TblMultuPayment ON dbo.TblPaymentType.PaymentID = dbo.TblMultuPayment.PaymentID LEFT OUTER JOIN"
    My_SQL = My_SQL & "                   dbo.BanksData ON dbo.TblPaymentType.BankId = dbo.BanksData.BankID"
    My_SQL = My_SQL & "     Where (dbo.TblMultuPayment.NoteID = " & val(XPTxtID.text) & ")"
    My_SQL = My_SQL & "   ORDER BY dbo.TblPaymentType.PaymentID"
    rs.Open My_SQL, Cn, adOpenKeyset, adLockReadOnly, adCmdText

    With Me.Grid22
        .rows = 2
        .Clear flexClearScrollable
        If rs.RecordCount > 0 Then
            .rows = rs.RecordCount + 2
            rs.MoveFirst
FillGridWithDataPayment = True
            For i = 1 To .rows - 1
                .TextMatrix(i, .ColIndex("Ser")) = i
             If SystemOptions.UserInterface = ArabicInterface Then
                .TextMatrix(i, .ColIndex("PaymentName")) = IIf(IsNull(rs.Fields("PaymentName").value), "‰ﬁœÌ", rs.Fields("PaymentName").value)
               Else
               .TextMatrix(i, .ColIndex("PaymentName")) = IIf(IsNull(rs.Fields("PaymentNamee").value), "Cash", rs.Fields("PaymentNamee").value)
               End If
               .TextMatrix(i, .ColIndex("Value")) = IIf(IsNull(rs.Fields("Value").value), "", rs.Fields("Value").value)
               .TextMatrix(i, .ColIndex("CardNo")) = IIf(IsNull(rs.Fields("CardNo").value), "", rs.Fields("CardNo").value)
               .TextMatrix(i, .ColIndex("MaxValue")) = IIf(IsNull(rs.Fields("MaxValue").value), 0, rs.Fields("MaxValue").value)
                .TextMatrix(i, .ColIndex("PaymentID")) = IIf(IsNull(rs.Fields("PaymentID").value), "", rs.Fields("PaymentID").value)
           
                .TextMatrix(i, .ColIndex("BankId")) = IIf(IsNull(rs.Fields("BankId").value), "", rs.Fields("BankId").value)
            
            .TextMatrix(i, .ColIndex("Accountsus")) = IIf(IsNull(rs.Fields("Accountsus").value), "", rs.Fields("Accountsus").value)
            .TextMatrix(i, .ColIndex("Accountcom")) = IIf(IsNull(rs.Fields("Accountcom").value), "", rs.Fields("Accountcom").value)
            .TextMatrix(i, .ColIndex("commision")) = IIf(IsNull(rs.Fields("commision").value), "", rs.Fields("commision").value)
           .TextMatrix(i, .ColIndex("bankAccount_Code")) = IIf(IsNull(rs.Fields("bankAccount_Code").value), "", rs.Fields("bankAccount_Code").value)
            
                rs.MoveNext
            Next

            rs.Close
            Else
            FillGridWithDataPayment = False
        End If

  '      .RowHeight(-1) = 300
    End With

ErrTrap:
End Function
Private Sub Grid22_AfterEdit(ByVal row As Long, ByVal Col As Long)
ReLineGrid2
End Sub
Private Sub CmdValue_Click(index As Integer)
LBLPayVal.Caption = 0
'TxtPayedValue.text = CmdValue(Index).Caption
LBLPayVal.Caption = CmdValue(index).Caption
        With Grid22
          .TextMatrix(.row, .ColIndex("Value")) = LBLPayVal.Caption
          End With
     ReLineGrid2
End Sub
Private Sub lblexit_Click(index As Integer)
If Me.TxtModFlg.text <> "R" Then
If index = 90 Then
If val(CboPayMentType.ListIndex) = 5 Then
If val(TxtRemainValue2.text) <> 0 Then
If SystemOptions.UserInterface = ArabicInterface Then
MsgBox "«·ﬁÌ„… «·„œŒ·… €Ì— ’ÕÌÕ…"
Else
MsgBox "The  value is incorrect"
End If
Exit Sub
End If
FramePay.Visible = False
End If
End If
Else
FramePay.Visible = False
End If
End Sub
Private Sub ReLineGrid2()
If Me.TxtModFlg = "R" Then Exit Sub
    On Error Resume Next
    Dim i As Integer
    Dim IntCounter As Integer
    Dim totalPayed As Double
    Dim visapayed As Double
 totalPayed = 0
 visapayed = 0
  With Grid22

        For i = .FixedRows To .rows - 1

            If .TextMatrix(i, .ColIndex("value")) <> "" Then
               ' IntCounter = IntCounter + 1
                totalPayed = totalPayed + .TextMatrix(i, .ColIndex("value"))
                If totalPayed > val(Me.TxtNetValue2.text) Then
                If SystemOptions.UserInterface = ArabicInterface Then
                MsgBox "·«Ì„ﬂ‰ «‰  ﬂÊ‰ «·ﬁÌ„… «ﬂ»— „‰ «·«Ã„«·Ì"
                Else
                 MsgBox "ERROR Incorrect Value" & CHR(13)
                End If
                .TextMatrix(i, .ColIndex("value")) = 0
                Exit Sub
                End If
            End If

        Next i

    End With
  TxtPayedValue2.text = totalPayed
    TxtRemainValue2.text = val(Me.TxtPayedValue2.text) - val(Me.TxtNetValue2.text)
End Sub
Private Sub Grid22_Click()
If TxtPayedValue2.text = "" Or val(TxtPayedValue2.text) = 0 Then
With Me.Grid22
.TextMatrix(.row, .ColIndex("Value")) = LBLPayVal.Caption
ReLineGrid2
End With
End If
End Sub
Private Sub CMDPAy_Click()
If val(CboPayMentType.ListIndex) = 5 Then
If val(TxtRemainValue2.text) <> 0 Then
If SystemOptions.UserInterface = ArabicInterface Then
MsgBox "«·ﬁÌ„… «·„œŒ·… €Ì— ’ÕÌÕ…"
Else
MsgBox "The  value is incorrect"
End If
Exit Sub
End If
FramePay.Visible = False
End If
    If SystemOptions.CanEditOnlyPayMethod And (Me.TxtModFlg = "E" Or Me.TxtModFlg = "R") Then
        Label20.Enabled = False
        lblexit(90).Enabled = False
        Ele(12).Enabled = False
        XPTab301.Enabled = False
    Else
         Label20.Enabled = True
         lblexit(90).Enabled = True
    
    End If


End Sub
Private Sub cleargrid()
    On Error Resume Next
    Dim i As Integer
 
  With Grid22

       ' For I = .FixedRows To .Rows - 1

         .TextMatrix(.row, .ColIndex("value")) = 0
          
       ' Next I

    End With
     TxtPayedValue2 = 0
    
End Sub

Sub SaveMultyPayment(Optional NoteID As Double)
Dim Rs3 As ADODB.Recordset
Dim i As Integer
Set Rs3 = New ADODB.Recordset
Dim sql As String
sql = "select * from TblMultuPayment where 1=-1 "
Rs3.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
With Grid22
For i = 1 To .rows - 1
If (.TextMatrix(i, .ColIndex("PaymentName"))) <> "" Then
Rs3.AddNew
Rs3("NoteID").value = NoteID
Rs3("PaymentID").value = val(.TextMatrix(i, .ColIndex("PaymentID")))
Rs3("Value").value = val(.TextMatrix(i, .ColIndex("Value")))
Rs3("CardNo").value = (.TextMatrix(i, .ColIndex("CardNo")))
Rs3("maxvalue").value = val((.TextMatrix(i, .ColIndex("MaxValue"))))

Rs3.update
End If
Next i
End With
End Sub

Private Sub TxtPaymentValue_Change()
If Me.TxtModFlg.text <> "R" Then
TxtPercentage.text = 0
If val(TxtCurrentBalance.text) > val(Me.TxtPaymentValue.text) Then
TxtPercentage.text = GetPercentage(0, day(XPDtbTrans.value))
ElseIf val(TxtCurrentBalance.text) = val(Me.TxtPaymentValue.text) Then
TxtPercentage.text = GetPercentage(1, day(XPDtbTrans.value))
Else
If SystemOptions.UserInterface = ArabicInterface Then
MsgBox "·«Ì„ﬂ‰ «‰  ﬂÊ‰ «·ﬁÌ„… «ﬂ»— „‰ «·—’Ìœ"
Else
MsgBox "Value can not be greater than balance"
End If
TxtPaymentValue.text = 0
TxtPaymentValue.SetFocus
XPTxtVal.text = 0
Exit Sub
End If
TxtPercentageValue.text = (val(TxtPaymentValue.text) * val(TxtPercentage.text)) / 100
XPTxtVal.text = val(TxtPaymentValue.text) - (val(TxtPaymentValue.text) * val(TxtPercentage.text)) / 100
End If
End Sub

Private Sub TxtPaymentValue_KeyPress(KeyAscii As Integer)
KeyAscii = KeyAscii_Num(KeyAscii, Me.TxtPaymentValue.text, 0)
End Sub

Private Sub txtTotal_GotFocus()
    mClick = True
End Sub

Private Sub txttotal_LostFocus()
ClaCul
End Sub

Private Sub txtTotal_Validate(Cancel As Boolean)
CalCulteVAT 0
End Sub

Private Sub TxtTransID_Change()

    If Me.TxtModFlg.text = "N" Or Me.TxtModFlg.text = "E" Then
        If Me.TxtTransID.text <> "" Then
            If Me.CboTrans.ListIndex = 0 Or Me.CboTrans.ListIndex = 1 Then
                Me.TxtTransSerial.text = GetTransIDSerial(1, val(Me.TxtTransID.text))
            Else
                Me.TxtTransSerial.text = Me.TxtTransID.text
            End If
        End If
    End If

End Sub

Private Sub TxtTransSerial_KeyPress(KeyAscii As Integer)
    KeyAscii = KeyAscii_Num(KeyAscii, Me.TxtTransSerial.text, 1)
End Sub

Private Sub TxtVAt2_Change()
TxtVATValue.text = txtVat2.text
End Sub

Private Sub TxtVATValue_Change()

If val(txtVat2.text) <> 0 Then
txtVat2.text = TxtVATValue.text
'XPTxtVal_Validate False
End If
End Sub

Private Sub VSFlexGrid1_AfterEdit(ByVal row As Long, ByVal Col As Long)
With VSFlexGrid1
Select Case .ColKey(Col)
Case "payed"
If .cell(flexcpChecked, row, .ColIndex("payed")) = flexChecked Then
.TextMatrix(row, .ColIndex("TransPayedValue")) = .TextMatrix(row, .ColIndex("RemainingValue"))
Else
.TextMatrix(row, .ColIndex("TransPayedValue")) = 0
End If
End Select
End With
RelineBuy
RelineBu22
End Sub
Sub CalCulteVAT(Optional Ind As Integer = 0)
Dim AccountVATCreit As String
Dim Percetage As Double

'XPDtbTrans.value = 100
'XPTxtVal = 100

If Me.TxtModFlg.text <> "R" And Me.TxtModFlg.text <> "" Then

CalcTotal Ind
If Option3.value = True And (val(DCboCashType.ListIndex) = 0 Or val(DCboCashType.ListIndex) = 1 Or val(DCboCashType.ListIndex) = 2 Or val(DCboCashType.ListIndex) = 3 Or val(DCboCashType.ListIndex) = 5 Or val(DCboCashType.ListIndex) = 7 Or val(DCboCashType.ListIndex) = 8 Or val(DCboCashType.ListIndex) = 9 Or val(DCboCashType.ListIndex) = 10 Or val(DCboCashType.ListIndex) = 11) Then
          
If SystemOptions.NotAllowedCalcVata Then
    TxtVATValue.text = 0
    txtVat2.text = 0
Else
    GetValueAddedAccount XPDtbTrans.value, AccountVATCreit, , 1, 23
         
    PercentgValueAddedAccount_Transec XPDtbTrans.value, 23, 0, AccountVATCreit, Percetage
     TxtVATValue.text = val(XPTxtVal.text) * Percetage / 100
     txtTotal = Round(val(val(XPTxtVal.text) + val(TxtVATValue.text)), 2)
End If
      
Else
TxtVATValue.text = 0
End If

End If

txtVat2.text = TxtVATValue.text
End Sub

Sub CalcTotal(Optional Ind As Integer)
    If Ind = 1 Then
        txtTotal = val(txtVat2) + val(XPTxtVal)
    ElseIf Ind = 0 Then
           Dim Percetage As Double
    Dim AccountVATCreit As String
     
    If SystemOptions.NotAllowedCalcVata Then
        TxtVATValue.text = 0
        txtVat2.text = 0
    Else
        PercentgValueAddedAccount_Transec XPDtbTrans.value, 23, 0, AccountVATCreit, Percetage
    End If
     
    
    'TxtVATValue.Text = val(XPTxtVal.Text) * Percetage / 100
    If Option3.value = True Then
    XPTxtVal.text = Round(val(txtTotal) / (Percetage / 100 + 1), 2)
    TxtVATValue.text = Round(val(XPTxtVal.text) * Percetage / 100, 2)
    txtVat2.text = TxtVATValue.text
    Else
      XPTxtVal.text = val(txtTotal.text)
    End If
    End If
End Sub
Private Sub VSFlexGrid1_BeforeEdit(ByVal row As Long, ByVal Col As Long, Cancel As Boolean)
With VSFlexGrid1
If Me.TxtModFlg.text <> "E" And Me.TxtModFlg.text <> "N" Then
Cancel = True
Exit Sub
End If
Select Case .ColKey(Col)
Case "TransPayedValue"
If .cell(flexcpChecked, row, .ColIndex("payed")) = flexChecked Then
Cancel = False
Else
End If

Case "NoteSerial1"
Cancel = True
Case "too"
Cancel = True
Case "NoteDate"
Cancel = True
Case "branch_name"
Cancel = True
Case "Note_Value"
Cancel = True
Case "PayedValue"
Cancel = True
Case "RemainingValue"
Cancel = True
Case "NetValue"
Cancel = True

End Select
End With
End Sub

Private Sub XPBtnMove_Click(index As Integer)
     On Error GoTo ErrTrap

    If Me.TxtModFlg.text = "N" Then
        clear_all Me
        Me.TxtModFlg.text = "R"
        XPBtnMove_Click (1)
    End If

    Select Case index

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

Public Sub Retrive(Optional Lngid As Long = 0)
          
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

            If rs.EOF Or rs.BOF Then
                Exit Sub
            End If
        End If
    End If
isFormFirstRun = False
    If Not IsNull(rs("general_cost_center").value) Then
        Me.DcCostCenter.BoundText = IIf(rs("general_cost_center").value = "", "", rs("general_cost_center").value)
    End If
    mClick = False
    TxtBillMaintNo.text = IIf(IsNull(rs("BillMaintNo").value), "", rs("BillMaintNo").value)
    TxtBillMaintID.text = IIf(IsNull(rs("BillMaintID").value), "", rs("BillMaintID").value)
    TxtBillTransNo.text = IIf(IsNull(rs("BillTransNo").value), "", rs("BillTransNo").value)
    TxtBillTransID.text = IIf(IsNull(rs("BillTransID").value), "", rs("BillTransID").value)
    
    txtContainerNo = IIf(IsNull(rs("ContainerNo").value), "", rs("ContainerNo").value)
    
    DCPreFix.text = IIf(IsNull(rs("Prefix").value), "", rs("Prefix").value)
    dcBranch.BoundText = IIf(IsNull(rs("branch_no").value), "", rs("branch_no").value)
    Me.DcEmp.BoundText = IIf(IsNull(rs("EmpId")), "", rs("EmpId"))
    Me.Text1.text = IIf(IsNull(rs("foxy_no").value), "", rs("foxy_no").value)
    XPTxtID.text = IIf(IsNull(rs("NoteID").value), "", val(rs("NoteID").value))
    TxtNoteSerial.text = IIf(IsNull(rs("NoteSerial").value), "", rs("NoteSerial").value)
    TxtNoteSerial1.text = IIf(IsNull(rs("NoteSerial1").value), "", rs("NoteSerial1").value)
    TxtManulaNO.text = IIf(IsNull(rs("ManulaNO").value), "", rs("ManulaNO").value)
    TxtBookNo.text = IIf(IsNull(rs("BookNo").value), "", rs("BookNo").value)
    Me.TxtContractNo.text = IIf(IsNull(rs("ContractNo").value), "", rs("ContractNo").value)
    Me.TxtContNo.text = IIf(IsNull(rs("ContNo").value), "", rs("ContNo").value)
    Me.TxtVATValue.text = IIf(IsNull(rs("VAT").value), "", rs("VAT").value)
    Me.oldtxtNoteSerial1.text = IIf(IsNull(rs("OldNoteSerial1").value), IIf(IsNull(rs("NoteSerial1").value), "", rs("NoteSerial1").value), rs("OldNoteSerial1").value)
    lbl(46).Caption = showLabel(TxtNoteSerial1, oldtxtNoteSerial1)
TxtCurrentBalance.text = IIf(IsNull(rs("CurrentBalance").value), "", rs("CurrentBalance").value)
TxtPaymentValue.text = IIf(IsNull(rs("PaymentValue").value), "", rs("PaymentValue").value)
TxtPercentage.text = IIf(IsNull(rs("Percentage").value), "", rs("Percentage").value)
TxtPercentageValue.text = IIf(IsNull(rs("PercentageValue").value), "", rs("PercentageValue").value)
 Dim mmm As String
  If Not (IsNull(rs("QrCodeImage").value)) Then
        LoadPictureFromDB Picture1, rs, "QrCodeImage", mmm
    Else
     Set Picture1.Picture = Nothing
    End If
    
    txtperson.text = IIf(IsNull(rs("person").value), "", rs("person").value)
    
    txtTradingContractID = IIf(IsNull(rs("TradingContractID").value), 0, rs("TradingContractID").value)
    
Option1.value = False
Option2.value = False
Option3.value = False
Option7.value = False
C1Elastic1.Visible = False
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
  
      CboStatus.ListIndex = IIf(IsNull(rs("Status").value), 0, rs("Status").value)
         
If IsNull(rs("commdiscounttype").value) Then
commdiscounttype.ListIndex = 0
Else
commdiscounttype.ListIndex = IIf(IsNull(rs("commdiscounttype").value), 0, rs("commdiscounttype").value)
     
End If
Commdiscountvalue.text = IIf(IsNull(rs("Commdiscountvalue").value), 0, (rs("Commdiscountvalue").value))
Commdiscountvalue1.text = IIf(IsNull(rs("Commdiscountvalue1").value), 0, (rs("Commdiscountvalue1").value))
Me.CommdiscountAccount.BoundText = IIf(IsNull(rs("CommdiscountAccount").value), "", rs("CommdiscountAccount").value)
'

 
   
    XPTxtVal.text = IIf(IsNull(rs("Note_Value").value), "", Trim(rs("Note_Value").value))
    
    Me.txtoldvalue.text = val(XPTxtVal.text)
    
    TXTBankName.text = IIf(IsNull(rs("BankName").value), "", Trim(rs("BankName").value))
 
    txtAdv_payment_value.text = IIf(IsNull(rs("Adv_payment_value").value), "", Trim(rs("Adv_payment_value").value))

    XPMTxtRemarks.text = IIf(IsNull(rs("Remark").value), "", Trim(rs("Remark").value))
    'dcproject.BoundText = IIf(IsNull(Rs("Remark").value), "", Trim(Rs("Remark").value))

    XPDtbTrans.value = IIf(IsNull(rs("NoteDate").value), Date, rs("NoteDate").value)
    Txt_DateHigri.value = IIf(IsNull(rs("NoteDateH").value), ToHijriDate(XPDtbTrans.value), rs("NoteDateH").value)
    DCboCashType.ListIndex = IIf(IsNull(rs("CashingType").value), -1, rs("CashingType").value)
    

    Me.DCCar.BoundText = IIf(IsNull(rs("CarId").value), "", rs("CarId").value)
    Me.DCDriver.BoundText = IIf(IsNull(rs("DriverId").value), "", rs("DriverId").value)

    Me.DCboUserName.BoundText = IIf(IsNull(rs("UserID").value), "", rs("UserID").value)
    DBCboClientName.BoundText = IIf(IsNull(rs("CusID").value), 0, rs("CusID").value)

  Me.DcbIqara.BoundText = val(IIf(IsNull(rs.Fields("akarid").value), 0, rs.Fields("akarid").value))
     Me.DcbUnitType.BoundText = val(IIf(IsNull(rs.Fields("UnitType").value), -1, rs.Fields("UnitType").value))
  DcbUnitType_Change
     Me.DcbUnitNo.BoundText = val(IIf(IsNull(rs.Fields("UnitNo").value), -1, rs.Fields("UnitNo").value))

TxtInterval.text = IIf(IsNull(rs("interval").value), 0, (rs("interval").value))
cbointervaltype.ListIndex = IIf(IsNull(rs("intervaltype").value), 0, (rs("intervaltype").value))
    txtrenterName.text = IIf(IsNull(rs("renterName").value), "", Trim(rs("renterName").value))



    '-----------------------------------------------------------------------------
    If IsNull(rs("NoteCashingType").value) Then
        Me.CboPayMentType.ListIndex = 0
        Me.DcboBox.BoundText = IIf(IsNull(rs("BoxID").value), 0, rs("BoxID").value)
    
        'project_Expensen_account
        Me.DcboBankName.BoundText = ""
        Me.TxtChequeNumber.text = ""
        Me.DCChequeBox.BoundText = ""
    ElseIf rs("NoteCashingType").value = 0 Then
        Me.CboPayMentType.ListIndex = 0
        Me.DcboBox.BoundText = IIf(IsNull(rs("BoxID").value), 0, rs("BoxID").value)
        Me.DcboBankName.BoundText = ""
        Me.TxtChequeNumber.text = ""
        Me.DCChequeBox.BoundText = ""
    ElseIf rs("NoteCashingType").value = 1 Then
        Me.CboPayMentType.ListIndex = 1
        Me.DcboBox.BoundText = ""
    
        Me.TxtChequeNumber.text = rs("ChqueNum").value
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

        Me.CboPayMentType.ListIndex = 2
        Me.DcboBox.BoundText = ""
        Me.DcboBankName.BoundText = rs("BankID").value
        Me.TxtChequeNumber.text = rs("ChqueNum").value
        Me.DtpChequeDueDate.value = rs("DueDate").value
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

        Me.CboPayMentType.ListIndex = 3
        Me.DcboBox.BoundText = ""
        Me.DcboBankName.BoundText = rs("BankID").value
        Me.TxtChequeNumber.text = rs("ChqueNum").value
        Me.DtpChequeDueDate.value = rs("DueDate").value
        Me.DCChequeBox.BoundText = ""
       ElseIf rs("NoteCashingType").value = 4 Then
        Me.CboPayMentType.ListIndex = 4
        Me.DcboBox.BoundText = ""
        Me.DcboBankName.BoundText = ""
        Me.TxtChequeNumber.text = ""
        Me.DCChequeBox.BoundText = ""
        DcbAccount.BoundText = IIf(IsNull(rs("AccountPaym").value), "", rs("AccountPaym").value)
      ElseIf rs("NoteCashingType").value = 5 Then
        Me.CboPayMentType.ListIndex = 5
        Me.DcboBox.BoundText = IIf(IsNull(rs("BoxID").value), 0, rs("BoxID").value)
        Me.DcboBankName.BoundText = ""
        Me.TxtChequeNumber.text = ""
        Me.DCChequeBox.BoundText = ""
    
    End If
    
    dBox = val(Me.DcboBox.BoundText)
    txtTotal.text = IIf(IsNull(rs("TotalNotesValue").value), val(txtVat2) + val(XPTxtVal), rs("TotalNotesValue").value)
    CboPayMentType_Change

    '-----------------------------------------------------------------------------
    If Not IsNull(rs("Transaction_ID").value) Then
        Me.ChkTrans.value = vbChecked
        'Me.ChkTrans.Enabled = True
        Set RsTemp = New ADODB.Recordset
        StrSQL = "Select * From Transactions Where Transaction_ID=" & rs("Transaction_ID").value
        RsTemp.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

        If Not (RsTemp.BOF Or RsTemp.EOF) Then
            Me.TxtTransID.text = RsTemp("Transaction_ID").value
            Me.TxtTransSerial.text = IIf(IsNull(RsTemp("Transaction_Serial").value), "", RsTemp("Transaction_Serial").value)

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
        Me.TxtTransID.text = rs("MaintananceID").value
        Me.TxtTransSerial.text = rs("MaintananceID").value
    ElseIf Not IsNull(rs("RevenuesID").value) Then
        Me.DcboRevenuesTypes.BoundText = rs("RevenuesID").value
        Me.ChkTrans.value = vbUnchecked
        Me.CboTrans.ListIndex = -1
        Me.TxtTransID.text = ""
        Me.TxtTransSerial.text = ""
    Else
        Me.ChkTrans.value = vbUnchecked
        Me.CboTrans.ListIndex = -1
        Me.TxtTransID.text = ""
        Me.TxtTransSerial.text = ""
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

  If DCboCashType.ListIndex = 11 Then
        
'        My_SQL = "  select id,Project_name from projects where not(REVENUE_account is null)" '
'        fill_combo Me.DBCboClientName, My_SQL
      
        DBCboClientName.BoundText = IIf(IsNull(rs("CusID").value), 0, rs("CusID").value)
        
        cus_or_sub = IIf(IsNull(rs("cus_or_sub").value), 0, rs("cus_or_sub").value)

        
        Option4.value = True
       

    End If

    If DCboCashType.ListIndex = 6 Then
        DCEmployee.BoundText = IIf(IsNull(rs("EmployeeID").value), "", rs("EmployeeID").value)
    End If
  
    If DCboCashType.ListIndex = 7 Then
        Me.DCAccounts.BoundText = IIf(IsNull(rs("AccountsCode").value), "", rs("AccountsCode").value)
    End If

 
    
    '-----------------------------------------------------------------------------
    If SystemOptions.SysAppAccoutingType = CompeleteAccounting Then
        StrSQL = "Select * From DOUBLE_ENTREY_VOUCHERS Where Notes_ID=" & val(Me.XPTxtID.text)
        StrSQL = StrSQL + " Order By DEV_ID_Line_No "
        Set RsDev = New ADODB.Recordset
        RsDev.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

        If Not (RsDev.BOF Or rs.EOF) Then
            Me.LblDevID.Caption = RsDev("Double_Entry_Vouchers_ID").value
            Me.lbl(33).Caption = IIf(IsNull(RsDev("Account_Interval_ID").value), 1, RsDev("Account_Interval_ID").value)
            RsDev.MoveFirst
Dim X As Integer

If val(Commdiscountvalue.text) = 0 Then
X = 2
Else
X = 3
End If
If DCAccounts.BoundText <> "" Then
    Me.DcboCreditSide.BoundText = ""
End If
            For i = 1 To X ' RsDev.RecordCount

                If RsDev("Credit_Or_Debit").value = 0 And i = 1 Then
                  Me.DcboDebitSide.BoundText = RsDev("Account_Code").value
                ElseIf RsDev("Credit_Or_Debit").value = 1 Then
                    Me.DcboCreditSide.BoundText = RsDev("Account_Code").value
                End If

                RsDev.MoveNext
            Next i

        End If
    End If
    If Trim(Me.DcboCreditSide.BoundText) = "" And DCAccounts.BoundText <> "" Then
        Me.DcboCreditSide.BoundText = DCAccounts.BoundText
    End If
 RetriveBillBuyData
 FillGridWithDataPayment
    '-----------------------------------------------------------------------------
    ChkTrans_Click
    '⁄—÷ «·„” Œ·’« 
    'If DCboCashType.ListIndex = 5 Then
    FillGridWithData val(Me.DBCboClientName.BoundText), TxtNoteSerial.text
    '⁄—÷ «·«ﬁ”« ÿ ·⁄ﬁÊœ  « ··«ÌÃ«—
       FillGridWithDataContract TxtContractNo.text
       
    '  End If
    XPTxtCurrent.Caption = rs.AbsolutePosition
    XPTxtCount.Caption = rs.RecordCount
    fillapprovData
 
    FramePay.Visible = False
 

'    txtTotal = val(TxtVAt2) + val(XPTxtVal)
    Exit Sub
ErrTrap:

End Sub

Private Sub SaveData()

    Dim Msg As String
    Dim RsTemp As New ADODB.Recordset
    Dim StrSQL As String
    Dim StrTemp As String
    Dim LngDevID As Long
    Dim RsDev As ADODB.Recordset
    Dim BeginTrans As Boolean
    Dim AccountVATCreit As String
    ' On Error GoTo ErrTrap
     Dim Posted As Integer
            If CheckAprroveScreen(Me.Name) = True Then
            Posted = 1
            Else
            Posted = 0
            End If
    If Me.TxtModFlg.text <> "R" Then
    Dim i As Integer
      
    Dim IntCounter As Integer
    Dim totalPayed As Double
    Dim visapayed As Double
 totalPayed = 0
 visapayed = 0
 
    If CboPayMentType.ListIndex = 5 Then '›Ì Õ«·Â «·„ ⁄œœ «· √ﬂœ „‰ ÿ—Ìﬁ… «·œ›⁄
           If val(TxtPayedValue2) <> val(XPTxtVal) + val(txtVat2) Then
             Msg = " Õ·Ì· «·ﬁÌ„Â «·„œ›Ê⁄Â «·„ ⁄œœÂ €Ì— „ÿ«»ﬁ… ·ﬁÌ„Â «·”‰œ "
            MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
            Exit Sub
           End If


    
    End If
        If DCboCashType.ListIndex = -1 Then
            Msg = "ÌÃ»  ÕœÌœ ‰Ê⁄ «·„ﬁ»Ê÷«  "
            MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
            DCboCashType.SetFocus
            Sendkeys "{F4}"
            Exit Sub
        End If



        If (commdiscounttype.ListIndex = 1 Or commdiscounttype.ListIndex = 2) And val(Commdiscountvalue.text) = 0 Then
            Msg = "ÌÃ»  ÕœÌœ    «·⁄„Ê·… "
            MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
             Commdiscountvalue.SetFocus
         '   SendKeys "{F4}"
            Exit Sub
        End If
        
        If (commdiscounttype.ListIndex = 1 Or commdiscounttype.ListIndex = 2) And CommdiscountAccount.BoundText = "" Then
            Msg = "ÌÃ»  ÕœÌœ   Õ”«» «·⁄„Ê·… "
            MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
            CommdiscountAccount.SetFocus
             Sendkeys "{F4}"
            Exit Sub
        End If
        
        
        If CboPayMentType.ListIndex = 4 Then
        
              If DcbAccount.BoundText = "" Then
            Msg = "ÌÃ»  ÕœÌœ   «·Õ”«»   «Ê·« "
            MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
            DcbAccount.SetFocus
             Sendkeys "{F4}"
            Exit Sub
        End If
        
        End If
      If CboPayMentType.ListIndex = 8 Then
        
    If val(TxtBillTransID.text) = 0 Then
            Msg = "ÌÃ» «œŒ«· —ﬁ„ «·›« Ê—… "
            MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
            TxtBillTransID.SetFocus
             Sendkeys "{F4}"
            Exit Sub
        End If
           If DBCboClientName.text = "" Then
                Msg = "ÌÃ» «Œ Ì«— «”„ «·⁄„Ì· "
                MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
                DBCboClientName.SetFocus
                Sendkeys "{F4}"
                Exit Sub
            End If
        End If
     If CboPayMentType.ListIndex = 9 Then
      If val(TxtBillMaintID.text) = 0 Then
            Msg = "ÌÃ» «œŒ«· —ﬁ„ «·›« Ê—… "
            MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
            TxtBillMaintID.SetFocus
             Sendkeys "{F4}"
            Exit Sub
        End If
            If DBCboClientName.text = "" Then
                Msg = "ÌÃ» «Œ Ì«— «”„ «·⁄„Ì· "
                MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
                DBCboClientName.SetFocus
                Sendkeys "{F4}"
                Exit Sub
            End If
        
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
            If DBCboClientName.text = "" Then
                Msg = "ÌÃ» «Œ Ì«— «”„ «·⁄„Ì· √Ê «·„Ê—œ"
                MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
                DBCboClientName.SetFocus
                Sendkeys "{F4}"
                Exit Sub
            End If
        End If
    
         '  If Me.DCboCashType.ListIndex = 8 Then
         '   If TxtContractNo.Text = "" Then
         '       Msg = "ÌÃ» « œŒ«· —ﬁ„ «·⁄ﬁœ"
         '       MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
         '       TxtContractNo.SetFocus
         '       SendKeys "{F4}"
         '       Exit Sub
         '   End If
        'End If
        
    
    
        If Me.DCboCashType.ListIndex = 5 Then
            If DBCboClientName.text = "" Then
                Msg = "ÌÃ» «Œ Ì«— «”„ ««·„‘—Ê⁄"
                MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
                DBCboClientName.SetFocus
                Sendkeys "{F4}"
                Exit Sub
            End If
        End If
    
        If Me.DCboCashType.ListIndex = 11 Then
            If DBCboClientName.text = "" Then
                Msg = "ÌÃ» «Œ Ì«— «”„ «·⁄„Ì·"
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
    
         '   If Me.DCboCashType.ListIndex = 8 Then
         '   If Me.TXTContNo.Text = "" Then
         '       If SystemOptions.UserInterface = ArabicInterface Then
         '           Msg = "ÌÃ»      «Œ Ì«— ⁄ﬁœ "
         '       Else
         '           Msg = "Select Contract Firstly"
         ''       End If
'
'                MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
'                TXTContNo.SetFocus
'                SendKeys "{F4}"
'                Exit Sub
'            End If
'        End If
        
    
         '  If Me.DCboCashType.ListIndex = 9 Then
         '   If Me.DcbIqara.BoundText = "" Then
         '       If SystemOptions.UserInterface = ArabicInterface Then
         '           Msg = "ÌÃ» «Œ Ì«— «”„ «·⁄ﬁ«—"
         '       Else
         '           Msg = "Select entity Firstly"
         '       End If
'
'                MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
'                DcbIqara.SetFocus
'                SendKeys "{F4}"
'                Exit Sub
'            End If
'
'
'                     If Me.DcbUnitType.BoundText = "" Then
'                If SystemOptions.UserInterface = ArabicInterface Then
''                    Msg = "ÌÃ» «Œ Ì«—    ‰Ê⁄ «·ÊÕœ…"
 '               Else
 '                   Msg = "Select unit type Firstly"
 '               End If
'
'                MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
'                DcbUnitType.SetFocus
'                SendKeys "{F4}"
'                Exit Sub
'            End If
            
'                      If Me.DcbUnitNo.BoundText = "" Then
'                If SystemOptions.UserInterface = ArabicInterface Then
'                    Msg = "ÌÃ» «Œ Ì«—    —ﬁ„ «·ÊÕœ…   "
'                Else
'                    Msg = "Select unit no Firstly"
'                End If
'
''                MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
 '               DcbUnitNo.SetFocus
 '               SendKeys "{F4}"
 '               Exit Sub
 '           End If
 '
            
       '    If Me.txtinterval.Text = "" Then
       '         If SystemOptions.UserInterface = ArabicInterface Then
       '             Msg = "ÌÃ»   ÕœÌœ «·„œ…   "
       '         Else
       '             Msg = "Select Account Firstly"
       '         End If
'
'                MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
'                txtinterval.SetFocus
'                SendKeys "{F4}"
'                Exit Sub
'            End If
'
'        End If
        
        If val(XPTxtVal.text) = 0 Then
            Msg = "ÌÃ» «œŒ«· ﬁÌ„… «·„ﬁ»Ê÷«  "
            MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
            '        XPTxtVal.SetFocus
            Exit Sub
        End If

        If Not IsNumeric(XPTxtVal.text) Then
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

            If Trim(Me.TxtTransSerial.text) = "" Then
                Msg = "»—Ã«¡ ≈œŒ«· —ﬁ„ «·›« Ê—…..!!!"
                MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
                TxtTransSerial.SetFocus
                Exit Sub
            Else

                If Me.CboTrans.ListIndex = 0 Then
                    StrTemp = GetTransIDSerial(0, , Trim(Me.TxtTransSerial.text), 2)

                    If CheckDebitTrans(val(StrTemp)) = False Then
                        Exit Sub
                    End If

                ElseIf Me.CboTrans.ListIndex = 1 Then
                    StrTemp = GetTransIDSerial(0, , Trim(Me.TxtTransSerial.text), 5)

                    If CheckDebitTrans(val(StrTemp)) = False Then
                        Exit Sub
                    End If

                ElseIf Me.CboTrans.ListIndex = 2 Then

                    If CheckDebitMaintaince(val(Me.TxtTransSerial.text)) = False Then
                        Exit Sub
                    End If

                ElseIf Me.CboTrans.ListIndex = 3 Then
                    Msg = "⁄›Ê« .. Ã«—Ï  ÿÊÌ— «·»—‰«„Ã .. ·⁄„· «·„ﬁ»Ê÷«  „‰ «·Œœ„« "
                    MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
                    Exit Sub
                End If
            End If
        End If

        If Me.CboPayMentType.ListIndex = -1 Then
            Msg = "ÌÃ» ≈Œ Ì«— ÿ—Ìﬁ… «·œ›⁄...!!"
            MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
            CboPayMentType.SetFocus
            Sendkeys "{F4}"
            Exit Sub
        End If

        If Me.CboPayMentType.ListIndex = 0 Then
            If Me.DcboBox.BoundText = "" Then
                Msg = "ÌÃ» ≈Œ Ì«— «·Œ“‰…...!!"
                MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
                DcboBox.SetFocus
                Sendkeys "{F4}"
                Exit Sub
            End If
            
  ElseIf Me.CboPayMentType.ListIndex = 5 And CheckMult_Cash() = True Then
                   
            GetUserData user_id, usertype, userbranchid, dstore, dBox, , EmpID
'DcEmp.BoundText = EmpID
' CboPaymentType.ListIndex = 0
DcboBox.BoundText = dBox

            If Me.DcboBox.BoundText = "" Then
                Msg = "ÌÃ» ≈Œ Ì«— «·Œ“‰…...!!"
                MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
'                DcboBox.SetFocus
                 Sendkeys "{F4}"
                Exit Sub
            End If
        ElseIf Me.CboPayMentType.ListIndex = 1 Then
      
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
    
                If TXTBankName.text = "" Then
                           
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
        
                If Trim$(Me.TxtChequeNumber.text) = "" Then
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

                If Trim$(Me.TxtChequeNumber.text) = "" Then
                    Msg = "ÌÃ» ﬂ «»… —ﬁ„ «·‘Ìﬂ...!!"
                    MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
                    TxtChequeNumber.SetFocus
                    Exit Sub
                End If
            End If
    
        ElseIf Me.CboPayMentType.ListIndex = 2 Then

            If Me.DcboBankName.BoundText = "" Then
                Msg = "ÌÃ» ≈Œ Ì«— «·»‰ﬂ...!!"
                MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
                DcboBankName.SetFocus
                Sendkeys "{F4}"
                Exit Sub
            End If

            If Trim$(Me.TxtChequeNumber.text) = "" Then
                Msg = "ÌÃ» ﬂ «»… —ﬁ„ «·ÕÊ«·Â...!!"
                MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
                TxtChequeNumber.SetFocus
                Exit Sub
            End If
     
        ElseIf Me.CboPayMentType.ListIndex = 3 Then

            If Me.DcboBankName.BoundText = "" Then
                Msg = "ÌÃ» ≈Œ Ì«— «·»‰ﬂ...!!"
                MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
                DcboBankName.SetFocus
                Sendkeys "{F4}"
                Exit Sub
            End If

            If Trim$(Me.TxtChequeNumber.text) = "" Then
                Msg = "ÌÃ» ﬂ «»… —ﬁ„ «·‘Ìﬂ...!!"
                MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
                TxtChequeNumber.SetFocus
                Exit Sub
            End If
     
        End If

        Dim notes_result As String
        Dim Vchr_result As String
        
        If TxtNoteSerial1.text = "" Then
            Vchr_result = Voucher_coding(val(my_branch), XPDtbTrans.value, 2, 4, , , DCPreFix.text)

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
    
        If TxtNoteSerial.text = "" Or val(TxtNoteSerial.text) = 0 Then
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

        If TxtModFlg.text = "N" Then
            XPTxtID.text = CStr(new_id("Notes", "NoteID", "", True))
            'Me.TxtNoteSerial.text = CStr(new_id("Notes", "NoteSerial", "", True, "NoteType=4"))
            rs.AddNew
       
            rs("NoteID").value = val(XPTxtID.text)
            Me.oldtxtNoteSerial1.text = Trim$(Me.TxtNoteSerial1.text)
         
        ElseIf TxtModFlg.text = "E" Then
            StrSQL = "Delete From DOUBLE_ENTREY_VOUCHERS Where Notes_ID=" & val(XPTxtID.text)
            Cn.Execute StrSQL, , adExecuteNoRecords
            StrSQL = "Delete From TblMultuPayment Where NoteID=" & val(XPTxtID.text)
            Cn.Execute StrSQL, , adExecuteNoRecords
        
            StrSQL = "Delete  marakes_taklefa_temp  where kedno =" & val(Text1.text)
            Cn.Execute StrSQL, , adExecuteNoRecords


   StrSQL = " delete   notes where NoteType= 2000   and  NoteSerial='" & TxtNoteSerial.text & "'"
  
Cn.Execute StrSQL

'StrSQL = "  delete TblBillBuyPayment2 where noteid=" & val(XPTxtID.Text)
' Cn.Execute StrSQL
         End If


             If DCboCashType.ListIndex = 5 Then
                '«·„‘«—Ì⁄
                Dim pstate As Integer '·Ê «·„‘—Ê⁄ «›  «ÕÌ
          
             '   account_codeLegal = get_project_Account(val(DBCboClientName.BoundText), "legal")
                     pstate = val(get_project_Account(val(DBCboClientName.BoundText), "pstate"))

    '   If pstate = 1 Then Option7.value = True Else Option7.value = False


      End If

        rs("branch_no").value = val(Me.dcBranch.BoundText)
        rs("EmpId").value = IIf(Me.DcEmp.BoundText = "", Null, (Me.DcEmp.BoundText))
        rs("foxy_no").value = val(Text1.text)
        rs("general_cost_center").value = IIf(Me.DcCostCenter.BoundText = "", "", Me.DcCostCenter.BoundText)
    rs("Prefix").value = IIf(DCPreFix.text = "", Null, DCPreFix.text)

        rs("CarId").value = IIf(Me.DCCar.BoundText = "", Null, (Me.DCCar.BoundText))
        rs("DriverId").value = IIf(Me.DCDriver.BoundText = "", Null, (Me.DCDriver.BoundText))
    
        If TxtNoteSerial1.text = "" Then
            TxtNoteSerial1.text = Voucher_coding(val(my_branch), XPDtbTrans.value, 2, 4, , , DCPreFix.text)
        End If
    
        If TxtNoteSerial.text = "" Or val(TxtNoteSerial.text) = 0 Then
            TxtNoteSerial.text = Notes_coding(val(my_branch), XPDtbTrans.value)
        End If
        
             If CboStatus.ListIndex <> 0 Then
        TxtNoteSerial.text = ""
        
        End If
        rs("TradingContractID").value = IIf(txtTradingContractID.text = "", 0, val(txtTradingContractID.text))
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
       
    
        rs("ContainerNo").value = IIf(Trim(Me.txtContainerNo.text) = "", Null, Trim(Me.txtContainerNo.text))
        rs("ManulaNO").value = IIf(Trim(Me.TxtManulaNO.text) = "", Null, Trim(Me.TxtManulaNO.text))
        rs("ManualNo").value = IIf(Trim(Me.TxtManulaNO.text) = "", Null, Trim(Me.TxtManulaNO.text))
        rs("BookNo").value = IIf(Trim(Me.TxtBookNo.text) = "", Null, Trim(Me.TxtBookNo.text))
        
        '
        rs("NoteSerial").value = IIf(Trim(Me.TxtNoteSerial.text) = "", Null, Trim(Me.TxtNoteSerial.text))
        rs("NoteSerial1").value = IIf(Trim(Me.TxtNoteSerial1.text) = "", Null, Trim(Me.TxtNoteSerial1.text))
        rs("OldNoteSerial1").value = Trim$(Me.oldtxtNoteSerial1.text) '
    
        rs("person").value = IIf(txtperson.text = "", "", Trim(txtperson.text))
        rs("Note_Value").value = IIf(XPTxtVal.text = "", Null, val(XPTxtVal.text))
        rs("Adv_payment_value").value = IIf(txtAdv_payment_value.text = "", Null, val(txtAdv_payment_value.text))
        rs("VAT").value = IIf(TxtVATValue.text = "", Null, val(TxtVATValue.text))
    
        '    Rs("Remark").value = IIf(dcproject.BoundText = "", "", Trim(dcproject.BoundText))
        If lblinvoices.Caption = "" Then
        rs("Remark").value = IIf(XPMTxtRemarks.text = "", "", Trim(XPMTxtRemarks.text))
        Else
        rs("Remark").value = IIf(XPMTxtRemarks.text = "", "", Trim(XPMTxtRemarks.text)) & vbEnter & lblinvoices.Caption
        End If
        
        rs("BankName").value = IIf(TXTBankName.text = "", "", Trim(TXTBankName.text))
        rs("NoteType").value = 4
        rs("NoteDate").value = XPDtbTrans.value
        rs("BillTransNo").value = TxtBillTransNo.text
        rs("BillTransID").value = val(TxtBillTransID.text)
        rs("BillMaintNo").value = TxtBillMaintNo.text
        rs("BillMaintID").value = val(TxtBillMaintID.text)
        'rs("NoteDate").value = Format$(Date, "dd-mm-yyyy")
        rs("NoteDateH").value = Me.Txt_DateHigri.value

        Select Case DCboCashType.ListIndex

            Case 0, 1

                If Me.ChkTrans.value = vbChecked Then
                    If Me.CboTrans.ListIndex = 0 Or Me.CboTrans.ListIndex = 1 Then
                        rs("Transaction_ID").value = val(Me.TxtTransID.text)
                        rs("MaintananceID").value = Null
                    ElseIf Me.CboTrans.ListIndex = 2 Then
                        rs("Transaction_ID").value = Null
                        rs("MaintananceID").value = val(Me.TxtTransID.text)
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

        rs("CashingType").value = DCboCashType.ListIndex
        calcnet
        '
        rs("TotalNotesValue").value = val(Me.txtTotal.text)
        rs("commdiscounttype").value = commdiscounttype.ListIndex
        rs("Commdiscountvalue").value = val(Me.Commdiscountvalue.text)
        rs("Commdiscountvalue1").value = val(Me.Commdiscountvalue1.text)
        rs("CommdiscountAccount").value = IIf(CommdiscountAccount.text = "", Null, CommdiscountAccount.BoundText)
        
        rs("Status").value = CboStatus.ListIndex
        rs("CurrentBalance").value = val(TxtCurrentBalance.text)
        rs("PaymentValue").value = val(TxtPaymentValue.text)
        rs("Percentage").value = val(TxtPercentage.text)
        rs("PercentageValue").value = val(TxtPercentageValue.text)
        
        If Me.DCboCashType.ListIndex = 0 Or Me.DCboCashType.ListIndex = 1 Or Me.DCboCashType.ListIndex = 2 Or Me.DCboCashType.ListIndex = 4 Or Me.DCboCashType.ListIndex = 8 Or Me.DCboCashType.ListIndex = 9 Or Me.DCboCashType.ListIndex = 10 Or Me.DCboCashType.ListIndex = 11 Or Me.DCboCashType.ListIndex = 12 Then
            rs("CusID").value = IIf(DBCboClientName.text = "", Null, DBCboClientName.BoundText)
     
        ElseIf Me.DCboCashType.ListIndex = 5 Then
            Dim X As Double
                    If IsNull(rs("note_count").value) Then
                         rs("note_count").value = CStr(new_id("Notes", "note_count", " ", True, " project_id=" & val(DBCboClientName.BoundText) & ""))
                    End If
            
            If Option4.value = True Then
                X = get_project_customer_id(val(DBCboClientName.BoundText), "End_user_Account")
            Else
                X = get_project_customer_id(DBCboClientName.BoundText, "sub_contractor_Account")
            End If

            rs("CusID").value = X
     
        Else
            rs("CusID").value = Null
        End If

        '--------------------------------------------------------------------------
        'ÿ—Ìﬁ… «·œ›⁄ «·‰ﬁœÏ «Ê «·‘Ìﬂ
        If Me.CboPayMentType.ListIndex = 0 Then
            rs("NoteCashingType").value = 0
            rs("BoxID").value = IIf(DcboBox.BoundText = "", Null, DcboBox.BoundText)
            rs("BankID").value = Null
            rs("ChqueNum").value = Null
            rs("DueDate").value = Null
        
        ElseIf Me.CboPayMentType.ListIndex = 1 Then
            rs("NoteCashingType").value = 1
            rs("BoxID").value = Null

            If SystemOptions.ChequeBox = False Then
        
                rs("BankID").value = val(Me.DcboBankName.BoundText)
            Else
                rs("BankID").value = Null
            End If
        
            rs("ChqueNum").value = Trim$(Me.TxtChequeNumber.text)
            rs("DueDate").value = Me.DtpChequeDueDate.value

            If SystemOptions.ChequeBox = True Then
                rs("ChequeBoxID").value = IIf(DCChequeBox.BoundText = "", Null, DCChequeBox.BoundText)
            Else
                rs("ChequeBoxID").value = Null
                
            End If
                
        ElseIf Me.CboPayMentType.ListIndex = 2 Then
            rs("NoteCashingType").value = 2
            rs("BoxID").value = Null
            rs("BankID").value = val(Me.DcboBankName.BoundText)
            rs("ChqueNum").value = Trim$(Me.TxtChequeNumber.text)
            rs("DueDate").value = Me.DtpChequeDueDate.value
            rs("ChequeBoxID").value = Null
                
        ElseIf Me.CboPayMentType.ListIndex = 3 Then
            rs("NoteCashingType").value = 3
            rs("BoxID").value = Null
            rs("BankID").value = val(Me.DcboBankName.BoundText)
            rs("ChqueNum").value = Trim$(Me.TxtChequeNumber.text)
            rs("DueDate").value = Me.DtpChequeDueDate.value
            rs("ChequeBoxID").value = Null
       ElseIf Me.CboPayMentType.ListIndex = 4 Then
            rs("NoteCashingType").value = 4
            rs("BoxID").value = Null
            rs("BankID").value = Null
            rs("ChqueNum").value = Null
            rs("DueDate").value = Null
            rs("AccountPaym").value = IIf(Trim(DcbAccount.BoundText) = "", Null, DcbAccount.BoundText)
         ElseIf Me.CboPayMentType.ListIndex = 5 Then
            rs("NoteCashingType").value = 5
            rs("BoxID").value = IIf(DcboBox.BoundText = "", Null, DcboBox.BoundText)
            rs("BankID").value = Null
            rs("ChqueNum").value = Null
            rs("DueDate").value = Null
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
    
     '  If DCboCashType.ListIndex = 8 Then
     '       rs("ContractNo").value = IIf(TxtContractNo.Text = "", Null, TxtContractNo.Text)
     '       rs("ContNo").value = IIf(TXTContNo.Text = "", Null, TXTContNo.Text)
     '       Else
     '        rs("ContractNo").value = Null
     '        rs("ContNo").value = Null
     '   End If
        
        
   '  If DCboCashType.ListIndex = 9 Then
   ' rs("akarid").value = IIf(val(Me.DcbIqara.BoundText) <> 0, val(DcbIqara.BoundText), Null)
   '  rs.Fields("UnitType").value = IIf(Me.DcbUnitType.BoundText <> "", val(DcbUnitType.BoundText), Null)
   '  rs.Fields("UnitNo").value = IIf(Me.DcbUnitNo.BoundText <> "", val(DcbUnitNo.BoundText), Null)
  '   rs("interval").value = IIf(txtinterval.Text = "", Null, val(txtinterval.Text))
  '   rs("intervaltype").value = val(cbointervaltype.ListIndex)
  '   rs("renterName").value = IIf(txtrenterName.Text = "", Null, txtrenterName.Text)
  '            If cbointervaltype.ListIndex = 0 Then
  '            rs("allowdate").value = DateAdd("d", val(txtinterval), XPDtbTrans.value)
  '            ElseIf cbointervaltype.ListIndex = 1 Then
  '            rs("allowdate").value = DateAdd("M", val(txtinterval), XPDtbTrans.value)
  '
  '          ElseIf cbointervaltype.ListIndex = 2 Then
  '            rs("allowdate").value = DateAdd("YYYY", val(txtinterval), XPDtbTrans.value)
  '
  '           End If
  '                rs("allowdateH").value = ToHijriDate(rs("allowdate").value)
  '
  '          Else
  '        rs("akarid").value = Null
  '   rs.Fields("UnitType").value = Null
  '   rs.Fields("UnitNo").value = Null
  '   rs("interval").value = Null
  '   rs("intervaltype").value = Null
  '   rs("renterName").value = Null
          
  '      End If
              
              
              
        
        rs("sanad_year").value = year(XPDtbTrans.value)
        rs("sanad_month").value = Month(XPDtbTrans.value)
    
        If DCboCashType.ListIndex = 5 Then
            rs("note_value_by_characters").value = Trim$(Me.lbl(18).Caption) 'WriteNo(val(Me.XPTxtVal.Text), 0, True)
        Else
            rs("note_value_by_characters").value = Trim$(Me.lbl(18).Caption)
        End If

        If Option4.value = True Then
            rs("cus_or_sub").value = 0 '⁄„Ì· ‰Â«∆Ì
        Else
            rs("cus_or_sub").value = 1 '⁄„Ì· »«ÿ‰
        End If
    
        rs.update

        saveChequeBoxContents (XPTxtID.text)
        SaveMultyPayment val(XPTxtID.text)
        '==========================================================================
    If SystemOptions.IsCheque = True And CboPayMentType.ListIndex = 1 Then GoTo endSave
        Line1 = setfoxy_Line
        Line2 = setfoxy_Line
        Line3 = setfoxy_Line
        Line4 = setfoxy_Line

        ' ”ÃÌ· ﬁÌÊœ
        If SystemOptions.SysAppAccoutingType = CompeleteAccounting Then
            LngDevID = new_id("DOUBLE_ENTREY_VOUCHERS", "Double_Entry_Vouchers_ID", "", True)
            Set RsDev = New ADODB.Recordset
        '    RsDev.Open "DOUBLE_ENTREY_VOUCHERS", Cn, adOpenStatic, adLockOptimistic, adCmdTable
                      StrSQL = "SELECT     * from dbo.DOUBLE_ENTREY_VOUCHERS Where (1 = -1)"
   RsDev.Open StrSQL, Cn, adOpenKeyset, adLockOptimistic, adCmdText
 Dim lineno As Integer
 lineno = 1
            '«·ÿ—› «·„œÌ‰
       Dim newdes As String
       newdes = ""
      If val(DCboCashType.ListIndex) = 5 Then
      If SystemOptions.UserInterface = ArabicInterface Then
         newdes = newdes & " " & " ·„‘—Ê⁄ "
         newdes = newdes & DBCboClientName.text
         newdes = newdes & " " & " ﬂÊœ «·„‘—Ê⁄ "
         newdes = newdes & TxtCustCode.text
       Else
          newdes = newdes & " " & " project "
         newdes = newdes & DBCboClientName.text
         newdes = newdes & " " & " Code "
         newdes = newdes & TxtCustCode.text
      End If
       End If
       If val(CboPayMentType.ListIndex) = 1 Then
       If SystemOptions.UserInterface = ArabicInterface Then
       newdes = newdes & " »‰«¡ ⁄·Ï ‘Ìﬂ —ﬁ„"
       Else
       newdes = newdes & " Based on check No."
       End If
       newdes = newdes & " " & TxtChequeNumber.text
       End If
       If val(CboPayMentType.ListIndex) = 2 Then
       If SystemOptions.UserInterface = ArabicInterface Then
       newdes = newdes & "»‰«¡ ⁄·Ï ÕÊ«·… »‰ﬂÌ…  —ﬁ„"
       Else
       newdes = newdes & "Based on bank transfer No."
       End If
       newdes = newdes & " " & TxtChequeNumber.text
       End If
        If val(CboPayMentType.ListIndex) = 3 Then
       If SystemOptions.UserInterface = ArabicInterface Then
       newdes = newdes & "»‰«¡ ⁄·Ï ‘Ìﬂ „Õ’· —ﬁ„"
       Else
       newdes = newdes & "Based on check  No."
       End If
       newdes = newdes & " " & TxtChequeNumber.text
       End If
       
       
               Dim BranchID As Integer
         Dim BranchID2 As Integer
         Dim BranchID3  As Integer
         
         Dim DeptSide As String
         Dim credit_side As String
         Dim credit_side3 As String
         
                  
               
                 
                  BranchID = val(Me.dcBranch.BoundText)
                If DCboCashType.ListIndex < 3 Then
                  BranchID2 = GetBrnchCustomer(val(DBCboClientName.BoundText))
                ElseIf DCboCashType.ListIndex = 5 Then
                    BranchID2 = GetBrnchproject(val(DBCboClientName.BoundText))
                ElseIf DCboCashType.ListIndex < 6 Then
                    BranchID2 = GetBrnchEmployee(val(DBCboClientName.BoundText))
                End If
                   
             BranchID3 = GetBrnchBank(val(DcboBankName.BoundText))
            If DCboCashType.ListIndex = 4 Then
              '  BranchID2 = val(Me.DcbEmpBranch.BoundText)
              BranchID2 = val(Me.dcBranch.BoundText)
             ElseIf DCboCashType.ListIndex = 1 Or DCboCashType.ListIndex = 0 Then
            
             
             End If
             

                                  DeptSide = getBranchCurrentAccount(BranchID)
                                                 credit_side = getBranchCurrentAccount(BranchID2)
                                                  credit_side3 = getBranchCurrentAccount(BranchID3)
                                                 
'
'                                                 If DCboCashType.ListIndex = 4 Or DCboCashType.ListIndex = 0 Or DCboCashType.ListIndex = 1 Then
'                    If BranchID <> BranchID2 And BranchID2 <> 0 And DCboCashType.ListIndex = 4 Then
'                        RsDev("branch_id").value = val(Me.DcbEmpBranch.BoundText)
'                    ElseIf BranchID <> BranchID2 And (DCboCashType.ListIndex = 0 Or DCboCashType.ListIndex = 1) Then
'                        RsDev("branch_id").value = GetBrnchCustomer(val(DBCboClientName.BoundText))
'                    Else
'                       RsDev("branch_id").value = val(Me.dcBranch.BoundText)
'                    End If
'
'   Else
'            RsDev("branch_id").value = val(Me.dcBranch.BoundText)
'   End If
'
 If val(CboPayMentType.ListIndex) = 5 Then
 PGMultyPayment val(XPTxtID.text), lineno, Line1, XPMTxtRemarks.text & CHR(13) & newdes, Posted
 Else
   If CboStatus.ListIndex = 0 Then
            RsDev.AddNew
            
            
            If DCboCashType.ListIndex = 4 Or DCboCashType.ListIndex = 0 Or DCboCashType.ListIndex = 1 Then
                If BranchID <> BranchID2 And BranchID2 <> 0 And DCboCashType.ListIndex = 4 Then
                    RsDev("branch_id").value = val(Me.dcBranch.BoundText)
                ElseIf ((BranchID <> BranchID2) Or ((BranchID <> BranchID3))) And (DCboCashType.ListIndex = 0 Or DCboCashType.ListIndex = 1) Then
                    RsDev("branch_id").value = GetBrnchBank(val(DcboBankName.BoundText))
                Else
                    RsDev("branch_id").value = val(Me.dcBranch.BoundText)
                End If
            
            Else
                RsDev("branch_id").value = val(Me.dcBranch.BoundText)
            End If
            'RsDev("branch_id").value = val(Me.dcBranch.BoundText)
            RsDev("Double_Entry_Vouchers_ID").value = LngDevID
            RsDev("DEV_ID_Line_No").value = lineno
            RsDev("DEV_ID_Line_No1").value = Line1
            If Posted = 1 Then
            RsDev("Posted").value = 1
            Else
            RsDev("Posted").value = Null
            End If
            RsDev("Account_Code").value = Me.DcboDebitSide.BoundText
            RsDev("NextAccount_Code").value = Me.DcboCreditSide.BoundText
            RsDev("Value").value = val(Me.XPTxtVal.text) - commvalue + val(TxtVATValue.text)
            RsDev("Credit_Or_Debit").value = 0
             RsDev("Double_Entry_Vouchers_Description").value = XPMTxtRemarks.text & CHR(13) & newdes
            
            'RsDev("Double_Entry_Vouchers_Description").value = dcproject.BoundText
            
            RsDev("Notes_ID").value = val(XPTxtID.text)
            RsDev("RecordDate").value = Me.XPDtbTrans.value
            RsDev("RecordDateH").value = ToHijriDate(XPDtbTrans.value)
            RsDev("UserID").value = Me.DCboUserName.BoundText
            RsDev("Account_Interval_ID").value = SystemOptions.SysCurrentAccountIntervalID
                If DCboCashType.ListIndex = 5 Then
           '  RsDev("project_id").value = IIf(DBCboClientName.BoundText = "", 0, DBCboClientName.BoundText)
               End If

            RsDev.update
            
            lineno = lineno + 1
            If Option3.value = True Then
            
                  RsDev.AddNew
                 RsDev("branch_id").value = val(Me.dcBranch.BoundText)
                 RsDev("Double_Entry_Vouchers_ID").value = LngDevID
                 RsDev("DEV_ID_Line_No").value = lineno
                 RsDev("DEV_ID_Line_No1").value = Line1
                 RsDev("Account_Code").value = DcboCreditSide.BoundText
                 RsDev("NextAccount_Code").value = DcboCreditSide.BoundText
                 RsDev("Value").value = val(TxtVATValue.text)
                 RsDev("Credit_Or_Debit").value = 0
                  RsDev("Double_Entry_Vouchers_Description").value = "œ›⁄… „ﬁœ„…" & DcboCreditSide.text & XPMTxtRemarks.text & CHR(13) & newdes
                 
                 'RsDev("Double_Entry_Vouchers_Description").value = dcproject.BoundText
                 
                 RsDev("Notes_ID").value = val(XPTxtID.text)
                 RsDev("RecordDate").value = Me.XPDtbTrans.value
                 RsDev("RecordDateH").value = ToHijriDate(XPDtbTrans.value)
                 RsDev("UserID").value = Me.DCboUserName.BoundText
                 RsDev("Account_Interval_ID").value = SystemOptions.SysCurrentAccountIntervalID
                     If DCboCashType.ListIndex = 5 Then
                '  RsDev("project_id").value = IIf(DBCboClientName.BoundText = "", 0, DBCboClientName.BoundText)
                    End If
                 If Posted = 1 Then
                 RsDev("Posted").value = 1
                 Else
                 RsDev("Posted").value = Null
                 End If
                 RsDev.update
            End If
            lineno = lineno + 1
            'End If
        ''/////////Œ’„ „”„ÊÕ »Â
        If val(DCboCashType.ListIndex) = 0 Then
        If SystemOptions.AllowAcceleratepayment = True And val(TxtPercentageValue.text) > 0 Then
                RsDev.AddNew
            RsDev("branch_id").value = val(Me.dcBranch.BoundText)
            RsDev("Double_Entry_Vouchers_ID").value = LngDevID
            RsDev("DEV_ID_Line_No").value = lineno
            RsDev("DEV_ID_Line_No1").value = Line1
            RsDev("Account_Code").value = get_account_code_branch(12, 0)
            RsDev("NextAccount_Code").value = DcboCreditSide.BoundText
            RsDev("Value").value = val(TxtPercentageValue.text)
            RsDev("Credit_Or_Debit").value = 0
            RsDev("Double_Entry_Vouchers_Description").value = XPMTxtRemarks.text & CHR(13) & "Œ’„ „”„ÊÕ »Â"
            RsDev("Notes_ID").value = val(XPTxtID.text)
            RsDev("RecordDate").value = Me.XPDtbTrans.value
            RsDev("RecordDateH").value = ToHijriDate(XPDtbTrans.value)
            RsDev("UserID").value = Me.DCboUserName.BoundText
            RsDev("Account_Interval_ID").value = SystemOptions.SysCurrentAccountIntervalID
            If Posted = 1 Then
            RsDev("Posted").value = 1
            Else
            RsDev("Posted").value = Null
            End If
            RsDev.update
            
            lineno = lineno + 1
            
            
            
       End If
    End If
    End If
   End If
'«·⁄„Ê·« 
If commvalue > 0 Then
   If CboStatus.ListIndex = 0 Then
            RsDev.AddNew
            RsDev("branch_id").value = val(Me.dcBranch.BoundText)
            RsDev("Double_Entry_Vouchers_ID").value = LngDevID
            RsDev("DEV_ID_Line_No").value = lineno
            RsDev("DEV_ID_Line_No1").value = Line1
            RsDev("Account_Code").value = Me.CommdiscountAccount.BoundText
            RsDev("NextAccount_Code").value = DcboCreditSide.BoundText
            RsDev("Value").value = commvalue
            RsDev("Credit_Or_Debit").value = 0
             RsDev("Double_Entry_Vouchers_Description").value = "Œ’„ ⁄„Ê·… ·’«·Õ" & CommdiscountAccount.text & XPMTxtRemarks.text & CHR(13) & newdes
            
            'RsDev("Double_Entry_Vouchers_Description").value = dcproject.BoundText
            
            RsDev("Notes_ID").value = val(XPTxtID.text)
            RsDev("RecordDate").value = Me.XPDtbTrans.value
            RsDev("RecordDateH").value = ToHijriDate(XPDtbTrans.value)
            RsDev("UserID").value = Me.DCboUserName.BoundText
            RsDev("Account_Interval_ID").value = SystemOptions.SysCurrentAccountIntervalID
                If DCboCashType.ListIndex = 5 Then
           '  RsDev("project_id").value = IIf(DBCboClientName.BoundText = "", 0, DBCboClientName.BoundText)
               End If
            If Posted = 1 Then
            RsDev("Posted").value = 1
            Else
            RsDev("Posted").value = Null
            End If
            RsDev.update
            
            lineno = lineno + 1
            End If
End If

'‰Â«Ì… «·⁄„Ê·« 
            
            
            '«·ÿ—› «·œ«∆‰
       If CboStatus.ListIndex = 0 Then
               If DCboCashType.ListIndex = 11 Then 'Õ«·Â ⁄œ… „” Œ·’« 
            'rs("project_id").value = IIf(DBCboClientName.BoundText = "", 0, DBCboClientName.BoundText)
            GLByProjectInvoice CDbl(LngDevID), CDbl(lineno), Line2
         Else
        
            RsDev.AddNew
            
            
              If DCboCashType.ListIndex = 4 Or DCboCashType.ListIndex = 0 Or DCboCashType.ListIndex = 1 Then
                If BranchID <> BranchID2 And BranchID2 <> 0 And DCboCashType.ListIndex = 4 Then
                    RsDev("branch_id").value = val(Me.dcBranch.BoundText)
                ElseIf BranchID <> BranchID2 And (DCboCashType.ListIndex = 0 Or DCboCashType.ListIndex = 1) Then
                    RsDev("branch_id").value = GetBrnchCustomer(val(DBCboClientName.BoundText))
                Else
                    RsDev("branch_id").value = val(Me.dcBranch.BoundText)
                End If
            
            Else
                RsDev("branch_id").value = val(Me.dcBranch.BoundText)
            End If
            RsDev("Double_Entry_Vouchers_ID").value = LngDevID
            RsDev("DEV_ID_Line_No").value = lineno
            RsDev("DEV_ID_Line_No1").value = Line2
            RsDev("Account_Code").value = Me.DcboCreditSide.BoundText
            RsDev("NextAccount_Code").value = Me.DcboDebitSide.BoundText
            
            If Option3.value = True Then
                RsDev("Value").value = val(Me.XPTxtVal.text) - commvalue + val(TxtVATValue.text)
            Else
                RsDev("Value").value = val(Me.XPTxtVal.text)
            End If
            RsDev("Credit_Or_Debit").value = 1
            If SystemOptions.PaymentIntoAccouStat = True And val(DCboCashType.ListIndex) = 5 Then
            RsDev("project_id").value = val(DBCboClientName.BoundText)
            RsDev("projectid").value = val(DBCboClientName.BoundText)
            End If
         '   RsDev("Double_Entry_Vouchers_Description").value = XPMTxtRemarks.text
           '  If DCboCashType.ListIndex = 9 Then
            
           '  newdes = "  ⁄—»Ê‰ ÕÃ“  «·ÊÕœ…   " & DcbUnitType.Text & "  »—ﬁ„   " & DcbUnitNo.Text & "  ··„” √Ã— " & txtrenterName
           ' End If
        '    If Me.DCboCashType = 0 Then
        '    newdes
            RsDev("Double_Entry_Vouchers_Description").value = XPMTxtRemarks.text & CHR(13) & newdes & CHR(13) & lblinvoices.Caption
               
               
            ' RsDev("Double_Entry_Vouchers_Description").value = dcproject.BoundText
            RsDev("Notes_ID").value = val(XPTxtID.text)
            RsDev("RecordDate").value = Me.XPDtbTrans.value
            RsDev("RecordDateH").value = ToHijriDate(XPDtbTrans.value)
            RsDev("UserID").value = Me.DCboUserName.BoundText
            RsDev("Account_Interval_ID").value = SystemOptions.SysCurrentAccountIntervalID
               If DCboCashType.ListIndex = 5 Then 'okkkkkkkk
                 RsDev("project_id").value = IIf(DBCboClientName.BoundText = "", 0, DBCboClientName.BoundText)
               End If
            RsDev("CarId").value = IIf(Me.DCCar.BoundText = "", Null, (Me.DCCar.BoundText))
              If Posted = 1 Then
            RsDev("Posted").value = 1
            Else
            RsDev("Posted").value = Null
            End If
            RsDev.update
            
            
            End If
       '''///////
       If val(DCboCashType.ListIndex) = 0 Then
       If SystemOptions.AllowAcceleratepayment = True And val(TxtPercentageValue.text) > 0 Then
       lineno = lineno + 1
            RsDev.AddNew
            RsDev("branch_id").value = val(Me.dcBranch.BoundText)
            RsDev("Double_Entry_Vouchers_ID").value = LngDevID
            RsDev("DEV_ID_Line_No").value = lineno
            RsDev("DEV_ID_Line_No1").value = Line2
            RsDev("Account_Code").value = Me.DcboCreditSide.BoundText
            RsDev("NextAccount_Code").value = Me.DcboDebitSide.BoundText
            RsDev("Value").value = val(Me.TxtPercentageValue.text)
            RsDev("Credit_Or_Debit").value = 1
            RsDev("Double_Entry_Vouchers_Description").value = XPMTxtRemarks.text & CHR(13) & "Œ’„ „”„ÊÕ »Â"
            RsDev("Notes_ID").value = val(XPTxtID.text)
            RsDev("RecordDate").value = Me.XPDtbTrans.value
            RsDev("RecordDateH").value = ToHijriDate(XPDtbTrans.value)
            RsDev("UserID").value = Me.DCboUserName.BoundText
            RsDev("Account_Interval_ID").value = SystemOptions.SysCurrentAccountIntervalID
             If Posted = 1 Then
            RsDev("Posted").value = 1
            Else
            RsDev("Posted").value = Null
            End If
            RsDev.update
      End If
      End If
End If
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
                
'If pstate = 1 Then Option7.value = True: GoTo LL

                If account_codeLegal = "" Or account_codeREVENUE_account = "" Then GoTo ll
            lineno = lineno + 1
            If CboStatus.ListIndex = 0 Then
                RsDev.AddNew
                RsDev("branch_id").value = val(Me.dcBranch.BoundText)
                RsDev("Double_Entry_Vouchers_ID").value = LngDevID
                RsDev("DEV_ID_Line_No").value = 3
                RsDev("DEV_ID_Line_No1").value = Line3
                RsDev("DEV_ID_Line_No").value = lineno
            If Posted = 1 Then
                RsDev("Posted").value = 1
            Else
                RsDev("Posted").value = Null
            End If
                RsDev("Account_Code").value = account_codeLegal
                RsDev("NextAccount_Code").value = DcboCreditSide.BoundText
                RsDev("Value").value = val(Me.XPTxtVal.text) - val(txtPayTotalProjVat)
                RsDev("Credit_Or_Debit").value = 0
                RsDev("Double_Entry_Vouchers_Description").value = XPMTxtRemarks.text & CHR(13) & lblinvoices.Caption
                'RsDev("Double_Entry_Vouchers_Description").value = dcproject.BoundText
            
                RsDev("Notes_ID").value = val(XPTxtID.text)
                RsDev("RecordDate").value = Me.XPDtbTrans.value
                RsDev("RecordDateH").value = ToHijriDate(XPDtbTrans.value)
                RsDev("UserID").value = Me.DCboUserName.BoundText
                RsDev("Account_Interval_ID").value = SystemOptions.SysCurrentAccountIntervalID

                If DCboCashType.ListIndex = 5 Then
 '                   RsDev("project_id").value = IIf(DBCboClientName.BoundText = "", 0, DBCboClientName.BoundText)
                End If

                RsDev.update
                '«·ÿ—› «·œ«∆‰
                RsDev.AddNew
                lineno = lineno + 1
                RsDev("branch_id").value = val(Me.dcBranch.BoundText)
                RsDev("Double_Entry_Vouchers_ID").value = LngDevID
                RsDev("DEV_ID_Line_No").value = lineno
                RsDev("DEV_ID_Line_No1").value = Line4
                RsDev("Account_Code").value = account_codeREVENUE_account
                RsDev("NextAccount_Code").value = DcboDebitSide.BoundText
                RsDev("Value").value = val(Me.XPTxtVal.text) - val(txtPayTotalProjVat)
                RsDev("Credit_Or_Debit").value = 1
                RsDev("Double_Entry_Vouchers_Description").value = XPMTxtRemarks.text & CHR(13) & lblinvoices.Caption
                ' RsDev("Double_Entry_Vouchers_Description").value = dcproject.BoundText
                RsDev("Notes_ID").value = val(XPTxtID.text)
                RsDev("RecordDate").value = Me.XPDtbTrans.value
                RsDev("RecordDateH").value = ToHijriDate(XPDtbTrans.value)
                RsDev("UserID").value = Me.DCboUserName.BoundText
                RsDev("Account_Interval_ID").value = SystemOptions.SysCurrentAccountIntervalID
          If Posted = 1 Then
                RsDev("Posted").value = 1
            Else
                RsDev("Posted").value = Null
          End If
                If DCboCashType.ListIndex = 5 Then
                   '  RsDev("project_id").value = IIf(DBCboClientName.BoundText = "", 0, DBCboClientName.BoundText)
                End If
    
                RsDev.update
                End If
ll:
            End If

            LblDevID.Caption = LngDevID
            lbl(33).Caption = SystemOptions.SysCurrentAccountIntervalID
        End If
If val(TxtVATValue.text) > 0 Then
lineno = lineno + 1
Line1 = Line1 + 1
            RsDev.AddNew
            RsDev("branch_id").value = val(Me.dcBranch.BoundText)
            RsDev("Double_Entry_Vouchers_ID").value = LngDevID
            RsDev("DEV_ID_Line_No").value = lineno
            RsDev("DEV_ID_Line_No1").value = Line1
            GetValueAddedAccount XPDtbTrans.value, , AccountVATCreit, 1, 23
            RsDev("Account_Code").value = AccountVATCreit
            RsDev("NextAccount_Code").value = DcboDebitSide.BoundText
            RsDev("Value").value = val(TxtVATValue.text)
            RsDev("Credit_Or_Debit").value = 1
            RsDev("Double_Entry_Vouchers_Description").value = "  «·ﬁÌ„… «·„÷«›… ··„⁄«„·«  «·„«·Ì…" & XPMTxtRemarks.text
            RsDev("Notes_ID").value = val(XPTxtID.text)
            RsDev("RecordDate").value = Me.XPDtbTrans.value
            RsDev("RecordDateH").value = ToHijriDate(XPDtbTrans.value)
            RsDev("UserID").value = Me.DCboUserName.BoundText
            RsDev("Account_Interval_ID").value = SystemOptions.SysCurrentAccountIntervalID
            If Posted = 1 Then
            RsDev("Posted").value = 1
            Else
            RsDev("Posted").value = Null
            End If
            RsDev.update
            
            lineno = lineno + 1
            End If



Dim total_value As Double
Dim LastLine
       If BranchID <> BranchID2 And BranchID2 <> 0 And SystemOptions.DontDistributeLegalACC = False Then
 total_value = val(Me.XPTxtVal.text) ' + val(Me.txtTransferExpenses.Text)
LastLine = lineno + 1
lineno = lineno + 1
OtherInformation.NextAccount_Code = DeptSide
Msg = "Ã«—Ï «·›—⁄ "
                                               If ModAccounts.AddNewDev(LngDevID, lineno, DeptSide, total_value, 0, Msg, val(XPTxtID.text), , , , XPDtbTrans.value, user_id, , , , , , , , , CDbl(LastLine), , , , , , , , , BranchID2, , , , , , , , , , , , , , , , , , , , , , , , , , , OtherInformation) = False Then
                                                                   
                                                              End If
                                                              'LastLine = LastLine + 1
                                                              LastLine = LastLine + 1
                                                        lineno = lineno + 1
                                                              OtherInformation.NextAccount_Code = credit_side
                                                        '????
                                                              If ModAccounts.AddNewDev(LngDevID, lineno, credit_side, total_value, 1, Msg, val(XPTxtID.text), , , , XPDtbTrans.value, user_id, , , , , , , , , CDbl(LastLine), , , , , , , , , BranchID, , , , , , , , , , , , , , , , , , , , , , , , , , , OtherInformation) = False Then
                                                                   
                                                              End If
                                                              
                                                        
                                    
                                                        
                                      LastLine = LastLine + 1
        

       'End If
       End If
       
       
       If BranchID <> BranchID3 And BranchID3 <> 0 And SystemOptions.DontDistributeLegalACC = False Then
 total_value = val(Me.XPTxtVal.text) ' + val(Me.txtTransferExpenses.Text)
'LastLine = 6
LastLine = LastLine + 1
OtherInformation.NextAccount_Code = DeptSide
Msg = "Ã«—Ï «·›—⁄ "
LastLine = LastLine + 1
lineno = lineno + 1
                                               If ModAccounts.AddNewDev(LngDevID, lineno, DeptSide, total_value, 0, Msg, val(XPTxtID.text), , , , XPDtbTrans.value, user_id, , , , , , , , , CDbl(LastLine), , , , , , , , , BranchID3, , , , , , , , , , , , , , , , , , , , , , , , , , , OtherInformation) = False Then
                                                                   
                                                              End If
                                                              'LastLine = LastLine + 1
                                                              LastLine = LastLine + 1
                                                              OtherInformation.NextAccount_Code = credit_side
                                                        '????
                                                        lineno = lineno + 1
                                                              If ModAccounts.AddNewDev(LngDevID, lineno, credit_side3, total_value, 1, Msg, val(XPTxtID.text), , , , XPDtbTrans.value, user_id, , , , , , , , , CDbl(LastLine), , , , , , , , , BranchID, , , , , , , , , , , , , , , , , , , , , , , , , , , OtherInformation) = False Then
                                                                   
                                                              End If
                                                              
                                                        
                                    
                                                        
                                      LastLine = LastLine + 1
        

       'End If
       End If
       
endSave:
        '==========================================================================
        Cn.CommitTrans
        BeginTrans = False
        XPTxtCurrent.Caption = rs.AbsolutePosition
        XPTxtCount.Caption = rs.RecordCount

        'Õ›Ÿ «·„” Œ·’« 
        If DCboCashType.ListIndex = 5 Or DCboCashType.ListIndex = 11 Then
          saveprojectBillPayment TxtNoteSerial.text, val(XPTxtVal.text), val(Me.XPTxtID.text)
  
        End If
    
        If DCboCashType.ListIndex = 5 Or DCboCashType.ListIndex = 11 Then
            FillGridWithData val(Me.DBCboClientName.BoundText), TxtNoteSerial.text
        End If
    
    
    
       'Õ›Ÿ «·«ﬁ”«ÿ ·⁄ﬁÊœ «·«ÌÃ«—
        If DCboCashType.ListIndex = 8 Then
'             saveContractInstallments val(Me.XPTxtID), XPDtbTrans.value, Txt_DateHigri.value, val(XPTxtVal.text), val(TXTContNo.text)
'
        End If
    
      '  If DCboCashType.ListIndex = 8 Then
      '         FillGridWithDataContract TxtContractNo.Text
      '  End If
        
        
        If Me.ChkTrans.value = vbUnchecked Then
            Me.CboTrans.ListIndex = -1
            Me.TxtTransSerial.text = ""
            Me.TxtTransID.text = ""
        End If
    saveBillBuy
        CuurentLogdata

'''' **************save Paydes***********************
         
   Dim PayDes As String
 Dim RowNum As Double
      PayDes = ""
    For RowNum = 1 To Grid22.rows - 1
            
                       If val(Grid22.TextMatrix(RowNum, Grid22.ColIndex("Value"))) <> 0 Then
                        
                                    'Check Repeat Serial
                                     
If PayDes <> "" Then
          PayDes = PayDes & CHR(13) & Grid22.TextMatrix(RowNum, Grid22.ColIndex("PaymentName")) & ":" & Grid22.TextMatrix(RowNum, Grid22.ColIndex("value"))
 Else
           PayDes = Grid22.TextMatrix(RowNum, Grid22.ColIndex("PaymentName")) & ":" & Grid22.TextMatrix(RowNum, Grid22.ColIndex("value"))
 End If
 End If
 Next RowNum
  
 Cn.Execute "update Notes set PayDes ='" & PayDes & "'   where NoteID=" & val(XPTxtID.text)
 '''' **************save Paydes***********************


        Select Case Me.TxtModFlg.text

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
    
        '   If Me.DcCostCenter.BoundText <> "" Then
        save_General_cost_center Me.DcCostCenter.BoundText, Me.DcCostCenter.text, "„ﬁ»Ê÷« ", Me.XPDtbTrans.value
        save_cost_center
        '   End If
        
        'Õ›Ÿ «·„’«—Ì› › ÃœÊ· «·„œ›Ê⁄«  Ê «·„ﬁ»Ê÷« 
     
        If SavePaymentAndReciveDetails(1, TxtNoteSerial.text, TxtNoteSerial1.text, "", XPDtbTrans.value) = True Then
        End If

        TxtModFlg.text = "R"
    End If

    WriteCustomerBalPublic Me.DcboCreditSide.BoundText, Balance, balanceString
    LblLink.Caption = balanceString
    WriteInfo
 RetriveBillBuyData
 '   If Option1.value = True And DCboCashType.ListIndex <> 8 Then
 ' If SystemOptions.EnableCustomerAging = True Then
      
      
  '      FIFO_FUNCTION val(DBCboClientName.BoundText)
 ' End If
  
   ' End If
   
    If Option2.value Then
      '  Distribute_to_bills Me.lblsqlstring, val(DBCboClientName.BoundText)
    End If
   
    TxtModFlg.text = "R"
    fillapprovData
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
Function CheckMult_Cash() As Boolean
Dim i As Integer
  With Grid22
      For i = 1 To .rows - 1
      If val(.TextMatrix(i, .ColIndex("Value"))) <> 0 And val(.TextMatrix(i, .ColIndex("PaymentID"))) = 0 Then
      CheckMult_Cash = True
      Exit Function
      End If
      Next i
      CheckMult_Cash = False
   End With
End Function
Sub PGMultyPayment(Optional general_noteid As Long, Optional ByRef lineno As Integer, Optional ByRef Line1 As Double, Optional StrTempDes As String, Optional Posted As Integer)
Dim StrMSG As String
Dim Commisionvalue As Double
Dim StrTempAccountCode As String
Dim i As Integer
Dim ValuGird As Double
Dim maxvalue As Double
Dim commision As Double
   Dim LngDevID As Long
    'Dim LngDevNO  As Integer
  '  Dim StrTempDes As String
    Dim SngTemp As Variant
    Dim TotalValue As Double
    On Error GoTo ErrTrap
    LngDevID = new_id("DOUBLE_ENTREY_VOUCHERS", "Double_Entry_Vouchers_ID", "")
  With Grid22
      For i = 1 To .rows - 1
      StrMSG = ""
      If val(.TextMatrix(i, .ColIndex("Value"))) <> 0 Then
      ValuGird = val(.TextMatrix(i, .ColIndex("Value"))) '* val(txt_Currency_rate.Text)
      StrMSG = " " & (.TextMatrix(i, .ColIndex("PaymentName")))
      If val(.TextMatrix(i, .ColIndex("PaymentID"))) = 0 Then
      StrTempAccountCode = GetMyAccountCode("TblBoxesData", "BoxID", val(Me.DcboBox.BoundText))
      OtherInformation.NextAccount_Code = DcboCreditSide.BoundText
            If ModAccounts.AddNewDev(LngDevID, lineno, StrTempAccountCode, ValuGird, 0, StrTempDes & StrMSG, general_noteid, , , , Me.XPDtbTrans.value, DCboUserName.BoundText, , , , , , , , , , , , , , , , , , val(Me.dcBranch.BoundText), , , , , , , , , , , , , , , , , , , , , , , , , Posted, , OtherInformation) = False Then
                GoTo ErrTrap
                End If
            lineno = lineno + 1
      ElseIf val(.TextMatrix(i, .ColIndex("PaymentID"))) > 0 Then
      commision = val(.TextMatrix(i, .ColIndex("commision")))
      maxvalue = val(get_TblPaymentTypet(val(.TextMatrix(i, .ColIndex("PaymentID"))), "MaxValue"))
          StrTempAccountCode = .TextMatrix(i, .ColIndex("bankAccount_Code"))
      If SystemOptions.AllowCommtionJEFromValueVisa = True Then

            
      If commision > 0 And .TextMatrix(i, .ColIndex("Accountcom")) <> "" Then
                Commisionvalue = (ValuGird * commision) / 100
                If maxvalue <> 0 And maxvalue < Commisionvalue Then
                Commisionvalue = maxvalue
                End If
        
            If ModAccounts.AddNewDev(LngDevID, lineno, .TextMatrix(i, .ColIndex("Accountcom")), Commisionvalue, 0, StrTempDes & "   " & .TextMatrix(i, .ColIndex("PaymentName")) & "⁄„Ê·… ", general_noteid, , , , Me.XPDtbTrans.value, Me.DCboUserName.BoundText, , , , , , , , , , , , , , , , , val(Me.dcBranch.BoundText), , , , , , , , , , , , , , , , , , , , , , , , , , Posted, , OtherInformation) = False Then
                GoTo ErrTrap
                End If
            lineno = lineno + 1
      End If
      ValuGird = ValuGird - Commisionvalue
                If ModAccounts.AddNewDev(LngDevID, lineno, StrTempAccountCode, ValuGird, 0, StrTempDes & StrMSG, general_noteid, , , , Me.XPDtbTrans.value, Me.DCboUserName.BoundText, val(Me.XPTxtBillID.text), , , , , , , , , , , , , , , , , val(Me.dcBranch.BoundText), , , , , , , , , , , , , , , , , , , , , , , , , Posted, , OtherInformation) = False Then
                GoTo ErrTrap
                End If
            lineno = lineno + 1
      Else
                   If ModAccounts.AddNewDev(LngDevID, lineno, StrTempAccountCode, ValuGird, 0, StrTempDes & StrMSG, general_noteid, , , , Me.XPDtbTrans.value, Me.DCboUserName.BoundText, val(Me.XPTxtBillID.text), , , , , , , , , , , , , , , , , val(Me.dcBranch.BoundText), , , , , , , , , , , , , , , , , , , , , , , , , Posted, , OtherInformation) = False Then
                GoTo ErrTrap
                End If
            lineno = lineno + 1
            
      If commision > 0 And .TextMatrix(i, .ColIndex("Accountcom")) <> "" Then
                Commisionvalue = (ValuGird * commision) / 100
                If maxvalue <> 0 And maxvalue < Commisionvalue Then
                Commisionvalue = maxvalue
                End If
        
            If ModAccounts.AddNewDev(LngDevID, lineno, .TextMatrix(i, .ColIndex("Accountcom")), Commisionvalue, 0, StrTempDes & "   " & .TextMatrix(i, .ColIndex("PaymentName")) & "⁄„Ê·… ", general_noteid, , , , Me.XPDtbTrans.value, Me.DCboUserName.BoundText, val(Me.XPTxtBillID.text), , , , , , , , , , , , , , , , , val(Me.dcBranch.BoundText), , , , , , , , , , , , , , , , , , , , , , , , , Posted, , OtherInformation) = False Then
                GoTo ErrTrap
                End If
            lineno = lineno + 1
            OtherInformation.NextAccount_Code = DcboDebitSide.BoundColumn
                 If ModAccounts.AddNewDev(LngDevID, lineno, StrTempAccountCode, Commisionvalue, 1, StrTempDes & "   " & .TextMatrix(i, .ColIndex("PaymentName")) & "⁄„Ê·… ", general_noteid, , , , Me.XPDtbTrans.value, Me.DCboUserName.BoundText, val(Me.XPTxtBillID.text), , , , , , , , , , , , , , , , , val(Me.dcBranch.BoundText), , , , , , , , , , , , , , , , , , , , , , , , , Posted, , OtherInformation) = False Then
                GoTo ErrTrap
                End If
            lineno = lineno + 1
      End If
      End If
   
      End If
          
          End If
     Next i
      End With
ErrTrap:
End Sub
Function saveChequeBoxContents(NoteID As Double)

    Dim i As Integer
    Dim rs As New ADODB.Recordset
 
    Dim StrSQL As String

    StrSQL = "Delete  TblChecqueBoxContent  where NoteID =" & NoteID
    Cn.Execute StrSQL, , adExecuteNoRecords

    
    If SystemOptions.IsCheque = True And CboPayMentType.ListIndex = 1 Then
    
    Else
        If val(DCChequeBox.BoundText) = 0 Then Exit Function
    End If
 
  '  rs.Open "TblChecqueBoxContent", Cn, adOpenStatic, adLockOptimistic, adCmdTable
    StrSQL = "SELECT     * from dbo.TblChecqueBoxContent Where (1 = -1)"
   rs.Open StrSQL, Cn, adOpenKeyset, adLockOptimistic, adCmdText
   
   
    rs.AddNew
    rs("noteid").value = NoteID
    rs("ChequeBoxID").value = val(DCChequeBox.BoundText)
            
    rs("RecordDate").value = XPDtbTrans.value
    rs("DueDate").value = DtpChequeDueDate.value
    rs("BankName").value = TXTBankName.text
     rs("BankId").value = val(DcboBankName.BoundText)
    rs("ChequeNo").value = TxtChequeNumber.text
    rs("ChequeValue").value = val(XPTxtVal.text) + val(txtVat2.text)
    
    rs("Remarks").value = DcboCreditSide.text
    rs("Deposited").value = 0
    rs("Collected").value = 0
    rs("CreditAccount").value = (DcboCreditSide.BoundText)
    
    
    
            If DCboCashType.ListIndex = 0 Then
                        rs("customeraccount").value = ModAccounts.GetMyAccountCode("TblCustemers", "CusID", val(Me.DBCboClientName.BoundText), "Account_code")
                        rs("customeraccount1").value = ModAccounts.GetMyAccountCode("TblCustemers", "CusID", val(Me.DBCboClientName.BoundText), "Account_code1")
                        rs("customeraccount2").value = ModAccounts.GetMyAccountCode("TblCustemers", "CusID", val(Me.DBCboClientName.BoundText), "Account_code2")
                        
             ElseIf DCboCashType.ListIndex = 5 Then
                       rs("customeraccount").value = get_project_customer_account(val(DBCboClientName.BoundText), "Account_Code")
                        rs("customeraccount1").value = get_project_customer_account(val(DBCboClientName.BoundText), "Account_Code1")
                        rs("customeraccount2").value = get_project_customer_account(val(DBCboClientName.BoundText), "Account_Code2")
                        
              
              
            End If
    
    rs.update
  
    rs.Close
End Function

Function save_cost_center()

    'on error resume next
    If Not IsNumeric(Text1.text) Then Exit Function
    Dim i As Integer
    Dim rs As ADODB.Recordset
    Set rs = New ADODB.Recordset
    Dim sql_str As String

    'Rs.Open "", Cn, adOpenStatic, adLockOptimistic, adCmdTable

    sql_str = "select * from marakes_taklefa_temp where kedno=" & Text1.text
    rs.Open sql_str, Cn, adOpenStatic, adLockOptimistic, adCmdText

    For i = 1 To rs.RecordCount
        rs("ok").value = 1
        rs("NoteDate").value = XPDtbTrans.value
        rs("NoteSerial").value = TxtNoteSerial.text
        rs("Remark").value = "”‰œ „ﬁ»Ê÷«     —ﬁ„ " & TxtNoteSerial1.text & "    " & Me.TxtCustCode
 
        rs.update
        rs.MoveNext
    Next i

End Function

Public Function save_General_cost_center(cost_center_id As String, _
                                         cost_center, _
                                         opr_type As String, _
                                         record_date As Date) 'As String, value As Double, depit_or_credit As Boolean, opr_id As Double, opr_type As String, account_no As String, account_name As String, line_no As Double, recorddate As Date)
    Dim i As Integer
    Dim rs As New ADODB.Recordset
 
    Dim StrSQL As String

    StrSQL = "Delete  marakes_taklefa_temp  where general_des=1 AND  kedno =" & val(Text1.text)
    Cn.Execute StrSQL, , adExecuteNoRecords
    
    If Me.DcCostCenter.BoundText = "" Then
        Exit Function
    End If
 
    'rs.Open "marakes_taklefa_temp", Cn, adOpenStatic, adLockOptimistic, adCmdTable
  StrSQL = "SELECT   *  from dbo.marakes_taklefa_temp Where (1 = -1)"
   rs.Open StrSQL, Cn, adOpenKeyset, adLockOptimistic, adCmdText
 
    'ÿ—› „œÌ‰
    '       rs.AddNew
    '       rs("cost_center_id").value = cost_center_id
    '       rs("cost_center").value = cost_center
    '       rs("value").value = XPTxtVal.text
    '       rs("depit_or_credit").value = "„œÌ‰"
    '       rs("opr_id").value = Me.Text1.text
    '       rs("kedno").value = Me.Text1.text
    '
    '       rs("opr_type").value = opr_type
    '       rs("account_name").value = DcboDebitSide.text
    '       rs("account_no").value = DcboDebitSide.BoundText
    '       rs("line_no").value = Line1
    '       rs("record_date").value = record_date
    '       rs.update
    'ÿ—› œ«∆‰
    rs.AddNew
    rs("general_des").value = 1
    rs("cost_center_id").value = cost_center_id
    rs("cost_center").value = cost_center
    rs("value").value = XPTxtVal.text
    rs("depit_or_credit").value = "œ«∆‰"
    rs("opr_id").value = Me.Text1.text
    rs("kedno").value = Me.Text1.text

    rs("opr_type").value = opr_type
    rs("account_name").value = DcboCreditSide.text
    rs("account_no").value = DcboCreditSide.BoundText
    rs("line_no").value = Line2
    rs("record_date").value = record_date
                    rs("description").value = XPMTxtRemarks.text
                    
    rs.update
 
    rs.Close
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
    total_value = val(XPTxtVal.text)
  
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

    txtAdv_payment_value.text = total_value
    change_adv_payment_value XPTxtID.text, total_value

    ' If IsNull(Rs3("UserName").value) Then FIFO_FUNCTION = "": Exit Function
  
    ' If Not IsNull(Rs3("UserName").value) Then get_user_name = Rs3("UserName").value: Exit Function
    Rs3.Close
 
End Function

Function FIFO_FUNCTION(CusID As Double)
    Dim Rs3 As ADODB.Recordset
    Set Rs3 = New ADODB.Recordset
    Dim sql As String
    Dim i As Integer
 If CusID = 0 Then Exit Function
   sql = " delete   notes where NoteType= 2000   and  NoteSerial='" & TxtNoteSerial.text & "'"
 'Cn.Execute sql
Cn.Execute sql


    sql = "SELECT CompanyCreditValues.*  FROM dbo.CompanyCreditValues() CompanyCreditValues  where   (cusid=" & CusID & " and requiredvalue>0  AND TRANSACTION_TYPE=21 )  order by duedate"
 
    Rs3.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
  
    If Rs3.RecordCount = 0 Then Exit Function
    Dim total_value As Double
    Dim current_value As Double
    total_value = val(Me.XPTxtVal.text)
  
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
    txtAdv_payment_value.text = total_value
  '  change_adv_payment_value XPTxtID.text, total_value
    ' If Not IsNull(Rs3("UserName").value) Then get_user_name = Rs3("UserName").value: Exit Function
    Rs3.Close

End Function

Function Add_new_notes(NoteDate As Date, NoteType As Integer, Note_Value As Double, Transaction_ID As Double, CusID As Double, BoxID As Integer, displayed As Integer, UserID As Integer)
    Dim RsDev As New ADODB.Recordset
   ' RsDev.Open "notes", Cn, adOpenStatic, adLockOptimistic, adCmdTable
      Dim StrSQL  As String
       StrSQL = "SELECT     dbo.Notes.* from dbo.Notes Where (NoteID = -1)"
   RsDev.Open StrSQL, Cn, adOpenKeyset, adLockOptimistic, adCmdText
   
   
   
    '
    Dim sql As String
    

    RsDev.AddNew
      
    RsDev("NoteID").value = CStr(new_id("Notes", "NoteID", "", True))
    RsDev("NoteSerial").value = TxtNoteSerial.text ' CStr(new_id("Notes", "NoteSerial", "", True, "NoteType=2000"))
              
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

    Select Case TxtModFlg.text

        Case "N"
            clear_all Me
            Me.TxtModFlg.text = "R"
            XPBtnMove_Click (1)

        Case "E"
            rs.Find "NoteID='" & val(XPTxtID.text) & "'", , adSearchForward, adBookmarkFirst

            If rs.EOF Or rs.BOF Then
                Me.TxtModFlg.text = "R"
                Exit Sub
            End If

            Retrive
            Me.TxtModFlg.text = "R"
    End Select

    Exit Sub
ErrTrap:
End Sub

Private Sub Del_Trans()
    Dim Msg As String
    On Error GoTo ErrTrap

  If XPTxtID.text <> "" Then
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
        Msg = Msg + (TxtNoteSerial.text) & CHR(13)
        Msg = Msg + " Â·  —€» ›Ì Õ–› Â–Â «·»Ì«‰« ø"

        If MsgBox(Msg, vbYesNo + vbQuestion + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title) = vbYes Then
            If Not rs.RecordCount < 1 Then
                CuurentLogdata ("D")
    Deletepost Me.Name, "Notes", "NoteID", 0, val(dcBranch.BoundText), val(XPTxtID.text), TxtNoteSerial1.text
    
                rs.delete
                Dim StrSQL As String
               ' StrSQL = "Delete From notes  Where  (NoteType=2000 OR NoteType=4 ) AND  NoteSerial=" & val(TxtNoteSerial.Text)
               ' Cn.Execute StrSQL, , adExecuteNoRecords
        
                StrSQL = "Delete  marakes_taklefa_temp  where kedno =" & val(Text1.text)
                Cn.Execute StrSQL, , adExecuteNoRecords
       
                StrSQL = "Delete From ReciveDetails Where NoteSerial1='" & val(TxtNoteSerial1.text) & "'"
                Cn.Execute StrSQL, , adExecuteNoRecords
    
                StrSQL = "Delete From ProjectBillBuy Where TxtNoteSerial='" & TxtNoteSerial.text & "'"
                Cn.Execute StrSQL, , adExecuteNoRecords
    
    
                StrSQL = "Delete From ContracttBillInstallmentsDone Where NoteID =" & val(Me.XPTxtID)
                Cn.Execute StrSQL, , adExecuteNoRecords
                 StrSQL = "Delete From TblMultuPayment Where NoteID =" & val(Me.XPTxtID)
                Cn.Execute StrSQL, , adExecuteNoRecords
                
                StrSQL = "Delete  TblChecqueBoxContent  where NoteID =" & val(Me.XPTxtID)
                Cn.Execute StrSQL, , adExecuteNoRecords
    
    DeleteBillBuy
              StrSQL = "Delete From TblNotesBillBuyPayment2 Where NoteID1=" & val(Me.XPTxtID.text) & " and TransType is null"
              Cn.Execute StrSQL, , adExecuteNoRecords
              StrSQL = "Delete From TblBillBuyPayment2 Where TypTrans IS NULL and  NoteID=" & val(Me.XPTxtID.text)
              Cn.Execute StrSQL, , adExecuteNoRecords
     StrSQL = " delete   notes where NoteType= 2000   and  NoteSerial='" & TxtNoteSerial.text & "'"
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
    
 TranslateForm Me, True
    lbl(43).Caption = "Cheque Box"
    lbl(50).Caption = "Car"
    lbl(49).Caption = "Driver"
    ALLButton6.Caption = "Show"
Option7.Caption = "Old Projects"
lbl(48).Caption = "Manual No."
Command1.Caption = "Show All"
lbl(66).Caption = "Total"
CmdAttach.Caption = "Attachments"
lbl(64).Caption = "Account"
CMDPAy.Caption = "Pay"
lbl(65).Caption = "VAT"
lbl(101).Caption = "Total"
lbl(67).Caption = "Bill No."
    lbl(100).Caption = "Paid"
    lbl(99).Caption = "Remaining"
lbl(51).Caption = "Book No."
lbl(56).Caption = "Comm. Dis."
lbl(57).Caption = "Comm. Acc."
Command9.Caption = "Show Bills"
Command10.Caption = "Cancel Payment"
Check1.RightToLeft = False
Check1.Caption = "Select All"
Frame12.Caption = "Data"
Label27.Caption = "Total"
FramePay.Caption = "Payments Data"
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
    Cmd(10).Caption = "Print 2"
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
    Me.XPTab301.TabCaption(2) = "Payments"
    Me.XPTab301.TabCaption(3) = "Approval Status"
    lbl(37).Caption = "Total Rec."""
    lbl(0).Caption = "Select bills"
    lbl(42).Caption = "Payed bills"
    CmdRemove.Caption = "Remove Row"

    lbl(58).Caption = "Status"
    
  With CboStatus
  .Clear
  .AddItem "Done"
  .AddItem "Pending"
  .AddItem "Cancelled"
  .AddItem "Lost"
  
  End With
  
    With Grid
        .TextMatrix(0, .ColIndex("NoteSerial1")) = "Progress Bill"
        .TextMatrix(0, .ColIndex("ManualNO")) = "ManualNO"
        .TextMatrix(0, .ColIndex("Ser")) = "Ser"
        .TextMatrix(0, .ColIndex("Select")) = "Select"
        .TextMatrix(0, .ColIndex("id")) = "Invoice No."
        .TextMatrix(0, .ColIndex("bill_date")) = "Invoice Date"
        .TextMatrix(0, .ColIndex("total")) = "Invoice Total"
        .TextMatrix(0, .ColIndex("ActualTotal")) = "Payed Totalt"
        .TextMatrix(0, .ColIndex("result")) = "Not Payed"
        .TextMatrix(0, .ColIndex("resultpercentage")) = "Not Payed%"
    End With
    With Grid22
    .TextMatrix(0, .ColIndex("PaymentName")) = "Payments"
    .TextMatrix(0, .ColIndex("Value")) = "Value"
    .TextMatrix(0, .ColIndex("CardNo")) = "Card No."
    End With
    With GRID1
    
        .TextMatrix(0, .ColIndex("NoteSerial1")) = "Progress Bill"
        .TextMatrix(0, .ColIndex("ManualNO")) = "ManualNO"
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
    
    With Grid2
        .TextMatrix(0, .ColIndex("Approved")) = "Approved"
        .TextMatrix(0, .ColIndex("levelName")) = "Level"
        .TextMatrix(0, .ColIndex("EmpName")) = "Employee"
        .TextMatrix(0, .ColIndex("ApprovDate")) = "Approve Date"
    End With
        
        Label1100.Caption = "Approval Requested by "
        Label24.Caption = "Approval Requested by "
        
    With Grid3
        .TextMatrix(0, .ColIndex("Ser")) = "No."
        .TextMatrix(0, .ColIndex("Select")) = "Select"
        .TextMatrix(0, .ColIndex("InstallNo")) = "Installment No"
        .TextMatrix(0, .ColIndex("Installdateh")) = "Hijri Date"
        .TextMatrix(0, .ColIndex("Installdate")) = "Date"
        .TextMatrix(0, .ColIndex("CommisionTypes")) = "Type"
        .TextMatrix(0, .ColIndex("RentValue")) = "Rent"
        .TextMatrix(0, .ColIndex("Insurance")) = "percentage "
        .TextMatrix(0, .ColIndex("Commissions")) = "Insurance"
        .TextMatrix(0, .ColIndex("Water")) = "Water"
        .TextMatrix(0, .ColIndex("Electric")) = "Electric"
        .TextMatrix(0, .ColIndex("TelandNet")) = "Services"
        .TextMatrix(0, .ColIndex("OldValue")) = "Remainder"
        .TextMatrix(0, .ColIndex("total")) = "Total"
        .TextMatrix(0, .ColIndex("RentValuePayed")) = "Payed Rent"
        .TextMatrix(0, .ColIndex("Installdate")) = "Date"
        .TextMatrix(0, .ColIndex("CommissionsPayed")) = "Payed Commissions"
        .TextMatrix(0, .ColIndex("InsurancePayed")) = "Payed Insurance"
        .TextMatrix(0, .ColIndex("WaterPayed")) = "Payed Water"
        .TextMatrix(0, .ColIndex("ElectricPayed")) = "Payed Electric"
        .TextMatrix(0, .ColIndex("TelandNetPayed")) = "Payed Services"
        .TextMatrix(0, .ColIndex("ActualTotal")) = "Payed"
        .TextMatrix(0, .ColIndex("Result")) = "Remainder"
        .TextMatrix(0, .ColIndex("ResultPercentage")) = "percentage "
    End With
    
        With Grid4
        .TextMatrix(0, .ColIndex("Ser")) = "No."
        .TextMatrix(0, .ColIndex("Select")) = "Select"
        .TextMatrix(0, .ColIndex("InstallNo")) = "Installment No"
        .TextMatrix(0, .ColIndex("Installdateh")) = "Hijri Date"
        .TextMatrix(0, .ColIndex("Installdate")) = "Date"
        
        .TextMatrix(0, .ColIndex("total")) = "Installment Value"
        .TextMatrix(0, .ColIndex("ActualTotal")) = "Total Payed"
        .TextMatrix(0, .ColIndex("RentValuePayed")) = "Payed Rent "
        .TextMatrix(0, .ColIndex("CommissionsPayed")) = "payed Commissions"
        .TextMatrix(0, .ColIndex("InsurancePayed")) = "Payed Insurance"
        .TextMatrix(0, .ColIndex("WaterPayed")) = "Payed Water"
        .TextMatrix(0, .ColIndex("ElectricPayed")) = "Payed Electric"
        .TextMatrix(0, .ColIndex("TelandNetPayed")) = "Payed Services"
        .TextMatrix(0, .ColIndex("Result")) = "Remainder"
        .TextMatrix(0, .ColIndex("ResultPercentage")) = "percentage "
        

        
        
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
    DCboCashType.AddItem "From Transportation "
    DCboCashType.AddItem "From  Contract"
    DCboCashType.AddItem "From  Bill Maintenance"
    DCboCashType.AddItem "Based on a maintenance card"
    DCboCashType.AddItem "Based on container contract "
    With Me.CboPayMentType
        .Clear
        .AddItem "Cash"
        .AddItem "Cheque"
        .AddItem "Bank Transfer"
        .AddItem "Coll. Cheque"
        .AddItem "Account"
    If SystemOptions.AllowAccountMultyPayed = True Then
    .AddItem "Multy"
    End If
    End With
    With Me.commdiscounttype
        .Clear
        .AddItem "NA"
        .AddItem "Value"
        .AddItem "Percemtage"
        
    End With
        With VSFlexGrid1

.TextMatrix(0, .ColIndex("Ser")) = "Serial"
.TextMatrix(0, .ColIndex("InstalValue")) = "Installment Value"
.TextMatrix(0, .ColIndex("haveqest")) = "Have Installments"
.TextMatrix(0, .ColIndex("payed")) = "Select"
.TextMatrix(0, .ColIndex("NoteSerial1")) = "Bill No"
.TextMatrix(0, .ColIndex("too")) = "Bill Supplier"
.TextMatrix(0, .ColIndex("NoteDate")) = "Date"
.TextMatrix(0, .ColIndex("branch_name")) = "Branch"
.TextMatrix(0, .ColIndex("Note_Value")) = "Original value"
.TextMatrix(0, .ColIndex("PayedValue")) = "Payed Value"
.TextMatrix(0, .ColIndex("RemainingValue")) = "Remaining"
.TextMatrix(0, .ColIndex("TransPayedValue")) = "Payed Trans"
.TextMatrix(0, .ColIndex("NetValue")) = "Net value"
.TextMatrix(0, .ColIndex("Show")) = "Show"
.TextMatrix(0, .ColIndex("DueDate")) = "Due Date"
End With
    With Me.CboTrans
        .Clear
        .AddItem "Sales invoice"
        .AddItem "Returned purchases"
        .AddItem "Delivery of maintenance for a client"
        .AddItem "Services"
    End With
 
Accredit.Caption = "Send for Approval"
 
End Sub
 Function saveBillBuy()
    Dim StrSQL As String
   ' Dim StrSQL  As String
    Dim i As Integer
    Dim Diff As Double
    Dim Note_Value1 As Double
    Diff = 0
Dim RsDetails As ADODB.Recordset
      If Me.TxtModFlg.text = "E" Then
    StrSQL = "Delete From TblNotesBillBuyPayment2 Where NoteID1=" & val(Me.XPTxtID.text) & " and TransType is null"
    Cn.Execute StrSQL, , adExecuteNoRecords
        StrSQL = "Delete From TblBillBuyPayment2 Where TypTrans IS NULL and  NoteID=" & val(Me.XPTxtID.text)
    Cn.Execute StrSQL, , adExecuteNoRecords
    End If

    Set RsDetails = New ADODB.Recordset
   ' RsDetails.Open "TblEmpAdvanceDetails", Cn, adOpenKeyset, adLockOptimistic, adCmdTable

    StrSQL = "SELECT     * from dbo.TblNotesBillBuyPayment2 Where (1 = -1)"
   RsDetails.Open StrSQL, Cn, adOpenKeyset, adLockOptimistic, adCmdText
    
    With VSFlexGrid1
    TxtValueTemp.text = val(XPTxtVal.text)
    For i = .FixedRows To .rows - 1
        If .cell(flexcpChecked, i, .ColIndex("payed")) = flexChecked Then
            RsDetails.AddNew
            RsDetails("NoteID1").value = val(XPTxtID.text)
            RsDetails("NoteID").value = val(.TextMatrix(i, .ColIndex("NoteID")))
            RsDetails("branch_no").value = val(.TextMatrix(i, .ColIndex("branch_no")))
            RsDetails("NoteSerial1").value = val(.TextMatrix(i, .ColIndex("NoteSerial1")))
            RsDetails("Note_Value").value = val(.TextMatrix(i, .ColIndex("Note_Value")))
            Note_Value1 = val(.TextMatrix(i, .ColIndex("RemainingValue")))
            Diff = 0
            If val(TxtValueTemp.text) > 0 Then
          If val(TxtValueTemp.text) <= Note_Value1 Then
          Diff = val(TxtValueTemp.text)
          TxtValueTemp.text = val(TxtValueTemp.text) - Note_Value1
          Else
          Diff = Note_Value1
          TxtValueTemp.text = val(TxtValueTemp.text) - Note_Value1
          End If
            End If
          ' .TextMatrix(i, .ColIndex("RemainingValue")) = val(.TextMatrix(i, .ColIndex("RemainingValue"))) - val(.TextMatrix(i, .ColIndex("RemainingValue")))
            .TextMatrix(i, .ColIndex("TransPayedValue")) = Diff
            
            RsDetails("PayedValue").value = val(.TextMatrix(i, .ColIndex("PayedValue")))
            
            RsDetails("too").value = (.TextMatrix(i, .ColIndex("too")))
            RsDetails("NoteDate").value = IIf((.TextMatrix(i, .ColIndex("NoteDate"))) = "", Null, (.TextMatrix(i, .ColIndex("NoteDate"))))
            If .TextMatrix(i, .ColIndex("DueDate")) <> "" And .TextMatrix(i, .ColIndex("DueDate")) <> " " Then
            RsDetails("DueDate").value = IIf((.TextMatrix(i, .ColIndex("DueDate"))) = "", Null, (.TextMatrix(i, .ColIndex("DueDate"))))
            Else
            RsDetails("DueDate").value = Null
            End If
            RsDetails("TransPayedValue").value = val(.TextMatrix(i, .ColIndex("TransPayedValue")))
           .TextMatrix(i, .ColIndex("NetValue")) = val(.TextMatrix(i, .ColIndex("RemainingValue"))) - val(.TextMatrix(i, .ColIndex("TransPayedValue")))
            RsDetails("NetValue").value = val(.TextMatrix(i, .ColIndex("NetValue")))
            RsDetails("RemainingValue").value = val(.TextMatrix(i, .ColIndex("RemainingValue")))
            RsDetails.update
                
            If val(val(.TextMatrix(i, .ColIndex("Transaction_Type")))) <> 9999 Then
                If val(val(.TextMatrix(i, .ColIndex("NetValue")))) = 0 Then
                    StrSQL = "Update Transactions Set  TotalPayed=1 Where Transaction_ID=" & val(.TextMatrix(i, .ColIndex("NoteID"))) & ""
                        Cn.Execute StrSQL, , adExecuteNoRecords
                     Else
                         StrSQL = "Update Transactions Set  TotalPayed=0 Where Transaction_ID=" & val(.TextMatrix(i, .ColIndex("NoteID"))) & ""
                        Cn.Execute StrSQL, , adExecuteNoRecords
                End If
            Else
                If val(val(.TextMatrix(i, .ColIndex("NetValue")))) = 0 Then
                    StrSQL = "Update TblTravDueK Set  TotalPayed=1 Where ID=" & val(.TextMatrix(i, .ColIndex("NoteID"))) & ""
                        Cn.Execute StrSQL, , adExecuteNoRecords
                     Else
                         StrSQL = "Update TblTravDueK Set  TotalPayed=0 Where ID=" & val(.TextMatrix(i, .ColIndex("NoteID"))) & ""
                        Cn.Execute StrSQL, , adExecuteNoRecords
                End If
            End If
      End If
    Next i
End With
    Set RsDetails = New ADODB.Recordset
   ' RsDetails.Open "TblEmpAdvanceDetails", Cn, adOpenKeyset, adLockOptimistic, adCmdTable

    StrSQL = "SELECT     * from dbo.TblBillBuyPayment2 Where (1 = -1)"
   RsDetails.Open StrSQL, Cn, adOpenKeyset, adLockOptimistic, adCmdText
    
    With VSFlexGrid1
    For i = .FixedRows To .rows - 1
        If .cell(flexcpChecked, i, .ColIndex("payed")) = flexChecked Then
            RsDetails.AddNew
            RsDetails("NoteID").value = val(XPTxtID.text)
            RsDetails("RecDate").value = XPDtbTrans.value
            RsDetails("Serial").value = TxtNoteSerial1.text
            RsDetails("Transaction_ID").value = val(.TextMatrix(i, .ColIndex("NoteID")))
            RsDetails("TransType").value = val(.TextMatrix(i, .ColIndex("Transaction_Type")))
            RsDetails("Note_Value").value = val(.TextMatrix(i, .ColIndex("Note_Value")))
            RsDetails("PayedValue").value = val(.TextMatrix(i, .ColIndex("TransPayedValue")))
            RsDetails.update
        End If
    Next i
End With

End Function
Sub RelineBuy()
    Dim IntCounter As Integer
    Dim Sm As Double
    Dim billDesStr As String
    billDesStr = "”œ«œ ›Ê« Ì— «—ﬁ«„"
    Sm = 0
    IntCounter = 0
    Dim i As Integer
    With Me.VSFlexGrid1
        For i = .FixedRows To .rows - 1
                If .cell(flexcpChecked, i, .ColIndex("payed")) = flexChecked Then
                IntCounter = IntCounter + 1
                .TextMatrix(i, .ColIndex("Ser")) = IntCounter
           Sm = Sm + val(.TextMatrix(i, .ColIndex("RemainingValue")))
           billDesStr = billDesStr & "," & (.TextMatrix(i, .ColIndex("NoteSerial1")))
           End If
           Next i
  
    End With
   Label28.Caption = Sm
   XPMTxtRemarks.text = billDesStr
End Sub
Function CuurentLogdata(Optional Currentmode As String)
    LogTextA = "    ‘«‘… " & ScreenNameArabic & CHR(13) & " —ﬁ„ «·”‰œ " & TxtNoteSerial1.text & CHR(13) & "   «· «—ÌŒ " & XPDtbTrans & CHR(13) & "   ‰Ê⁄ «·„ﬁ»Ê÷«  " & DCboCashType & CHR(13) & "   «·›—⁄  " & dcBranch & CHR(13) & "   «·«”„  " & DBCboClientName & CHR(13) & "   ﬁÌ„Â «·„ﬁ»Ê÷«   " & XPTxtVal & CHR(13) & "   ÿ—Ìﬁ… «·ﬁ»÷ " & CboPayMentType & CHR(13) & "   «·Œ“Ì‰…  " & DcboBox & CHR(13) & "   «·»‰ﬂ  " & DcboBankName & CHR(13) & "   —ﬁ„ «·‘Ìﬂ  " & TxtChequeNumber & CHR(13) & "    «—ÌŒ «·«” Õﬁ«ﬁ  " & DtpChequeDueDate & CHR(13) & "     »‰«¡ ⁄·Ï   " & XPMTxtRemarks & CHR(13) & "   —ﬁ„ «·ﬁÌœ   " & TxtNoteSerial & CHR(13) & "   —ﬁ„ «·ﬁÌœ   " & TxtNoteSerial & CHR(13) & "ÿ—› „œÌ‰  " & DcboDebitSide & CHR(13) & " ÿ—› œ«∆‰ " & DcboCreditSide & CHR(13) & " «·„‰œÊ» " & DcEmp
                        
    LogTexte = "    Screen  " & ScreenNameEnglish & CHR(13) & " Vchr. NO.  " & TxtNoteSerial1.text & CHR(13) & "   Date " & XPDtbTrans & CHR(13) & "  Payment Type " & DCboCashType & CHR(13) & "   Branch  " & dcBranch & CHR(13) & "   Name  " & DBCboClientName & CHR(13) & "  Value" & XPTxtVal & CHR(13) & "   Cash/   Cheque " & CboPayMentType & CHR(13) & "   Box  " & DcboBox & CHR(13) & "   Bank  " & DcboBankName & CHR(13) & "   Cheque No" & TxtChequeNumber & CHR(13) & "  Due Date  " & DtpChequeDueDate & CHR(13) & " Ge NO.  " & TxtNoteSerial & CHR(13) & "Debit " & DcboDebitSide & CHR(13) & "Credit " & DcboCreditSide & CHR(13) & " UserName " & DCboUserName & CHR(13) & " Sales Person " & DcEmp
                     
    If Currentmode <> "D" Then
        AddToLogFile CInt(user_id), 4, Date, Time, LogTextA, LogTexte, Me.Name, Me.TxtModFlg, , , val(TxtNoteSerial), TxtNoteSerial1
    Else
        AddToLogFile CInt(user_id), 4, Date, Time, LogTextA, LogTexte, Me.Name, "D", , , val(TxtNoteSerial), TxtNoteSerial1
    End If
    
End Function
Private Sub TxtNetValue2_Change()
    TxtRemainValue2.text = val(Me.TxtPayedValue2.text) - val(Me.TxtNetValue2.text)
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, _
                         Shift As Integer)
'    On Error GoTo ErrTrap

    If KeyCode = vbKeyReturn Then
        If Me.TxtModFlg.text = "R" Then
            Cmd_Click (0)
        Else
            Sendkeys "{TAB}"
        End If
    End If

    If Me.TxtModFlg.text = "R" Then
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

    If Me.TxtModFlg.text <> "R" Then

        Select Case Me.TxtModFlg.text

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

    If Trim(TxtNoteSerial1.text) <> "" Then
        oldtxtNoteSerial1.text = TxtNoteSerial1.text
    End If

    TxtNoteSerial.text = ""
    TxtNoteSerial1.text = ""
    
    If Me.TxtModFlg.text <> "R" Then
     
    Txt_DateHigri.value = ToHijriDate(XPDtbTrans.value)
'       TxtContNo_Change
End If


End Sub

Private Sub Txt_DateHigri_LostFocus()
      If Me.TxtModFlg.text <> "R" Then
             
             XPDtbTrans.value = ToGregorianDate(Txt_DateHigri.value)

               
        End If
End Sub

Private Sub DcbAccount_Change()
DcbAccount_Click (0)
End Sub

Private Sub DcbAccount_Click(Area As Integer)
TxtAccount.text = getAccountSerial_Code("Account_Serial", "Account_Code", DcbAccount.BoundText)
'If Me.TxtModFlg.Text <> "R" Then
        If CboPayMentType.ListIndex = 4 Then
            Me.DcboDebitSide.BoundText = DcbAccount.BoundText
        End If
' End If
 
End Sub
Private Sub TxtAccount_KeyPress(KeyAscii As Integer)
DcbAccount.BoundText = getAccountSerial_Code("Account_Code", "Account_Serial", TxtAccount.text)
End Sub
Sub ClaCul()

    'Me.lbl(18).Caption = WriteNo(Me.XPTxtVal.text, 0, True)
    'txtAdv_payment_value.text = Format(Val(XPTxtVal.text), "#,###." & String(Abs(SystemOptions.SysDefCurrencyForamt), "0"))
calcnet
    If SystemOptions.NotAllowedCalcVata Then
        TxtVATValue.text = 0
        txtVat2.text = 0
        
    End If
    If SystemOptions.UserInterface = ArabicInterface Then
        Me.lbl(18).Caption = WriteNo(Format(val(XPTxtVal.text) + val(TxtVATValue.text), "0.00"), 0, True, ".", , 0)

    Else
 
        Me.lbl(18).Caption = WriteNo(Format(val(XPTxtVal.text) + val(TxtVATValue.text), "0.00"), 0, True, ".", , 1)

    End If

    'If TxtModFlg.text = "N" Or TxtModFlg.text = "E" And Option3.value = True Then
    If TxtModFlg.text = "N" Or TxtModFlg.text = "E" Then
        txtAdv_payment_value.text = XPTxtVal.text
    End If
CalCulteVAT 1
   If SystemOptions.UserInterface = ArabicInterface Then
        Me.lbl(18).Caption = WriteNo(Format(val(XPTxtVal.text) + val(TxtVATValue.text), "0.00"), 0, True, ".", , 0)

    Else
 
        Me.lbl(18).Caption = WriteNo(Format(val(XPTxtVal.text) + val(TxtVATValue.text), "0.00"), 0, True, ".", , 1)

    End If
End Sub
Private Sub XPTxtVal_Change()
    If SystemOptions.UserInterface = ArabicInterface Then
        Me.lbl(18).Caption = WriteNo(Format(val(XPTxtVal.text) + val(TxtVATValue.text), "0.00"), 0, True, ".", , 0)

    Else
 
        Me.lbl(18).Caption = WriteNo(Format(val(XPTxtVal.text) + val(TxtVATValue.text), "0.00"), 0, True, ".", , 1)

    End If
End Sub

Private Sub XPTxtVal_KeyPress(KeyAscii As Integer)
    KeyAscii = KeyAscii_Num(KeyAscii, XPTxtVal.text, 0)
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
                Msg = "⁄›Ê« «·›« Ê—… —ﬁ„ " & Trim(Me.TxtTransSerial.text)
                Msg = Msg & CHR(13) & "›« Ê—… ‰ﬁœÌ… ...Ê·«Ì„ﬂ‰  Õ’Ì· ·Â« „ﬁ»Ê÷« "
                MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
                TxtTransSerial.SetFocus
                Exit Function
            End If

            If Me.DBCboClientName.BoundText <> IIf(IsNull(RsTemp("CusID").value), "", RsTemp("CusID").value) Then
                Msg = "⁄›Ê« «·›« Ê—… —ﬁ„ " & Trim(Me.TxtTransSerial.text)
                Msg = Msg & CHR(13) & "·Ì”  „”Ã·… „⁄ «·⁄„Ì· " & Me.DBCboClientName.text
                MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
                TxtTransSerial.SetFocus
                Exit Function
            End If

            If LngTransID <> val(Me.TxtTransID.text) Then
                Me.TxtTransID.text = LngTransID
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

            If DblCreditNoteValue < val(Me.XPTxtVal.text) Then
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

            If Me.TxtModFlg.text = "E" Then
                StrSQL = StrSQL + " And Notes.NoteID <>" & Me.XPTxtID.text & ""
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
                ElseIf RsTemp("SumNote_Value").value + val(Me.XPTxtVal.text) > DblCreditNoteValue Then
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
            Msg = "⁄›Ê« «·›« Ê—… —ﬁ„ " & Trim(Me.TxtTransSerial.text)
            Msg = Msg & CHR(13) & "·Ì”  „”Ã·… „⁄ «·⁄„Ì· " & Me.DBCboClientName.text
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
                Msg = "⁄›Ê« «·›« Ê—… —ﬁ„ " & Trim(Me.TxtTransSerial.text)
                Msg = Msg & CHR(13) & "›« Ê—… ‰ﬁœÌ… ...Ê·«Ì„ﬂ‰  Õ’Ì· ·Â« „ﬁ»Ê÷« "
                MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
                TxtTransSerial.SetFocus
                Exit Function
            End If

            If Me.DBCboClientName.BoundText <> IIf(IsNull(RsTemp("CusID").value), "", RsTemp("CusID").value) Then
                Msg = "⁄›Ê« «·›« Ê—… —ﬁ„ " & Trim(Me.TxtTransSerial.text)
                Msg = Msg & CHR(13) & "·Ì”  „”Ã·… „⁄ «·⁄„Ì· " & Me.DBCboClientName.text
                MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
                TxtTransSerial.SetFocus
                Exit Function
            End If

            If LngTransID <> val(Me.TxtTransID.text) Then
                Me.TxtTransID.text = LngTransID
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

            If DblCreditNoteValue < val(Me.XPTxtVal.text) Then
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

            If Me.TxtModFlg.text = "E" Then
                StrSQL = StrSQL + " And Notes.NoteID <>" & Me.XPTxtID.text & ""
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
                ElseIf RsTemp("SumNote_Value").value + val(Me.XPTxtVal.text) > DblCreditNoteValue Then
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
            Msg = "⁄›Ê« «·›« Ê—… —ﬁ„ " & Trim(Me.TxtTransSerial.text)
            Msg = Msg & CHR(13) & "·Ì”  „”Ã·… „⁄ «·⁄„Ì· " & Me.DBCboClientName.text
            MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
            TxtTransSerial.SetFocus
            Exit Function
        End If
    End If

    CheckDebitMaintaince = True
    Exit Function
ErrTrap:
End Function

Public Function CheckDebitService()

End Function

Private Sub WriteInfo()
    Dim rs As ADODB.Recordset
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
    StrSQL = " SELECT     SUM(Note_Value) AS SumX, NoteCashingType"
    StrSQL = StrSQL + " From Notes "
    StrSQL = StrSQL + " Where (NoteType = 4) "
    StrSQL = StrSQL + " AND NoteDate=" & SQLDate(Date, True)
    StrSQL = StrSQL + " GROUP BY NoteCashingType"
    StrSQL = StrSQL + " Order BY NoteCashingType"
    Set rs = New ADODB.Recordset
    rs.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

    If Not (rs.BOF Or rs.EOF) Then
        rs.MoveFirst

        For i = 0 To rs.RecordCount - 1

            If rs("NoteCashingType").value = 0 Then
                Me.LblLinkInfo(0).Caption = rs("SumX").value
            ElseIf rs("NoteCashingType").value = 1 Then
                Me.LblLinkInfo(1).Caption = rs("SumX").value
            End If

            rs.MoveNext
        Next

        Me.LblLinkInfo(6).Caption = val(Me.LblLinkInfo(0).Caption) + val(Me.LblLinkInfo(1).Caption)
    Else
        Me.LblLinkInfo(0).Caption = 0
        Me.LblLinkInfo(1).Caption = 0
        Me.LblLinkInfo(6).Caption = 0
    End If

    '------------------------------------------------------------------------------
    '„ﬁ»Ê÷«  «·√”»Ê⁄ «·Õ«·Ï
    StrSQL = " SELECT     SUM(Note_Value) AS SumX, NoteCashingType"
    StrSQL = StrSQL + " From Notes "
    StrSQL = StrSQL + " Where (NoteType = 4) "
    StrSQL = StrSQL + " AND NoteDate >=" & SQLDate(StartWeekDate, True)
    StrSQL = StrSQL + " AND NoteDate <=" & SQLDate(EndWeekDate, True)
    StrSQL = StrSQL + " GROUP BY NoteCashingType"
    StrSQL = StrSQL + " Order BY NoteCashingType"
    Set rs = New ADODB.Recordset
    rs.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

    If Not (rs.BOF Or rs.EOF) Then
        rs.MoveFirst

        For i = 0 To rs.RecordCount - 1

            If rs("NoteCashingType").value = 0 Then
                Me.LblLinkInfo(2).Caption = rs("SumX").value
            ElseIf rs("NoteCashingType").value = 1 Then
                Me.LblLinkInfo(3).Caption = rs("SumX").value
            End If

            rs.MoveNext
        Next

        Me.LblLinkInfo(7).Caption = val(Me.LblLinkInfo(2).Caption) + val(Me.LblLinkInfo(3).Caption)
    Else
        Me.LblLinkInfo(0).Caption = 0
        Me.LblLinkInfo(1).Caption = 0
        Me.LblLinkInfo(7).Caption = 0
    End If

    '------------------------------------------------------------------------------
    '„ﬁ»Ê÷«  «·‘Â— «·Õ«·Ï
    StrSQL = " SELECT     SUM(Note_Value) AS SumX, NoteCashingType"
    StrSQL = StrSQL + " From Notes "
    StrSQL = StrSQL + " Where (NoteType = 4) "
    StrSQL = StrSQL + " AND Month(NoteDate)=" & Month(Date) & ""
    StrSQL = StrSQL + " AND Year(NoteDate)=" & year(Date) & ""
    StrSQL = StrSQL + " GROUP BY NoteCashingType"
    StrSQL = StrSQL + " Order BY NoteCashingType"
    Set rs = New ADODB.Recordset
    rs.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

    If Not (rs.BOF Or rs.EOF) Then
        rs.MoveFirst

        For i = 0 To rs.RecordCount - 1

            If rs("NoteCashingType").value = 0 Then
                Me.LblLinkInfo(4).Caption = rs("SumX").value
            ElseIf rs("NoteCashingType").value = 1 Then
                Me.LblLinkInfo(5).Caption = rs("SumX").value
            End If

            rs.MoveNext
        Next

        Me.LblLinkInfo(8).Caption = val(Me.LblLinkInfo(4).Caption) + val(Me.LblLinkInfo(5).Caption)
    Else
        Me.LblLinkInfo(4).Caption = 0
        Me.LblLinkInfo(5).Caption = 0
        Me.LblLinkInfo(8).Caption = 0
    End If

End Sub

Private Sub XPTxtVal_Validate(Cancel As Boolean)
ClaCul
End Sub



Private Sub chkPaymentPermission(Optional ByVal IsEdit As Boolean = True)
 
    Dim ctl As Control
    On Error Resume Next
    If Not isChkPaymentType And Not IsEdit Then
        Exit Sub
    End If
    For Each ctl In Me.Controls
        Debug.Print ctl.Name

'        If TypeOf ctl Is ComboBox Then If ctl.Tag <> "not" Then ctl.ListIndex = -1
'        If TypeOf ctl Is OptionButton Then If ctl.Tag <> "not" Then ctl.value = False
'        If TypeOf ctl Is CheckBox Then If ctl.Tag <> "not" Then ctl.value = False
'        If TypeOf ctl Is DataCombo Then If ctl.Tag <> "not" Then ctl.BoundText = ""
        
        
        If TypeOf ctl Is frame Then ctl.Enabled = True: GoTo NextCtl
        If isChkPaymentType Then
            If TypeOf ctl Is C1Elastic Then ctl.Enabled = True: GoTo NextCtl
            If TypeOf ctl Is C1Tab Then ctl.Enabled = True: GoTo NextCtl
            If TypeOf ctl Is Label Then ctl.Enabled = True: GoTo NextCtl
        End If
        
        If ctl.Name = "Cmd" Or ctl.Name = "TxtModFlg" Or ctl.Name = "TxtModFlg1" Or ctl.Name = "TxtModFlg2" Then
            GoTo NextCtl
        Else
            If IsEdit Then
                isChkPaymentType = True
              '  ctl.Tag = ctl.Enabled
                ctl.Enabled = False
            Else
              '  ctl.Enabled = IIf(UCase(ctl.Tag) = "-1", True, False)
            End If
            
            
            
        End If
        Select Case ctl.Name
        Case "TxtNetValue2", "TxtPayedValue2", "TxtRemainValue2", "Command4", "Grid22", "Label20", "lblexit", "CmdValue", "CMDPAy", "CmdNos", "CboPaymentType", "DcboBox", "TXTBankName", "TxtChequeNumber", "DtpChequeDueDate", "Text4", "DTPicker1"
            ctl.Enabled = True
        Case ""
        End Select
       ' if ctl.Name = "TxtNetValue2" Or TxtPayedValue2
        '    If TypeOf Ctl Is TextBox And Ctl.name <> "not" Then Ctl.text = ""
        

        '    If TypeOf Ctl Is XPDatePicker30 Then Ctl.CurrentDate = ""
NextCtl:
    Next
    If IsEdit Then
        dcBranch.Enabled = False
        DBCboClientName.Enabled = False
    Else
        dcBranch.Enabled = True
        DBCboClientName.Enabled = True
    End If
  '  CboPayMentType_Change
End Sub


Private Sub GetDefaultEnabled(Optional ByVal IsEdit As Boolean = True)
 
    Dim ctl As Control
    On Error Resume Next
    
    For Each ctl In Me.Controls
        Debug.Print ctl.Name
        

'        If TypeOf ctl Is ComboBox Then If ctl.Tag <> "not" Then ctl.ListIndex = -1
'        If TypeOf ctl Is OptionButton Then If ctl.Tag <> "not" Then ctl.value = False
'        If TypeOf ctl Is CheckBox Then If ctl.Tag <> "not" Then ctl.value = False
'        If TypeOf ctl Is DataCombo Then If ctl.Tag <> "not" Then ctl.BoundText = ""
        
        
        If TypeOf ctl Is frame Then ctl.Enabled = True: GoTo NextCtl
        If isChkPaymentType Then
            If TypeOf ctl Is C1Elastic Then ctl.Enabled = True: GoTo NextCtl
            If TypeOf ctl Is C1Tab Then ctl.Enabled = True: GoTo NextCtl
            If TypeOf ctl Is Label Then ctl.Enabled = True: GoTo NextCtl
        End If
        
        If ctl.Name = "Cmd" Or ctl.Name = "TxtModFlg" Or ctl.Name = "TxtModFlg1" Or ctl.Name = "TxtModFlg2" Then
            GoTo NextCtl
        Else
            If Not IsEdit Then
                
                ctl.Tag = ctl.Enabled
                
            Else
                ctl.Enabled = IIf(UCase(ctl.Tag) = "-1", True, False)
                
            End If
            Debug.Print ctl.Enabled
            
            
        End If
       ' if ctl.Name = "TxtNetValue2" Or TxtPayedValue2
        '    If TypeOf Ctl Is TextBox And Ctl.name <> "not" Then Ctl.text = ""
        

        '    If TypeOf Ctl Is XPDatePicker30 Then Ctl.CurrentDate = ""
NextCtl:
    Next
  '  CboPayMentType_Change
End Sub
'
Function GetBrnchproject(Optional CusID As Double) As Integer
Dim sql As String
Dim rs2 As ADODB.Recordset
Set rs2 = New ADODB.Recordset
sql = " Select Branch_NO as BranchId from projects"

sql = sql & " Where (Id = " & CusID & ")"
rs2.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
If rs2.RecordCount > 0 Then
GetBrnchproject = IIf(IsNull(rs2("BranchId").value), val(Me.dcBranch.BoundText), rs2("BranchId").value)
Else
GetBrnchproject = val(Me.dcBranch.BoundText)
End If
If GetBrnchproject = 0 Then
GetBrnchproject = val(Me.dcBranch.BoundText)
End If
End Function


Function GetBrnchEmployee(Optional CusID As Double) As Integer
Dim sql As String
Dim rs2 As ADODB.Recordset
Set rs2 = New ADODB.Recordset
sql = " Select BranchId from TblEmployee"

sql = sql & " Where (Id = " & CusID & ")"
rs2.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
If rs2.RecordCount > 0 Then
GetBrnchEmployee = IIf(IsNull(rs2("BranchId").value), val(Me.dcBranch.BoundText), rs2("BranchId").value)
Else
GetBrnchEmployee = val(Me.dcBranch.BoundText)
End If
If GetBrnchEmployee = 0 Then
GetBrnchEmployee = val(Me.dcBranch.BoundText)
End If
End Function




Function GetBrnchCustomer(Optional CusID As Double) As Integer
Dim sql As String
Dim rs2 As ADODB.Recordset
Set rs2 = New ADODB.Recordset
sql = " SELECT     BranchId"
sql = sql & " From dbo.TblCustemers"
sql = sql & " Where (CusID = " & CusID & ")"
rs2.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
If rs2.RecordCount > 0 Then
GetBrnchCustomer = IIf(IsNull(rs2("BranchId").value), val(Me.dcBranch.BoundText), rs2("BranchId").value)
Else
GetBrnchCustomer = val(Me.dcBranch.BoundText)
End If
If GetBrnchCustomer = 0 Then
GetBrnchCustomer = val(Me.dcBranch.BoundText)
End If
End Function

Function GetBrnchBank(Optional BankID As Double) As Integer
Dim sql As String
Dim rs2 As ADODB.Recordset
Set rs2 = New ADODB.Recordset
sql = " SELECT     BranchId"
sql = sql & " From dbo.BanksData"
sql = sql & " Where (BankID = " & BankID & ")"
rs2.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
If rs2.RecordCount > 0 Then
GetBrnchBank = IIf(IsNull(rs2("BranchId").value), val(Me.dcBranch.BoundText), rs2("BranchId").value)
Else
GetBrnchBank = val(Me.dcBranch.BoundText)
End If
If GetBrnchBank = 0 Then
GetBrnchBank = val(Me.dcBranch.BoundText)
End If
End Function

