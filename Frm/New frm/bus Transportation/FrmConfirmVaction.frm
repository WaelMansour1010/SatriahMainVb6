VERSION 5.00
Object = "{0AFE7BE0-11B7-4A3E-978D-D4501E9A57FE}#1.0#0"; "c1sizer.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{FE5DCFAD-BC1D-11D2-94CF-004005455FAA}#1.4#0"; "ImpulseButton.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form FrmConfirmVaction 
   BackColor       =   &H00E2E9E9&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "‘«‘… «À»«  «· ⁄ÿ·"
   ClientHeight    =   5730
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9285
   Icon            =   "FrmConfirmVaction.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   RightToLeft     =   -1  'True
   ScaleHeight     =   5730
   ScaleWidth      =   9285
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
      Height          =   5724
      Left            =   0
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   0
      Width           =   9288
      _cx             =   16378
      _cy             =   10107
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
         Height          =   732
         Left            =   120
         TabIndex        =   30
         TabStop         =   0   'False
         Top             =   3720
         Width           =   5772
         _cx             =   10186
         _cy             =   1296
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
            Left            =   3204
            RightToLeft     =   -1  'True
            TabIndex        =   34
            Top             =   240
            Width           =   828
         End
         Begin VB.Label XPTxtCount 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Height          =   312
            Left            =   240
            RightToLeft     =   -1  'True
            TabIndex        =   33
            Top             =   240
            Width           =   708
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   " «·”Ã· «·Õ«·Ì:"
            Height          =   312
            Index           =   2
            Left            =   4080
            RightToLeft     =   -1  'True
            TabIndex        =   32
            Top             =   240
            Width           =   1152
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   " ⁄œœ «·”Ã·« :"
            Height          =   312
            Index           =   4
            Left            =   984
            RightToLeft     =   -1  'True
            TabIndex        =   31
            Top             =   240
            Width           =   1152
         End
      End
      Begin C1SizerLibCtl.C1Elastic C1Elastic2 
         Height          =   2892
         Left            =   120
         TabIndex        =   26
         TabStop         =   0   'False
         Top             =   720
         Width           =   9012
         _cx             =   15901
         _cy             =   5106
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
         Begin VB.TextBox txtDC 
            Alignment       =   1  'Right Justify
            Height          =   312
            Left            =   360
            Locked          =   -1  'True
            MaxLength       =   10
            RightToLeft     =   -1  'True
            TabIndex        =   43
            Top             =   1440
            Width           =   3288
         End
         Begin VB.TextBox txtdayvalue 
            Alignment       =   1  'Right Justify
            Height          =   312
            Left            =   7272
            MaxLength       =   10
            RightToLeft     =   -1  'True
            TabIndex        =   41
            Top             =   3720
            Visible         =   0   'False
            Width           =   768
         End
         Begin VB.TextBox txtRemarks 
            Alignment       =   1  'Right Justify
            Height          =   672
            Left            =   384
            MaxLength       =   50
            RightToLeft     =   -1  'True
            TabIndex        =   10
            Top             =   2160
            Width           =   7536
         End
         Begin VB.TextBox txtID 
            Alignment       =   1  'Right Justify
            Enabled         =   0   'False
            Height          =   312
            Left            =   4992
            Locked          =   -1  'True
            MaxLength       =   10
            RightToLeft     =   -1  'True
            TabIndex        =   1
            Top             =   360
            Width           =   2928
         End
         Begin MSComCtl2.DTPicker FromDate 
            Height          =   336
            Left            =   2064
            TabIndex        =   4
            Top             =   720
            Width           =   1620
            _ExtentX        =   2858
            _ExtentY        =   582
            _Version        =   393216
            CalendarBackColor=   12648447
            CustomFormat    =   "yyyy/M/d"
            Format          =   98762755
            CurrentDate     =   38718
         End
         Begin Dynamic_Byte.NourHijriCal FromDateH 
            Height          =   336
            Left            =   360
            TabIndex        =   5
            Top             =   720
            Width           =   1596
            _ExtentX        =   2805
            _ExtentY        =   582
         End
         Begin MSDataListLib.DataCombo dcDuration 
            Height          =   288
            Left            =   4992
            TabIndex        =   3
            Top             =   720
            Width           =   2928
            _ExtentX        =   5159
            _ExtentY        =   556
            _Version        =   393216
            Style           =   2
            Text            =   ""
            RightToLeft     =   -1  'True
         End
         Begin MSDataListLib.DataCombo dcVacationType 
            Height          =   288
            Left            =   360
            TabIndex        =   2
            Top             =   360
            Width           =   3300
            _ExtentX        =   5821
            _ExtentY        =   556
            _Version        =   393216
            Style           =   2
            Text            =   ""
            RightToLeft     =   -1  'True
         End
         Begin MSDataListLib.DataCombo dcCity 
            Height          =   288
            Left            =   4992
            TabIndex        =   9
            Top             =   1440
            Width           =   2928
            _ExtentX        =   5159
            _ExtentY        =   556
            _Version        =   393216
            Text            =   ""
            RightToLeft     =   -1  'True
         End
         Begin MSComCtl2.DTPicker ToDate 
            Height          =   336
            Left            =   2064
            TabIndex        =   7
            Top             =   1080
            Width           =   1620
            _ExtentX        =   2858
            _ExtentY        =   582
            _Version        =   393216
            CalendarBackColor=   12648447
            CustomFormat    =   "yyyy/M/d"
            Format          =   98762755
            CurrentDate     =   38718
         End
         Begin Dynamic_Byte.NourHijriCal TODateH 
            Height          =   336
            Left            =   360
            TabIndex        =   8
            Top             =   1080
            Width           =   1596
            _ExtentX        =   2805
            _ExtentY        =   582
         End
         Begin MSDataListLib.DataCombo dcMonth 
            Height          =   288
            Left            =   4992
            TabIndex        =   6
            Top             =   1080
            Width           =   2928
            _ExtentX        =   5159
            _ExtentY        =   556
            _Version        =   393216
            Style           =   2
            Text            =   ""
            RightToLeft     =   -1  'True
         End
         Begin MSDataListLib.DataCombo DcMangerialArea 
            Height          =   288
            Left            =   4992
            TabIndex        =   44
            Top             =   1800
            Width           =   2928
            _ExtentX        =   5159
            _ExtentY        =   556
            _Version        =   393216
            Style           =   2
            Text            =   ""
            RightToLeft     =   -1  'True
         End
         Begin MSDataListLib.DataCombo DCboUserName 
            Height          =   288
            Left            =   360
            TabIndex        =   47
            Top             =   1800
            Width           =   3288
            _ExtentX        =   5794
            _ExtentY        =   556
            _Version        =   393216
            Enabled         =   0   'False
            Text            =   ""
            RightToLeft     =   -1  'True
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "Õ—— »Ê«”ÿ… : "
            Height          =   288
            Index           =   13
            Left            =   3708
            RightToLeft     =   -1  'True
            TabIndex        =   48
            Top             =   1812
            Width           =   900
         End
         Begin VB.Label Label4 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "«·„‰ÿﬁ… «·«œ«—Ì…"
            ForeColor       =   &H00000000&
            Height          =   372
            Left            =   7824
            TabIndex        =   45
            Top             =   1800
            Width           =   1068
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "ﬁÌ„… «·ÌÊ„"
            Height          =   288
            Index           =   9
            Left            =   7872
            RightToLeft     =   -1  'True
            TabIndex        =   42
            Top             =   3720
            Visible         =   0   'False
            Width           =   1020
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "⁄œœ «·«Ì«„"
            Height          =   312
            Index           =   7
            Left            =   3720
            RightToLeft     =   -1  'True
            TabIndex        =   40
            Top             =   1440
            Width           =   900
         End
         Begin VB.Label Label2 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "«·› —…"
            ForeColor       =   &H00000000&
            Height          =   372
            Left            =   7824
            TabIndex        =   39
            Top             =   1080
            Width           =   1068
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "«·Ï  «—ÌŒ "
            Height          =   312
            Index           =   5
            Left            =   3708
            RightToLeft     =   -1  'True
            TabIndex        =   38
            Top             =   1080
            Width           =   900
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "«·„Õ«›Ÿ…"
            ForeColor       =   &H00000000&
            Height          =   372
            Left            =   7824
            TabIndex        =   37
            Top             =   1440
            Width           =   1068
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "„‰  «—ÌŒ "
            Height          =   312
            Index           =   6
            Left            =   3708
            RightToLeft     =   -1  'True
            TabIndex        =   36
            Top             =   720
            Width           =   900
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "„·«ÕŸ« "
            Height          =   312
            Index           =   1
            Left            =   7908
            RightToLeft     =   -1  'True
            TabIndex        =   35
            Top             =   2160
            Width           =   984
         End
         Begin VB.Label Label3 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "«·”‰… «·œ—«”Ì…"
            ForeColor       =   &H00000000&
            Height          =   372
            Left            =   7824
            TabIndex        =   29
            Top             =   720
            Width           =   1068
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "„”·”·"
            Height          =   288
            Index           =   0
            Left            =   7872
            RightToLeft     =   -1  'True
            TabIndex        =   28
            Top             =   360
            Width           =   1020
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "‰Ê⁄ «·⁄ÿ·…"
            Height          =   312
            Index           =   3
            Left            =   3648
            RightToLeft     =   -1  'True
            TabIndex        =   27
            Top             =   360
            Width           =   984
         End
      End
      Begin C1SizerLibCtl.C1Elastic EleHeader 
         Height          =   588
         Left            =   0
         TabIndex        =   20
         TabStop         =   0   'False
         Top             =   0
         Width           =   9432
         _cx             =   16642
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
         Appearance      =   0
         MousePointer    =   0
         Version         =   801
         BackColor       =   16777215
         ForeColor       =   4210688
         FloodColor      =   6553600
         ForeColorDisabled=   -2147483631
         Caption         =   "   ‘«‘… «À»«  «· ⁄ÿ·   "
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
            TabIndex        =   21
            Top             =   180
            Visible         =   0   'False
            Width           =   855
         End
         Begin ImpulseButton.ISButton XPBtnMove 
            Height          =   345
            Index           =   0
            Left            =   1155
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
            ButtonImage     =   "FrmConfirmVaction.frx":038A
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
            ButtonImage     =   "FrmConfirmVaction.frx":0724
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
            TabIndex        =   24
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
            ButtonImage     =   "FrmConfirmVaction.frx":0ABE
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
            TabIndex        =   25
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
            ButtonImage     =   "FrmConfirmVaction.frx":0E58
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
      Begin C1SizerLibCtl.C1Elastic C1Elastic5 
         Height          =   816
         Left            =   120
         TabIndex        =   11
         TabStop         =   0   'False
         Top             =   4560
         Width           =   9012
         _cx             =   15901
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
            Height          =   528
            Index           =   0
            Left            =   7968
            TabIndex        =   12
            Top             =   144
            Width           =   984
            _ExtentX        =   1746
            _ExtentY        =   926
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
            ButtonImage     =   "FrmConfirmVaction.frx":11F2
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
            Height          =   528
            Index           =   1
            Left            =   6996
            TabIndex        =   13
            Top             =   144
            Width           =   876
            _ExtentX        =   1535
            _ExtentY        =   926
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
            ButtonImage     =   "FrmConfirmVaction.frx":7A54
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
            Height          =   528
            Index           =   2
            Left            =   5916
            TabIndex        =   14
            Top             =   144
            Width           =   984
            _ExtentX        =   1746
            _ExtentY        =   926
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
            ButtonImage     =   "FrmConfirmVaction.frx":E2B6
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
            Height          =   528
            Index           =   3
            Left            =   4932
            TabIndex        =   15
            Top             =   144
            Width           =   984
            _ExtentX        =   1746
            _ExtentY        =   926
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
            ButtonImage     =   "FrmConfirmVaction.frx":14B18
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
            Height          =   528
            Index           =   4
            Left            =   3924
            TabIndex        =   16
            Top             =   144
            Width           =   984
            _ExtentX        =   1746
            _ExtentY        =   926
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
            ButtonImage     =   "FrmConfirmVaction.frx":1B37A
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
            Height          =   528
            Index           =   6
            Left            =   984
            TabIndex        =   18
            Top             =   144
            Width           =   984
            _ExtentX        =   1746
            _ExtentY        =   926
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
            ButtonImage     =   "FrmConfirmVaction.frx":21BDC
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
            Height          =   528
            Left            =   120
            TabIndex        =   19
            Top             =   144
            Width           =   864
            _ExtentX        =   1535
            _ExtentY        =   926
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
            ButtonImage     =   "FrmConfirmVaction.frx":4B7FE
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
            Height          =   528
            Index           =   7
            Left            =   2964
            TabIndex        =   17
            Top             =   144
            Width           =   876
            _ExtentX        =   1535
            _ExtentY        =   926
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
            ButtonImage     =   "FrmConfirmVaction.frx":52060
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
            Height          =   528
            Index           =   5
            Left            =   2040
            TabIndex        =   46
            Top             =   144
            Width           =   876
            _ExtentX        =   1535
            _ExtentY        =   926
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
            ButtonImage     =   "FrmConfirmVaction.frx":588C2
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
   End
End
Attribute VB_Name = "FrmConfirmVaction"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rs As ADODB.Recordset
Dim Rs_Temp As ADODB.Recordset
Dim Rs_Temp2 As ADODB.Recordset
Dim rsVendor As ADODB.Recordset
Dim TTP As clstooltip

Dim FromDate_ As Date
Dim ToDate_ As Date
Dim FromDateH_ As String
Dim ToDateH_ As String


Private Sub Cmd_Click(Index As Integer)
 '    On Error GoTo ErrTrap
 
 

    Select Case Index
        Case 0
            If DoPremis(Do_New, Me.Name, True) = False Then
                Exit Sub
            End If
            TxtModFlg.Text = "N"
            clear_all Me
            txtID.Text = CStr(new_id("TblConfirmVacation", "ID", "", True))
         '   txtName.SetFocus
        Case 1
                                     If ChekClodePeriod(FromDate.value) = True Then
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
            
            
             If ISAllowDeleteUpdateContract = False Then
                MsgBox ("·«Ì„ﬂ‰ «· ⁄œÌ· ⁄·Ï «·”‰œ »”»» ⁄„· ÿ·» ’—› ⁄·ÌÂ")
                Exit Sub
            Else
                    TxtModFlg.Text = "E"
            End If
        Case 2

                                               If ChekClodePeriod(FromDate.value) = True Then
               If SystemOptions.UserInterface = ArabicInterface Then
                MsgBox "Ì—ÃÏ  €ÌÌ—  «—ÌŒ «·Õ—ﬂ… ·«‰ Â–Â «·› —… „€·ﬁ…"
               Else
               MsgBox "Please Change Date Becouse This is Period is Closed"
              End If
              Exit Sub
              End If
            SaveData

        Case 3
            Undo

        Case 4
                                     If ChekClodePeriod(FromDate.value) = True Then
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
            
            
             If ISAllowDeleteUpdateContract = False Then
                MsgBox ("·«Ì„ﬂ‰ Õ–› «·”‰œ »”»» ⁄„· ÿ·» ’—› ⁄·ÌÂ")
                Exit Sub
             Else
                Del_Company
             End If
        Case 5
                    Unload FrmSearch_BasicData
                    FrmSearch_BasicData.SendForm = "confirmVacation"
                    FrmSearch_BasicData.show vbModal
        Case 6
            Unload Me
         Case 7
            print_report
    End Select

    Exit Sub
ErrTrap:
End Sub

Private Sub CmdHelp_Click()
    SystemOptions.SysHelp.HHTopicID = Me.HelpContextID
    SystemOptions.SysHelp.HHDisplayTopicID Me.hwnd
End Sub

 

Private Sub dcContract_Click(Area As Integer)

 


End Sub

Private Sub dcViolation_Click(Area As Integer)

End Sub



Private Sub dtpDate_Change()
        VBA.Calendar = vbCalGreg
       ' dtpDateH.value = ToHijriDate(dtpDate.value)
End Sub

Private Sub dtpDateH_LostFocus()
'dtpDate.value = ToGregorianDate(dtpDateH.value)
End Sub

Private Sub Dtp_CallbackKeyDown(ByVal KeyCode As Integer, ByVal Shift As Integer, ByVal CallbackField As String, CallbackDate As Date)

End Sub

Private Sub CmdAttach_Click()
    On Error Resume Next
          If DoPremis(Do_Attach, Me.Name, True) = False Then
                Exit Sub
            End If
            
ShowAttachments txtID, "15062020001"

End Sub

Private Sub dcCity_Click(Area As Integer)
Dim str As String
Set Rs_Temp = New ADODB.Recordset
Set DcMangerialArea.RowSource = Rs_Temp
If SystemOptions.UserInterface = ArabicInterface Then
    str = " Select ID , Name   from TblManagerialArea  where cityid = " & val(dcCity.BoundText)
Else
    str = " Select ID , NameE   from TblManagerialArea  where cityid = " & val(dcCity.BoundText)
End If
fill_combo DcMangerialArea, str
DcMangerialArea.Refresh
End Sub

Private Sub dcDuration_Change()
Dim i  As Integer, str As String, Typ As Integer
    i = val(dcDuration.BoundText)
    If i > 0 Then
        str = "  select id , Name  from TblDurations_Details where did =   " & i
        fill_combo dcMonth, str
    Else
        str = "  select id , Name  from TblDurations_Details where did =   " & -1
        fill_combo dcMonth, str
    End If
    
    str = " select * from TblDurations where id =   " & i
    Set Rs_Temp = New ADODB.Recordset
    Rs_Temp.Open str, Cn, adOpenStatic, adLockOptimistic, adCmdText
    
    If Rs_Temp.RecordCount > 0 Then
         Typ = IIf(IsNull(Rs_Temp("type").value), -1, Rs_Temp("type").value)
         If Typ = 0 Then
                FromDate.Enabled = True
                ToDate.Enabled = True
                FromDateH.Enabled = False
                ToDateH.Enabled = False
         ElseIf Typ = 1 Then
                FromDate.Enabled = False
                ToDate.Enabled = False
                FromDateH.Enabled = True
                ToDateH.Enabled = True
         End If
      
        
         
    End If
End Sub

Private Sub dcDuration_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyF3 Then
   '  Unload FrmSearch_Duration
   '    FrmSearch_Duration.SendForm = "ConfirmVacation"
   '     FrmSearch_Duration.show
End If
End Sub

Private Sub dcMonth_Click(Area As Integer)
Dim str As String
  str = " select * from TblDurations_details where id =   " & val(dcMonth.BoundText)
  Set Rs_Temp = New ADODB.Recordset
  Rs_Temp.Open str, Cn, adOpenStatic, adLockOptimistic, adCmdText
    
   If Rs_Temp.RecordCount > 0 Then
        FromDate_ = IIf(IsNull(Rs_Temp("FromDate").value), Date, Rs_Temp("FromDate").value)
        ToDate_ = IIf(IsNull(Rs_Temp("ToDate").value), Date, Rs_Temp("ToDate").value)
        FromDateH_ = IIf(IsNull(Rs_Temp("FromDateH").value), ToHijriDate(Date), Rs_Temp("FromDateH").value)
        ToDateH_ = IIf(IsNull(Rs_Temp("ToDateH").value), ToHijriDate(Date), Rs_Temp("ToDateH").value)
        
        FromDate.value = FromDate_
        ToDate.value = ToDate_
        FromDateH.value = FromDateH_
        ToDateH.value = ToDateH_
       
        
        If FromDate.Enabled = True Then
              '  FromDate.MinDate = Rs_Temp("FromDate").value
              '  FromDate.MaxDate = Rs_Temp("ToDate").value
              '  ToDate.MinDate = Rs_Temp("FromDate").value
              '  ToDate.MaxDate = Rs_Temp("ToDate").value
        Else
            '    FromDateH.Mi = rs_temp("FromDate").value
            '    FromDate.MaxDate = rs_temp("ToDate").value
            '    ToDate.MinDate = rs_temp("FromDate").value
            '    ToDate.MaxDate = rs_temp("ToDate").value
        End If
   End If


End Sub

Private Sub dcVacationType_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyF3 Then
        Unload FrmSearch_BasicData
        FrmSearch_BasicData.SendForm = "ConfirmVacation"
        FrmSearch_BasicData.show
        
End If
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
    Dim str As String

    
    Set Dcombos = New ClsDataCombos
    'Dcombos.GetCustomersSuppliers 2, dcVendor
    Dcombos.getCountriesGovernments dcCity
    Dcombos.GetUsers Me.DCboUserName
         
   str = "  select id , name  from TblDurations  "
   fill_combo dcDuration, str
   
   str = "   select id , name from TblVacationTypes  "
   fill_combo dcVacationType, str
   
   
   str = "  select id , name from TblManagerialArea "
   fill_combo DcMangerialArea, str
   
   str = "   "
    
    If SystemOptions.UserInterface = EnglishInterface Then
        SetInterface Me
        ChangeLang
    End If

    LogTextA = "   «·œŒÊ· «·Ì ‘«‘… " & " «À»«  «· ⁄ÿ·  "
    LogTexte = " Open Window " & " Confirm  Violation "
    AddToLogFile CInt(user_id), 0, Date, Time, LogTextA, LogTexte, Me.Name, "O", "", ""

    Dim My_SQL As String
       
    
    Resize_Form Me
    
    AddTip
    Set rs = New ADODB.Recordset
    Dim StrSQL As String
    StrSQL = "SELECT  *  From TblConfirmVacation "
    rs.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText
      
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
    LogTextA = "     «·Œ—ÊÃ „‰ ‘«‘… " & "  »Ì«‰«  «À»«  «· ⁄ÿ·  "
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


Private Sub FromDate_Change()
If FromDate.value < FromDate_ Or FromDate.value > ToDate_ Then
        FromDate.value = FromDate_
        Exit Sub
End If

FromDateH.value = ToHijriDate(FromDate.value)

Dim str As String
str = CStr(DateDiff("d", FromDate.value, ToDate.value) + 1)
txtDC.Text = str
End Sub


Private Sub FromDate_GotFocus()
If dcDuration.BoundText = "" Then
    MsgBox ("«Œ — «·”‰… «·œ—«”Ì… «Ê·« ")
    dcDuration.SetFocus
    Exit Sub
End If

If dcMonth.BoundText = "" Then
    MsgBox ("«Œ — «·› —… «Ê·« ")
    dcMonth.SetFocus
    Exit Sub
End If
End Sub

Private Sub FromDateH_GotFocus()
If dcDuration.BoundText = "" Then
    MsgBox ("«Œ — «·”‰… «·œ—«”Ì… «Ê·« ")
    dcDuration.SetFocus
    Exit Sub
End If

If dcMonth.BoundText = "" Then
    MsgBox ("«Œ — «·› —… «Ê·« ")
    dcMonth.SetFocus
    Exit Sub
End If
End Sub

Private Sub Fromdateh_LostFocus()

If FromDateH.value < FromDateH_ Or FromDateH.value > ToDateH_ Then
        FromDateH.value = FromDateH_
        Exit Sub
End If


VBA.Calendar = vbCalGreg
FromDate.value = ToGregorianDate(FromDateH.value)
txtDC.Text = DateDiff("d", FromDateH.value, ToDateH.value) + 1
End Sub


Private Sub ToDate_Change()
    
If ToDate.value < FromDate_ Or ToDate.value > ToDate_ Then
        ToDate.value = ToDate_
        Exit Sub
End If

ToDateH.value = ToHijriDate(ToDate.value)
txtDC.Text = DateDiff("d", FromDate.value, ToDate.value) + 1
End Sub

Private Sub ToDate_GotFocus()
If dcDuration.BoundText = "" Then
    MsgBox ("«Œ — «·”‰… «·œ—«”Ì… «Ê·« ")
    dcDuration.SetFocus
    Exit Sub
End If

If dcMonth.BoundText = "" Then
    MsgBox ("«Œ — «·› —… «Ê·« ")
    dcMonth.SetFocus
    Exit Sub
End If
End Sub

Private Sub TODateH_GotFocus()
If dcDuration.BoundText = "" Then
    MsgBox ("«Œ — «·”‰… «·œ—«”Ì… «Ê·« ")
    dcDuration.SetFocus
    Exit Sub
End If

If dcMonth.BoundText = "" Then
    dcMonth.SetFocus
    MsgBox ("«Œ — «·› —… «Ê·« ")
    Exit Sub
End If
End Sub

Private Sub ToDateH_LostFocus()

If ToDateH.value < FromDateH_ Or ToDateH.value > ToDateH_ Then
        ToDateH.value = ToDateH_
        Exit Sub
End If


VBA.Calendar = vbCalGreg
ToDate.value = ToGregorianDate(ToDateH.value)
txtDC.Text = DateDiff("d", FromDateH.value, ToDateH.value) + 1
End Sub

Private Sub TxtModFlg_Change()
    On Error GoTo ErrTrap

    Select Case Me.TxtModFlg.Text

        Case "R"

            If SystemOptions.UserInterface = ArabicInterface Then
                Me.Caption = "»Ì«‰«  «À»«  «· ⁄ÿ· "
            Else
                Me.Caption = "Violation Types"
            End If

            Me.Cmd(2).Enabled = False
            Me.Cmd(3).Enabled = False
        
            Me.Cmd(0).Enabled = True
            Me.Cmd(1).Enabled = True
            Me.Cmd(4).Enabled = True
        
            Me.XPBtnMove(0).Enabled = True
            Me.XPBtnMove(1).Enabled = True
            Me.XPBtnMove(2).Enabled = True
            Me.XPBtnMove(3).Enabled = True
        
            Me.txtID.locked = True
            'Me.txtName.locked = True
          '  Me.XPMTxtRemark.locked = True

            If rs.RecordCount < 1 Then
                Me.XPBtnMove(0).Enabled = False
                Me.XPBtnMove(1).Enabled = False
                Me.XPBtnMove(2).Enabled = False
                Me.XPBtnMove(3).Enabled = False
                Me.Cmd(1).Enabled = False
                Me.Cmd(4).Enabled = False
            End If
C1Elastic2.Enabled = False
        Case "N"

            If SystemOptions.UserInterface = ArabicInterface Then
                Me.Caption = "»Ì«‰«  «À»«  «· ⁄ÿ·( ÃœÌœ )"
            Else
                Me.Caption = "Violation Types (New)"
            End If
        
            Me.Cmd(2).Enabled = True
            Me.Cmd(3).Enabled = True

            If SystemOptions.UserInterface = ArabicInterface Then
                Me.Caption = "»Ì«‰«  «À»«  «· ⁄ÿ·( ÃœÌœ )"
            Else
                Me.Caption = "Violation Types(New)"
            End If
        
            Me.Cmd(0).Enabled = False
            Me.Cmd(1).Enabled = False
            Me.Cmd(4).Enabled = False
            Me.txtID.locked = True
          '  Me.txtName.locked = False
       C1Elastic2.Enabled = True
        Case "E"

            If SystemOptions.UserInterface = ArabicInterface Then
                Me.Caption = "»Ì«‰«  «À»«  «· ⁄ÿ· (  ⁄œÌ· )"
            Else
                Me.Caption = "Violation Types(Edit)"
            End If
        
            Me.Cmd(2).Enabled = True
            Me.Cmd(3).Enabled = True
        
            Me.Cmd(0).Enabled = False
            Me.Cmd(1).Enabled = False
            Me.Cmd(4).Enabled = False
        
            Me.XPBtnMove(0).Enabled = False
            Me.XPBtnMove(1).Enabled = False
            Me.XPBtnMove(2).Enabled = False
            Me.XPBtnMove(3).Enabled = False
        
            Me.txtID.locked = True
           ' Me.txtName.locked = False
       '     Me.XPMTxtRemark.locked = False
       C1Elastic2.Enabled = True
    End Select

    Exit Sub
ErrTrap:
End Sub

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
            rs.find "ID =" & Lngid, , adSearchForward, adBookmarkFirst

            If rs.EOF Or rs.BOF Then
                Exit Sub
            End If
        End If
    End If


    txtID.Text = IIf(IsNull(rs("ID").value), "", val(rs("ID").value))
    dcDuration.BoundText = IIf(IsNull(rs("DurationID").value), "", Trim(rs("DurationID").value))
    dcCity.BoundText = IIf(IsNull(rs("CityID").value), "", Trim(rs("CityID").value))
    dcVacationType.BoundText = IIf(IsNull(rs("VacationTypeID").value), "", Trim(rs("VacationTypeID").value))
    dcMonth.BoundText = IIf(IsNull(rs("MonthID").value), "", rs("MonthID").value)
    txtdayvalue.Text = IIf(IsNull(rs("DayValue").value), "", rs("DayValue").value)
    FromDate.value = IIf(IsNull(rs("FromDate").value), Date, Trim(rs("FromDate").value))
    FromDateH.value = IIf(IsNull(rs("FromDateH").value), ToHijriDate(Date), Trim(rs("FromDateH").value))
    ToDate.value = IIf(IsNull(rs("ToDate").value), Date, Trim(rs("ToDate").value))
    ToDateH.value = IIf(IsNull(rs("TODateH").value), ToHijriDate(Date), Trim(rs("TODateH").value))
    txtRemarks.Text = IIf(IsNull(rs("remarks").value), "", rs("remarks").value)
    DcMangerialArea.BoundText = IIf(IsNull(rs("MangerialAreaID").value), "", rs("MangerialAreaID").value)
    txtDC.Text = IIf(IsNull(rs("daycount").value), "", rs("daycount").value)
     Me.DCboUserName.BoundText = IIf(IsNull(rs("UserID").value), "", rs("UserID").value)
    XPTxtCurrent.Caption = rs.AbsolutePosition
    txtDC.Text = IIf(IsNull(rs("Daycount").value), "", rs("Daycount").value)
    XPTxtCount.Caption = rs.RecordCount
    Exit Sub
ErrTrap:
End Sub




Private Sub TxtName_GotFocus()
On Error Resume Next
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
    
        If dcDuration.BoundText = "" Then
            MsgBox "„‰ ›÷·ﬂ «Œ — «·› —… ", vbOKOnly + vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
            dcDuration.SetFocus
            SendKeys ("{F4}")
            Exit Sub
        End If

        If dcVacationType.BoundText = "" Then
            MsgBox "„‰ ›÷·ﬂ  «Œ — ‰Ê⁄ «· ⁄ÿ·  ", vbOKOnly + vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
            dcVacationType.SetFocus
            'SendKeys ("{F4}")
            Exit Sub
        End If
        
       If DcMangerialArea.BoundText = "" Then
            MsgBox "„‰ ›÷·ﬂ  «Œ — «·„‰ÿﬁ… «·«œ«—Ì… ", vbOKOnly + vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
            DcMangerialArea.SetFocus
            SendKeys ("{F4}")
            Exit Sub
        End If
        
        
          If dcMonth.BoundText = "" Then
            MsgBox "„‰ ›÷·ﬂ  «Œ — «·› —… ", vbOKOnly + vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
            dcMonth.SetFocus
            'SendKeys ("{F4}")
            Exit Sub
        End If
       
        Select Case Me.TxtModFlg.Text
            Case "N"
            rs.AddNew
            txtID.Text = CStr(new_id("TblConfirmVacation", "ID", "", True))
            Case "E"
              '  StrSQL = "select * From  TblViolationTypes where Name='" & Trim(txtName.text) & "'"
                StrSQL = "delete From TblConfirmVacation_Details where  HID =" & val(txtID.Text)
                Cn.Execute StrSQL, , adExecuteNoRecords
        End Select

        Cn.BeginTrans
        BeginTrans = True
          
        rs("ID").value = val(txtID.Text)
        rs("DurationID").value = IIf(dcDuration.BoundText = "", Null, dcDuration.BoundText)
        rs("VacationTypeID").value = IIf(dcVacationType.BoundText = "", Null, dcVacationType.BoundText)
        rs("CityID").value = IIf(dcCity.BoundText = "", Null, dcCity.BoundText)
        rs("Remarks") = txtRemarks.Text
        rs("FromDate") = IIf(IsNull(FromDate.value), Date, FromDate.value)
        rs("FromDateH") = IIf(IsNull(FromDateH.value), ToHijriDate(Date), FromDateH.value)
        rs("ToDate") = IIf(IsNull(ToDate.value), Date, ToDate.value)
        rs("ToDateH") = IIf(IsNull(ToDateH.value), ToHijriDate(Date), ToDateH.value)
        rs("MonthID") = IIf(dcMonth.BoundText = "", Null, dcMonth.BoundText)
        rs("CreationDate") = Date
        rs("DayValue") = val(txtdayvalue.Text)
        rs("MangerialAreaID") = val(DcMangerialArea.BoundText)
        rs("UserID") = user_id
        rs("daycount") = val(txtDC.Text)
        rs.update
        
        
        Dim str As String, k As Integer
        str = " select * from tblschoolefile where CityID =  " & val(dcCity.BoundText)
        Set Rs_Temp = New ADODB.Recordset
        Rs_Temp.Open str, Cn, adOpenStatic, adLockOptimistic, adCmdText
        
        If Rs_Temp.RecordCount > 0 Then
                For k = 0 To Rs_Temp.RecordCount - 1
                         Set rsVendor = New ADODB.Recordset
                         rsVendor.Open "TblConfirmVacation_Details", Cn, adOpenStatic, adLockOptimistic, adCmdTable
                         rsVendor.AddNew
                         rsVendor("ID").value = CStr(new_id("TblConfirmVacation_Details", "ID", "", True))
                         rsVendor("HID").value = txtID.Text
                         rsVendor("SchoolFileID").value = IIf(IsNull(Rs_Temp("id").value), "", Rs_Temp("id").value)
                         rsVendor("DayCount").value = (DateDiff("d", FromDate.value, ToDate.value) + 1)
                         rsVendor("DayValue").value = val(txtdayvalue.Text)
                        rsVendor("FromDate") = IIf(IsNull(FromDate.value), Date, FromDate.value)
                        rsVendor("FromDateH") = IIf(IsNull(FromDateH.value), ToHijriDate(Date), FromDateH.value)
                        rsVendor("ToDate") = IIf(IsNull(ToDate.value), Date, ToDate.value)
                        rsVendor("ToDateH") = IIf(IsNull(ToDateH.value), ToHijriDate(Date), ToDateH.value)
                         
                         rsVendor.update
                         Rs_Temp.MoveNext
               Next
        End If
        
        
        Cn.CommitTrans
        
        BeginTrans = False
        XPTxtCurrent.Caption = rs.AbsolutePosition
        XPTxtCount.Caption = rs.RecordCount
       'CuurentLogdata

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

Private Sub Undo()
    On Error GoTo ErrTrap

    Select Case TxtModFlg.Text

        Case "N"
            clear_all Me
            Me.TxtModFlg.Text = "R"
            XPBtnMove_Click (1)

        Case "E"
            rs.find "ID='" & val(txtID.Text) & "'", , adSearchForward, adBookmarkFirst

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

    
        Msg = "”Ì „ Õ–› »Ì«‰«  «À»«  «· ⁄ÿ· —ﬁ„ " & CHR(13)
        Msg = Msg + (txtID.Text) & CHR(13)
        Msg = Msg + " Â·  —€» ›Ì Õ–› Â–Â «·»Ì«‰« ø"

        If MsgBox(Msg, vbYesNo + vbQuestion + vbMsgBoxRight + vbMsgBoxRtlReading, App.title) = vbYes Then
            
            If Not rs.RecordCount < 1 Then
                StrSQL = "delete From TblConfirmVacation where  ID =" & val(txtID.Text)
                Cn.Execute StrSQL, , adExecuteNoRecords
                
                  StrSQL = "delete From TblConfirmVacation_Details where  HID =" & val(txtID.Text)
                  Cn.Execute StrSQL, , adExecuteNoRecords
                
                   StrSQL = "SELECT  *  From TblConfirmVacation  "
                   Set rs = New ADODB.Recordset
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
    Msg = "·«Ì„ﬂ‰ Õ–› Â–« «·”Ã· · ﬂ«„· «·»Ì«‰«  " & CHR(13) & " ÊÃœ »Ì«‰«  „— »ÿ… »Â–« «· ⁄ÿ· "
    Msg = Msg & CHR(13) & Err.description
    MsgBox Msg, vbMsgBoxRight + vbMsgBoxRtlReading + vbExclamation, App.title
    rs.CancelUpdate
    'End If
End Sub


Private Sub AddTip()
    Dim Wrap As String
    On Error GoTo ErrTrap
    Set TTP = New clstooltip
    Wrap = CHR(13) + CHR(10)

    With TTP
        .Create Me.hwnd, "»Ì«‰«  «À»«  «· ⁄ÿ· ", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
        .AddControl Cmd(0), "ÃœÌœ ..." & Wrap & "·«÷«›… »Ì«‰«  «À»«  «· ⁄ÿ·  ÃœÌœ" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
    End With

    With TTP
        .Create Me.hwnd, "»Ì«‰«  «À»«  «· ⁄ÿ·  ", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
        .AddControl Cmd(1), " ⁄œÌ· ..." & Wrap & "· ⁄œÌ· »Ì«‰«  «À»«  «· ⁄ÿ· " & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
    End With

    With TTP
        .Create Me.hwnd, "»Ì«‰«  «À»«  «· ⁄ÿ·  ", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
        .AddControl Cmd(2), "Õ›Ÿ ..." & Wrap & "·Õ›Ÿ »Ì«‰«  «À»«  «· ⁄ÿ·  «·ÃœÌœ" & Wrap & "·Õ›Ÿ «· ⁄œÌ·« " & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
    End With

    With TTP
        .Create Me.hwnd, "»Ì«‰«  «À»«  «· ⁄ÿ· ", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
        .AddControl Cmd(3), " —«Ã⁄ ..." & Wrap & "·· —«Ã⁄ ⁄‰ ⁄„·Ì… «·«÷«›…" & Wrap & "··· —«Ã⁄ ⁄‰ ⁄„·Ì… «· ⁄œÌ·" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
    End With

    With TTP
        .Create Me.hwnd, "»Ì«‰«  «À»«  «· ⁄ÿ· ", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
        .AddControl Cmd(4), "Õ–› ..." & Wrap & "·Õ–› »Ì«‰«  Â–« «À»«  «· ⁄ÿ· " & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
    End With

    With TTP
        .Create Me.hwnd, "»Ì«‰« «À»«  «· ⁄ÿ· ", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
        .AddControl Cmd(5), "»ÕÀ ..." & Wrap & "···»ÕÀ ⁄‰ «À»«  «· ⁄ÿ· " & Wrap & "Ì‰ÿ»ﬁ ⁄·ÌÂ ‘—Êÿ „⁄Ì‰…" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
    End With

    With TTP
        .Create Me.hwnd, "»Ì«‰«  «À»«  «· ⁄ÿ· ", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
        .AddControl Cmd(6), "Œ—ÊÃ ..." & Wrap & "·«€·«ﬁ Â–Â «·‰«›–…" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
    End With

    With TTP
        .Create Me.hwnd, "»Ì«‰«  «À»«  «· ⁄ÿ· ", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
        .AddControl XPBtnMove(1), "«·√Ê· ..." & Wrap & "··«‰ ﬁ«· «·Ï √Ê· ”Ã·" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
    End With

    With TTP
        .Create Me.hwnd, "»Ì«‰«  «· ⁄ÿ·", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
        .AddControl XPBtnMove(0), "«·”«»ﬁ ..." & Wrap & "··«‰ ﬁ«· «·Ï «·”Ã· «·”«»ﬁ" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
    End With

    With TTP
        .Create Me.hwnd, "»Ì«‰«  «À»«  «· ⁄ÿ· ", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
        .AddControl XPBtnMove(3), "«· «·Ì ..." & Wrap & "··«‰ ﬁ«· «·Ï «·”Ã· «· «·Ì" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
    End With

    With TTP
        .Create Me.hwnd, "»Ì«‰«  «À»«  «· ⁄ÿ· ", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
        .AddControl XPBtnMove(2), "«·√ŒÌ— ..." & Wrap & "··«‰ ﬁ«· «·Ï ¬Œ— ”Ã·" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
    End With

    With TTP
        .Create Me.hwnd, "»Ì«‰«   «À»«  «· ⁄ÿ· ", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
    '    .AddControl CmdHelp, "„”«⁄œ… ..." & Wrap & "·· ⁄—› ⁄·Ï ÊŸÌ›… Â–Â «·‰«›–…" & Wrap & "ÊﬂÌ›Ì… «· ⁄«„· „⁄Â«" & Wrap & "≈÷€ÿ Â‰«" & Wrap, True
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

Public Function ISAllowDeleteUpdateContract() As Boolean
Dim EntryCreated As Boolean
Dim str As String



str = str & "   SELECT H.IDAC, dbo.TblSchooleFile.ManagerialID, H.DurationID, DD.MonthID, dbo.TblExchangeRequest.EntryCreated"
str = str & "   FROM     dbo.TblAttributionContract AS H INNER JOIN"
str = str & "   dbo.TblVehicleAllocation_Details AS D ON H.IDAC = D.IDVA INNER JOIN"
str = str & "   dbo.TblAttributionInstallmentDivided AS DD ON D.ID = DD.DetailsID INNER JOIN"
str = str & "   dbo.TblSchooleFile ON D.SchoolFileID = dbo.TblSchooleFile.ID INNER JOIN"
str = str & "   dbo.TblManagerialArea ON dbo.TblSchooleFile.ManagerialID = dbo.TblManagerialArea.ID INNER JOIN"
str = str & "   dbo.TblExchangeRequest ON DD.REID = dbo.TblExchangeRequest.ID"
str = str & "   Where (dd.RE_Paid = 1) And (dbo.TblExchangeRequest.EntryCreated = 1)"

str = str & "   And TblSchooleFile.CityID = " & val(dcCity.BoundText) & " And dd.MonthID = " & val(dcMonth.BoundText) & " And H.DurationID = " & val(dcDuration.BoundText)

Set Rs_Temp2 = New ADODB.Recordset
Rs_Temp2.Open str, Cn, adOpenStatic, adLockOptimistic
If Rs_Temp2.RecordCount > 0 Then
' EntryCreated = IIf(IsNull(Rs_Temp2("EntryCreated").value), 0, Rs_Temp2("EntryCreated").value)
' If EntryCreated = 0 Then '·„ Ì „ «‰‘«¡ ·=«·ﬁÌœ
        ISAllowDeleteUpdateContract = False
        Exit Function
  Else
     ISAllowDeleteUpdateContract = True
        Exit Function
  End If
        
'End If

ISAllowDeleteUpdateContract = True

End Function





Function print_report(Optional NoteSerial As Integer)
     
    Dim MySQL As String
    Dim RsData As New ADODB.Recordset
    Dim xApp As New CRAXDRT.Application
    Dim xReport As CRAXDRT.Report
    Dim CViewer As ClsReportViewer
    Dim StrReportTitle As String
    Dim StrFileName As String
    Dim Msg As String
            
MySQL = MySQL & "   SELECT dbo.TblCountriesGovernments.GovernmentName, dbo.TblConfirmVacation.CityID, dbo.TblConfirmVacation.DurationID, dbo.TblConfirmVacation.VacationTypeID,"
MySQL = MySQL & "                     dbo.TblConfirmVacation.Remarks, dbo.TblConfirmVacation.FromDate, dbo.TblConfirmVacation.ToDate, dbo.TblConfirmVacation.FromDateH,"
MySQL = MySQL & "                     dbo.TblConfirmVacation.ToDateH, dbo.TblConfirmVacation.UserID, dbo.TblConfirmVacation.MonthID, dbo.TblConfirmVacation.DayValue,"
MySQL = MySQL & "                     dbo.TblConfirmVacation.MangerialAreaID, dbo.TblManagerialArea.Name AS MAName, dbo.TblDurations.Name AS DurName, dbo.TblDurations_Details.Name AS MonthName,"
MySQL = MySQL & "                     dbo.TblConfirmVacation.ID, dbo.TblVacationTypes.Name AS TypeName"
MySQL = MySQL & "   FROM     dbo.TblConfirmVacation INNER JOIN"
MySQL = MySQL & "                     dbo.TblDurations ON dbo.TblConfirmVacation.DurationID = dbo.TblDurations.ID INNER JOIN"
MySQL = MySQL & "                     dbo.TblDurations_Details ON dbo.TblConfirmVacation.MonthID = dbo.TblDurations_Details.ID INNER JOIN"
MySQL = MySQL & "                     dbo.TblManagerialArea ON dbo.TblConfirmVacation.MangerialAreaID = dbo.TblManagerialArea.ID INNER JOIN"
MySQL = MySQL & "                     dbo.TblCountriesGovernments ON dbo.TblConfirmVacation.CityID = dbo.TblCountriesGovernments.GovernmentID INNER JOIN"
MySQL = MySQL & "                     dbo.TblVacationTypes ON dbo.TblConfirmVacation.VacationTypeID = dbo.TblVacationTypes.ID"


  MySQL = MySQL & "   where  TblConfirmVacation.id  = " & val(txtID.Text)
     MySQL = MySQL & "  order by TblConfirmVacation.id "
     
    If SystemOptions.UserInterface = ArabicInterface Then
        StrFileName = App.path & "\Reports\REPORTS NEW\" & "rpt_VacationReceipt.rpt"
    Else
        StrFileName = App.path & "\Reports\REPORTS NEW\" & "rpt_VacationReceipt.rpt"
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

