VERSION 5.00
Object = "{0AFE7BE0-11B7-4A3E-978D-D4501E9A57FE}#1.0#0"; "c1sizer.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{FE5DCFAD-BC1D-11D2-94CF-004005455FAA}#1.4#0"; "ImpulseButton.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form Frm_JournalSearch 
   BackColor       =   &H00E2E9E9&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "«·»ÕÀ ⁄‰ ﬁÌÊœ «·ÌÊ„Ì…"
   ClientHeight    =   6765
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6315
   Icon            =   "Frm_JournalSearch.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   RightToLeft     =   -1  'True
   ScaleHeight     =   6765
   ScaleWidth      =   6315
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   8
      Charset         =   178
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Begin VB.OptionButton OptType 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   "≈Œ Ì«— √‰Ê«⁄ «·ﬁÌÊœ"
      Height          =   315
      Index           =   1
      Left            =   90
      RightToLeft     =   -1  'True
      TabIndex        =   25
      Top             =   690
      Width           =   2895
   End
   Begin VB.OptionButton OptType 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   "Ã„Ì⁄ √‰Ê«⁄ «·ﬁÌÊœ"
      Height          =   315
      Index           =   0
      Left            =   3030
      RightToLeft     =   -1  'True
      TabIndex        =   24
      Top             =   690
      Width           =   3195
   End
   Begin C1SizerLibCtl.C1Elastic EleOptions 
      Height          =   3405
      Index           =   3
      Left            =   60
      TabIndex        =   12
      TabStop         =   0   'False
      Top             =   1050
      Width           =   6195
      _cx             =   10927
      _cy             =   6006
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
      Begin VB.CheckBox ChkMain 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "’Ì«‰…"
         Height          =   255
         Left            =   3420
         RightToLeft     =   -1  'True
         TabIndex        =   40
         Top             =   2940
         Width           =   2625
      End
      Begin VB.CheckBox ChkNotes 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "Œ’Ê„«  „ﬂ ”»…"
         Height          =   225
         Index           =   2
         Left            =   600
         RightToLeft     =   -1  'True
         TabIndex        =   23
         Tag             =   "10"
         Top             =   1260
         Width           =   2265
      End
      Begin VB.Frame Fra 
         BackColor       =   &H00E2E9E9&
         Caption         =   "„⁄«„·«  „«·Ì…"
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
         Height          =   2985
         Index           =   1
         Left            =   60
         RightToLeft     =   -1  'True
         TabIndex        =   18
         Top             =   30
         Width           =   2895
         Begin VB.CheckBox ChkNotes 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "⁄Ã“ ›Ï ‰ﬁœÌ… «·Œ“‰…"
            Height          =   225
            Index           =   10
            Left            =   540
            RightToLeft     =   -1  'True
            TabIndex        =   39
            Tag             =   "12"
            Top             =   2670
            Width           =   2265
         End
         Begin VB.CheckBox ChkNotes 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "“Ì«œ… ›Ï ‰ﬁœÌ… «·Œ“‰…"
            Height          =   225
            Index           =   9
            Left            =   540
            RightToLeft     =   -1  'True
            TabIndex        =   38
            Tag             =   "11"
            Top             =   2430
            Width           =   2265
         End
         Begin VB.CheckBox ChkNotes 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "”Õ» „‰ «·Œ“‰…"
            Height          =   225
            Index           =   8
            Left            =   540
            RightToLeft     =   -1  'True
            TabIndex        =   35
            Tag             =   "8"
            Top             =   2190
            Width           =   2265
         End
         Begin VB.CheckBox ChkNotes 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "≈Ìœ«⁄ ›Ï «·Œ“‰…"
            Height          =   225
            Index           =   7
            Left            =   540
            RightToLeft     =   -1  'True
            TabIndex        =   34
            Tag             =   "7"
            Top             =   1950
            Width           =   2265
         End
         Begin VB.CheckBox ChkNotes 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   " Õ’Ì· Ê”œ«œ «·‘Ìﬂ« "
            Height          =   225
            Index           =   6
            Left            =   540
            RightToLeft     =   -1  'True
            TabIndex        =   33
            Tag             =   "25"
            Top             =   1710
            Width           =   2265
         End
         Begin VB.CheckBox ChkNotes 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   " Õ’Ì· Ê”œ«œ «·√ﬁ”«ÿ"
            Height          =   225
            Index           =   5
            Left            =   540
            RightToLeft     =   -1  'True
            TabIndex        =   32
            Tag             =   "25"
            Top             =   1470
            Width           =   2265
         End
         Begin VB.CheckBox ChkNotes 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "Œ’Ê„«  „”„ÊÕ… "
            Height          =   225
            Index           =   4
            Left            =   540
            RightToLeft     =   -1  'True
            TabIndex        =   31
            Tag             =   "9"
            Top             =   990
            Width           =   2265
         End
         Begin VB.CheckBox ChkNotes 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "„’—Ê›« "
            Height          =   225
            Index           =   3
            Left            =   540
            RightToLeft     =   -1  'True
            TabIndex        =   27
            Tag             =   "3"
            Top             =   270
            Width           =   2265
         End
         Begin VB.CheckBox ChkNotes 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "„œ›Ê⁄« "
            Height          =   225
            Index           =   1
            Left            =   540
            RightToLeft     =   -1  'True
            TabIndex        =   22
            Tag             =   "5"
            Top             =   750
            Width           =   2265
         End
         Begin VB.CheckBox ChkNotes 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "„ﬁ»Ê÷« "
            Height          =   225
            Index           =   0
            Left            =   540
            RightToLeft     =   -1  'True
            TabIndex        =   19
            Tag             =   "4"
            Top             =   510
            Width           =   2265
         End
      End
      Begin VB.Frame Fra 
         BackColor       =   &H00E2E9E9&
         Caption         =   "Õ—ﬂ«   Ã«—Ì…"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   2805
         Index           =   0
         Left            =   2970
         RightToLeft     =   -1  'True
         TabIndex        =   13
         Top             =   30
         Width           =   3165
         Begin VB.CheckBox ChkTrans 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   " ”ÊÌ… «·„Œ“Ê‰(»«·‰ﬁ’)"
            Height          =   225
            Index           =   9
            Left            =   480
            RightToLeft     =   -1  'True
            TabIndex        =   37
            Tag             =   "15"
            Top             =   2520
            Width           =   2595
         End
         Begin VB.CheckBox ChkTrans 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   " ”ÊÌ… «·„Œ“Ê‰(»«·“Ì«œ…)"
            Height          =   225
            Index           =   8
            Left            =   480
            RightToLeft     =   -1  'True
            TabIndex        =   36
            Tag             =   "16"
            Top             =   2280
            Width           =   2595
         End
         Begin VB.CheckBox ChkTrans 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "’—› »÷«⁄… („”ÕÊ»«  ‘Œ’Ì…)"
            Height          =   225
            Index           =   7
            Left            =   480
            RightToLeft     =   -1  'True
            TabIndex        =   30
            Tag             =   "17"
            Top             =   2040
            Width           =   2595
         End
         Begin VB.CheckBox ChkTrans 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "’—› »÷«⁄…(·“Ê„ «·⁄„·)"
            Height          =   255
            Index           =   6
            Left            =   480
            RightToLeft     =   -1  'True
            TabIndex        =   29
            Tag             =   "18"
            Top             =   1770
            Width           =   2595
         End
         Begin VB.CheckBox ChkTrans 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "’—› »÷«⁄…( ·›Ì« )"
            Height          =   255
            Index           =   5
            Left            =   480
            RightToLeft     =   -1  'True
            TabIndex        =   28
            Tag             =   "8"
            Top             =   1500
            Width           =   2595
         End
         Begin VB.CheckBox ChkTrans 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "«·—’Ìœ «·≈›  «ÕÌ"
            Height          =   225
            Index           =   4
            Left            =   480
            RightToLeft     =   -1  'True
            TabIndex        =   26
            Tag             =   "3"
            Top             =   1260
            Width           =   2595
         End
         Begin VB.CheckBox ChkTrans 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "›Ê« Ì— „— Ã⁄«  «·„‘ —Ì« "
            Height          =   225
            Index           =   3
            Left            =   480
            RightToLeft     =   -1  'True
            TabIndex        =   17
            Tag             =   "5"
            Top             =   1020
            Width           =   2595
         End
         Begin VB.CheckBox ChkTrans 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "›Ê« Ì— „— Ã⁄«  «·„»Ì⁄« "
            Height          =   255
            Index           =   2
            Left            =   480
            RightToLeft     =   -1  'True
            TabIndex        =   16
            Tag             =   "9"
            Top             =   750
            Width           =   2595
         End
         Begin VB.CheckBox ChkTrans 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "›Ê« Ì— «·‘—«¡"
            Height          =   225
            Index           =   1
            Left            =   480
            RightToLeft     =   -1  'True
            TabIndex        =   15
            Tag             =   "1"
            Top             =   510
            Width           =   2595
         End
         Begin VB.CheckBox ChkTrans 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "›Ê« Ì— «·»Ì⁄"
            Height          =   225
            Index           =   0
            Left            =   480
            RightToLeft     =   -1  'True
            TabIndex        =   14
            Tag             =   "2"
            Top             =   270
            Width           =   2595
         End
      End
   End
   Begin C1SizerLibCtl.C1Elastic EleHeader 
      Height          =   645
      Left            =   0
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   0
      Width           =   6315
      _cx             =   11139
      _cy             =   1138
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial (Arabic)"
         Size            =   14.25
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
      BackColor       =   16777215
      ForeColor       =   4210688
      FloodColor      =   6553600
      ForeColorDisabled=   -2147483631
      Picture         =   "Frm_JournalSearch.frx":08CA
      Caption         =   "«·»ÕÀ ⁄‰ ﬁÌÊœ «·ÌÊ„Ì…"
      Align           =   1
      AutoSizeChildren=   0
      BorderWidth     =   6
      ChildSpacing    =   4
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
      FrameColor      =   4210688
      FrameShadow     =   -2147483632
      FloodStyle      =   1
      _GridInfo       =   ""
      AccessibleName  =   ""
      AccessibleDescription=   ""
      AccessibleValue =   ""
      AccessibleRole  =   9
   End
   Begin C1SizerLibCtl.C1Elastic EleOptions 
      Height          =   1755
      Index           =   0
      Left            =   30
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   4470
      Width           =   6225
      _cx             =   10980
      _cy             =   3096
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
      ForeColor       =   128
      FloodColor      =   6553600
      ForeColorDisabled=   -2147483631
      Caption         =   "ŒÌ«—«  √Œ—Ï"
      Align           =   0
      AutoSizeChildren=   0
      BorderWidth     =   2
      ChildSpacing    =   1
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
      Frame           =   4
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
      Begin C1SizerLibCtl.C1Elastic EleOptions 
         Height          =   1035
         Index           =   1
         Left            =   660
         TabIndex        =   41
         TabStop         =   0   'False
         Top             =   210
         Width           =   3165
         _cx             =   5583
         _cy             =   1826
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
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
         ForeColor       =   128
         FloodColor      =   6553600
         ForeColorDisabled=   -2147483631
         Caption         =   "Õ«·… «·ﬁÌœ"
         Align           =   0
         AutoSizeChildren=   0
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
            Caption         =   "«·ﬂ·"
            Height          =   315
            Index           =   2
            Left            =   60
            RightToLeft     =   -1  'True
            TabIndex        =   44
            Top             =   420
            Value           =   -1  'True
            Width           =   945
         End
         Begin VB.OptionButton Opt 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "€Ì— „—Õ·"
            Height          =   315
            Index           =   1
            Left            =   1020
            RightToLeft     =   -1  'True
            TabIndex        =   43
            Top             =   420
            Width           =   1095
         End
         Begin VB.OptionButton Opt 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "„—Õ·"
            Height          =   315
            Index           =   0
            Left            =   2160
            RightToLeft     =   -1  'True
            TabIndex        =   42
            Top             =   420
            Width           =   945
         End
      End
      Begin C1SizerLibCtl.C1Elastic EleOptions 
         Height          =   1035
         Index           =   2
         Left            =   3840
         TabIndex        =   7
         TabStop         =   0   'False
         Top             =   210
         Width           =   2295
         _cx             =   4048
         _cy             =   1826
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
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
         ForeColor       =   128
         FloodColor      =   6553600
         ForeColorDisabled=   -2147483631
         Caption         =   " «—ÌŒ  Õ—Ì— «·ﬁÌœ"
         Align           =   0
         AutoSizeChildren=   0
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
         Begin MSComCtl2.DTPicker DTPDev_From 
            Height          =   345
            Left            =   90
            TabIndex        =   8
            Top             =   240
            Width           =   1635
            _ExtentX        =   2884
            _ExtentY        =   609
            _Version        =   393216
            CalendarBackColor=   12648447
            CalendarTitleBackColor=   10383715
            CheckBox        =   -1  'True
            CustomFormat    =   "yyyy/M/d"
            Format          =   92274691
            CurrentDate     =   37773
         End
         Begin MSComCtl2.DTPicker DTPDEV_TO 
            Height          =   345
            Left            =   90
            TabIndex        =   9
            Top             =   630
            Width           =   1635
            _ExtentX        =   2884
            _ExtentY        =   609
            _Version        =   393216
            CalendarBackColor=   12648447
            CalendarTitleBackColor=   10383715
            CheckBox        =   -1  'True
            CustomFormat    =   "yyyy/M/d"
            Format          =   92274691
            CurrentDate     =   37773
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "≈·Ï "
            Height          =   345
            Index           =   3
            Left            =   1770
            RightToLeft     =   -1  'True
            TabIndex        =   11
            Top             =   660
            Width           =   405
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "„‰"
            Height          =   345
            Index           =   0
            Left            =   1770
            RightToLeft     =   -1  'True
            TabIndex        =   10
            Top             =   270
            Width           =   405
         End
      End
      Begin MSDataListLib.DataCombo DcboUsers 
         Height          =   315
         Index           =   0
         Left            =   3150
         TabIndex        =   2
         Top             =   1320
         Width           =   1845
         _ExtentX        =   3254
         _ExtentY        =   556
         _Version        =   393216
         BackColor       =   -2147483624
         Text            =   ""
         RightToLeft     =   -1  'True
      End
      Begin MSDataListLib.DataCombo DcboUsers 
         Height          =   315
         Index           =   1
         Left            =   120
         TabIndex        =   21
         Top             =   1320
         Width           =   1905
         _ExtentX        =   3360
         _ExtentY        =   556
         _Version        =   393216
         Enabled         =   0   'False
         BackColor       =   -2147483624
         Text            =   ""
         RightToLeft     =   -1  'True
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "«·„” Œœ„ «·ﬁ«∆„ »«· —ÕÌ·"
         Height          =   405
         Index           =   2
         Left            =   2040
         RightToLeft     =   -1  'True
         TabIndex        =   20
         Top             =   1260
         Width           =   1065
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "«·„” Œœ„ «·„Õ——"
         Height          =   255
         Index           =   1
         Left            =   4980
         RightToLeft     =   -1  'True
         TabIndex        =   1
         Top             =   1320
         Width           =   1155
      End
   End
   Begin ImpulseButton.ISButton Cmd_Exit 
      Cancel          =   -1  'True
      Height          =   345
      Left            =   1230
      TabIndex        =   4
      Top             =   6330
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   609
      ButtonStyle     =   1
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
      ColorHoverText  =   0
      DrawFocusRectangle=   0   'False
      ColorToggledHoverText=   0
   End
   Begin ImpulseButton.ISButton Cmd_Get 
      Height          =   345
      Left            =   2250
      TabIndex        =   5
      Top             =   6330
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   609
      ButtonStyle     =   1
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
      ColorHoverText  =   0
      DrawFocusRectangle=   0   'False
      ColorToggledHoverText=   0
   End
   Begin ImpulseButton.ISButton Cmd_Clear 
      Height          =   345
      Left            =   3270
      TabIndex        =   6
      Top             =   6330
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   609
      ButtonStyle     =   1
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
      ColorButton     =   14871017
      ColorHighlight  =   16777215
      ColorHoverText  =   0
      DrawFocusRectangle=   0   'False
      ColorToggledHoverText=   0
   End
End
Attribute VB_Name = "Frm_JournalSearch"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
Dim RsLoadJournal   As New ADODB.Recordset
Dim TTP As New clstooltip
Dim IntCounter      As Long

Private Sub cmd_clear_Click()

    Me.DTPDev_From.value = Date
    Me.DTPDEV_TO.value = Date
    Me.Opt(2).value = True
    Me.DcboUsers(0).BoundText = ""

End Sub

Private Sub Cmd_Clear_MouseEnter()

    cmd_clear.FontBold = True

End Sub

Private Sub Cmd_Clear_MouseLeave()
    cmd_clear.FontBold = False
End Sub

Private Sub Cmd_exit_Click()
    Unload Me
End Sub

Private Sub Cmd_Exit_MouseEnter()
    Cmd_Exit.FontBold = True
End Sub

Private Sub Cmd_Exit_MouseLeave()
    Cmd_Exit.FontBold = False
End Sub

Private Sub Cmd_Get_Click()
    Dim StrSQL As String
    StrSQL = StrBuild

    If StrSQL = "" Then
        Exit Sub
    End If

    Frm_General_Journal.Retrive StrSQL
    Frm_General_Journal.DTPDev_From.value = Me.DTPDev_From.value
    Frm_General_Journal.DTPDEV_TO.value = Me.DTPDEV_TO.value
End Sub

Private Sub Cmd_Get_MouseEnter()
    Cmd_Get.FontBold = True
End Sub

Private Sub Cmd_Get_MouseLeave()
    Cmd_Get.FontBold = False
End Sub

Private Sub Form_Load()
    Dim Dcombos As New ClsDataCombos

    SetDtpickerDate Me.DTPDev_From
    SetDtpickerDate Me.DTPDEV_TO

    Dcombos.GetUsers DcboUsers(0)
    Dcombos.GetUsers DcboUsers(1)

    If SystemOptions.UserInterface = EnglishInterface Then
        SetInterface Me
        ChangeLang
    End If

    OptType(0).value = True
    OptType_Click (0)
    AddTip
    CenterForm Me

    FormPostion Me, GetPostion
End Sub

Private Sub AddTip()
    Dim Msg As String

    Dim Wrap As String
    Wrap = Chr(13) + Chr(10)

    If SystemOptions.UserInterface = ArabicInterface Then

        With TTP
            .Create Me.hwnd, OptType(0).Caption, 1, 15204351, -2147483630
            .MaxWidth = 4000
            .VisibleTime = 10000
            .DelayTime = 300
            Msg = "Â–« «·ŒÌ«— ÌﬁÊ„ »≈” —Ã«⁄ Ã„Ì⁄ √‰Ê«⁄ «·ﬁÌÊœ" & Wrap & "«·„”Ã·…. "
            .AddControl OptType(0), Msg, True
        End With

        With TTP
            .Create Me.hwnd, OptType(1).Caption, 1, 15204351, -2147483630
            .MaxWidth = 4000
            .VisibleTime = 10000
            .DelayTime = 300
            Msg = "Â–« «·ŒÌ«— Ì ÌÕ ·ﬂ ≈Œ Ì«— «·ﬁÌÊœ«·„”Ã·…" & Wrap & "«· Ï  —Ìœ ≈” —Ã«⁄Â« „À·«:-" & Wrap & "ﬁÌÊœ ›Ê« Ì— «·„»Ì⁄« -ﬁÌÊœ ›Ê« Ì— «·„‘ —Ì« " & Wrap & "«Ê „— Ã⁄«  «·„»Ì⁄«  «Ê „— Ã⁄«  «·„‘ —Ì«    " & Wrap & "...«Ê Õ Ï «·ﬁÌÊœ «·„”Ã·… „‰ ﬁ»· «·„” Œœ„" & Wrap & "›Ï ‘«‘…  Õ—Ì— ﬁÌÊœ «·ÌÊ„Ì…."
            .AddControl OptType(1), Msg, True
        End With

        '
        With TTP
            .Create Me.hwnd, ChkTrans(0).Caption, 1, 15204351, -2147483630
            .MaxWidth = 4000
            .VisibleTime = 10000
            .DelayTime = 300
            Msg = "Â–« «·ŒÌ«— ÌﬁÊ„ »≈” —Ã«⁄ «·ﬁÌÊœ«·„”Ã·…" & Wrap & "⁄·Ï ›Ê« Ì— «·»Ì⁄ ( «·„”Ã·… „‰ ‘«‘… " & Wrap & "›« Ê—… «·»Ì⁄ )."
            .AddControl ChkTrans(0), Msg, True
        End With

        With TTP
            .Create Me.hwnd, ChkTrans(1).Caption, 1, 15204351, -2147483630
            .MaxWidth = 4000
            .VisibleTime = 10000
            .DelayTime = 300
            Msg = "Â–« «·ŒÌ«— ÌﬁÊ„ »≈” —Ã«⁄ «·ﬁÌÊœ«·„”Ã·…" & Wrap & "⁄·Ï ›Ê« Ì— «·‘—«¡ ( «·„”Ã·… „‰ ‘«‘… " & Wrap & "›« Ê—… «·‘—«¡ )."
            .AddControl ChkTrans(1), Msg, True
        End With

        With TTP
            .Create Me.hwnd, ChkTrans(2).Caption, 1, 15204351, -2147483630
            .MaxWidth = 4000
            .VisibleTime = 10000
            .DelayTime = 300
            Msg = "Â–« «·ŒÌ«— ÌﬁÊ„ »≈” —Ã«⁄ «·ﬁÌÊœ«·„”Ã·…" & Wrap & "⁄·Ï ›Ê« Ì— „— Ã⁄ «·„»Ì⁄«  ( «·„”Ã·… " & Wrap & "„‰ ‘«‘… „— Ã⁄ «·„»Ì⁄«  )."
            .AddControl ChkTrans(2), Msg, True
        End With

        With TTP
            .Create Me.hwnd, ChkTrans(3).Caption, 1, 15204351, -2147483630
            .MaxWidth = 4000
            .VisibleTime = 10000
            .DelayTime = 300
            Msg = "Â–« «·ŒÌ«— ÌﬁÊ„ »≈” —Ã«⁄ «·ﬁÌÊœ«·„”Ã·…" & Wrap & "⁄·Ï ›Ê« Ì— „— Ã⁄ «·„‘ —Ì«  ( «·„”Ã·… " & Wrap & "„‰ ‘«‘… „— Ã⁄ «·„‘ —Ì«  )."
            .AddControl ChkTrans(3), Msg, True
        End With

        With TTP
            .Create Me.hwnd, ChkNotes(0).Caption, 1, 15204351, -2147483630
            .MaxWidth = 4000
            .VisibleTime = 10000
            .DelayTime = 300
            Msg = "Â–« «·ŒÌ«— ÌﬁÊ„ »≈” —Ã«⁄ «·ﬁÌÊœ«·„”Ã·…" & Wrap & "⁄·Ï «·„ﬁ»Ê÷«  ( «· Ï ”Ã·  „‰ ‘«‘… " & Wrap & "«·„ﬁ»Ê÷«  )."
            .AddControl ChkNotes(0), Msg, True
        End With

        With TTP
            .Create Me.hwnd, ChkNotes(1).Caption, 1, 15204351, -2147483630
            .MaxWidth = 4000
            .VisibleTime = 10000
            .DelayTime = 300
            Msg = "Â–« «·ŒÌ«— ÌﬁÊ„ »≈” —Ã«⁄ «·ﬁÌÊœ«·„”Ã·…" & Wrap & "⁄·Ï «·„œ›Ê⁄«  ( «· Ï ”Ã·  „‰ ‘«‘… " & Wrap & "«·„œ›Ê⁄«  )."
            .AddControl ChkNotes(1), Msg, True
        End With

        With TTP
            .Create Me.hwnd, ChkNotes(2).Caption, 1, 15204351, -2147483630
            .MaxWidth = 4000
            .VisibleTime = 10000
            .DelayTime = 300
            Msg = "Â–« «·ŒÌ«— ÌﬁÊ„ »≈” —Ã«⁄ «·ﬁÌÊœ«·„”Ã·…" & Wrap & "»Ê«”ÿ… «·„” Œœ„ ( «· Ï ”Ã·  „‰ ‘«‘… " & Wrap & " Õ—Ì— ﬁÌÊœ «·ÌÊ„Ì… )."
            .AddControl ChkNotes(2), Msg, True
        End With

        With TTP
            .Create Me.hwnd, "»ÕÀ", 1, 15204351, -2147483630
            .MaxWidth = 4000
            .VisibleTime = 10000
            .DelayTime = 300
            .AddControl Cmd_Get, "»œ¡ «·»ÕÀ «·ﬁÌÊœ ..." & Wrap & "≈÷€ÿ Â‰« ·»œ¡ ⁄„·Ì… «·»ÕÀ ⁄‰ «·ﬁÌÊœ " & Wrap & "ÿ»ﬁ« ··‘—Êÿ «· Ï Õœœ Â«.  ", True
        End With

        With TTP
            .Create Me.hwnd, "„”Õ", 1, 15204351, -2147483630
            .MaxWidth = 4000
            .VisibleTime = 10000
            .DelayTime = 300
            .AddControl cmd_clear, "„”Õ «·‘—Êÿ «·„Õœœ… ..." & Wrap & "„”Õ «·‘—Êÿ «·„⁄—Ê÷… ··»œ¡ ›Ï " & Wrap & "⁄„·Ì… »ÕÀ ÃœÌœ….", True
        End With

        With TTP
            .Create Me.hwnd, "Œ—ÊÃ", 1, 15204351, -2147483630
            .MaxWidth = 4000
            .VisibleTime = 10000
            .DelayTime = 300
            .AddControl Cmd_Exit, "«·Œ—ÊÃ ..." & Wrap & "«·Œ—ÊÃ „‰ ‘«‘… «·»ÕÀ ⁄‰ ﬁÌÊœ «·ÌÊ„Ì… " & Wrap & "Ê«·⁄Êœ… ≈·Ï ‘«‘…  —ÕÌ· «·ÌÊ„Ì….", True
        End With

        With TTP
            .Create Me.hwnd, "Õ«·… «·ﬁÌœ(„—Õ·)", 1, 15204351, -2147483630
            .MaxWidth = 4000
            .VisibleTime = 10000
            .DelayTime = 300
            .AddControl Opt(0), "„—Õ· ..." & Wrap & "Ì÷Ì› ≈·Ï ‘—Êÿ «·»ÕÀ «‰ Ì „ " & Wrap & "«·»ÕÀ ⁄‰ ﬁÌÊœ «·ÌÊ„Ì… «·„—Õ·….", True
        End With

        With TTP
            .Create Me.hwnd, "Õ«·… «·ﬁÌœ(€Ì— „—Õ·)", 1, 15204351, -2147483630
            .MaxWidth = 4000
            .VisibleTime = 10000
            .DelayTime = 300
            .AddControl Opt(1), " €Ì— „—Õ· ..." & Wrap & "Ì÷Ì› ≈·Ï ‘—Êÿ «·»ÕÀ «‰ Ì „ " & Wrap & "«·»ÕÀ ⁄‰ ﬁÌÊœ «·ÌÊ„Ì…«·€Ì— „—Õ·….", True
        End With

        With TTP
            .Create Me.hwnd, "Õ«·… «·ﬁÌœ(«·ﬂ·)", 1, 15204351, -2147483630
            .MaxWidth = 4000
            .VisibleTime = 10000
            .DelayTime = 300
            .AddControl Opt(2), " «·ﬂ· ..." & Wrap & "Ì÷Ì› ≈·Ï ‘—Êÿ «·»ÕÀ «‰ Ì „ " & Wrap & "«·»ÕÀ ⁄‰ Ã„Ì⁄ ﬁÌÊœ «·ÌÊ„Ì…" & Wrap & "”Ê«¡ ﬂ«‰   ·ﬂ «·ﬁÌÊœ „—Õ·… «„ €Ì— „—Õ·…", True
        End With

        With TTP
            .Create Me.hwnd, " «—ÌŒ  Õ—Ì— «·ﬁÌœ(»œ«Ì…)", 1, 15204351, -2147483630
            .MaxWidth = 4000
            .VisibleTime = 10000
            .DelayTime = 300
            .AddControl DTPDev_From, " «—ÌŒ »œ«Ì… «·› —… ..." & Wrap & "» ›⁄Ì· ⁄·«„… «·√Œ Ì«— Ì „ ≈” —Ã«⁄ " & Wrap & "ﬁÌÊœ «·ÌÊ„Ì… «·„Õ——… »œ«Ì… „‰ Â–« «· «—ÌŒ" & Wrap & "Ê⁄œ„  ›⁄Ì· ⁄·«„… «·√Œ Ì«— Ì „ ≈” —Ã«⁄ «·ﬁÌÊœ" & Wrap & "«·„Õ——… „‰ »œ«Ì… ⁄„· «·»—‰«„Ã", True
        End With

        With TTP
            .Create Me.hwnd, " «—ÌŒ  Õ—Ì— «·ﬁÌœ(‰Â«Ì…)", 1, 15204351, -2147483630
            .MaxWidth = 4000
            .VisibleTime = 10000
            .DelayTime = 300
            .AddControl DTPDEV_TO, " «—ÌŒ ‰Â«Ì… «·› —… ..." & Wrap & "» ›⁄Ì· ⁄·«„… «·√Œ Ì«— Ì „ ≈” —Ã«⁄ " & Wrap & "ﬁÌÊœ «·ÌÊ„Ì… «·„Õ——… Õ Ï Â–« «· «—ÌŒ" & Wrap & "Ê⁄œ„  ›⁄Ì· ⁄·«„… «·√Œ Ì«— Ì „ ≈” —Ã«⁄ «·ﬁÌÊœ" & Wrap & "«·„Õ——… Õ Ï «·√‰", True
        End With

        With TTP
            .Create Me.hwnd, "«·„” Œœ„ «·„Õ——", 1, 15204351, -2147483630
            .MaxWidth = 4000
            .VisibleTime = 10000
            .DelayTime = 300
            .AddControl DcboUsers(0), "«·„” Œœ„ «·„Õ——..." & Wrap & "»≈Œ Ì«— „” Œœ„ „⁄Ì‰  " & Wrap & "Ì÷«› ≈·Ï ‘—Êÿ «·»ÕÀ «‰  ﬂÊ‰ «·ﬁÌÊœ" & Wrap & "Õ——  «Ê ”Ã·  »Ê«”ÿ… Â–« «·„” Œœ„", True
        End With

    Else

        With TTP
            .Create Me.hwnd, OptType(0).Caption, 1, 15204351, -2147483630, False
            .MaxWidth = 4000
            .VisibleTime = 10000
            .DelayTime = 300
            Msg = "This Options allows you to retrive " & Wrap & "all issued journal. "
            .AddControl OptType(0), Msg, False
        End With

        With TTP
            .Create Me.hwnd, OptType(1).Caption, 1, 15204351, -2147483630, False
            .MaxWidth = 4000
            .VisibleTime = 10000
            .DelayTime = 300
            Msg = "This option allows you to customize which" & Wrap & "journal you will retrive like:-" & Wrap & "Sales Journal,Purchases journal" & Wrap & "Retrun Sales Journal,Retrun Purchase Journal" & Wrap & "...in addition you can retrive the manual" & Wrap & "journal( edit journal screen)."
            .AddControl OptType(1), Msg, False
        End With

        With TTP
            .Create Me.hwnd, ChkTrans(0).Caption, 1, 15204351, -2147483630, False
            .MaxWidth = 4000
            .VisibleTime = 10000
            .DelayTime = 300
            Msg = "This option allows you to retrive the" & Wrap & "journal on sales account. " & Wrap & "(Record From Bill Invoice Screen)."
            .AddControl ChkTrans(0), Msg, False
        End With

        With TTP
            .Create Me.hwnd, ChkTrans(1).Caption, 1, 15204351, -2147483630, False
            .MaxWidth = 4000
            .VisibleTime = 10000
            .DelayTime = 300
            Msg = "This option allows you to retrive the" & Wrap & "journal on purchase account. " & Wrap & "(Record From purchase Invoice Screen)."
            .AddControl ChkTrans(1), Msg, False
        End With

        With TTP
            .Create Me.hwnd, ChkTrans(2).Caption, 1, 15204351, -2147483630, False
            .MaxWidth = 4000
            .VisibleTime = 10000
            .DelayTime = 300
            Msg = "This option allows you to retrive the" & Wrap & "journal on retrun sales account. " & Wrap & "(Record From Retrun Sales Screen)."
            .AddControl ChkTrans(2), Msg, False
        End With

        With TTP
            .Create Me.hwnd, ChkTrans(3).Caption, 1, 15204351, -2147483630, False
            .MaxWidth = 4000
            .VisibleTime = 10000
            .DelayTime = 300
            Msg = "This option allows you to retrive the" & Wrap & "journal on retrun purchase account. " & Wrap & "(Record From Retrun purchase Screen)."
            .AddControl ChkTrans(3), Msg, False
        End With

        With TTP
            .Create Me.hwnd, ChkNotes(0).Caption, 1, 15204351, -2147483630, False
            .MaxWidth = 4000
            .VisibleTime = 10000
            .DelayTime = 300
            Msg = "This option allows you to retrive the" & Wrap & "journal on notes receivable account" & Wrap & "(Record From Notes Receivable Screen)."
            .AddControl ChkNotes(0), Msg, False
        End With

        With TTP
            .Create Me.hwnd, ChkNotes(1).Caption, 1, 15204351, -2147483630, False
            .MaxWidth = 4000
            .VisibleTime = 10000
            .DelayTime = 300
            Msg = "This option allows you to retrive the" & Wrap & "journal on notes payable account" & Wrap & "(Record From Notes Payable Screen)."
            .AddControl ChkNotes(1), Msg, False
        End With

        With TTP
            .Create Me.hwnd, ChkNotes(2).Caption, 1, 15204351, -2147483630, False
            .MaxWidth = 4000
            .VisibleTime = 10000
            .DelayTime = 300
            Msg = "This option allows you to retrive the" & Wrap & "manual edited journal" & Wrap & "(Record From Edit Journal Screen)."
            .AddControl ChkNotes(2), Msg, False
        End With

        With TTP
            .Create Me.hwnd, "Beginning Date", 1, 15204351, -2147483630, False
            .MaxWidth = 4000
            .VisibleTime = 10000
            .DelayTime = 300
            Msg = "When Check is enabled...search will" & Wrap & "start from this date."
            .AddControl DTPDev_From, Msg, False
        End With

        With TTP
            .Create Me.hwnd, "End Date", 1, 15204351, -2147483630, False
            .MaxWidth = 4000
            .VisibleTime = 10000
            .DelayTime = 300
            Msg = "When check is enabled...search will" & Wrap & "end to this date."
            .AddControl DTPDEV_TO, Msg, False
        End With

        With TTP
            .Create Me.hwnd, "Voucher State(Posted)", 1, 15204351, -2147483630, False
            .MaxWidth = 4000
            .VisibleTime = 10000
            .DelayTime = 300
            Msg = "This option when enabled... " & Wrap & "the search will retrive the " & Wrap & "posted voucheres only." & Wrap & "(POSTED JOURNAL)."
            .AddControl Opt(0), Msg, False
        End With

        With TTP
            .Create Me.hwnd, "Voucher State(Not Posted)", 1, 15204351, -2147483630, False
            .MaxWidth = 4000
            .VisibleTime = 10000
            .DelayTime = 300
            Msg = "This option when enabled... " & Wrap & "the search will retrive the " & Wrap & "not posted voucheres only." & Wrap & "(NOT POSTED JOURNAL)."
            .AddControl Opt(1), Msg, False
        End With

        With TTP
            .Create Me.hwnd, "Voucher State(All)", 1, 15204351, -2147483630, False
            .MaxWidth = 4000
            .VisibleTime = 10000
            .DelayTime = 300
            Msg = "This option when enabled... " & Wrap & "the search will retrive the " & Wrap & "All voucheres(ALL JOURNAL)."
            .AddControl Opt(2), Msg, False
        End With

        With TTP
            .Create Me.hwnd, "Issued By(Users)", 1, 15204351, -2147483630, False
            .MaxWidth = 4000
            .VisibleTime = 10000
            .DelayTime = 300
            Msg = "If you want to retrive the journal" & Wrap & "which issued by specific user. "
            .AddControl DcboUsers(0), Msg, False
        End With

        With TTP
            .Create Me.hwnd, "Posted By(Users)", 1, 15204351, -2147483630, False
            .MaxWidth = 4000
            .VisibleTime = 10000
            .DelayTime = 300
            Msg = "If you want to retrive the journal" & Wrap & "which posted by specific user. "
            .AddControl DcboUsers(1), Msg, False
        End With

        With TTP
            .Create Me.hwnd, "Clear", 1, 15204351, -2147483630, False
            .MaxWidth = 4000
            .VisibleTime = 10000
            .DelayTime = 300
            Msg = "Clear all search conditions" & Wrap & "to start a new search ."
            .AddControl cmd_clear, Msg, False
        End With

        With TTP
            .Create Me.hwnd, "Search", 1, 15204351, -2147483630, False
            .MaxWidth = 4000
            .VisibleTime = 10000
            .DelayTime = 300
            .AddControl Cmd_Get, "Click here to start search process.", False
        End With

        With TTP
            .Create Me.hwnd, "Exit", 1, 15204351, -2147483630, False
            .MaxWidth = 4000
            .VisibleTime = 10000
            .DelayTime = 300
            Msg = "Exit from search journal screen."
            .AddControl Cmd_Exit, Msg, False
        End With

    End If

End Sub

Private Sub ChangeLang()
    Me.Caption = "Journal Search"
    Me.EleHeader.Caption = Me.Caption
    OptType(0).Caption = "All Journal Types"
    OptType(1).Caption = "Choose Journal Types"
    Fra(0).Caption = "Inventory Transactions"
    ChkTrans(0).Caption = "Bill Invoices"
    ChkTrans(1).Caption = "Purchase Invoices"
    ChkTrans(2).Caption = "Retrun Sales Invoices"
    ChkTrans(3).Caption = "Retrun Purchase Invoices"
    Fra(1).Caption = "Financial Transactions"
    ChkNotes(0).Caption = "Notes Receivable"
    ChkNotes(1).Caption = "Notes Payable"
    ChkNotes(2).Caption = "Edited Journal"

    EleOptions(0).Caption = "More Search Options"
    EleOptions(1).Caption = "Posting State"
    Opt(0).Caption = "Posted"
    Opt(1).Caption = "Not Posted"
    Opt(2).Caption = "ALL"
    lbl(0).Caption = "From"
    lbl(3).Caption = "To"

    lbl(1).Caption = "Issued User"
    lbl(2).Caption = "Posted User"

    EleOptions(2).Caption = "Issued Date"
    cmd_clear.Caption = "&Clear"
    Cmd_Get.Caption = "&Search"
    Cmd_Exit.Caption = "E&xit"
End Sub

Private Sub Form_Unload(Cancel As Integer)

    If RsLoadJournal.State = adStateOpen Then
        RsLoadJournal.Close
    End If

    Set RsLoadJournal = Nothing
    Set TTP = Nothing

    FormPostion Me, SavePostion
End Sub

Private Sub Opt_Click(Index As Integer)
    DcboUsers(1).Enabled = Opt(0).value
End Sub

Private Sub OptType_Click(Index As Integer)
    Dim i As Integer

    For i = ChkTrans.LBound To ChkTrans.UBound
        ChkTrans(i).Enabled = OptType(1).value
    Next i

    For i = ChkNotes.LBound To ChkNotes.UBound
        ChkNotes(i).Enabled = OptType(1).value
    Next i

    Fra(0).Enabled = OptType(1).value
    Fra(1).Enabled = OptType(1).value
End Sub

Private Sub SetFgAccounts(StrWhere As String)
    Dim rs As New ADODB.Recordset
    Dim StrSQL As String
    Dim i As Integer
    Dim BolAccountRTL As Boolean
    'If SystemOptions.UserShowDataAccounts = ShowArabicData Then
    '    BolAccountRTL = True
    'ElseIf SystemOptions.UserShowDataAccounts = ShowEnglishData Then
    '    BolAccountRTL = False
    'Else
    '    BolAccountRTL = IIf(SystemOptions.UserInterface = ArabicInterface, True, False)
    'End If
    Exit Sub
    StrSQL = "SELECT Sum(IIf(DOUBLE_ENTREY_VOUCHERS.Credit_Or_Debit=0 " & "And NOTES.NotePosted=True,DOUBLE_ENTREY_VOUCHERS.Value,0)) AS DebitPosted," & "Sum(IIf(DOUBLE_ENTREY_VOUCHERS.Credit_Or_Debit=0 And NOTES.NotePosted=False," & "DOUBLE_ENTREY_VOUCHERS.Value,0)) AS DebitNotPosted," & "Sum(IIf(DOUBLE_ENTREY_VOUCHERS.Credit_Or_Debit=1 And NOTES.NotePosted=True," & "DOUBLE_ENTREY_VOUCHERS.Value,0)) AS CreditPosted," & "Sum(IIf(DOUBLE_ENTREY_VOUCHERS.Credit_Or_Debit=1 And notes.NotePosted=False," & "DOUBLE_ENTREY_VOUCHERS.Value,0)) AS CreditNotPosted, ACCOUNTS.Account_Code," & "ACCOUNTS.Account_Name, ACCOUNTS.Account_Serial, ACCOUNTS.Account_NameEng "
    StrSQL = StrSQL + " FROM ((Transaction_Type RIGHT JOIN TRANSACTION_HEADER ON " & "Transaction_Type.TransactionID = TRANSACTION_HEADER.Transaction_Type) " & "RIGHT JOIN (NOTES INNER JOIN (ACCOUNTS INNER JOIN DOUBLE_ENTREY_VOUCHERS " & "ON ACCOUNTS.Account_Code = DOUBLE_ENTREY_VOUCHERS.Account_Code) ON NOTES.Note_ID =" & "DOUBLE_ENTREY_VOUCHERS.Notes_Id) ON TRANSACTION_HEADER.Transaction_Header_ID =" & "NOTES.Transaction_Header_ID) LEFT JOIN NOTES AS NOTES_1 ON NOTES.Return_Note_ID =" & "NOTES_1.Note_ID "
    StrSQL = StrSQL + StrWhere

    StrSQL = StrSQL + " Group By ACCOUNTS.Account_Code, ACCOUNTS.Account_Name," & "ACCOUNTS.Account_Serial ,ACCOUNTS.Account_NameEng "
    StrSQL = StrSQL + " Order By ACCOUNTS.Account_Serial "
    rs.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

    With Frm_General_Journal.FgAccountsValue
        .Clear flexClearScrollable, flexClearEverything
        .Rows = .FixedRows + 1

        If Not (rs.BOF Or rs.EOF) Then
            .Rows = .FixedRows + rs.RecordCount
            rs.MoveFirst

            For i = .FixedRows To .Rows - 1
                .TextMatrix(i, .ColIndex("Account_code")) = IIf(IsNull(rs("Account_Code").value), "", rs("Account_Code").value)
                .TextMatrix(i, .ColIndex("Account_Serial")) = IIf(IsNull(rs("Account_Serial").value), "", rs("Account_Serial").value)

                If BolAccountRTL = True Then
                    .TextMatrix(i, .ColIndex("Account_Name")) = IIf(IsNull(rs("Account_Name").value), "", rs("Account_Name").value)
                Else
                    .TextMatrix(i, .ColIndex("Account_Name")) = IIf(IsNull(rs("Account_NameEng").value), "", rs("Account_NameEng").value)
                End If

                .TextMatrix(i, .ColIndex("Debit_Posted")) = IIf(IsNull(rs("DebitPosted").value), "", rs("DebitPosted").value)
                .TextMatrix(i, .ColIndex("Credit_Posted")) = IIf(IsNull(rs("CreditPosted").value), "", rs("CreditPosted").value)
                .TextMatrix(i, .ColIndex("Debit_NotPosted")) = IIf(IsNull(rs("DebitNotPosted").value), "", rs("DebitNotPosted").value)
                .TextMatrix(i, .ColIndex("Credit_NotPosted")) = IIf(IsNull(rs("CreditNotPosted").value), "", rs("CreditNotPosted").value)
                rs.MoveNext
            Next i

        End If

        .AutoSize 0, .ColIndex("Account_Name"), False
    End With

End Sub

Private Function StrBuild() As String
    Dim Begine          As Boolean
    Dim StrWhere        As String
    Dim StrSQL          As String
    Dim StrSubWhere     As String
    Dim BolCheckTrans   As Boolean
    Dim BolChecked      As Boolean
    Dim StrOrder2        As String

    If OptType(1).value = True Then

        For IntCounter = ChkTrans.LBound To ChkTrans.UBound

            If ChkTrans(IntCounter).value = vbChecked Then
                BolChecked = True
                Exit For
            End If

        Next IntCounter

        For IntCounter = ChkNotes.LBound To ChkNotes.UBound

            If ChkNotes(IntCounter).value = vbChecked Then
                BolChecked = True
                Exit For
            End If

        Next IntCounter

        If BolChecked = False Then
            GetMsgs 83, vbExclamation
            Exit Function
        End If
    End If

    If SystemOptions.SysDataBaseType = SQLServerDataBase Then
        StrSQL = "SELECT  * From QryDevReport Where Double_Entry_Vouchers_ID <> 0 "
    ElseIf SystemOptions.SysDataBaseType = AccessDataBase Then
        Exit Function
    End If

    If Not IsNull(Me.DTPDev_From.value) Then
        StrWhere = " AND RecordDate >=" & SQLDate(Me.DTPDev_From.value, True) & ""
    End If

    If Not IsNull(Me.DTPDEV_TO.value) Then
        StrWhere = StrWhere + " And RecordDate <=" & SQLDate(Me.DTPDEV_TO.value, True) & ""
    End If

    '------------------------
    If OptType(1).value = True Then
        StrWhere = StrWhere + " And ("
        StrSubWhere = ""

        For IntCounter = ChkTrans.LBound To ChkTrans.UBound

            If ChkTrans(IntCounter).value = vbChecked Then
                BolCheckTrans = True
                StrSubWhere = StrSubWhere + " Transaction_Type =" & ChkTrans(IntCounter).Tag & " OR "
            End If

        Next IntCounter

        If StrSubWhere <> "" Then
            StrSubWhere = left(StrSubWhere, Len(StrSubWhere) - 3)
            StrSubWhere = "(" & StrSubWhere & ")"
            StrWhere = StrWhere & StrSubWhere
        End If

        StrSubWhere = ""

        For IntCounter = ChkNotes.LBound To ChkNotes.UBound

            If ChkNotes(IntCounter).value = vbChecked Then
                StrSubWhere = StrSubWhere + " NoteType =" & ChkNotes(IntCounter).Tag & " OR "
            End If

        Next IntCounter

        If StrSubWhere <> "" Then
            StrSubWhere = left(StrSubWhere, Len(StrSubWhere) - 3)
            StrSubWhere = "(" & StrSubWhere & ")"

            If BolCheckTrans = True Then
                StrWhere = StrWhere & " OR " & StrSubWhere
            Else
                StrWhere = StrWhere & StrSubWhere
            End If
        End If

        StrWhere = StrWhere & ")"
    End If

    '------------------------
    If Opt(0).value = True Then '„—Õ·
        StrWhere = StrWhere + " And Posted =1"
    ElseIf Opt(1).value = True Then '€Ì— „—Õ·
        StrWhere = StrWhere + " And Posted =0"
    End If

    If Me.DcboUsers(0).BoundText <> "" Then
        'StrWhere = StrWhere + " And NOTES.Issued_By =" & DcboUsers(0).BoundText & ""
    End If

    If Me.DcboUsers(1).Enabled = True Then
        '    If Me.DcboUsers(1).BoundText <> "" Then
        '    If Begine = True Then
        '        StrWhere = StrWhere + " And NOTES.PostedBy =" & DcboUsers(1).BoundText & ""
        '    Else
        '        Begine = True
        '        StrWhere = " Where NOTES.PostedBy =" & DcboUsers(1).BoundText & ""
        '    End If
        'End If
    End If

    StrOrder2 = "ORDER BY Double_Entry_Vouchers_ID,Credit_Or_Debit"
    StrSQL = StrSQL + StrWhere + StrOrder2
    StrBuild = StrSQL
    SetFgAccounts StrWhere
    StrBuild = StrSQL
End Function
