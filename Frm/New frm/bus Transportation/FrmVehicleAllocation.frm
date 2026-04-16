VERSION 5.00
Object = "{C115893A-A3BF-43AF-B28D-69DB846077F3}#1.0#0"; "vsflex8u.ocx"
Object = "{0AFE7BE0-11B7-4A3E-978D-D4501E9A57FE}#1.0#0"; "c1sizer.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{FE5DCFAD-BC1D-11D2-94CF-004005455FAA}#1.4#0"; "ImpulseButton.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form FrmVehicleAllocation 
   BackColor       =   &H00E2E9E9&
   BorderStyle     =   1  'Fixed Single
   Caption         =   " Œ’Ì’ «·Õ«›·« "
   ClientHeight    =   10065
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   15075
   Icon            =   "FrmVehicleAllocation.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   RightToLeft     =   -1  'True
   ScaleHeight     =   10065
   ScaleWidth      =   15075
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
   Begin C1SizerLibCtl.C1Elastic C1Elastic1 
      Height          =   10065
      Left            =   0
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   0
      Width           =   15075
      _cx             =   26591
      _cy             =   17754
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
      Begin VB.CommandButton Command1 
         Caption         =   "Õ–› «·ﬂ·"
         Height          =   492
         Left            =   120
         RightToLeft     =   -1  'True
         TabIndex        =   83
         Top             =   7095
         Width           =   1092
      End
      Begin VB.CommandButton Command2 
         Caption         =   " Õ–› ”ÿ—"
         Height          =   492
         Left            =   1320
         RightToLeft     =   -1  'True
         TabIndex        =   82
         Top             =   7095
         Width           =   1092
      End
      Begin C1SizerLibCtl.C1Elastic C1Elastic3 
         Height          =   1350
         Left            =   120
         TabIndex        =   27
         TabStop         =   0   'False
         Top             =   7695
         Width           =   3975
         _cx             =   7011
         _cy             =   2381
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
         Begin VB.Label Lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "0"
            ForeColor       =   &H000000C0&
            Height          =   372
            Index           =   11
            Left            =   120
            RightToLeft     =   -1  'True
            TabIndex        =   33
            Top             =   120
            Width           =   852
         End
         Begin VB.Label Label2 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "0"
            ForeColor       =   &H000000C0&
            Height          =   300
            Left            =   456
            TabIndex        =   32
            Top             =   540
            Width           =   516
         End
         Begin VB.Label Lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "0"
            ForeColor       =   &H000000C0&
            Height          =   288
            Index           =   12
            Left            =   120
            RightToLeft     =   -1  'True
            TabIndex        =   31
            Top             =   948
            Width           =   852
         End
         Begin VB.Label Lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "«Ã„«·Ï «·ÿ·«» «·–Ì‰  „  ”ﬂÌ‰Â„"
            ForeColor       =   &H000000C0&
            Height          =   372
            Index           =   6
            Left            =   1596
            RightToLeft     =   -1  'True
            TabIndex        =   30
            Top             =   120
            Width           =   2136
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "⁄œœ «·ÿ·«» «·„ ⁄«ﬁœ ⁄·ÌÂ„"
            ForeColor       =   &H000000C0&
            Height          =   300
            Index           =   1
            Left            =   1596
            TabIndex        =   29
            Top             =   540
            Width           =   2136
         End
         Begin VB.Label Lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "«·›—ﬁ «·ﬁ«»· ··«”‰«œ ··„ ⁄ÂœÌ‰"
            ForeColor       =   &H000000C0&
            Height          =   288
            Index           =   10
            Left            =   1188
            RightToLeft     =   -1  'True
            TabIndex        =   28
            Top             =   948
            Width           =   2544
         End
      End
      Begin C1SizerLibCtl.C1Elastic C1Elastic2 
         Height          =   2340
         Left            =   7368
         TabIndex        =   20
         TabStop         =   0   'False
         Top             =   840
         Width           =   7668
         _cx             =   13520
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
         Begin VB.TextBox txtStudentCount1 
            Alignment       =   1  'Right Justify
            Enabled         =   0   'False
            Height          =   300
            Left            =   240
            MaxLength       =   50
            RightToLeft     =   -1  'True
            TabIndex        =   64
            Top             =   1092
            Width           =   2028
         End
         Begin VB.TextBox txtIDVA 
            Alignment       =   1  'Right Justify
            Height          =   288
            Left            =   4020
            Locked          =   -1  'True
            RightToLeft     =   -1  'True
            TabIndex        =   61
            Top             =   168
            Width           =   1992
         End
         Begin VB.TextBox txtProcess 
            Alignment       =   1  'Right Justify
            Height          =   324
            Left            =   4020
            MaxLength       =   50
            RightToLeft     =   -1  'True
            TabIndex        =   1
            Top             =   -72
            Visible         =   0   'False
            Width           =   1992
         End
         Begin VB.TextBox txtMinistryNo 
            Alignment       =   1  'Right Justify
            Height          =   300
            Left            =   4020
            MaxLength       =   50
            RightToLeft     =   -1  'True
            TabIndex        =   3
            Top             =   792
            Width           =   1992
         End
         Begin VB.TextBox txtDepend 
            Alignment       =   1  'Right Justify
            Height          =   324
            Left            =   660
            MaxLength       =   50
            RightToLeft     =   -1  'True
            TabIndex        =   5
            Top             =   2940
            Visible         =   0   'False
            Width           =   1632
         End
         Begin VB.TextBox txtStudentCount 
            Alignment       =   1  'Right Justify
            Enabled         =   0   'False
            Height          =   300
            Left            =   240
            Locked          =   -1  'True
            MaxLength       =   50
            RightToLeft     =   -1  'True
            TabIndex        =   7
            Top             =   480
            Width           =   2028
         End
         Begin MSDataListLib.DataCombo dcSchoolFile 
            Height          =   288
            Left            =   240
            TabIndex        =   4
            Top             =   792
            Width           =   2028
            _ExtentX        =   3572
            _ExtentY        =   556
            _Version        =   393216
            Style           =   2
            BackColor       =   16777215
            Text            =   ""
            RightToLeft     =   -1  'True
         End
         Begin MSComCtl2.DTPicker dtpProcessDate 
            Height          =   336
            Left            =   240
            TabIndex        =   12
            TabStop         =   0   'False
            Top             =   168
            Width           =   2028
            _ExtentX        =   3572
            _ExtentY        =   582
            _Version        =   393216
            CalendarBackColor=   12648447
            CalendarTitleBackColor=   10383715
            CustomFormat    =   "yyyy/M/d"
            Format          =   99549187
            CurrentDate     =   37140
         End
         Begin MSComCtl2.DTPicker DTPicker1 
            Height          =   312
            Left            =   240
            TabIndex        =   9
            TabStop         =   0   'False
            Top             =   1488
            Visible         =   0   'False
            Width           =   2028
            _ExtentX        =   3572
            _ExtentY        =   556
            _Version        =   393216
            CalendarBackColor=   12648447
            CalendarTitleBackColor=   10383715
            CustomFormat    =   "yyyy/M/d"
            Format          =   99549187
            CurrentDate     =   37140
         End
         Begin MSComCtl2.DTPicker DTPicker2 
            Height          =   312
            Left            =   4020
            TabIndex        =   8
            TabStop         =   0   'False
            Top             =   1488
            Visible         =   0   'False
            Width           =   1992
            _ExtentX        =   3519
            _ExtentY        =   556
            _Version        =   393216
            CalendarBackColor=   12648447
            CalendarTitleBackColor=   10383715
            CustomFormat    =   "yyyy/M/d"
            Format          =   99549187
            CurrentDate     =   37140
         End
         Begin Dynamic_Byte.NourHijriCal NourHijriCal1 
            Height          =   324
            Left            =   240
            TabIndex        =   11
            Top             =   1848
            Visible         =   0   'False
            Width           =   2028
            _ExtentX        =   3625
            _ExtentY        =   582
         End
         Begin Dynamic_Byte.NourHijriCal NourHijriCal2 
            Height          =   336
            Left            =   4020
            TabIndex        =   10
            Top             =   1848
            Visible         =   0   'False
            Width           =   1992
            _ExtentX        =   3519
            _ExtentY        =   582
         End
         Begin MSDataListLib.DataCombo dcDuration 
            Height          =   288
            Left            =   4020
            TabIndex        =   2
            Top             =   1080
            Width           =   1992
            _ExtentX        =   3519
            _ExtentY        =   556
            _Version        =   393216
            Style           =   2
            BackColor       =   16777215
            Text            =   ""
            RightToLeft     =   -1  'True
         End
         Begin ImpulseButton.ISButton Cmd 
            Height          =   288
            Index           =   5
            Left            =   300
            TabIndex        =   6
            Top             =   2820
            Visible         =   0   'False
            Width           =   360
            _ExtentX        =   635
            _ExtentY        =   503
            ButtonPositionImage=   1
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
            ButtonImage     =   "FrmVehicleAllocation.frx":038A
            ColorButton     =   14871017
            ColorHighlight  =   16777215
            ColorHoverText  =   16711680
            ColorShadow     =   -2147483637
            ColorOutline    =   0
            DrawFocusRectangle=   0   'False
            ColorToggledHoverText=   16711680
            ColorTextShadow =   -2147483637
         End
         Begin MSDataListLib.DataCombo dcMinistryContract 
            Height          =   288
            Left            =   4020
            TabIndex        =   63
            Top             =   492
            Width           =   1992
            _ExtentX        =   3519
            _ExtentY        =   556
            _Version        =   393216
            Style           =   2
            BackColor       =   16777215
            Text            =   ""
            RightToLeft     =   -1  'True
         End
         Begin VB.Label Lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "»«ﬁÏ «· Œ’Ì’"
            Height          =   312
            Index           =   25
            Left            =   2376
            RightToLeft     =   -1  'True
            TabIndex        =   84
            Top             =   480
            Width           =   1464
         End
         Begin VB.Label Lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "«·”‰… «·œ—«”Ì…"
            Height          =   300
            Index           =   17
            Left            =   6672
            RightToLeft     =   -1  'True
            TabIndex        =   62
            Top             =   1080
            Width           =   828
         End
         Begin VB.Label Lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   " «—ÌŒ ‰Â«Ì… «· Œ’Ì’ „"
            Height          =   324
            Index           =   20
            Left            =   2520
            RightToLeft     =   -1  'True
            TabIndex        =   60
            Top             =   1488
            Visible         =   0   'False
            Width           =   1380
         End
         Begin VB.Label Lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   " «—ÌŒ »œ«Ì… «· Œ’Ì’ Â‹ "
            Height          =   336
            Index           =   19
            Left            =   6120
            RightToLeft     =   -1  'True
            TabIndex        =   59
            Top             =   1848
            Visible         =   0   'False
            Width           =   1380
         End
         Begin VB.Label Lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   " «—ÌŒ ‰Â«Ì… «· Œ’Ì’ Â‹‹ "
            Height          =   336
            Index           =   18
            Left            =   2400
            RightToLeft     =   -1  'True
            TabIndex        =   58
            Top             =   1848
            Visible         =   0   'False
            Width           =   1500
         End
         Begin VB.Label Lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   " «—ÌŒ »œ«Ì… «· Œ’Ì’ „"
            Height          =   324
            Index           =   16
            Left            =   6120
            RightToLeft     =   -1  'True
            TabIndex        =   57
            Top             =   1488
            Visible         =   0   'False
            Width           =   1380
         End
         Begin VB.Label Lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "»‰«¡ ⁄·Ï ⁄ﬁœ "
            Height          =   240
            Index           =   15
            Left            =   6360
            RightToLeft     =   -1  'True
            TabIndex        =   56
            Top             =   492
            Width           =   1140
         End
         Begin VB.Label Lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "«·—ﬁ„ «·Ê“«—Ï ··„œ—”…"
            Height          =   312
            Index           =   9
            Left            =   6036
            RightToLeft     =   -1  'True
            TabIndex        =   55
            Top             =   792
            Width           =   1464
         End
         Begin VB.Label Label3 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   " «—ÌŒ «· Œ’Ì’"
            ForeColor       =   &H00000000&
            Height          =   264
            Left            =   2904
            TabIndex        =   54
            Top             =   168
            Width           =   996
         End
         Begin VB.Label Lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "—ﬁ„ «· Œ’Ì’"
            Height          =   204
            Index           =   7
            Left            =   6036
            RightToLeft     =   -1  'True
            TabIndex        =   53
            Top             =   168
            Width           =   1464
         End
         Begin VB.Label Lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "«”„ «·„œ—”…"
            Height          =   300
            Index           =   1
            Left            =   3072
            RightToLeft     =   -1  'True
            TabIndex        =   22
            Top             =   792
            Width           =   828
         End
         Begin VB.Label Lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "⁄œœ «·ÿ·«» ··„œ—”…"
            Height          =   216
            Index           =   5
            Left            =   2400
            RightToLeft     =   -1  'True
            TabIndex        =   21
            Top             =   1092
            Width           =   1500
         End
      End
      Begin C1SizerLibCtl.C1Elastic EleHeader 
         Height          =   684
         Left            =   0
         TabIndex        =   14
         TabStop         =   0   'False
         Top             =   0
         Width           =   15108
         _cx             =   26644
         _cy             =   1217
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
         Caption         =   "     Œ’Ì’ «·Õ«›·«   "
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
            Left            =   12960
            RightToLeft     =   -1  'True
            TabIndex        =   15
            Top             =   1320
            Visible         =   0   'False
            Width           =   855
         End
         Begin ImpulseButton.ISButton XPBtnMove 
            Height          =   345
            Index           =   0
            Left            =   1155
            TabIndex        =   16
            Top             =   120
            Width           =   495
            _ExtentX        =   873
            _ExtentY        =   609
            ButtonStyle     =   1
            ButtonPositionImage=   4
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
            ButtonImage     =   "FrmVehicleAllocation.frx":0CB6
            ColorButton     =   16777215
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
            TabIndex        =   17
            Top             =   120
            Width           =   495
            _ExtentX        =   873
            _ExtentY        =   609
            ButtonStyle     =   1
            ButtonPositionImage=   4
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
            ButtonImage     =   "FrmVehicleAllocation.frx":1050
            ColorButton     =   16777215
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
            TabIndex        =   18
            Top             =   120
            Width           =   495
            _ExtentX        =   873
            _ExtentY        =   609
            ButtonStyle     =   1
            ButtonPositionImage=   4
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
            ButtonImage     =   "FrmVehicleAllocation.frx":13EA
            ColorButton     =   16777215
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
            TabIndex        =   19
            Top             =   120
            Width           =   495
            _ExtentX        =   873
            _ExtentY        =   609
            ButtonStyle     =   1
            ButtonPositionImage=   4
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
            ButtonImage     =   "FrmVehicleAllocation.frx":1784
            ColorButton     =   16777215
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
      Begin VSFlex8UCtl.VSFlexGrid FgInstallments 
         Height          =   3765
         Left            =   120
         TabIndex        =   13
         Top             =   3240
         Width           =   14895
         _cx             =   26273
         _cy             =   6641
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
         Cols            =   13
         FixedRows       =   1
         FixedCols       =   1
         RowHeightMin    =   0
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   -1  'True
         FormatString    =   $"FrmVehicleAllocation.frx":1B1E
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
      Begin C1SizerLibCtl.C1Elastic C1Elastic4 
         Height          =   2352
         Left            =   120
         TabIndex        =   34
         TabStop         =   0   'False
         Top             =   840
         Width           =   7140
         _cx             =   12594
         _cy             =   4154
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Enabled         =   0   'False
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
         Begin VB.TextBox XPTxtBoxName 
            Alignment       =   1  'Right Justify
            Height          =   372
            Left            =   120
            MaxLength       =   50
            RightToLeft     =   -1  'True
            TabIndex        =   35
            Top             =   276
            Width           =   5880
         End
         Begin MSComCtl2.DTPicker dtpToDate 
            Height          =   396
            Left            =   1176
            TabIndex        =   36
            TabStop         =   0   'False
            Top             =   1548
            Width           =   1080
            _ExtentX        =   1905
            _ExtentY        =   688
            _Version        =   393216
            CalendarBackColor=   12648447
            CalendarTitleBackColor=   10383715
            CustomFormat    =   "yyyy/M/d"
            Format          =   99549187
            CurrentDate     =   37140
         End
         Begin MSComCtl2.DTPicker dtpFromDate 
            Height          =   396
            Left            =   4836
            TabIndex        =   37
            TabStop         =   0   'False
            Top             =   1548
            Width           =   1188
            _ExtentX        =   2090
            _ExtentY        =   688
            _Version        =   393216
            CalendarBackColor=   12648447
            CalendarTitleBackColor=   10383715
            CustomFormat    =   "yyyy/M/d"
            Format          =   99549187
            CurrentDate     =   37140
         End
         Begin Dynamic_Byte.NourHijriCal dtpFromDateH 
            Height          =   396
            Left            =   3540
            TabIndex        =   38
            Top             =   1548
            Width           =   1308
            _ExtentX        =   2302
            _ExtentY        =   688
         End
         Begin Dynamic_Byte.NourHijriCal dtpToDateH 
            Height          =   396
            Left            =   120
            TabIndex        =   39
            Top             =   1548
            Width           =   1068
            _ExtentX        =   1879
            _ExtentY        =   688
         End
         Begin Dynamic_Byte.NourHijriCal dtpSContractDateH 
            Height          =   312
            Left            =   3540
            TabIndex        =   40
            Top             =   1104
            Width           =   1308
            _ExtentX        =   2302
            _ExtentY        =   556
         End
         Begin MSComCtl2.DTPicker dtpSContractDate 
            Height          =   312
            Left            =   4836
            TabIndex        =   41
            TabStop         =   0   'False
            Top             =   1104
            Width           =   1188
            _ExtentX        =   2090
            _ExtentY        =   556
            _Version        =   393216
            CalendarBackColor=   12648447
            CalendarTitleBackColor=   10383715
            CustomFormat    =   "yyyy/M/d"
            Format          =   99549187
            CurrentDate     =   37140
         End
         Begin MSComCtl2.DTPicker dtpEContractDate 
            Height          =   312
            Left            =   1176
            TabIndex        =   42
            TabStop         =   0   'False
            Top             =   1104
            Width           =   1080
            _ExtentX        =   1905
            _ExtentY        =   556
            _Version        =   393216
            CalendarBackColor=   12648447
            CalendarTitleBackColor=   10383715
            CustomFormat    =   "yyyy/M/d"
            Format          =   99549187
            CurrentDate     =   37140
         End
         Begin Dynamic_Byte.NourHijriCal dtpEContractDateH 
            Height          =   312
            Left            =   120
            TabIndex        =   43
            Top             =   1104
            Width           =   1068
            _ExtentX        =   1879
            _ExtentY        =   556
         End
         Begin MSDataListLib.DataCombo dcVendor 
            Height          =   288
            Left            =   120
            TabIndex        =   44
            Top             =   720
            Width           =   2112
            _ExtentX        =   3731
            _ExtentY        =   556
            _Version        =   393216
            Style           =   2
            BackColor       =   16777215
            Text            =   ""
            RightToLeft     =   -1  'True
         End
         Begin MSDataListLib.DataCombo dcCity 
            Height          =   288
            Left            =   3540
            TabIndex        =   45
            Top             =   720
            Width           =   2460
            _ExtentX        =   4339
            _ExtentY        =   556
            _Version        =   393216
            Style           =   2
            BackColor       =   16777215
            Text            =   ""
            RightToLeft     =   -1  'True
         End
         Begin VB.Label Lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "„”„Ï «· ⁄«ﬁœ"
            Height          =   372
            Index           =   14
            Left            =   5784
            RightToLeft     =   -1  'True
            TabIndex        =   52
            Top             =   276
            Width           =   1188
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "«·Ï  «—ÌŒ"
            ForeColor       =   &H00000000&
            Height          =   456
            Index           =   0
            Left            =   2628
            TabIndex        =   51
            Top             =   1548
            Width           =   804
         End
         Begin VB.Label Label4 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   " «—ÌŒ"
            ForeColor       =   &H00000000&
            Height          =   300
            Left            =   6240
            TabIndex        =   50
            Top             =   1548
            Width           =   804
         End
         Begin VB.Label Lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "  «· ⁄«ﬁœ „Ì·«œÏ"
            Height          =   312
            Index           =   13
            Left            =   6012
            RightToLeft     =   -1  'True
            TabIndex        =   49
            Top             =   1104
            Width           =   960
         End
         Begin VB.Label Lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "  «‰ Â«¡ «· ⁄«ﬁœ „Ì·«œÏ"
            Height          =   576
            Index           =   8
            Left            =   2364
            RightToLeft     =   -1  'True
            TabIndex        =   48
            Top             =   1104
            Width           =   1068
         End
         Begin VB.Label Lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "«”„ «·„‰ÿﬁ…"
            Height          =   384
            Index           =   3
            Left            =   6012
            RightToLeft     =   -1  'True
            TabIndex        =   47
            Top             =   720
            Width           =   960
         End
         Begin VB.Label Lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "«·«œ«—… «· ⁄·Ì„Ì…"
            Height          =   384
            Index           =   0
            Left            =   2364
            RightToLeft     =   -1  'True
            TabIndex        =   46
            Top             =   720
            Width           =   1068
         End
      End
      Begin C1SizerLibCtl.C1Elastic C1Elastic6 
         Height          =   1350
         Left            =   4200
         TabIndex        =   65
         TabStop         =   0   'False
         Top             =   7695
         Width           =   3975
         _cx             =   7011
         _cy             =   2381
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
         Begin VB.Label Lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "«·›—ﬁ "
            ForeColor       =   &H000000C0&
            Height          =   288
            Index           =   24
            Left            =   1188
            RightToLeft     =   -1  'True
            TabIndex        =   71
            Top             =   948
            Width           =   2544
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "«Ã„«·Ï ⁄œœ ÿ·«» «·„œ—”…  „ »ﬁÌ"
            ForeColor       =   &H000000C0&
            Height          =   300
            Index           =   2
            Left            =   1596
            TabIndex        =   70
            Top             =   540
            Width           =   2136
         End
         Begin VB.Label Lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "«Ã„«·Ï «·ÿ·«» «·–Ì‰  „  ”ﬂÌ‰Â„"
            ForeColor       =   &H000000C0&
            Height          =   375
            Index           =   23
            Left            =   1350
            RightToLeft     =   -1  'True
            TabIndex        =   69
            Top             =   120
            Width           =   2370
         End
         Begin VB.Label Lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "0"
            ForeColor       =   &H000000C0&
            Height          =   288
            Index           =   22
            Left            =   120
            RightToLeft     =   -1  'True
            TabIndex        =   68
            Top             =   948
            Width           =   852
         End
         Begin VB.Label Label5 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "0"
            ForeColor       =   &H000000C0&
            Height          =   300
            Left            =   456
            TabIndex        =   67
            Top             =   540
            Width           =   516
         End
         Begin VB.Label Lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "0"
            ForeColor       =   &H000000C0&
            Height          =   372
            Index           =   21
            Left            =   120
            RightToLeft     =   -1  'True
            TabIndex        =   66
            Top             =   120
            Width           =   852
         End
      End
      Begin C1SizerLibCtl.C1Elastic C1Elastic5 
         Height          =   810
         Left            =   120
         TabIndex        =   72
         TabStop         =   0   'False
         Top             =   9120
         Width           =   14820
         _cx             =   26141
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
            Height          =   540
            Index           =   0
            Left            =   13020
            TabIndex        =   73
            Top             =   144
            Width           =   1572
            _ExtentX        =   2778
            _ExtentY        =   953
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
            ButtonImage     =   "FrmVehicleAllocation.frx":1D15
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
            Height          =   540
            Index           =   1
            Left            =   11364
            TabIndex        =   74
            Top             =   144
            Width           =   1608
            _ExtentX        =   2831
            _ExtentY        =   953
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
            ButtonImage     =   "FrmVehicleAllocation.frx":8577
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
            Height          =   540
            Index           =   2
            Left            =   9492
            TabIndex        =   75
            Top             =   120
            Width           =   1848
            _ExtentX        =   3254
            _ExtentY        =   953
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
            ButtonImage     =   "FrmVehicleAllocation.frx":EDD9
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
            Height          =   540
            Index           =   3
            Left            =   7896
            TabIndex        =   76
            Top             =   144
            Width           =   1536
            _ExtentX        =   2699
            _ExtentY        =   953
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
            ButtonImage     =   "FrmVehicleAllocation.frx":1563B
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
            Height          =   540
            Index           =   4
            Left            =   6276
            TabIndex        =   77
            Top             =   144
            Width           =   1548
            _ExtentX        =   2725
            _ExtentY        =   953
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
            ButtonImage     =   "FrmVehicleAllocation.frx":1BE9D
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
            Height          =   540
            Index           =   6
            Left            =   1716
            TabIndex        =   78
            Top             =   144
            Width           =   1476
            _ExtentX        =   2593
            _ExtentY        =   953
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
            ButtonImage     =   "FrmVehicleAllocation.frx":226FF
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
            Height          =   540
            Left            =   120
            TabIndex        =   79
            Top             =   144
            Width           =   1548
            _ExtentX        =   2725
            _ExtentY        =   953
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
            ButtonImage     =   "FrmVehicleAllocation.frx":4C321
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
            Height          =   540
            Index           =   7
            Left            =   4800
            TabIndex        =   80
            Top             =   144
            Width           =   1452
            _ExtentX        =   2566
            _ExtentY        =   953
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
            ButtonImage     =   "FrmVehicleAllocation.frx":52B83
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
            Height          =   540
            Index           =   9
            Left            =   3240
            TabIndex        =   81
            Top             =   144
            Width           =   1464
            _ExtentX        =   2593
            _ExtentY        =   953
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
            ButtonImage     =   "FrmVehicleAllocation.frx":593E5
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
      Begin VB.Label XPTxtCurrent 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Height          =   375
         Left            =   11730
         RightToLeft     =   -1  'True
         TabIndex        =   26
         Top             =   8625
         Width           =   990
      End
      Begin VB.Label XPTxtCount 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Height          =   375
         Left            =   8280
         RightToLeft     =   -1  'True
         TabIndex        =   25
         Top             =   8625
         Width           =   855
      End
      Begin VB.Label Lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   " «·”Ã· «·Õ«·Ì:"
         Height          =   375
         Index           =   2
         Left            =   12645
         RightToLeft     =   -1  'True
         TabIndex        =   24
         Top             =   8625
         Width           =   1365
      End
      Begin VB.Label Lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   " ⁄œœ «·”Ã·« :"
         Height          =   375
         Index           =   4
         Left            =   9030
         RightToLeft     =   -1  'True
         TabIndex        =   23
         Top             =   8625
         Width           =   1410
      End
   End
End
Attribute VB_Name = "FrmVehicleAllocation"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rs As ADODB.Recordset
Dim Rs_Temp As ADODB.Recordset
Dim rs_det As ADODB.Recordset
Dim rsVendor As ADODB.Recordset


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
          txtIDVA.Text = CStr(new_id("TblVehicleAllocation", "IDVA", "", True))
          FgInstallments.Rows = 2
          '  XPTxtBoxName.SetFocus
          Clear_SumationLabels
          
        Case 1

            If DoPremis(Do_Edit, Me.Name, True) = False Then
                Exit Sub
            End If

            TxtModFlg.Text = "E"
            FgInstallments.Rows = FgInstallments.Rows + 1
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
                      FrmSearch_MinistryContract.SendForm = "VA"
             FrmSearch_MinistryContract.show
        Case 6
            Unload Me
         Case 7
       Me.print_VA
   Case 9
            Unload FrmSearch_MinistryContract
            FrmSearch_MinistryContract.SendForm = "VA2"
            FrmSearch_MinistryContract.show
    End Select

    Exit Sub
ErrTrap:
End Sub
Private Sub Clear_SumationLabels()

     
         lbl(11).Caption = ""
         lbl(12).Caption = ""
         Label2.Caption = ""
         lbl(21).Caption = ""
         Label5.Caption = ""
         lbl(22).Caption = ""

End Sub

Private Sub CmdHelp_Click()
    SystemOptions.SysHelp.HHTopicID = Me.HelpContextID
    SystemOptions.SysHelp.HHDisplayTopicID Me.hwnd
End Sub



Private Sub Option1_Click()
'If Option1.value = True Then
'    Frame1.Visible = False
'Else
'    Frame1.Visible = True
'End If
End Sub

 


Private Sub CmdAttach_Click()
            On Error Resume Next
'ShowAttachments XPTxtBoxID, "0701201405"
 
    On Error Resume Next
          If DoPremis(Do_Attach, Me.Name, True) = False Then
                Exit Sub
            End If
            
ShowAttachments txtIDVA, "15062020008"

End Sub

Private Sub Command1_Click()
FgInstallments.Rows = 1
End Sub

Private Sub Command2_Click()
If FgInstallments.Row < FgInstallments.FixedRows Then Exit Sub
FgInstallments.RemoveItem (FgInstallments.Row)

End Sub

Private Sub dcDuration_Click(Area As Integer)
Calc_SchoolRes
Get_Resdent
End Sub

Private Sub dcDuration_KeyUp(KeyCode As Integer, Shift As Integer)

If KeyCode = vbKeyF3 Then
    'Unload FrmSearch_Duration
    'FrmSearch_Duration.SendForm = "VA"
    'FrmSearch_Duration.show vbModal
End If

End Sub

Private Sub dcMinistryContract_Change()
On Error Resume Next
       Dim StrSQL As String
       txtStudentCount.Text = ""
       Set Rs_Temp = New ADODB.Recordset
       StrSQL = " select * from tblministrycontract where IDMC =   " & dcMinistryContract.BoundText
       Rs_Temp.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText
       If (Rs_Temp.RecordCount > 0) Then
            txtStudentCount.Text = IIf(IsNull(Rs_Temp("studentcount").value), "", Rs_Temp("studentcount").value)
            XPTxtBoxName.Text = IIf(IsNull(Rs_Temp("Name").value), "", Rs_Temp("Name").value)
            dcCity.BoundText = IIf(IsNull(Rs_Temp("CityID").value), "", Rs_Temp("CityID").value)
            DCVendor.BoundText = IIf(IsNull(Rs_Temp("VendorID").value), "", Rs_Temp("VendorID").value)
            dtpSContractDate.value = IIf(IsNull(Rs_Temp("StartContractDate").value), "", Rs_Temp("StartContractDate").value)
            dtpEContractDate.value = IIf(IsNull(Rs_Temp("EndContractDate").value), "", Rs_Temp("EndContractDate").value)
            dtpSContractDateH.value = IIf(IsNull(Rs_Temp("StartContractDateh").value), "", Rs_Temp("StartContractDateh").value)
            dtpSContractDateH.value = IIf(IsNull(Rs_Temp("EndContractDateH").value), "", Rs_Temp("EndContractDateH").value)
            dtpFromDate.value = IIf(IsNull(Rs_Temp("fromdate").value), "", Rs_Temp("fromdate").value)
            dtpFromDateH.value = IIf(IsNull(Rs_Temp("fromdateh").value), "", Rs_Temp("fromdateh").value)
       Else
            txtStudentCount.Text = ""
            XPTxtBoxName.Text = ""
            dcCity.BoundText = ""
            DCVendor.BoundText = ""
            dtpSContractDate.value = Date
            dtpEContractDate.value = Date
            dtpSContractDateH.value = ToHijriDate(Date)
            dtpSContractDateH.value = ToHijriDate(Date)
            dtpFromDate.value = Date
            dtpFromDateH.value = Rs_Temp("fromdateh").value
            
       End If
       Calc_SchoolRes
       Get_Resdent
End Sub

Private Sub Get_Resdent()
    Dim str As String
    Dim alloc As Double, StudentCount As Double, attr As Double
    
   If dcDuration.BoundText = "" Or dcMinistryContract.BoundText = "" Then
   Exit Sub
   End If
    
    
   ' str = "select StudentCount  from TblMinistryContract where ProcessNo =  '" & txtMinistryContractNo.text & "'"
    str = "select StudentCount  from TblMinistryContract where IDMC =  " & val(dcMinistryContract.BoundText)
    Set Rs_Temp = New ADODB.Recordset
    Rs_Temp.Open str, Cn, adOpenStatic, adLockOptimistic, adCmdText
    If Rs_Temp.RecordCount > 0 Then
        StudentCount = IIf(IsNull(Rs_Temp("StudentCount").value), 0, Rs_Temp("StudentCount").value)
    End If
    
    str = " select * from TblVehicleAllocation where IDMC = " & val(dcMinistryContract.BoundText) & " and DurationID =  " & val(dcDuration.BoundText)
    Set Rs_Temp = New ADODB.Recordset
    Rs_Temp.Open str, Cn, adOpenStatic, adLockOptimistic, adCmdText
    If Rs_Temp.RecordCount > 0 Then
        alloc = IIf(IsNull(Rs_Temp("StudentAlloc").value), 0, Rs_Temp("StudentAlloc").value)
    End If
    
    str = " select  sum (COALESCE (studentcount , 0)) sumstudent from TblAttributionContract where idmc  =   " & dcMinistryContract.BoundText & " and DurationID =  " & val(dcDuration.BoundText)
    Set Rs_Temp = New ADODB.Recordset
    Rs_Temp.Open str, Cn, adOpenStatic, adLockOptimistic, adCmdText
    If Rs_Temp.RecordCount > 0 Then
        attr = IIf(IsNull(Rs_Temp("sumstudent").value), 0, Rs_Temp("sumstudent").value)
    End If
    
    txtStudentCount.Text = StudentCount - alloc - attr
    
End Sub



Private Sub dcMinistryContract_KeyUp(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyF3 Then
  
             Unload FrmSearch_MinistryContract
        
        FrmSearch_MinistryContract.SendForm = "VA"
        FrmSearch_MinistryContract.show
    End If
End Sub

Private Sub Calc_SchoolRes()

      Dim StrSQL As String
       If dcSchoolFile.BoundText = "" Then Exit Sub
       Set Rs_Temp = New ADODB.Recordset
       StrSQL = " select *  from tblschoolefile where id =    " & val(dcSchoolFile.BoundText)
       Rs_Temp.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText
       If (Rs_Temp.RecordCount > 0) Then
            txtMinistryNo.Text = IIf(IsNull(Rs_Temp("ministerno").value), "", Rs_Temp("ministerno").value)
            txtStudentCount1.Text = IIf(IsNull(Rs_Temp("StudentCount").value), "", Rs_Temp("StudentCount").value)
       Else
              txtMinistryNo.Text = ""
              txtStudentCount1 = ""
       End If



    Dim tot As Integer
    Set Rs_Temp = New ADODB.Recordset
    StrSQL = " select sum(d.StudentCount) tot from  TblVehicleAllocation H , TblVehicleAllocation_Details D where h.IDVA = d.IDva and type = 2  and IDMC = " & val(dcMinistryContract.BoundText) & " and DurationID = " & val(dcDuration.BoundText) & " and H.SchoolFileID   =" & val(dcSchoolFile.BoundText)
    Rs_Temp.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText
    If Rs_Temp.RecordCount > 0 Then
            tot = IIf(IsNull(Rs_Temp("tot").value), 0, Rs_Temp("tot").value)
    End If

    Set Rs_Temp = New ADODB.Recordset
    StrSQL = " select sum(d.StudentCount) tot from  TblAttributionContract  H , TblVehicleAllocation_Details D where h.IDAC  = d.IDva and type = 3  and H.IDMC = " & val(dcMinistryContract.BoundText) & " and h.DurationID = " & val(dcDuration.BoundText) & " and D.SchoolFileID = " & val(dcSchoolFile.BoundText)
    Rs_Temp.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText
    If Rs_Temp.RecordCount > 0 Then
            tot = tot + IIf(IsNull(Rs_Temp("tot").value), 0, Rs_Temp("tot").value)
    End If
        
    txtStudentCount1.Text = val(txtStudentCount1.Text) - tot
        
End Sub

Private Sub dcSchoolFile_Click(Area As Integer)
   Calc_SchoolRes
End Sub

Private Sub dcSchoolFile_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF3 Then
        Unload FrmSearch_BasicData
        FrmSearch_BasicData.SendForm = "VA"
        FrmSearch_BasicData.show vbModal
    End If
 
End Sub

Private Sub DTPicker1_Change()
        NourHijriCal1.value = ToHijriDate(DTPicker1.value)
End Sub

Private Sub DTPicker2_Change()
        NourHijriCal2.value = ToHijriDate(DTPicker2.value)
End Sub

Private Sub FgInstallments_AfterEdit(ByVal Row As Long, ByVal Col As Long)
    
    Dim StrAccountCode As String
    Dim Msg As String
    Dim rs As New ADODB.Recordset
    Dim StrSQL As String
    Dim ClsAcc As New ClsAccounts
    Dim LngRow As Long
    Dim sql As String
 
    With FgInstallments

Select Case .ColKey(Col)
  Case "CarNo"
              Add_Board val(.ComboData), Row, Col, (FgInstallments.TextMatrix(FgInstallments.Row, FgInstallments.ColIndex("CarNo")))
              
   Case "BoardNo"
              Add_Board val(.ComboData), Row, Col
              
              
  Case "Driver"
                
                StrAccountCode = " select * from TblEmployee  where  Emp_ID =    " & val(.ComboData)
                Set rs = Nothing
                rs.Open StrAccountCode, Cn, adOpenStatic, adLockReadOnly, adCmdText
                If Not (rs.BOF Or rs.EOF) Then
                     .TextMatrix(Row, .ColIndex("DriverID")) = IIf(IsNull(rs("Emp_ID").value), "", rs("Emp_ID").value)
                End If
    
    
   Case "count"
   
                If val(FgInstallments.TextMatrix(FgInstallments.Row, FgInstallments.ColIndex("count"))) > val(FgInstallments.TextMatrix(FgInstallments.Row, FgInstallments.ColIndex("allow"))) Then
                        MsgBox ("·« Ì„ﬂ‰ «‰ Ì Ã«Ê“  ⁄œœ  «·„ﬁ«⁄œ «·„Œ’’… «·„”„ÊÕ  »Â ·· Œ’Ì’")
                                FgInstallments.TextMatrix(FgInstallments.Row, FgInstallments.ColIndex("count")) = FgInstallments.TextMatrix(FgInstallments.Row, FgInstallments.ColIndex("allow"))
                End If
                If val(FgInstallments.TextMatrix(FgInstallments.Row, FgInstallments.ColIndex("count"))) > val(FgInstallments.TextMatrix(FgInstallments.Row, FgInstallments.ColIndex("carstudentcount"))) Then
                        MsgBox ("·« Ì„ﬂ‰ «‰ Ì Ã«Ê“  ”⁄… «·Õ«›·… ")
                                FgInstallments.TextMatrix(FgInstallments.Row, FgInstallments.ColIndex("count")) = FgInstallments.TextMatrix(FgInstallments.Row, FgInstallments.ColIndex("carstudentcount"))
                End If
                Cslculation_Count
              
    End Select
  End With
    

End Sub

Public Sub Cslculation_Count()
                 Cal_Student
                Cal_Student1
                If val(lbl(22)) < 0 Then
                    MsgBox ("·ﬁœ  Ã«Ê“  Õœ «· Œ’Ì’ «·„”„ÊÕ »… ··„œ—”…")
                      FgInstallments.TextMatrix(FgInstallments.Row, FgInstallments.ColIndex("count")) = "0"
                      Exit Sub
                End If
              
                If val(lbl(12)) < 0 Then
                    MsgBox ("·ﬁœ  Ã«Ê“  Õœ «· Œ’Ì’ «·„”„ÊÕ »… ··⁄ﬁœ")
                      FgInstallments.TextMatrix(FgInstallments.Row, FgInstallments.ColIndex("count")) = "0"
                      Exit Sub
                End If
                Cal_Student
                Cal_Student1


End Sub

Public Sub Add_Board(ID As Integer, Row As Long, Col As Long, Optional CarNo As String)

   Dim StrAccountCode As String
   Dim StrSQL As String
   
  With FgInstallments
   
  .TextMatrix(Row, .ColIndex("CarNo")) = 0
                .TextMatrix(Row, .ColIndex("Chasis")) = 0
                .TextMatrix(Row, .ColIndex("Remarks")) = ""
                .TextMatrix(Row, .ColIndex("Driver")) = ""
                .TextMatrix(Row, .ColIndex("Count")) = ""
                              
               StrAccountCode = ID
               StrSQL = "   SELECT dbo.TblCarsData.id,TblCarsData.MaxCap   ,dbo.TblCarsData.Branch_NO, dbo.TblCarsData.code, dbo.TblCarsData.Fullcode, dbo.TblCarsData.prifix, dbo.TblCarsData.CarsTypeId,"
               StrSQL = StrSQL & "       dbo.TblCarsData.LicenseNO, dbo.TblCarsData.Name, dbo.TblCarsData.BoardNO, dbo.TblCarsData.Model, dbo.TblCarsData.PurchaseDate, dbo.TblCarsData.LastKMCounter,"
               StrSQL = StrSQL & "       dbo.TblCarsData.InsuranceCompanyId, dbo.TblCarsData.LicenseExpireDate, dbo.TblCarsData.Emp_id, dbo.TblCarsData.InsuranceExpireDate,"
               StrSQL = StrSQL & "       dbo.TblCarsData.TestExpireDate, dbo.TblCarsData.Notes, dbo.TblCarsData.LicenseExpireDateH, dbo.TblCarsData.InsuranceExpireDateH, dbo.TblCarsData.TestExpireDateH,"
               StrSQL = StrSQL & "       dbo.TblCarsData.fixedAssetid, dbo.TblCarsData.VehicleLong, dbo.TblCarsData.EquQty, dbo.TblCarsData.Capacity, dbo.TblCarsData.ContractID,"
               StrSQL = StrSQL & "       dbo.TblCarsData.EndContractDate, dbo.TblCarsData.SetCount, dbo.TblCarsData.Rate, dbo.TblCarsData.EndContractDateH, dbo.TblEmployee.Emp_Namee,"
               StrSQL = StrSQL & "       dbo.TblEmployee.emp_Name , dbo.TblEmployee.Emp_mobile"
               StrSQL = StrSQL & "    FROM     dbo.TblCarsData LEFT OUTER JOIN"
             If CarNo = "" Then
               StrSQL = StrSQL & "     dbo.TblEmployee ON dbo.TblCarsData.Emp_id = dbo.TblEmployee.Emp_ID where   ID = '" & StrAccountCode & "'"
             Else
                StrSQL = StrSQL & "     dbo.TblEmployee ON dbo.TblCarsData.Emp_id = dbo.TblEmployee.Emp_ID where   TblCarsData.FullCode = '" & CarNo & "'"
             End If
             
               Set Rs_Temp = New ADODB.Recordset
               Rs_Temp.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

               If Not (Rs_Temp.BOF Or Rs_Temp.EOF) Then
               .TextMatrix(Row, .ColIndex("CarID")) = IIf(IsNull(Rs_Temp("ID").value), "", Rs_Temp("ID").value)
                    .TextMatrix(Row, .ColIndex("BoardNo")) = IIf(IsNull(Rs_Temp("BoardNo").value), "", Rs_Temp("BoardNo").value)
                    .TextMatrix(Row, .ColIndex("carstudentcount")) = IIf(IsNull(Rs_Temp("capacity").value), "", Rs_Temp("capacity").value)
                    .TextMatrix(Row, .ColIndex("Driver")) = IIf(IsNull(Rs_Temp("Emp_Name").value), "", Rs_Temp("Emp_Name").value)
                    .TextMatrix(Row, .ColIndex("DriverID")) = IIf(IsNull(Rs_Temp("Emp_ID").value), "", Rs_Temp("Emp_ID").value)
                    .TextMatrix(Row, .ColIndex("Tel")) = IIf(IsNull(Rs_Temp("Emp_mobile").value), "", Rs_Temp("Emp_mobile").value)
                    .TextMatrix(Row, .ColIndex("CarNo")) = IIf(IsNull(Rs_Temp("FullCode").value), "", Rs_Temp("FullCode").value)
                    .TextMatrix(Row, .ColIndex("MaxCap")) = IIf(IsNull(Rs_Temp("MaxCap").value), "", Rs_Temp("MaxCap").value)
                     .TextMatrix(Row, .ColIndex("Serial")) = .Row
                    .Rows = .Rows + 1
                End If
                
                
                If Not StrAccountCode = "" Then
                    cal_Cars StrAccountCode, Row
                End If
                
                

End With

End Sub

Private Sub cal_Cars(StrAccountCode As String, Row As Long)
    Dim StrSQL As String

     StrSQL = " select  d.CarID  , sum( d.studentcount ) sumation  ,v.DurationID  from TblVehicleAllocation  v , TblVehicleAllocation_Details d  "
     StrSQL = StrSQL & "  where v.IDVA = d.IDVA  and  type = 2 and   carid = " & val(StrAccountCode)
     StrSQL = StrSQL & " and v.DurationID = " & val(dcDuration.BoundText) & " and v.IDMC = " & val(dcMinistryContract.BoundText)
     If TxtModFlg.Text = "E" Then
            StrSQL = StrSQL & " and v.IDVA <>   " & val(txtIDVA.Text)
     End If
     StrSQL = StrSQL & " group by   d.CarID ,v.DurationID "
     Set Rs_Temp = New ADODB.Recordset
     Rs_Temp.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText
     Dim carall As Integer
     If Rs_Temp.RecordCount > 0 Then
              carall = IIf(IsNull(Rs_Temp("sumation").value), 0, Rs_Temp("sumation").value)
     End If
     
     
    Dim i  As Integer, j As Integer
    With FgInstallments
    For i = .FixedRows To FgInstallments.Rows - 1
        If i <> Row Then
              If .TextMatrix(i, .ColIndex("CarID")) = StrAccountCode Then
                        j = j + val(.TextMatrix(i, .ColIndex("count")))
              End If
        End If
    Next
    
     
    ' .TextMatrix(Row, .ColIndex("allow")) = val(.TextMatrix(Row, .ColIndex("carstudentcount"))) - carall
      .TextMatrix(Row, .ColIndex("allow")) = val(.TextMatrix(Row, .ColIndex("MaxCap"))) - carall - j
     End With
     Cslculation_Count


End Sub



Private Sub Cal_Student()
Dim i As Integer, j As Integer

i = 0

With FgInstallments
For j = 1 To .Rows - 1
    If .TextMatrix(j, .ColIndex("BoardNo")) <> "" Then
        i = i + val(.TextMatrix(j, .ColIndex("count")))
    End If
Next
End With

lbl(12).Caption = val(txtStudentCount.Text) - i
Label2.Caption = val(txtStudentCount.Text)
lbl(11).Caption = i

End Sub

Private Sub Cal_Student1()
Dim i As Integer, j As Integer

i = 0

With FgInstallments
For j = 1 To .Rows - 1
    If .TextMatrix(j, .ColIndex("BoardNo")) <> "" Then
    
        i = i + val(.TextMatrix(j, .ColIndex("count")))
    
    End If
    
Next
End With

lbl(22).Caption = val(txtStudentCount1.Text) - i
Label5.Caption = val(txtStudentCount1.Text)
lbl(21).Caption = i

End Sub




Private Sub FgInstallments_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
With FgInstallments
     
     
    If dcDuration.BoundText = "" Or dcMinistryContract.BoundText = "" Then
            MsgBox ("«Œ — «·”‰… «·œ—«”Ì… «Ê·«")
            Cancel = True
            Exit Sub
    End If
     
     If dcMinistryContract.BoundText = "" Then
            MsgBox ("«Œ —  «·⁄ﬁœ «·Ê“«—Ï «Ê·«")
            Cancel = True
            Exit Sub
    End If
     
  If dcSchoolFile.BoundText = "" Then
            MsgBox ("«Œ —  «·„œ—”… «Ê·«")
            Cancel = True
            Exit Sub
    End If
     
     
     Select Case .ColKey(Col)
     
    Case "CarNo"
         .ComboList = ""
         
    Case "Chasis"
            .ComboList = ""
            
    Case "count"
        .ComboList = ""
        
        Dim s As String
        s = FgInstallments.TextMatrix(FgInstallments.Row, .ColIndex("CarID"))
        If s <> "" Then
                cal_Cars s, FgInstallments.Row
        End If
        
        
    Case "Remarks"
            .ComboList = ""
    Case "Tel"
            .ComboList = ""
    Case "allow"
    .ComboList = ""
     Cancel = True
    Case "carstudentcount"
    .ComboList = ""
    Cancel = True
    
   Case "MaxCap"
   .ComboList = ""
    Cancel = True
    
            End Select
            
            End With
            
End Sub

Private Sub FgInstallments_KeyUp(KeyCode As Integer, Shift As Integer)
        If FgInstallments.Col = FgInstallments.ColIndex("BoardNo") Then
                   If KeyCode = vbKeyF3 Then
                            Unload FrmCasrShearches
                             FrmCasrShearches.SendForm = "VA"
                             Load FrmCasrShearches
                             FrmCasrShearches.show vbModal
                   End If
        End If
End Sub

Private Sub FgInstallments_StartEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
 Dim rs As New ADODB.Recordset
    Dim StrSQL  As String
    Dim StrAccountType As String
    Dim StrComboList As String
    Dim Msg As String
    
     With FgInstallments
     
     Select Case .ColKey(Col)
     
    
     Case "BoardNo"
                    
          StrSQL = "  SELECT  ID ,BoardNo  from TblCarsData ORDER BY ID "
          rs.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText
          StrComboList = FgInstallments.BuildComboList(rs, "BoardNo", "ID")
           If StrComboList <> "" Then
                 StrComboList = "|" & StrComboList
           End If
          .ComboList = StrComboList
          
       Case "Driver"
         ' StrSQL = "   select   EmpID, Emp_Name,  Emp_Code,  DrivValue, Emp_mobile  from tblCarDrivers ,TblEmployee where tblCarDrivers.EmpID = TblEmployee.Emp_ID "
         
          StrSQL = "  select   e.Emp_ID Emp_ID , e.Emp_Name   Emp_Name  from TblEmployee e, TblEmpJobsTypes  j"
          StrSQL = StrSQL & "   Where e.JobTypeID = j.JobTypeID"
          StrSQL = StrSQL & "     and  ( j.JobTypeName like '%”«∆ﬁ%'  or j.JobTypeNamee like '%driver%')"
          
          rs.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText
          StrComboList = FgInstallments.BuildComboList(rs, "Emp_Name", "Emp_ID")
           If StrComboList <> "" Then
                 StrComboList = "|" & StrComboList
           End If
          .ComboList = StrComboList
          FgInstallments.Rows = FgInstallments.Rows + 1
        
        
     End Select
   End With
   
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
    Dcombos.getCountriesGovernments Me.dcCity
   
   Dim str As String, str1     As String
   If SystemOptions.UserInterface = ArabicInterface Then
    str = "Select ID , Name from tblSchooleFile "
    str1 = "Select ID , Name   from TblManagerialArea "
   Else
   str = "Select ID , NameE from tblSchooleFile "
   str1 = "Select ID , NameE   from TblManagerialArea "
   End If
   fill_combo dcSchoolFile, str
   fill_combo DCVendor, str1
   
   str = " select ID ,Name  from TblDurations  "
   fill_combo dcDuration, str
   
   
    str = " select IDMC , MinistryContractNo  from tblministrycontract "
    fill_combo dcMinistryContract, str
  
   
   
    If SystemOptions.UserInterface = EnglishInterface Then
        SetInterface Me
        ChangeLang
    End If
    Dim My_SQL As String
    
    Resize_Form Me
'    AddTip
    Set rs = New ADODB.Recordset
    Dim StrSQL As String
    StrSQL = "SELECT  *  From TblVehicleAllocation"
    rs.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText
    Set Cmd(5).ButtonImage = mdifrmmain.ImgLstTree.ListImages("Search").Picture
  
        
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

 
   lbl(0).Caption = "No."
   lbl(3).Caption = " Name Ar"
   lbl(7).Caption = " Name En"
   Label3.Caption = "City"
   
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
    LogTextA = "     «·Œ—ÊÃ „‰ ‘«‘… " & "  Œ’Ì’ «·Õ«›·«   "
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

Private Sub NourHijriCal1_LostFocus()
         VBA.Calendar = vbCalGreg
    DTPicker1.value = ToGregorianDate(NourHijriCal1.value)
End Sub

Private Sub NourHijriCal2_LostFocus()
    VBA.Calendar = vbCalGreg
    DTPicker2.value = ToGregorianDate(NourHijriCal2.value)
End Sub

Private Sub Text6_Change()
retriveMinistry
End Sub

Private Sub Text6_KeyPress(KeyAscii As Integer)

If KeyAscii = 13 Then
   retriveMinistry
End If

End Sub

Private Sub retriveMinistry()
     Dim rs As New ADODB.Recordset
    Dim StrSQL As String
                
       If dcMinistryContract.BoundText = "" Then Exit Sub
       
      StrSQL = " select * from TblMinistryContract where IDMC =  " & dcMinistryContract.BoundText
      Set rs = Nothing
      rs.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

      If Not (rs.BOF Or rs.EOF) Then
      
            XPTxtBoxName.Text = IIf(IsNull(rs("Name").value), "", rs("Name").value)
            dcCity.BoundText = IIf(IsNull(rs("CityID").value), "", rs("CityID").value)
            DCVendor.BoundText = IIf(IsNull(rs("VendorID").value), "", rs("VendorID").value)
            dtpFromDate.value = IIf(IsNull(rs("FromDate").value), Date, rs("FromDate").value)
            dtpFromDate.value = IIf(IsNull(rs("ToDate").value), Date, rs("ToDate").value)
            dtpFromDateH.value = IIf(IsNull(rs("FromDateH").value), ToHijriDate(Date), rs("FromDateH").value)
            dtpFromDateH.value = IIf(IsNull(rs("ToDateH").value), ToHijriDate(Date), rs("ToDateH").value)
      Else
     '       Text6.text = ""
            XPTxtBoxName.Text = ""
            dcCity.BoundText = ""
            DCVendor.BoundText = ""
        
      End If

End Sub


Private Sub txtMinistryNo_Change()
Dim val1, val2

If txtMinistryNo.Text = "" Then Exit Sub
Dim str As String
    str = " select *  from tblschoolefile where ministerNo =   '" & txtMinistryNo.Text & "'"
    Set Rs_Temp = New ADODB.Recordset
    Rs_Temp.Open str, Cn, adOpenStatic, adLockOptimistic, adCmdText
    If Rs_Temp.RecordCount > 0 Then
        Rs_Temp.MoveFirst '
        dcSchoolFile.BoundText = IIf(IsNull(Rs_Temp("ID").value), "", Rs_Temp("ID").value)
     Else
        dcSchoolFile.BoundText = ""
    End If
 End Sub

Private Sub txtMinistryNo_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF3 Then
        Unload FrmSearch_BasicData
        FrmSearch_BasicData.SendForm = "VA"
        FrmSearch_BasicData.show
    End If
End Sub

Private Sub TxtModFlg_Change()
    On Error GoTo ErrTrap

    Select Case Me.TxtModFlg.Text

        Case "R"

            If SystemOptions.UserInterface = ArabicInterface Then
                Me.Caption = "»Ì«‰«   Œ’Ì’ «·Õ«›·« "
            Else
                Me.Caption = "Boxes Data"
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
        
           
            If rs.RecordCount < 1 Then
                Me.XPBtnMove(0).Enabled = False
                Me.XPBtnMove(1).Enabled = False
                Me.XPBtnMove(2).Enabled = False
                Me.XPBtnMove(3).Enabled = False
                Me.Cmd(1).Enabled = False
                Me.Cmd(4).Enabled = False
            End If
           
           C1Elastic2.Enabled = False
           FgInstallments.Enabled = False
        Case "N"

            If SystemOptions.UserInterface = ArabicInterface Then
                Me.Caption = " Œ’Ì’ «·Õ«›·« ( ÃœÌœ )"
            Else
                Me.Caption = "Boxes Data(New)"
            End If
        
            Me.Cmd(2).Enabled = True
            Me.Cmd(3).Enabled = True

            If SystemOptions.UserInterface = ArabicInterface Then
                Me.Caption = " Œ’Ì’ «·Õ«›·« ( ÃœÌœ )"
            Else
                Me.Caption = "Boxes Data(New)"
            End If
        
            Me.Cmd(0).Enabled = False
            Me.Cmd(1).Enabled = False
            Me.Cmd(4).Enabled = False
            Me.Cmd(9).Enabled = False
        
                   Me.XPBtnMove(0).Enabled = False
                   Me.XPBtnMove(1).Enabled = False
                   Me.XPBtnMove(2).Enabled = False
                   Me.XPBtnMove(3).Enabled = False
        
                
           C1Elastic2.Enabled = True
           FgInstallments.Enabled = True
        Case "E"

            If SystemOptions.UserInterface = ArabicInterface Then
                Me.Caption = " Œ’Ì’ «·Õ«›·«  (  ⁄œÌ· )"
            Else
                Me.Caption = "Boxes Data(Edit)"
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
            
          C1Elastic2.Enabled = True
           FgInstallments.Enabled = True
       
    End Select

    Exit Sub
ErrTrap:
End Sub

Public Sub Retrive(Optional Lngid As Long = 0)

    Dim SngCusBegainAccount As Single

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
        rs.find "IDVA =" & Lngid, , adSearchForward, adBookmarkFirst
        If rs.BOF Or rs.EOF Then
            Exit Sub
        End If
    End If
    Clear_SumationLabels

    Me.txtIDVA.Text = IIf(IsNull(rs("IDVA").value), "", rs("IDVA").value)
    txtProcess.Text = txtIDVA.Text
    txtStudentCount.Text = IIf(IsNull(rs("StudentCount").value), "", rs("StudentCount").value)
    dtpProcessDate.value = IIf(IsNull(rs("ProcessDate").value), Date, rs("ProcessDate").value)
    txtMinistryNo.Text = IIf(IsNull(rs("MinistryNo").value), "", rs("MinistryNo").value)
    dcSchoolFile.BoundText = IIf(IsNull(rs("SchoolFileID").value), "", (rs("SchoolFileID").value))
    txtDepend.Text = IIf(IsNull(rs("Depend").value), "", rs("Depend").value)
    dtpFromDate.value = IIf(IsNull(rs("FromDate").value), Date, rs("FromDate").value)
    dtpToDate.value = IIf(IsNull(rs("FromDate").value), Date, rs("FromDate").value)
    dtpFromDateH.value = IIf(IsNull(rs("FromDateH").value), ToHijriDate(Date), rs("FromDateH").value)
    dtpToDateH.value = IIf(IsNull(rs("ToDateH").value), ToHijriDate(Date), rs("ToDateH").value)
    dcDuration.BoundText = IIf(IsNull(rs("DurationID").value), "", (rs("DurationID").value))
    dcMinistryContract.BoundText = IIf(IsNull(rs("IDMC").value), "", rs("IDMC").value)
 
    lbl(12).Caption = IIf(IsNull(rs("StudentAttrib").value), "", (rs("StudentAttrib").value))
    lbl(11).Caption = IIf(IsNull(rs("StudentAlloc").value), "", (rs("StudentAlloc").value))
    
     XPTxtCurrent.Caption = rs.AbsolutePosition
     XPTxtCount.Caption = rs.RecordCount
    
    
       Dim j As Integer
       Dim VendorSQL As String

       VendorSQL = " SELECT dbo.TblEmployee.Emp_Name AS Receipnt, dbo.TblEmployee.Emp_Code, dbo.TblEmployee.Emp_mobile, dbo.TblVehicleAllocation_Details.DriverID, "
       VendorSQL = VendorSQL & "  dbo.TblCarsData.Name AS CarName, dbo.TblVehicleAllocation_Details.CarID, dbo.TblCarsData.BoardNO,dbo.TblCarsData.FullCode, dbo.TblVehicleAllocation_Details.IDVA,"
       VendorSQL = VendorSQL & "              dbo.TblVehicleAllocation_Details.Type , dbo.TblVehicleAllocation_Details.studentcount"
       VendorSQL = VendorSQL & "  ,  VehicleSiteCount  ,VehicleAvailableSite  , TblVehicleAllocation_Details.MaxCap"
       VendorSQL = VendorSQL & "  FROM     dbo.TblEmployee RIGHT OUTER JOIN"
       VendorSQL = VendorSQL & "             dbo.TblCarsData INNER JOIN"
       VendorSQL = VendorSQL & "             dbo.TblVehicleAllocation_Details ON dbo.TblCarsData.id = dbo.TblVehicleAllocation_Details.CarID ON dbo.TblEmployee.Emp_ID = dbo.TblVehicleAllocation_Details.DriverID"
       VendorSQL = VendorSQL & "   where Type = 2 and  IDVA = " & (txtIDVA.Text)
      
       Set rsVendor = New ADODB.Recordset
       rsVendor.Open VendorSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText
  '     rsVendor.MoveFirst
       
       FgInstallments.Rows = 1
       With FgInstallments
       FgInstallments.Rows = rsVendor.RecordCount + 1
       For j = 1 To rsVendor.RecordCount
              .TextMatrix(j, .ColIndex("Serial")) = j
              .TextMatrix(j, .ColIndex("CarID")) = IIf(IsNull(rsVendor("CarID").value), "", rsVendor("CarID").value)
              .TextMatrix(j, .ColIndex("DriverID")) = IIf(IsNull(rsVendor("DriverID").value), "", rsVendor("DriverID").value)
            '  .TextMatrix(j, .ColIndex("Driver")) = IIf(IsNull(rsVendor("Emp_Name").value), "", rsVendor("Emp_Name").value)
              .TextMatrix(j, .ColIndex("Driver")) = IIf(IsNull(rsVendor("Receipnt").value), "", rsVendor("Receipnt").value)
.TextMatrix(j, .ColIndex("BoardNo")) = IIf(IsNull(rsVendor("BoardNo").value), "", rsVendor("BoardNo").value)

              .TextMatrix(j, .ColIndex("CarNo")) = IIf(IsNull(rsVendor("FullCode").value), "", rsVendor("FullCode").value)
              .TextMatrix(j, .ColIndex("count")) = IIf(IsNull(rsVendor("StudentCount").value), "", rsVendor("StudentCount").value)
              .TextMatrix(j, .ColIndex("carstudentcount")) = IIf(IsNull(rsVendor("VehicleSiteCount").value), "", rsVendor("VehicleSiteCount").value)
              .TextMatrix(j, .ColIndex("allow")) = IIf(IsNull(rsVendor("VehicleAvailableSite").value), "", rsVendor("VehicleAvailableSite").value)
              .TextMatrix(j, .ColIndex("MaxCap")) = IIf(IsNull(rsVendor("MaxCap").value), "", rsVendor("MaxCap").value)
                           
              rsVendor.MoveNext
       Next
       End With
       Cal_Student
        Cal_Student1
    Exit Sub
ErrTrap:
End Sub

Private Sub txtStudentCount_Change()
Cal_Student
End Sub

Private Sub txtStudentCount1_Change()
Cal_Student1
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
  '  LogTextA = "  Õ›Ÿ ‘«‘… " & " »Ì«‰«   «·Œ“‰ Ê «·⁄Âœ " & Chr(13) & " ﬂÊœ «·Œ“Ì‰…/«·⁄ÂœÂ  " & XPTxtBoxID.text & Chr(13) & " «·›—⁄ " & dcBranch.text & Chr(13) & "«·«”„ ⁄—»Ì  " & XPTxtBoxName & Chr(13) & " «·‰Ê⁄     "
'
'    If Option1.value = True Then
'        LogTextA = LogTextA & " Œ“Ì‰…  "
'    ElseIf Option2.value = True Then
'        LogTextA = LogTextA & "  ⁄Âœ…  "
'
'    End If
'
'    LogTextA = LogTextA & Chr(13) & "› Õ Õ«›Ÿ… «·‘Ìﬂ«  "
'
'    If chkChequeBox.value = vbChecked Then
'        LogTextA = LogTextA & "‰⁄„ "
'    Else
'        LogTextA = LogTextA & "·« "
'    End If
'
'    LogTextA = LogTextA & Chr(13) & "«”„ «·„ÊŸ›   " & dcEmp.text
'
'    LogTextA = LogTextA & Chr(13) & " ÿ»Ì⁄Â «·—’Ìœ «·«›  «ÕÌ   "
'
'    If OptType(0).value = True Then
'        LogTextA = LogTextA & "„œÌ‰"
''    ElseIf OptType(1).value = True Then
'        LogTextA = LogTextA & "œ«∆‰"
'    ElseIf OptType(2).value = True Then
'        LogTextA = LogTextA & "€Ì— „Õœœ"
'    End If
'
'    LogTextA = LogTextA & Chr(13) & " ﬁÌ„… «·—’Ìœ «·«›  «ÕÌ  " & TxtOpenBalance
'    LogTextA = LogTextA & Chr(13) & "„·«ÕŸ«    " & XPMTxtRemark
'
'    'sssssssssssssssssss
'    LogTextE = "  Save Screen " & " Boxes Data " & Chr(13) & " Code " & XPTxtBoxID.text & Chr(13) & " Branch " & dcBranch.text & Chr(13) & "«Name  " & XPTxtBoxName & Chr(13) & " Type     "
'
'    If Option1.value = True Then
'        LogTextE = LogTextE & " Box  "
'    ElseIf Option2.value = True Then
'        LogTextE = LogTextE & "  Era  "
'
'    End If
'
'    LogTextE = LogTextE & Chr(13) & "Open CHeque Box "
'
'    If chkChequeBox.value = vbChecked Then
'        LogTextE = LogTextE & "Yes "
'    Else
'        LogTextE = LogTextE & "No "
'    End If
'
'    LogTextE = LogTextE & Chr(13) & " Employee Name" & dcEmp.text
'
'    LogTextE = LogTextE & Chr(13) & "Opening Balance Type"
'
'    If OptType(0).value = True Then
'        LogTextE = LogTextE & "Debit"
'    ElseIf OptType(1).value = True Then
'        LogTextE = LogTextE & "Credit"
'    ElseIf OptType(2).value = True Then
'        LogTextE = LogTextE & "Na"
'    End If
'
'    LogTextE = LogTextE & Chr(13) & " Opening Balance  Value " & TxtOpenBalance
'    LogTextE = LogTextE & Chr(13) & " Remarks   " & XPMTxtRemark
'
'    If Currentmode <> "D" Then
'        AddToLogFile CInt(user_id), 0, Date, Time, LogTextA, LogTextE, Me.name, Me.TxtModFlg, "", ""
'    Else
'        AddToLogFile CInt(user_id), 0, Date, Time, LogTextA, LogTextE, Me.name, "D", "", ""
'    End If

End Function
 
Private Sub SaveData()
    Dim Msg As String
    Dim StrSQL As String
    
    Dim BeginTrans As Boolean
   ' On Error GoTo ErrTrap

    If Me.TxtModFlg.Text <> "R" Then
    
        If dcSchoolFile.BoundText = "" Then
            MsgBox "„‰ ›÷·ﬂ √œŒ· «Œ — «·„œ—”…", vbOKOnly + vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
            dcSchoolFile.SetFocus
        'attia    SendKeys ("{F4}")
            Exit Sub
        End If
        
       If dcDuration.BoundText = "" Then
            MsgBox ("«Œ — «·”‰… «·œ—«”Ì… «Ê·«")
            'dcDuration.backcolor = RGB(255, 0, 0)
            Exit Sub
       End If
       
        If dcMinistryContract.BoundText = "" Then
            MsgBox ("«Œ — «·⁄ﬁœ «·Ê“«—Ï «Ê·«")
            'dcDuration.backcolor = RGB(255, 0, 0)
            Exit Sub
       End If
                
       If val(lbl(22).Caption) < 0 Then
            MsgBox ("·« Ì„ﬂ‰ «‰ Ì Ã«Ê“ ⁄œœ «·ÿ·»Â «·–Ì‰  „  ”ﬂÌ‰Â„ ⁄œœ ÿ·»… «·„œ—”…")
            Exit Sub
       End If
                
      If val(lbl(21).Caption) <= 0 Then
            MsgBox ("·«»œ „‰  ”ﬂÌ‰ ÿ·»… ﬁ»· « „«„ ⁄„·Ì… «·Õ›Ÿ")
            Exit Sub
       End If
                
         
       Select Case Me.TxtModFlg.Text
            Case "N"
            rs.AddNew
            txtIDVA.Text = CStr(new_id("TblVehicleAllocation", "IDVA", "", True))
            Case "E"
              StrSQL = "delete From TblVehicleAllocation_Details  where type = 2 and  IDVA =" & (txtIDVA.Text)
            Cn.Execute StrSQL, , adExecuteNoRecords
       End Select

        Cn.BeginTrans
        BeginTrans = True
        txtIDVA.Text = new_id("TblVehicleAllocation", "IDVA", "", True)
        rs("IDVA").value = (txtIDVA.Text)
        rs("ProcessNo").value = txtIDVA.Text
        rs("StudentCount").value = val(txtStudentCount.Text)
        rs("Depend").value = txtDepend.Text
        rs("ProcessDate").value = dtpProcessDate.value
        rs("MinistryNo").value = txtMinistryNo.Text
        rs("SchoolFileID").value = IIf(dcSchoolFile.BoundText = "", Null, dcSchoolFile.BoundText)
        rs("StudentCount").value = (txtStudentCount.Text)
        rs("Fromdate").value = IIf(IsNull(dtpFromDate.value), Date, dtpFromDate.value)
        rs("Todate").value = IIf(IsNull(dtpFromDate.value), Date, dtpFromDate.value)
        rs("FromdateH").value = IIf(IsNull(dtpFromDateH.value), ToHijriDate(Date), dtpFromDateH.value)
        rs("TodateH").value = IIf(IsNull(dtpToDate.value), ToHijriDate(Date), dtpToDateH.value)
        rs("DurationID").value = val(dcDuration.BoundText)
        rs("StudentAttrib").value = val(lbl(12).Caption)
        rs("StudentAlloc").value = val(lbl(11).Caption)
        rs("IDMC").value = dcMinistryContract.BoundText
        
        rs.update
        
        Dim j As Integer
        Dim rs_det As ADODB.Recordset
        Set rs_det = New ADODB.Recordset
        rs_det.Open "TblVehicleAllocation_Details", Cn, adOpenStatic, adLockOptimistic, adCmdTable
        
        With FgInstallments
         For j = 1 To FgInstallments.Rows - 1
            If .TextMatrix(j, .ColIndex("BoardNo")) <> "" Then
                rs_det.AddNew
                rs_det("ID").value = new_id("TblVehicleAllocation_Details", "ID", "", True)
                rs_det("IDVA").value = val(txtIDVA.Text)
                rs_det("CarID").value = IIf(.TextMatrix(j, .ColIndex("CarID")) = "", Null, .TextMatrix(j, .ColIndex("CarID")))
                rs_det("StudentCount").value = IIf(.TextMatrix(j, .ColIndex("count")) = "", 0, val(.TextMatrix(j, .ColIndex("count"))))
                rs_det("DriverID").value = IIf(.TextMatrix(j, .ColIndex("DriverID")) = "", Null, .TextMatrix(j, .ColIndex("DriverID")))
                
                rs_det("VehicleSiteCount").value = IIf(.TextMatrix(j, .ColIndex("carstudentcount")) = "", 0, val(.TextMatrix(j, .ColIndex("carstudentcount"))))
                rs_det("VehicleAvailableSite").value = IIf(.TextMatrix(j, .ColIndex("allow")) = "", Null, .TextMatrix(j, .ColIndex("allow")))
                rs_det("MaxCap").value = IIf(.TextMatrix(j, .ColIndex("MaxCap")) = "", Null, .TextMatrix(j, .ColIndex("MaxCap")))
                rs_det("Type").value = 2
                rs_det.update
            End If
         Next
        End With
                
        Cn.CommitTrans
        BeginTrans = False
        XPTxtCurrent.Caption = rs.AbsolutePosition
        XPTxtCount.Caption = rs.RecordCount
       'CuurentLogdata

        Select Case Me.TxtModFlg.Text

            Case "N"

                If SystemOptions.UserInterface = ArabicInterface Then
                    Msg = "  „ Õ›Ÿ «·»Ì«‰«   " & CHR(13)
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
          '  rs.find "ID='" & val(XPTxtBoxID.text) & "'", , adSearchForward, adBookmarkFirst

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
            
        If txtIDVA.Text <> "" Then

    
        Msg = "”Ì „ Õ–› »Ì«‰«   Œ’Ì’ «·Õ«›·«  —ﬁ„ " & CHR(13)
        Msg = Msg + (txtIDVA.Text) & CHR(13)
        Msg = Msg + " Â·  —€» ›Ì Õ–› Â–Â «·»Ì«‰« ø"

        If MsgBox(Msg, vbYesNo + vbQuestion + vbMsgBoxRight + vbMsgBoxRtlReading, App.title) = vbYes Then
            
    
            If Not rs.RecordCount < 1 Then
            
                StrSQL = "delete From TblVehicleAllocation_Details where type =2 and   IDVA  =" & val(txtIDVA.Text)
                Cn.Execute StrSQL, , adExecuteNoRecords
                        
                StrSQL = "delete From TblVehicleAllocation where  IDVA  =" & val(txtIDVA.Text)
                Cn.Execute StrSQL, , adExecuteNoRecords
                
                   StrSQL = "SELECT  *  From  TblVehicleAllocation "
                   rs.Close
                   rs.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText
                

                If rs.RecordCount < 1 Then
                    clear_all Me
                    TxtModFlg_Change
                    XPTxtCurrent.Caption = 0
                    XPTxtCount.Caption = 0
                    Clear_SumationLabels
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
    Msg = "·«Ì„ﬂ‰ Õ–› Â–« «·”Ã· · ﬂ«„· «·»Ì«‰«  " & CHR(13) & " ÊÃœ »Ì«‰«  „— »ÿ… » Œ’Ì’ «·Õ«›·«  "
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
        .Create Me.hwnd, "»Ì«‰«  «·Œ“‰", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
        .AddControl Cmd(0), "ÃœÌœ ..." & Wrap & "·«÷«›… »Ì«‰«   Œ’Ì’ ÃœÌœ" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
    End With

    With TTP
        .Create Me.hwnd, "»Ì«‰«   Œ’Ì’ «·Õ«›·«  ", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
        .AddControl Cmd(1), " ⁄œÌ· ..." & Wrap & "· ⁄œÌ· »Ì«‰«   Œ’Ì’ «·Õ«›·« " & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
    End With

    With TTP
        .Create Me.hwnd, "»Ì«‰«  «·Œ“‰", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
        .AddControl Cmd(2), "Õ›Ÿ ..." & Wrap & "·Õ›Ÿ »Ì«‰«   Œ’Ì’ «·Õ«›·«  «·ÃœÌœ" & Wrap & "·Õ›Ÿ «· ⁄œÌ·« " & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
    End With

    With TTP
        .Create Me.hwnd, "»Ì«‰«  «·Œ“‰", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
        .AddControl Cmd(3), " —«Ã⁄ ..." & Wrap & "·· —«Ã⁄ ⁄‰ ⁄„·Ì… «·«÷«›…" & Wrap & "··· —«Ã⁄ ⁄‰ ⁄„·Ì… «· ⁄œÌ·" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
    End With

    With TTP
        .Create Me.hwnd, "»Ì«‰«  «·Œ“‰", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
        .AddControl Cmd(4), "Õ–› ..." & Wrap & "·Õ–› »Ì«‰«   Œ’Ì’ «·Õ«›·« " & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
    End With

    With TTP
        .Create Me.hwnd, "»Ì«‰«  «·Œ“‰", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
        .AddControl Cmd(5), "»ÕÀ ..." & Wrap & "···»ÕÀ ⁄‰  Œ’Ì’ «·Õ«›·«  " & Wrap & "Ì‰ÿ»ﬁ ⁄·ÌÂ ‘—Êÿ „⁄Ì‰…" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
    End With

    With TTP
        .Create Me.hwnd, "»Ì«‰«   Œ’Ì’ «·Õ«›·« ", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
        .AddControl Cmd(6), "Œ—ÊÃ ..." & Wrap & "·«€·«ﬁ Â–Â «·‰«›–…" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
    End With

    With TTP
        .Create Me.hwnd, "»Ì«‰«   Œ’Ì’ «·Õ«›·« ", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
        .AddControl XPBtnMove(1), "«·√Ê· ..." & Wrap & "··«‰ ﬁ«· «·Ï √Ê· ”Ã·" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
    End With

    With TTP
        .Create Me.hwnd, "»Ì«‰«   Œ’Ì’ «·Õ«›·« ", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
        .AddControl XPBtnMove(0), "«·”«»ﬁ ..." & Wrap & "··«‰ ﬁ«· «·Ï «·”Ã· «·”«»ﬁ" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
    End With

    With TTP
        .Create Me.hwnd, "»Ì«‰«   Œ’Ì’ «·Õ«›·« ", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
        .AddControl XPBtnMove(3), "«· «·Ì ..." & Wrap & "··«‰ ﬁ«· «·Ï «·”Ã· «· «·Ì" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
    End With

    With TTP
        .Create Me.hwnd, "»Ì«‰«   Œ’Ì’ «·Õ«›·« ", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
        .AddControl XPBtnMove(2), "«·√ŒÌ— ..." & Wrap & "··«‰ ﬁ«· «·Ï ¬Œ— ”Ã·" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
    End With

    With TTP
        .Create Me.hwnd, "»Ì«‰«   Œ’Ì’ «·Õ«›·« ", 1, 15204351, -2147483630
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


Function print_report()
    
    Dim MySQL As String
    Dim RsData As New ADODB.Recordset
    Dim xApp As New CRAXDRT.Application
    Dim xReport As CRAXDRT.Report
    Dim CViewer As ClsReportViewer
    Dim StrReportTitle As String
    Dim StrFileName As String
    Dim Msg As String


MySQL = MySQL & "   SELECT dbo.TblCarsData.BoardNO, dbo.TblVehicleAllocation.DurationID, dbo.TblVehicleAllocation.MinistryNo, dbo.TblVehicleAllocation.StudentAlloc,"
MySQL = MySQL & "   dbo.TblVehicleAllocation.StudentAttrib, dbo.TblVehicleAllocation.ToDateH, dbo.TblVehicleAllocation.FromDateH, dbo.TblVehicleAllocation.ToDate,"
MySQL = MySQL & "   dbo.TblVehicleAllocation.FromDate, dbo.TblVehicleAllocation.StudentCount, dbo.TblVehicleAllocation.IDMC, dbo.TblVehicleAllocation.SchoolFileID,"
MySQL = MySQL & "   dbo.TblVehicleAllocation.ProcessNo, dbo.TblVehicleAllocation.IDVA, dbo.TblSchooleFile.Name AS SchoolName, dbo.TblSchooleFile.StudentCount AS SchoolStudentCount,"
MySQL = MySQL & "   dbo.TblSchooleFile.ministerNo AS SchoolministerNo, dbo.TblDurations.Name AS DurationName, dbo.TblDurations.FromDate AS DurFromDate,"
MySQL = MySQL & "   dbo.TblDurations.FromDateH AS DurFromDateH, dbo.TblDurations.ToDate AS DurToDate, dbo.TblDurations.TODateH AS DurToDateH, dbo.TblEmployee.Emp_Name,"
MySQL = MySQL & "   dbo.TblEmployee.Emp_Name1, dbo.TblEmployee.Fullcode, dbo.TblVehicleAllocation_Details.Type, dbo.TblMinistryContract.MinistryContractNo,"
MySQL = MySQL & "   dbo.TblVehicleAllocation.ProcessDate, dbo.TblVehicleAllocation_Details.Capecity, dbo.TblVehicleAllocation_Details.MaxCap, dbo.TblCarsData.code,"
MySQL = MySQL & "   dbo.TblVehicleAllocation_Details.StudentCount AS StudentCount1, dbo.TblVehicleAllocation_Details.VehicleAvailableSite, dbo.TblEmployee.NumEkama,"
MySQL = MySQL & "   dbo.TblEmployee.Emp_Phone , dbo.TblEmployee.Emp_mobile, dbo.TblVehicleAllocation_Details.CarID , dbo.TblVehicleAllocation_Details.VehicleSiteCount"
MySQL = MySQL & "   FROM     dbo.TblMinistryContract INNER JOIN"
MySQL = MySQL & "   dbo.TblDurations INNER JOIN"
MySQL = MySQL & "   dbo.TblVehicleAllocation ON dbo.TblDurations.ID = dbo.TblVehicleAllocation.DurationID INNER JOIN"
MySQL = MySQL & "   dbo.TblSchooleFile ON dbo.TblVehicleAllocation.SchoolFileID = dbo.TblSchooleFile.ID INNER JOIN"
MySQL = MySQL & "   dbo.TblVehicleAllocation_Details ON dbo.TblVehicleAllocation.IDVA = dbo.TblVehicleAllocation_Details.IDVA ON"
MySQL = MySQL & "   dbo.TblMinistryContract.IDMC = dbo.TblVehicleAllocation.IDMC INNER JOIN"
MySQL = MySQL & "   dbo.TblCarsData ON dbo.TblVehicleAllocation_Details.CarID = dbo.TblCarsData.id left outer JOIN"
MySQL = MySQL & "   dbo.TblEmployee ON dbo.TblVehicleAllocation_Details.DriverID = dbo.TblEmployee.Emp_ID"
MySQL = MySQL & "   WHERE   dbo.TblVehicleAllocation_Details.Type = 2 "





  MySQL = MySQL & "   and  TblVehicleAllocation.IDVA = " & val(txtIDVA.Text)
     
    
    


    If SystemOptions.UserInterface = ArabicInterface Then
        StrFileName = App.path & "\Reports\REPORTS NEW\" & "rpt_VehicleAllocationHeader.rpt"
    Else
        StrFileName = App.path & "\Reports\REPORTS NEW\" & "rpt_VehicleAllocationHeader.rpt"
    End If

    If Dir(StrFileName) = "" Then

        Screen.MousePointer = vbDefault
        Exit Function
    End If

    Set RsData = New ADODB.Recordset
    RsData.Open MySQL, Cn, adOpenStatic, adLockReadOnly, adCmdText
 
    Screen.MousePointer = vbArrowHourglass
    Set xReport = xApp.OpenReport(StrFileName)
    Dim cCompanyInfo As New ClsCompanyInfo

    xReport.EnableParameterPrompting = False
    Set CViewer = New ClsReportViewer
    CViewer.FireReport xReport, WindowTarget, "", , , , StrFileName

    RsData.Close
    Set RsData = Nothing
    Screen.MousePointer = vbDefault



End Function

'//////////////////////////////////////////////////////////////////////////////////////


Function print_VA(Optional Opt As Integer = 1)
    
     
    Dim MySQL As String
    Dim RsData As New ADODB.Recordset
    Dim xApp As New CRAXDRT.Application
    Dim xReport As CRAXDRT.Report
    Dim CViewer As ClsReportViewer
    Dim StrReportTitle As String
    Dim StrFileName As String
    Dim Msg As String
   

MySQL = MySQL & "   SELECT dbo.TblCarsData.BoardNO, dbo.TblVehicleAllocation.DurationID, dbo.TblVehicleAllocation.MinistryNo, dbo.TblVehicleAllocation.StudentAlloc,"
MySQL = MySQL & "   dbo.TblVehicleAllocation.StudentAttrib, dbo.TblVehicleAllocation.ToDateH, dbo.TblVehicleAllocation.FromDateH, dbo.TblVehicleAllocation.ToDate,"
MySQL = MySQL & "   dbo.TblVehicleAllocation.FromDate, dbo.TblVehicleAllocation.StudentCount, dbo.TblVehicleAllocation.IDMC, dbo.TblVehicleAllocation.SchoolFileID,"
MySQL = MySQL & "   dbo.TblVehicleAllocation.ProcessNo, dbo.TblVehicleAllocation.IDVA, dbo.TblSchooleFile.Name AS SchoolName, dbo.TblSchooleFile.StudentCount AS SchoolStudentCount,"
MySQL = MySQL & "   dbo.TblSchooleFile.ministerNo AS SchoolministerNo, dbo.TblDurations.Name AS DurationName, dbo.TblDurations.FromDate AS DurFromDate,"
MySQL = MySQL & "   dbo.TblDurations.FromDateH AS DurFromDateH, dbo.TblDurations.ToDate AS DurToDate, dbo.TblDurations.TODateH AS DurToDateH, dbo.TblEmployee.Emp_Name,"
MySQL = MySQL & "   dbo.TblEmployee.Emp_Name1, dbo.TblEmployee.Fullcode, dbo.TblVehicleAllocation_Details.Type, dbo.TblMinistryContract.MinistryContractNo,"
MySQL = MySQL & "   dbo.TblVehicleAllocation.ProcessDate, dbo.TblVehicleAllocation_Details.Capecity, dbo.TblVehicleAllocation_Details.MaxCap, dbo.TblCarsData.code,"
MySQL = MySQL & "   dbo.TblVehicleAllocation_Details.StudentCount AS StudentCount1, dbo.TblVehicleAllocation_Details.VehicleAvailableSite, dbo.TblEmployee.NumEkama,"
MySQL = MySQL & "   dbo.TblEmployee.Emp_Phone , dbo.TblEmployee.Emp_mobile, dbo.TblVehicleAllocation_Details.CarID , dbo.TblVehicleAllocation_Details.VehicleSiteCount"
MySQL = MySQL & "   FROM     dbo.TblMinistryContract INNER JOIN"
MySQL = MySQL & "   dbo.TblDurations INNER JOIN"
MySQL = MySQL & "   dbo.TblVehicleAllocation ON dbo.TblDurations.ID = dbo.TblVehicleAllocation.DurationID INNER JOIN"
MySQL = MySQL & "   dbo.TblSchooleFile ON dbo.TblVehicleAllocation.SchoolFileID = dbo.TblSchooleFile.ID INNER JOIN"
MySQL = MySQL & "   dbo.TblVehicleAllocation_Details ON dbo.TblVehicleAllocation.IDVA = dbo.TblVehicleAllocation_Details.IDVA ON"
MySQL = MySQL & "   dbo.TblMinistryContract.IDMC = dbo.TblVehicleAllocation.IDMC INNER JOIN"
MySQL = MySQL & "   dbo.TblCarsData ON dbo.TblVehicleAllocation_Details.CarID = dbo.TblCarsData.id left outer JOIN"
MySQL = MySQL & "   dbo.TblEmployee ON dbo.TblVehicleAllocation_Details.DriverID = dbo.TblEmployee.Emp_ID"
MySQL = MySQL & "   WHERE   dbo.TblVehicleAllocation_Details.Type = 2 "





  MySQL = MySQL & "   and  TblVehicleAllocation.IDVA = " & val(txtIDVA.Text)
     
    
    MySQL = MySQL & "   order by  TblVehicleAllocation.IDVA "



    If SystemOptions.UserInterface = ArabicInterface Then
        StrFileName = App.path & "\Reports\REPORTS NEW\" & "rpt_VehicleAllocationHeader.rpt"
    Else
        StrFileName = App.path & "\Reports\REPORTS NEW\" & "rpt_VehicleAllocationHeader.rpt"
    End If




'MySQL = " order by TblAttributionContract.idac "
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


