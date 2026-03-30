VERSION 5.00
Object = "{C115893A-A3BF-43AF-B28D-69DB846077F3}#1.0#0"; "vsflex8u.ocx"
Object = "{0AFE7BE0-11B7-4A3E-978D-D4501E9A57FE}#1.0#0"; "c1sizer.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{FE5DCFAD-BC1D-11D2-94CF-004005455FAA}#1.4#0"; "ImpulseButton.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form FrmVehicleAllocation 
   BackColor       =   &H00E2E9E9&
   BorderStyle     =   1  'Fixed Single
   Caption         =   " Œ’Ì’ «·Õ«›·« "
   ClientHeight    =   9840
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   15075
   Icon            =   "VehicleAllocation.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   RightToLeft     =   -1  'True
   ScaleHeight     =   9840
   ScaleWidth      =   15075
   WindowState     =   2  'Maximized
   Begin C1SizerLibCtl.C1Elastic C1Elastic1 
      Height          =   9840
      Left            =   0
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   0
      Width           =   15075
      _cx             =   26591
      _cy             =   17357
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
         Height          =   1425
         Left            =   120
         TabIndex        =   19
         TabStop         =   0   'False
         Top             =   7140
         Width           =   3975
         _cx             =   7011
         _cy             =   2514
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
            Height          =   375
            Index           =   11
            Left            =   120
            RightToLeft     =   -1  'True
            TabIndex        =   25
            Top             =   120
            Width           =   855
         End
         Begin VB.Label Label2 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "0"
            ForeColor       =   &H000000C0&
            Height          =   390
            Left            =   450
            TabIndex        =   24
            Top             =   540
            Width           =   510
         End
         Begin VB.Label Lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "0"
            ForeColor       =   &H000000C0&
            Height          =   285
            Index           =   12
            Left            =   120
            RightToLeft     =   -1  'True
            TabIndex        =   23
            Top             =   1035
            Width           =   855
         End
         Begin VB.Label Lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "«Ã„«·Ï «·ÿ·«» «·–Ì‰  „  ”ﬂÌ‰Â„"
            ForeColor       =   &H000000C0&
            Height          =   375
            Index           =   6
            Left            =   630
            RightToLeft     =   -1  'True
            TabIndex        =   22
            Top             =   120
            Width           =   3090
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "⁄œœ «·ÿ·«» «·„ ⁄«ﬁœ ⁄·ÌÂ„"
            ForeColor       =   &H000000C0&
            Height          =   390
            Index           =   1
            Left            =   1590
            TabIndex        =   21
            Top             =   540
            Width           =   2130
         End
         Begin VB.Label Lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "«·›—ﬁ «·ﬁ«»· ··«”‰«œ ··„ ⁄ÂœÌ‰"
            ForeColor       =   &H000000C0&
            Height          =   285
            Index           =   10
            Left            =   1185
            RightToLeft     =   -1  'True
            TabIndex        =   20
            Top             =   1035
            Width           =   2550
         End
      End
      Begin C1SizerLibCtl.C1Elastic C1Elastic2 
         Height          =   2280
         Left            =   7395
         TabIndex        =   7
         TabStop         =   0   'False
         Top             =   840
         Width           =   7515
         _cx             =   13256
         _cy             =   4022
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
         Begin VB.TextBox XPTxtBoxNamee 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   3705
            MaxLength       =   50
            RightToLeft     =   -1  'True
            TabIndex        =   54
            Top             =   255
            Width           =   1845
         End
         Begin VB.TextBox Text7 
            Alignment       =   1  'Right Justify
            Height          =   270
            Left            =   3705
            MaxLength       =   50
            RightToLeft     =   -1  'True
            TabIndex        =   13
            Top             =   585
            Width           =   1845
         End
         Begin VB.TextBox Text6 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   3705
            MaxLength       =   50
            RightToLeft     =   -1  'True
            TabIndex        =   9
            Top             =   855
            Width           =   1845
         End
         Begin VB.TextBox Text8 
            Alignment       =   1  'Right Justify
            Enabled         =   0   'False
            Height          =   255
            Left            =   195
            MaxLength       =   50
            RightToLeft     =   -1  'True
            TabIndex        =   8
            Top             =   855
            Width           =   1860
         End
         Begin MSDataListLib.DataCombo DataCombo1 
            Height          =   315
            Left            =   195
            TabIndex        =   10
            Top             =   585
            Width           =   1860
            _ExtentX        =   3281
            _ExtentY        =   556
            _Version        =   393216
            BackColor       =   16777215
            Text            =   ""
            RightToLeft     =   -1  'True
         End
         Begin MSComCtl2.DTPicker EndDate 
            Height          =   285
            Left            =   195
            TabIndex        =   55
            TabStop         =   0   'False
            Top             =   255
            Width           =   1860
            _ExtentX        =   3281
            _ExtentY        =   503
            _Version        =   393216
            CalendarBackColor=   12648447
            CalendarTitleBackColor=   10383715
            CustomFormat    =   "yyyy/M/d"
            Format          =   98762755
            CurrentDate     =   37140
         End
         Begin MSComCtl2.DTPicker DTPicker1 
            Height          =   270
            Left            =   195
            TabIndex        =   61
            TabStop         =   0   'False
            Top             =   1170
            Width           =   1890
            _ExtentX        =   3334
            _ExtentY        =   476
            _Version        =   393216
            CalendarBackColor=   12648447
            CalendarTitleBackColor=   10383715
            CustomFormat    =   "yyyy/M/d"
            Format          =   98762755
            CurrentDate     =   37140
         End
         Begin MSComCtl2.DTPicker DTPicker2 
            Height          =   270
            Left            =   3705
            TabIndex        =   62
            TabStop         =   0   'False
            Top             =   1170
            Width           =   1845
            _ExtentX        =   3254
            _ExtentY        =   476
            _Version        =   393216
            CalendarBackColor=   12648447
            CalendarTitleBackColor=   10383715
            CustomFormat    =   "yyyy/M/d"
            Format          =   98762755
            CurrentDate     =   37140
         End
         Begin Dynamic_Byte.NourHijriCal NourHijriCal1 
            Height          =   285
            Left            =   195
            TabIndex        =   63
            Top             =   1485
            Width           =   1890
            _ExtentX        =   3334
            _ExtentY        =   503
         End
         Begin Dynamic_Byte.NourHijriCal NourHijriCal2 
            Height          =   285
            Left            =   3705
            TabIndex        =   64
            Top             =   1485
            Width           =   1845
            _ExtentX        =   3254
            _ExtentY        =   503
         End
         Begin VB.Label Lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   " «—ÌŒ ‰Â«Ì… «· Œ’Ì’ „"
            Height          =   285
            Index           =   20
            Left            =   2160
            RightToLeft     =   -1  'True
            TabIndex        =   67
            Top             =   1170
            Width           =   1545
         End
         Begin VB.Label Lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   " «—ÌŒ »œ«Ì… «· Œ’Ì’ Â‹ "
            Height          =   285
            Index           =   19
            Left            =   5580
            RightToLeft     =   -1  'True
            TabIndex        =   66
            Top             =   1485
            Width           =   1800
         End
         Begin VB.Label Lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   " «—ÌŒ ‰Â«Ì… «· Œ’Ì’ Â‹‹ "
            Height          =   285
            Index           =   18
            Left            =   1980
            RightToLeft     =   -1  'True
            TabIndex        =   65
            Top             =   1485
            Width           =   1725
         End
         Begin VB.Label Lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   " «—ÌŒ »œ«Ì… «· Œ’Ì’ „"
            Height          =   285
            Index           =   16
            Left            =   5580
            RightToLeft     =   -1  'True
            TabIndex        =   60
            Top             =   1170
            Width           =   1800
         End
         Begin VB.Label Lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "»‰«¡ ⁄·Ï ⁄ﬁœ —ﬁ„"
            Height          =   285
            Index           =   15
            Left            =   5640
            RightToLeft     =   -1  'True
            TabIndex        =   59
            Top             =   855
            Width           =   1545
         End
         Begin VB.Label Lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "«·—ﬁ„ «·Ê“«—Ï ··„œ—”…"
            Height          =   270
            Index           =   9
            Left            =   5760
            RightToLeft     =   -1  'True
            TabIndex        =   58
            Top             =   585
            Width           =   1425
         End
         Begin VB.Label Label3 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   " «—ÌŒ"
            ForeColor       =   &H00000000&
            Height          =   240
            Left            =   2895
            TabIndex        =   57
            Top             =   255
            Width           =   720
         End
         Begin VB.Label Lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "—ﬁ„ «·⁄„·Ì…"
            Height          =   285
            Index           =   7
            Left            =   5760
            RightToLeft     =   -1  'True
            TabIndex        =   56
            Top             =   255
            Width           =   1425
         End
         Begin VB.Label Lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "«”„ «·„œ—”…"
            Height          =   270
            Index           =   1
            Left            =   2835
            RightToLeft     =   -1  'True
            TabIndex        =   12
            Top             =   585
            Width           =   780
         End
         Begin VB.Label Lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "⁄œœ «·ÿ·«»"
            Height          =   255
            Index           =   5
            Left            =   2670
            RightToLeft     =   -1  'True
            TabIndex        =   11
            Top             =   855
            Width           =   945
         End
      End
      Begin C1SizerLibCtl.C1Elastic EleHeader 
         Height          =   690
         Left            =   0
         TabIndex        =   1
         TabStop         =   0   'False
         Top             =   0
         Width           =   15105
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
            TabIndex        =   2
            Top             =   1320
            Visible         =   0   'False
            Width           =   855
         End
         Begin ImpulseButton.ISButton XPBtnMove 
            Height          =   345
            Index           =   0
            Left            =   1155
            TabIndex        =   3
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
            ButtonImage     =   "VehicleAllocation.frx":038A
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
            TabIndex        =   4
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
            ButtonImage     =   "VehicleAllocation.frx":0724
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
            TabIndex        =   5
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
            ButtonImage     =   "VehicleAllocation.frx":0ABE
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
            TabIndex        =   6
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
            ButtonImage     =   "VehicleAllocation.frx":0E58
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
         Height          =   3735
         Left            =   120
         TabIndex        =   14
         Top             =   3345
         Width           =   14850
         _cx             =   26194
         _cy             =   6588
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
         Cols            =   8
         FixedRows       =   1
         FixedCols       =   1
         RowHeightMin    =   0
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   -1  'True
         FormatString    =   $"VehicleAllocation.frx":11F2
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
      Begin C1SizerLibCtl.C1Elastic C1Elastic5 
         Height          =   810
         Left            =   0
         TabIndex        =   26
         TabStop         =   0   'False
         Top             =   8820
         Width           =   14925
         _cx             =   26326
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
            Height          =   645
            Index           =   0
            Left            =   16470
            TabIndex        =   27
            Top             =   135
            Width           =   2115
            _ExtentX        =   3731
            _ExtentY        =   1138
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
            ButtonImage     =   "VehicleAllocation.frx":1312
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
            Height          =   645
            Index           =   1
            Left            =   14355
            TabIndex        =   28
            Top             =   135
            Width           =   2025
            _ExtentX        =   3572
            _ExtentY        =   1138
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
            ButtonImage     =   "VehicleAllocation.frx":7B74
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
            Height          =   645
            Index           =   2
            Left            =   11745
            TabIndex        =   29
            Top             =   135
            Width           =   2625
            _ExtentX        =   4630
            _ExtentY        =   1138
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
            ButtonImage     =   "VehicleAllocation.frx":E3D6
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
            Height          =   645
            Index           =   3
            Left            =   9360
            TabIndex        =   30
            Top             =   135
            Width           =   2295
            _ExtentX        =   4048
            _ExtentY        =   1138
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
            ButtonImage     =   "VehicleAllocation.frx":14C38
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
            Height          =   645
            Index           =   4
            Left            =   7155
            TabIndex        =   31
            Top             =   135
            Width           =   2100
            _ExtentX        =   3704
            _ExtentY        =   1138
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
            ButtonImage     =   "VehicleAllocation.frx":1B49A
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
            Height          =   645
            Index           =   6
            Left            =   2700
            TabIndex        =   32
            Top             =   135
            Width           =   2295
            _ExtentX        =   4048
            _ExtentY        =   1138
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
            ButtonImage     =   "VehicleAllocation.frx":21CFC
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
            Height          =   645
            Left            =   150
            TabIndex        =   33
            Top             =   135
            Width           =   2460
            _ExtentX        =   4339
            _ExtentY        =   1138
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
            ButtonImage     =   "VehicleAllocation.frx":4B91E
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
            Height          =   645
            Index           =   7
            Left            =   4995
            TabIndex        =   34
            Top             =   135
            Width           =   2115
            _ExtentX        =   3731
            _ExtentY        =   1138
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
            ButtonImage     =   "VehicleAllocation.frx":52180
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
      Begin C1SizerLibCtl.C1Elastic C1Elastic4 
         Height          =   2295
         Left            =   150
         TabIndex        =   35
         TabStop         =   0   'False
         Top             =   840
         Width           =   7245
         _cx             =   12779
         _cy             =   4048
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
            Height          =   360
            Left            =   120
            MaxLength       =   50
            RightToLeft     =   -1  'True
            TabIndex        =   36
            Top             =   270
            Width           =   6075
         End
         Begin MSComCtl2.DTPicker dtpToDate 
            Height          =   390
            Left            =   1215
            TabIndex        =   37
            TabStop         =   0   'False
            Top             =   1140
            Width           =   1110
            _ExtentX        =   1958
            _ExtentY        =   688
            _Version        =   393216
            CalendarBackColor=   12648447
            CalendarTitleBackColor=   10383715
            CustomFormat    =   "yyyy/M/d"
            Format          =   98762755
            CurrentDate     =   37140
         End
         Begin MSComCtl2.DTPicker dtpFromDate 
            Height          =   390
            Left            =   4995
            TabIndex        =   38
            TabStop         =   0   'False
            Top             =   1140
            Width           =   1230
            _ExtentX        =   2170
            _ExtentY        =   688
            _Version        =   393216
            CalendarBackColor=   12648447
            CalendarTitleBackColor=   10383715
            CustomFormat    =   "yyyy/M/d"
            Format          =   98762755
            CurrentDate     =   37140
         End
         Begin Dynamic_Byte.NourHijriCal dtpFromDateH 
            Height          =   390
            Left            =   3660
            TabIndex        =   39
            Top             =   1140
            Width           =   1350
            _ExtentX        =   2381
            _ExtentY        =   688
         End
         Begin Dynamic_Byte.NourHijriCal dtpToDateH 
            Height          =   390
            Left            =   120
            TabIndex        =   40
            Top             =   1140
            Width           =   1110
            _ExtentX        =   1958
            _ExtentY        =   688
         End
         Begin Dynamic_Byte.NourHijriCal dtpSContractDateH 
            Height          =   390
            Left            =   3660
            TabIndex        =   41
            Top             =   1575
            Width           =   1350
            _ExtentX        =   2381
            _ExtentY        =   688
         End
         Begin MSComCtl2.DTPicker dtpSContractDate 
            Height          =   390
            Left            =   4995
            TabIndex        =   42
            TabStop         =   0   'False
            Top             =   1575
            Width           =   1230
            _ExtentX        =   2170
            _ExtentY        =   688
            _Version        =   393216
            CalendarBackColor=   12648447
            CalendarTitleBackColor=   10383715
            CustomFormat    =   "yyyy/M/d"
            Format          =   98762755
            CurrentDate     =   37140
         End
         Begin MSComCtl2.DTPicker dtpEContractDate 
            Height          =   390
            Left            =   1215
            TabIndex        =   43
            TabStop         =   0   'False
            Top             =   1575
            Width           =   1110
            _ExtentX        =   1958
            _ExtentY        =   688
            _Version        =   393216
            CalendarBackColor=   12648447
            CalendarTitleBackColor=   10383715
            CustomFormat    =   "yyyy/M/d"
            Format          =   98762755
            CurrentDate     =   37140
         End
         Begin Dynamic_Byte.NourHijriCal dtpEContractDateH 
            Height          =   390
            Left            =   120
            TabIndex        =   44
            Top             =   1575
            Width           =   1110
            _ExtentX        =   1958
            _ExtentY        =   688
         End
         Begin MSDataListLib.DataCombo dcVendor 
            Height          =   315
            Left            =   120
            TabIndex        =   45
            Top             =   705
            Width           =   2190
            _ExtentX        =   3863
            _ExtentY        =   556
            _Version        =   393216
            BackColor       =   16777215
            Text            =   ""
            RightToLeft     =   -1  'True
         End
         Begin MSDataListLib.DataCombo dcCity 
            Height          =   315
            Left            =   3660
            TabIndex        =   46
            Top             =   705
            Width           =   2535
            _ExtentX        =   4471
            _ExtentY        =   556
            _Version        =   393216
            BackColor       =   16777215
            Text            =   ""
            RightToLeft     =   -1  'True
         End
         Begin VB.Label Lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "„”„Ï «· ⁄«ﬁœ"
            Height          =   360
            Index           =   14
            Left            =   5970
            RightToLeft     =   -1  'True
            TabIndex        =   53
            Top             =   270
            Width           =   1230
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "«·Ï  «—ÌŒ"
            ForeColor       =   &H00000000&
            Height          =   450
            Index           =   0
            Left            =   2715
            TabIndex        =   52
            Top             =   1140
            Width           =   810
         End
         Begin VB.Label Label4 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "„‰  «—ÌŒ"
            ForeColor       =   &H00000000&
            Height          =   285
            Left            =   6360
            TabIndex        =   51
            Top             =   1140
            Width           =   825
         End
         Begin VB.Label Lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "  «· ⁄«ﬁœ „Ì·«œÏ"
            Height          =   390
            Index           =   13
            Left            =   6210
            RightToLeft     =   -1  'True
            TabIndex        =   50
            Top             =   1575
            Width           =   990
         End
         Begin VB.Label Lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "  «‰ Â«¡ «· ⁄«ﬁœ „Ì·«œÏ"
            Height          =   660
            Index           =   8
            Left            =   2445
            RightToLeft     =   -1  'True
            TabIndex        =   49
            Top             =   1575
            Width           =   1110
         End
         Begin VB.Label Lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "«”„ «·„‰ÿﬁ…"
            Height          =   390
            Index           =   3
            Left            =   6210
            RightToLeft     =   -1  'True
            TabIndex        =   48
            Top             =   705
            Width           =   990
         End
         Begin VB.Label Lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "«·«œ«—… «· ⁄·Ì„Ì…"
            Height          =   390
            Index           =   0
            Left            =   2445
            RightToLeft     =   -1  'True
            TabIndex        =   47
            Top             =   705
            Width           =   1110
         End
      End
      Begin VB.Label XPTxtCurrent 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Height          =   345
         Left            =   11730
         RightToLeft     =   -1  'True
         TabIndex        =   18
         Top             =   8160
         Width           =   990
      End
      Begin VB.Label XPTxtCount 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Height          =   345
         Left            =   8160
         RightToLeft     =   -1  'True
         TabIndex        =   17
         Top             =   8160
         Width           =   855
      End
      Begin VB.Label Lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   " «·”Ã· «·Õ«·Ì:"
         Height          =   345
         Index           =   2
         Left            =   12645
         RightToLeft     =   -1  'True
         TabIndex        =   16
         Top             =   8160
         Width           =   1365
      End
      Begin VB.Label Lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   " ⁄œœ «·”Ã·« :"
         Height          =   345
         Index           =   4
         Left            =   9030
         RightToLeft     =   -1  'True
         TabIndex        =   15
         Top             =   8160
         Width           =   1410
      End
   End
End
Attribute VB_Name = "FrmVehicleAllocation"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
 
