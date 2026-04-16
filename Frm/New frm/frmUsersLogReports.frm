VERSION 5.00
Object = "{0AFE7BE0-11B7-4A3E-978D-D4501E9A57FE}#1.0#0"; "c1sizer.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{FE5DCFAD-BC1D-11D2-94CF-004005455FAA}#1.4#0"; "ImpulseButton.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form FrmUsersLogReports 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   " ‹‹ﬁ‹‹«—Ì‹‹‹— «·„” Œœ„Ì‰"
   ClientHeight    =   7770
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   9900
   HelpContextID   =   470
   Icon            =   "frmUsersLogReports.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   RightToLeft     =   -1  'True
   ScaleHeight     =   7770
   ScaleWidth      =   9900
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   8
      Charset         =   178
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Begin C1SizerLibCtl.C1Tab TabMain 
      Height          =   7725
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   9855
      _cx             =   17383
      _cy             =   13626
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
      Caption         =   " ﬁ—Ì— «·„” Œœ„Ì‰ | ﬁ«—Ì— ’·«ÕÌ«  «·„” Œœ„Ì‰"
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
      DogEars         =   -1  'True
      MultiRow        =   0   'False
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
      Begin C1SizerLibCtl.C1Elastic Ele30 
         Height          =   7350
         Index           =   2
         Left            =   45
         TabIndex        =   1
         TabStop         =   0   'False
         Top             =   45
         Width           =   9765
         _cx             =   17224
         _cy             =   12965
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
         Frame           =   0
         FrameStyle      =   5
         FrameWidth      =   1
         FrameColor      =   -2147483628
         FrameShadow     =   -2147483632
         FloodStyle      =   1
         _GridInfo       =   ""
         AccessibleName  =   ""
         AccessibleDescription=   ""
         AccessibleValue =   ""
         AccessibleRole  =   9
         Begin VB.TextBox Text4 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   3600
            RightToLeft     =   -1  'True
            TabIndex        =   49
            Top             =   2040
            Width           =   1575
         End
         Begin VB.TextBox TxtDescription 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00C0FFFF&
            Height          =   885
            Left            =   840
            MultiLine       =   -1  'True
            RightToLeft     =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   44
            Top             =   3120
            Width           =   4335
         End
         Begin VB.TextBox txtNoteSeial 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   840
            RightToLeft     =   -1  'True
            TabIndex        =   10
            Top             =   2760
            Width           =   4335
         End
         Begin VB.TextBox txtNoteSeial1 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   840
            RightToLeft     =   -1  'True
            TabIndex        =   9
            Top             =   2400
            Width           =   4335
         End
         Begin VB.TextBox Txt_order_no 
            Alignment       =   1  'Right Justify
            Height          =   330
            Left            =   3000
            RightToLeft     =   -1  'True
            TabIndex        =   8
            Top             =   7920
            Visible         =   0   'False
            Width           =   1665
         End
         Begin VB.TextBox Text1 
            Alignment       =   1  'Right Justify
            Height          =   330
            Left            =   3000
            RightToLeft     =   -1  'True
            TabIndex        =   7
            Top             =   8280
            Visible         =   0   'False
            Width           =   1665
         End
         Begin VB.CheckBox ChkUser 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "„” Œœ„ „Õœœ"
            Height          =   255
            Left            =   6360
            RightToLeft     =   -1  'True
            TabIndex        =   6
            Top             =   960
            Width           =   2295
         End
         Begin VB.CheckBox ChkNotesType 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "Õ—ﬂ«  „⁄Ì‰…"
            Height          =   255
            Left            =   6360
            RightToLeft     =   -1  'True
            TabIndex        =   5
            Top             =   1320
            Width           =   2295
         End
         Begin VB.CheckBox ChkType 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "ÿ»Ì⁄Â «·Õ—ﬂ…"
            Height          =   255
            Left            =   6360
            RightToLeft     =   -1  'True
            TabIndex        =   4
            Top             =   1680
            Width           =   2295
         End
         Begin VB.ComboBox CboTransactionType 
            Enabled         =   0   'False
            Height          =   315
            Left            =   840
            RightToLeft     =   -1  'True
            TabIndex        =   3
            Top             =   1680
            Width           =   4335
         End
         Begin VB.CheckBox ChkScreen 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "‘«‘… „⁄Ì‰…"
            Height          =   255
            Left            =   6360
            RightToLeft     =   -1  'True
            TabIndex        =   2
            Top             =   2040
            Width           =   2295
         End
         Begin C1SizerLibCtl.C1Elastic Ele 
            Height          =   1185
            Index           =   1
            Left            =   3300
            TabIndex        =   11
            TabStop         =   0   'False
            Top             =   4260
            Width           =   4785
            _cx             =   8440
            _cy             =   2090
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
            Caption         =   " ÕœÌœ «·› —… «·“„‰Ì…"
            Align           =   0
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
            Style           =   1
            TagSplit        =   2
            PicturePos      =   4
            CaptionStyle    =   0
            ResizeFonts     =   0   'False
            GridRows        =   0
            GridCols        =   0
            Frame           =   0
            FrameStyle      =   5
            FrameWidth      =   1
            FrameColor      =   -2147483628
            FrameShadow     =   -2147483632
            FloodStyle      =   1
            _GridInfo       =   ""
            AccessibleName  =   ""
            AccessibleDescription=   ""
            AccessibleValue =   ""
            AccessibleRole  =   9
            Begin MSComCtl2.DTPicker DTPickerAccFrom 
               BeginProperty DataFormat 
                  Type            =   1
                  Format          =   "dd/MM/yyyy"
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   11265
                  SubFormatType   =   3
               EndProperty
               Height          =   345
               Left            =   1860
               TabIndex        =   12
               ToolTipText     =   "„‰  «—ÌŒ ﬁœÌ„"
               Top             =   240
               Width           =   1500
               _ExtentX        =   2646
               _ExtentY        =   609
               _Version        =   393216
               CalendarBackColor=   -2147483624
               CalendarTitleBackColor=   10383715
               CheckBox        =   -1  'True
               CustomFormat    =   "yyyy/M/d"
               Format          =   98959363
               CurrentDate     =   37357
            End
            Begin MSComCtl2.DTPicker DTPickerAccTo 
               BeginProperty DataFormat 
                  Type            =   1
                  Format          =   "dd/MM/yyyy"
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   11265
                  SubFormatType   =   3
               EndProperty
               Height          =   345
               Left            =   1860
               TabIndex        =   13
               ToolTipText     =   " ≈·Ï  «—ÌŒ √ÕœÀ"
               Top             =   600
               Width           =   1500
               _ExtentX        =   2646
               _ExtentY        =   609
               _Version        =   393216
               CalendarBackColor=   -2147483624
               CalendarTitleBackColor=   10383715
               CheckBox        =   -1  'True
               CustomFormat    =   "yyyy/M/d"
               Format          =   98959363
               CurrentDate     =   37357
            End
            Begin MSComCtl2.DTPicker XPDtbTransTimeFrom 
               Height          =   375
               Left            =   60
               TabIndex        =   47
               Top             =   210
               Width           =   1725
               _ExtentX        =   3043
               _ExtentY        =   661
               _Version        =   393216
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Times New Roman"
                  Size            =   9.75
                  Charset         =   178
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               CheckBox        =   -1  'True
               Format          =   98959362
               CurrentDate     =   38784
            End
            Begin MSComCtl2.DTPicker XPDtbTransTimeTo 
               Height          =   405
               Left            =   30
               TabIndex        =   48
               Top             =   600
               Width           =   1755
               _ExtentX        =   3096
               _ExtentY        =   714
               _Version        =   393216
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Times New Roman"
                  Size            =   9.75
                  Charset         =   178
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               CheckBox        =   -1  'True
               Format          =   98959362
               CurrentDate     =   38784
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "≈·Ï"
               Height          =   285
               Index           =   2
               Left            =   3360
               RightToLeft     =   -1  'True
               TabIndex        =   15
               Top             =   600
               Width           =   555
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "„‰"
               Height          =   285
               Index           =   4
               Left            =   3360
               RightToLeft     =   -1  'True
               TabIndex        =   14
               Top             =   285
               Width           =   555
            End
         End
         Begin ImpulseButton.ISButton CmdAccount 
            Height          =   405
            Left            =   600
            TabIndex        =   16
            Top             =   5040
            Width           =   1305
            _ExtentX        =   2302
            _ExtentY        =   714
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
            ButtonImage     =   "frmUsersLogReports.frx":038A
            ColorButton     =   14871017
            ColorHoverText  =   16777215
            DrawFocusRectangle=   0   'False
            ColorToggledHoverText=   16777215
         End
         Begin MSDataListLib.DataCombo DCNotesTypes 
            Height          =   315
            Left            =   840
            TabIndex        =   17
            Top             =   1320
            Width           =   4305
            _ExtentX        =   7594
            _ExtentY        =   556
            _Version        =   393216
            Enabled         =   0   'False
            Text            =   ""
            RightToLeft     =   -1  'True
         End
         Begin MSDataListLib.DataCombo DCScreen 
            Height          =   315
            Left            =   840
            TabIndex        =   18
            Top             =   2040
            Width           =   2745
            _ExtentX        =   4842
            _ExtentY        =   556
            _Version        =   393216
            Enabled         =   0   'False
            Text            =   ""
            RightToLeft     =   -1  'True
         End
         Begin MSDataListLib.DataCombo DCboUserName 
            Height          =   315
            Left            =   840
            TabIndex        =   19
            Top             =   960
            Width           =   4305
            _ExtentX        =   7594
            _ExtentY        =   556
            _Version        =   393216
            Enabled         =   0   'False
            Text            =   ""
            RightToLeft     =   -1  'True
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "≈·Ï"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   13.5
               Charset         =   178
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   435
            Index           =   10
            Left            =   6060
            RightToLeft     =   -1  'True
            TabIndex        =   46
            Top             =   7680
            Width           =   480
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "«·‘—Õ"
            Height          =   315
            Index           =   9
            Left            =   5160
            RightToLeft     =   -1  'True
            TabIndex        =   45
            Top             =   3240
            Width           =   975
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "—ﬁ„ «·ﬁÌœ"
            Height          =   315
            Index           =   8
            Left            =   5160
            RightToLeft     =   -1  'True
            TabIndex        =   28
            Top             =   2760
            Width           =   975
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "—ﬁ„ «·Õ—ﬂ…"
            Height          =   315
            Index           =   7
            Left            =   5160
            RightToLeft     =   -1  'True
            TabIndex        =   27
            Top             =   2400
            Width           =   975
         End
         Begin VB.Label LblAccountName 
            Alignment       =   2  'Center
            BackColor       =   &H00C0C8C0&
            Caption         =   "”Ã· ‰‘«ÿ «·„” Œœ„Ì‰"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   13.5
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   645
            Left            =   0
            RightToLeft     =   -1  'True
            TabIndex        =   26
            Top             =   75
            Width           =   9030
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "—ﬁ„ «·ÿ·»Ì…"
            ForeColor       =   &H00000000&
            Height          =   285
            Index           =   17
            Left            =   4680
            RightToLeft     =   -1  'True
            TabIndex        =   25
            Top             =   8760
            Visible         =   0   'False
            Width           =   1335
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "—ﬁ„ √„— «·«‰ «Ã"
            ForeColor       =   &H00000000&
            Height          =   285
            Index           =   0
            Left            =   5940
            RightToLeft     =   -1  'True
            TabIndex        =   24
            Top             =   8520
            Visible         =   0   'False
            Width           =   1335
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "«”„ «·„” Œœ„"
            Height          =   315
            Index           =   1
            Left            =   5160
            RightToLeft     =   -1  'True
            TabIndex        =   23
            Top             =   960
            Width           =   975
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "Õœœ «·Õ—ﬂ…"
            Height          =   315
            Index           =   3
            Left            =   5160
            RightToLeft     =   -1  'True
            TabIndex        =   22
            Top             =   1320
            Width           =   975
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "ÿ»Ì⁄Â «·Õ—ﬂ…"
            Height          =   315
            Index           =   5
            Left            =   5160
            RightToLeft     =   -1  'True
            TabIndex        =   21
            Top             =   1680
            Width           =   975
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "«”„ «·‘«‘…"
            Height          =   315
            Index           =   6
            Left            =   5160
            RightToLeft     =   -1  'True
            TabIndex        =   20
            Top             =   2040
            Width           =   975
         End
      End
      Begin C1SizerLibCtl.C1Elastic Ele20 
         Height          =   7350
         Index           =   0
         Left            =   10500
         TabIndex        =   29
         TabStop         =   0   'False
         Top             =   45
         Width           =   9765
         _cx             =   17224
         _cy             =   12965
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
         Frame           =   0
         FrameStyle      =   5
         FrameWidth      =   1
         FrameColor      =   -2147483628
         FrameShadow     =   -2147483632
         FloodStyle      =   1
         _GridInfo       =   ""
         AccessibleName  =   ""
         AccessibleDescription=   ""
         AccessibleValue =   ""
         AccessibleRole  =   9
         Begin VB.CommandButton Clear 
            Caption         =   "„”Õ"
            Height          =   495
            Left            =   4560
            RightToLeft     =   -1  'True
            TabIndex        =   43
            Top             =   4200
            Width           =   1815
         End
         Begin VB.CommandButton PrintRep 
            Caption         =   "⁄—÷ «· ﬁ—Ì— "
            Height          =   495
            Left            =   6600
            RightToLeft     =   -1  'True
            TabIndex        =   42
            Top             =   4200
            Width           =   1815
         End
         Begin VB.CheckBox ShowDeniedUsers 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "«ŸÂ«— «·ﬂ·"
            Height          =   300
            Left            =   2520
            RightToLeft     =   -1  'True
            TabIndex        =   41
            Top             =   360
            Width           =   3420
         End
         Begin VB.CheckBox ShowAdvPerm 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "⁄—÷ «·’·«ÕÌ«  «·„Œ’’…"
            Height          =   300
            Left            =   6120
            RightToLeft     =   -1  'True
            TabIndex        =   40
            Top             =   360
            Width           =   2220
         End
         Begin VB.ListBox SelectedEmployeesList 
            Height          =   2790
            ItemData        =   "frmUsersLogReports.frx":0724
            Left            =   240
            List            =   "frmUsersLogReports.frx":072B
            RightToLeft     =   -1  'True
            TabIndex        =   39
            Top             =   840
            Width           =   3375
         End
         Begin VB.ListBox EmployeesList 
            Height          =   2790
            ItemData        =   "frmUsersLogReports.frx":0746
            Left            =   5040
            List            =   "frmUsersLogReports.frx":074D
            RightToLeft     =   -1  'True
            TabIndex        =   34
            Top             =   840
            Width           =   3375
         End
         Begin VB.TextBox Text2 
            Alignment       =   1  'Right Justify
            Height          =   330
            Left            =   3000
            RightToLeft     =   -1  'True
            TabIndex        =   31
            Top             =   5760
            Visible         =   0   'False
            Width           =   1665
         End
         Begin VB.TextBox Text3 
            Alignment       =   1  'Right Justify
            Height          =   330
            Left            =   3000
            RightToLeft     =   -1  'True
            TabIndex        =   30
            Top             =   5880
            Visible         =   0   'False
            Width           =   1665
         End
         Begin VB.Label Label5 
            Alignment       =   2  'Center
            BackColor       =   &H00E2E9E9&
            Caption         =   "<"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   13.5
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   705
            Left            =   4080
            RightToLeft     =   -1  'True
            TabIndex        =   38
            Top             =   2355
            Width           =   495
         End
         Begin VB.Label Label6 
            Alignment       =   2  'Center
            BackColor       =   &H00E2E9E9&
            Caption         =   "<<"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   13.5
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   360
            Left            =   4080
            RightToLeft     =   -1  'True
            TabIndex        =   37
            Top             =   3045
            Width           =   495
         End
         Begin VB.Label Label7 
            Alignment       =   2  'Center
            BackColor       =   &H00E2E9E9&
            Caption         =   ">>"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   13.5
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   720
            Left            =   4080
            RightToLeft     =   -1  'True
            TabIndex        =   36
            Top             =   1650
            Width           =   495
         End
         Begin VB.Label Label8 
            Alignment       =   2  'Center
            BackColor       =   &H00E2E9E9&
            Caption         =   ">"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   13.5
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   705
            Left            =   4080
            RightToLeft     =   -1  'True
            TabIndex        =   35
            Top             =   960
            Width           =   495
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "—ﬁ„ √„— «·«‰ «Ã"
            ForeColor       =   &H00000000&
            Height          =   285
            Index           =   15
            Left            =   4680
            RightToLeft     =   -1  'True
            TabIndex        =   33
            Top             =   5880
            Visible         =   0   'False
            Width           =   1335
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "—ﬁ„ «·ÿ·»Ì…"
            ForeColor       =   &H00000000&
            Height          =   285
            Index           =   16
            Left            =   4680
            RightToLeft     =   -1  'True
            TabIndex        =   32
            Top             =   5880
            Visible         =   0   'False
            Width           =   1335
         End
      End
   End
End
Attribute VB_Name = "FrmUsersLogReports"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim StrAccountCode As String
Dim StrAccountName As String
  
Private Sub ChangeLang()
    'Label1.Caption = "Des"

    Me.Caption = "Users LogFile Reports"
    LblAccountName.Caption = Me.Caption
    lbl(7).Caption = "Tran."
     lbl(8).Caption = "GL."
     Ele(1).Caption = "Interval"
     lbl(4).Caption = "From"
     lbl(2).Caption = "To"
     CmdAccount.Caption = "Print"
     
     
     
 
    With CboTransactionType
        .Clear
        .AddItem "LogIn/Out "
        .AddItem " Open \Close Window"
        .AddItem "New"
        .AddItem "Edit"
        .AddItem "Delete"
        .AddItem "View Report"
      
    End With
 
    ChkUser.Caption = "Specific User"
    lbl(1).Caption = "User"
 
    ChkNotesType.Caption = "Specific Transaction"
    lbl(3).Caption = "Transaction"
 
    ChkType.Caption = "Specific Type"
    lbl(5).Caption = "Type"
 
    ChkScreen.Caption = "Specific Window"
    lbl(6).Caption = "Window"
    
    '########################################## Khaled's Code ###################################
    ShowAdvPerm.Caption = "Show Specific Permissions"
    ShowDeniedUsers.Caption = "Show All"
    PrintRep.Caption = "Print Report"
    Clear.Caption = "Clear List"
    TabMain.Caption = " User Reports | User permissions Reports "
    '############################################################################################
 
End Sub

Private Sub ChkNotesType_Click()

    If ChkNotesType.value = vbChecked Then
        DCNotesTypes.Enabled = True
        DCNotesTypes.BoundText = ""
    Else
        DCNotesTypes.Enabled = False
        DCNotesTypes.BoundText = ""
    End If

End Sub

Private Sub ChkScreen_Click()

    If ChkScreen.value = vbChecked Then
        DCScreen.Enabled = True
        DCScreen.BoundText = ""
    Else
        DCScreen.Enabled = False
        DCScreen.BoundText = ""
    End If

End Sub

Private Sub ChkType_Click()

    If ChkType.value = vbChecked Then
        CboTransactionType.Enabled = True
        CboTransactionType.Text = ""
    Else
        CboTransactionType.Enabled = False
        CboTransactionType.Text = ""
    End If

End Sub

Private Sub ChkUser_Click()

    If ChkUser.value = vbChecked Then
        DCboUserName.Enabled = True
        DCboUserName.BoundText = ""
    Else
        DCboUserName.Enabled = False
        DCboUserName.BoundText = ""
    End If
     
End Sub


Private Sub CmdAccount_Click()
    Dim i As Integer
    Dim cAccountReport As ClsAccReports
    Dim whrstr As String
    Dim Currenttype As String
    whrstr = "1=1"
 
    If CboTransactionType.ListIndex = 0 Then
        Currenttype = "L"
    ElseIf CboTransactionType.ListIndex = 1 Then
        Currenttype = "O"
    ElseIf CboTransactionType.ListIndex = 2 Then
        Currenttype = "N"
    ElseIf CboTransactionType.ListIndex = 3 Then
        Currenttype = "E"
    ElseIf CboTransactionType.ListIndex = 4 Then
        Currenttype = "D"
    ElseIf CboTransactionType.ListIndex = 5 Then
        Currenttype = "V"
    End If
       
    If ChkUser.value = vbChecked And DCboUserName.BoundText = "" Then
        If SystemOptions.UserInterface = ArabicInterface Then
            MsgBox "«Œ — «”„ „” Œœ„", vbCritical
        Else
            MsgBox "Specify User Name", vbCritical
        End If

        DCboUserName.SetFocus
        SendKeys ("{F4}")
        Exit Sub
    ElseIf ChkUser.value = vbChecked And DCboUserName.BoundText <> "" Then
        whrstr = whrstr & " and  LogFile.UserID=" & val(DCboUserName.BoundText)
    End If
   
    If ChkNotesType.value = vbChecked And DCNotesTypes.BoundText = "" Then
        If SystemOptions.UserInterface = ArabicInterface Then
            MsgBox "«Œ —  «·Õ—ﬂ…  ", vbCritical
        Else
            MsgBox "Specify Transaction", vbCritical
        End If

        DCNotesTypes.SetFocus
        SendKeys ("{F4}")
        Exit Sub
    ElseIf ChkNotesType.value = vbChecked And DCNotesTypes.BoundText <> "" Then
        whrstr = whrstr & " and  LogFile.NotesType=" & val(DCNotesTypes.BoundText)
    End If
    
    If ChkType.value = vbChecked And CboTransactionType.ListIndex = -1 Then
        If SystemOptions.UserInterface = ArabicInterface Then
            MsgBox "«Œ — ÿ»Ì⁄Â  «·Õ—ﬂ…  ", vbCritical
        Else
            MsgBox "Specify Transaction type", vbCritical
        End If

        CboTransactionType.SetFocus
        SendKeys ("{F4}")
        Exit Sub
    ElseIf ChkType.value = vbChecked And CboTransactionType.ListIndex <> -1 Then
   
        whrstr = whrstr & " and  LogFile.TransactionType='" & Currenttype & "'"
   
    End If
   
    If ChkScreen.value = vbChecked And DCScreen.BoundText = "" Then
        If SystemOptions.UserInterface = ArabicInterface Then
            MsgBox "«Œ —     «·‘«‘…  ", vbCritical
        Else
            MsgBox "Specify Window  ", vbCritical
        End If

        DCScreen.SetFocus
        SendKeys ("{F4}")
        Exit Sub
    ElseIf ChkScreen.value = vbChecked And DCScreen.BoundText <> "" Then
        whrstr = whrstr & " and  LogFile.Remarks='" & DCScreen.BoundText & "'"
    End If
   
    If (TxtDescription.Text) <> "" Then
 
        whrstr = whrstr & " and  ( LogFile.Description like '%" & (TxtDescription.Text) & "%'"
        whrstr = whrstr & "or LogFile.descriptione  like '%" & (TxtDescription.Text) & "%' )"
        
   
    End If
    
    If Not IsNull(Me.DTPickerAccFrom.value) Then
        whrstr = whrstr + " and   LogDate >=" & Trim(ReformatDate(CStr(DTPickerAccFrom.value), True))
        
            '" & SQLDate(Me.BegineDate, True) & ""
            'CONVERT(DATETIME, " & Me.BegineDate & " , 102)
    '        WHERE LogDate > {ts '2020-03-05 00:00:00.001'}
    '  AND LogDate < {ts '2020-03-05 12:08:00.000'}

    End If

    If Not IsNull(Me.DTPickerAccTo.value) Then
        'MySQL = MySQL + " and LogDate <=" & SQLDate(EndDate, True)
        whrstr = whrstr + " and   LogDate <=    " & Trim(ReformatDate(CStr(DTPickerAccTo.value), False))
    End If
    
          If Not IsNull(Me.XPDtbTransTimeFrom.value) Then
       
                   whrstr = whrstr & " AND CAST(LogTime as time) >=N'" & FormatDateTime(Me.XPDtbTransTimeFrom.value, vbShortTime) & "'"
      End If
       If Not IsNull(Me.XPDtbTransTimeFrom.value) Then
            
                   whrstr = whrstr & " AND CAST(LogTime as time)<=N'" & FormatDateTime(Me.XPDtbTransTimeTo.value, vbShortTime) & "'"
      End If
      
      
'

       If val(txtNoteSeial.Text) <> 0 Then
 
        whrstr = whrstr & " and  LogFile.NotesSerial=" & (txtNoteSeial.Text)
   
    End If
    
    If val(txtNoteSeial1.Text) <> 0 Then
 
        whrstr = whrstr & " and  LogFile.NotesSerial1='" & (txtNoteSeial1.Text) & "'"
   
    End If
   
    Set cAccountReport = New ClsAccReports
    Screen.MousePointer = 11
                
 '   cAccountReport.BegineDate = Me.DTPickerAccFrom.value
 '   cAccountReport.EndDate = DateAdd("D", 0, DTPickerAccTo)
    cAccountReport.ShowLogFile whrstr
    Set cAccountReport = Nothing
 
End Sub
Function ReformatDate(StrDate As String, IsFrom As Boolean) As String
    Dim strDateparts() As String
    Dim dt As Date
    
    strDateparts = Split(StrDate, "/")
    dt = DateSerial(strDateparts(2), strDateparts(1), strDateparts(0))
    ReformatDate = Format(dt, "yyyy-mm-dd")
    If IsFrom Then
        ReformatDate = "{ts '" & ReformatDate & " 00:00:00.000'}"
    Else
        ReformatDate = "{ts '" & ReformatDate & " 23:59:59.001'}"
    End If
     
End Function
Private Sub Form_Load()
    Resize_Form Me, NoChangeInSize
    StrAccountCode = ""
 
    Dim StrSQL As String
 
    Dim Msg As String
 
    With CboTransactionType
        .Clear
        .AddItem " ”ÃÌ· œŒÊ· ÊŒ—ÊÃ "
        .AddItem "› Õ / €·ﬁ ‘«‘…"
        .AddItem "ÃœÌœ"
        .AddItem " ⁄œÌ·"
        .AddItem "Õ–›"
        .AddItem "⁄—÷  ﬁ—Ì—"
     
    End With

    StrSQL = "SELECT UserID,UserName From TblUsers    order by UserName"
    fill_combo DCboUserName, StrSQL

    If SystemOptions.UserInterface = ArabicInterface Then
       ' StrSQL = "SELECT NotesType,NotesTypeName From TblNotesTypes order by NotesTypeName "
        
   StrSQL = "     SELECT NotesType,RTRIM( LTRIM( NotesTypeName)) as NotesTypeName From TblNotesTypes"
StrSQL = StrSQL & "  order by NotesType, NotesTypeName"

    Else
       StrSQL = "     SELECT NotesType,RTRIM( LTRIM( NotesTypeNamee)) as NotesTypeNamee From TblNotesTypes"
StrSQL = StrSQL & "  order by NotesType, NotesTypeNamee"


'        StrSQL = "SELECT NotesType,NotesTypeNamee From TblNotesTypes  order by NotesTypeNamee"
    End If

    fill_combo DCNotesTypes, StrSQL

    If SystemOptions.UserInterface = ArabicInterface Then
        StrSQL = "SELECT LTRIM(RTRIM(ScreenName)) as ScreenName ,LTRIM(RTRIM(screencaption)) as screencaption From Screens order by LTRIM(RTRIM(screencaption)) "
    Else:
        StrSQL = "SELECT LTRIM(RTRIM(ScreenName)) as ScreenName ,LTRIM(RTRIM(ScreenTitleEng)) From Screens  order by  LTRIM(RTRIM(ScreenTitleEng)) "
    End If

    fill_combo DCScreen, StrSQL

    Set cSearchDcbo = New clsDCboSearch
    Set cSearchDcbo.Client = Me.DCboUserName

    SetDtpickerDate Me.DTPickerAccFrom
    SetDtpickerDate Me.DTPickerAccTo

    Dim FirstPeriodDateInthisYear  As Date
    getFirstPeriodDateInthisYear FirstPeriodDateInthisYear
    Me.DTPickerAccFrom = FirstPeriodDateInthisYear
    Me.DTPickerAccTo = Date
 
    If SystemOptions.UserInterface = EnglishInterface Then
        SetInterface Me
        ChangeLang
    End If
    
   '########################################################## Khaled's Code ########################################################
   FillEmployeesList
   '#################################################################################################################################
    
End Sub
 '########################################################## Khaled's Code ########################################################
Function FillEmployeesList()
    Dim sql As String
    Dim Rs2 As ADODB.Recordset
    Dim i As Integer
    Set Rs2 = New ADODB.Recordset
    sql = " SELECT * from  TblUsers where UserID <> 1"
    Rs2.Open sql, Cn, adOpenStatic, adLockReadOnly, adCmdText
    Me.EmployeesList.Clear
    Me.SelectedEmployeesList.Clear
    If Rs2.RecordCount > 0 Then
        For i = 1 To Rs2.RecordCount
            If SystemOptions.UserInterface = ArabicInterface Then
                EmployeesList.AddItem IIf(IsNull(Rs2("UserName").value), "", Rs2("UserName").value)
            Else
                EmployeesList.AddItem IIf(IsNull(Rs2("UserName").value), "", Rs2("UserName").value)
            End If
            EmployeesList.ItemData(EmployeesList.NewIndex) = IIf(IsNull(Rs2("UserID").value), 0, Rs2("UserID").value)
            Rs2.MoveNext
        Next i

    End If
    Rs2.Close
End Function
'------------------------ Add All Employees --------------------------
Private Sub Label7_Click()
    Dim i As Integer
    Me.SelectedEmployeesList.Clear
    For i = 0 To Me.EmployeesList.ListCount - 1
        Me.SelectedEmployeesList.AddItem EmployeesList.List(i)
        SelectedEmployeesList.ItemData(i) = EmployeesList.ItemData(i)
    Next i
End Sub
'----------------- Add one Employee at a time  ----------------------
Private Sub Label8_Click()
Dim Rs1  As ADODB.Recordset
Dim i, k As Integer
 If Me.EmployeesList.ListIndex > -1 Then
    Me.SelectedEmployeesList.AddItem EmployeesList.List(EmployeesList.ListIndex)
    SelectedEmployeesList.ItemData(SelectedEmployeesList.NewIndex) = EmployeesList.ItemData(EmployeesList.ListIndex)
End If
End Sub
'--------------- Remove one Employee at a time -----------------
Private Sub Label5_Click()
If SelectedEmployeesList.ListIndex > -1 Then
SelectedEmployeesList.RemoveItem (SelectedEmployeesList.ListIndex)
End If
End Sub
Private Sub Label6_Click()
SelectedEmployeesList.Clear
End Sub
Private Sub Clear_Click()
SelectedEmployeesList.Clear
End Sub
Private Sub PrintRep_Click()
Dim SqlQu As String
Dim UserIDs As String
Dim Msg As String

UserIDs = "0"

Dim i As Integer
 If Me.SelectedEmployeesList.ListCount > 0 Then
     For i = 0 To Me.SelectedEmployeesList.ListCount - 1
       UserIDs = UserIDs & "," & SelectedEmployeesList.ItemData(i)
     Next
End If
If UserIDs = "0" Then
  If SystemOptions.UserInterface = ArabicInterface Then
    Msg = "«·—Ã«¡  ÕœÌœ „” Œœ„ Ê«Õœ ⁄·Ï «·√ﬁ·"
    Else
    Msg = "Please select at least one user "
  End If
  MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
  Exit Sub
End If

SqlQu = "SELECT Distinct dbo.TblUsers.UserID, dbo.Screens.ScreenName, dbo.Screens.ScreenCaption, dbo.TblUsers.Empid, dbo.TblUsers.UserName, dbo.ScreenJuncUser.CanAdd,"
SqlQu = SqlQu & "dbo.ScreenJuncUser.CanEdit ,dbo.ScreenJuncUser.CanDelete, dbo.ScreenJuncUser.CanPrint, dbo.ScreenJuncUser.CanSearch, dbo.ScreenJuncUser.FullAccess,"
SqlQu = SqlQu & "dbo.ScreenJuncUser.CanShow, dbo.TblUsers.InvPrices,dbo.TblUsers.ShowInvProfit, dbo.TblUsers.InvPrices1, dbo.TblUsers.InvPrices2, dbo.TblUsers.FixedCustomer,"
SqlQu = SqlQu & "dbo.TblUsers.ShowBillCommisions, dbo.TblUsers.HideCost, dbo.TblUsers.hideColumn,dbo.TblUsers.ExceedShipment, dbo.TblUsers.AllowSett1, dbo.TblUsers.AllowSett,"
SqlQu = SqlQu & "dbo.TblUsers.Allowpayroll, dbo.TblUsers.AllowBigAccount, dbo.TblUsers.AllowRequestgl, dbo.TblUsers.Allowrank,dbo.TblUsers.AllowOrbonDate,"
SqlQu = SqlQu & "dbo.TblUsers.AllowCreateHajomraVoucher, dbo.TblUsers.AllowCompChanPrice, dbo.TblUsers.AllowChangeSalesAtTransfer, dbo.TblUsers.AllowSalesSaveWithoutCostPrice,"
SqlQu = SqlQu & "dbo.TblUsers.AllowChanProjectBillPrice , dbo.Screens.ScreenType , dbo.ScreenJuncUser.Attachments "
SqlQu = SqlQu & "FROM dbo.TblUsers INNER JOIN dbo.ScreenJuncUser ON dbo.TblUsers.UserID = dbo.ScreenJuncUser.User_ID INNER JOIN dbo.Screens ON dbo.ScreenJuncUser.ScreenName = dbo.Screens.ScreenName "
SqlQu = SqlQu & "Where dbo.TblUsers.UserID in (" & UserIDs & ")"   'User IDs Goes here

print_report_UserPerm SqlQu
End Sub
Function print_report_UserPerm(Optional NoteSerial As String)
     
    Set rs = New ADODB.Recordset
    rs.Open NoteSerial, Cn, adOpenStatic, adLockReadOnly, adCmdText
     
    Dim MySQL As String
    Dim RsData As New ADODB.Recordset
    Dim xApp As New CRAXDRT.Application
    Dim xReport As CRAXDRT.Report
    Dim CViewer As ClsReportViewer
    Dim StrReportTitle As String
    Dim StrFileName As String
    Dim Msg As String

     If SystemOptions.UserInterface = ArabicInterface Then
            StrFileName = App.path & "\REPORTS\REPORTS NEW\" & "PermissionsReport.rpt"
            Else
             StrFileName = App.path & "\REPORTS\REPORTS NEW\" & "PermissionsReportE.rpt"
       End If
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
        StrReportTitle = " ﬁ—Ì— ’·«ÕÌ«  «·„” Œœ„Ì‰" '& StrAccountName
    Else
        xReport.ParameterFields(1).AddCurrentValue cCompanyInfo.ArabCompanyName ' RPTCompany_Name_Eng
        StrReportTitle = "User Permissions Report"
    End If
    
    xReport.ParameterFields(2).AddCurrentValue " "
    xReport.ParameterFields(3).AddCurrentValue user_name
    
    If ShowAdvPerm.value = 0 Then
    xReport.ParameterFields(4).AddCurrentValue False
    Else
    xReport.ParameterFields(4).AddCurrentValue True
    End If
    If ShowDeniedUsers.value = 0 Then
    xReport.ParameterFields(5).AddCurrentValue False
    Else
    xReport.ParameterFields(5).AddCurrentValue True
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

Private Sub Text4_Change()
Dim StrSQL As String
 

    If SystemOptions.UserInterface = ArabicInterface Then
        StrSQL = "SELECT LTRIM(RTRIM(ScreenName)) as ScreenName ,LTRIM(RTRIM(screencaption)) as screencaption From Screens"
        StrSQL = StrSQL & " where screencaption like'%" & Text4.Text & "%'  "
       StrSQL = StrSQL & " order by LTRIM(RTRIM(screencaption)) "
    Else:
        StrSQL = "SELECT LTRIM(RTRIM(ScreenName)) as ScreenName ,LTRIM(RTRIM(ScreenTitleEng)) From Screens  "
        StrSQL = StrSQL & " where ScreenTitleEng like'%" & Text4.Text & "%'  "
       StrSQL = StrSQL & " order by  LTRIM(RTRIM(ScreenTitleEng)) "
       
 
    End If

    fill_combo DCScreen, StrSQL
    
    fill_combo DCScreen, StrSQL
End Sub

 '#################################################################################################################################
Private Sub XPDtbTransTimeFrom_CallbackKeyDown(ByVal KeyCode As Integer, ByVal Shift As Integer, ByVal CallbackField As String, CallbackDate As Date)

End Sub
