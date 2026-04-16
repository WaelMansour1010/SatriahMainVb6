VERSION 5.00
Object = "{C115893A-A3BF-43AF-B28D-69DB846077F3}#1.0#0"; "vsflex8u.ocx"
Object = "{0AFE7BE0-11B7-4A3E-978D-D4501E9A57FE}#1.0#0"; "c1sizer.ocx"
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{CFC0A331-9521-11D5-B9E6-5A06F6000000}#1.0#0"; "VDSCombo.DLL"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{FE5DCFAD-BC1D-11D2-94CF-004005455FAA}#1.4#0"; "ImpulseButton.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "mscomctl.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form FrmInpoutWorkOrder 
   Caption         =   "”‰œ «” ·«„ «‰ «Ã  «„"
   ClientHeight    =   7935
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   14790
   HelpContextID   =   100
   Icon            =   "FrmInpoutWorkOrder.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   Picture         =   "FrmInpoutWorkOrder.frx":038A
   RightToLeft     =   -1  'True
   ScaleHeight     =   7935
   ScaleWidth      =   14790
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
   Begin C1SizerLibCtl.C1Elastic C1ElasticMain 
      Height          =   7935
      Left            =   0
      TabIndex        =   9
      TabStop         =   0   'False
      Top             =   0
      Width           =   14790
      _cx             =   26088
      _cy             =   13996
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
      AutoSizeChildren=   8
      BorderWidth     =   1
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
      GridRows        =   5
      GridCols        =   1
      Frame           =   3
      FrameStyle      =   0
      FrameWidth      =   1
      FrameColor      =   -2147483628
      FrameShadow     =   -2147483632
      FloodStyle      =   1
      _GridInfo       =   $"FrmInpoutWorkOrder.frx":2B2C
      AccessibleName  =   ""
      AccessibleDescription=   ""
      AccessibleValue =   ""
      AccessibleRole  =   9
      Begin C1SizerLibCtl.C1Elastic Ele 
         Height          =   1695
         Index           =   5
         Left            =   15
         TabIndex        =   15
         TabStop         =   0   'False
         Top             =   630
         Width           =   14760
         _cx             =   26035
         _cy             =   2990
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
         Begin VB.ComboBox cmbTypeTrans 
            Height          =   315
            Left            =   9660
            Locked          =   -1  'True
            RightToLeft     =   -1  'True
            Style           =   2  'Dropdown List
            TabIndex        =   222
            Top             =   1200
            Width           =   3045
         End
         Begin VB.TextBox TxtBillComment 
            Alignment       =   1  'Right Justify
            Height          =   1035
            Left            =   3990
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   213
            Top             =   540
            Width           =   4395
         End
         Begin VB.TextBox TxtInspectionReport 
            Alignment       =   1  'Right Justify
            Height          =   375
            Left            =   3960
            RightToLeft     =   -1  'True
            TabIndex        =   211
            Top             =   120
            Width           =   1455
         End
         Begin VB.Frame Frame5 
            BackColor       =   &H00E2E9E9&
            Caption         =   "»Ì«‰«  ﬁÌœ «·”‰œ"
            Height          =   855
            Left            =   -240
            RightToLeft     =   -1  'True
            TabIndex        =   190
            Top             =   0
            Width           =   3255
            Begin VB.TextBox TxtNoteSerial 
               Alignment       =   1  'Right Justify
               Height          =   315
               Left            =   1440
               RightToLeft     =   -1  'True
               TabIndex        =   192
               Top             =   240
               Width           =   1185
            End
            Begin ImpulseButton.ISButton Cmd 
               CausesValidation=   0   'False
               Height          =   375
               Index           =   10
               Left            =   120
               TabIndex        =   191
               Top             =   240
               Width           =   1215
               _ExtentX        =   2143
               _ExtentY        =   661
               ButtonStyle     =   1
               ButtonPositionImage=   1
               Caption         =   "ÿ»«⁄Â «·ﬁÌœ"
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
         End
         Begin VB.TextBox TXTOrderNO1 
            Alignment       =   1  'Right Justify
            Height          =   375
            Left            =   6840
            RightToLeft     =   -1  'True
            TabIndex        =   187
            Top             =   120
            Width           =   1455
         End
         Begin VB.Frame Frame3 
            Caption         =   "„⁄·Ê„«  «·”‰œ"
            Height          =   1575
            Left            =   -1560
            RightToLeft     =   -1  'True
            TabIndex        =   158
            Top             =   720
            Visible         =   0   'False
            Width           =   2655
            Begin VB.TextBox Text7 
               Alignment       =   1  'Right Justify
               Height          =   285
               Left            =   120
               RightToLeft     =   -1  'True
               TabIndex        =   162
               Top             =   1080
               Width           =   1695
            End
            Begin VB.ComboBox CBOSource 
               Height          =   315
               Left            =   120
               RightToLeft     =   -1  'True
               Style           =   2  'Dropdown List
               TabIndex        =   159
               Top             =   360
               Width           =   1695
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               BackStyle       =   0  'Transparent
               Caption         =   "—ﬁ„ ›« Ê—… «·„‘ —Ì« "
               Height          =   270
               Index           =   59
               Left            =   960
               RightToLeft     =   -1  'True
               TabIndex        =   161
               Top             =   720
               Width           =   1500
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               BackStyle       =   0  'Transparent
               Caption         =   "„’œ—Â"
               Height          =   195
               Index           =   58
               Left            =   1965
               RightToLeft     =   -1  'True
               TabIndex        =   160
               Top             =   390
               Width           =   570
            End
         End
         Begin VB.TextBox Text6 
            Alignment       =   1  'Right Justify
            Height          =   315
            Left            =   12705
            RightToLeft     =   -1  'True
            TabIndex        =   128
            Top             =   480
            Width           =   930
         End
         Begin VB.Frame Frame1 
            BackColor       =   &H00E2E9E9&
            Height          =   1575
            Left            =   -4380
            RightToLeft     =   -1  'True
            TabIndex        =   119
            Top             =   1200
            Visible         =   0   'False
            Width           =   6645
            Begin VB.ComboBox CBOPriceType 
               Height          =   315
               Left            =   120
               RightToLeft     =   -1  'True
               Style           =   2  'Dropdown List
               TabIndex        =   156
               Top             =   600
               Width           =   1935
            End
            Begin VB.TextBox Text3 
               Alignment       =   1  'Right Justify
               Height          =   285
               Left            =   3360
               RightToLeft     =   -1  'True
               TabIndex        =   126
               Top             =   1080
               Width           =   2175
            End
            Begin VB.ComboBox CBoBasedON 
               Height          =   315
               Left            =   3360
               RightToLeft     =   -1  'True
               Style           =   2  'Dropdown List
               TabIndex        =   124
               Top             =   240
               Width           =   2175
            End
            Begin VB.TextBox TXT_order_no 
               Alignment       =   1  'Right Justify
               Height          =   285
               Left            =   120
               RightToLeft     =   -1  'True
               TabIndex        =   122
               Top             =   240
               Width           =   1935
            End
            Begin VB.TextBox Text4 
               Alignment       =   1  'Right Justify
               Height          =   285
               Left            =   3360
               RightToLeft     =   -1  'True
               TabIndex        =   121
               Top             =   720
               Width           =   2175
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "ÿ—ÌﬁÂ «· ”⁄Ì—"
               Height          =   195
               Index           =   57
               Left            =   2085
               RightToLeft     =   -1  'True
               TabIndex        =   157
               Top             =   630
               Width           =   1050
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "—ﬁ„   «·‘Õ‰Â"
               Height          =   270
               Index           =   51
               Left            =   5640
               RightToLeft     =   -1  'True
               TabIndex        =   127
               Top             =   1080
               Width           =   900
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "»‰«¡ ⁄·Ï"
               Height          =   195
               Index           =   56
               Left            =   5820
               RightToLeft     =   -1  'True
               TabIndex        =   125
               Top             =   270
               Width           =   645
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "«·—ﬁ„"
               Height          =   270
               Index           =   55
               Left            =   2280
               RightToLeft     =   -1  'True
               TabIndex        =   123
               Top             =   240
               Width           =   900
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "—ﬁ„  «·«⁄ „«œ "
               Height          =   270
               Index           =   52
               Left            =   5640
               RightToLeft     =   -1  'True
               TabIndex        =   120
               Top             =   720
               Width           =   900
            End
         End
         Begin VB.TextBox TxtNoteSerial1 
            Alignment       =   1  'Right Justify
            Height          =   315
            Left            =   11895
            RightToLeft     =   -1  'True
            TabIndex        =   116
            Top             =   120
            Width           =   1710
         End
         Begin VB.TextBox TXTNoteID 
            Alignment       =   1  'Right Justify
            Height          =   315
            Left            =   0
            RightToLeft     =   -1  'True
            TabIndex        =   115
            Top             =   960
            Visible         =   0   'False
            Width           =   1275
         End
         Begin VB.TextBox Text2 
            Alignment       =   1  'Right Justify
            Enabled         =   0   'False
            Height          =   285
            Left            =   6255
            RightToLeft     =   -1  'True
            TabIndex        =   114
            Top             =   960
            Visible         =   0   'False
            Width           =   540
         End
         Begin VB.TextBox Txt_EXport 
            Alignment       =   1  'Right Justify
            Height          =   405
            Left            =   6315
            RightToLeft     =   -1  'True
            TabIndex        =   111
            Top             =   1440
            Visible         =   0   'False
            Width           =   1140
         End
         Begin VB.TextBox Text1 
            Alignment       =   1  'Right Justify
            Enabled         =   0   'False
            Height          =   285
            Left            =   5355
            RightToLeft     =   -1  'True
            TabIndex        =   110
            Top             =   960
            Visible         =   0   'False
            Width           =   585
         End
         Begin VB.TextBox TxtCusID 
            Alignment       =   1  'Right Justify
            Height          =   315
            Left            =   8520
            RightToLeft     =   -1  'True
            TabIndex        =   2
            Top             =   1200
            Visible         =   0   'False
            Width           =   915
         End
         Begin VB.TextBox TxtStoreID 
            Alignment       =   1  'Right Justify
            Height          =   315
            Left            =   12720
            RightToLeft     =   -1  'True
            TabIndex        =   4
            Top             =   855
            Width           =   915
         End
         Begin VB.TextBox XPTxtBillID 
            Alignment       =   1  'Right Justify
            Height          =   345
            Left            =   5490
            Locked          =   -1  'True
            RightToLeft     =   -1  'True
            TabIndex        =   25
            Top             =   1020
            Visible         =   0   'False
            Width           =   1170
         End
         Begin VB.TextBox TxtTransSerial 
            Alignment       =   1  'Right Justify
            Height          =   345
            Left            =   6705
            RightToLeft     =   -1  'True
            TabIndex        =   0
            Top             =   900
            Visible         =   0   'False
            Width           =   630
         End
         Begin VB.TextBox XPTxtDiscountVal 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   585
            MaxLength       =   10
            RightToLeft     =   -1  'True
            TabIndex        =   7
            Top             =   480
            Visible         =   0   'False
            Width           =   2805
         End
         Begin VB.ComboBox XPCboDiscountType 
            Height          =   315
            Left            =   585
            RightToLeft     =   -1  'True
            Style           =   2  'Dropdown List
            TabIndex        =   6
            Top             =   1560
            Visible         =   0   'False
            Width           =   2805
         End
         Begin VB.ComboBox CboPayMentType 
            Height          =   315
            Left            =   570
            RightToLeft     =   -1  'True
            Style           =   2  'Dropdown List
            TabIndex        =   8
            Top             =   840
            Visible         =   0   'False
            Width           =   2805
         End
         Begin MSDataListLib.DataCombo DBCboClientName 
            Height          =   315
            Left            =   5460
            TabIndex        =   3
            Top             =   1200
            Visible         =   0   'False
            Width           =   3030
            _ExtentX        =   5345
            _ExtentY        =   556
            _Version        =   393216
            Text            =   ""
            RightToLeft     =   -1  'True
         End
         Begin MSDataListLib.DataCombo DCboStoreName 
            Height          =   315
            Left            =   9660
            TabIndex        =   5
            Top             =   840
            Width           =   3030
            _ExtentX        =   5345
            _ExtentY        =   556
            _Version        =   393216
            Text            =   ""
            RightToLeft     =   -1  'True
         End
         Begin MSComCtl2.DTPicker XPDtbBill 
            Height          =   315
            Left            =   9675
            TabIndex        =   1
            Top             =   120
            Width           =   1290
            _ExtentX        =   2275
            _ExtentY        =   556
            _Version        =   393216
            Format          =   232980481
            CurrentDate     =   38784
         End
         Begin ImpulseButton.ISButton XPBtnNewClients 
            Height          =   420
            Left            =   6690
            TabIndex        =   16
            TabStop         =   0   'False
            Top             =   795
            Width           =   960
            _ExtentX        =   1693
            _ExtentY        =   741
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
            ButtonImage     =   "FrmInpoutWorkOrder.frx":2B90
            ColorButton     =   14871017
            ColorHighlight  =   16777215
            ColorShadow     =   -2147483631
            ColorOutline    =   -2147483631
            DrawFocusRectangle=   0   'False
            DisabledImageStyle=   1
         End
         Begin ImpulseButton.ISButton CmdConvert 
            Height          =   390
            Left            =   480
            TabIndex        =   113
            Top             =   1200
            Visible         =   0   'False
            Width           =   2235
            _ExtentX        =   3942
            _ExtentY        =   688
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
         Begin MSDataListLib.DataCombo dcBranch 
            Height          =   315
            Left            =   9660
            TabIndex        =   117
            Top             =   480
            Width           =   3015
            _ExtentX        =   5318
            _ExtentY        =   556
            _Version        =   393216
            BackColor       =   16777215
            Text            =   ""
            RightToLeft     =   -1  'True
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "„·«ÕŸ« "
            Height          =   315
            Index           =   86
            Left            =   8520
            TabIndex        =   214
            Top             =   600
            Width           =   810
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "—ﬁ„  ﬁ—Ì— «·›Õ’"
            Height          =   270
            Index           =   65
            Left            =   5535
            RightToLeft     =   -1  'True
            TabIndex        =   212
            Top             =   240
            Width           =   1170
         End
         Begin VB.Label Label3 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "«·›—⁄"
            ForeColor       =   &H00000000&
            Height          =   375
            Left            =   13695
            TabIndex        =   118
            Top             =   480
            Width           =   975
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            Caption         =   "«·„’—Ê›«  «·«Œ—Ï"
            Height          =   255
            Left            =   4140
            RightToLeft     =   -1  'True
            TabIndex        =   112
            Top             =   2160
            Width           =   2790
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "«·„Œ“‰"
            Height          =   210
            Index           =   4
            Left            =   13935
            RightToLeft     =   -1  'True
            TabIndex        =   24
            Top             =   855
            Width           =   735
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "‰Ê⁄ «·Œ’„"
            Height          =   345
            Index           =   5
            Left            =   3390
            RightToLeft     =   -1  'True
            TabIndex        =   23
            Top             =   1335
            Visible         =   0   'False
            Width           =   1575
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "«·„Ê—œ"
            Height          =   285
            Index           =   6
            Left            =   8175
            RightToLeft     =   -1  'True
            TabIndex        =   22
            Top             =   1215
            Visible         =   0   'False
            Width           =   1095
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "—ﬁ„ «·”‰œ"
            Height          =   300
            Index           =   8
            Left            =   13935
            RightToLeft     =   -1  'True
            TabIndex        =   21
            Top             =   105
            Width           =   735
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "„‰ «„— «‰ «Ã —ﬁ„"
            Height          =   270
            Index           =   9
            Left            =   8415
            RightToLeft     =   -1  'True
            TabIndex        =   20
            Top             =   120
            Width           =   1170
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   " «—ÌŒ «·”‰œ"
            Height          =   270
            Index           =   7
            Left            =   10950
            RightToLeft     =   -1  'True
            TabIndex        =   19
            Top             =   120
            Width           =   840
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "ÿ—Ìﬁ… «·œ›⁄"
            Height          =   300
            Index           =   10
            Left            =   3390
            RightToLeft     =   -1  'True
            TabIndex        =   18
            Top             =   855
            Visible         =   0   'False
            Width           =   1575
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "ﬁÌ„… «·Œ’„"
            Height          =   285
            Index           =   11
            Left            =   3390
            RightToLeft     =   -1  'True
            TabIndex        =   17
            Top             =   480
            Visible         =   0   'False
            Width           =   1575
         End
      End
      Begin C1SizerLibCtl.C1Elastic Ele 
         Height          =   600
         Index           =   6
         Left            =   15
         TabIndex        =   10
         TabStop         =   0   'False
         Top             =   15
         Width           =   14760
         _cx             =   26035
         _cy             =   1058
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
         Caption         =   "”‰œ «” ·«„ «‰ «Ã  «„ "
         Align           =   0
         AutoSizeChildren=   7
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
         Begin VB.CheckBox chkIgnorDetails 
            Alignment       =   1  'Right Justify
            Caption         =   " Ã«Â· «· ›«’Ì·"
            Height          =   270
            Left            =   11610
            RightToLeft     =   -1  'True
            TabIndex        =   221
            Top             =   270
            Value           =   1  'Checked
            Visible         =   0   'False
            Width           =   1875
         End
         Begin VB.CommandButton cmdReSave 
            Caption         =   "÷»ÿ «·Õ—ﬂ« "
            Height          =   285
            Left            =   10500
            TabIndex        =   218
            Top             =   240
            Visible         =   0   'False
            Width           =   1125
         End
         Begin VB.TextBox txtPassword 
            Height          =   315
            IMEMode         =   3  'DISABLE
            Left            =   4320
            PasswordChar    =   "*"
            TabIndex        =   217
            Top             =   330
            Width           =   735
         End
         Begin VB.CheckBox chkIsBranch 
            Caption         =   "»«·›—⁄"
            Height          =   225
            Index           =   0
            Left            =   6420
            TabIndex        =   216
            Top             =   330
            Visible         =   0   'False
            Width           =   945
         End
         Begin VB.CheckBox chkIsBranch 
            Caption         =   "»œÊ‰ ﬁÌœ"
            Height          =   225
            Index           =   1
            Left            =   5160
            TabIndex        =   215
            Top             =   360
            Visible         =   0   'False
            Width           =   945
         End
         Begin VB.TextBox oldtxtNoteSerial1 
            Alignment       =   1  'Right Justify
            Height          =   315
            Left            =   0
            RightToLeft     =   -1  'True
            TabIndex        =   188
            Top             =   0
            Visible         =   0   'False
            Width           =   1425
         End
         Begin VB.TextBox TxtModFlg 
            Alignment       =   1  'Right Justify
            BackColor       =   &H000000FF&
            Height          =   345
            Left            =   3735
            RightToLeft     =   -1  'True
            TabIndex        =   105
            Top             =   90
            Visible         =   0   'False
            Width           =   810
         End
         Begin VB.TextBox TxtFillData 
            Alignment       =   1  'Right Justify
            BackColor       =   &H000000FF&
            Height          =   315
            Left            =   4590
            RightToLeft     =   -1  'True
            TabIndex        =   104
            Top             =   105
            Visible         =   0   'False
            Width           =   540
         End
         Begin ImpulseButton.ISButton CmdNotes 
            Height          =   390
            Left            =   4710
            TabIndex        =   31
            Top             =   90
            Width           =   1185
            _ExtentX        =   2090
            _ExtentY        =   688
            ButtonStyle     =   1
            ButtonPositionImage=   3
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
            ButtonImage     =   "FrmInpoutWorkOrder.frx":2F2A
            DrawFocusRectangle=   0   'False
         End
         Begin ImpulseButton.ISButton XPBtnMove 
            Height          =   345
            Index           =   0
            Left            =   2520
            TabIndex        =   12
            Top             =   105
            Width           =   960
            _ExtentX        =   1693
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
            ButtonImage     =   "FrmInpoutWorkOrder.frx":32C4
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
            Left            =   1350
            TabIndex        =   13
            Top             =   105
            Width           =   975
            _ExtentX        =   1720
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
            ButtonImage     =   "FrmInpoutWorkOrder.frx":365E
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
            Left            =   3555
            TabIndex        =   11
            Top             =   105
            Width           =   1005
            _ExtentX        =   1773
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
            ButtonImage     =   "FrmInpoutWorkOrder.frx":39F8
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
            Left            =   120
            TabIndex        =   14
            Top             =   105
            Width           =   1095
            _ExtentX        =   1931
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
            ButtonImage     =   "FrmInpoutWorkOrder.frx":3D92
            ColorHighlight  =   4194304
            ColorHoverText  =   16777215
            ColorShadow     =   -2147483631
            ColorOutline    =   -2147483631
            DrawFocusRectangle=   0   'False
            DisabledImageStyle=   1
            ColorToggledHoverText=   16777215
            ColorTextShadow =   16777215
         End
         Begin ImpulseButton.ISButton CmdRetruns 
            Height          =   390
            Left            =   5955
            TabIndex        =   32
            Top             =   90
            Width           =   1035
            _ExtentX        =   1826
            _ExtentY        =   688
            ButtonStyle     =   1
            ButtonPositionImage=   3
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
            ButtonImage     =   "FrmInpoutWorkOrder.frx":412C
            DrawFocusRectangle=   0   'False
         End
         Begin ImpulseButton.ISButton CmdInfo 
            Height          =   480
            Left            =   8610
            TabIndex        =   155
            Top             =   0
            Visible         =   0   'False
            Width           =   615
            _ExtentX        =   1085
            _ExtentY        =   847
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
            ButtonImage     =   "FrmInpoutWorkOrder.frx":46C6
            ButtonImageHover=   "FrmInpoutWorkOrder.frx":53A0
            ColorButton     =   16777215
            DrawFocusRectangle=   0   'False
         End
         Begin MSComCtl2.DTPicker txtFromDateReSave 
            Height          =   315
            Left            =   9090
            TabIndex        =   219
            Top             =   270
            Visible         =   0   'False
            Width           =   1380
            _ExtentX        =   2434
            _ExtentY        =   556
            _Version        =   393216
            Format          =   232980481
            CurrentDate     =   38784
         End
         Begin MSComCtl2.DTPicker txtToDateReSave 
            Height          =   315
            Left            =   7530
            TabIndex        =   220
            Top             =   300
            Visible         =   0   'False
            Width           =   1560
            _ExtentX        =   2752
            _ExtentY        =   556
            _Version        =   393216
            Format          =   232980481
            CurrentDate     =   38784
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            BackStyle       =   0  'Transparent
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
            Index           =   62
            Left            =   4680
            RightToLeft     =   -1  'True
            TabIndex        =   189
            Top             =   0
            Width           =   6795
         End
      End
      Begin C1SizerLibCtl.C1Tab XPTab301 
         Height          =   4575
         Left            =   15
         TabIndex        =   26
         Top             =   2340
         Width           =   14760
         _cx             =   26035
         _cy             =   8070
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
         BackColor       =   14871017
         ForeColor       =   0
         FrontTabColor   =   14871017
         BackTabColor    =   12648447
         TabOutlineColor =   -2147483632
         FrontTabForeColor=   16711680
         Caption         =   "«·√’‰«›|«·„’«—Ì› «· ﬁœÌ—ÌÂ|≈” ﬁÿ«⁄«  ⁄·Ï «·›« Ê—…|«·„’—Ê›« |«·›Ê« Ì— «·„«·Ì…|«·„’—Ê›«  «· ﬁœÌ—Ì…"
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
         Picture(0)      =   "FrmInpoutWorkOrder.frx":607A
         Picture(1)      =   "FrmInpoutWorkOrder.frx":6414
         Flags(1)        =   2
         Flags(2)        =   2
         Flags(3)        =   2
         Flags(4)        =   2
         Flags(5)        =   2
         Begin C1SizerLibCtl.C1Elastic C1Elastic2 
            Height          =   4110
            Left            =   16605
            TabIndex        =   201
            TabStop         =   0   'False
            Top             =   45
            Width           =   14670
            _cx             =   25876
            _cy             =   7250
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
            Begin VB.Frame Frame2 
               Height          =   4110
               Left            =   0
               RightToLeft     =   -1  'True
               TabIndex        =   202
               Top             =   0
               Width           =   14670
               Begin VB.CommandButton Command6 
                  Caption         =   "Command6"
                  Height          =   375
                  Left            =   6840
                  RightToLeft     =   -1  'True
                  TabIndex        =   206
                  Top             =   3000
                  Visible         =   0   'False
                  Width           =   1815
               End
               Begin VB.TextBox TXTToTAlELSHahn 
                  Alignment       =   1  'Right Justify
                  Height          =   405
                  Left            =   0
                  Locked          =   -1  'True
                  RightToLeft     =   -1  'True
                  TabIndex        =   205
                  Text            =   "0"
                  Top             =   3000
                  Visible         =   0   'False
                  Width           =   1935
               End
               Begin VB.TextBox Text5 
                  Alignment       =   1  'Right Justify
                  Height          =   405
                  Left            =   10200
                  RightToLeft     =   -1  'True
                  TabIndex        =   204
                  Top             =   2880
                  Width           =   1890
               End
               Begin VB.CommandButton Command2 
                  Caption         =   "⁄—÷ «·„’—Ê›« "
                  Height          =   480
                  Left            =   11880
                  RightToLeft     =   -1  'True
                  TabIndex        =   203
                  Top             =   3360
                  Visible         =   0   'False
                  Width           =   2220
               End
               Begin VSFlex8UCtl.VSFlexGrid Grid 
                  Height          =   2325
                  Left            =   120
                  TabIndex        =   207
                  Tag             =   "1"
                  Top             =   480
                  Width           =   14055
                  _cx             =   24791
                  _cy             =   4101
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
                  Rows            =   50
                  Cols            =   10
                  FixedRows       =   1
                  FixedCols       =   1
                  RowHeightMin    =   300
                  RowHeightMax    =   0
                  ColWidthMin     =   0
                  ColWidthMax     =   0
                  ExtendLastCol   =   0   'False
                  FormatString    =   $"FrmInpoutWorkOrder.frx":67AE
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
                  BackColor       =   &H00E2E9E9&
                  BackStyle       =   0  'Transparent
                  Caption         =   "«Ã„«·Ì «·„’—Ê›« "
                  Height          =   285
                  Index           =   60
                  Left            =   1680
                  RightToLeft     =   -1  'True
                  TabIndex        =   210
                  Top             =   3000
                  Visible         =   0   'False
                  Width           =   1800
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  BackStyle       =   0  'Transparent
                  Caption         =   "”‰œ«  «·’—›"
                  Height          =   285
                  Index           =   54
                  Left            =   11280
                  RightToLeft     =   -1  'True
                  TabIndex        =   209
                  Top             =   120
                  Width           =   2520
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  BackStyle       =   0  'Transparent
                  Caption         =   "«Ã„«·Ì  ”‰œ«  «·„’—Ê›« "
                  Height          =   285
                  Index           =   53
                  Left            =   12150
                  RightToLeft     =   -1  'True
                  TabIndex        =   208
                  Top             =   3000
                  Width           =   1920
               End
            End
         End
         Begin C1SizerLibCtl.C1Elastic C1Elastic1 
            Height          =   4110
            Left            =   16305
            TabIndex        =   193
            TabStop         =   0   'False
            Top             =   45
            Width           =   14670
            _cx             =   25876
            _cy             =   7250
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
            Begin VB.Frame Frame4 
               Height          =   4110
               Left            =   0
               RightToLeft     =   -1  'True
               TabIndex        =   194
               Top             =   0
               Width           =   14670
               Begin VB.CommandButton Command5 
                  Caption         =   " Œ’Ì’"
                  Height          =   480
                  Left            =   12120
                  RightToLeft     =   -1  'True
                  TabIndex        =   197
                  Top             =   3240
                  Visible         =   0   'False
                  Width           =   2220
               End
               Begin VB.CommandButton Command4 
                  Caption         =   "⁄—÷ «·›Ê« Ì— «·„«·Ì…"
                  Height          =   480
                  Left            =   7560
                  RightToLeft     =   -1  'True
                  TabIndex        =   196
                  Top             =   2880
                  Width           =   2220
               End
               Begin VB.TextBox txt_total_bill 
                  Height          =   405
                  Left            =   10200
                  RightToLeft     =   -1  'True
                  TabIndex        =   195
                  Top             =   2880
                  Width           =   1770
               End
               Begin VSFlex8UCtl.VSFlexGrid grid4 
                  Height          =   2325
                  Left            =   240
                  TabIndex        =   198
                  Tag             =   "1"
                  Top             =   480
                  Width           =   14055
                  _cx             =   24791
                  _cy             =   4101
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
                  Rows            =   50
                  Cols            =   11
                  FixedRows       =   1
                  FixedCols       =   1
                  RowHeightMin    =   300
                  RowHeightMax    =   0
                  ColWidthMin     =   0
                  ColWidthMax     =   0
                  ExtendLastCol   =   0   'False
                  FormatString    =   $"FrmInpoutWorkOrder.frx":692E
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
                  BackColor       =   &H00E2E9E9&
                  BackStyle       =   0  'Transparent
                  Caption         =   "«·›Ê« Ì— «·„«·ÌÂ"
                  Height          =   285
                  Index           =   64
                  Left            =   12000
                  RightToLeft     =   -1  'True
                  TabIndex        =   200
                  Top             =   120
                  Width           =   2040
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  BackStyle       =   0  'Transparent
                  Caption         =   "«Ã„«·Ì «·›Ê« Ì—"
                  Height          =   285
                  Index           =   61
                  Left            =   12150
                  RightToLeft     =   -1  'True
                  TabIndex        =   199
                  Top             =   2880
                  Width           =   2040
               End
            End
         End
         Begin C1SizerLibCtl.C1Elastic Ele 
            Height          =   4110
            Index           =   0
            Left            =   45
            TabIndex        =   27
            TabStop         =   0   'False
            Top             =   45
            Width           =   14670
            _cx             =   25876
            _cy             =   7250
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
            AutoSizeChildren=   8
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
            GridRows        =   3
            GridCols        =   1
            Frame           =   3
            FrameStyle      =   0
            FrameWidth      =   1
            FrameColor      =   -2147483628
            FrameShadow     =   -2147483632
            FloodStyle      =   1
            _GridInfo       =   $"FrmInpoutWorkOrder.frx":6AF2
            AccessibleName  =   ""
            AccessibleDescription=   ""
            AccessibleValue =   ""
            AccessibleRole  =   9
            Begin VSFlex8UCtl.VSFlexGrid FG 
               Height          =   2970
               Left            =   30
               TabIndex        =   28
               Top             =   735
               Width           =   14610
               _cx             =   25770
               _cy             =   5239
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
               Cols            =   21
               FixedRows       =   1
               FixedCols       =   1
               RowHeightMin    =   300
               RowHeightMax    =   0
               ColWidthMin     =   0
               ColWidthMax     =   0
               ExtendLastCol   =   -1  'True
               FormatString    =   $"FrmInpoutWorkOrder.frx":6B3F
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
            Begin MSComctlLib.Toolbar TBar 
               Height          =   630
               Left            =   30
               TabIndex        =   29
               Top             =   3720
               Width           =   14610
               _ExtentX        =   25770
               _ExtentY        =   1111
               ButtonWidth     =   609
               ButtonHeight    =   1005
               Appearance      =   1
               _Version        =   393216
            End
            Begin C1SizerLibCtl.C1Elastic Ele 
               Height          =   690
               Index           =   4
               Left            =   30
               TabIndex        =   163
               TabStop         =   0   'False
               Top             =   30
               Width           =   14610
               _cx             =   25770
               _cy             =   1217
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
               Begin VB.TextBox TxtPrice 
                  Alignment       =   1  'Right Justify
                  Height          =   345
                  Left            =   735
                  MaxLength       =   10
                  RightToLeft     =   -1  'True
                  TabIndex        =   167
                  Top             =   315
                  Width           =   1305
               End
               Begin VB.TextBox TxtSerial 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00FFFFFF&
                  Enabled         =   0   'False
                  Height          =   345
                  Left            =   3705
                  MaxLength       =   20
                  RightToLeft     =   -1  'True
                  TabIndex        =   166
                  Top             =   315
                  Width           =   3225
               End
               Begin VB.TextBox TxtQuantity 
                  Alignment       =   1  'Right Justify
                  Height          =   345
                  Left            =   2040
                  MaxLength       =   10
                  RightToLeft     =   -1  'True
                  TabIndex        =   165
                  Top             =   315
                  Width           =   1665
               End
               Begin VB.ComboBox CboItemCase 
                  Height          =   315
                  Left            =   7035
                  RightToLeft     =   -1  'True
                  Style           =   2  'Dropdown List
                  TabIndex        =   164
                  Top             =   315
                  Width           =   2205
               End
               Begin MSDataListLib.DataCombo DCboItemsName 
                  Height          =   315
                  Left            =   9345
                  TabIndex        =   168
                  Top             =   315
                  Width           =   3330
                  _ExtentX        =   5874
                  _ExtentY        =   556
                  _Version        =   393216
                  Text            =   ""
                  RightToLeft     =   -1  'True
               End
               Begin MSDataListLib.DataCombo DCboItemsCode 
                  Height          =   315
                  Left            =   12765
                  TabIndex        =   169
                  Top             =   315
                  Width           =   1845
                  _ExtentX        =   3254
                  _ExtentY        =   556
                  _Version        =   393216
                  Text            =   ""
                  RightToLeft     =   -1  'True
               End
               Begin ImpulseButton.ISButton CmdAdd 
                  Height          =   345
                  Left            =   0
                  TabIndex        =   170
                  Top             =   315
                  Width           =   735
                  _ExtentX        =   1296
                  _ExtentY        =   609
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
                  ButtonImage     =   "FrmInpoutWorkOrder.frx":6E9D
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
               Begin VB.Label lbl 
                  Alignment       =   2  'Center
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "«· ﬂ·›…"
                  Height          =   270
                  Index           =   26
                  Left            =   735
                  RightToLeft     =   -1  'True
                  TabIndex        =   176
                  Top             =   30
                  Width           =   1305
               End
               Begin VB.Label lbl 
                  Alignment       =   2  'Center
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "«·ﬂ„Ì…"
                  Height          =   270
                  Index           =   27
                  Left            =   2040
                  RightToLeft     =   -1  'True
                  TabIndex        =   175
                  Top             =   30
                  Width           =   1665
               End
               Begin VB.Label lbl 
                  Alignment       =   2  'Center
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "«·”Ì—Ì«·"
                  Height          =   270
                  Index           =   28
                  Left            =   3705
                  RightToLeft     =   -1  'True
                  TabIndex        =   174
                  Top             =   30
                  Width           =   3225
               End
               Begin VB.Label lbl 
                  Alignment       =   2  'Center
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "Õ«·… «·’‰›"
                  Height          =   270
                  Index           =   29
                  Left            =   7035
                  RightToLeft     =   -1  'True
                  TabIndex        =   173
                  Top             =   30
                  Width           =   2205
               End
               Begin VB.Label lbl 
                  Alignment       =   2  'Center
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "≈”„ «·’‰›"
                  Height          =   270
                  Index           =   30
                  Left            =   9345
                  RightToLeft     =   -1  'True
                  TabIndex        =   172
                  Top             =   30
                  Width           =   3330
               End
               Begin VB.Label lbl 
                  Alignment       =   2  'Center
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "ﬂÊœ «·’‰›"
                  Height          =   270
                  Index           =   31
                  Left            =   12765
                  RightToLeft     =   -1  'True
                  TabIndex        =   171
                  Top             =   30
                  Width           =   1845
               End
            End
            Begin VB.Label LblItemsCount 
               Alignment       =   2  'Center
               BackColor       =   &H00404040&
               ForeColor       =   &H0000FFFF&
               Height          =   360
               Left            =   30
               RightToLeft     =   -1  'True
               TabIndex        =   30
               Top             =   3720
               Width           =   14610
            End
         End
         Begin C1SizerLibCtl.C1Elastic Ele 
            Height          =   4110
            Index           =   2
            Left            =   15405
            TabIndex        =   33
            TabStop         =   0   'False
            Top             =   45
            Width           =   14670
            _cx             =   25876
            _cy             =   7250
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
            BackColor       =   255
            ForeColor       =   -2147483630
            FloodColor      =   6553600
            ForeColorDisabled=   -2147483631
            Caption         =   ""
            Align           =   0
            AutoSizeChildren=   8
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
            GridRows        =   3
            GridCols        =   4
            Frame           =   3
            FrameStyle      =   0
            FrameWidth      =   1
            FrameColor      =   -2147483628
            FrameShadow     =   -2147483632
            FloodStyle      =   1
            _GridInfo       =   $"FrmInpoutWorkOrder.frx":7237
            AccessibleName  =   ""
            AccessibleDescription=   ""
            AccessibleValue =   ""
            AccessibleRole  =   9
            Begin C1SizerLibCtl.C1Elastic Ele 
               Height          =   1875
               Index           =   10
               Left            =   13815
               TabIndex        =   34
               TabStop         =   0   'False
               Top             =   2235
               Visible         =   0   'False
               Width           =   855
               _cx             =   1508
               _cy             =   3307
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
               AutoSizeChildren=   8
               BorderWidth     =   1
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
               GridRows        =   4
               GridCols        =   4
               Frame           =   3
               FrameStyle      =   0
               FrameWidth      =   1
               FrameColor      =   -2147483628
               FrameShadow     =   -2147483632
               FloodStyle      =   1
               _GridInfo       =   $"FrmInpoutWorkOrder.frx":72A7
               AccessibleName  =   ""
               AccessibleDescription=   ""
               AccessibleValue =   ""
               AccessibleRole  =   9
               Begin C1SizerLibCtl.C1Elastic Ele 
                  Height          =   120
                  Index           =   14
                  Left            =   15
                  TabIndex        =   35
                  TabStop         =   0   'False
                  Top             =   15
                  Width           =   825
                  _cx             =   1455
                  _cy             =   212
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
                  Begin VB.CheckBox XPChkPayType 
                     Alignment       =   1  'Right Justify
                     BackColor       =   &H00E2E9E9&
                     Caption         =   "‘Ìﬂ« "
                     Height          =   105
                     Index           =   2
                     Left            =   480
                     RightToLeft     =   -1  'True
                     TabIndex        =   36
                     Top             =   30
                     Width           =   60
                  End
                  Begin ImpulseButton.ISButton CmdCheque 
                     Height          =   105
                     Left            =   150
                     TabIndex        =   37
                     Top             =   30
                     Width           =   75
                     _ExtentX        =   132
                     _ExtentY        =   185
                     ButtonStyle     =   1
                     Caption         =   " ”ÃÌ· «·‘Ìﬂ« "
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
                     DrawFocusRectangle=   0   'False
                  End
                  Begin VB.Label lbl 
                     Alignment       =   1  'Right Justify
                     BackColor       =   &H00E2E9E9&
                     Height          =   105
                     Index           =   18
                     Left            =   225
                     RightToLeft     =   -1  'True
                     TabIndex        =   41
                     Top             =   30
                     Width           =   45
                  End
                  Begin VB.Label lbl 
                     Alignment       =   1  'Right Justify
                     BackColor       =   &H00E2E9E9&
                     Caption         =   "≈Ã„«·Ï ﬁÌ„… «·‘Ìﬂ« "
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
                     Height          =   105
                     Index           =   16
                     Left            =   285
                     RightToLeft     =   -1  'True
                     TabIndex        =   40
                     Top             =   30
                     Width           =   75
                  End
                  Begin VB.Label lbl 
                     Alignment       =   1  'Right Justify
                     BackColor       =   &H00E2E9E9&
                     Caption         =   "⁄œœ «·‘Ìﬂ« "
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
                     Height          =   105
                     Index           =   17
                     Left            =   420
                     RightToLeft     =   -1  'True
                     TabIndex        =   39
                     Top             =   30
                     Width           =   60
                  End
                  Begin VB.Label lbl 
                     Alignment       =   1  'Right Justify
                     BackColor       =   &H00E2E9E9&
                     Height          =   105
                     Index           =   19
                     Left            =   360
                     RightToLeft     =   -1  'True
                     TabIndex        =   38
                     Top             =   30
                     Width           =   60
                  End
               End
               Begin VSFlex8UCtl.VSFlexGrid FgCheques 
                  Height          =   1710
                  Left            =   165
                  TabIndex        =   42
                  Top             =   150
                  Width           =   675
                  _cx             =   1191
                  _cy             =   3016
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
                  FormatString    =   $"FrmInpoutWorkOrder.frx":7317
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
            End
            Begin C1SizerLibCtl.C1Elastic Ele 
               Height          =   3870
               Index           =   7
               Left            =   13815
               TabIndex        =   43
               TabStop         =   0   'False
               Top             =   240
               Width           =   855
               _cx             =   1508
               _cy             =   6826
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
               AutoSizeChildren=   8
               BorderWidth     =   1
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
               GridRows        =   3
               GridCols        =   4
               Frame           =   3
               FrameStyle      =   0
               FrameWidth      =   1
               FrameColor      =   -2147483628
               FrameShadow     =   -2147483632
               FloodStyle      =   1
               _GridInfo       =   $"FrmInpoutWorkOrder.frx":744B
               AccessibleName  =   ""
               AccessibleDescription=   ""
               AccessibleValue =   ""
               AccessibleRole  =   9
               Begin VSFlex8UCtl.VSFlexGrid FgInstallments 
                  Height          =   3750
                  Left            =   645
                  TabIndex        =   44
                  Top             =   105
                  Visible         =   0   'False
                  Width           =   195
                  _cx             =   344
                  _cy             =   6615
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
                  Cols            =   5
                  FixedRows       =   1
                  FixedCols       =   1
                  RowHeightMin    =   0
                  RowHeightMax    =   0
                  ColWidthMin     =   0
                  ColWidthMax     =   0
                  ExtendLastCol   =   -1  'True
                  FormatString    =   $"FrmInpoutWorkOrder.frx":74B3
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
               Begin C1SizerLibCtl.C1Elastic Ele 
                  Height          =   60
                  Index           =   13
                  Left            =   15
                  TabIndex        =   45
                  TabStop         =   0   'False
                  Top             =   3795
                  Width           =   825
                  _cx             =   1455
                  _cy             =   106
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
                  Begin VB.Label LblInstallmentType 
                     Alignment       =   1  'Right Justify
                     BackColor       =   &H00E2E9E9&
                     ForeColor       =   &H00FF0000&
                     Height          =   45
                     Left            =   75
                     RightToLeft     =   -1  'True
                     TabIndex        =   60
                     Top             =   15
                     Width           =   30
                  End
                  Begin VB.Label lbl 
                     Alignment       =   1  'Right Justify
                     BackColor       =   &H00E2E9E9&
                     Caption         =   "› —… «· ﬁ”Ìÿ"
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
                     Height          =   45
                     Index           =   42
                     Left            =   120
                     RightToLeft     =   -1  'True
                     TabIndex        =   59
                     Top             =   15
                     Width           =   60
                  End
                  Begin VB.Label LblFirstInstallDate 
                     Alignment       =   1  'Right Justify
                     BackColor       =   &H00E2E9E9&
                     Height          =   45
                     Left            =   180
                     RightToLeft     =   -1  'True
                     TabIndex        =   58
                     Top             =   15
                     Width           =   45
                  End
                  Begin VB.Label lbl 
                     Alignment       =   1  'Right Justify
                     BackColor       =   &H00E2E9E9&
                     Caption         =   "«Ê· ﬁ”ÿ"
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
                     Height          =   45
                     Index           =   40
                     Left            =   225
                     RightToLeft     =   -1  'True
                     TabIndex        =   57
                     Top             =   15
                     Width           =   30
                  End
                  Begin VB.Label LblInstallCount 
                     Alignment       =   1  'Right Justify
                     BackColor       =   &H00E2E9E9&
                     Height          =   45
                     Left            =   255
                     RightToLeft     =   -1  'True
                     TabIndex        =   56
                     Top             =   15
                     Width           =   30
                  End
                  Begin VB.Label lbl 
                     Alignment       =   1  'Right Justify
                     BackColor       =   &H00E2E9E9&
                     Caption         =   "⁄œœ «·√ﬁ”«ÿ"
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
                     Height          =   45
                     Index           =   38
                     Left            =   285
                     RightToLeft     =   -1  'True
                     TabIndex        =   55
                     Top             =   15
                     Width           =   45
                  End
                  Begin VB.Label LblInstallTotal 
                     Alignment       =   1  'Right Justify
                     BackColor       =   &H00E2E9E9&
                     Height          =   45
                     Left            =   330
                     RightToLeft     =   -1  'True
                     TabIndex        =   54
                     Top             =   15
                     Width           =   30
                  End
                  Begin VB.Label lbl 
                     Alignment       =   1  'Right Justify
                     BackColor       =   &H00E2E9E9&
                     Caption         =   "«·„»·€ «·ﬂ·Ï"
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
                     Height          =   45
                     Index           =   36
                     Left            =   360
                     RightToLeft     =   -1  'True
                     TabIndex        =   53
                     Top             =   15
                     Width           =   45
                  End
                  Begin VB.Label LblPrecenType 
                     Alignment       =   1  'Right Justify
                     BackColor       =   &H00E2E9E9&
                     Height          =   45
                     Left            =   450
                     RightToLeft     =   -1  'True
                     TabIndex        =   52
                     Top             =   15
                     Width           =   45
                  End
                  Begin VB.Label lbl 
                     Alignment       =   1  'Right Justify
                     BackColor       =   &H00E2E9E9&
                     Caption         =   "‰Ê⁄ «·›«∆œ…"
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
                     Height          =   45
                     Index           =   34
                     Left            =   495
                     RightToLeft     =   -1  'True
                     TabIndex        =   51
                     Top             =   15
                     Width           =   45
                  End
                  Begin VB.Label lbl 
                     Alignment       =   1  'Right Justify
                     BackColor       =   &H00E2E9E9&
                     Caption         =   "‰”»… «·›«∆œ…"
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
                     Height          =   45
                     Index           =   35
                     Left            =   435
                     RightToLeft     =   -1  'True
                     TabIndex        =   50
                     Top             =   15
                     Width           =   15
                  End
                  Begin VB.Label LblPrecenValue 
                     Alignment       =   1  'Right Justify
                     BackColor       =   &H00E2E9E9&
                     Height          =   45
                     Left            =   420
                     RightToLeft     =   -1  'True
                     TabIndex        =   49
                     Top             =   15
                     Width           =   15
                  End
                  Begin VB.Label LblInstallSeprator 
                     Alignment       =   1  'Right Justify
                     BackColor       =   &H00E2E9E9&
                     ForeColor       =   &H00FF0000&
                     Height          =   45
                     Left            =   105
                     RightToLeft     =   -1  'True
                     TabIndex        =   48
                     Top             =   15
                     Width           =   15
                  End
                  Begin VB.Label LblStartValue 
                     Alignment       =   1  'Right Justify
                     BackColor       =   &H00E2E9E9&
                     Height          =   45
                     Left            =   0
                     RightToLeft     =   -1  'True
                     TabIndex        =   47
                     Top             =   15
                     Width           =   15
                  End
                  Begin VB.Label lbl 
                     Alignment       =   1  'Right Justify
                     BackColor       =   &H00E2E9E9&
                     Caption         =   "«·ﬁÌ„… «·„»œ∆Ì…"
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
                     Height          =   45
                     Index           =   37
                     Left            =   15
                     RightToLeft     =   -1  'True
                     TabIndex        =   46
                     Top             =   15
                     Width           =   60
                  End
               End
               Begin C1SizerLibCtl.C1Elastic Ele 
                  Height          =   75
                  Index           =   12
                  Left            =   15
                  TabIndex        =   61
                  TabStop         =   0   'False
                  Top             =   15
                  Width           =   825
                  _cx             =   1455
                  _cy             =   132
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
                  Begin VB.CheckBox XPChkPayType 
                     Alignment       =   1  'Right Justify
                     BackColor       =   &H00E2E9E9&
                     Caption         =   "¬Ã· "
                     Height          =   60
                     Index           =   1
                     Left            =   480
                     RightToLeft     =   -1  'True
                     TabIndex        =   65
                     Top             =   15
                     Width           =   60
                  End
                  Begin VB.TextBox XPTxtValue 
                     Alignment       =   1  'Right Justify
                     Height          =   75
                     Index           =   1
                     Left            =   390
                     MaxLength       =   10
                     RightToLeft     =   -1  'True
                     TabIndex        =   64
                     Top             =   15
                     Width           =   75
                  End
                  Begin VB.TextBox XPTxtSerial 
                     Alignment       =   1  'Right Justify
                     Height          =   75
                     Index           =   1
                     Left            =   285
                     Locked          =   -1  'True
                     RightToLeft     =   -1  'True
                     TabIndex        =   63
                     Top             =   15
                     Width           =   75
                  End
                  Begin VB.CheckBox ChkInstall 
                     Alignment       =   1  'Right Justify
                     BackColor       =   &H00E2E9E9&
                     Caption         =   " ﬁ”Ìÿ"
                     Height          =   90
                     Left            =   90
                     RightToLeft     =   -1  'True
                     TabIndex        =   62
                     Top             =   0
                     Width           =   60
                  End
                  Begin ImpulseButton.ISButton CmdINSTALLMENT 
                     Height          =   90
                     Left            =   0
                     TabIndex        =   66
                     Top             =   0
                     Width           =   90
                     _ExtentX        =   159
                     _ExtentY        =   159
                     ButtonStyle     =   1
                     ButtonPositionImage=   1
                     Caption         =   "Õ”«» «·√ﬁ”«ÿ"
                     BackColor       =   14871017
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
                     ButtonImage     =   "FrmInpoutWorkOrder.frx":7584
                     ColorButton     =   14871017
                     ColorHighlight  =   16777215
                     ColorHoverText  =   16711680
                     ColorShadow     =   4210752
                     ColorOutline    =   0
                     DrawFocusRectangle=   0   'False
                     ColorToggledHoverText=   16711680
                     ColorTextShadow =   4210752
                  End
                  Begin MSComCtl2.DTPicker DtpDelayDate 
                     Height          =   75
                     Left            =   150
                     TabIndex        =   67
                     Top             =   15
                     Width           =   90
                     _ExtentX        =   159
                     _ExtentY        =   132
                     _Version        =   393216
                     Format          =   201195521
                     CurrentDate     =   38784
                  End
                  Begin VB.Label lbl 
                     Alignment       =   1  'Right Justify
                     BackColor       =   &H00E2E9E9&
                     Caption         =   "„”·”·"
                     Height          =   75
                     Index           =   14
                     Left            =   360
                     RightToLeft     =   -1  'True
                     TabIndex        =   70
                     Top             =   15
                     Width           =   30
                  End
                  Begin VB.Label lbl 
                     Alignment       =   1  'Right Justify
                     BackColor       =   &H00E2E9E9&
                     Caption         =   "«·ﬁÌ„…"
                     Height          =   75
                     Index           =   15
                     Left            =   465
                     RightToLeft     =   -1  'True
                     TabIndex        =   69
                     Top             =   15
                     Width           =   15
                  End
                  Begin VB.Label lbl 
                     Alignment       =   1  'Right Justify
                     BackColor       =   &H00E2E9E9&
                     Caption         =   " «—ÌŒ «·«” Õﬁ«ﬁ"
                     Height          =   90
                     Index           =   21
                     Left            =   240
                     RightToLeft     =   -1  'True
                     TabIndex        =   68
                     Top             =   0
                     Width           =   45
                  End
               End
            End
            Begin C1SizerLibCtl.C1Elastic Ele 
               Height          =   240
               Index           =   11
               Left            =   13815
               TabIndex        =   71
               TabStop         =   0   'False
               Top             =   0
               Visible         =   0   'False
               Width           =   855
               _cx             =   1508
               _cy             =   423
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
               Begin MSDataListLib.DataCombo DcboCurrency 
                  Height          =   315
                  Left            =   75
                  TabIndex        =   80
                  Top             =   120
                  Visible         =   0   'False
                  Width           =   120
                  _ExtentX        =   212
                  _ExtentY        =   556
                  _Version        =   393216
                  Text            =   ""
                  RightToLeft     =   -1  'True
               End
               Begin VB.CheckBox XPChkPayType 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "‰ﬁœ«"
                  Height          =   345
                  Index           =   0
                  Left            =   765
                  RightToLeft     =   -1  'True
                  TabIndex        =   74
                  Top             =   90
                  Width           =   60
               End
               Begin VB.TextBox XPTxtSerial 
                  Alignment       =   1  'Right Justify
                  Height          =   375
                  Index           =   0
                  Left            =   450
                  Locked          =   -1  'True
                  RightToLeft     =   -1  'True
                  TabIndex        =   73
                  Top             =   60
                  Width           =   105
               End
               Begin VB.TextBox XPTxtValue 
                  Alignment       =   1  'Right Justify
                  Height          =   375
                  Index           =   0
                  Left            =   615
                  MaxLength       =   10
                  RightToLeft     =   -1  'True
                  TabIndex        =   72
                  Top             =   60
                  Width           =   90
               End
               Begin MSDataListLib.DataCombo DcboBox 
                  Height          =   315
                  Left            =   240
                  TabIndex        =   75
                  Top             =   105
                  Width           =   135
                  _ExtentX        =   238
                  _ExtentY        =   556
                  _Version        =   393216
                  Text            =   ""
                  RightToLeft     =   -1  'True
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "«”„ «·Œ“‰…"
                  Height          =   345
                  Index           =   2
                  Left            =   375
                  RightToLeft     =   -1  'True
                  TabIndex        =   79
                  Top             =   120
                  Width           =   75
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "„”·”·"
                  Height          =   345
                  Index           =   12
                  Left            =   555
                  RightToLeft     =   -1  'True
                  TabIndex        =   78
                  Top             =   90
                  Width           =   60
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "«·ﬁÌ„…"
                  Height          =   345
                  Index           =   13
                  Left            =   705
                  RightToLeft     =   -1  'True
                  TabIndex        =   77
                  Top             =   90
                  Width           =   45
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "«·⁄„·…"
                  Height          =   225
                  Index           =   20
                  Left            =   180
                  RightToLeft     =   -1  'True
                  TabIndex        =   76
                  Top             =   150
                  Visible         =   0   'False
                  Width           =   45
               End
            End
         End
         Begin C1SizerLibCtl.C1Elastic Ele 
            Height          =   4110
            Index           =   15
            Left            =   15705
            TabIndex        =   81
            TabStop         =   0   'False
            Top             =   45
            Width           =   14670
            _cx             =   25876
            _cy             =   7250
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial (Arabic)"
               Size            =   12
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
            ForeColor       =   -2147483630
            FloodColor      =   6553600
            ForeColorDisabled=   -2147483631
            Caption         =   ""
            Align           =   0
            AutoSizeChildren=   8
            BorderWidth     =   1
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
            Style           =   0
            TagSplit        =   2
            PicturePos      =   4
            CaptionStyle    =   0
            ResizeFonts     =   0   'False
            GridRows        =   7
            GridCols        =   1
            Frame           =   3
            FrameStyle      =   0
            FrameWidth      =   1
            FrameColor      =   -2147483628
            FrameShadow     =   -2147483632
            FloodStyle      =   1
            _GridInfo       =   $"FrmInpoutWorkOrder.frx":791E
            AccessibleName  =   ""
            AccessibleDescription=   ""
            AccessibleValue =   ""
            AccessibleRole  =   9
            Begin VB.TextBox TxtBillCommentx 
               Alignment       =   1  'Right Justify
               Height          =   1620
               Left            =   15
               MultiLine       =   -1  'True
               RightToLeft     =   -1  'True
               ScrollBars      =   2  'Vertical
               TabIndex        =   97
               Top             =   2475
               Width           =   14640
            End
            Begin C1SizerLibCtl.C1Elastic Ele 
               Height          =   1425
               Index           =   18
               Left            =   15
               TabIndex        =   82
               TabStop         =   0   'False
               Top             =   1035
               Width           =   14640
               _cx             =   25823
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
               Appearance      =   5
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
               Begin VB.CheckBox ChkTaxSerivce 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "÷—»Ì… Œœ„…"
                  Height          =   525
                  Left            =   435
                  RightToLeft     =   -1  'True
                  TabIndex        =   84
                  Top             =   285
                  Visible         =   0   'False
                  Width           =   45
               End
               Begin VB.TextBox TxtTaxServiceValue 
                  Alignment       =   1  'Right Justify
                  Height          =   810
                  Left            =   315
                  MaxLength       =   4
                  RightToLeft     =   -1  'True
                  TabIndex        =   83
                  Top             =   135
                  Visible         =   0   'False
                  Width           =   45
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Enabled         =   0   'False
                  Height          =   525
                  Index           =   49
                  Left            =   240
                  RightToLeft     =   -1  'True
                  TabIndex        =   109
                  Top             =   210
                  Visible         =   0   'False
                  Width           =   45
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "«·ﬁÌ„…"
                  Enabled         =   0   'False
                  Height          =   735
                  Index           =   43
                  Left            =   360
                  RightToLeft     =   -1  'True
                  TabIndex        =   86
                  Top             =   210
                  Visible         =   0   'False
                  Width           =   30
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "%"
                  BeginProperty Font 
                     Name            =   "Tahoma"
                     Size            =   9.75
                     Charset         =   178
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   600
                  Index           =   47
                  Left            =   300
                  RightToLeft     =   -1  'True
                  TabIndex        =   85
                  Top             =   210
                  Visible         =   0   'False
                  Width           =   15
               End
            End
            Begin C1SizerLibCtl.C1Elastic Ele 
               Height          =   1425
               Index           =   17
               Left            =   15
               TabIndex        =   87
               TabStop         =   0   'False
               Top             =   1035
               Width           =   14640
               _cx             =   25823
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
               Appearance      =   5
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
               Begin VB.CheckBox ChkTaxStamp 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "œ„€…"
                  Height          =   600
                  Left            =   435
                  RightToLeft     =   -1  'True
                  TabIndex        =   89
                  Top             =   210
                  Visible         =   0   'False
                  Width           =   45
               End
               Begin VB.TextBox TxtTaxStampValue 
                  Alignment       =   1  'Right Justify
                  Height          =   810
                  Left            =   315
                  MaxLength       =   4
                  RightToLeft     =   -1  'True
                  TabIndex        =   88
                  Top             =   135
                  Visible         =   0   'False
                  Width           =   45
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Enabled         =   0   'False
                  Height          =   525
                  Index           =   33
                  Left            =   255
                  RightToLeft     =   -1  'True
                  TabIndex        =   108
                  Top             =   210
                  Visible         =   0   'False
                  Width           =   45
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "«·ﬁÌ„…"
                  Enabled         =   0   'False
                  Height          =   660
                  Index           =   41
                  Left            =   360
                  RightToLeft     =   -1  'True
                  TabIndex        =   91
                  Top             =   210
                  Visible         =   0   'False
                  Width           =   30
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "$"
                  BeginProperty Font 
                     Name            =   "Tahoma"
                     Size            =   9.75
                     Charset         =   178
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   600
                  Index           =   48
                  Left            =   300
                  RightToLeft     =   -1  'True
                  TabIndex        =   90
                  Top             =   210
                  Visible         =   0   'False
                  Width           =   15
               End
            End
            Begin C1SizerLibCtl.C1Elastic Ele 
               Height          =   1005
               Index           =   16
               Left            =   15
               TabIndex        =   92
               TabStop         =   0   'False
               Top             =   15
               Width           =   14640
               _cx             =   25823
               _cy             =   1773
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
               Appearance      =   5
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
               Begin VB.CheckBox ChkTaxAdd 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "÷—»Ì»… Œ’„ Ê≈÷«›… (√—»«Õ  Ã«—Ì…)"
                  Height          =   990
                  Left            =   135
                  RightToLeft     =   -1  'True
                  TabIndex        =   94
                  Top             =   75
                  Width           =   30
               End
               Begin VB.TextBox TxtTaxAddValue 
                  Alignment       =   1  'Right Justify
                  Height          =   795
                  Left            =   105
                  MaxLength       =   4
                  RightToLeft     =   -1  'True
                  TabIndex        =   93
                  Top             =   135
                  Width           =   15
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Enabled         =   0   'False
                  Height          =   585
                  Index           =   32
                  Left            =   75
                  RightToLeft     =   -1  'True
                  TabIndex        =   107
                  Top             =   210
                  Width           =   15
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "«·ﬁÌ„…"
                  Enabled         =   0   'False
                  Height          =   660
                  Index           =   39
                  Left            =   120
                  RightToLeft     =   -1  'True
                  TabIndex        =   96
                  Top             =   210
                  Width           =   15
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "%"
                  BeginProperty Font 
                     Name            =   "Tahoma"
                     Size            =   9.75
                     Charset         =   178
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   585
                  Index           =   46
                  Left            =   90
                  RightToLeft     =   -1  'True
                  TabIndex        =   95
                  Top             =   210
                  Width           =   15
               End
            End
            Begin C1SizerLibCtl.C1Elastic Ele 
               Height          =   1005
               Index           =   8
               Left            =   15
               TabIndex        =   98
               TabStop         =   0   'False
               Top             =   15
               Width           =   14640
               _cx             =   25823
               _cy             =   1773
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
               Appearance      =   5
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
               Begin VB.TextBox XPTxtTaxValue 
                  Alignment       =   1  'Right Justify
                  Height          =   780
                  Left            =   105
                  MaxLength       =   4
                  RightToLeft     =   -1  'True
                  TabIndex        =   100
                  Top             =   180
                  Width           =   15
               End
               Begin VB.CheckBox XPChkTAX 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "÷—»Ì»… «·„»Ì⁄« "
                  Height          =   525
                  Left            =   135
                  RightToLeft     =   -1  'True
                  TabIndex        =   99
                  Top             =   270
                  Width           =   30
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Enabled         =   0   'False
                  Height          =   525
                  Index           =   25
                  Left            =   75
                  RightToLeft     =   -1  'True
                  TabIndex        =   106
                  Top             =   270
                  Width           =   15
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "«·ﬁÌ„…"
                  Enabled         =   0   'False
                  Height          =   510
                  Index           =   22
                  Left            =   120
                  RightToLeft     =   -1  'True
                  TabIndex        =   102
                  Top             =   315
                  Width           =   15
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "%"
                  BeginProperty Font 
                     Name            =   "Tahoma"
                     Size            =   9.75
                     Charset         =   178
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   600
                  Index           =   45
                  Left            =   90
                  RightToLeft     =   -1  'True
                  TabIndex        =   101
                  Top             =   270
                  Width           =   15
               End
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "≈÷«›… √Ì… „·«ÕŸ«  ⁄·Ï «·›« Ê—…"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   12
                  Charset         =   178
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H000000C0&
               Height          =   1620
               Index           =   44
               Left            =   15
               RightToLeft     =   -1  'True
               TabIndex        =   103
               Top             =   2475
               Width           =   14640
            End
         End
         Begin C1SizerLibCtl.C1Elastic Ele 
            Height          =   4110
            Index           =   9
            Left            =   16005
            TabIndex        =   177
            TabStop         =   0   'False
            Top             =   45
            Width           =   14670
            _cx             =   25876
            _cy             =   7250
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
            Begin VB.TextBox TXTFactoryExpenses 
               Alignment       =   2  'Center
               Height          =   405
               Left            =   6720
               RightToLeft     =   -1  'True
               TabIndex        =   178
               Top             =   2880
               Width           =   1215
            End
            Begin VSFlex8Ctl.VSFlexGrid Fg_Journal 
               Height          =   2340
               Left            =   600
               TabIndex        =   179
               Top             =   480
               Width           =   12600
               _cx             =   22225
               _cy             =   4128
               Appearance      =   1
               BorderStyle     =   1
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
               Rows            =   1
               Cols            =   9
               FixedRows       =   1
               FixedCols       =   0
               RowHeightMin    =   300
               RowHeightMax    =   0
               ColWidthMin     =   0
               ColWidthMax     =   0
               ExtendLastCol   =   0   'False
               FormatString    =   $"FrmInpoutWorkOrder.frx":7995
               ScrollTrack     =   0   'False
               ScrollBars      =   2
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
               Begin VB.PictureBox PicDes 
                  BorderStyle     =   0  'None
                  Height          =   1635
                  Left            =   240
                  RightToLeft     =   -1  'True
                  ScaleHeight     =   1635
                  ScaleWidth      =   2925
                  TabIndex        =   180
                  Top             =   960
                  Visible         =   0   'False
                  Width           =   2925
                  Begin VB.TextBox TxtDes 
                     Alignment       =   1  'Right Justify
                     BackColor       =   &H80000018&
                     BorderStyle     =   0  'None
                     Height          =   1125
                     Left            =   30
                     MultiLine       =   -1  'True
                     RightToLeft     =   -1  'True
                     ScrollBars      =   3  'Both
                     TabIndex        =   181
                     Top             =   360
                     Visible         =   0   'False
                     Width           =   2115
                  End
                  Begin VB.Label LblDes 
                     Alignment       =   1  'Right Justify
                     BackColor       =   &H8000000C&
                     Caption         =   "Ì„ﬂ‰ﬂ ﬂ «»…  ⁄·Ìﬁ Â‰«:"
                     ForeColor       =   &H0000C8FF&
                     Height          =   315
                     Left            =   0
                     RightToLeft     =   -1  'True
                     TabIndex        =   182
                     Top             =   0
                     Width           =   2445
                  End
               End
               Begin VDSCOMBOLibCtl.SmartCombo CboDes 
                  Height          =   315
                  Left            =   240
                  TabIndex        =   183
                  ToolTipText     =   "ﬂ «»…  ⁄·Ìﬁ"
                  Top             =   600
                  Visible         =   0   'False
                  Width           =   2955
                  _cx             =   1973752924
                  _cy             =   1973748268
                  Alignment       =   0
                  Appearance      =   3
                  AutoSearch      =   0   'False
                  BackColor       =   -2147483624
                  BackgroundColor =   -2147483633
                  BorderColor     =   0
                  BorderVisible   =   -1  'True
                  Caption         =   "SmartCombo1"
                  CaptionAlignment=   4
                  CaptionBackColor=   -2147483633
                  BeginProperty CaptionFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "MS Sans Serif"
                     Size            =   8.25
                     Charset         =   178
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  CaptionForeColor=   -2147483630
                  CaptionHeight   =   15
                  CaptionOnTop    =   0   'False
                  CaptionMultiLine=   0
                  Checkbox3D      =   0   'False
                  CheckboxAlignment=   5
                  CheckboxBackColor=   16777215
                  CheckboxSize    =   13
                  CheckboxValue   =   0
                  BrowsePictureAlignment=   5
                  BrowsePictureStretchH=   0
                  BrowsePictureStretchV=   0
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
                  ForeColor       =   0
                  Gap             =   0
                  HideSelection   =   -1  'True
                  Locked          =   0   'False
                  MaxLength       =   0
                  MultiLine       =   0
                  OnFocus         =   3
                  PasswordChar    =   ""
                  Picture         =   "FrmInpoutWorkOrder.frx":7AF5
                  PictureAlignment=   5
                  PictureBackColor=   -2147483624
                  PictureStretchH =   0
                  PictureStretchV =   0
                  Redraw          =   -1  'True
                  ScrollBar       =   0
                  Style           =   0
                  Text            =   ""
                  UnderLine       =   0   'False
                  Enabled0        =   -1  'True
                  Position0       =   0
                  Tip0            =   "Caption"
                  Visible0        =   0   'False
                  Width0          =   90
                  Enabled1        =   -1  'True
                  Position1       =   1
                  Tip1            =   ""
                  Visible1        =   -1  'True
                  Width1          =   32
                  Enabled2        =   -1  'True
                  Position2       =   2
                  Tip2            =   "Check Box (Space, Ctrl + Space)"
                  Visible2        =   0   'False
                  Width2          =   16
                  Enabled3        =   -1  'True
                  Position3       =   3
                  Tip3            =   "ﬂ «»…  ⁄·Ìﬁ"
                  Visible3        =   -1  'True
                  Width3          =   145
                  Enabled4        =   -1  'True
                  Position4       =   4
                  Tip4            =   "Left Spinner (Alt + Left)"
                  Visible4        =   0   'False
                  Width4          =   16
                  Enabled5        =   -1  'True
                  Position5       =   5
                  Tip5            =   "Right Spinner (Alt + Right)"
                  Visible5        =   0   'False
                  Width5          =   16
                  Enabled6        =   -1  'True
                  Position6       =   6
                  Tip6            =   "Up Spinner (Ctrl + Up)"
                  Visible6        =   0   'False
                  Width6          =   16
                  Enabled7        =   -1  'True
                  Position7       =   7
                  Tip7            =   "Down Spinner (Ctrl + Down)"
                  Visible7        =   0   'False
                  Width7          =   16
                  Enabled8        =   -1  'True
                  Position8       =   8
                  Tip8            =   "Browse (Alt + Enter)"
                  Visible8        =   0   'False
                  Width8          =   16
                  Enabled9        =   -1  'True
                  Position9       =   9
                  Tip9            =   " (Alt + Down, F4)"
                  Visible9        =   -1  'True
                  Width9          =   16
                  Enabled10       =   -1  'True
                  Position10      =   10
                  Tip10           =   "Right Arrow (Alt + >)"
                  Visible10       =   0   'False
                  Width10         =   16
               End
            End
            Begin ImpulseButton.ISButton Cmd 
               Height          =   390
               Index           =   9
               Left            =   11640
               TabIndex        =   184
               Top             =   2880
               Width           =   1530
               _ExtentX        =   2699
               _ExtentY        =   688
               ButtonStyle     =   1
               ButtonPositionImage=   1
               Caption         =   "Õ–› ”ÿ—"
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
               ButtonImage     =   "FrmInpoutWorkOrder.frx":808F
               ColorButton     =   14871017
               DrawFocusRectangle=   0   'False
            End
            Begin VB.Label Label19 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "«Œ Ì«— «·„’—Ê›«  «· ﬁœÌ—ÌÂ"
               Height          =   255
               Left            =   10920
               RightToLeft     =   -1  'True
               TabIndex        =   186
               Top             =   240
               Width           =   2415
            End
            Begin VB.Label Label18 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "«Ã„«·Ì  «·„’«—Ì› «· ﬁœÌ—ÌÂ"
               Height          =   375
               Left            =   7920
               RightToLeft     =   -1  'True
               TabIndex        =   185
               Top             =   3000
               Width           =   2055
            End
         End
      End
      Begin C1SizerLibCtl.C1Elastic Ele 
         Height          =   540
         Index           =   1
         Left            =   15
         TabIndex        =   129
         TabStop         =   0   'False
         Top             =   7380
         Width           =   14760
         _cx             =   26035
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
         Begin ImpulseButton.ISButton Cmd 
            Height          =   540
            Index           =   0
            Left            =   13200
            TabIndex        =   130
            Top             =   0
            Width           =   1470
            _ExtentX        =   2593
            _ExtentY        =   953
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
            Height          =   540
            Index           =   1
            Left            =   11550
            TabIndex        =   131
            Top             =   0
            Width           =   1470
            _ExtentX        =   2593
            _ExtentY        =   953
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
            Height          =   540
            Index           =   2
            Left            =   9870
            TabIndex        =   132
            Top             =   0
            Width           =   1500
            _ExtentX        =   2646
            _ExtentY        =   953
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
            Height          =   540
            Index           =   3
            Left            =   8250
            TabIndex        =   133
            Top             =   0
            Width           =   1515
            _ExtentX        =   2672
            _ExtentY        =   953
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
            Height          =   540
            Index           =   4
            Left            =   6600
            TabIndex        =   134
            Top             =   0
            Width           =   1485
            _ExtentX        =   2619
            _ExtentY        =   953
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
            Height          =   540
            Index           =   5
            Left            =   4920
            TabIndex        =   135
            Top             =   0
            Width           =   1500
            _ExtentX        =   2646
            _ExtentY        =   953
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
            Height          =   540
            Index           =   6
            Left            =   30
            TabIndex        =   136
            Top             =   0
            Width           =   1440
            _ExtentX        =   2540
            _ExtentY        =   953
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
            Height          =   540
            Index           =   7
            Left            =   3225
            TabIndex        =   137
            Top             =   0
            Width           =   1515
            _ExtentX        =   2672
            _ExtentY        =   953
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
            Height          =   540
            Left            =   1605
            TabIndex        =   138
            Top             =   0
            Width           =   1515
            _ExtentX        =   2672
            _ExtentY        =   953
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
      End
      Begin C1SizerLibCtl.C1Elastic Ele 
         Height          =   435
         Index           =   3
         Left            =   15
         TabIndex        =   139
         TabStop         =   0   'False
         Top             =   6930
         Width           =   14760
         _cx             =   26035
         _cy             =   767
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
            Alignment       =   2  'Center
            BackColor       =   &H000000FF&
            Height          =   435
            Left            =   7635
            Locked          =   -1  'True
            RightToLeft     =   -1  'True
            TabIndex        =   140
            TabStop         =   0   'False
            Top             =   0
            Visible         =   0   'False
            Width           =   330
         End
         Begin MSDataListLib.DataCombo DCboUserName 
            Height          =   315
            Left            =   3270
            TabIndex        =   141
            Top             =   90
            Width           =   1620
            _ExtentX        =   2858
            _ExtentY        =   556
            _Version        =   393216
            Enabled         =   0   'False
            Text            =   ""
            RightToLeft     =   -1  'True
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "«·≈Ã„«·Ï"
            Height          =   255
            Index           =   3
            Left            =   14130
            RightToLeft     =   -1  'True
            TabIndex        =   154
            Top             =   90
            Width           =   540
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "—ﬁ„ «·”Ã·"
            Height          =   255
            Index           =   0
            Left            =   2100
            RightToLeft     =   -1  'True
            TabIndex        =   153
            Top             =   90
            Width           =   795
         End
         Begin VB.Label XPTxtCurrent 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Height          =   255
            Left            =   1350
            RightToLeft     =   -1  'True
            TabIndex        =   152
            Top             =   90
            Width           =   720
         End
         Begin VB.Label XPTxtCount 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Height          =   255
            Left            =   45
            RightToLeft     =   -1  'True
            TabIndex        =   151
            Top             =   90
            Width           =   750
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "«·„” Œœ„"
            Height          =   255
            Index           =   1
            Left            =   4815
            RightToLeft     =   -1  'True
            TabIndex        =   150
            Top             =   90
            Width           =   750
         End
         Begin VB.Label LblTotal 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00FF0000&
            BorderStyle     =   1  'Fixed Single
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   12
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   435
            Left            =   8370
            RightToLeft     =   -1  'True
            TabIndex        =   149
            Top             =   0
            Visible         =   0   'False
            Width           =   1545
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "/"
            Height          =   255
            Index           =   23
            Left            =   825
            RightToLeft     =   -1  'True
            TabIndex        =   148
            Top             =   90
            Width           =   225
         End
         Begin VB.Label LblTotalAll 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00FF0000&
            BorderStyle     =   1  'Fixed Single
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   12
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   435
            Left            =   12705
            RightToLeft     =   -1  'True
            TabIndex        =   147
            Top             =   0
            Width           =   1320
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "Œ’Ê„« "
            Height          =   255
            Index           =   50
            Left            =   12045
            RightToLeft     =   -1  'True
            TabIndex        =   146
            Top             =   90
            Visible         =   0   'False
            Width           =   645
         End
         Begin VB.Label LblDiscountsTotal 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00FF0000&
            BorderStyle     =   1  'Fixed Single
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   12
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   435
            Left            =   10725
            RightToLeft     =   -1  'True
            TabIndex        =   145
            Top             =   0
            Visible         =   0   'False
            Width           =   1215
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "«·’«›Ì"
            Height          =   255
            Index           =   24
            Left            =   9975
            RightToLeft     =   -1  'True
            TabIndex        =   144
            Top             =   90
            Visible         =   0   'False
            Width           =   660
         End
         Begin VB.Label LblTotalQty 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00FF0000&
            BorderStyle     =   1  'Fixed Single
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   12
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   345
            Left            =   5625
            TabIndex        =   143
            Top             =   0
            Visible         =   0   'False
            Width           =   960
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "«Ã„«·Ì «·ﬂ„ÌÂ"
            Height          =   255
            Index           =   63
            Left            =   6780
            TabIndex        =   142
            Top             =   180
            Visible         =   0   'False
            Width           =   1080
         End
      End
   End
End
Attribute VB_Name = "FrmInpoutWorkOrder"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rs As ADODB.Recordset
Dim NewGrid As ClsGrid
Dim TTP As clstooltip
Dim BuyReport As ClsBuyReport
Dim cSearchDcbo(3) As clsDCboSearch
Public BolPrint As Boolean
Dim WithEvents m_MnuShowNewItemsPrices As Menu
Attribute m_MnuShowNewItemsPrices.VB_VarHelpID = -1
Dim WithEvents m_MenuViewList As Menu
Attribute m_MenuViewList.VB_VarHelpID = -1
Dim WithEvents m_MenuShowItemCostEffect As Menu
Attribute m_MenuShowItemCostEffect.VB_VarHelpID = -1
Dim WithEvents m_FrmSearch As Form
Attribute m_FrmSearch.VB_VarHelpID = -1
Dim general_noteid As Long
Dim RsNotesGeneral As ADODB.Recordset
Dim IsSaveWithOutMsg As Boolean
Function CuurentLogdata(Optional Currentmode As String)
    LogTextA = "    ‘«‘… " & ScreenNameArabic & CHR(13) & " —ﬁ„ «·«–‰   " & TxtNoteSerial1.text & CHR(13) & "  «—ÌŒ «·«–‰ " & XPDtbBill.value & CHR(13) & " «·›—⁄   " & dcBranch.text & CHR(13) & "—ﬁ„ «„— «·«‰ «Ã  " & TXTOrderNO1 & CHR(13) & " «·„Œ“‰  " & DCboStoreName.text
                     
    LogTexte = "    Screen  " & ScreenNameEnglish & CHR(13) & " Vchr. No.  " & TxtNoteSerial1.text & CHR(13) & "Date " & XPDtbBill.value & CHR(13) & " Branch   " & dcBranch.text & CHR(13) & " To  Order No " & TXTOrderNO1 & CHR(13) & " Inventory  " & DCboStoreName.text
                     
    If Currentmode <> "D" Then
        AddToLogFile CInt(user_id), 250, Date, Time, LogTextA, LogTexte, Me.Name, Me.TxtModFlg, , , Me.TxtNoteSerial, Me.TxtNoteSerial1
    Else
        AddToLogFile CInt(user_id), 250, Date, Time, LogTextA, LogTexte, Me.Name, "D", , , Me.TxtNoteSerial, Me.TxtNoteSerial1
    End If
    
End Function

Public Sub Convert()
    Cmd_Click (0)
End Sub

Public Sub Cala()
    NewGrid.Calculate 1, , , True
End Sub

Private Sub CBoBasedON_Change()

    If Me.CBoBasedON.ListIndex = 0 Then

    ElseIf Me.CBoBasedON.ListIndex = 1 Then
    If SystemOptions.UserInterface = ArabicInterface Then
        lbl(55).Caption = "—ﬁ„ «·«„—"
     Else
     lbl(55).Caption = "Order No."
     End If
    ElseIf Me.CBoBasedON.ListIndex = 2 Then
    If SystemOptions.UserInterface = ArabicInterface Then
        lbl(55).Caption = "—ﬁ„ «·ﬁ« Ê—… «·„»œ∆ÌÂ"
    Else
    lbl(55).Caption = "Bill No."
    End If
    End If

End Sub

Private Sub CBoBasedON_Click()
    CBoBasedON_Change
End Sub

Private Sub ChkInstall_Click()

    If ChkInstall.value = vbChecked Then
        Me.CmdINSTALLMENT.Enabled = True
    Else
        Me.CmdINSTALLMENT.Enabled = False

        With Me.FgInstallments
            .Clear flexClearScrollable, flexClearEverything
            .rows = .FixedRows
            LblPrecenType.Caption = ""
            LblPrecenValue.Caption = ""
            LblInstallTotal.Caption = ""
            LblInstallCount.Caption = ""
            LblFirstInstallDate.Caption = ""
            LblInstallmentType.Caption = ""
        End With

    End If

End Sub

Private Sub ChkTaxAdd_Click()

    If ChkTaxAdd.value = Checked Then
        TxtTaxAddValue.Enabled = True
        lbl(39).Enabled = True
        lbl(46).Enabled = True
    Else
        TxtTaxAddValue.text = ""
        TxtTaxAddValue.Enabled = False
        lbl(39).Enabled = False
        lbl(46).Enabled = False
    End If

    Exit Sub
ErrTrap:
End Sub

Private Sub ChkTaxSerivce_Click()
    On Error GoTo ErrTrap

    If ChkTaxSerivce.value = Checked Then
        TxtTaxServiceValue.Enabled = True
        lbl(43).Enabled = True
        lbl(47).Enabled = True
    Else
        TxtTaxServiceValue.text = ""
        TxtTaxServiceValue.Enabled = False
        lbl(43).Enabled = False
        lbl(47).Enabled = False
    End If

    Exit Sub
ErrTrap:
End Sub

Private Sub ChkTaxStamp_Click()

    If ChkTaxStamp.value = Checked Then
        TxtTaxStampValue.Enabled = True
        lbl(41).Enabled = True
        lbl(48).Enabled = True
    Else
        TxtTaxStampValue.text = ""
        TxtTaxStampValue.Enabled = False
        lbl(41).Enabled = False
        lbl(48).Enabled = False
    End If

    Exit Sub
ErrTrap:
End Sub

Private Sub Cmd_Click(Index As Integer)
  On Error GoTo ErrTrap
    Dim AskOption As Boolean
    Dim intDef As Integer
    Dim Msg As String

    BolPrint = True

    Select Case Index

        Case 0

            If DoPremis(Do_New, Me.Name, True) = False Then
                Exit Sub
            End If

            clear_all Me
            TxtModFlg.text = "N"
            SetDefaults
            NewGrid.GridDefaultValue 1
            dcBranch.BoundText = Current_branch
            Me.DCboUserName.BoundText = user_id
            intDef = GetSetting(StrAppRegPath, "DefaultOptions", "DefaultSup", 1)
            DBCboClientName.BoundText = intDef
        
            intDef = GetSetting(StrAppRegPath, "DefaultOptions", "DefaultPurchaseStore", 1)
            DCboStoreName.BoundText = intDef
        
            '   XPTab301.CurrTab = 0
            FG.SetFocus
            FG.Col = FG.ColIndex("Code")
            FG.Row = FG.rows - 1
            Fg_Journal.Clear flexClearScrollable, flexClearEverything
            Fg_Journal.rows = 2
            Fg_Journal.Enabled = True
            Me.CBoBasedON.ListIndex = 0
            CboPriceType.ListIndex = 0
            CBOSource.ListIndex = 0

            If Voucher_coding(val(my_branch), XPDtbBill.value, 19, 250) = "" And val(my_branch) <> 0 Then
                TxtNoteSerial1.locked = False
            Else
                TxtNoteSerial1.locked = True
 
            End If
 
        Case 1
                    If ChekClodePeriod(XPDtbBill.value) = True Then
               If SystemOptions.UserInterface = ArabicInterface Then
                MsgBox "Ì—ÃÏ  €ÌÌ—  «—ÌŒ «·Õ—ﬂ… ·«‰ Â–Â «·› —… „€·ﬁ…"
               Else
               MsgBox "Please Change Date Becouse This is Period is Closed"
              End If
              Exit Sub
              End If
            If cmdReSave.Visible = False Then
            If Text1.text <> "" And SystemOptions.AllowEditProductionOutManulay = False Then
            If SystemOptions.UserInterface = ArabicInterface Then
                Msg = " „  ÕÊÌ· Â–« «·«–‰  »√„— «‰ «Ã   ›·« Ì„ﬂ‰ﬂ «· ⁄œÌ· " & Space$(5) & TXTOrderNO1.text
              Else
                   Msg = "Can Not Edit " & Space$(5) & TXTOrderNO1.text
         
            End If
                MsgBox Msg, vbOKOnly + vbMsgBoxRight + vbExclamation + vbMsgBoxRtlReading, App.Title
                Exit Sub
            End If
        
            If DoPremis(Do_Edit, Me.Name, True) = False Then
                Exit Sub
            End If
        End If
     '       If SystemOptions.usertype = UserNormal Then
     '           If AvailableDeal = False Then
     '               Exit Sub
     '           End If
     '       End If

            TxtModFlg.text = "E"
            If Trim(txtPassword) <> "Alex2025" Then
            Me.DCboUserName.BoundText = user_id
            End If
            Me.DcboBox.BoundText = 1
            Fg_Journal.rows = Fg_Journal.rows + 1
            Fg_Journal.Enabled = True
            CuurentLogdata

        Case 2
        
              If ChekClodePeriod(XPDtbBill.value) = True Then
               If SystemOptions.UserInterface = ArabicInterface Then
                MsgBox "Ì—ÃÏ  €ÌÌ—  «—ÌŒ «·Õ—ﬂ… ·«‰ Â–Â «·› —… „€·ﬁ…"
               Else
               MsgBox "Please Change Date Becouse This is Period is Closed"
              End If
              Exit Sub
              End If
              
            '                   If Me.TxtModFlg.text = "N" Then
             
            'If SystemOptions.UserType = UserAdminAll Then
            If Trim(dcBranch.BoundText) = "" Then
                If SystemOptions.UserInterface = EnglishInterface Then
                    Msg = "Specify Departement"
                Else
                    Msg = "Õœœ «·›—⁄ «Ê·«  "
                End If

                MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
                dcBranch.SetFocus
                Sendkeys "{F4}"
                Screen.MousePointer = vbDefault
                Exit Sub
            End If

            my_branch = Me.dcBranch.BoundText
  
            'End If

            '                       If TxtNoteSerial.text = "" Then
             
            '             End If
             
            SaveData
            CuurentLogdata

        Case 3
            Undo

        Case 4
                If ChekClodePeriod(XPDtbBill.value) = True Then
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

            If Text1.text <> "" Then
                Msg = " „  ÕÊÌ· Â–« «·«–‰  »√„— «‰ «Ã   ›·« Ì„ﬂ‰ﬂ «·Õ–› " & Space$(5) & TXTOrderNO1.text
                MsgBox Msg, vbOKOnly + vbMsgBoxRight + vbExclamation + vbMsgBoxRtlReading, App.Title
                Exit Sub
            End If

            If SystemOptions.usertype = UserNormal Then
                Msg = "·Ì” ·ﬂ Õﬁ Õ–› ›Ï «·›Ê« Ì—"
                MsgBox Msg, vbOKOnly + vbMsgBoxRight + vbExclamation + vbMsgBoxRtlReading, App.Title
                Exit Sub
            End If

            Del_TransAction

        Case 5

            If DoPremis(Do_Search, Me.Name, True) = False Then
                Exit Sub
            End If

            If m_FrmSearch Is Nothing Then
                Set m_FrmSearch = New FrmBuySearch
                m_FrmSearch.DealingForm = ProductionMaterialReciveVoucher
                m_FrmSearch.Caption = "«·»ÕÀ ⁄‰ ”‰œ«  «” ·«„  „—«Õ· «‰ «Ã  «„  "
                Set m_FrmSearch.RetrunFrm = Me
                m_FrmSearch.show vbModal ', mdifrmmain
            Else
                Msg = "Â‰«ﬂ ‘«‘… »ÕÀ "
                Msg = Msg & CHR(13) & "Ÿ«Â—… «„«„ﬂ ›⁄·«...·«Ì„ﬂ‰ ⁄—÷ «ﬂÀ— „‰ ‘«‘… »ÕÀ  "
                MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
                m_FrmSearch.Visible = True
                m_FrmSearch.ZOrder 0
                m_FrmSearch.SetFocus
            End If

        Case 6
            Unload Me

        Case 7

            If DoPremis(Do_Print, Me.Name, True) = False Then
                Exit Sub
            End If

            AskOption = GetSetting(StrAppRegPath, "View_Type", "ShowMe", False)

            If AskOption = False Then
                FrmPrintOptions.show vbModal
            End If

            If BolPrint = False Then
                Exit Sub
            End If

            printing
        
        Case 10
'            ShowGL_cc TxtNoteSerial.text, , 200, val(Me.TxtNoteID.text)
            ShowGL_cc TxtNoteSerial.text, , 200, val(Me.TXTNoteID.text)
    End Select

    Exit Sub
ErrTrap:
End Sub

Private Sub SumChecks()

    With Me.FgCheques

        If .rows > 1 Then
            Me.lbl(19).Caption = .Aggregate(flexSTCount, .FixedRows, .ColIndex("CheckNumber"), .rows - 1, .ColIndex("CheckNumber"))
            Me.lbl(18).Caption = .Aggregate(flexSTSum, .FixedRows, .ColIndex("CheckValue"), .rows - 1, .ColIndex("CheckValue"))
        Else
            Me.lbl(19).Caption = 0
            Me.lbl(18).Caption = 0
        End If

    End With

End Sub

Private Sub CmdConvert_Click()
    Dim RowNum As Integer
    Dim Frm As Form
    Dim Msg As String

    If Text1.text <> "" Then
        Msg = " „  ÕÊÌ· Â–« «·«–‰ »—ﬁ„ ›« Ê—…  " & Space$(5) & Text1.text
        MsgBox Msg, vbOKOnly + vbMsgBoxRight + vbExclamation + vbMsgBoxRtlReading, App.Title
        Exit Sub
    End If

    On Error GoTo ErrTrap
    Screen.MousePointer = vbArrowHourglass
    Set Frm = New FrmBillBuy

    With Frm
        .Convert
        '    .XPTxtBillID.Text = XPTxtBillID.Text
        .XPDtbBill.value = XPDtbBill.value
        .DBCboClientName.BoundText = DBCboClientName.BoundText
        .DCboStoreName.BoundText = DCboStoreName.BoundText
        .CboPayMentType.ListIndex = 0 ' CboPaymentType.ListIndex
        .Text1.text = TxtTransSerial.text
        .Text2.text = XPTxtBillID.text
    
        For RowNum = 1 To FG.rows - 1

            If .FG.TextMatrix(.FG.rows - 1, .FG.ColIndex("Code")) <> "" Then
                .FG.rows = .FG.rows + 1
            End If

            .FG.TextMatrix(.FG.rows - 1, .FG.ColIndex("Name")) = IIf(FG.TextMatrix(RowNum, FG.ColIndex("Name")) = "", "", FG.TextMatrix(RowNum, FG.ColIndex("Name")))
            .FG.TextMatrix(.FG.rows - 1, .FG.ColIndex("Code")) = IIf(FG.TextMatrix(RowNum, FG.ColIndex("Code")) = "", "", FG.TextMatrix(RowNum, FG.ColIndex("Code")))
            .FG.TextMatrix(.FG.rows - 1, .FG.ColIndex("ItemCase")) = IIf(FG.TextMatrix(RowNum, FG.ColIndex("ItemCase")) = "", "", FG.TextMatrix(RowNum, FG.ColIndex("ItemCase")))
            .FG.TextMatrix(.FG.rows - 1, .FG.ColIndex("HaveSerial")) = IIf(FG.TextMatrix(RowNum, FG.ColIndex("HaveSerial")) = "", "", FG.TextMatrix(RowNum, FG.ColIndex("HaveSerial")))
            .FG.TextMatrix(.FG.rows - 1, .FG.ColIndex("Count")) = IIf(FG.TextMatrix(RowNum, FG.ColIndex("Count")) = "", "", FG.TextMatrix(RowNum, FG.ColIndex("Count")))
            .FG.TextMatrix(.FG.rows - 1, .FG.ColIndex("Price")) = IIf(FG.TextMatrix(RowNum, FG.ColIndex("Price")) = "", "", FG.TextMatrix(RowNum, FG.ColIndex("Price")))
            .FG.TextMatrix(.FG.rows - 1, .FG.ColIndex("DiscountType")) = IIf(FG.TextMatrix(RowNum, FG.ColIndex("DiscountType")) = "", "", FG.TextMatrix(RowNum, FG.ColIndex("DiscountType")))
            Dim StrSQL As String
            Dim RsUnit As New ADODB.Recordset
            StrSQL = "SELECT TOP 100 PERCENT dbo.TblItemsUnits.UnitID, dbo.TblUnites.UnitName, dbo.Transactions.Transaction_Serial,dbo.Transactions.Transaction_Type FROM dbo.Transaction_Details INNER JOIN dbo.Transactions ON dbo.Transaction_Details.Transaction_ID = dbo.Transactions.Transaction_ID INNER JOIN dbo.TblUnites INNER JOIN dbo.TblItemsUnits ON dbo.TblUnites.UnitID = dbo.TblItemsUnits.UnitID ON dbo.Transaction_Details.UnitId = dbo.TblUnites.UnitID AND dbo.Transaction_Details.Item_ID = dbo.TblItemsUnits.ItemID WHERE (dbo.Transactions.Transaction_Serial = '" & TxtTransSerial & "') AND (dbo.Transactions.Transaction_Type = 28) AND (dbo.TblItemsUnits.ItemID = " & FG.TextMatrix(RowNum, FG.ColIndex("Code")) & ") ORDER BY dbo.TblItemsUnits.SecOrder"
            Set RsUnit = New ADODB.Recordset
            RsUnit.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText
        
            .FG.cell(flexcpData, RowNum, .FG.ColIndex("UnitID")) = IIf(IsNull(RsUnit("UnitID")), "", (RsUnit("UnitID").value))
            .FG.TextMatrix(RowNum, .FG.ColIndex("UnitID")) = IIf(IsNull(RsUnit("UnitName")), "", (RsUnit("UnitName").value))

            '        FG.Cell(flexcpData, I, FG.ColIndex("UnitID")) = IIf(IsNull(RsDetails("UnitID")), "", (RsDetails("UnitID").Value))
            '        FG.TextMatrix(I, FG.ColIndex("UnitID")) = IIf(IsNull(RsDetails("UnitName")), "", (RsDetails("UnitName").Value))
            '           StrSQL = "SELECT dbo.Transactions.Transaction_Type, dbo.Transaction_Details.UnitId, dbo.TblUnites.UnitName, dbo.Transactions.Transaction_Serial FROM dbo.Transaction_Details INNER JOIN dbo.Transactions ON dbo.Transaction_Details.Transaction_ID = dbo.Transactions.Transaction_ID INNER JOIN dbo.TblUnites ON dbo.Transaction_Details.UnitId = dbo.TblUnites.UnitID WHERE (dbo.Transactions.Transaction_Type = 19) AND (dbo.Transactions.Transaction_Serial = '" & TxtTransSerial & "')"
            '        .FG.Cell(flexcpData, .FG.Rows - 1, FG.ColIndex("UnitID")) = 1 'FG.Cell(flexcpData, RowNum, FG.ColIndex("UnitID")) ' IIf(IsNull(RsUnit("UnitID")), "", (RsUnit("UnitID").Value))
            '        .FG.TextMatrix(.FG.Rows - 1, FG.ColIndex("UnitID")) = "Ã—«„" 'FG.TextMatrix(RowNum, FG.ColIndex("UnitID")) ' IIf(IsNull(RsUnit("UnitName")), "", (RsUnit("UnitName").Value))

        Next RowNum

        .Cala
    End With

    Screen.MousePointer = vbDefault
    Exit Sub
ErrTrap:
    Screen.MousePointer = vbDefault

End Sub

Private Sub CmdHelp_Click()
    SystemOptions.SysHelp.HHTopicID = Me.HelpContextID
    SystemOptions.SysHelp.HHDisplayTopicID Me.hWnd
End Sub

Private Sub CmdInfo_Click()
    Me.PopupMenu mdifrmmain.MnuInvPurchase
End Sub

Private Sub CmdINSTALLMENT_Click()
    On Error GoTo ErrTrap
    Dim Msg As String
    Dim i As Integer

    If XPTxtValue(1).text = "" Then
        Msg = "ÌÃ»  ÕœÌœ «·ﬁÌ„… «·¬Ã·… ﬁ»·  ”ÃÌ· «·√ﬁ”«ÿ"
        MsgBox Msg, vbOKOnly + vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title

        If XPTxtValue(1).Enabled = True Then
            XPTxtValue(1).SetFocus
        End If

        Exit Sub
    End If

    Load FrmInstallMent
    Set FrmInstallMent.Frm = Me

    With FrmInstallMent

        If Me.TxtModFlg.text = "R" Then
            .Tag = "R"
            .Retrive val(XPTxtValue(1).Tag)
        Else
            .Tag = "N"
            .Txt(1).text = XPTxtValue(1).text
            .LblNoteID.Caption = XPTxtSerial(1).text
            .CboPrecenType.ListIndex = val(Me.LblPrecenType.Tag)
            .Txt(3).text = val(LblPrecenValue.Caption)
            .Txt(5).text = val(LblInstallCount.Caption)

            If IsDate(Me.LblFirstInstallDate.Caption) Then
                .Dtp_First.value = Me.LblFirstInstallDate.Caption
            End If

            .Txt(7).text = val(LblInstallSeprator.Caption)

            If val(LblInstallmentType.Tag) = 0 Then
                .OptInt(0).value = True
            ElseIf val(LblInstallmentType.Tag) = 1 Then
                .OptInt(1).value = True
            ElseIf val(LblInstallmentType.Tag) = 2 Then
                .OptInt(2).value = True
            End If

            With .FG
                .rows = Me.FgInstallments.rows

                For i = 1 To Me.FgInstallments.rows - 1
                    .TextMatrix(i, .ColIndex("Serial")) = i
                    .TextMatrix(i, .ColIndex("Value")) = Me.FgInstallments.TextMatrix(i, Me.FgInstallments.ColIndex("Value"))
                    .TextMatrix(i, .ColIndex("Due_Date")) = Me.FgInstallments.TextMatrix(i, Me.FgInstallments.ColIndex("Due_Date"))
                Next i

                .AutoSize 0, .Cols - 1, False
            End With

        End If

        .show vbModal
    End With

    Exit Sub
ErrTrap:
End Sub

Private Sub CmdNotes_Click()
    ShowRelatedNotes val(Me.XPTxtBillID.text), 1
End Sub

Private Sub CmdNotes_MouseMove(Button As Integer, _
                               Shift As Integer, _
                               X As Single, _
                               Y As Single)
    Dim StrTemp As String

    If val(Me.CmdNotes.Tag) = 0 Then
        Me.CmdNotes.ToolTipText = ""
    Else
        StrTemp = " ÊÃœ ⁄·Ï Â–Â «·Õ—ﬂ… ⁄„·Ì«  „«·Ì… „ﬁœ«—Â« : " & val(Me.CmdNotes.Tag)
        Me.CmdNotes.ToolTipText = StrTemp
    End If

End Sub

Private Sub cmdReSave_Click()
    Dim s As String
    Dim i As Double
    XPBtnMove_Click (2)
    DoEvents
    For i = 1 To rs.RecordCount
        Cmd_Click (1)
        DoEvents
        ' NewGrid.DtpBillDate_Change
        DoEvents
        DoEvents
        IsSaveWithOutMsg = True
        DoEvents
        Cmd_Click (2)
        DoEvents
        XPBtnMove_Click (0)
    Next i
    IsSaveWithOutMsg = False
    MsgBox " „ «·Õ›Ÿ"
End Sub

Private Sub CmdRetruns_MouseMove(Button As Integer, _
                                 Shift As Integer, _
                                 X As Single, _
                                 Y As Single)
    Dim StrTemp As String

    If val(Me.CmdRetruns.Tag) = 0 Then
        Me.CmdRetruns.ToolTipText = ""
    Else
        StrTemp = " ÊÃœ ⁄·Ï Â–Â «·Õ—ﬂ… Õ—ﬂ«   Ã«—Ì… √Œ—Ï ·Â« ⁄·«ﬁ… »Â« ≈Ã„«·ÌÂ«: " & val(Me.CmdRetruns.Tag)
        Me.CmdRetruns.ToolTipText = StrTemp
    End If

End Sub

Private Sub Cmmadd_Click()
    'Dim D As Double
    'D = Me.Grid.TextMatrix(1, Me.Grid.ColIndex("select"))
    'Dim I As Integer
    '
    'Txt_EXport.text = 0
    '     For I = 1 To Grid.Rows - 1
    '
    '        If Val(Grid.TextMatrix(I, Grid.ColIndex("select"))) = -1 Then
    '
    '        Txt_EXport.text = Val(Txt_EXport.text) + Val(Grid.TextMatrix(I, Grid.ColIndex("note_value")))
    '
    '        End If
    '        Next
End Sub

Private Sub DBCboClientName_Change()
    Dim StrSQL As String
    Dim RsTemp As ADODB.Recordset

    On Error GoTo ErrTrap

    If Me.TxtModFlg.text = "N" Or Me.TxtModFlg.text = "E" Then
        If DBCboClientName.BoundText <> "" Then
            If DBCboClientName.BoundText = 1 Or DBCboClientName.BoundText = 2 Then
                CboPayMentType.locked = True
                '  CboPayMentType.ListIndex = 0
            Else
                CboPayMentType.locked = False
            End If
        End If
    End If

    If Me.TxtModFlg.text = "N" Or Me.TxtModFlg.text = "E" Then
        StrSQL = "Select * From TblCustemers Where CusID=" & val(DBCboClientName.BoundText)
        Set RsTemp = New ADODB.Recordset
        RsTemp.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

        If Not (RsTemp.BOF Or RsTemp.EOF) Then
            If Not (IsNull(RsTemp("Trans_DiscountTypePur").value)) Then
                If RsTemp("Trans_DiscountTypePur").value = 0 Then
                    Me.XPCboDiscountType.ListIndex = 0
                    Me.XPTxtDiscountVal.text = 0
                ElseIf RsTemp("Trans_DiscountTypePur").value = 1 Then
                    Me.XPCboDiscountType.ListIndex = 1
                    Me.XPTxtDiscountVal.text = IIf(IsNull(RsTemp("Trans_DiscountPur").value), "", RsTemp("Trans_DiscountPur").value)
                ElseIf RsTemp("Trans_DiscountTypePur").value = 2 Then
                    Me.XPCboDiscountType.ListIndex = 2
                    Me.XPTxtDiscountVal.text = IIf(IsNull(RsTemp("Trans_DiscountPur").value), "", RsTemp("Trans_DiscountPur").value)
                End If

            Else
                Me.XPCboDiscountType.ListIndex = 0
                Me.XPTxtDiscountVal.text = 0
            End If

        Else
            Me.XPCboDiscountType.ListIndex = 0
            Me.XPTxtDiscountVal.text = 0
        End If

        RsTemp.Close
        Set RsTemp = Nothing
    End If
    
    Exit Sub
ErrTrap:
End Sub

Private Sub DBCboClientName_Click(Area As Integer)
    DBCboClientName_Change
End Sub

Private Sub DCboItemsCode_KeyUp(KeyCode As Integer, _
                                Shift As Integer)
        
    If KeyCode = vbKeyF3 Then
        Load FrmItemSearch
        FrmItemSearch.RetrunType = 14
        FrmItemSearch.show vbModal
    End If

End Sub

Private Sub DCboStoreName_Change()
 TxtStoreID.text = getStoreCoding(val(DCboStoreName.BoundText))
 
    If Me.TxtModFlg = "N" Or Me.TxtModFlg = "E" Then

     If CheckStoreCoding(val(dcBranch.BoundText), 19) = True Then
     TxtNoteSerial.text = ""
    TxtNoteSerial1.text = ""

     End If
     
    End If

End Sub

Private Sub DCboStoreName_KeyUp(KeyCode As Integer, _
                                Shift As Integer)

    If KeyCode = vbKeyF5 Then
        Dim Dcombos As ClsDataCombos
        
        Set Dcombos = New ClsDataCombos
        Dcombos.GetStores Me.DCboStoreName
       
    End If

End Sub

Private Sub Dcbranch_Change()
    If Me.TxtModFlg = "N" Or Me.TxtModFlg = "E" Then
        TxtNoteSerial.text = ""
    TxtNoteSerial1.text = ""
    End If

End Sub

Private Sub Dcbranch_Click(Area As Integer)
    TxtNoteSerial.text = ""
    TxtNoteSerial1.text = ""
End Sub

Private Sub dcBranch_KeyUp(KeyCode As Integer, _
                           Shift As Integer)

    If KeyCode = vbKeyF5 Then
        Dim Dcombos As ClsDataCombos
        
        Set Dcombos = New ClsDataCombos
        Dcombos.GetBranches dcBranch
    End If

End Sub

Private Sub Ele_DblClick(Index As Integer)
    On Error GoTo ErrTrap

    Select Case Index

        Case 6

            If Me.WindowState = vbNormal Then
                Me.WindowState = vbMaximized
            Else
                Me.WindowState = vbNormal
            End If

    End Select

    Exit Sub
ErrTrap:
End Sub

Private Sub ReLineGrid()
    Dim i As Integer
    Dim IntCounter As Integer

    With Fg_Journal

        For i = .FixedRows To .rows - 1

            If .TextMatrix(i, .ColIndex("AccountCode")) <> "" Then
                IntCounter = IntCounter + 1
                .TextMatrix(i, .ColIndex("LineNo")) = IntCounter
            End If

        Next i

    End With

End Sub

Private Sub FG_AfterEdit(ByVal Row As Long, _
                         ByVal Col As Long)

    If Me.TxtModFlg <> "E" Then Exit Sub

    '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\LogFile Saving//////////////////////////////////////////////////////////////////////////
    If Col = FG.ColIndex("Code") Or Col = FG.ColIndex("Name") Then
        RegisterItemData Me.Name, Me.TxtModFlg, FG.cell(flexcpTextDisplay, Row, FG.ColIndex("Code")), FG.cell(flexcpTextDisplay, Row, FG.ColIndex("Name")), , , , , , , , , , val(Me.TxtNoteSerial), Me.TxtNoteSerial1, 250
    ElseIf Col = FG.ColIndex("UnitID") Then
        RegisterItemData Me.Name, Me.TxtModFlg, FG.cell(flexcpTextDisplay, Row, FG.ColIndex("Code")), FG.cell(flexcpTextDisplay, Row, FG.ColIndex("Name")), FG.cell(flexcpTextDisplay, Row, FG.ColIndex("UnitID")), , , , , , , , , val(Me.TxtNoteSerial), Me.TxtNoteSerial1, 250
    ElseIf Col = FG.ColIndex("Count") Then
        RegisterItemData Me.Name, Me.TxtModFlg, FG.cell(flexcpTextDisplay, Row, FG.ColIndex("Code")), FG.cell(flexcpTextDisplay, Row, FG.ColIndex("Name")), , (FG.TextMatrix(Row, FG.ColIndex("Count"))), , , , , , , , val(Me.TxtNoteSerial), Me.TxtNoteSerial1, 250
    ElseIf Col = FG.ColIndex("Price") Then
        RegisterItemData Me.Name, Me.TxtModFlg, FG.cell(flexcpTextDisplay, Row, FG.ColIndex("Code")), FG.cell(flexcpTextDisplay, Row, FG.ColIndex("Name")), , , (FG.TextMatrix(Row, FG.ColIndex("Price"))), , , , , , , val(Me.TxtNoteSerial), Me.TxtNoteSerial1, 250
    ElseIf Col = FG.ColIndex("ColorID") Then
        RegisterItemData Me.Name, Me.TxtModFlg, FG.cell(flexcpTextDisplay, Row, FG.ColIndex("Code")), FG.cell(flexcpTextDisplay, Row, FG.ColIndex("Name")), , , , , FG.cell(flexcpTextDisplay, Row, FG.ColIndex("ColorID")), , , , , val(Me.TxtNoteSerial), Me.TxtNoteSerial1, 250
    ElseIf Col = FG.ColIndex("ItemSize") Then
        RegisterItemData Me.Name, Me.TxtModFlg, FG.cell(flexcpTextDisplay, Row, FG.ColIndex("Code")), FG.cell(flexcpTextDisplay, Row, FG.ColIndex("Name")), , , , , , FG.cell(flexcpTextDisplay, Row, FG.ColIndex("ItemSize")), , , , val(Me.TxtNoteSerial), Me.TxtNoteSerial1, 250
    ElseIf Col = FG.ColIndex("ClassId") Then
        RegisterItemData Me.Name, Me.TxtModFlg, FG.cell(flexcpTextDisplay, Row, FG.ColIndex("Code")), FG.cell(flexcpTextDisplay, Row, FG.ColIndex("Name")), , , , , , , FG.cell(flexcpTextDisplay, Row, FG.ColIndex("ClassId")), , , val(Me.TxtNoteSerial), Me.TxtNoteSerial1, 250
    ElseIf Col = FG.ColIndex("DiscountType") Then
        RegisterItemData Me.Name, Me.TxtModFlg, FG.cell(flexcpTextDisplay, Row, FG.ColIndex("Code")), FG.cell(flexcpTextDisplay, Row, FG.ColIndex("Name")), , , , , , , , FG.cell(flexcpTextDisplay, Row, FG.ColIndex("DiscountType")), , val(Me.TxtNoteSerial), Me.TxtNoteSerial1, 250
    ElseIf Col = FG.ColIndex("DiscountVal") Then
        RegisterItemData Me.Name, Me.TxtModFlg, FG.cell(flexcpTextDisplay, Row, FG.ColIndex("Code")), FG.cell(flexcpTextDisplay, Row, FG.ColIndex("Name")), , , , , , , , , FG.TextMatrix(Row, FG.ColIndex("DiscountVal")), val(Me.TxtNoteSerial), Me.TxtNoteSerial1, 250

    End If

    '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\LogFile Saving//////////////////////////////////////////////////////////////////////////

End Sub

Private Sub FG_CellButtonClick(ByVal Row As Long, ByVal Col As Long)
With FG
Select Case .ColKey(Col)
Case "Show"
If val(.TextMatrix(Row, .ColIndex("NoSerial"))) <> 0 Then
Load FrmTestCertificate
FrmTestCertificate.FindRec (val(.TextMatrix(Row, .ColIndex("NoSerial"))))
FrmTestCertificate.show
End If
End Select
End With
End Sub

Private Sub Fg_DblClick()
    'FrmItemsDetails.Show
End Sub

Public Sub Fg_Journal_AfterEdit(ByVal Row As Long, _
                                ByVal Col As Long)
 
    Dim StrAccountCode As String
    Dim Msg As String
    Dim rs As New ADODB.Recordset
    Dim StrSQL As String
    Dim ClsAcc As New ClsAccounts
    Dim LngRow As Long

    With Fg_Journal

        Select Case .ColKey(Col)
 
            Case "AccountName"
                '  .TextMatrix(Row, .ColIndex("userid")) = user_id
                        
                StrAccountCode = .ComboData
                LngRow = .FindRow(StrAccountCode, .FixedRows, .ColIndex("AccountCode"), False, True)
                .TextMatrix(Row, .ColIndex("AccountCode")) = StrAccountCode
                .TextMatrix(Row, .ColIndex("ExpensesID")) = get_Expenses_id(StrAccountCode)
                .TextMatrix(Row, .ColIndex("LineNo1")) = setfoxy_Line

                If SystemOptions.UserInterface = ArabicInterface Then
                    StrSQL = "select * from Expenses_accounts where Account_Code='" & StrAccountCode & "'"
                Else
                    StrSQL = "select * from Expenses_accounts_eng where Account_Code='" & StrAccountCode & "'"
                End If
            
                rs.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText
                     
                If rs.RecordCount > 0 Then
                    .TextMatrix(Row, .ColIndex("des")) = IIf(IsNull(rs("parent_account").value), "", rs("parent_account").value)
                Else
                    .TextMatrix(Row, .ColIndex("des")) = ""
                End If

            Case "value"
                Dim sgl As String
               
                Me.TXTFactoryExpenses.text = .Aggregate(flexSTSum, .FixedRows, .ColIndex("Value"), .rows - 1, .ColIndex("Value"))
                '    sgl = "update  marakes_taklefa_temp  set value=0 where  line_no=" & Val(Fg_Journal.TextMatrix(Fg_Journal.Row, Fg_Journal.ColIndex("LineNo1")))
                '     Cn.Execute sgl, , adExecuteNoRecords
        
                '  Me.XPTxtVal.text = Format(Me.XPTxtVal.text, SystemOptions.SysDefCurrencyForamt)
        End Select

        Me.TXTFactoryExpenses.text = .Aggregate(flexSTSum, .FixedRows, .ColIndex("Value"), .rows - 1, .ColIndex("Value"))

        ' Me.XPTxtVal.text = Format(Me.XPTxtVal.text, SystemOptions.SysDefCurrencyForamt)
        'to Add new row if needed
        If Row = .rows - 1 Then
            .rows = .rows + 1
        End If

        ' ReLineGrid
    End With

    ReLineGrid
End Sub

Private Sub Fg_Journal_BeforeEdit(ByVal Row As Long, _
                                  ByVal Col As Long, _
                                  Cancel As Boolean)

    With Fg_Journal

        If Row > .FixedRows Then
            '  If .TextMatrix(Row - 1, .ColIndex("AccountCode")) = "" Then
            '      Cancel = True
            '  End If
        End If

        Select Case .ColKey(Col)

            Case "value"
                .ComboList = ""

            Case "des"
                .ComboList = ""
        
            Case "Order_No"
                .ComboList = ""
        
                '  Cancel = True
            
        End Select

    End With

End Sub

Private Sub Fg_Journal_DblClick()
    Exit Sub
  
    Static lNoteRow&, lNoteCol&, r&, c&

    With Fg_Journal
        ' clicking? no work
        'If Button <> 0 Then Exit Sub
        ' get mouse coordinates
        r = Fg_Journal.Row
        c = Fg_Journal.Col

        If Fg_Journal.ColKey(c) <> "Des" Then
            CboDes.Visible = False
            Exit Sub
        End If

        If Fg_Journal.TextMatrix(r, c) = "" Then
            'Exit Sub
        End If

        If .TextMatrix(r, .ColIndex("AccountCode")) = "" Then
            Exit Sub
        End If

        ' same cell or neighbour? no work
        '    If r = lNoteRow And C = lNoteCol Then Exit Sub
        '    If r = lNoteRow And C = lNoteCol + 1 Then Exit Sub

        ' other cell, hide current note, if any
        If lNoteRow >= 0 And lNoteCol >= 0 Then
            Fg_Journal.SetFocus
            lNoteRow = -1
            lNoteCol = -1
        End If

        ' no note to show? then bail out
        If r <= 0 Or c <= 0 Then Exit Sub
        If typename(Fg_Journal.cell(flexcpData, r, c)) <> "String" Then
            TxtDes.text = ""
        Else
            '
            TxtDes.text = Fg_Journal.cell(flexcpData, r, c)
        End If

        ' show new note
        CboDes.Move .CellLeft, .CellTop, .CellWidth, .CellHeight
        CboDes.Visible = True
        CboDes.ZOrder 0
        CboDes.SetFocus
        'save coordinates for next time
        lNoteRow = r
        lNoteCol = c
    End With

End Sub

Private Sub Fg_Journal_KeyUp(KeyCode As Integer, _
                             Shift As Integer)

    With Fg_Journal

        Select Case .ColKey(.Col)

            Case "Order_No"
                           
                If KeyCode = vbKeyF3 Then
                    Order_no_search.show
                    Order_no_search.RetrunType = 4
                End If

            Case "AccountName"

                If KeyCode = vbKeyF3 Then
                    FrmExpensesSearch.show
                    FrmExpensesSearch.RetrunType = 3
                End If
 
        End Select

    End With

End Sub

Private Sub Fg_Journal_StartEdit(ByVal Row As Long, _
                                 ByVal Col As Long, _
                                 Cancel As Boolean)
    Dim rs As New ADODB.Recordset
    Dim Rs1 As New ADODB.Recordset

    Dim StrSQL  As String
    Dim StrAccountType As String
    Dim StrComboList As String
    Dim StrComboList1 As String

    Dim Msg As String

    'Case "DebitName"
    'StrAccountType = " (ACCOUNTS.Account_Code Like 'a1%' Or ACCOUNTS.Account_Code Like 'a3%')"
    'Case "CreditName"
    'StrAccountType = " (ACCOUNTS.Account_Code Like 'a2%' Or ACCOUNTS.Account_Code Like 'a4%')"
    With Fg_Journal

        Select Case .ColKey(Col)

            Case "AccountName"
                StrSQL = "select * from Expenses_accounts"
                rs.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText
                StrComboList = Fg_Journal.BuildComboList(rs, "Account_Name", "Account_Code")

                If StrComboList <> "" Then
                    StrComboList = "|" & StrComboList
                End If

                .ComboList = StrComboList

            Case "opr_fullcode"
                StrSQL = "  select fullcode,name from terms_operations "
                rs.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText
                StrComboList1 = Fg_Journal.BuildComboList(rs, "fullcode", "fullcode")

                If StrComboList1 <> "" Then
                    StrComboList1 = "|" & StrComboList1
                End If

                .ComboList = StrComboList1
         
        End Select

    End With

End Sub

Private Sub fg_StartEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
With FG
Select Case .ColKey(Col)
Case "Show"
.ColComboList(.ColIndex("Show")) = "..."
End Select
End With
End Sub

Private Sub Form_Activate()
    Set m_MnuShowNewItemsPrices = mdifrmmain.MnuInvPurchaseMnu2
    Set m_MenuViewList = mdifrmmain.MnuInvPurchaseMnu1
    Set m_MenuShowItemCostEffect = mdifrmmain.MnuInvPurchaseMnu4
End Sub

Private Sub CmdRetruns_Click()
    ShowRelatedTransactions val(Me.XPTxtBillID.text), 1
End Sub

Private Sub Grid_StartEdit(ByVal Row As Long, _
                           ByVal Col As Long, _
                           Cancel As Boolean)
    Dim rs As New ADODB.Recordset
    Dim StrSQL  As String
    Dim StrAccountType As String
    Dim StrComboList As String
    Dim Msg As String

    'Case "DebitName"
    'StrAccountType = " (ACCOUNTS.Account_Code Like 'a1%' Or ACCOUNTS.Account_Code Like 'a3%')"
    'Case "CreditName"
    'StrAccountType = " (ACCOUNTS.Account_Code Like 'a2%' Or ACCOUNTS.Account_Code Like 'a4%')"
    With Grid

        Select Case .ColKey(Col)

            Case "ItemName"
       
                StrSQL = "Select * from QRY_temp_bill_items"
                
                rs.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText
                'StrComboList = grid4.BuildComboList(rs, "ItemName", "ItemID")
                Debug.Print StrSQL
 
                If StrComboList <> "" Then
                    StrComboList = "|" & StrComboList
                End If

                .ComboList = StrComboList
        End Select

    End With

End Sub

Private Sub Grid_AfterEdit(ByVal Row As Long, _
                           ByVal Col As Long)
    Dim StrAccountCode As String
    Dim Msg As String
    Dim rs As New ADODB.Recordset
    Dim StrSQL As String
    Dim ClsAcc As New ClsAccounts
    Dim LngRow As Long
    Dim sql As String
       
    With Grid

        Select Case .ColKey(Col)
   
            Case "ItemID"
          
                .TextMatrix(Row, Col) = Trim(.TextMatrix(Row, Col))
    
                StrSQL = "select * from QRY_temp_bill_items where ItemID=" & Trim(.TextMatrix(Row, Col))
                Set rs = Nothing
                rs.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText
            
                If Not (rs.BOF Or rs.EOF) Then
    
                    .TextMatrix(Row, .ColIndex("ItemName")) = IIf(IsNull(rs("ItemName").value), "", rs("ItemName").value)
                    .TextMatrix(Row, .ColIndex("ItemCode")) = IIf(IsNull(rs("ItemCode").value), "", rs("ItemCode").value)
                    
                Else
            
                    .TextMatrix(Row, .ColIndex("ItemName")) = ""
                    .TextMatrix(Row, .ColIndex("ItemCode")) = ""
                    .TextMatrix(Row, .ColIndex("ItemID")) = ""
 
                End If
 
                check_item_Exist_in_Grid val(.TextMatrix(Row, .ColIndex("ItemID"))), val(.TextMatrix(Row, .ColIndex("Note_value")))

            Case "ItemCode"
          
                .TextMatrix(Row, Col) = Trim(.TextMatrix(Row, Col))

                If .TextMatrix(Row, Col) = "" Then
                    Exit Sub
                End If

                StrSQL = "select * from QRY_temp_bill_items where ItemCode='" & Trim(.TextMatrix(Row, Col)) & "'"
                Set rs = Nothing
                rs.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

                If Not (rs.BOF Or rs.EOF) Then
    
                    .TextMatrix(Row, .ColIndex("ItemName")) = IIf(IsNull(rs("ItemName").value), "", rs("ItemName").value)
                    .TextMatrix(Row, .ColIndex("ItemID")) = IIf(IsNull(rs("ItemID").value), "", rs("ItemID").value)
                    
                Else
                    .TextMatrix(Row, .ColIndex("ItemCode")) = ""
                    .TextMatrix(Row, .ColIndex("ItemName")) = ""
                    .TextMatrix(Row, .ColIndex("ItemID")) = ""
 
                End If
 
            Case "ItemName"
                  
                StrAccountCode = .ComboData
                LngRow = .FindRow(StrAccountCode, .FixedRows, .ColIndex("ItemID"), False, True)
    
                Set ClsAcc = New ClsAccounts
      
                .TextMatrix(Row, .ColIndex("ItemID")) = StrAccountCode
                 
                StrSQL = "select * from QRY_temp_bill_items where ItemID= " & val(StrAccountCode)
                Set rs = Nothing
                rs.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

                If Not (rs.BOF Or rs.EOF) Then
            
                    .TextMatrix(Row, .ColIndex("ItemCode")) = rs("ItemCode").value
                Else
                    .TextMatrix(Row, .ColIndex("ItemCode")) = ""
                    .TextMatrix(Row, .ColIndex("ItemID")) = ""
                    .TextMatrix(Row, .ColIndex("ItemName")) = ""
                   
                End If

        End Select

        'to Add new row if needed
        If Row = .rows - 1 Then
            '    .Rows = .Rows + 1
        End If

    End With

    Expenses_update_total
End Sub

Private Sub Grid_BeforeEdit(ByVal Row As Long, _
                            ByVal Col As Long, _
                            Cancel As Boolean)

    With Grid

        If .ColKey(Col) <> "ItemName" Then
            .ComboList = ""
        End If
   
    End With

End Sub

Function Expenses_update_total()
    Dim i As Integer
    On Error Resume Next

    If Grid.rows = 1 Then Exit Function
    Txt_EXport.text = 0

    With Grid

        For i = 1 To Grid.rows - 1
        
            If .cell(flexcpChecked, i, .ColIndex("select")) = flexChecked And Grid.TextMatrix(i, Grid.ColIndex("ItemID")) = "" Then
                Txt_EXport.text = val(Txt_EXport.text) + val(Grid.TextMatrix(i, Grid.ColIndex("note_value")))
            End If
            
            If val(Grid.TextMatrix(i, Grid.ColIndex("select"))) = 0 Then
                Grid.TextMatrix(i, Grid.ColIndex("ItemID")) = ""
                Grid.TextMatrix(i, Grid.ColIndex("ItemCode")) = ""
                Grid.TextMatrix(i, Grid.ColIndex("ItemName")) = ""
            
            End If
            
        Next
 
    End With
       
End Function

Function Retrive_Expenses_Vouchers()
    '   ????? ?????? ?????????

    With Me.Grid
        .rows = .FixedRows
        .ExtendLastCol = True
        .RowHeightMin = 300
        .Editable = flexEDKbdMouse
        .ExplorerBar = flexExSortShowAndMove

        '    .AutoSize 0, .Cols - 1, False
    End With

    Dim i As Integer
    Dim RsExp As ADODB.Recordset
    Dim My_SQL As String

    Set RsExp = New ADODB.Recordset

    My_SQL = "SELECT dbo.Notes.NoteID,dbo.Notes.buy,dbo.Notes.NoteSerial,dbo.notes.ItemID , dbo.Notes.Note_Value, dbo.ExpensesType.Name ,  dbo.ExpensesType.Account_Code FROM dbo.Notes INNER JOIN dbo.ExpensesType ON dbo.Notes.ExpensesID = dbo.ExpensesType.ID Where (dbo.Notes.NoteType = 3   and order_no='" & Me.TXT_order_no.text & "' and(Transaction_ID1 is null or Transaction_ID1=" & val(Me.XPTxtBillID.text) & ")  )  "
    'My_SQL = ""

    RsExp.Open My_SQL, Cn, adOpenKeyset, adLockReadOnly, adCmdText
    Dim StrSQL  As String

    With Me.Grid
        .rows = 1
        .Clear flexClearScrollable

        If RsExp.RecordCount > 0 Then
            .rows = RsExp.RecordCount + 1
            RsExp.MoveFirst

            For i = 1 To .rows - 1
                   
                .TextMatrix(i, .ColIndex("ItemID")) = IIf(IsNull(RsExp.Fields("ItemID").value), "", RsExp.Fields("ItemID").value)
    
                StrSQL = "select * from TblItems where ItemID=" & val(.TextMatrix(i, .ColIndex("ItemID")))
                Dim rs As New ADODB.Recordset
                Set rs = Nothing
                rs.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

                If Not (rs.BOF Or rs.EOF) Then
    
                    .TextMatrix(i, .ColIndex("ItemName")) = IIf(IsNull(rs("ItemName").value), "", rs("ItemName").value)
                    .TextMatrix(i, .ColIndex("ItemCode")) = IIf(IsNull(rs("ItemCode").value), "", rs("ItemCode").value)
                    
                Else
            
                    .TextMatrix(i, .ColIndex("ItemName")) = ""
                    .TextMatrix(i, .ColIndex("ItemCode")) = ""
 
                End If
               
                .TextMatrix(i, .ColIndex("Name")) = IIf(IsNull(RsExp.Fields("Name").value), "", RsExp.Fields("Name").value)
               
                .TextMatrix(i, .ColIndex("NoteSerial")) = IIf(IsNull(RsExp.Fields("NoteSerial").value), "", RsExp.Fields("NoteSerial").value)
            
                .TextMatrix(i, .ColIndex("NoteID")) = IIf(IsNull(RsExp.Fields("NoteID").value), "", RsExp.Fields("NoteID").value)
           
                .TextMatrix(i, .ColIndex("Note_Value")) = IIf(IsNull(RsExp.Fields("Note_Value").value), "", RsExp.Fields("Note_Value").value)
                .TextMatrix(i, .ColIndex("Account_Code")) = IIf(IsNull(RsExp.Fields("Account_Code").value), "", RsExp.Fields("Account_Code").value)
            
                If IsNull(RsExp.Fields("buy").value) Then
                    .TextMatrix(i, .ColIndex("Select")) = 0
                Else

                    If RsExp.Fields("buy").value = False Then
                        .TextMatrix(i, .ColIndex("Select")) = 0
                    ElseIf RsExp.Fields("buy").value = True Then
                        .TextMatrix(i, .ColIndex("Select")) = 1
                    Else
                        .TextMatrix(i, .ColIndex("Select")) = 0
                    End If
           
                End If
           
                RsExp.MoveNext
            Next

            RsExp.Close
        End If

        .RowHeight(-1) = 300
    End With

    Grid.Visible = True

    '   ????? ?????? ?????????

    Expenses_update_total

End Function
 
Private Function check_item_Exist_in_Grid(ItemID As Integer, _
                                          value As Single, _
                                          Optional addition As Boolean)
    Dim i As Integer

    With FG

        For i = 1 To FG.rows - 1

            If .TextMatrix(i, .ColIndex("Code")) = ItemID Then
                If addition = False Then
                    .TextMatrix(i, .ColIndex("LineShahn")) = value
                Else
                    .TextMatrix(i, .ColIndex("LineShahn")) = val(.TextMatrix(i, .ColIndex("LineShahn"))) + value
                End If

                Exit Function
    
            End If

        Next i

    End With
 
End Function

Private Sub LblTotal_MouseMove(Button As Integer, _
                               Shift As Integer, _
                               X As Single, _
                               Y As Single)
    LblTotal.ToolTipText = WriteNo(LblTotal.Caption, 0, True)
End Sub

Private Sub m_FrmSearch_Unload(Cancel As Integer)
    Set m_FrmSearch = Nothing
End Sub

Private Sub m_MenuShowItemCostEffect_Click()

    If Me.TxtModFlg.text = "R" Then
        ShowItemCostEffectForTrans 1, , Trim$(Me.TxtTransSerial.text)
    End If

End Sub

Private Sub m_MenuViewList_Click()
    Dim FrmView As FrmViewList
    Dim FG As VSFlex8UCtl.VSFlexGrid
    Dim StrSQL As String
    Dim rs As ADODB.Recordset
    Dim StrComboList As String
    Dim GrdBack As ClsBackGroundPic
    Dim cProgress As ClsProgress
    Dim BolFrmLoaded As Boolean
    Set FrmView = New FrmViewList
    Set FG = FrmView.vsfGroup1.VSFlexGrid

    With FG
        .Cols = 9
        .RowHeightMin = 320
        .TextMatrix(0, 0) = "—ﬁ„ «·»—‰«„Ã"
        .ColKey(0) = "Transaction_ID"
        .TextMatrix(0, 1) = "—ﬁ„ «·›« Ê—…"
        .TextMatrix(0, 2) = " «—ÌŒ «·›« Ê—…"
        .ColDataType(2) = flexDTDate
        .TextMatrix(0, 3) = "«”„ «·„Ê—œ"
        .TextMatrix(0, 4) = "ÿ—Ìﬁ… «·œ›⁄"
        StrComboList = "#0;‰ﬁœÏ|#1;√Ã·"
        .ColComboList(4) = StrComboList
    
        .TextMatrix(0, 5) = "«”„ «·„Œ“‰"
        .TextMatrix(0, 6) = "‰Ê⁄ «·Œ’„"
        .TextMatrix(0, 7) = "ﬁÌ„… «·Œ’„"
        .TextMatrix(0, 8) = "≈Ã„«·Ï «·›« Ê—…"

        ',
        'QryTransactionsTotal.TransSum
        'QryTransactionsTotal.TransNet,
        If SystemOptions.SysDataBaseType = SQLServerDataBase Then
            StrSQL = "SELECT TOP 100 PERCENT QryTransactionsTotal.Transaction_ID," & "QryTransactionsTotal.Transaction_Serial, QryTransactionsTotal.Transaction_Date, " & "dbo.TblCustemers.CusName, QryTransactionsTotal.PaymentType, dbo.TblStore.StoreName," & "QryTransactionsTotal.Trans_DiscountType,QryTransactionsTotal.Trans_Discount ," & "QryTransactionsTotal.TotalAfterTax "
            StrSQL = StrSQL + " FROM dbo.QryTransactionsTotal() QryTransactionsTotal LEFT OUTER JOIN "
            StrSQL = StrSQL + "dbo.TblStore ON QryTransactionsTotal.StoreID = dbo.TblStore.StoreID " & "LEFT OUTER JOIN dbo.TblCustemers ON QryTransactionsTotal.CusID = dbo.TblCustemers.CusID"
            StrSQL = StrSQL + " Where (QryTransactionsTotal.Transaction_Type = 1)"
            StrSQL = StrSQL + " ORDER BY QryTransactionsTotal.Transaction_ID "
        ElseIf SystemOptions.SysDataBaseType = AccessDataBase Then
            StrSQL = "SELECT QryTransactionsTotal.Transaction_ID , QryTransactionsTotal.Transaction_Serial," & "QryTransactionsTotal.Transaction_Date,TblCustemers.CusName, QryTransactionsTotal.PaymentType," & "TblStore.StoreName,TblEmployee.Emp_Name ,QryTransactionsTotal.Trans_DiscountType," & "QryTransactionsTotal.Trans_Discount,QryTransactionsTotal.TotalAfterTax "
            StrSQL = StrSQL + "FROM (TblEmployee RIGHT JOIN (TblCustemers RIGHT JOIN QryTransactionsTotal " & "ON TblCustemers.CusID = QryTransactionsTotal.CusID) ON TblEmployee.Emp_ID = QryTransactionsTotal.Emp_ID) " & "LEFT JOIN TblStore ON QryTransactionsTotal.StoreID = TblStore.StoreID "
            StrSQL = StrSQL + " WHERE QryTransactionsTotal.Transaction_Type= 1 "
            StrSQL = StrSQL + " Order  By QryTransactionsTotal.Transaction_ID"
        End If

        Set rs = New ADODB.Recordset
        rs.Open StrSQL, Cn, adOpenKeyset, adLockReadOnly, adAsyncExecute + adAsyncFetch
        Set cProgress = New ClsProgress
        BolFrmLoaded = True
        cProgress.ProgressType = Waiting
        cProgress.StartProgress

        Do While rs.State = adStateExecuting
            DoEvents
        Loop

        If BolFrmLoaded = True Then
            cProgress.StopProgess
            Set cProgress = Nothing
        End If

        Set .DataSource = rs
        .TextMatrix(0, 0) = "—ﬁ„ «·»—‰«„Ã"
        .ColKey(0) = "Transaction_ID"
        .TextMatrix(0, 1) = "—ﬁ„ «·›« Ê—…"
        .TextMatrix(0, 2) = " «—ÌŒ «·›« Ê—…"
        .ColDataType(2) = flexDTDate
        .TextMatrix(0, 3) = "«”„ «·„Ê—œ"
        .TextMatrix(0, 4) = "ÿ—Ìﬁ… «·œ›⁄"
        StrComboList = "#0;‰ﬁœÏ|#1;√Ã·"
        .ColComboList(4) = StrComboList
        .TextMatrix(0, 5) = "«”„ «·„Œ“‰"
        .TextMatrix(0, 6) = "‰Ê⁄ «·Œ’„"
        .TextMatrix(0, 7) = "ﬁÌ„… «·Œ’„"
        .TextMatrix(0, 8) = "≈Ã„«·Ï «·›« Ê—…"
        .ColKey(8) = "TotalAfterTax"
        'Rs.Close
        'Set Rs = Nothing
    End With

    Set GrdBack = New ClsBackGroundPic
    FrmView.vsfGroup1.VSFlexGrid.WallPaper = GrdBack.Picture
    FrmView.vsfGroup1.SetRTL = True
    FrmView.vsfGroup1.TotalOnColKey = "TotalAfterTax"
    FrmView.vsfGroup1.update
    FrmView.BolRetrunOnDblClick = True
    FrmView.SetDblClickRetrun Me, "Transaction_ID"
    FrmView.Caption = "⁄—÷ ‘Ã—Ï ÃœÊ·Ï ·›Ê« Ì— «·„‘ —Ì« "
    FrmView.show
End Sub

Private Sub m_MnuShowNewItemsPrices_Click()

    If Not NewGrid Is Nothing Then
        NewGrid.ShowNewItemsPrice
    End If

End Sub

Private Sub Txt_EXport_GotFocus()
    'On Error GoTo ErrTrap

    'With Me.Grid
    '    .Rows = .FixedRows
    '    .ExtendLastCol = True
    '    .RowHeightMin = 300
    '    .Editable = flexEDKbdMouse
    '    .ExplorerBar = flexExSortShowAndMove
    '
    '    .AutoSize 0, .Cols - 1, False
    'End With

    'Dim I As Integer
    'Dim Rs As ADODB.Recordset
    'Dim My_SQL As String
    '
    'Set Rs = New ADODB.Recordset
    '
    'My_SQL = "SELECT dbo.Notes.NoteID , dbo.Notes.Note_Value, dbo.ExpensesType.Name FROM dbo.Notes INNER JOIN dbo.ExpensesType ON dbo.Notes.ExpensesID = dbo.ExpensesType.ID Where (dbo.Notes.NoteType = 3)"
    '
    ''    My_SQL = "select * From TblEmployee  where DateEndPasp < getdate()"
    '
    'Rs.Open My_SQL, Cn, adOpenKeyset, adLockReadOnly, adCmdText

    'With Me.Grid
    '    .Rows = 2
    '    .Clear flexClearScrollable
    '    If Rs.RecordCount > 0 Then
    '        .Rows = Rs.RecordCount + 1
    '        Rs.MoveFirst
    '        For I = 1 To .Rows - 1
    ''             .TextMatrix(i, .ColIndex("Ser")) = i
    ''
    '             .TextMatrix(I, .ColIndex("Name")) = IIf(IsNull(Rs.Fields("Name").Value), _
    '            "", Rs.Fields("Name").Value)
    '
    '            .TextMatrix(I, .ColIndex("NoteID")) = IIf(IsNull(Rs.Fields("NoteID").Value), _
    '            "", Rs.Fields("NoteID").Value)
    '
    '                        .TextMatrix(I, .ColIndex("Note_Value")) = IIf(IsNull(Rs.Fields("Note_Value").Value), _
    '            "", Rs.Fields("Note_Value").Value)
    '
    '            Rs.MoveNext
    '        Next
    '       Rs.Close
    '    End If
    '    .RowHeight(-1) = 300
    'End With
    'ErrTrap:

    'Dim StrSQL As String
    'Dim i As Double
    '
    'StrSQL = "SELECT dbo.Notes.NoteID , dbo.Notes.Note_Value, dbo.ExpensesType.Name FROM dbo.Notes INNER JOIN dbo.ExpensesType ON dbo.Notes.ExpensesID = dbo.ExpensesType.ID Where (dbo.Notes.NoteType = 3)"
    'Set Rs = New ADODB.Recordset
    'Rs.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText
    'If Not (Rs.BOF Or Rs.EOF) Then
    '
    '
    '
    'Rs.MoveFirst
    '   For i = 0 To Rs.RecordCount - 1
    '
    '   lstExp.AddItem Rs("NoteID") & Space$(5) & Rs("Note_Value") & Space$(5) & Rs("Name")
    '
    '    Rs.MoveNext
    '
    '    Next
    'End If

End Sub

Private Sub Txt_EXport_MouseMove(Button As Integer, _
                                 Shift As Integer, _
                                 X As Single, _
                                 Y As Single)
    Txt_EXport.ToolTipText = "„Ã„Ê⁄ «·„’—Ê›«  « Ê„« ÌﬂÌ« ⁄·Ï «–‰ «·«÷«›… "
End Sub

Private Sub Txt_order_no_Change()
    Retrive_Expenses_Vouchers
End Sub

Private Sub TxtFillData_Change()

    If TxtFillData.text = "F" Then
        NewGrid.Calculate 1, , , True
    End If

End Sub

Private Sub TXTOrderNO1_KeyUp(KeyCode As Integer, _
                              Shift As Integer)

    If KeyCode = vbKeyF3 Then
        Order_no_search2.show
        Order_no_search2.RetrunType = 6
 
    End If

End Sub

Private Sub txtPassword_Change()
If Trim(txtPassword) = "Alex2025" Then
    cmdReSave.Visible = True
    txtFromDateReSave.Visible = True
    txtToDateReSave.Visible = True
    chkIsBranch(0).Visible = True
    chkIsBranch(1).Visible = True
    chkIgnorDetails.Visible = True
        chkIgnorDetails.value = 1
Else
    cmdReSave.Visible = False
    txtFromDateReSave.Visible = False
    txtToDateReSave.Visible = False
   chkIsBranch(0).Visible = False
   chkIsBranch(1).Visible = False
   chkIgnorDetails.Visible = False
        
End If
txtFromDateReSave.value = Date
txtToDateReSave.value = Date
End Sub

Private Sub XPBtnMove_Click(Index As Integer)
    On Error GoTo ErrTrap

    If Me.TxtModFlg.text = "N" Then
        clear_all Me
        Me.TxtModFlg.text = "R"
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

         '   If Not (rs.EOF Or rs.BOF) Then
             '   rs.MoveLast
           ' End If
DisplayRec:
         Me.TxtModFlg.text = ""
        Dim StrSQL As String
     StrSQL = "SELECT * FROM Transactions WHERE Transaction_Type=28 "
     
StrSQL = StrSQL & "  AND      BranchId in(" & Current_branchSql & ")"
            If SystemOptions.usertype <> UserAdminAll Then
           '     StrSQL = StrSQL & " AND   BranchId=" & Current_branch
            End If


     If SystemOptions.usertype <> UserAdminAll Then
 
          If SystemOptions.FixedCustomer = 1 Then
            StrSQL = StrSQL & " and  UserID = " & user_id
             End If
  
  
        Me.dcBranch.Enabled = True
      
      
    End If
    
            If SystemOptions.SortInvoiceByEntry Then
                StrSQL = StrSQL & " Order by Transaction_ID"
            Else
                StrSQL = StrSQL & " Order by noteserial1"
            End If
                
      If cmdReSave.Visible = True Then
    
    StrSQL = " SELECT * FROM Transactions WHERE Transaction_Type = 28 "
    StrSQL = StrSQL & "   and ( Transaction_Date >= " & SQLDate(txtFromDateReSave.value, True) & " and "
    StrSQL = StrSQL & "   Transaction_Date <=   " & SQLDate(txtToDateReSave.value, True) & " )"
    
    StrSQL = StrSQL & "   AND Transaction_ID IN (SELECT"
    StrSQL = StrSQL & "           tt.Transaction_ID"
    StrSQL = StrSQL & "       FROM dbo.Transactions TT"
    StrSQL = StrSQL & "       INNER JOIN dbo.Transaction_Details"
    StrSQL = StrSQL & "           ON TT.Transaction_ID = dbo.Transaction_Details.Transaction_ID"
    StrSQL = StrSQL & "       Where"
    StrSQL = StrSQL & "       (TT.Transaction_Type =28)"
    StrSQL = StrSQL & "       AND"
    StrSQL = StrSQL & "       (price <>GardQty ))"
    StrSQL = StrSQL & "   ORDER BY Transaction_Date, BranchId, Transaction_ID Desc"

    
    If chkIsBranch(0).value = vbChecked And val(Me.dcBranch.BoundText) > 0 Then
        StrSQL = StrSQL & "  and BranchID =  " & val(Me.dcBranch.BoundText)
    End If
     If chkIsBranch(1).value = vbChecked Then
        StrSQL = StrSQL & "  and Transaction_ID in "
        StrSQL = StrSQL & "  ( Select Transaction_ID from Transactions where Transaction_Type=28 and NoteId not In (SELECT IsNull(notes_id,0) FROM DOUBLE_ENTREY_VOUCHERS where Credit_Or_Debit = 0))"
    End If

End If
                
            Set rs = New ADODB.Recordset
            rs.Open StrSQL, Cn, adOpenKeyset, adLockOptimistic, adCmdText


            If Not (rs.EOF Or rs.BOF) Then
                rs.MoveLast
            End If
        Me.TxtModFlg.text = "R"
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
    On Error GoTo ErrTrap

    If KeyCode = vbKeyReturn Then
        If Me.TxtModFlg.text = "R" Then
            '        Cmd_Click (0)
        Else
            '        SendKeys "{TAB}"
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

    If KeyCode = vbKeyF3 Then
        If Cmd(5).Enabled = False Then Exit Sub
        Cmd_Click (5)
    End If

    If KeyCode = vbKeyF6 Then
        If Cmd(7).Enabled = False Then Exit Sub
        Cmd_Click (7)
    End If

    If KeyCode = vbKeyF2 Then
        If TxtModFlg.text = "N" Or TxtModFlg.text = "E" Then
            'XPBtnAdd_Click
        End If
    End If

    If KeyCode = vbKeyF3 Then
        If TxtModFlg.text = "N" Or TxtModFlg.text = "E" Then
            'XPBtnRemove_Click
        End If
    End If

    If KeyCode = vbKeyF5 Then
        If TxtModFlg.text = "N" Or TxtModFlg.text = "E" Then
            XPBtnNewClients_Click
        End If
    End If

    If Shift = 2 Then
        If KeyCode = vbKeySpace Then
            If TxtModFlg.text = "N" Or TxtModFlg.text = "E" Then
                'XPFillData_Click
            End If
        End If
    End If

    If Shift = 2 Then
        XPTab301.SetFocus

        If KeyCode = vbKeyTab Then
            If XPTab301.CurrTab = 0 Then
                XPTab301.CurrTab = 0

                If XPChkPayType(0).Enabled = True Then
                    XPChkPayType(0).SetFocus
                End If

            Else
                XPTab301.CurrTab = 0
                FG.SetFocus
            End If
        End If
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

Private Sub ChangeLang()

    Dim XPic As IPictureDisp
    Set XPic = Me.XPBtnMove(1).ButtonImage
    Set Me.XPBtnMove(1).ButtonImage = Me.XPBtnMove(2).ButtonImage
    Set Me.XPBtnMove(2).ButtonImage = XPic

    Set XPic = Me.XPBtnMove(0).ButtonImage
    Set Me.XPBtnMove(0).ButtonImage = Me.XPBtnMove(3).ButtonImage
    Set Me.XPBtnMove(3).ButtonImage = XPic
lbl(65).Caption = "Q.A. No."
lbl(86).Caption = "Remarks"
    'Me.LblShortcutKeys.Caption = "(New F12 OR Enter) ,(Edit F11),(Save F10),(Undo F9),(Delete F8),(Search F7)"
    Me.Caption = "Production Recive Voucher"
    Ele(6).Caption = Me.Caption
    lbl(8).Caption = " ID"
    lbl(7).Caption = " Date"
    lbl(6).Caption = "Vendor Name"
    lbl(4).Caption = "Store "
    lbl(9).Caption = "Work Order NO "
    'lbl(25).Caption = "Employee "
    lbl(10).Caption = "Payment Type"
    lbl(5).Caption = "Discount Type"
    lbl(11).Caption = "Value"

    Label1.Caption = "Another Expenses"
    CmdConvert.Caption = "Convert to bill"

    'lbl(22).Caption = "Profit Value"
    'lbl(23).Caption = "Profit Perce"

    lbl(3).Caption = " Total Cost "
    lbl(50).Caption = "Disc"
    lbl(24).Caption = " Net:"

    lbl(1).Caption = " By:"
    lbl(0).Caption = "Rec. Count:"

    lbl(31).Caption = "Item Code"
    lbl(30).Caption = "Item Name"
    lbl(29).Caption = " Case"
    lbl(28).Caption = " Serial"
    lbl(27).Caption = "QTY"
    lbl(26).Caption = "Price"
    lbl(32).Caption = "Sales Type"
    lbl(33).Caption = "Cash Customer"
    Me.Cmd(0).Caption = "New"
    Me.Cmd(1).Caption = "Edit"
    Me.Cmd(2).Caption = "Save"
    Me.Cmd(3).Caption = "Undo"
    Me.Cmd(4).Caption = "Delete"
    Me.Cmd(5).Caption = "Search"
    Me.Cmd(6).Caption = "Exit"
    Me.Cmd(7).Caption = "Print"
    Me.CmdHelp.Caption = "Help"

    Me.XPTab301.TabCaption(0) = "Items"
    Me.XPTab301.TabCaption(3) = "Expense"
    Me.XPTab301.TabCaption(4) = "Financial Invoices "
    Me.XPTab301.TabCaption(5) = "  Another Expenss"

    Label3.Caption = "Branch"
    lbl(56).Caption = "Based On"
    lbl(52).Caption = "LC No:"
    lbl(51).Caption = "Order No:"
    lbl(57).Caption = "Pricing"
    Frame3.Caption = "Info"
    lbl(58).Caption = "Source"
    lbl(63).Caption = "Total Qty"
    lbl(55).Caption = "NO:"
           
    lbl(59).Caption = "Purchase Inv No:"
    Cmd(10).Caption = "Print JL"
    Frame5.Caption = "JL NO"
 
    Me.XPTab301.TabCaption(1) = "Notes"
    lbl(20).Caption = "Payment Method"
    XPChkPayType(0).Caption = "Cahs"
    XPChkPayType(1).Caption = "Due"
    XPChkPayType(0).Caption = "Check"
    lbl(13).Caption = "Value"
    lbl(15).Caption = "Value"
    lbl(16).Caption = "Value"
    lbl(12).Caption = "Serial"
    lbl(14).Caption = "Serial"
    '    lbl(11).Caption = "Box Name"
    lbl(21).Caption = "Due Date"
    
    lbl(18).Caption = "Check NO."
    lbl(17).Caption = "Bank Name"
    lbl(19).Caption = "Due Date"
    CmdINSTALLMENT.Caption = "INSTALLMENT"
    Me.XPTab301.TabCaption(2) = "Comment On Invoice"
    Me.Ele(15).Caption = "Write any Comments about this Invoice"

    With Me.FG
        .TextMatrix(0, .ColIndex("NewItem")) = "NewItem"
         .TextMatrix(0, .ColIndex("Show")) = "Show Test"
 .TextMatrix(0, .ColIndex("NoSerial")) = "No.Test"
    End With
 
    With Me.Grid

        .TextMatrix(0, .ColIndex("Select")) = "Select"
        .TextMatrix(0, .ColIndex("NoteID")) = "NoteID"
        .TextMatrix(0, .ColIndex("NoteSerial")) = "NoteID"

        .TextMatrix(0, .ColIndex("Note_Value")) = "Note_Value"
        .TextMatrix(0, .ColIndex("name")) = "name"

        .TextMatrix(0, .ColIndex("ItemCode")) = "Item Code"
        .TextMatrix(0, .ColIndex("ItemName")) = "Item Name"
    End With

    With Me.Grid4
        '

        .TextMatrix(0, .ColIndex("Select")) = "Select"
        .TextMatrix(0, .ColIndex("NoteSerial1")) = "NoteID"
        .TextMatrix(0, .ColIndex("name")) = "Account Name"

        .TextMatrix(0, .ColIndex("Note_Value")) = "Note_Value"
        .TextMatrix(0, .ColIndex("ItemCode")) = "Item Code"
        .TextMatrix(0, .ColIndex("ItemName")) = "Item Name"

    End With
 
    Cmd(9).Caption = "Delete Row"
    Label18.Caption = "Total"
    Label19.Caption = "Anothe Expenses"
    lbl(64).Caption = "Financial Invoices"
    lbl(61).Caption = " Total"
    Command4.Caption = "View Fin Invoices"
    lbl(54).Caption = "Expenses Vouchers"
    lbl(53).Caption = " Total"
 
    With Me.Fg_Journal
        .TextMatrix(0, .ColIndex("LineNo")) = "I"
        .TextMatrix(0, .ColIndex("AccountName")) = "Expenses Name"
        .TextMatrix(0, .ColIndex("value")) = "value"

        .TextMatrix(0, .ColIndex("des")) = "des"
    End With
 
End Sub

Private Sub Form_Load()
    Dim RsClients As New ADODB.Recordset
    Dim StrSQL As String
    Dim Num As Integer
    Dim StrList As String
    Dim Dcombos As ClsDataCombos
    Dim BGround As New ClsBackGroundPic
    Dim RsNote As New ADODB.Recordset

    On Error GoTo ErrTrap

    If SystemOptions.UserInterface = EnglishInterface Then
        SetInterface Me
        ChangeLang
    End If

    ScreenNameArabic = "”‰œ «” ·«„ «‰ «Ã  „—«Õ·  «„  "
    ScreenNameEnglish = " Productuion Recieve Voucher Steps "
    RegisterLogInOut Me.Name, ScreenNameArabic, ScreenNameEnglish, "1", 250

    Dim My_SQL As String
    My_SQL = "  select branch_id,branch_name from TblBranchesData order by branch_id   "
    fill_combo dcBranch, My_SQL
 
    If SystemOptions.usertype <> UserAdminAll Then
 
        Me.dcBranch.Enabled = True
    End If

    Set Cmd(0).ButtonImage = mdifrmmain.ImgLstTree.ListImages("New").Picture
    Set Cmd(1).ButtonImage = mdifrmmain.ImgLstTree.ListImages("Edit").Picture
    Set Cmd(2).ButtonImage = mdifrmmain.ImgLstTree.ListImages("save").Picture
    Set Cmd(3).ButtonImage = mdifrmmain.ImgLstTree.ListImages("Undo").Picture
    Set Cmd(4).ButtonImage = mdifrmmain.ImgLstTree.ListImages("Del").Picture
    Set Cmd(5).ButtonImage = mdifrmmain.ImgLstTree.ListImages("Search").Picture
    Set Cmd(6).ButtonImage = mdifrmmain.ImgLstTree.ListImages("Exit").Picture
    Set Cmd(7).ButtonImage = mdifrmmain.ImgLstTree.ListImages("Print").Picture
    Set CmdHelp.ButtonImage = mdifrmmain.ImgLstTree.ListImages("Help").Picture

    SetDtpickerDate XPDtbBill
    Set NewGrid = New ClsGrid
    NewGrid.GridTrans = ProductionMaterialReciveVoucher
 
    Set NewGrid.Grid = Me.FG
    Set NewGrid.TxtInvID = Me.XPTxtBillID
    Set NewGrid.TxtModFlag = Me.TxtModFlg
    Set NewGrid.txtTotal = Me.XPTxtSum
    Set NewGrid.CboDiscount_Type = XPCboDiscountType
    Set NewGrid.TxtDiscount_Val = XPTxtDiscountVal
    Set NewGrid.TxtValueCash = XPTxtValue(0)
    Set NewGrid.TxtValueDelay = XPTxtValue(1)
    Set NewGrid.TxtValuechque = XPTxtValue(2)
    Set NewGrid.TxtFillData = TxtFillData
    Set NewGrid.DtpBillDate = Me.XPDtbBill
    Set NewGrid.StoreName = Me.DCboStoreName
    Set NewGrid.GrdTBar = Me.TBar
    Set NewGrid.LblTotalQty = Me.LblTotalQty
    '-----------------------------------------------------------------------------
    Set NewGrid.TxtTaxValue = Me.XPTxtTaxValue
    Set NewGrid.TxtAddTax = Me.TxtTaxAddValue
    Set NewGrid.TxtStampTax = Me.TxtTaxStampValue
    Set NewGrid.TxtServiceTax = Me.TxtTaxServiceValue
    '-----------------------------------------------------------------------------
    ' ⁄»∆… »Ì«‰«  «·√’‰«›
    Set NewGrid.DCboItemName = DCboItemsName
    Set NewGrid.DCboItemCode = DCboItemsCode
        Set NewGrid.StoreName = DCboStoreName
            Set NewGrid.DtpBillDate = Me.XPDtbBill
            
    Set NewGrid.CboItemCase = CboItemCase
    Set NewGrid.CmdAddData = CmdAdd
    Set NewGrid.TxtSerial = TxtSerial
    Set NewGrid.TxtQuantity = TxtQuantity
    Set NewGrid.TxtPrice = TxtPrice
    Set NewGrid.LblItemsCount = Me.LblItemsCount
    Set NewGrid.LblTotalAll = Me.LblTotalAll
    Set NewGrid.LblDiscountsTotal = Me.LblDiscountsTotal
    Set NewGrid.LblTaxSalesValue = Me.lbl(25)
    Set NewGrid.LblTaxAddValue = Me.lbl(32)
    Set NewGrid.LblTaxStampValue = Me.lbl(33)
    Set NewGrid.LblTaxServiceValue = Me.lbl(49)

    FG.WallPaper = BGround.Picture

    AddTip
    XPTab301.CurrTab = 0
    XPDtbBill.value = Date
    
            
        With cmbTypeTrans
            .Clear
            .AddItem "Acceptable quantities"
            .AddItem "Rejected quantities"
        End With


    If SystemOptions.UserInterface = EnglishInterface Then

        With XPCboDiscountType
            .Clear
            .AddItem "NO"
            .AddItem "Value  "
            .AddItem "Percentage"
        End With

        With CboPayMentType
            .Clear
            .AddItem "Cash"
            .AddItem "Credit"
        End With

    Else

        With XPCboDiscountType
            .Clear
            .AddItem "·«ÌÊÃœ Œ’„"
            .AddItem "Œ’„ »ﬁÌ„…"
            .AddItem "Œ’„ »‰”»…"
        End With

        With CboPayMentType
            .Clear
            .AddItem "‰ﬁœ«"
            .AddItem "¬Ã·"
        End With

    End If

    With Me.CBoBasedON
        .Clear
        .AddItem "»·«"
        .AddItem "√„— ‘—¡"
        .AddItem "›« Ê—… „»œ∆ÌÂ"
    End With

    With Me.CBOSource
        .Clear
        .AddItem "ÌœÊÌ"
        .AddItem "√·Ì "
     
    End With

    With Me.CboPriceType
        .Clear
        .AddItem "€Ì— „Õœœ"
        .AddItem "  «Œ— ”⁄— ‘—«¡"
        .AddItem "   ﬂ·›Â ÌœÊÌ  "
        .AddItem "  «” ·«„ ﬂ„Ì«  ›ﬁÿ"
 
    End With

    NewGrid.FillGrid
    Set Dcombos = New ClsDataCombos
    Dcombos.GetBoxes Me.DcboBox
    Dcombos.GetUsers Me.DCboUserName
    Dcombos.GetCustomersSuppliers 0, Me.DBCboClientName, True
    Set cSearchDcbo(0) = New clsDCboSearch
    Set cSearchDcbo(0).Client = Me.DBCboClientName
    cSearchDcbo(0).SetBuddyText Me.TxtCusID
    Dcombos.GetStores Me.DCboStoreName
    Set cSearchDcbo(2) = New clsDCboSearch
    Set cSearchDcbo(2).Client = Me.DCboStoreName
  '  cSearchDcbo(2).SetBuddyText Me.TxtStoreID
    '-----------------------------------------
    SetDtpickerDate Me.DtpDelayDate
    '≈⁄œ«œ Ã—œ «·√ﬁ”«ÿ
    ChkInstall.value = Unchecked
    ChkInstall.Enabled = False

    With Me.FgInstallments
        .rows = .FixedRows
        Set .WallPaper = BGround.Picture
        .RowHeightMin = 300
        .AutoSize 0, .Cols - 1, False
    End With

    With Me.FgCheques
        .rows = .FixedRows
        Set .WallPaper = BGround.Picture
        .RowHeightMin = 300
        .AutoSize 0, .Cols - 1, False
    End With

    Me.XPChkTAX.value = vbUnchecked
    XPChkTAX_Click
    Me.ChkTaxAdd.value = vbUnchecked
    ChkTaxAdd_Click
    Me.ChkTaxStamp.value = vbUnchecked
    ChkTaxStamp_Click
    Me.ChkTaxSerivce.value = vbUnchecked
    ChkTaxSerivce_Click
    '-----------------------------------------------------------------------------
    StrSQL = "SELECT * FROM Transactions WHERE Transaction_Type=28"
StrSQL = StrSQL & "  AND      BranchId in(" & Current_branchSql & ")"
    If SystemOptions.usertype <> UserAdminAll Then
        'StrSQL = StrSQL & " AND   BranchId=" & Current_branch
    End If

    Set rs = New ADODB.Recordset
    rs.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText

    If Not (rs.BOF Or rs.EOF) Then
        rs.MoveLast
    End If

    Retrive
    Me.TxtModFlg.text = "R"
    Resize_Form Me, TransactionSize
If SystemOptions.HideCost = True Then
LblTotalAll.Visible = False
LblTotal.Visible = False

TxtPrice.Visible = False
       FG.ColHidden(FG.ColIndex("Price")) = True
       FG.ColHidden(FG.ColIndex("Valu")) = True


End If
    If OPEN_NEW_SCREEN = True Then
        Cmd_Click (0)
    End If

    Exit Sub
ErrTrap:
End Sub

Private Sub Form_Unload(Cancel As Integer)
    On Error GoTo ErrTrap
    Dim i As Integer
    RegisterLogInOut Me.Name, ScreenNameArabic, ScreenNameEnglish, , 250

    If rs.State = adStateOpen Then
        If Not (rs.EOF Or rs.BOF) Then
            If rs.EditMode <> adEditNone Then
                rs.CancelUpdate
            End If
        End If

        rs.Close
    End If

    For i = LBound(cSearchDcbo) To UBound(cSearchDcbo)
        Set cSearchDcbo(i) = Nothing
    Next i

    Set rs = Nothing
    Set TTP = Nothing
    NewGrid.Class_Terminate
    Set NewGrid = Nothing
    Set BuyReport = Nothing
    Set m_MnuShowNewItemsPrices = Nothing

    If Not m_FrmSearch Is Nothing Then
        Unload m_FrmSearch
        Set m_FrmSearch = Nothing
    End If

    Exit Sub
ErrTrap:
End Sub

Private Sub TxtModFlg_Change()
    On Error GoTo ErrTrap

    Select Case Me.TxtModFlg.text

        Case "R"

            If SystemOptions.UserInterface = ArabicInterface Then
                Me.Caption = "”‰œ «” ·«„ «‰ «Ã  «„   "
            Else
                Me.Caption = "Production Recieve Voucher     "
            End If
    
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
            XPBtnNewClients.Enabled = False
        
            XPCboDiscountType.locked = True
            Me.XPDtbBill.Enabled = False
            Me.DBCboClientName.locked = True
            Me.DCboStoreName.locked = True
            Me.XPTxtDiscountVal.locked = True
        
            XPChkPayType(0).Enabled = False
            XPChkPayType(1).Enabled = False
            XPChkPayType(2).Enabled = False
        
            XPTxtValue(0).Enabled = False
            XPTxtSerial(0).Enabled = False
            XPTxtValue(1).Enabled = False
            XPTxtSerial(1).Enabled = False
            FG.Editable = flexEDKbdMouse
        '    FG.Editable = flexEDNone

            If rs.RecordCount < 1 Then
                Me.XPBtnMove(0).Enabled = False
                Me.XPBtnMove(1).Enabled = False
                Me.XPBtnMove(2).Enabled = False
                Me.XPBtnMove(3).Enabled = False
                Me.Cmd(1).Enabled = False
                Me.Cmd(4).Enabled = False
                Me.Cmd(5).Enabled = False
                Me.Cmd(7).Enabled = False
            
            End If
        
            CboPayMentType.locked = True
            DtpDelayDate.Enabled = False
            Ele(4).Enabled = False
        
            XPChkTAX.Enabled = False
            ChkTaxAdd.Enabled = False
            ChkTaxSerivce.Enabled = False
            ChkTaxStamp.Enabled = False
        
        Case "N"
            '      Me.Caption = "«–‰ «÷«›… ( ÃœÌœ )"
            Me.Cmd(2).Enabled = True
            Me.Cmd(3).Enabled = True
        
            Me.Cmd(0).Enabled = False
            Me.Cmd(1).Enabled = False
            Me.Cmd(4).Enabled = False
            Me.Cmd(5).Enabled = False
            Me.Cmd(7).Enabled = False
        
            '   Me.XPBtnMove(0).Enabled = False
            '   Me.XPBtnMove(1).Enabled = False
            '   Me.XPBtnMove(2).Enabled = False
            '   Me.XPBtnMove(3).Enabled = False
            XPBtnNewClients.Enabled = True
            FG.Enabled = True
            FG.rows = 2
            XPCboDiscountType.locked = False
            Me.XPDtbBill.Enabled = True
            Me.DBCboClientName.locked = False
            Me.DCboStoreName.locked = False
            Me.XPTxtDiscountVal.locked = False
            XPChkPayType(0).Enabled = True
            XPChkPayType(1).Enabled = True
            XPChkPayType(2).Enabled = True
            XPChkPayType(0).value = Unchecked
            XPChkPayType(1).value = Unchecked
            XPChkPayType(2).value = Unchecked
            FG.Editable = flexEDKbdMouse
            XPDtbBill.value = Date
            '        XPFillData.Enabled = True
            XPCboDiscountType.ListIndex = 0
            CboPayMentType.ListIndex = 0
            CboPayMentType.locked = False
            DtpDelayDate.Enabled = True
            DtpDelayDate.value = Date
            Ele(4).Enabled = True
        
            CboItemCase.ListIndex = 0
        
            XPChkTAX.Enabled = True
            ChkTaxAdd.Enabled = True
            ChkTaxSerivce.Enabled = True
            ChkTaxStamp.Enabled = True

        Case "E"
            '      Me.Caption = "«–‰ «÷«›… (  ⁄œÌ· )"
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
            XPBtnNewClients.Enabled = True
        
            FG.Enabled = True
            XPCboDiscountType.locked = False
            Me.XPDtbBill.Enabled = True
            Me.DBCboClientName.locked = False
            Me.DCboStoreName.locked = False
            Me.XPTxtDiscountVal.locked = False
        
            XPChkPayType(0).Enabled = True
            XPChkPayType(1).Enabled = True
            XPChkPayType(2).Enabled = True
            DtpDelayDate.Enabled = True
        
            If XPChkPayType(0).value = Checked Then
                XPChkPayType_Click (0)
            End If

            If XPChkPayType(1).value = Checked Then
                XPChkPayType_Click (1)
            End If

            If XPChkPayType(2).value = Checked Then
                XPChkPayType_Click (2)
            End If

            If CboPayMentType.ListIndex = 0 Then
                CboPayMentType_Change
            End If

            FG.Editable = flexEDKbdMouse
        
            CboPayMentType.locked = False
            DBCboClientName_Change
            Ele(4).Enabled = True
            XPChkTAX.Enabled = True
            ChkTaxAdd.Enabled = True
            ChkTaxSerivce.Enabled = True
            ChkTaxStamp.Enabled = True
    End Select

    Exit Sub
ErrTrap:
    Stop
End Sub

Public Sub Retrive(Optional Lngid As Long = 0)
    Dim RsDetails As ADODB.Recordset
    Dim StrSQL As String
    Dim RsNotes As ADODB.Recordset
    Dim RsTemp As ADODB.Recordset
    Dim RsTest As ADODB.Recordset
    Dim Num As Long
    Dim Msg As String
    Dim i As Integer
    Dim LngPartID As Long
    Dim RsPartDetails As ADODB.Recordset
    'Dim rs As ADODB.Record
    On Error GoTo ErrTrap
    '---------------------------------------------
    'Here We Reset all Setting
    Me.CmdNotes.Visible = False
    Me.CmdNotes.Tag = ""
    Me.CmdRetruns.Visible = False
    Me.CmdRetruns.Tag = ""
    ChkTaxAdd.value = vbUnchecked
    Me.TxtTaxAddValue.text = ""
    ChkTaxStamp.value = vbUnchecked
    Me.TxtTaxStampValue.text = ""
    ChkTaxStamp.value = vbUnchecked
    Me.TxtTaxStampValue.text = ""
    ChkTaxSerivce.value = vbUnchecked
    Me.TxtTaxServiceValue.text = ""

    '---------------------------------------------
    '---------------------------------------------
    If rs.RecordCount < 1 Then
        XPTxtCurrent.Caption = 0
        XPTxtCount.Caption = 0
        Exit Sub
    End If

    If rs.EOF Or rs.BOF Then
        Exit Sub
    End If

    If Lngid <> 0 Then
        rs.Find "Transaction_ID=" & Lngid, , adSearchForward, adBookmarkFirst

        If rs.BOF Or rs.EOF Then
            Exit Sub
        End If
    End If

    TxtFillData.text = "T"
    Screen.MousePointer = vbArrowHourglass
    dcBranch.BoundText = IIf(IsNull(rs("BranchId").value), "", rs("BranchId").value)
    Me.TXTOrderNO1.text = IIf(IsNull(rs("WorkOrderNO").value), "", (rs("WorkOrderNO").value))

    Me.TxtNoteSerial.text = IIf(IsNull(rs("NoteSerial").value), "", (rs("NoteSerial").value))
    Me.TxtNoteSerial1.text = IIf(IsNull(rs("NoteSerial1").value), "", (rs("NoteSerial1").value))
    Me.oldtxtNoteSerial1.text = IIf(IsNull(rs("OldNoteSerial1").value), IIf(IsNull(rs("NoteSerial1").value), "", rs("NoteSerial1").value), rs("OldNoteSerial1").value)

TxtBillComment.text = IIf(IsNull(rs("TransactionComment").value), "", (rs("TransactionComment").value))

TxtInspectionReport.text = IIf(IsNull(rs("InspectionReport").value), "", (rs("InspectionReport").value))

    lbl(62).Caption = showLabel(TxtNoteSerial1, oldtxtNoteSerial1)

    Me.TXTNoteID.text = IIf(IsNull(rs("NoteID").value), "", (rs("NoteID").value))
    XPTxtBillID.text = IIf(IsNull(rs("Transaction_ID").value), "", (rs("Transaction_ID").value))
    TxtTransSerial.text = IIf(IsNull(rs("Transaction_Serial").value), "", (rs("Transaction_Serial").value))
    XPDtbBill.value = IIf(IsNull(rs("Transaction_Date").value), "", (rs("Transaction_Date").value))
    XPCboDiscountType.ListIndex = IIf(IsNull(rs("Trans_DiscountType").value), 0, rs("Trans_DiscountType").value)
    cmbTypeTrans.ListIndex = IIf(IsNull(rs("TypeTrans").value), 0, rs("TypeTrans").value)
    XPTxtDiscountVal.text = IIf(IsNull(rs("Trans_Discount").value), "", Trim(rs("Trans_Discount").value))
    CboPayMentType.ListIndex = IIf(IsNull(rs("PaymentType").value), 0, rs("PaymentType").value)
    Me.DBCboClientName.BoundText = IIf(IsNull(rs("CusID").value), "", rs("CusID").value)
    Me.DCboUserName.BoundText = IIf(IsNull(rs("UserID").value), "", rs("UserID").value)
    Me.DCboStoreName.BoundText = IIf(IsNull(rs("StoreID").value), "", rs("StoreID").value)
    '÷—»Ì… «·„»Ì⁄« 
    XPTxtTaxValue.text = IIf(IsNull(rs("TaxValue").value), "", (rs("TaxValue").value))
    XPChkTAX.value = IIf(rs("TaxFound") = True, Checked, Unchecked)
    Dim Myrec As New ADODB.Recordset
    Dim Mytotal As Integer
    Dim MySQL As String

    MySQL = "SELECT Sum (Notes.Note_Value) AS [TotalRevenue] FROM Notes where NumOrderInpot = " & val(TxtTransSerial)
    Set RsNotes = New ADODB.Recordset
    RsNotes.Open MySQL, Cn, adOpenStatic, adLockOptimistic, adCmdText

    If Not (RsNotes.BOF Or RsNotes.EOF) Then
        Txt_EXport.text = IIf(IsNull(RsNotes("TotalRevenue").value), "", (RsNotes("TotalRevenue").value))
    End If

    Text1.text = IIf(IsNull(rs("nots").value), "", (rs("nots").value))
    'Txt_EXport.text = IIf(IsNull(rs("Shahne").Value), "", (rs("Shahne").Value))

    '÷—»Ì… «·Œ’„ Ê«·≈÷«›…
    If Not IsNull(rs("TaxAddValue").value) Then
        If rs("TaxAddValue").value > 0 Then
            ChkTaxAdd.value = vbChecked
            Me.TxtTaxAddValue.text = rs("TaxAddValue").value
        End If
    End If

    '÷—»Ì… «·œ„€…
    If Not IsNull(rs("TaxStampValue").value) Then
        If rs("TaxStampValue").value > 0 Then
            ChkTaxStamp.value = vbChecked
            Me.TxtTaxStampValue.text = rs("TaxStampValue").value
        End If
    End If

    '÷—»Ì… «·Œœ„…
    If Not IsNull(rs("TaxServiceValue").value) Then
        If rs("TaxServiceValue").value > 0 Then
            ChkTaxSerivce.value = vbChecked
            Me.TxtTaxServiceValue.text = rs("TaxServiceValue").value
        End If
    End If

    FG.rows = 2
    FG.Clear flexClearScrollable, flexClearEverything
    XPTxtSum.text = ""

    StrSQL = "SELECT TblItems.HaveSerial, * FROM TblItems INNER JOIN Transaction_Details " & "ON TblItems.ItemID = Transaction_Details.Item_ID INNER JOIN dbo.TblUnites ON dbo.Transaction_Details.UnitID = dbo.TblUnites.UnitID"
    StrSQL = StrSQL + "  where Transaction_ID=" & val(rs("Transaction_ID").value)

    Set RsDetails = New ADODB.Recordset
    RsDetails.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText

    If Not (RsDetails.EOF Or RsDetails.BOF) Then
        FG.rows = RsDetails.RecordCount + 1

        For Num = 1 To RsDetails.RecordCount
            FG.TextMatrix(Num, FG.ColIndex("NoSerial")) = IIf(IsNull(RsDetails("NoSerial")), "", (RsDetails("NoSerial").value))
            FG.TextMatrix(Num, FG.ColIndex("FoxyNo")) = IIf(IsNull(RsDetails("FoxyNo")), "", (RsDetails("FoxyNo").value))
            FG.TextMatrix(Num, FG.ColIndex("Code")) = IIf(IsNull(RsDetails("Item_ID")), "", (RsDetails("Item_ID").value))
            FG.TextMatrix(Num, FG.ColIndex("Name")) = IIf(IsNull(RsDetails("Item_ID")), "", Trim(RsDetails("Item_ID").value))
            FG.TextMatrix(Num, FG.ColIndex("Count")) = IIf(IsNull(RsDetails("showQty")), "", (RsDetails("showQty").value))
            FG.TextMatrix(Num, FG.ColIndex("Serial")) = IIf(IsNull(RsDetails("ItemSerial")), "", Trim(RsDetails("ItemSerial").value))
            FG.TextMatrix(Num, FG.ColIndex("DamageshowQty")) = IIf(IsNull(RsDetails("DamageshowQty")), "", (RsDetails("DamageshowQty").value))
        
        FG.TextMatrix(Num, FG.ColIndex("NProductionOrderNO")) = IIf(IsNull(RsDetails("NProductionOrderNO")), "", Trim(RsDetails("NProductionOrderNO").value))
        
            If RsDetails("HaveSerial") = True Then
                FG.TextMatrix(Num, FG.ColIndex("HaveSerial")) = True
            End If

            FG.TextMatrix(Num, FG.ColIndex("Price")) = IIf(IsNull(RsDetails("showPrice")), "", (RsDetails("showPrice").value))

            If FG.TextMatrix(Num, FG.ColIndex("Price")) = "" Then
                FG.TextMatrix(Num, FG.ColIndex("Price")) = IIf(IsNull(RsDetails("Price")), "", (RsDetails("Price").value))
            End If

            If SystemOptions.SysDataBaseType = SQLServerDataBase Then
                FG.TextMatrix(Num, FG.ColIndex("ItemCase")) = IIf(IsNull(RsDetails("ItemCase")), "", (RsDetails("ItemCase").value))
            Else
                FG.TextMatrix(Num, FG.ColIndex("ItemCase")) = IIf(IsNull(RsDetails("Transaction_Details.ItemCase")), "", (RsDetails("Transaction_Details.ItemCase").value))
            End If

            FG.TextMatrix(Num, FG.ColIndex("DiscountType")) = IIf(IsNull(RsDetails("ItemDiscountType")), "", (RsDetails("ItemDiscountType").value))
            FG.TextMatrix(Num, FG.ColIndex("DiscountVal")) = IIf(IsNull(RsDetails("ItemDiscount")), "", (RsDetails("ItemDiscount").value))
            FG.TextMatrix(Num, FG.ColIndex("guaranteeTime")) = IIf(IsNull(RsDetails("guaranteeTime")), "", (RsDetails("guaranteeTime").value))
            FG.TextMatrix(Num, FG.ColIndex("Remarks")) = IIf(IsNull(RsDetails("Remarks")), "", (RsDetails("Remarks").value))
        
            FG.TextMatrix(Num, FG.ColIndex("ColorID")) = IIf(IsNull(RsDetails("ColorID")), 1, (RsDetails("ColorID").value))
            FG.TextMatrix(Num, FG.ColIndex("ItemSize")) = IIf(IsNull(RsDetails("ItemSize")), 1, (RsDetails("ItemSize").value))
        
            FG.cell(flexcpData, Num, FG.ColIndex("UnitID")) = IIf(IsNull(RsDetails("UnitID")), "", (RsDetails("UnitID").value))
            FG.TextMatrix(Num, FG.ColIndex("UnitID")) = IIf(IsNull(RsDetails("UnitName")), "", (RsDetails("UnitName").value))
            FG.TextMatrix(Num, FG.ColIndex("order_no")) = IIf(IsNull(RsDetails("order_no").value), "", RsDetails("order_no").value)
        
            FG.TextMatrix(Num, FG.ColIndex("ColorID")) = IIf(IsNull(RsDetails("ColorID")), 1, (RsDetails("ColorID").value))
            FG.TextMatrix(Num, FG.ColIndex("ItemSize")) = IIf(IsNull(RsDetails("ItemSize")), 1, (RsDetails("ItemSize").value))
            FG.TextMatrix(Num, FG.ColIndex("ClassID")) = IIf(IsNull(RsDetails("ClassID")), 1, (RsDetails("ClassID").value))
            FG.TextMatrix(Num, FG.ColIndex("OrderArrivalDate")) = IIf(IsNull(RsDetails("OrderArrivalDate").value), "", RsDetails("OrderArrivalDate").value)
            FG.TextMatrix(Num, FG.ColIndex("order_no")) = IIf(IsNull(RsDetails("order_no").value), "", RsDetails("order_no").value)
            FG.TextMatrix(Num, FG.ColIndex("FoxyNo")) = IIf(IsNull(RsDetails("FoxyNo")), "", RsDetails("FoxyNo").value)
        
        
                       FG.TextMatrix(Num, FG.ColIndex("Remarks")) = IIf(IsNull(RsDetails("Remarks")), "", (RsDetails("Remarks").value))
                    FG.TextMatrix(Num, FG.ColIndex("L")) = IIf(IsNull(RsDetails("L")), "", (RsDetails("L").value))
                    FG.TextMatrix(Num, FG.ColIndex("W")) = IIf(IsNull(RsDetails("W")), "", (RsDetails("W").value))
                    FG.TextMatrix(Num, FG.ColIndex("OldID")) = IIf(IsNull(RsDetails("OldID")), "", (RsDetails("OldID").value))
                    FG.TextMatrix(Num, FG.ColIndex("H1")) = IIf(IsNull(RsDetails("H1")), "", (RsDetails("H1").value))
                    FG.TextMatrix(Num, FG.ColIndex("H2")) = IIf(IsNull(RsDetails("H2")), "", (RsDetails("H2").value))
                    FG.TextMatrix(Num, FG.ColIndex("NoCount")) = IIf(IsNull(RsDetails("NoCount")), "", (RsDetails("NoCount").value))
                    FG.TextMatrix(Num, FG.ColIndex("Area")) = IIf(IsNull(RsDetails("Area")), "", (RsDetails("Area").value))
                    FG.TextMatrix(Num, FG.ColIndex("Height")) = IIf(IsNull(RsDetails("Height")), "", (RsDetails("Height").value))
                    FG.TextMatrix(Num, FG.ColIndex("length")) = IIf(IsNull(RsDetails("length")), "", (RsDetails("length").value))
                    FG.TextMatrix(Num, FG.ColIndex("Width")) = IIf(IsNull(RsDetails("Width")), "", (RsDetails("Width").value))
                    FG.TextMatrix(Num, FG.ColIndex("ColorID")) = IIf(IsNull(RsDetails("ColorID")), 1, (RsDetails("ColorID").value))
                    FG.TextMatrix(Num, FG.ColIndex("ItemSize")) = IIf(IsNull(RsDetails("ItemSize")), 1, (RsDetails("ItemSize").value))
                    FG.TextMatrix(Num, FG.ColIndex("ClassID")) = IIf(IsNull(RsDetails("ClassID")), 1, (RsDetails("ClassID").value))
            RsDetails.MoveNext

            If FG.rows > 10 Then
                If Num = 8 Then FG.Refresh
            End If

        Next Num

        FG.AutoSize 0, FG.Cols - 1, False
    End If

    XPChkPayType(0).value = Unchecked
    XPChkPayType(1).value = Unchecked
    XPChkPayType(2).value = Unchecked
    XPTxtValue(0).text = ""
    XPTxtValue(1).text = ""

    XPTxtSerial(0).text = ""
    XPTxtSerial(1).text = ""
    DtpDelayDate.value = Date
    StrSQL = "select * From Notes where Transaction_ID=" & val(rs("Transaction_ID").value)
    Set RsNotes = New ADODB.Recordset
    RsNotes.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText

    If Not (RsNotes.EOF Or RsNotes.BOF) Then

        For Num = 1 To RsNotes.RecordCount

            If RsNotes("NoteType").value = 0 Then
                XPChkPayType(0).value = Checked
                XPChkPayType_Click (0)
                'Me.TxtNoteID(0).text = IIf(IsNull(RsNotes("NOTEID").Value), "", (RsNotes("NOTEID").Value))
                XPTxtValue(0).text = IIf(IsNull(RsNotes("Note_Value").value), "", (RsNotes("Note_Value").value))
                XPTxtSerial(0).text = IIf(IsNull(RsNotes("NoteSerial").value), "", Trim(RsNotes("NoteSerial").value))
                Me.DcboBox.BoundText = IIf(IsNull(RsNotes("BoxID").value), "", RsNotes("BoxID").value)
            End If

            If RsNotes("NoteType").value = 1 Then
                XPChkPayType(1).value = Checked
                XPChkPayType_Click (1)
                'Me.TxtNoteID(1).text = IIf(IsNull(RsNotes("NOTEID").Value), "", (RsNotes("NOTEID").Value))
                XPTxtValue(1).text = IIf(IsNull(RsNotes("Note_Value").value), "", (RsNotes("Note_Value").value))
                XPTxtValue(1).Tag = IIf(IsNull(RsNotes("NoteID").value), "", (RsNotes("NoteID").value))
                XPTxtSerial(1).text = IIf(IsNull(RsNotes("NoteSerial").value), "", Trim(RsNotes("NoteSerial").value))
                DtpDelayDate.value = IIf(IsNull(RsNotes("DueDate").value), "", (RsNotes("DueDate").value))
            End If

            If RsNotes("NoteType").value = 13 Then
                XPChkPayType(2).value = Checked
                XPChkPayType_Click (2)
            End If
        
            RsNotes.MoveNext
        Next Num

    End If

    Set RsNotes = New ADODB.Recordset
    StrSQL = "SELECT Notes.NoteID, Notes.NoteDate, Notes.NoteType, Notes.NoteSerial," & "Notes.Note_Value, Notes.BankID,BanksData.BankName , Notes.ChqueNum, Notes.DueDate "
    StrSQL = StrSQL + " FROM Notes INNER JOIN BanksData ON Notes.BankID = BanksData.BankID "
    StrSQL = StrSQL + " Where NoteType=13 AND NOTES.Transaction_ID=" & val(rs("Transaction_ID").value)
    StrSQL = StrSQL + " Order BY Notes.NoteID"
    RsNotes.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

    With Me.FgCheques
        .rows = .FixedRows

        If Not (RsNotes.BOF Or RsNotes.EOF) Then
            .rows = .FixedRows + RsNotes.RecordCount

            For i = .FixedRows To .rows - 1
                .TextMatrix(i, .ColIndex("CheckValue")) = IIf(IsNull(RsNotes("Note_Value").value), "", RsNotes("Note_Value").value)
                .TextMatrix(i, .ColIndex("CheckNumber")) = IIf(IsNull(RsNotes("ChqueNum").value), "", RsNotes("ChqueNum").value)
                .TextMatrix(i, .ColIndex("BankID")) = IIf(IsNull(RsNotes("BankID").value), "", RsNotes("BankID").value)
                .TextMatrix(i, .ColIndex("BankName")) = IIf(IsNull(RsNotes("BankName").value), "", RsNotes("BankName").value)

                If Not IsNull(RsNotes("DueDate").value) Then
                    .TextMatrix(i, .ColIndex("DueDate")) = DisplayDate(RsNotes("DueDate").value)
                Else
                    .TextMatrix(i, .ColIndex("DueDate")) = ""
                End If

                RsNotes.MoveNext
            Next i

        End If

        .AutoSize 0, .Cols - 1, False
        SumChecks
    End With

    '⁄—÷ «·√ﬁ”«ÿ «·„”Ã·… ⁄·Ï «·›« Ê—…
    If XPTxtValue(1).Tag <> "" Then
        StrSQL = "Select * From InstallMent where NoteID=" & XPTxtValue(1).Tag
        Set RsTest = New ADODB.Recordset
        RsTest.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

        If Not (RsTest.EOF Or RsTest.BOF) Then
            CmdINSTALLMENT.Enabled = True
            CmdINSTALLMENT.Caption = "⁄—÷ «·√ﬁ”«ÿ «·„”Ã·…"
            LngPartID = RsTest("PartID").value
            Me.LblPrecenType.Tag = RsTest("InterestType").value

            If RsTest("InterestType").value = 0 Then
                LblPrecenType.Caption = "‰”»… „∆ÊÌ…"
            ElseIf RsTest("InterestType").value = 1 Then
                LblPrecenType.Caption = "ﬁÌ„… À«» …"
            ElseIf RsTest("InterestType").value = 2 Then
                LblPrecenType.Caption = "·«ÌÊÃœ"
            End If

            Me.LblPrecenValue.Caption = RsTest("InterestVal").value
            Me.LblInstallTotal.Caption = RsTest("Total").value
            Me.LblInstallCount.Caption = RsTest("InstallCount").value
            Me.LblFirstInstallDate.Caption = DisplayDate(RsTest("FirstInstallDate").value)
            Me.LblInstallmentType.Tag = RsTest("InstallmentType").value

            If RsTest("InstallmentType").value = 0 Then
                LblInstallmentType.Caption = "ÌÊ„"
            ElseIf RsTest("InstallmentType").value = 1 Then
                LblInstallmentType.Caption = "‘Â—"
            ElseIf RsTest("InstallmentType").value = 2 Then
                LblInstallmentType.Caption = "”‰…"
            End If

            Me.LblInstallSeprator.Caption = RsTest("InstallSeprator").value
            Me.LblStartValue.Caption = IIf(IsNull(RsTest("StartValue").value), "", RsTest("StartValue").value)
            Set RsPartDetails = New ADODB.Recordset
            StrSQL = "Select * From InstallMentDetails Where PartID=" & LngPartID
            RsPartDetails.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

            If Not (RsPartDetails.BOF Or RsPartDetails.EOF) Then
                RsPartDetails.MoveFirst

                With Me.FgInstallments
                    .rows = .FixedRows + RsPartDetails.RecordCount

                    For i = .FixedRows To .rows - 1
                        .TextMatrix(i, .ColIndex("QestID")) = IIf(IsNull(RsPartDetails("QestID").value), "", RsPartDetails("QestID").value)
                        .TextMatrix(i, .ColIndex("Value")) = IIf(IsNull(RsPartDetails("Value").value), "", RsPartDetails("Value").value)

                        If Not IsNull(RsPartDetails("DueDate").value) Then
                            .TextMatrix(i, .ColIndex("Due_Date")) = DisplayDate(RsPartDetails("DueDate").value)
                        Else
                            .TextMatrix(i, .ColIndex("Due_Date")) = ""
                        End If
 
                        RsPartDetails.MoveNext
                    Next i

                End With

            End If

        Else
            CmdINSTALLMENT.Enabled = False
            CmdINSTALLMENT.Caption = " ﬁ”Ìÿ «·ﬁÌ„… «·¬Ã·…"
        End If
    End If

    NewGrid.Calculate 1, , , True
    Dim SngRelatedNotesValues As Single
    Me.CmdNotes.Visible = ShowRelatedNotes(val(Me.XPTxtBillID.text), 0, SngRelatedNotesValues)
    Me.CmdNotes.Tag = SngRelatedNotesValues

    SngRelatedNotesValues = 0
    Me.CmdRetruns.Visible = ShowRelatedTransactions(val(Me.XPTxtBillID.text), 0, SngRelatedNotesValues)
    Me.CmdRetruns.Tag = SngRelatedNotesValues
    '-----------------------------------------------------------------------------------------------
    Screen.MousePointer = vbDefault
    TxtFillData.text = "F"
    XPTxtCurrent.Caption = rs.AbsolutePosition
    XPTxtCount.Caption = rs.RecordCount
    Exit Sub
ErrTrap:
    Msg = "Œÿ« ›Ï ≈” —Ã«⁄ «·»Ì«‰« ..!!!"
    Msg = Msg & CHR(13) & Err.Description
    Msg = Msg & CHR(13) & Err.Number
    Msg = Msg & CHR(13) & Err.Source
    MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
    Screen.MousePointer = vbDefault
End Sub

Private Sub Undo()
    Dim Msg As String

    On Error GoTo ErrTrap

    Select Case TxtModFlg.text

        Case "N"
            Msg = "”Ê› Ì „ «· —«Ã⁄ ›Ï  ”ÃÌ· Â–Â «·›« Ê—… .."
            Msg = Msg & CHR(13) & "›Â· «‰  „ «ﬂœ „‰ «·√” „—«— ..!!"

            If MsgBox(Msg, vbQuestion + vbYesNo + vbMsgBoxRight + vbMsgBoxRtlReading + vbDefaultButton2, App.Title) = vbYes Then
                clear_all Me
                Me.TxtModFlg.text = "R"
                XPBtnMove_Click (1)
            End If

        Case "E"
            Msg = "”Ê› Ì „ «· —«Ã⁄ ›Ï  ⁄œÌ· Â–Â «·›« Ê—… .."
            Msg = Msg & CHR(13) & "›Â· «‰  „ «ﬂœ „‰ «·√” „—«— ..!!"

            If MsgBox(Msg, vbQuestion + vbYesNo + vbMsgBoxRight + vbMsgBoxRtlReading + vbDefaultButton2, App.Title) = vbYes Then
                rs.Find "Transaction_ID='" & val(XPTxtBillID.text) & "'", , adSearchForward, adBookmarkFirst

                If rs.EOF Or rs.BOF Then
                    Me.TxtModFlg.text = "R"
                    Exit Sub
                End If

                If Not rs.EOF Or rs.BOF Then
                    Me.TxtModFlg.text = "R"
                    Retrive
                End If
            End If

    End Select

    Exit Sub
ErrTrap:
End Sub

Private Sub Del_TransAction()
    Dim Msg As String
    Dim StrSQL As String
    Dim BegainTrans As Boolean
    On Error GoTo ErrTrap

    If XPTxtBillID.text <> "" Then
        Msg = "”Ì „ Õ–› »Ì«‰«  «·⁄„·Ì… —ﬁ„ " & CHR(13)
        Msg = Msg + (XPTxtBillID.text) & CHR(13)
        Msg = Msg + " Â·  —€» ›Ì Õ–› Â–Â «·»Ì«‰« ø"

        If MsgBox(Msg, vbYesNo + vbQuestion + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title) = vbYes Then

            '        If AvailableDeal = True Then
            If Not rs.RecordCount < 1 Then
                Cn.BeginTrans
                BegainTrans = True
                StrSQL = "Delete From DOUBLE_ENTREY_VOUCHERS  " & "Where DOUBLE_ENTREY_VOUCHERS.Transaction_ID=" & rs("Transaction_ID").value
                Cn.Execute StrSQL, , adExecuteNoRecords
                
                StrSQL = "delete From Notes where noteid=" & val(TXTNoteID.text)
                Cn.Execute StrSQL, , adExecuteNoRecords
                CuurentLogdata ("D")
                rs.delete
                Cn.CommitTrans
                BegainTrans = False
                rs.MoveFirst

                If rs.RecordCount < 1 Then
                    clear_all Me
                    TxtModFlg_Change
                    XPTxtCurrent.Caption = 0
                    XPTxtCount.Caption = 0
                Else
                    Retrive
                End If
            End If

            '        End If
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
    Msg = "·«Ì„ﬂ‰ Õ–› Â–« «·”Ã· · ﬂ«„· «·»Ì«‰«  " & CHR(13) & " ÊÃœ »Ì«‰«  „— »ÿ… »Â–« «·”Ã· "
    Msg = Msg & CHR(13) & Err.Number
    Msg = Msg & CHR(13) & Err.Description
    Msg = Msg & CHR(13) & Err.Source
    MsgBox Msg, vbMsgBoxRight + vbMsgBoxRtlReading + vbExclamation, App.Title

    If BegainTrans = True Then
        rs.CancelUpdate
        Cn.RollbackTrans
        BegainTrans = False
    End If

End Sub

Private Sub AddTip()
    Dim Wrap As String
    Set TTP = New clstooltip
    On Error GoTo ErrTrap
    Wrap = CHR(13) + CHR(10)

    With TTP
        .Create Me.hWnd, "»Ì«‰«  «–‰ «·«÷«›…", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
        .AddControl Cmd(0), "ÃœÌœ ..." & Wrap & "·«÷«›… »Ì«‰«  ⁄„·Ì… «–‰ «÷«›… ÃœÌœ…" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«" & Wrap & "„›« ÌÕ «·«Œ ’«— F12 OR Enter", True
    End With

    With TTP
        .Create Me.hWnd, "»Ì«‰«  «–‰ «÷«›…", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
        .AddControl Cmd(7), "ÿ»«⁄… ..." & Wrap & "·⁄—÷ «·»Ì«‰«  «·Õ«·Ì… ›Ì  ﬁ—Ì— " & Wrap & " Ì„ﬂ‰ ÿ»«⁄ Â ⁄‰ ÿ—Ìﬁ «·ÿ«»⁄…" & Wrap & "„›« ÌÕ «·«Œ ’«— F6", True
    End With

    With TTP
        .Create Me.hWnd, "»Ì«‰«  «–‰ «÷«›…", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
        .AddControl Cmd(1), " ⁄œÌ· ..." & Wrap & "· ⁄œÌ· »Ì«‰«  «–‰ «·«÷«›…" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«" & Wrap & "„›« ÌÕ «·«Œ ’«— F11", True
    End With

    With TTP
        .Create Me.hWnd, "»Ì«‰«  «–‰ «÷«›…", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
        .AddControl Cmd(2), "Õ›Ÿ ..." & Wrap & "·Õ›Ÿ »Ì«‰«  ⁄„·Ì… «–‰ «·«÷«›…" & Wrap & "·Õ›Ÿ «· ⁄œÌ·« " & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«" & Wrap & "„›« ÌÕ «·«Œ ’«— F10", True
    End With

    With TTP
        .Create Me.hWnd, "»Ì«‰«  «–‰ «·«÷«›…", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
        .AddControl Cmd(3), " —«Ã⁄ ..." & Wrap & "·· —«Ã⁄ ⁄‰ ⁄„·Ì… «–‰ «·«÷«›…" & Wrap & "··· —«Ã⁄ ⁄‰ ⁄„·Ì… «· ⁄œÌ·" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«" & Wrap & "„›« ÌÕ «·«Œ ’«— F9", True
    End With

    With TTP
        .Create Me.hWnd, "»Ì«‰«  ⁄„·Ì«  «–‰ «·«÷«›…", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
        .AddControl Cmd(4), "Õ–› ..." & Wrap & "·Õ–› »Ì«‰«  ⁄„·Ì… «–‰ «·«÷«›…" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«" & Wrap & "„›« ÌÕ «·«Œ ’«— F8", True
    End With

    With TTP
        .Create Me.hWnd, "»Ì«‰«  «–‰ «·«÷«›…", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
        .AddControl Cmd(5), "»ÕÀ ..." & Wrap & "···»ÕÀ ⁄‰ ⁄„·Ì… ‘—«¡" & Wrap & "Ì‰ÿ»ﬁ ⁄·ÌÂ« ‘—Êÿ „⁄Ì‰…" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«" & Wrap & "„›« ÌÕ «·«Œ ’«— F7", True
    End With

    With TTP
        .Create Me.hWnd, "»Ì«‰«  «–‰ «·«÷«›…", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
        .AddControl Cmd(6), "Œ—ÊÃ ..." & Wrap & "·«€·«ﬁ Â–Â «·‰«›–…" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«" & Wrap & "„›« ÌÕ «·«Œ ’«— Ctrl + X", True
    End With

    'With TTP
    '   .Create Me.hwnd, "»Ì«‰«  ›« Ê—… «·‘—«¡", 1, 15204351, -2147483630
    '   .MaxWidth = 4000
    '   .VisibleTime = 9000
    '   .DelayTime = 600
    '   .AddControl XPBtnAdd, _
    '    "≈÷«›… «·√’‰«› ..." & Wrap & _
    '    " ·«÷«›… ’‰› ÃœÌœ" & Wrap & _
    '    " ›ﬁÿ ≈÷€ÿ Â‰«" & Wrap & _
    '    "„›« ÌÕ «·«Œ ’«— F2", True
    'End With
    'With TTP
    '   .Create Me.hwnd, "»Ì«‰«  ›« Ê—… «·‘—«¡", 1, 15204351, -2147483630
    '   .MaxWidth = 4000
    '   .VisibleTime = 9000
    '   .DelayTime = 600
    '   .AddControl XPBtnRemove, _
    '    "Õ–› ’‰› ..." & Wrap & _
    '    "·Õ–› √Õœ «·√’‰«›" & Wrap & _
    '    " ÕœœÂ Ê«÷€ÿ Â‰«" & Wrap & _
    '    "„›« ÌÕ «·«Œ ’«— F3", True
    'End With
    With TTP
        .Create Me.hWnd, "»Ì«‰«  «–‰ «·«÷«›…", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
        .AddControl XPBtnNewClients, "≈÷«›… ⁄„Ì· ÃœÌœ ..." & Wrap & "· ”ÃÌ· »Ì«‰«  ⁄„Ì· ÃœÌœ" & Wrap & " «÷€ÿ Â‰«" & Wrap & "„›« ÌÕ «·«Œ ’«— F5", True
    End With

    With TTP
        .Create Me.hWnd, "»Ì«‰«  «–‰ «·«÷«›…", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
        .AddControl XPBtnMove(1), "«·√Ê· ..." & Wrap & "··«‰ ﬁ«· «·Ï √Ê· ”Ã·" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
    End With

    With TTP
        .Create Me.hWnd, "»Ì«‰«  «–‰ «·«÷«›…", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
        .AddControl XPBtnMove(0), "«·”«»ﬁ ..." & Wrap & "··«‰ ﬁ«· «·Ï «·”Ã· «·”«»ﬁ" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
    End With

    With TTP
        .Create Me.hWnd, "»Ì«‰«  «–‰ «·«÷«›…", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
        .AddControl XPBtnMove(3), "«· «·Ì ..." & Wrap & "··«‰ ﬁ«· «·Ï «·”Ã· «· «·Ì" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
    End With

    With TTP
        .Create Me.hWnd, "»Ì«‰«  «–‰ «·«÷«›…", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
        .AddControl XPBtnMove(2), "«·√ŒÌ— ..." & Wrap & "··«‰ ﬁ«· «·Ï ¬Œ— ”Ã·" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
    End With

    'With TTP
    '   .Create Me.hwnd, "»Ì«‰«  ›« Ê—… «·‘—«¡", 1, 15204351, -2147483630
    '   .MaxWidth = 4000
    '   .VisibleTime = 9000
    '   .DelayTime = 600
    '   .AddControl XPFillData, _
    '    " ⁄»∆… »Ì«‰«  «·√’‰«›" & Wrap & _
    '    "· ⁄»∆… »Ì«‰«  «·√’‰«› ›Ì" & Wrap & _
    '    "›Ì ‰«›–… ÕÊ«—" & Wrap & _
    '    "  ≈÷€ÿ Â‰«" & Wrap & _
    '    "„›« ÌÕ «·«Œ ’«— Ctrl + Space", True
    'End With
    With TTP
        .Create Me.hWnd, "»Ì«‰«  «–‰ «·«÷«›…", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
        .AddControl CmdHelp, "„”«⁄œ… ..." & Wrap & "·· ⁄—› ⁄·Ï ÊŸÌ›… Â–Â «·‰«›–…" & Wrap & "ÊﬂÌ›Ì… «· ⁄«„· „⁄Â«" & Wrap & "≈÷€ÿ Â‰«" & Wrap, True
    End With

    Exit Sub
ErrTrap:
End Sub

Private Sub SaveData()
    Dim RSTransDetails     As ADODB.Recordset
    Dim RsNotes            As ADODB.Recordset
    Dim RsTemp             As New ADODB.Recordset
    Dim Msg                As String
    Dim Mytot              As String
    Dim RowNum             As Integer
    Dim StrSQL             As String
    Dim StrSqlDel          As String
    Dim SearchResault      As Integer
    Dim note_id            As Long
    Dim RsDetalis          As ADODB.Recordset
    Dim BeginTrans         As Boolean
    Dim LnItemID           As Long
    Dim i                  As Long
    Dim StrCurrentItemName As String
    Dim DblNotesTotal      As Double

    Dim IntLineNO          As Integer
    Dim StrAccountCode     As String
    '****************************
    '· Ã«Â· Õ›Ÿ «· ›«’Ì· „⁄ «⁄«œÂ Ÿ»ÿ «·Õ—ﬂ« 
    Dim mSaveDetails       As Boolean
    mSaveDetails = (IsSaveWithOutMsg And chkIgnorDetails.value = 1) Or Not IsSaveWithOutMsg
    '***********************
    On Error GoTo ErrTrap

    'If Trim(Me.TxtTransSerial.text) = "" Then
    '    Msg = "ÌÃ» ﬂ «»… —ﬁ„ «–‰ «·«÷«›… ..!!!"
    '    MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
    '    Me.TxtTransSerial.SetFocus
    '    Exit Sub
    'End If
    If Me.TxtModFlg.text = "N" Then
        '    If RepeatSerial(Trim(Me.TxtTransSerial.text), 20, 0, Val(Me.DBCboClientName.BoundText)) = True Then
        '        Exit Sub
        '    End If
    ElseIf Me.TxtModFlg.text = "E" Then
        '    If RepeatSerial(Trim(Me.TxtTransSerial.text), 1, Val(Me.XPTxtBillID.text), _
        '        Val(Me.DBCboClientName.BoundText)) = True Then
        '        Exit Sub
        '    End If
    End If

    Screen.MousePointer = vbArrowHourglass
    If Not IsSaveWithOutMsg Then
        If Trim(dcBranch.BoundText) = "" Then
            If SystemOptions.UserInterface = EnglishInterface Then
                Msg = "Specify Departement"
            Else
                Msg = "ÌÃ»  ÕœÌœ «”„    «·›—⁄"
            End If

            MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
            dcBranch.SetFocus
            Sendkeys "{F4}"
            Screen.MousePointer = vbDefault
            Exit Sub
        End If

        '   If DBCboClientName.text = "" Then
        '       Msg = "„‰ ›÷·ﬂ √œŒ· «”„ «·⁄„Ì·"
        '       MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
        '       DBCboClientName.SetFocus
        '       SendKeys "{F4}"
        '       Screen.MousePointer = vbDefault
        '       Exit Sub
        '   End If

        If DCboStoreName.text = "" Then
            Msg = "„‰ ›÷·ﬂ Õœœ «”„ «·„Œ“‰"
            MsgBox Msg, vbOKOnly + vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
            DCboStoreName.SetFocus
            Sendkeys "{F4}"
            Screen.MousePointer = vbDefault
            Exit Sub
        End If

        If XPCboDiscountType.ListIndex = 1 Or XPCboDiscountType.ListIndex = 2 Then
            If XPTxtDiscountVal.text = "" Then
                Msg = "ÌÃ»  ÕœÌœ ﬁÌ„… «·Œ’„ «·ﬂ·Ì ⁄·Ï «·›« Ê—…"
                MsgBox Msg, vbOKOnly + vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
                XPTxtDiscountVal.SetFocus
                Screen.MousePointer = vbDefault
                Exit Sub
            End If

            If Not IsNumeric(XPTxtDiscountVal.text) Then
                Msg = "ﬁÌ„… «·Œ’„ «·ﬂ·Ì ⁄·Ï «·›« Ê—… ÌÃ» √‰  ﬂÊ‰ ﬁÌ„… —ﬁ„Ì…"
                MsgBox Msg, vbOKOnly + vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
                XPTxtDiscountVal.SetFocus
                Screen.MousePointer = vbDefault
                Exit Sub
            End If

            XPTxtDiscountVal.SetFocus
        End If

        '    If CboPayMentType.ListIndex = -1 Then
        '        Msg = "ÌÃ»  ÕœÌœ ÿ—Ìﬁ… «·œ›⁄"
        '        MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
        '        CboPayMentType.SetFocus
        '        SendKeys "{F4}"
        '        Screen.MousePointer = vbDefault
        '        Exit Sub
        '    End If

        '   If XPChkPayType(0).value = vbChecked Then
        '       If Me.DcboBox.BoundText = "" Then
        '           Msg = "ÌÃ»  ÕœÌœ «”„ «·Œ“‰…...!!!"
        '           MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
        '           Screen.MousePointer = vbDefault
        '           Exit Sub
        '       End If

        '       If Me.TxtModFlg.Text = "N" Then
        '           If CheckBoxAccount(Me.DcboBox.BoundText, val(Me.XPTxtValue(0).Text), Me.XPDtbBill.value) = False Then
        '               Screen.MousePointer = vbDefault
        '               Exit Sub
        '           End If

        '       ElseIf Me.TxtModFlg.Text = "E" Then

        '           If CheckBoxAccount(Me.DcboBox.BoundText, val(Me.XPTxtValue(0).Text), Me.XPDtbBill.value, , , val(Me.XPTxtValue(0).Tag)) = False Then
        '               Screen.MousePointer = vbDefault
        '               Exit Sub
        '           End If
        
        '       End If
        '   End If

        '   If val(Me.XPTxtValue(1).Text) > 0 Then
        '       If ChkInstall.value = vbChecked Then
        '        If val(Me.LblInstallTotal.Caption) = 0 Then
        '            Msg = "ÌÃ» Õ”«» «·√ﬁ”«ÿ ﬁ»· ⁄„·Ì… «·Õ›Ÿ..!!!"
        '            MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
        '            Me.XPTab301.CurrTab = 1
        '            Screen.MousePointer = vbDefault
        '            Exit Sub
        '        End If

        '        If val(Me.LblInstallTotal.Caption) <> val(Me.XPTxtValue(1).Text) Then
        '            Me.XPTxtValue(1).Text = val(Me.LblInstallTotal.Caption)
        '        End If
        '       End If
        '   End If

        'If XPChkPayType(2).value = vbChecked Then
        '    If val(Me.lbl(18).Caption) = 0 Then
        '        Msg = "ÌÃ» ≈œŒ«· «·‘Ìﬂ«  ﬁ»· ⁄„·Ì… «·Õ›Ÿ..!!!"
        '        MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
        '        Me.XPTab301.CurrTab = 1
        '        Screen.MousePointer = vbDefault
        '        Exit Sub
        '    End If
        'End If

        ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        'If DBCboClientName.BoundText = 1 Then
        '    MsgBox "ÌÃ» «Œ Ì«— „Ê—œ √Œ—"
        ' Exit Sub
        'End If

        'Check the Items Grid
    End If
    XPTab301.CurrTab = 0

    If NewGrid.CheckDataEntered = False Then
        Exit Sub
    End If

    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    '    If Me.TxtModFlg.text = "E" Then
    '        If EditTransStatus(Val(Me.XPTxtBillID.text), "E", NewGrid) = False Then
    '            Exit Sub
    '        End If
    '    End If
    '---------------------------------------------------------------
    Cn.Execute "delete DOUBLE_ENTREY_VOUCHERS where Transaction_ID = " & val(Text2.text)

    If NewGrid.Calculate(1, , False, True) = False Then
        Screen.MousePointer = vbDefault
        Exit Sub
    End If

    '-------------------------------
    If Me.XPChkPayType(0).value = vbChecked Then
        DblNotesTotal = val(Me.XPTxtValue(0).text)
    End If

    If Me.XPChkPayType(1).value = vbChecked Then
        DblNotesTotal = DblNotesTotal + val(Me.XPTxtValue(1).text)
    End If

    If Me.XPChkPayType(2).value = vbChecked Then
        DblNotesTotal = DblNotesTotal + val(Me.lbl(18).Caption)
    End If

    'If DblNotesTotal <> Val(LblTotal.Caption) Then
    '    Msg = "≈Ã„«·Ï «·√Ê—«ﬁ «·„«·Ì… €Ì— „ ”«ÊÏ „⁄ ≈Ã„«·Ï «·›« Ê—…...!!!"
    '    MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
    '    Exit Sub
    'End If

    Set RsNotesGeneral = New ADODB.Recordset
                       
    If Notes_coding(val(my_branch), XPDtbBill.value) = "error" Then
        MsgBox " ·« Ì„ﬂ‰ «÷«›… ﬁÌÊœ ·Â–… «·⁄„·Ì… ·«‰ﬂ  ⁄œÌ  «·Õœ «·«ﬁ’Ì ··ﬁÌÊœ ﬂ„« Õœœ  ›Ì  —ﬁÌ„ «·”‰œ«  ": Exit Sub
    Else
                       
        If Notes_coding(val(my_branch), XPDtbBill.value) = "" Then
            MsgBox " ·«Ì„ﬂ‰ «‰‘«¡ «·ﬁÌœ ·Â–« «·„” ‰œ ·«‰ﬂ Õœœ   —ﬁÌ„ ﬁÌÊœ ÌœÊÌ  ": Exit Sub
        Else
            TxtNoteSerial.text = Notes_coding(val(my_branch), XPDtbBill.value)
        End If
    End If

    '                     End If
    my_branch = val(Me.dcBranch.BoundText)
    Dim TxtNoteSerial1str As String
        
    If TxtNoteSerial1.text = "" Then
        TxtNoteSerial1str = Voucher_coding(val(my_branch), XPDtbBill.value, 19, 250, , 28, , val(DCboStoreName.BoundText))
        If TxtNoteSerial1str = "error" Then
            MsgBox " ·« Ì„ﬂ‰ «÷«›… ”‰œ «·«÷«ﬁÂ ÃœÌœ ·«‰ﬂ  ⁄œÌ  «·Õœ «·–Ì ﬁ„  » ÕœÌœ… „‰ «·”‰œ«   ": Exit Sub
        Else
                       
            If TxtNoteSerial1str = "" Then
                MsgBox " ·«»œ „‰ ﬂ «»… —ﬁ„  ”‰œ «·«÷«ﬁÂ  ÌœÊÌ« ﬂ„« Õœœ   ": Exit Sub
            Else
                TxtNoteSerial1.text = TxtNoteSerial1str
            End If
        End If
    End If
    
    '    RsNotesGeneral.Open "[Notes]", Cn, adOpenStatic, adLockOptimistic, adCmdTable
    StrSQL = "SELECT     dbo.Notes.* from dbo.Notes Where (NoteID = -1)"
    RsNotesGeneral.Open StrSQL, Cn, adOpenKeyset, adLockOptimistic, adCmdText
     
    If Me.TxtModFlg.text = "N" Then
        XPTxtBillID.text = CStr(new_id("Transactions", "Transaction_ID", "", True))
        TxtTransSerial.text = CStr(new_id("Transactions", "Transaction_Serial", "", True, "Transaction_Type= 28"))
     
        Me.oldtxtNoteSerial1.text = Trim$(Me.TxtNoteSerial1.text)
    Else
        '   StrSqlDel = "delete From Transaction_Details where Transaction_ID=" & Val(rs("Transaction_ID").value)
        '   Cn.Execute StrSqlDel, , adExecuteNoRecords
        '   StrSqlDel = "delete From Notes where Transaction_ID=" & Val(rs("Transaction_ID").value)
        '   Cn.Execute StrSqlDel, , adExecuteNoRecords
        '
        '   StrSQL = "Delete From DOUBLE_ENTREY_VOUCHERS Where Transaction_ID=" & Val(Me.XPTxtBillID.text)
        '   Cn.Execute StrSQL, , adExecuteNoRecords
        '
        StrSqlDel = "delete From Notes where noteid=" & val(TXTNoteID.text)
        Cn.Execute StrSqlDel, , adExecuteNoRecords
        
        general_noteid = val(TXTNoteID.text)
    End If

    If SystemOptions.TypicalProduction = True Then
        ' do nothing
        TxtNoteSerial.text = ""
    Else
        RsNotesGeneral.AddNew
        RsNotesGeneral("NoteID").value = CStr(new_id("Notes", "NoteID", "", True))
        general_noteid = RsNotesGeneral("NoteID").value
        TXTNoteID.text = general_noteid
        ' RsNotesGeneral("Transaction_ID").value = Val(XPTxtBillID.text)
        RsNotesGeneral("NoteDate").value = XPDtbBill.value
        RsNotesGeneral("NoteType").value = 250 ' «–‰ «÷«›…
        RsNotesGeneral("Note_Value").value = val(LblTotal.Caption)
        RsNotesGeneral("NoteSerial").value = IIf(Trim(Me.TxtNoteSerial.text) = "", Null, Trim(Me.TxtNoteSerial.text))
        ' RsNotesGeneral("NoteSerial1").value = IIf(Trim(Me.TxtNoteSerial1.Text) = "", Null, Trim(Me.TxtNoteSerial1.Text))
        RsNotesGeneral("remark").value = IIf(Trim(Me.TxtNoteSerial1.text) = "", Null, Trim(Me.TxtNoteSerial1.text))
        
        RsNotesGeneral("OldNoteSerial1").value = Trim$(Me.oldtxtNoteSerial1.text) '
        
        RsNotesGeneral("numbering_type").value = sand_numbering_type(0) '”‰œ «·ﬁÌœ
        RsNotesGeneral("numbering_type1").value = sand_numbering_type(19) '  «–‰ «÷«›…
        RsNotesGeneral("sanad_year").value = year(XPDtbBill.value)
        RsNotesGeneral("sanad_month").value = Month(XPDtbBill.value)
        RsNotesGeneral("branch_no").value = val(Me.dcBranch.BoundText)
        
        'RsNotes("note_value_by_characters").value = Trim$(Me.lbl(18).Caption)
        RsNotesGeneral.update
    End If

    '---------Start Saving------------------------------------------------
    Set RSTransDetails = New ADODB.Recordset
    Set RsNotes = New ADODB.Recordset
    '   RSTransDetails.Open "[Transaction_Details]", Cn, adOpenStatic, adLockOptimistic, adCmdTable
    '   RsNotes.Open "[Notes]", Cn, adOpenStatic, adLockOptimistic, adCmdTable
    StrSQL = "SELECT     dbo.Notes.* from dbo.Notes Where (NoteID = -1)"
    RsNotes.Open StrSQL, Cn, adOpenKeyset, adLockOptimistic, adCmdText
 
    StrSQL = "SELECT     dbo.Transaction_Details.* from dbo.Transaction_Details Where (Transaction_ID = -1)"
    RSTransDetails.Open StrSQL, Cn, adOpenKeyset, adLockOptimistic, adCmdText
    
    Screen.MousePointer = vbArrowHourglass
    Cn.BeginTrans
    BeginTrans = True

    If Me.TxtModFlg.text = "N" Then
        XPTxtBillID.text = CStr(new_id("Transactions", "Transaction_ID", "", True))
        rs.AddNew
        rs("Transaction_ID").value = val(XPTxtBillID.text)
    ElseIf Me.TxtModFlg.text = "E" Then

        If rs("Transaction_ID").value <> val(XPTxtBillID.text) Then
            rs.Find "Transaction_ID=" & val(XPTxtBillID.text), , adSearchForward, 1
        End If
         
        ' If Not cmdReSave.Visible Then
        If mSaveDetails Then
            StrSqlDel = "delete From Transaction_Details where Transaction_ID=" & val(rs("Transaction_ID").value)
            Cn.Execute StrSqlDel, , adExecuteNoRecords
        End If
        ' End If
        StrSqlDel = "delete From Notes where Transaction_ID=" & val(rs("Transaction_ID").value)
        Cn.Execute StrSqlDel, , adExecuteNoRecords
        StrSQL = "Delete From DOUBLE_ENTREY_VOUCHERS Where Transaction_ID=" & val(Me.XPTxtBillID.text)
        Cn.Execute StrSQL, , adExecuteNoRecords
    End If

    rs("BranchId").value = IIf(Me.dcBranch.BoundText = "", 0, val(dcBranch.BoundText))
    rs("NoteSerial").value = IIf(Trim(Me.TxtNoteSerial.text) = "", Null, Trim(Me.TxtNoteSerial.text))
    rs("NoteSerial1").value = IIf(Trim(Me.TxtNoteSerial1.text) = "", Null, Trim(Me.TxtNoteSerial1.text))
    rs("OldNoteSerial1").value = Trim$(Me.oldtxtNoteSerial1.text) '
    rs("NoteId").value = val(TXTNoteID.text)
    rs("Transaction_Serial").value = IIf(Trim(Me.TxtTransSerial.text) = "", Null, Trim(Me.TxtTransSerial.text))
    rs("TransactionComment").value = IIf(Trim$(TxtBillComment.text) = "", Null, Trim$(TxtBillComment.text))
    rs("InspectionReport").value = IIf(Trim$(TxtInspectionReport.text) = "", Null, Trim$(TxtInspectionReport.text))
     
    rs("WorkOrderNO").value = IIf(Trim(Me.TXTOrderNO1.text) = "", Null, Trim(Me.TXTOrderNO1.text))
 
    rs("Transaction_Date").value = XPDtbBill.value
    rs("Transaction_Type").value = 28 '1
    rs("UserID").value = user_id
    rs("Shahne").value = val(Txt_EXport.text)
    rs("nots").value = Text1.text

    If XPCboDiscountType.ListIndex = -1 Then
        rs("Trans_DiscountType").value = 0
    Else
        rs("Trans_DiscountType").value = val(XPCboDiscountType.ListIndex)
    End If



    If cmbTypeTrans.ListIndex = -1 Then
        rs("TypeTrans").value = 0
    Else
        rs("TypeTrans").value = val(cmbTypeTrans.ListIndex)
    End If




    If XPCboDiscountType.ListIndex = -1 Or XPCboDiscountType.ListIndex = 0 Then
        rs("Trans_Discount").value = Null
    Else
        rs("Trans_Discount").value = IIf(XPTxtDiscountVal.text = "", Null, (XPTxtDiscountVal.text))
    End If

    If CboPayMentType.ListIndex = -1 Then
        rs("PaymentType").value = 0
    Else
        rs("PaymentType").value = val(CboPayMentType.ListIndex)
    End If

    rs("CusID").value = IIf(DBCboClientName.BoundText = "", Null, (DBCboClientName.BoundText))
    rs("StoreID").value = IIf(DCboStoreName.BoundText = "", Null, (DCboStoreName.BoundText))
    rs("TaxFound").value = IIf(XPChkTAX.value = Checked, True, False)
    rs("TaxValue").value = IIf(XPTxtTaxValue.text = "", Null, val(XPTxtTaxValue.text))

    '÷—»Ì… Œ’„ Ê≈÷«›…
    If ChkTaxAdd.value = vbChecked And val(Me.TxtTaxAddValue.text) > 0 Then
        rs("TaxAddValue").value = val(Me.TxtTaxAddValue.text)
    Else
        rs("TaxAddValue").value = 0
    End If

    '÷—»Ì… œ„€…
    If ChkTaxStamp.value = vbChecked And val(Me.TxtTaxStampValue.text) > 0 Then
        rs("TaxStampValue").value = val(Me.TxtTaxStampValue.text)
    Else
        rs("TaxStampValue").value = 0
    End If

    '÷—»Ì… Œœ„…
    If ChkTaxSerivce.value = vbChecked And val(Me.TxtTaxServiceValue.text) > 0 Then
        rs("TaxServiceValue").value = val(Me.TxtTaxServiceValue.text)
    Else
        rs("TaxServiceValue").value = 0
    End If

    rs.update

    If Me.TxtModFlg.text = "E" Then
        StrSqlDel = "delete From Notes where Transaction_ID=" & val(rs("Transaction_ID").value)
        Cn.Execute StrSqlDel, , adExecuteNoRecords
        If cmdReSave.Visible Then GoTo Line15
        
        StrSqlDel = "delete From Transaction_Details where Transaction_ID=" & val(rs("Transaction_ID").value)
        Cn.Execute StrSqlDel, , adExecuteNoRecords
       
    End If
    If mSaveDetails Then
        For RowNum = 1 To FG.rows - 1

            'Check Repeat Serial
            If FG.TextMatrix(RowNum, FG.ColIndex("Serial")) <> "" Then
                StrSQL = "select * From Transaction_Details where ItemSerial='" & FG.TextMatrix(RowNum, FG.ColIndex("Serial")) & "'"
                StrSQL = StrSQL + " and Transaction_ID =" & XPTxtBillID.text
                Set RsTemp = New ADODB.Recordset
                RsTemp.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

                If Not (RsTemp.EOF Or RsTemp.BOF) Then
                    Msg = "«·”Ì—Ì«· «·Œ«’ »«·’‰›" & CHR(13)
                    Msg = Msg + FG.cell(flexcpTextDisplay, RowNum, FG.ColIndex("name")) & CHR(13)
                    Msg = Msg + " „ √œŒ«·Â ·ﬁÿ⁄… √Œ—Ï ›Ì Â–Â «·›« Ê—…"
                    MsgBox Msg, vbOKOnly + vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
                    RsTemp.Close
                    XPTab301.CurrTab = 0
                    FG.Row = RowNum
                    FG.Col = FG.ColIndex("name")
                    FG.ShowCell RowNum, FG.ColIndex("name")
                    FG.SetFocus
                    Screen.MousePointer = vbDefault
                    BeginTrans = False
                    Cn.RollbackTrans
                    Exit Sub
                End If

                RsTemp.Close
            End If

            If FG.TextMatrix(RowNum, FG.ColIndex("Code")) <> "" Then
                RSTransDetails.AddNew
                RSTransDetails("FoxyNo").value = IIf((FG.TextMatrix(RowNum, FG.ColIndex("FoxyNo")) = ""), Null, FG.TextMatrix(RowNum, FG.ColIndex("FoxyNo")))
        
                RSTransDetails("BranchId").value = IIf(Me.dcBranch.BoundText = "", 0, val(dcBranch.BoundText))
                RSTransDetails("Transaction_ID").value = val(XPTxtBillID.text)
                RSTransDetails("Item_ID").value = IIf((FG.TextMatrix(RowNum, FG.ColIndex("Code")) = ""), Null, val(FG.TextMatrix(RowNum, FG.ColIndex("Code"))))
                RSTransDetails("Quantity").value = IIf((FG.TextMatrix(RowNum, FG.ColIndex("Count")) = ""), Null, val(FG.TextMatrix(RowNum, FG.ColIndex("Count"))))

                '            RSTransDetails("ItemName").Value = IIf((FG.TextMatrix(RowNum, FG.ColIndex("Name")) _
                '            = ""), "", Val(FG.TextMatrix(RowNum, FG.ColIndex("Name"))))
                If Not FG.TextMatrix(RowNum, FG.ColIndex("Name")) = "" Then
                    StrSQL = "select * From TblItems where ItemID=" & FG.TextMatrix(RowNum, FG.ColIndex("Name"))
                    RsTemp.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

                    If Not (RsTemp.EOF Or RsTemp.BOF) Then
                        If RsTemp("HaveSerial").value = True Then
                            RSTransDetails("ItemSerial").value = IIf((FG.TextMatrix(RowNum, FG.ColIndex("Serial")) = ""), Null, (FG.TextMatrix(RowNum, FG.ColIndex("Serial"))))
                        End If
                    End If

                    RsTemp.Close
                End If

                RSTransDetails("ItemCase").value = IIf((FG.TextMatrix(RowNum, FG.ColIndex("ItemCase")) = ""), Null, (FG.TextMatrix(RowNum, FG.ColIndex("ItemCase"))))
                RSTransDetails("showPrice").value = IIf((FG.TextMatrix(RowNum, FG.ColIndex("Price")) = ""), Null, val(FG.TextMatrix(RowNum, FG.ColIndex("Price"))))
                RSTransDetails("ItemDiscountType").value = IIf((FG.TextMatrix(RowNum, FG.ColIndex("DiscountType")) = ""), Null, (FG.TextMatrix(RowNum, FG.ColIndex("DiscountType"))))
                RSTransDetails("ItemDiscount").value = IIf((FG.TextMatrix(RowNum, FG.ColIndex("DiscountVal")) = ""), Null, val(FG.TextMatrix(RowNum, FG.ColIndex("DiscountVal"))))
                RSTransDetails("guaranteeTime").value = IIf((FG.TextMatrix(RowNum, FG.ColIndex("guaranteeTime")) = ""), Null, val(FG.TextMatrix(RowNum, FG.ColIndex("guaranteeTime"))))
        
                RSTransDetails("NProductionOrderNO").value = IIf((FG.TextMatrix(RowNum, FG.ColIndex("NProductionOrderNO")) = ""), Null, (FG.TextMatrix(RowNum, FG.ColIndex("NProductionOrderNO"))))
        
                RSTransDetails("Remarks").value = IIf((FG.TextMatrix(RowNum, FG.ColIndex("Remarks")) = ""), Null, Trim$(FG.TextMatrix(RowNum, FG.ColIndex("Remarks"))))
         
                '.TextMatrix(LngRow, .ColIndex("ColorID")) = 1
                '.TextMatrix(LngRow, .ColIndex("ItemSize")) = 0
        
                RSTransDetails("ColorID").value = IIf((FG.TextMatrix(RowNum, FG.ColIndex("ColorID")) = ""), 1, val(FG.TextMatrix(RowNum, FG.ColIndex("ColorID"))))
                RSTransDetails("ItemSize").value = IIf((FG.TextMatrix(RowNum, FG.ColIndex("ItemSize")) = ""), 1, Trim$(FG.TextMatrix(RowNum, FG.ColIndex("ItemSize"))))
                RSTransDetails("ClassId").value = IIf((FG.TextMatrix(RowNum, FG.ColIndex("ClassId")) = ""), 1, val(FG.TextMatrix(RowNum, FG.ColIndex("ClassId"))))
           
                RSTransDetails("UnitID").value = IIf(FG.cell(flexcpData, RowNum, FG.ColIndex("UnitID")) = "", Null, (FG.cell(flexcpData, RowNum, FG.ColIndex("UnitID"))))
                RSTransDetails("ShowQty").value = IIf((FG.TextMatrix(RowNum, FG.ColIndex("Count")) = ""), Null, val(FG.TextMatrix(RowNum, FG.ColIndex("Count"))))
             
                Dim RsUnitData   As ADODB.Recordset
                Dim LngCurItemID As Long
                Dim LngUnitID    As Long
                Dim DblQty       As Double
        
                LngCurItemID = val(FG.TextMatrix(RowNum, FG.ColIndex("Code")))
                LngUnitID = val(FG.cell(flexcpData, RowNum, FG.ColIndex("UnitID")))
                DblQty = val(FG.TextMatrix(RowNum, FG.ColIndex("Count")))

                StrSQL = "Select * From TblItemsUnits Where ItemID=" & LngCurItemID
                StrSQL = StrSQL + " AND UnitID=" & LngUnitID
                Set RsUnitData = New ADODB.Recordset
                RsUnitData.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

                If Not (rs.BOF Or rs.EOF) Then
                    RSTransDetails("QtyBySmalltUnit").value = RsUnitData("UnitFactor").value
                    RSTransDetails("Quantity").value = RSTransDetails("QtyBySmalltUnit").value * RSTransDetails("showqty").value
            
                    RSTransDetails("DamageshowQty").value = IIf((FG.TextMatrix(RowNum, FG.ColIndex("DamageshowQty")) = ""), Null, val(FG.TextMatrix(RowNum, FG.ColIndex("DamageshowQty"))))
                    RSTransDetails("DamageQty").value = RSTransDetails("QtyBySmalltUnit").value * RSTransDetails("DamageshowQty").value
            
                End If

                '     RSTransDetails("price").value = Round(FG.TextMatrix(RowNum, FG.ColIndex("Price")) / RSTransDetails("Quantity").value, 2)
                RSTransDetails("Price").value = val(IIf((FG.TextMatrix(RowNum, FG.ColIndex("Price")) = ""), Null, val(FG.TextMatrix(RowNum, FG.ColIndex("Price"))))) / RSTransDetails("QtyBySmalltUnit").value
        
                '        Dim RsPrice As New ADODB.Recordset
                '        Set RsPrice = New ADODB.Recordset
                '
                '        RsPrice.Open "select UnitPurPrice from TblItemsUnits where ItemID=" & FG.TextMatrix(RowNum, FG.ColIndex("Code")) & " and UnitID=" & FG.Cell(flexcpData, RowNum, FG.ColIndex("UnitID")), Cn, adOpenStatic, adLockOptimistic, adCmdTable
            
                ' RSTransDetails("price").Value = Round(FG.TextMatrix(RowNum, FG.ColIndex("Valu")) / RSTransDetails("Quantity").Value, 2)
                If val(Txt_EXport.text) > 0 Then
                    Dim mm    As String
                    Dim Myprc As String
                    mm = MsgBox(" Â· Â‰«ﬂ „’«—Ì› √Œ—Ï ⁄·Ï Â–« «·«–‰ ... «–«  „  Õ„Ì· Â–Â «·„’—Ê›«  ›·« ÌÕﬁ ·ﬂ «· ⁄œÌ·", vbYesNo)

                    If mm = vbYes Then

                        '   Â·  „  ÕÊÌ· «·«–‰ «·Ï ›« Ê—…
                        If Text1.text <> "" Then
                            
                            RSTransDetails("ToTAlELSHahn") = (((RSTransDetails("showPrice") * RSTransDetails("ShowQty")) / val(LblTotal.Caption)) * val(Txt_EXport.text)) / RSTransDetails("ShowQty")
                      
                            Myprc = RSTransDetails("showprice").value / RSTransDetails("QtyBySmalltUnit").value
                         
                            Myprc = (RSTransDetails("ToTAlELSHahn").value / RSTransDetails("QtyBySmalltUnit").value) + Myprc
                            RSTransDetails("Price").value = Myprc
                               
                            Mytot = RSTransDetails("showprice").value + RSTransDetails("ToTAlELSHahn")
                            RSTransDetails("showprice").value = Mytot
                        Else
                            MsgBox "ÌÃ»  ÕÊÌ· «·«–‰ «·Ï ›« Ê—… ﬁ»·  Õ„Ì·Â »«· ﬂ«·Ì› «·«Œ—Ï"
                        End If
                    End If

                    ' Round(FG.TextMatrix(RowNum, FG.ColIndex("Valu")) / RSTransDetails("Quantity").Value, 2)
                End If

                RSTransDetails("order_no").value = IIf(FG.TextMatrix(RowNum, FG.ColIndex("order_no")) = "", Null, FG.TextMatrix(RowNum, FG.ColIndex("order_no")))
        
                RSTransDetails("ColorID").value = IIf((FG.TextMatrix(RowNum, FG.ColIndex("ColorID")) = ""), 1, val(FG.TextMatrix(RowNum, FG.ColIndex("ColorID"))))
            
                RSTransDetails("ItemSize").value = IIf((FG.TextMatrix(RowNum, FG.ColIndex("ItemSize")) = ""), 1, val(FG.TextMatrix(RowNum, FG.ColIndex("ItemSize"))))
            
                RSTransDetails("ClassId").value = IIf((FG.TextMatrix(RowNum, FG.ColIndex("ClassId")) = ""), 1, val(FG.TextMatrix(RowNum, FG.ColIndex("ClassId"))))
            
            
                With FG
                i = RowNum
                RSTransDetails("Remarks").value = IIf((.TextMatrix(i, .ColIndex("Remarks"))) = "", Null, (.TextMatrix(i, .ColIndex("Remarks"))))
                RSTransDetails("LineId").value = i
                RSTransDetails("L").value = IIf((FG.TextMatrix(i, FG.ColIndex("L")) = ""), Null, (FG.TextMatrix(i, FG.ColIndex("L"))))
                RSTransDetails("W").value = IIf((FG.TextMatrix(i, FG.ColIndex("W")) = ""), Null, (FG.TextMatrix(i, FG.ColIndex("W"))))
                RSTransDetails("H1").value = IIf((FG.TextMatrix(i, FG.ColIndex("H1")) = ""), Null, (FG.TextMatrix(i, FG.ColIndex("H1"))))
                RSTransDetails("H2").value = IIf((FG.TextMatrix(i, FG.ColIndex("H2")) = ""), Null, (FG.TextMatrix(i, FG.ColIndex("H2"))))
                RSTransDetails("NoCount").value = IIf((FG.TextMatrix(i, FG.ColIndex("NoCount")) = ""), Null, val(FG.TextMatrix(i, FG.ColIndex("NoCount"))))
                RSTransDetails("Width").value = IIf((FG.TextMatrix(i, FG.ColIndex("Width")) = ""), Null, val(FG.TextMatrix(i, FG.ColIndex("Width"))))
                RSTransDetails("length").value = IIf((FG.TextMatrix(i, FG.ColIndex("length")) = ""), Null, val(FG.TextMatrix(i, FG.ColIndex("length"))))
                RSTransDetails("Height").value = IIf((FG.TextMatrix(i, FG.ColIndex("Height")) = ""), Null, val(FG.TextMatrix(i, FG.ColIndex("Height"))))
                RSTransDetails("Area").value = IIf((FG.TextMatrix(i, FG.ColIndex("Area")) = ""), Null, val(FG.TextMatrix(i, FG.ColIndex("Area"))))
                RSTransDetails("OldID").value = val(FG.TextMatrix(i, FG.ColIndex("OldID")))
                RSTransDetails("ColorID").value = IIf((FG.TextMatrix(i, FG.ColIndex("ColorID")) = ""), 1, val(FG.TextMatrix(i, FG.ColIndex("ColorID"))))
                RSTransDetails("ItemSize").value = IIf((FG.TextMatrix(i, FG.ColIndex("ItemSize")) = ""), 1, Trim$(FG.TextMatrix(i, FG.ColIndex("ItemSize"))))
                RSTransDetails("ClassId").value = IIf((FG.TextMatrix(i, FG.ColIndex("ClassId")) = ""), 1, val(FG.TextMatrix(i, FG.ColIndex("ClassId"))))
                End With

                RSTransDetails("BranchId").value = IIf(Me.dcBranch.BoundText = "", 0, val(dcBranch.BoundText))
                RSTransDetails("OrderArrivalDate").value = IIf(Not IsDate(FG.TextMatrix(RowNum, FG.ColIndex("OrderArrivalDate"))), Null, FG.TextMatrix(RowNum, FG.ColIndex("OrderArrivalDate")))
                RSTransDetails("order_no").value = IIf(FG.TextMatrix(RowNum, FG.ColIndex("order_no")) = "", Null, FG.TextMatrix(RowNum, FG.ColIndex("order_no")))
                RSTransDetails("FoxyNo").value = IIf((FG.TextMatrix(RowNum, FG.ColIndex("FoxyNo")) = ""), Null, FG.TextMatrix(RowNum, FG.ColIndex("FoxyNo")))
             
                Dim OldQty  As Double
                Dim OldCost As Double
                Dim NewQty  As Double
                Dim NewCost As Double
               
                getItemCostData XPDtbBill.value, RSTransDetails("Item_ID").value, val(DCboStoreName.BoundText), val(Me.XPTxtBillID.text), OldQty, OldCost, NewQty, NewCost, , LngUnitID
                RSTransDetails("OldQty").value = NewQty
                RSTransDetails("OldCost").value = NewCost
       
                RSTransDetails("NewQty").value = RSTransDetails("Quantity").value + RSTransDetails("OldQty").value
                If (RSTransDetails("Quantity").value + RSTransDetails("OldQty").value) <> 0 Then
                    RSTransDetails("NewCost").value = ((RSTransDetails("OldQty").value * RSTransDetails("OldCost").value) + (RSTransDetails("Quantity").value * RSTransDetails("Price").value)) / (RSTransDetails("Quantity").value + RSTransDetails("OldQty").value)
                Else
                    RSTransDetails("NewCost").value = 0
                End If
            
                RSTransDetails.update
            End If

        Next RowNum
    End If
Line15:
    '------------------------------------------------------------------------------
    '„‰ Â‰« «·ﬂÊœ  Ê«ﬁ›
    '------------------------------------------------------------------------------
    'If Me.XPChkPayType(0).Value = Checked Then
    '    RsNotes.AddNew
    '    RsNotes("NoteID").Value = CStr(new_id("Notes", "NoteID", "", True))
    '    Note_ID = RsNotes("NoteID").Value
    '    If Me.TxtModFlg.text = "N" Then
    '        RsNotes("NoteSerial").Value = CStr(new_id("Notes", "NoteSerial", "", True))
    '        XPTxtSerial(0).text = RsNotes("NoteSerial").Value
    '    ElseIf Trim(XPTxtSerial(0).text) <> "" Then
    '        RsNotes("NoteSerial").Value = Trim(XPTxtSerial(0).text)
    '    Else
    '        RsNotes("NoteSerial").Value = CStr(new_id("Notes", "NoteSerial", "", True))
    '        XPTxtSerial(0).text = RsNotes("NoteSerial").Value
    '    End If
    '    RsNotes("Transaction_ID").Value = Val(XPTxtBillID.text)
    '    RsNotes("NoteDate").Value = XPDtbBill.Value
    '    RsNotes("NoteType").Value = 0
    '    RsNotes("Note_Value").Value = _
    '    IIf(XPTxtValue(0).text = "", Null, Val(XPTxtValue(0).text))
    '    RsNotes("Member_ID").Value = _
    '    IIf(DBCboClientName.BoundText = "", Null, Val(DBCboClientName.BoundText))
    '    RsNotes("BankID").Value = Null
    '    RsNotes("BoxID").Value = IIf(DcboBox.BoundText = "", Null, Val(DcboBox.BoundText))
    '    RsNotes("CusID").Value = Null
    '    RsNotes.update
    '    Me.XPTxtValue(0).Tag = RsNotes("NoteID").Value
    '    '--------------------------------------------------------------------------
    'End If
    'If Me.XPChkPayType(1).Value = Checked Then
    '    RsNotes.AddNew
    '    RsNotes("NoteID").value = CStr(new_id("Notes", "NoteID", "", True))
    '    note_id = RsNotes("NoteID").value
    '    RsNotes("NoteDate").value = XPDtbBill.value
    '    If Me.TxtModFlg.text = "N" Then
    '        RsNotes("NoteSerial").value = CStr(new_id("Notes", "NoteSerial", "", True))
    '        XPTxtSerial(1).text = RsNotes("NoteSerial").value
    '    ElseIf Trim(XPTxtSerial(1).text) <> "" Then
    '        RsNotes("NoteSerial").value = Trim(XPTxtSerial(1).text)
    '    Else
    '        RsNotes("NoteSerial").value = CStr(new_id("Notes", "NoteSerial", "", True))
    '         XPTxtSerial(1).text = RsNotes("NoteSerial").value
    '    End If
    '    RsNotes("branch_no").value = Val(Me.DcBranch.BoundText)
    '    RsNotes("Transaction_ID").value = Val(XPTxtBillID.text)
    '    RsNotes("NoteType").value = 1
    '    RsNotes("Note_Value").value = Val(LblTotalAll.Caption)
    '
    '    RsNotes("Member_ID").value = _
    '    IIf(DBCboClientName.BoundText = "", Null, Val(DBCboClientName.BoundText))
    '    RsNotes("BankID").value = Null
    '    RsNotes("CusID").value = Null
    '    RsNotes("DueDate").value = DtpDelayDate.value
    '    RsNotes.update
    '    Me.XPTxtValue(1).Tag = RsNotes("NoteID").value
    'End If
    'If Me.XPChkPayType(2).Value = Checked Then
    '    With Me.FgCheques
    '        For I = .FixedRows To .Rows - 1
    '            RsNotes.AddNew
    '                RsNotes("NoteID").Value = CStr(new_id("Notes", "NoteID", "", True))
    '                Note_ID = RsNotes("NoteID").Value
    '                RsNotes("NoteDate").Value = XPDtbBill.Value
    '                RsNotes("Transaction_ID").Value = Val(XPTxtBillID.text)
    '                RsNotes("NoteType").Value = 13
    '                RsNotes("Note_Value").Value = Val(.TextMatrix(I, .ColIndex("CheckValue")))
    '                RsNotes("BankID").Value = Val(.TextMatrix(I, .ColIndex("BankID")))
    '                RsNotes("ChqueNum").Value = Trim$(.TextMatrix(I, .ColIndex("CheckNumber")))
    '                RsNotes("DueDate").Value = CDate(Trim$(.TextMatrix(I, .ColIndex("DueDate"))))
    '                RsNotes("Member_ID").Value = Val(Me.DBCboClientName.BoundText)
    '                RsNotes("CUSID").Value = Val(Me.DBCboClientName.BoundText)
    '            RsNotes.update
    '            '--------------------------------------------------------------------------
    '        Next I
    '    End With
    'End If
    'Õ›Ÿ «·√›”«ÿ
    'If Me.XPChkPayType(1).Value = Checked Then
    '    If ChkInstall.Value = vbChecked Then
    '        'Save installment Data
    '        Set RsTemp = New ADODB.Recordset
    '        RsTemp.Open "InstallMent", Cn, adOpenStatic, adLockOptimistic, adCmdTable
    '        RsTemp.AddNew
    '            RsTemp("PartID").Value = CStr(new_id("InstallMent", "PartID", "", True))
    '            RsTemp("NoteID").Value = Note_ID
    '            RsTemp("BasicAmmount").Value = IIf(XPTxtValue(1).text = "", 0, Val(XPTxtValue(1).text))
    '            RsTemp("InterestType").Value = Val(Me.LblPrecenType.Tag)
    '            RsTemp("InterestVal").Value = Val(LblPrecenValue.Caption)
    '            RsTemp("Total").Value = Val(LblInstallTotal.Caption)
    '            RsTemp("InstallCount").Value = Val(LblInstallCount.Caption)
    '            RsTemp("FirstInstallDate").Value = CDate(Me.LblFirstInstallDate.Caption)
    '            If Val(LblInstallmentType.Tag) = 0 Then
    '                RsTemp("InstallmentType").Value = 0
    '            ElseIf Val(LblInstallmentType.Tag) = 1 Then
    '                RsTemp("InstallmentType").Value = 1
    '            ElseIf Val(LblInstallmentType.Tag) = 2 Then
    '                RsTemp("InstallmentType").Value = 2
    '            End If
    '            RsTemp("InstallSeprator").Value = Val(Me.LblInstallSeprator.Caption)
    '            RsTemp("StartValue").Value = IIf(Val(Me.LblStartValue.Caption) = 0, Null, Val(Me.LblStartValue.Caption))
    '            RsTemp("CustID").Value = IIf(DBCboClientName.BoundText = "", Null, Val(DBCboClientName.BoundText))
    '            RsTemp("Type").Value = 1
    '        RsTemp.update
    '        'save installment Details
    '        Set RsDetalis = New ADODB.Recordset
    '        RsDetalis.Open "InstallMentDetails", Cn, adOpenStatic, adLockOptimistic, adCmdTable
    '        With Me.FgInstallments
    '            For RowNum = 1 To .Rows - 1
    '                RsDetalis.AddNew
    '                    RsDetalis("QestID").Value = CStr(new_id("InstallMentDetails", "QestID", "", True))
    '                    RsDetalis("PartID").Value = RsTemp("PartID").Value
    '                    RsDetalis("QeqtNum").Value = IIf(.TextMatrix(RowNum, .ColIndex("Serial")) = "", "", .TextMatrix(RowNum, .ColIndex("Serial")))
    '                    RsDetalis("Value").Value = IIf(.TextMatrix(RowNum, .ColIndex("Value")) = "", "", Val(.TextMatrix(RowNum, .ColIndex("Value"))))
    '                    RsDetalis("DueDate").Value = IIf(.TextMatrix(RowNum, .ColIndex("Due_Date")) = "", "", .TextMatrix(RowNum, .ColIndex("Due_Date")))
    '                    RsDetalis("Receipt").Value = False
    '                RsDetalis.update
    '            Next RowNum
    '        End With
    '    End If
    'End If
    Dim LngDevID             As Long
    Dim LngDevNO             As Integer
    Dim StrTempAccountCode   As String
    Dim StrTempDes           As String
    Dim SngTemp              As Variant
    Dim Account_Code_dynamic As String

    If SystemOptions.TypicalProduction = True Then
        GoTo tp
    End If

Dim s As String

s = "Select IsNotCreateEntry from TblStore where isnull(IsNotCreateEntry,0) = 1 and StoreId = " & val(DCboStoreName.BoundText)

Dim rsDummy As New ADODB.Recordset
rsDummy.Open s, Cn, adOpenKeyset, adLockReadOnly

If Not rsDummy.EOF Then
        GoTo ⁄ —Ì”
End If


rsDummy.Close


    LngDevID = new_id("DOUBLE_ENTREY_VOUCHERS", "Double_Entry_Vouchers_ID", "")
    '«·ÿ—› «·„œÌ‰
    SngTemp = NewGrid.GetItemsTotal(ItemsGoodType)

    If SngTemp > 0 Then
        If detect_inventory_work_type = 1 Then

            Account_Code_dynamic = get_account_code_branch(0, my_branch)
        
            If Account_Code_dynamic = "NO branch" Then
                MsgBox "·„ Ì „ «‰‘«¡ «·›—⁄", vbCritical
                GoTo ErrTrap
            Else

                If Account_Code_dynamic = "NO account" Then
                    MsgBox "·„ Ì „  ÕœÌœ Õ”«»  «·„Œ“Ê‰ ›Ì «·›—⁄ ·Â–… «·⁄„·Ì…", vbCritical
                    GoTo ErrTrap
         
                End If
            End If

            StrTempAccountCode = Account_Code_dynamic '„Œ“Ê‰ «·»÷«⁄…
            ' StrTempAccountCode = "a1a2a5" '„Œ“Ê‰ «·»÷«⁄…
            StrTempDes = "  ”‰œ «” ·«„ —ﬁ„ " & Me.TxtNoteSerial1.text
            LngDevNO = LngDevNO + 1

            If ModAccounts.AddNewDev(LngDevID, LngDevNO, StrTempAccountCode, SngTemp, 0, StrTempDes, general_noteid, , , , Me.XPDtbBill.value, Me.DCboUserName.BoundText, val(Me.XPTxtBillID.text), , , , , , , , , , , , , , , , , val(Me.dcBranch.BoundText)) = False Then
                GoTo ErrTrap
            End If

        ElseIf detect_inventory_work_type = 2 Then
            '«·„Œ“Ê‰ «·”·⁄Ì ⁄·Ï „” ÊÏ «·„Œ“‰
    
            Account_Code_dynamic = get_store_Account(DCboStoreName.BoundText, "Account_Code")

            If Account_Code_dynamic = "" Then
                MsgBox "·„ Ì „  ÕœÌœ Õ”«»  ··„Œ“Ê‰ «·”·⁄Ì ·Â–« «·„Œ“‰ ›Ì Â–« «·›—⁄    ", vbCritical
                GoTo ErrTrap
            End If
    
            StrTempAccountCode = Account_Code_dynamic  '„Õ“Ê‰ «·”·⁄Ì ··„Œ“‰
            ' StrTempAccountCode = "a1a2a5" '„Õ“Ê‰ «·»÷«⁄…
            StrTempDes = "  ”‰œ «” ·«„ —ﬁ„ " & Me.TxtNoteSerial1.text
            LngDevNO = LngDevNO + 1

            If ModAccounts.AddNewDev(LngDevID, LngDevNO, StrTempAccountCode, SngTemp, 0, StrTempDes, general_noteid, , , , Me.XPDtbBill.value, Me.DCboUserName.BoundText, val(Me.XPTxtBillID.text), , , , , , , , , , , , , , , , , val(Me.dcBranch.BoundText)) = False Then
                GoTo ErrTrap
            End If

        ElseIf detect_inventory_work_type = 3 Then
            Dim groupAccount As String
             
            Dim line_value   As Single

            With FG

                For i = 1 To FG.rows - 1

                    If FG.TextMatrix(i, FG.ColIndex("Code")) <> "" Then
    
                        ' groupAccount = get_item_group_account(FG.TextMatrix(i, FG.ColIndex("Code")), DCboStoreName.BoundText, 2)
                        groupAccount = get_item_group_account_inventory(FG.TextMatrix(i, FG.ColIndex("Code")), DCboStoreName.BoundText, 0)

                        If groupAccount = "Error" Then
                            If SystemOptions.UserInterface = ArabicInterface Then
                                MsgBox "«·’‰› ›Ì «·”ÿ— —ﬁ„" & i & " €Ì— „Õœœ —ﬁ„ Õ”«»  «·„Œ“Ê‰ «·”⁄·⁄Ì ··„Œ“‰ «·„Õœœ   ·„Ã„Ê⁄ …"
                            Else
                                MsgBox "Item in line no " & i & "Group Name Account Not Defined"
                            End If

                            GoTo ErrTrap
                        End If

                        line_value = FG.TextMatrix(i, FG.ColIndex("Price")) * FG.TextMatrix(i, FG.ColIndex("Count"))
    
                        StrTempDes = " ”‰œ «” ·«„   —ﬁ„ " & Me.TxtNoteSerial1.text
                        LngDevNO = LngDevNO + 1

                        If ModAccounts.AddNewDev(LngDevID, LngDevNO, groupAccount, line_value, 0, StrTempDes, general_noteid, , , , Me.XPDtbBill.value, Me.DCboUserName.BoundText, val(Me.XPTxtBillID.text), , , , , , , , , , , , , , , , , val(Me.dcBranch.BoundText)) = False Then
                            GoTo ErrTrap
                        End If
    
                    End If

                Next i

            End With

        End If

        '«·ÿ—› «·œ«∆‰
        SngTemp = NewGrid.GetItemsTotal(ItemsGoodType)

        If SngTemp > 0 Then
            If detect_inventory_work_type = 1 Or detect_inventory_work_type = 2 Then

                Account_Code_dynamic = get_account_code_branch(37, my_branch)
        
                If Account_Code_dynamic = "NO branch" Then
                    MsgBox "·„ Ì „ «‰‘«¡ «·›—⁄", vbCritical
                    GoTo ErrTrap
                Else

                    If Account_Code_dynamic = "NO account" Then
                        MsgBox "·„ Ì „  ÕœÌœ Õ”«»  ﬂ«·Ì› «·«‰ «Ã „Ê«œ ›Ì «·›—⁄ ·Â–… «·⁄„·Ì…", vbCritical
                        GoTo ErrTrap
         
                    End If
                End If

                StrTempAccountCode = Account_Code_dynamic '«·„‘ —Ì« 
            
                StrTempDes = "«–‰ «÷«›…" & Me.TxtTransSerial.text
                LngDevNO = LngDevNO + 1

                If ModAccounts.AddNewDev(LngDevID, LngDevNO, StrTempAccountCode, SngTemp, 1, StrTempDes, general_noteid, , , , Me.XPDtbBill.value, Me.DCboUserName.BoundText, val(Me.XPTxtBillID.text), , , , , , , , , , , , , , , , , val(Me.dcBranch.BoundText)) = False Then
                    GoTo ErrTrap
                End If
         
            ElseIf detect_inventory_work_type = 3 Then

                With FG

                    For i = 1 To FG.rows - 1

                        If FG.TextMatrix(i, FG.ColIndex("Code")) <> "" Then
    
                            groupAccount = get_item_group_account_in_branch(FG.TextMatrix(i, FG.ColIndex("Code")), val(my_branch), 4)

                            '  groupAccount = get_item_group_account_inventory(FG.TextMatrix(I, FG.ColIndex("Code")), DCboStoreName.BoundText, 4)
                            If groupAccount = "Error" Then
                                If SystemOptions.UserInterface = ArabicInterface Then
                                    MsgBox "«·’‰› ›Ì «·”ÿ— —ﬁ„" & i & " €Ì— „Õœœ —ﬁ„ Õ”«»   «·„‘ —Ì«    ··„Œ“‰ «·„Õœœ   ·„Ã„Ê⁄ …"
                                Else
                                    MsgBox "Item in line no " & i & "Group Name Account Not Defined"
                                End If

                                GoTo ErrTrap
                            End If

                            line_value = FG.TextMatrix(i, FG.ColIndex("Price")) * FG.TextMatrix(i, FG.ColIndex("Count"))
    
                            StrTempDes = "«–‰ «÷«›…" & Me.TxtTransSerial.text
                            LngDevNO = LngDevNO + 1

                            If ModAccounts.AddNewDev(LngDevID, LngDevNO, groupAccount, line_value, 1, StrTempDes, general_noteid, , , , Me.XPDtbBill.value, Me.DCboUserName.BoundText, val(Me.XPTxtBillID.text), , , , , , , , , , , , , , , , , val(Me.dcBranch.BoundText)) = False Then
                                GoTo ErrTrap
                            End If
    
                        End If

                    Next i

                End With

            End If
        End If

        '
        '        Account_Code_dynamic = get_store_Account(DCboStoreName.BoundText, "Account_Code")
        '        If Account_Code_dynamic = "" Then
        '         MsgBox "·„ Ì „  ÕœÌœ Õ”«»  ··„Œ“Ê‰ «·”·⁄Ì ·Â–« «·„Œ“‰ ›Ì Â–« «·›—⁄    ", vbCritical
        '        GoTo ErrTrap
        '        End If
        '
        '    StrTempAccountCode = Account_Code_dynamic '„Œ“Ê‰ «·»÷«⁄…0
        '  '  StrTempAccountCode = "a1a2a5" '„Œ“Ê‰ «·»÷«⁄…
        '    StrTempDes = "«–‰ «÷«›… —ﬁ„ " & Me.TxtTransSerial.text
        '    LngDevNO = LngDevNO + 1
        '    If ModAccounts.AddNewDev(LngDevID, LngDevNO, StrTempAccountCode, SngTemp, _
        '        0, StrTempDes, , , , , Me.XPDtbBill.value, Me.DCboUserName.BoundText, Val(Me.XPTxtBillID.text)) = False Then
        '        GoTo ErrTrap
        '    End If
    End If

tp:

    If XPChkTAX.value = vbChecked Then
        '  StrTempAccountCode = "a1a3a5" '÷—»Ì… „»Ì⁄«  „œÌ‰…
        '  SngTemp = Val(Me.lbl(25).Caption)
        '  StrTempDes = "«–‰ «÷«›…  —ﬁ„ " & Me.TxtTransSerial.text
        '  LngDevNO = LngDevNO + 1
        '  If ModAccounts.AddNewDev(LngDevID, LngDevNO, StrTempAccountCode, SngTemp, _
        '     0, StrTempDes, , , , , Me.XPDtbBill.value, Me.DCboUserName.BoundText, Val(Me.XPTxtBillID.text)) = False Then
        '      GoTo ErrTrap
        '  End If
    End If

    If Me.ChkTaxAdd.value = vbChecked Then
        '  StrTempAccountCode = "a2a5a4" '÷—»Ì… √—»«Õ  Ã«—Ì… (Œ’„ Ê≈÷«›…
        '  StrTempDes = "«–‰ «÷«›… —ﬁ„ " & Me.TxtTransSerial.text
        '  SngTemp = Val(Me.lbl(32).Caption)
        '  LngDevNO = LngDevNO + 1
        '  If ModAccounts.AddNewDev(LngDevID, LngDevNO, StrTempAccountCode, SngTemp, _
        '      0, StrTempDes, , , , , Me.XPDtbBill.value, Me.DCboUserName.BoundText, Val(Me.XPTxtBillID.text)) = False Then
        '      GoTo ErrTrap
        '  End If
    End If

    '«·œ«∆‰
    'If CboPaymentType.ListIndex = 0 Then  'Me.XPChkPayType(0).Value = vbChecked Then
    '    '«·Œ“Ì‰…
    '    StrTempAccountCode = GetMyAccountCode("TblBoxesData", "BoxID", Val(Me.DcboBox.BoundText))
    '    StrTempDes = "«–‰ «÷«›… —ﬁ„ " & Me.TxtTransSerial.text
    '
    '    SngTemp = DisplayCurrency(Val(Me.XPTxtValue(0).text))
    '    LngDevNO = LngDevNO + 1
    '    If ModAccounts.AddNewDev(LngDevID, LngDevNO, StrTempAccountCode, SngTemp, _
    '        1, StrTempDes, , , , , Me.XPDtbBill.value, Me.DCboUserName.BoundText, Val(Me.XPTxtBillID.text)) = False Then
    '        GoTo ErrTrap
    '    End If
    'End If
    'If CboPaymentType.ListIndex = 1 Then 'Me.XPChkPayType(1).Value = vbChecked Then
    '    '«·√Ã·
    '    StrTempAccountCode = GetMyAccountCode("TblCustemers", "CusID", Val(Me.DBCboClientName.BoundText))
    '    StrTempDes = "«–‰ «÷«›… —ﬁ„ " & Me.TxtTransSerial.text
    '    LngDevNO = LngDevNO + 1
    '    If ModAccounts.AddNewDev(LngDevID, LngDevNO, StrTempAccountCode, Val(Me.lbltotal.Caption), _
    '        1, StrTempDes, , , , , Me.XPDtbBill.value, Me.DCboUserName.BoundText, Val(Me.XPTxtBillID.text)) = False Then
    '        GoTo ErrTrap
    '    End If
    'End If
    'If Me.XPChkPayType(2).value = vbChecked Then
    '  '  StrTempAccountCode = "a2a3a2" '√Ê—«ﬁ «·œ›⁄
    '  '  StrTempDes = "⁄œœ " & Me.lbl(19).Caption & "  ‘Ìﬂ«  " & Chr(13)
    '  '  StrTempDes = StrTempDes & "«–‰ «÷«›… —ﬁ„ " & Me.TxtTransSerial.text
    '  '  LngDevNO = LngDevNO + 1
    '  '  If ModAccounts.AddNewDev(LngDevID, LngDevNO, StrTempAccountCode, Val(Me.lbl(18).Caption), _
    '  '      1, StrTempDes, , , , , Me.XPDtbBill.value, Me.DCboUserName.BoundText, Val(Me.XPTxtBillID.text)) = False Then
    '  '      GoTo ErrTrap
    '  '  End If
    'End If
    'If Val(Me.LblDiscountsTotal.Caption) > 0 Then
    '         Account_Code_dynamic = get_account_code_branch(13, my_branch)
    '
    '        If Account_Code_dynamic = "NO branch" Then
    '        MsgBox "·„ Ì „ «‰‘«¡ «·›—⁄", vbCritical
    '        GoTo ErrTrap
    '        Else
    '        If Account_Code_dynamic = "NO account" Then
    '           MsgBox "·„ Ì „  ÕœÌœ Õ”«»     «·Œ’„ «·„ﬂ ”» ›Ì «·›—⁄ ·Â–… «·⁄„·Ì…", vbCritical
    '        GoTo ErrTrap
    '
    '        End If
    '        End If
    '    StrTempAccountCode = Account_Code_dynamic '«·Œ’„ «·„ﬂ ”»13
    '  '  StrTempAccountCode = "a4a4" '«·Œ’„ «·„ﬂ ”»
    '    StrTempDes = "«–‰ «÷«›… —ﬁ„ " & Me.TxtTransSerial.text
    '    LngDevNO = LngDevNO + 1
    '    If ModAccounts.AddNewDev(LngDevID, LngDevNO, StrTempAccountCode, Val(Me.LblDiscountsTotal.Caption), _
    '        1, StrTempDes, , , , , Me.XPDtbBill.value, Me.DCboUserName.BoundText, Val(Me.XPTxtBillID.text)) = False Then
    '        GoTo ErrTrap
    '    End If
    'End If
⁄ —Ì”:
    Cn.CommitTrans
    BeginTrans = False
    If IsSaveWithOutMsg Then Exit Sub
    XPTxtCurrent.Caption = rs.AbsolutePosition
    XPTxtCount.Caption = rs.RecordCount

    Select Case Me.TxtModFlg.text

        Case "N"

            If SystemOptions.UserInterface = ArabicInterface Then
                Msg = "  „ Õ›Ÿ »Ì«‰«  Â–Â «·⁄„·Ì…" & CHR(13)
                Msg = Msg + "Â·  —€» ›Ì ≈÷«›… »Ì«‰«  √Œ—Ì"
            Else
                Msg = " Data Was Saved do you want Another Entry" & CHR(13)
    
            End If

            If MsgBox(Msg, vbYesNo + vbQuestion + vbMsgBoxRight + vbMsgBoxRtlReading + vbDefaultButton2, App.Title) = vbYes Then
                Cmd_Click (0)
                Screen.MousePointer = vbDefault
                Exit Sub
            End If
        
        Case "E"

            If SystemOptions.UserInterface = ArabicInterface Then
                MsgBox " „ Õ›Ÿ Â–Â «· ⁄œÌ·« ", vbInformation + vbOKOnly + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
            Else
                Msg = " chages Was Saved " & CHR(13)
                MsgBox Msg, vbInformation + vbOKOnly + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
            End If
    
            lbl(62).Caption = showLabel(TxtNoteSerial1, oldtxtNoteSerial1)
    
    End Select

    TxtModFlg.text = "R"
    UpdateTransCost val(Me.XPTxtBillID.text)

    If SystemOptions.SysMainStockCostMethod = ModernWeightAverage Then
        '›Ï Õ«·… «‰  ﬂÊ‰ ÿ—Ìﬁ… Õ”«» „ Ê”ÿ «· ﬂ·›…
        'ÂÊ
        'ModernWeightAverage
        '·«»œ «‰ ÌﬁÊ„ «·»—‰«„Ã » ⁄œÌ· ﬁÌ„… „ Ê”ÿ «· ﬂ·›… ··√’‰«›
        '«·„ÊÃÊœ… ›Ï «·›« Ê—…
    End If

    Screen.MousePointer = vbDefault
    Exit Sub
ErrTrap:

    'Stop
    'Resume
    If BeginTrans = True Then
        BeginTrans = False
        Cn.RollbackTrans
    End If

    Screen.MousePointer = vbDefault
    Msg = "⁄›Ê«...ÕœÀ Œÿ√ „« √À‰«¡ Õ›Ÿ Â–Â «·»Ì«‰«  " & CHR(13)
    Msg = Msg & Err.Description & CHR(13)
    Msg = Msg & Err.Number & CHR(13)
    Msg = Msg & Err.Source
    MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
End Sub

Private Sub XPBtnNewClients_Click()

    'With FrmAddNewCustemer
    '    '    .Tag = "x"
    '    .DealingForm = PurchaseTransaction
    '    Set .DcboCustomers = DBCboClientName
    '    .Caption = "≈÷«›… „Ê—œ ÃœÌœ"
    '    .lbl(1).Caption = "ﬂÊœ «·„Ê—œ"
    '    .lbl(0).Caption = "«”„ «·„Ê—œ"
    '    .AddType = 2
    '    .show vbModal
    'End With

End Sub

Private Sub XPCboDiscountType_Change()
    XPCboDiscountType_Click
End Sub

Private Sub XPCboDiscountType_Click()
    On Error GoTo ErrTrap

    If XPCboDiscountType.ListIndex = 0 Or XPCboDiscountType.ListIndex = 3 Or XPCboDiscountType.ListIndex = -1 Then
        lbl(11).Enabled = False
        XPTxtDiscountVal.Enabled = False
        XPTxtDiscountVal.text = ""
    Else
        lbl(11).Enabled = True
        XPTxtDiscountVal.Enabled = True
        XPTxtDiscountVal.text = ""
    End If

    If Me.TxtModFlg.text = "N" Or Me.TxtModFlg.text = "E" Then
        If FG.TextMatrix(1, FG.ColIndex("Code")) <> "" Then
            NewGrid.Calculate 1
        End If
    End If

    Exit Sub
ErrTrap:
End Sub

Private Sub XPChkPayType_Click(Index As Integer)
    On Error GoTo ErrTrap

    Select Case Index

        Case 0

            If XPChkPayType(0).value = Checked Then
                If Me.TxtModFlg.text = "N" Then
                    XPTxtValue(0).text = ""
                    XPTxtSerial(0).text = ""
                End If

                If Me.TxtModFlg.text = "N" Or Me.TxtModFlg.text = "E" Then
                    XPTxtValue(0).Enabled = True
                    '                XPTxtSerial(0).Enabled = True
                    XPTxtValue(0).locked = False
                    '                XPTxtSerial(0).Locked = False
                End If

            Else
                XPTxtValue(0).Enabled = False
                XPTxtValue(0).text = ""
                '            XPTxtSerial(0).Enabled = False
            End If

        Case 1

            If XPChkPayType(1).value = Checked Then
                If Me.TxtModFlg.text = "N" Then
                    XPTxtValue(1).text = ""
                    DtpDelayDate.value = Date
                End If

                If Me.TxtModFlg.text = "N" Or Me.TxtModFlg.text = "E" Then
                    XPTxtValue(1).Enabled = True
                    XPTxtValue(1).locked = False
                    DtpDelayDate.Enabled = True
                Else
                    DtpDelayDate.Enabled = False
                End If

                Me.ChkInstall.Enabled = True
            Else
                XPTxtValue(1).Enabled = False
                XPTxtValue(1).text = ""
                Me.ChkInstall.Enabled = False
            End If

        Case 2

            If XPChkPayType(2).value = Checked And Me.TxtModFlg.text <> "R" Then
                Me.CmdCheque.Enabled = True
            Else
                Me.CmdCheque.Enabled = False
                Me.lbl(18).Caption = 0
                Me.lbl(19).Caption = 0
                Me.FgCheques.rows = Me.FgCheques.FixedRows
            End If

    End Select

    Exit Sub
ErrTrap:
End Sub

Private Sub XPChkTAX_Click()
    On Error GoTo ErrTrap

    If XPChkTAX.value = Checked Then
        XPTxtTaxValue.Enabled = True
        lbl(22).Enabled = True
        lbl(45).Enabled = True
    Else
        XPTxtTaxValue.text = ""
        XPTxtTaxValue.Enabled = False
        lbl(22).Enabled = False
        lbl(45).Enabled = False
    End If

    Exit Sub
ErrTrap:
End Sub

Private Sub XPDtbBill_Change()

    If Trim(TxtNoteSerial1.text) <> "" Then
        oldtxtNoteSerial1.text = TxtNoteSerial1.text
    End If

    TxtNoteSerial.text = ""
    TxtNoteSerial1.text = ""
End Sub

Private Sub XPTab301_Click()
    On Error GoTo ErrTrap

    If Me.TxtModFlg.text <> "R" Then
    
    End If

    Exit Sub
ErrTrap:
End Sub

Private Sub printing()
    On Error GoTo ErrTrap
    Dim BuyReport As ClsBuyReport

    If Not XPTxtBillID.text Then
        Set BuyReport = New ClsBuyReport
        BuyReport.ShowBuyData XPTxtBillID.text, 4, True, LblTotal.Caption
    End If

    Exit Sub
ErrTrap:
 
End Sub

Private Function AvailableDeal() As Boolean
    On Error GoTo ErrTrap
    Dim RowNum As Integer
    Dim Msg As String
    Dim StrSQL As String
    Dim RsTemp As ADODB.Recordset
    Dim RsSalle As ADODB.Recordset
    Dim LngItemID As Long

    For RowNum = 1 To FG.rows - 1

        If FG.TextMatrix(RowNum, FG.ColIndex("Code")) <> "" Then
            StrSQL = "select * From QryDelPurchase where Transaction_Date >=" & SQLDate(XPDtbBill.value, True) & ""
            StrSQL = StrSQL + " and Item_ID=" & FG.TextMatrix(RowNum, FG.ColIndex("Code"))

            '        If FG.TextMatrix(RowNum, FG.ColIndex("HaveSerial")) <> "" Then
            '            If FG.TextMatrix(RowNum, FG.ColIndex("HaveSerial")) = True Then
            If FG.cell(flexcpChecked, RowNum, FG.ColIndex("HaveSerial")) = flexChecked Then
                If FG.TextMatrix(RowNum, FG.ColIndex("Serial")) <> "" Then
                    StrSQL = StrSQL + " and ItemSerial='" & FG.TextMatrix(RowNum, FG.ColIndex("Serial")) & "'"
                End If

                '            End If
            End If

            Set RsSalle = New ADODB.Recordset
            RsSalle.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText

            If Not (RsSalle.EOF Or RsSalle.BOF) Then
                If FG.cell(flexcpChecked, RowNum, FG.ColIndex("HaveSerial")) = flexChecked Then

                    With FrmAlarm
                        .Tag = "x"
                        .DealingForm = PurchaseTransaction
                        .show vbModal
                    End With

                    AvailableDeal = False
                    Exit Function
                    '                End If
                    RsTemp.Close
                Else
                    Set RsTemp = New ADODB.Recordset
                    LngItemID = val(FG.TextMatrix(RowNum, FG.ColIndex("Code")))
                    Set RsTemp = GetItemQuantityStock(LngItemID, Me.DCboStoreName.BoundText, Me.XPDtbBill.value, val(Me.XPTxtBillID.text))

                    If Not (RsTemp.EOF Or RsTemp.BOF) Then
                        If val(RsTemp("QTY").value) < val(FG.TextMatrix(RowNum, FG.ColIndex("Count"))) Then

                            With FrmAlarm
                                .DealingForm = PurchaseTransaction
                                .show vbModal
                            End With

                            AvailableDeal = False
                            Exit Function
                        End If
                    End If

                    RsTemp.Close
                End If
            End If

            RsSalle.Close
        End If

    Next RowNum

    AvailableDeal = True
    Exit Function
ErrTrap:
End Function

Private Sub Form_QueryUnload(Cancel As Integer, _
                             UnloadMode As Integer)
    Dim IntResult As String
    Dim StrMSG As String

    On Error GoTo ErrTrap

    If Me.TxtModFlg.text = "" Then Exit Sub
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

Private Sub CboPayMentType_Change()
    On Error GoTo ErrTrap

    If Me.TxtModFlg.text = "N" Or Me.TxtModFlg.text = "E" Then
        If CboPayMentType.ListIndex = 0 Then
            XPChkPayType(0).Enabled = False
            XPChkPayType(1).Enabled = False
            XPChkPayType(2).Enabled = False
            XPChkPayType(0).value = Checked
            XPChkPayType(1).value = Unchecked
            XPChkPayType(2).value = Unchecked
            XPTxtValue(0).text = XPTxtSum.text
        Else
            XPChkPayType(0).Enabled = True
            XPChkPayType(1).Enabled = True
            XPChkPayType(2).Enabled = True
            XPChkPayType(0).value = Unchecked
            XPChkPayType(1).value = Unchecked
            XPChkPayType(2).value = Unchecked
            XPTxtValue(0).text = ""
        End If
    End If

    Exit Sub
ErrTrap:

End Sub

Private Sub CboPayMentType_Click()
    CboPayMentType_Change
End Sub

Private Sub XPTxtBillID_Change()
    Retrive_Expenses_Vouchers
End Sub

Private Sub XPTxtDiscountVal_KeyPress(KeyAscii As Integer)
    KeyAscii = KeyAscii_Num(KeyAscii, Me.XPTxtDiscountVal.text, 0)
End Sub

Private Sub XPTxtSum_Change()
    On Error GoTo ErrTrap

    If CboPayMentType.ListIndex = 0 Then
        XPChkPayType(0).value = Checked
        XPTxtValue(0).text = XPTxtSum.text
    End If

    Me.LblTotal.Caption = XPTxtSum.text
    Exit Sub
ErrTrap:
End Sub

Public Function RepeatSerial(StrSerial As String, _
                             IntTransType As Integer, _
                             Optional IntTransID As Long = 0, _
                             Optional LngCusID As Long = 0) As Boolean

    Dim rs As ADODB.Recordset
    Dim StrSQL As String
    Dim Msg As String
    RepeatSerial = False

    If SystemOptions.SysDataBaseType = SQLServerDataBase Then
        StrSQL = "SELECT QryTransactionsTotal.Transaction_ID," & "QryTransactionsTotal.TransNet, QryTransactionsTotal.Transaction_Serial, " & "QryTransactionsTotal.Transaction_Date , QryTransactionsTotal.Transaction_Type," & "dbo.TblCustemers.CusName"
        StrSQL = StrSQL + " FROM dbo.QryTransactionsTotal() QryTransactionsTotal INNER JOIN " & "dbo.TblCustemers ON QryTransactionsTotal.CusID = dbo.TblCustemers.CusID"
        StrSQL = StrSQL + " Where QryTransactionsTotal.Transaction_Serial ='" & StrSerial & "'"
        StrSQL = StrSQL + " AND QryTransactionsTotal.Transaction_Type=" & IntTransType & ""

        If LngCusID <> 0 Then
            StrSQL = StrSQL + " AND dbo.TblCustemers.CusID=" & LngCusID & ""
        End If

        If IntTransID <> 0 Then
            StrSQL = StrSQL + " AND QryTransactionsTotal.Transaction_ID <> " & IntTransID & ""
        End If

        Set rs = New ADODB.Recordset
        rs.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

        If Not (rs.BOF Or rs.EOF) Then
            Msg = "—ﬁ„ «·›« Ê—… „ÊÃÊœ „”»ﬁ« ›Ï «·»—‰«„Ã øø" & CHR(13)
            Msg = Msg + "„⁄·Ê„«  ⁄‰ «·›« Ê—… «·„”Ã·…:-" & CHR(13)
        
            Msg = Msg + "—ﬁ„ «·›« Ê—… ›Ï «·»—‰«„Ã:" & rs("Transaction_ID").value & CHR(13)
            Msg = Msg + "„”·”· «·›« Ê—…:" & rs("Transaction_Serial").value & CHR(13)
            Msg = Msg + " «—ÌŒ  ”ÃÌ· «·›« Ê—…:" & rs("Transaction_Date").value & CHR(13)
            Msg = Msg + "«”„ «·⁄„Ì· «Ê «·„Ê—œ:" & rs("CusName").value & CHR(13)
            MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
            RepeatSerial = True
        End If

        rs.Close
        Set rs = Nothing

    End If

End Function

Private Sub SetDefaults()
    Dim StrTemp As String
    Dim RsTemp As ADODB.Recordset

    If SystemOptions.SysPurDateTakeType = InvDateFromLocalCompuer Then
        XPDtbBill.value = Date
    ElseIf SystemOptions.SysPurDateTakeType = InvDateFromServerComputer Then
        StrTemp = "select Getdate() as ServerDate"
        Set RsTemp = New ADODB.Recordset
        RsTemp.Open StrTemp, Cn, adOpenStatic, adLockReadOnly, adCmdText

        If Not (RsTemp.BOF Or RsTemp.EOF) Then
            If Not IsNull(RsTemp("ServerDate").value) Then
                XPDtbBill.value = Format(RsTemp("ServerDate").value, "yyyy/M/d")
            End If

            'XPDtbBill.Value = IIf(IsNull(RsTemp("ServerDate").Value), Date, (RsTemp("ServerDate").Value))
        End If

        RsTemp.Close
        Set RsTemp = Nothing
    End If

    If Not (rs.BOF Or rs.EOF) Then
        rs.MoveLast

        If SystemOptions.SysPurDateTakeType = InvDateFromLastInvDate Then
            XPDtbBill.value = IIf(IsNull(rs("Transaction_Date").value), Date, (rs("Transaction_Date").value))
        End If
    End If

    Me.DcboBox.BoundText = 1
    Me.CboPayMentType.ListIndex = 1

End Sub
