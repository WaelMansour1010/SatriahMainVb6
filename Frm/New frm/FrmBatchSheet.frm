VERSION 5.00
Object = "{C115893A-A3BF-43AF-B28D-69DB846077F3}#1.0#0"; "vsflex8u.ocx"
Object = "{0AFE7BE0-11B7-4A3E-978D-D4501E9A57FE}#1.0#0"; "c1sizer.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{FE5DCFAD-BC1D-11D2-94CF-004005455FAA}#1.4#0"; "ImpulseButton.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#12.0#0"; "SUITEC~1.OCX"
Begin VB.Form FrmBatchSheet 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   10065
   ClientLeft      =   1410
   ClientTop       =   2970
   ClientWidth     =   17475
   Icon            =   "FrmBatchSheet.frx":0000
   LinkTopic       =   "Form3"
   MDIChild        =   -1  'True
   RightToLeft     =   -1  'True
   ScaleHeight     =   10065
   ScaleWidth      =   17475
   WindowState     =   2  'Maximized
   Begin C1SizerLibCtl.C1Elastic C1Elastic1 
      Height          =   10065
      Left            =   0
      TabIndex        =   40
      TabStop         =   0   'False
      Top             =   0
      Width           =   17475
      _cx             =   30824
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
      Begin VB.Frame FraHeader 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         Height          =   660
         Left            =   0
         RightToLeft     =   -1  'True
         TabIndex        =   63
         Top             =   0
         Width           =   17505
         Begin VB.TextBox tXTRootAccount 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   3240
            RightToLeft     =   -1  'True
            TabIndex        =   65
            Top             =   360
            Visible         =   0   'False
            Width           =   2055
         End
         Begin VB.TextBox TxtName 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   0
            Left            =   5880
            RightToLeft     =   -1  'True
            TabIndex        =   64
            Top             =   360
            Visible         =   0   'False
            Width           =   2055
         End
         Begin ImpulseButton.ISButton btnLast 
            Height          =   315
            Left            =   450
            TabIndex        =   66
            Top             =   240
            Width           =   405
            _ExtentX        =   714
            _ExtentY        =   556
            ButtonStyle     =   1
            ButtonPositionImage=   1
            Caption         =   ""
            BackColor       =   16777215
            FontSize        =   12
            FontName        =   "Arial"
            FontBold        =   -1  'True
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   12
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ButtonImage     =   "FrmBatchSheet.frx":6852
            ColorButton     =   16777215
            AcclimateGrayTones=   -1  'True
            DrawFocusRectangle=   0   'False
            DisabledImageExtraction=   0
            DisabledImageStyle=   1
         End
         Begin ImpulseButton.ISButton btnNext 
            Height          =   315
            Left            =   915
            TabIndex        =   67
            Top             =   240
            Width           =   405
            _ExtentX        =   714
            _ExtentY        =   556
            ButtonStyle     =   1
            ButtonPositionImage=   1
            Caption         =   ""
            BackColor       =   16777215
            FontSize        =   12
            FontName        =   "Arial"
            FontBold        =   -1  'True
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   12
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ButtonImage     =   "FrmBatchSheet.frx":6BEC
            ColorButton     =   16777215
            DrawFocusRectangle=   0   'False
            DisabledImageStyle=   1
         End
         Begin ImpulseButton.ISButton btnPrevious 
            Height          =   315
            Left            =   1515
            TabIndex        =   68
            Top             =   240
            Width           =   405
            _ExtentX        =   714
            _ExtentY        =   556
            ButtonStyle     =   1
            ButtonPositionImage=   1
            Caption         =   ""
            BackColor       =   16777215
            FontSize        =   12
            FontName        =   "Arial"
            FontBold        =   -1  'True
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   12
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ButtonImage     =   "FrmBatchSheet.frx":6F86
            ColorButton     =   16777215
            DrawFocusRectangle=   0   'False
            DisabledImageStyle=   1
         End
         Begin ImpulseButton.ISButton btnFirst 
            Height          =   315
            Left            =   2040
            TabIndex        =   69
            Top             =   240
            Width           =   405
            _ExtentX        =   714
            _ExtentY        =   556
            ButtonStyle     =   1
            ButtonPositionImage=   1
            Caption         =   ""
            BackColor       =   16777215
            FontSize        =   12
            FontName        =   "Arial"
            FontBold        =   -1  'True
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   12
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ButtonImage     =   "FrmBatchSheet.frx":7320
            ColorButton     =   16777215
            DrawFocusRectangle=   0   'False
            DisabledImageStyle=   1
         End
         Begin VB.Image Image1 
            Height          =   615
            Left            =   13200
            Picture         =   "FrmBatchSheet.frx":76BA
            Stretch         =   -1  'True
            Top             =   120
            Visible         =   0   'False
            Width           =   735
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Ê—ﬁ…  ’ÕÌÕ"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   18
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   495
            Index           =   2
            Left            =   8880
            RightToLeft     =   -1  'True
            TabIndex        =   70
            Top             =   120
            Width           =   4080
         End
      End
      Begin C1SizerLibCtl.C1Elastic C1Elastic2 
         Height          =   735
         Left            =   0
         TabIndex        =   41
         TabStop         =   0   'False
         Top             =   600
         Width           =   17520
         _cx             =   30903
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
         Begin VB.TextBox TxtCode 
            Alignment       =   1  'Right Justify
            Height          =   315
            Left            =   120
            RightToLeft     =   -1  'True
            TabIndex        =   103
            Top             =   0
            Visible         =   0   'False
            Width           =   1545
         End
         Begin VB.TextBox TxtRefNo 
            Alignment       =   1  'Right Justify
            Height          =   315
            Left            =   240
            RightToLeft     =   -1  'True
            TabIndex        =   4
            Top             =   240
            Width           =   1665
         End
         Begin VB.TextBox TxtBatchNo 
            Alignment       =   1  'Right Justify
            Height          =   315
            Left            =   2880
            RightToLeft     =   -1  'True
            TabIndex        =   3
            Top             =   240
            Width           =   1545
         End
         Begin VB.TextBox TxtSerial1 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFFC0&
            Enabled         =   0   'False
            Height          =   315
            Left            =   13920
            RightToLeft     =   -1  'True
            TabIndex        =   0
            Top             =   240
            Width           =   2040
         End
         Begin MSComCtl2.DTPicker XPDtbTrans 
            Height          =   315
            Left            =   11385
            TabIndex        =   1
            Top             =   240
            Width           =   1590
            _ExtentX        =   2805
            _ExtentY        =   556
            _Version        =   393216
            Format          =   94175233
            CurrentDate     =   38784
         End
         Begin MSDataListLib.DataCombo Dcbranch 
            Bindings        =   "FrmBatchSheet.frx":8ABF
            Height          =   315
            Left            =   5640
            TabIndex        =   2
            Top             =   240
            Width           =   5040
            _ExtentX        =   8890
            _ExtentY        =   556
            _Version        =   393216
            Style           =   2
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
            Alignment       =   2  'Center
            BackColor       =   &H00E2E9E9&
            Caption         =   "—ﬁ„ «·„—Ã⁄"
            Height          =   285
            Index           =   6
            Left            =   2055
            TabIndex        =   74
            Top             =   240
            Width           =   750
         End
         Begin VB.Label lbl 
            Alignment       =   2  'Center
            BackColor       =   &H00E2E9E9&
            Caption         =   "—ﬁ„ «· ’ÕÌÕ"
            Height          =   285
            Index           =   5
            Left            =   4575
            TabIndex        =   73
            Top             =   240
            Width           =   870
         End
         Begin VB.Label lbl 
            Alignment       =   2  'Center
            BackColor       =   &H00E2E9E9&
            Caption         =   "«· «—ÌŒ"
            Height          =   285
            Index           =   2
            Left            =   13035
            TabIndex        =   44
            Top             =   240
            Width           =   900
         End
         Begin VB.Label lbl 
            Alignment       =   2  'Center
            BackColor       =   &H00E2E9E9&
            Caption         =   "—ﬁ„ «·Õ—ﬂ…"
            Height          =   285
            Index           =   4
            Left            =   16305
            RightToLeft     =   -1  'True
            TabIndex        =   43
            Top             =   240
            Width           =   930
         End
         Begin VB.Label lbl 
            Alignment       =   2  'Center
            BackColor       =   &H00E2E9E9&
            Caption         =   "«·›—⁄"
            Height          =   285
            Index           =   7
            Left            =   10200
            TabIndex        =   42
            Top             =   240
            Width           =   1620
         End
      End
      Begin C1SizerLibCtl.C1Elastic C1Elastic4 
         Height          =   2895
         Left            =   0
         TabIndex        =   45
         TabStop         =   0   'False
         Top             =   1440
         Width           =   17520
         _cx             =   30903
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
         Begin VB.TextBox Text4 
            Alignment       =   1  'Right Justify
            Height          =   315
            Left            =   5805
            TabIndex        =   7
            TabStop         =   0   'False
            Top             =   120
            Width           =   1440
         End
         Begin VB.TextBox TxtRemarks 
            Alignment       =   1  'Right Justify
            Height          =   315
            Left            =   120
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   11
            TabStop         =   0   'False
            Top             =   480
            Width           =   7155
         End
         Begin VB.TextBox TxtSpecSeries 
            Alignment       =   1  'Right Justify
            Height          =   315
            Left            =   8880
            TabIndex        =   24
            TabStop         =   0   'False
            Top             =   2400
            Width           =   2880
         End
         Begin VB.TextBox TxtMasterNo 
            Alignment       =   1  'Right Justify
            Height          =   315
            Left            =   13140
            TabIndex        =   23
            TabStop         =   0   'False
            Top             =   2400
            Width           =   2880
         End
         Begin VB.Frame Frame1 
            BackColor       =   &H00E2E9E9&
            Height          =   735
            Left            =   120
            RightToLeft     =   -1  'True
            TabIndex        =   81
            Top             =   840
            Width           =   8655
            Begin VB.TextBox TxtNoPermix2 
               Alignment       =   1  'Right Justify
               Height          =   315
               Left            =   2280
               TabIndex        =   104
               TabStop         =   0   'False
               Top             =   0
               Visible         =   0   'False
               Width           =   960
            End
            Begin VB.TextBox TxtMinQty 
               Alignment       =   1  'Right Justify
               Height          =   315
               Left            =   6480
               Locked          =   -1  'True
               TabIndex        =   101
               TabStop         =   0   'False
               Top             =   240
               Width           =   960
            End
            Begin VB.TextBox TxtNoPermix 
               Alignment       =   1  'Right Justify
               Height          =   315
               Left            =   120
               TabIndex        =   18
               TabStop         =   0   'False
               Top             =   240
               Width           =   960
            End
            Begin VB.TextBox TxtQtyPermix 
               Alignment       =   1  'Right Justify
               Height          =   315
               Left            =   1920
               TabIndex        =   17
               TabStop         =   0   'False
               Top             =   240
               Width           =   960
            End
            Begin VB.TextBox TxtPlanQty 
               Alignment       =   1  'Right Justify
               Height          =   315
               Left            =   4080
               TabIndex        =   16
               TabStop         =   0   'False
               Top             =   240
               Width           =   960
            End
            Begin VB.Label lbl 
               Alignment       =   2  'Center
               BackColor       =   &H00E2E9E9&
               Caption         =   "«ﬁ· ﬂ„Ì…"
               Height          =   195
               Index           =   26
               Left            =   7320
               RightToLeft     =   -1  'True
               TabIndex        =   102
               Top             =   240
               Width           =   1260
            End
            Begin VB.Label lbl 
               Alignment       =   2  'Center
               BackColor       =   &H00E2E9E9&
               Caption         =   "«·⁄œœ"
               Height          =   195
               Index           =   11
               Left            =   930
               RightToLeft     =   -1  'True
               TabIndex        =   84
               Top             =   240
               Width           =   1245
            End
            Begin VB.Label lbl 
               Alignment       =   2  'Center
               BackColor       =   &H00E2E9E9&
               Caption         =   "«·”⁄…"
               Height          =   195
               Index           =   12
               Left            =   2880
               RightToLeft     =   -1  'True
               TabIndex        =   83
               Top             =   240
               Width           =   1140
            End
            Begin VB.Label lbl 
               Alignment       =   2  'Center
               BackColor       =   &H00E2E9E9&
               Caption         =   "ﬂ„Ì… «·Œÿ…"
               Height          =   195
               Index           =   10
               Left            =   5160
               RightToLeft     =   -1  'True
               TabIndex        =   82
               Top             =   240
               Width           =   1260
            End
         End
         Begin VB.Frame Frame10 
            BackColor       =   &H00E2E9E9&
            Caption         =   "»‰«¡ ⁄·Ï"
            Height          =   735
            Left            =   8880
            RightToLeft     =   -1  'True
            TabIndex        =   78
            Top             =   840
            Width           =   8655
            Begin VB.CommandButton Command1 
               Caption         =   "View"
               Height          =   330
               Left            =   0
               RightToLeft     =   -1  'True
               TabIndex        =   96
               Top             =   270
               Width           =   615
            End
            Begin XtremeSuiteControls.RadioButton BasedRd 
               Height          =   375
               Index           =   0
               Left            =   6600
               TabIndex        =   12
               Top             =   240
               Width           =   1455
               _Version        =   786432
               _ExtentX        =   2566
               _ExtentY        =   661
               _StockProps     =   79
               Caption         =   "«Ê«„— «·‘€·"
               BackColor       =   14871017
               UseVisualStyle  =   -1  'True
               TextAlignment   =   1
               RightToLeft     =   -1  'True
            End
            Begin MSComCtl2.DTPicker FromDate 
               Height          =   330
               Left            =   3000
               TabIndex        =   14
               Top             =   270
               Width           =   1575
               _ExtentX        =   2778
               _ExtentY        =   582
               _Version        =   393216
               Format          =   94175235
               CurrentDate     =   41640
            End
            Begin MSComCtl2.DTPicker ToDate 
               Height          =   330
               Left            =   720
               TabIndex        =   15
               Top             =   270
               Width           =   1575
               _ExtentX        =   2778
               _ExtentY        =   582
               _Version        =   393216
               Format          =   94175235
               CurrentDate     =   41640
            End
            Begin XtremeSuiteControls.RadioButton BasedRd 
               Height          =   375
               Index           =   1
               Left            =   5160
               TabIndex        =   13
               Top             =   240
               Width           =   1335
               _Version        =   786432
               _ExtentX        =   2355
               _ExtentY        =   661
               _StockProps     =   79
               Caption         =   "Œÿ… «·«‰ «Ã"
               BackColor       =   14871017
               UseVisualStyle  =   -1  'True
               TextAlignment   =   1
               RightToLeft     =   -1  'True
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "„‰"
               Height          =   375
               Index           =   15
               Left            =   4530
               RightToLeft     =   -1  'True
               TabIndex        =   80
               Top             =   240
               Width           =   540
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "≈·Ï"
               Height          =   195
               Index           =   13
               Left            =   2310
               RightToLeft     =   -1  'True
               TabIndex        =   79
               Top             =   240
               Width           =   480
            End
         End
         Begin VB.TextBox Text3 
            Alignment       =   1  'Right Justify
            Height          =   315
            Left            =   14565
            TabIndex        =   9
            TabStop         =   0   'False
            Top             =   480
            Width           =   1440
         End
         Begin VB.TextBox TxtItemCode 
            Alignment       =   1  'Right Justify
            Height          =   315
            Left            =   14565
            TabIndex        =   5
            TabStop         =   0   'False
            Top             =   120
            Width           =   1440
         End
         Begin MSDataListLib.DataCombo DcbItem 
            Bindings        =   "FrmBatchSheet.frx":8AD4
            Height          =   315
            Left            =   8880
            TabIndex        =   6
            Top             =   120
            Width           =   5580
            _ExtentX        =   9843
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
         Begin MSDataListLib.DataCombo DcbGroup 
            Bindings        =   "FrmBatchSheet.frx":8AE9
            Height          =   315
            Left            =   8880
            TabIndex        =   10
            Top             =   480
            Width           =   5580
            _ExtentX        =   9843
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
         Begin MSDataListLib.DataCombo DcbItemFinish 
            Bindings        =   "FrmBatchSheet.frx":8AFE
            Height          =   315
            Left            =   120
            TabIndex        =   8
            Top             =   120
            Width           =   5580
            _ExtentX        =   9843
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
         Begin MSDataListLib.DataCombo DcbPermixerMach 
            Bindings        =   "FrmBatchSheet.frx":8B13
            Height          =   315
            Left            =   8880
            TabIndex        =   19
            Top             =   1680
            Width           =   7140
            _ExtentX        =   12594
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
         Begin MSDataListLib.DataCombo DcbExtruderMach 
            Bindings        =   "FrmBatchSheet.frx":8B28
            Height          =   315
            Left            =   120
            TabIndex        =   20
            Top             =   1680
            Width           =   7140
            _ExtentX        =   12594
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
         Begin MSDataListLib.DataCombo DcbBlenderMach 
            Bindings        =   "FrmBatchSheet.frx":8B3D
            Height          =   315
            Left            =   120
            TabIndex        =   22
            Top             =   2040
            Width           =   7140
            _ExtentX        =   12594
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
         Begin MSDataListLib.DataCombo DcbGrinderMach 
            Bindings        =   "FrmBatchSheet.frx":8B52
            Height          =   315
            Left            =   8880
            TabIndex        =   21
            Top             =   2040
            Width           =   7140
            _ExtentX        =   12594
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
         Begin MSDataListLib.DataCombo DcbCustomer 
            Bindings        =   "FrmBatchSheet.frx":8B67
            Height          =   315
            Left            =   7320
            TabIndex        =   107
            Top             =   480
            Visible         =   0   'False
            Width           =   1500
            _ExtentX        =   2646
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
            Alignment       =   2  'Center
            BackColor       =   &H00E2E9E9&
            Caption         =   "„·«ÕŸ« "
            Height          =   285
            Index           =   21
            Left            =   7200
            RightToLeft     =   -1  'True
            TabIndex        =   91
            Top             =   480
            Width           =   1410
         End
         Begin VB.Label lbl 
            Alignment       =   2  'Center
            BackColor       =   &H00E2E9E9&
            Caption         =   "—ﬁ„ «· Œ’Ì’"
            Height          =   285
            Index           =   20
            Left            =   11640
            RightToLeft     =   -1  'True
            TabIndex        =   90
            Top             =   2400
            Width           =   1410
         End
         Begin VB.Label lbl 
            Alignment       =   2  'Center
            BackColor       =   &H00E2E9E9&
            Caption         =   "—ﬁ„ «· Õﬂ„"
            Height          =   285
            Index           =   19
            Left            =   16080
            RightToLeft     =   -1  'True
            TabIndex        =   89
            Top             =   2400
            Width           =   1290
         End
         Begin VB.Label lbl 
            Alignment       =   2  'Center
            BackColor       =   &H00E2E9E9&
            Caption         =   "«·… «·ﬂ»”"
            Height          =   285
            Index           =   18
            Left            =   7320
            RightToLeft     =   -1  'True
            TabIndex        =   88
            Top             =   1680
            Width           =   1410
         End
         Begin VB.Label lbl 
            Alignment       =   2  'Center
            BackColor       =   &H00E2E9E9&
            Caption         =   "«·… «·Œ·ÿ"
            Height          =   285
            Index           =   17
            Left            =   7320
            RightToLeft     =   -1  'True
            TabIndex        =   87
            Top             =   2040
            Width           =   1410
         End
         Begin VB.Label lbl 
            Alignment       =   2  'Center
            BackColor       =   &H00E2E9E9&
            Caption         =   "«·… «·Œ·ÿ «·»œ«∆Ì"
            Height          =   285
            Index           =   16
            Left            =   16080
            RightToLeft     =   -1  'True
            TabIndex        =   86
            Top             =   1680
            Width           =   1290
         End
         Begin VB.Label lbl 
            Alignment       =   2  'Center
            BackColor       =   &H00E2E9E9&
            Caption         =   "«·… «·ÿÕ‰"
            Height          =   285
            Index           =   14
            Left            =   16080
            RightToLeft     =   -1  'True
            TabIndex        =   85
            Top             =   2040
            Width           =   1290
         End
         Begin VB.Label lbl 
            Alignment       =   2  'Center
            BackColor       =   &H00E2E9E9&
            Caption         =   "«·’‰› «·‰Â«∆Ì"
            Height          =   285
            Index           =   3
            Left            =   7560
            RightToLeft     =   -1  'True
            TabIndex        =   77
            Top             =   120
            Width           =   1050
         End
         Begin VB.Label lbl 
            Alignment       =   2  'Center
            BackColor       =   &H00E2E9E9&
            Caption         =   "«·„Ã„Ê⁄…"
            Height          =   285
            Index           =   1
            Left            =   16080
            RightToLeft     =   -1  'True
            TabIndex        =   76
            Top             =   480
            Width           =   1290
         End
         Begin VB.Label lbl 
            Alignment       =   2  'Center
            BackColor       =   &H00E2E9E9&
            Caption         =   "«·’‰›"
            Height          =   285
            Index           =   0
            Left            =   16080
            RightToLeft     =   -1  'True
            TabIndex        =   75
            Top             =   120
            Width           =   1290
         End
      End
      Begin C1SizerLibCtl.C1Elastic C1Elastic5 
         Height          =   615
         Left            =   0
         TabIndex        =   46
         TabStop         =   0   'False
         Top             =   8835
         Width           =   17475
         _cx             =   30824
         _cy             =   1085
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
         Begin C1SizerLibCtl.C1Elastic C1Elastic6 
            Height          =   405
            Left            =   360
            TabIndex        =   49
            TabStop         =   0   'False
            Top             =   105
            Width           =   4080
            _cx             =   7197
            _cy             =   714
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
            Begin VB.Label Label2 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "«·”Ã· «·Õ«·Ì:"
               Height          =   210
               Index           =   0
               Left            =   2970
               RightToLeft     =   -1  'True
               TabIndex        =   53
               Top             =   120
               Width           =   975
            End
            Begin VB.Label Label2 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "⁄œœ «·”Ã·« :"
               Height          =   210
               Index           =   1
               Left            =   1050
               RightToLeft     =   -1  'True
               TabIndex        =   52
               Top             =   120
               Width           =   975
            End
            Begin VB.Label LabCurrRec 
               Alignment       =   2  'Center
               BackColor       =   &H00E2E9E9&
               ForeColor       =   &H00800000&
               Height          =   210
               Left            =   2145
               RightToLeft     =   -1  'True
               TabIndex        =   51
               Top             =   135
               Width           =   675
            End
            Begin VB.Label LabCountRec 
               Alignment       =   2  'Center
               BackColor       =   &H00E2E9E9&
               ForeColor       =   &H00C00000&
               Height          =   210
               Left            =   120
               RightToLeft     =   -1  'True
               TabIndex        =   50
               Top             =   120
               Width           =   780
            End
         End
         Begin MSDataListLib.DataCombo DCboUserName 
            Height          =   315
            Left            =   7785
            TabIndex        =   47
            Top             =   105
            Width           =   5925
            _ExtentX        =   10451
            _ExtentY        =   556
            _Version        =   393216
            Enabled         =   0   'False
            Text            =   ""
            RightToLeft     =   -1  'True
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Õ—— »Ê«”ÿ…  "
            Height          =   225
            Index           =   8
            Left            =   14010
            TabIndex        =   48
            Top             =   105
            Width           =   1485
         End
      End
      Begin C1SizerLibCtl.C1Elastic C1Elastic7 
         Height          =   615
         Left            =   0
         TabIndex        =   54
         TabStop         =   0   'False
         Top             =   9450
         Width           =   17475
         _cx             =   30824
         _cy             =   1085
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
         Align           =   2
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
         Begin ImpulseButton.ISButton btnNew 
            Height          =   330
            Left            =   15795
            TabIndex        =   55
            ToolTipText     =   "· ”ÃÌ· »Ì«‰«  ÃœÌœ…"
            Top             =   120
            Width           =   1230
            _ExtentX        =   2170
            _ExtentY        =   582
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
            ButtonImage     =   "FrmBatchSheet.frx":8B7C
            ColorButton     =   14871017
            DrawFocusRectangle=   0   'False
         End
         Begin ImpulseButton.ISButton btnSave 
            Height          =   330
            Left            =   11970
            TabIndex        =   56
            ToolTipText     =   "Õ›Ÿ «·»Ì«‰«  œ«Œ· ﬁ«⁄œ… «·»Ì«‰« "
            Top             =   120
            Width           =   1335
            _ExtentX        =   2355
            _ExtentY        =   582
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
            ButtonImage     =   "FrmBatchSheet.frx":F3DE
            ColorButton     =   14871017
            DrawFocusRectangle=   0   'False
         End
         Begin ImpulseButton.ISButton btnModify 
            Height          =   330
            Left            =   14250
            TabIndex        =   57
            ToolTipText     =   "· ⁄œÌ· «·»Ì«‰«  «·Õ«·Ì…"
            Top             =   120
            Width           =   1215
            _ExtentX        =   2143
            _ExtentY        =   582
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
            ButtonImage     =   "FrmBatchSheet.frx":F778
            ColorButton     =   14871017
            DrawFocusRectangle=   0   'False
         End
         Begin ImpulseButton.ISButton BtnUndo 
            Height          =   330
            Left            =   10290
            TabIndex        =   58
            ToolTipText     =   "·· —«Ã⁄ ⁄‰ «·ÕœÀ Ê«·—ÃÊ⁄ «·Ï «·Ê÷⁄ «·ÿ»Ì⁄Ì"
            Top             =   120
            Width           =   1455
            _ExtentX        =   2566
            _ExtentY        =   582
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
            ButtonImage     =   "FrmBatchSheet.frx":15FDA
            ColorButton     =   14871017
            DrawFocusRectangle=   0   'False
         End
         Begin ImpulseButton.ISButton btnDelete 
            Height          =   330
            Left            =   8625
            TabIndex        =   59
            ToolTipText     =   "Õ–› «·»Ì«‰«  «·„Õœœ…"
            Top             =   120
            Width           =   1320
            _ExtentX        =   2328
            _ExtentY        =   582
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
            ButtonImage     =   "FrmBatchSheet.frx":16374
            ColorButton     =   14871017
            DrawFocusRectangle=   0   'False
         End
         Begin ImpulseButton.ISButton btnCancel 
            Height          =   330
            Left            =   0
            TabIndex        =   60
            ToolTipText     =   "«·Œ—ÊÃ «·Ï  «·‰«›–… «·—∆Ì”Ì…"
            Top             =   120
            Width           =   1215
            _ExtentX        =   2143
            _ExtentY        =   582
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
            ButtonImage     =   "FrmBatchSheet.frx":1690E
            ColorButton     =   14871017
            DrawFocusRectangle=   0   'False
            DisabledImageStyle=   1
         End
         Begin ImpulseButton.ISButton ISButton5 
            Height          =   405
            Left            =   7305
            TabIndex        =   61
            TabStop         =   0   'False
            ToolTipText     =   "ÿ»«⁄… «·»Ì«‰«  "
            Top             =   120
            Width           =   1005
            _ExtentX        =   1773
            _ExtentY        =   714
            ButtonStyle     =   1
            ButtonPositionImage=   1
            Caption         =   "ÿ»«⁄… "
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
            ButtonImage     =   "FrmBatchSheet.frx":16CA8
            ColorButton     =   14871017
            DrawFocusRectangle=   0   'False
            DisabledImageStyle=   1
         End
         Begin ImpulseButton.ISButton ISButton8 
            Height          =   330
            Left            =   4185
            TabIndex        =   62
            TabStop         =   0   'False
            ToolTipText     =   "(Ctrl+F)  ··»ÕÀ ≈÷€ÿ Â–« «·„› «Õ √Ê ≈÷€ÿ "
            Top             =   120
            Visible         =   0   'False
            Width           =   960
            _ExtentX        =   1693
            _ExtentY        =   582
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
            ButtonImage     =   "FrmBatchSheet.frx":1D50A
            ColorButton     =   14871017
            DrawFocusRectangle=   0   'False
         End
         Begin ImpulseButton.ISButton ISButton2 
            Height          =   405
            Left            =   5265
            TabIndex        =   94
            TabStop         =   0   'False
            ToolTipText     =   "ÿ»«⁄… «·»Ì«‰«  "
            Top             =   120
            Width           =   1575
            _ExtentX        =   2778
            _ExtentY        =   714
            ButtonStyle     =   1
            ButtonPositionImage=   1
            Caption         =   "ÿ»«⁄… «· Œ’Ì’"
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
            ButtonImage     =   "FrmBatchSheet.frx":1D8A4
            ColorButton     =   14871017
            DrawFocusRectangle=   0   'False
            DisabledImageStyle=   1
         End
         Begin ImpulseButton.ISButton Accredit 
            Height          =   345
            Left            =   1440
            TabIndex        =   95
            Top             =   120
            Width           =   2745
            _ExtentX        =   4842
            _ExtentY        =   609
            ButtonPositionImage=   1
            Caption         =   " ÕÊÌ· «·Ï «„— «‰ «Ã"
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
      End
      Begin C1SizerLibCtl.C1Elastic C1Elastic3 
         Height          =   4575
         Left            =   0
         TabIndex        =   71
         TabStop         =   0   'False
         Top             =   4290
         Width           =   17475
         _cx             =   30824
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
            Height          =   225
            Index           =   0
            Left            =   15780
            TabIndex        =   25
            Top             =   2430
            Width           =   1530
            _ExtentX        =   2699
            _ExtentY        =   397
            ButtonStyle     =   1
            ButtonPositionImage=   1
            Caption         =   " Õ–› ”ÿ—"
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
            ButtonImage     =   "FrmBatchSheet.frx":24106
            ColorButton     =   14871017
            DrawFocusRectangle=   0   'False
         End
         Begin ImpulseButton.ISButton Cmd 
            Height          =   225
            Index           =   1
            Left            =   14115
            TabIndex        =   26
            Top             =   2430
            Width           =   1500
            _ExtentX        =   2646
            _ExtentY        =   397
            ButtonStyle     =   1
            ButtonPositionImage=   1
            Caption         =   " Õ–› «·ﬂ·"
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
            ButtonImage     =   "FrmBatchSheet.frx":246A0
            ColorButton     =   14871017
            DrawFocusRectangle=   0   'False
         End
         Begin VSFlex8UCtl.VSFlexGrid VSFlexGrid1 
            Height          =   1425
            Left            =   120
            TabIndex        =   72
            Top             =   2760
            Width           =   17265
            _cx             =   30454
            _cy             =   2514
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
            Rows            =   12
            Cols            =   6
            FixedRows       =   1
            FixedCols       =   1
            RowHeightMin    =   320
            RowHeightMax    =   0
            ColWidthMin     =   0
            ColWidthMax     =   0
            ExtendLastCol   =   -1  'True
            FormatString    =   $"FrmBatchSheet.frx":24C3A
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
         Begin VSFlex8UCtl.VSFlexGrid FG 
            Height          =   2235
            Left            =   0
            TabIndex        =   92
            Top             =   0
            Width           =   17310
            _cx             =   30533
            _cy             =   3942
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
            Rows            =   1
            Cols            =   15
            FixedRows       =   1
            FixedCols       =   1
            RowHeightMin    =   300
            RowHeightMax    =   0
            ColWidthMin     =   0
            ColWidthMax     =   0
            ExtendLastCol   =   -1  'True
            FormatString    =   $"FrmBatchSheet.frx":24D0E
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
            WallPaperAlignment=   0
            AccessibleName  =   ""
            AccessibleDescription=   ""
            AccessibleValue =   ""
            AccessibleRole  =   24
         End
         Begin ImpulseButton.ISButton Cmd 
            Height          =   225
            Index           =   2
            Left            =   15795
            TabIndex        =   27
            Top             =   4200
            Width           =   1530
            _ExtentX        =   2699
            _ExtentY        =   397
            ButtonStyle     =   1
            ButtonPositionImage=   1
            Caption         =   " Õ–› ”ÿ—"
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
            ButtonImage     =   "FrmBatchSheet.frx":24F09
            ColorButton     =   14871017
            DrawFocusRectangle=   0   'False
         End
         Begin ImpulseButton.ISButton Cmd 
            Height          =   225
            Index           =   3
            Left            =   14130
            TabIndex        =   28
            Top             =   4200
            Width           =   1500
            _ExtentX        =   2646
            _ExtentY        =   397
            ButtonStyle     =   1
            ButtonPositionImage=   1
            Caption         =   " Õ–› «·ﬂ·"
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
            ButtonImage     =   "FrmBatchSheet.frx":254A3
            ColorButton     =   14871017
            DrawFocusRectangle=   0   'False
         End
         Begin VB.Label lbl 
            Alignment       =   2  'Center
            BackColor       =   &H00E2E9E9&
            Caption         =   "«Ã„«·Ì «· ﬂ·›…"
            Height          =   285
            Index           =   28
            Left            =   4800
            RightToLeft     =   -1  'True
            TabIndex        =   106
            Top             =   2400
            Width           =   1410
         End
         Begin VB.Label lbl 
            Alignment       =   2  'Center
            BackColor       =   &H00E2E9E9&
            Caption         =   "0"
            Height          =   285
            Index           =   27
            Left            =   3120
            RightToLeft     =   -1  'True
            TabIndex        =   105
            Top             =   2400
            Width           =   1410
         End
         Begin VB.Label lbl 
            Alignment       =   2  'Center
            BackColor       =   &H00E2E9E9&
            Caption         =   "0"
            Height          =   285
            Index           =   25
            Left            =   120
            RightToLeft     =   -1  'True
            TabIndex        =   100
            Top             =   2400
            Width           =   1770
         End
         Begin VB.Label lbl 
            Alignment       =   2  'Center
            BackColor       =   &H00E2E9E9&
            Caption         =   "«·«Ã„«·Ì"
            Height          =   285
            Index           =   24
            Left            =   2040
            RightToLeft     =   -1  'True
            TabIndex        =   99
            Top             =   2400
            Width           =   810
         End
         Begin VB.Label lbl 
            Alignment       =   2  'Center
            BackColor       =   &H00E2E9E9&
            Caption         =   "0"
            Height          =   285
            Index           =   23
            Left            =   6105
            RightToLeft     =   -1  'True
            TabIndex        =   98
            Top             =   2400
            Width           =   1410
         End
         Begin VB.Label lbl 
            Alignment       =   2  'Center
            BackColor       =   &H00E2E9E9&
            Caption         =   "«Ã„«·Ì «·ﬂ„Ì« "
            Height          =   285
            Index           =   9
            Left            =   7785
            RightToLeft     =   -1  'True
            TabIndex        =   97
            Top             =   2400
            Width           =   1410
         End
         Begin VB.Label lbl 
            Alignment       =   2  'Center
            BackColor       =   &H00E2E9E9&
            Caption         =   "»Ì«‰«  «· Œ’Ì’"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C00000&
            Height          =   405
            Index           =   22
            Left            =   8865
            RightToLeft     =   -1  'True
            TabIndex        =   93
            Top             =   2400
            Width           =   3555
         End
      End
   End
   Begin VB.TextBox TxtModFlg 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H0000FF00&
      Enabled         =   0   'False
      Height          =   285
      Left            =   18480
      RightToLeft     =   -1  'True
      TabIndex        =   33
      Text            =   "modflag"
      Top             =   4200
      Visible         =   0   'False
      Width           =   465
   End
   Begin VB.ComboBox CmbType 
      BackColor       =   &H80000018&
      Height          =   315
      ItemData        =   "FrmBatchSheet.frx":25A3D
      Left            =   18360
      List            =   "FrmBatchSheet.frx":25A4D
      RightToLeft     =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   32
      Top             =   3120
      Visible         =   0   'False
      Width           =   1005
   End
   Begin VB.TextBox TxtVac_ID 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H000000FF&
      Height          =   240
      Left            =   18480
      RightToLeft     =   -1  'True
      TabIndex        =   31
      Top             =   1200
      Visible         =   0   'False
      Width           =   945
   End
   Begin VB.Frame Frmo2 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      Height          =   375
      Left            =   18480
      RightToLeft     =   -1  'True
      TabIndex        =   30
      Top             =   1680
      Visible         =   0   'False
      Width           =   3105
   End
   Begin VB.TextBox Emp_id 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   18600
      RightToLeft     =   -1  'True
      TabIndex        =   29
      Top             =   5640
      Visible         =   0   'False
      Width           =   1575
   End
   Begin MSDataListLib.DataCombo DCUser 
      CausesValidation=   0   'False
      Height          =   315
      Left            =   18720
      TabIndex        =   34
      Tag             =   "„‰ ›÷·ﬂ √œŒ· —ﬁ„ «·ﬁ÷Ì…"
      Top             =   960
      Width           =   2100
      _ExtentX        =   3704
      _ExtentY        =   556
      _Version        =   393216
      Enabled         =   0   'False
      BackColor       =   -2147483624
      Text            =   ""
      RightToLeft     =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSDataListLib.DataCombo DCPreFix 
      Height          =   315
      Left            =   18360
      TabIndex        =   35
      Top             =   2280
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   556
      _Version        =   393216
      BackColor       =   16777215
      Text            =   ""
      RightToLeft     =   -1  'True
   End
   Begin MSComctlLib.ImageList GrdImageList 
      Left            =   18480
      Top             =   3720
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   8
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmBatchSheet.frx":25A66
            Key             =   "CompanyName"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmBatchSheet.frx":25E00
            Key             =   "Ser"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmBatchSheet.frx":2619A
            Key             =   "Vac_Name"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmBatchSheet.frx":26534
            Key             =   "ShareCount"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmBatchSheet.frx":268CE
            Key             =   "Dis_Count"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmBatchSheet.frx":26C68
            Key             =   "Bouns"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmBatchSheet.frx":27002
            Key             =   "SharesValue"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmBatchSheet.frx":2759C
            Key             =   "BuyValue"
         EndProperty
      EndProperty
   End
   Begin ImpulseButton.ISButton BtnUpdate 
      Height          =   330
      Left            =   18480
      TabIndex        =   36
      TabStop         =   0   'False
      ToolTipText     =   " ÕœÌÀ ﬁ«⁄œ… «·»Ì«‰« "
      Top             =   5040
      Visible         =   0   'False
      Width           =   1200
      _ExtentX        =   2117
      _ExtentY        =   582
      ButtonStyle     =   1
      ButtonPositionImage=   1
      Caption         =   " ÕœÌÀ"
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
      ButtonImage     =   "FrmBatchSheet.frx":27936
      ColorButton     =   14871017
      DrawFocusRectangle=   0   'False
   End
   Begin ImpulseButton.ISButton ISButton1 
      Height          =   405
      Left            =   18840
      TabIndex        =   38
      TabStop         =   0   'False
      ToolTipText     =   "ÿ»«⁄… «·»Ì«‰«  "
      Top             =   120
      Width           =   1245
      _ExtentX        =   2196
      _ExtentY        =   714
      ButtonStyle     =   1
      ButtonPositionImage=   1
      Caption         =   "ÿ»«⁄… "
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
      ButtonImage     =   "FrmBatchSheet.frx":2E198
      ColorButton     =   14871017
      DrawFocusRectangle=   0   'False
      DisabledImageStyle=   1
   End
   Begin ImpulseButton.ISButton btnQuery 
      Height          =   330
      Left            =   19800
      TabIndex        =   39
      TabStop         =   0   'False
      ToolTipText     =   "(Ctrl+F)  ··»ÕÀ ≈÷€ÿ Â–« «·„› «Õ √Ê ≈÷€ÿ "
      Top             =   120
      Width           =   1320
      _ExtentX        =   2328
      _ExtentY        =   582
      ButtonStyle     =   1
      ButtonPositionImage=   1
      Caption         =   "»ÕÀ"
      BackColor       =   14871017
      FontSize        =   9.75
      FontName        =   "Arial"
      FontBold        =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ButtonImage     =   "FrmBatchSheet.frx":349FA
      ColorButton     =   14871017
      DrawFocusRectangle=   0   'False
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E0E0E0&
      Caption         =   "«·„” Œœ„"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Index           =   13
      Left            =   18360
      RightToLeft     =   -1  'True
      TabIndex        =   37
      Top             =   360
      Width           =   855
   End
End
Attribute VB_Name = "FrmBatchSheet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
 Option Explicit
 Dim RsSavRec As ADODB.Recordset
 Dim StrSQL As String
 Dim RsDevsub As ADODB.Recordset
 Dim BKGrndPic As ClsBackGroundPic
 Dim RecId As String
 Dim Account_Code_dynamic As String
 Dim RevenueAccount As String
 Dim II As Long
 Public LonRow As Double
Public LngCol As Double

Private Sub Accredit_Click()
Dim i As Integer
If Me.TxtModFlg.Text <> "E" And Me.TxtModFlg.Text <> "N" Then
If CheckItem() = False Then
If SystemOptions.UserInterface = ArabicInterface Then
MsgBox "Â–« «·’‰› ·Ì” ’‰› „‰ Ã"
Else
MsgBox "This is item not Product item"
End If
Exit Sub
End If
If val(TxtQtyPermix.Text) = 0 Then
If SystemOptions.UserInterface = ArabicInterface Then
MsgBox "·« ÊÃœ ﬂ„Ì…"
Else
MsgBox "There is no quantity"
End If
Exit Sub
Else
Dim UnitID As Long
GetDefaultItemUnit val(Me.DcbItem.BoundText), UnitID
Dim NoOfworkorder As Double
        If SystemOptions.BatchCreateManyworkOrder = True Then
             NoOfworkorder = val(TxtNoPermix)
        Else
                NoOfworkorder = 1
        
        End If
For i = 1 To NoOfworkorder
SavePalnOrder UnitID
Next i
'If SystemOptions.UserInterface = ArabicInterface Then
'MsgBox " „ «·Õ›Ÿ »‰Ã«Õ"
'Else
'MsgBox "Saved Successfully"
'End If
CheckIsPlanOrder
End If
End If
End Sub

Private Sub BasedRd_Click(Index As Integer)
FromDate_Change
End Sub

Private Sub Cmd_Click(Index As Integer)
If Me.TxtModFlg.Text <> "R" Then
Select Case Index
Case 0
RemoveGridRow2
Case 1
 FG.Clear flexClearScrollable, flexClearEverything
            FG.Rows = 2
Case 2
RemoveGridRow
Case 3
 VSFlexGrid1.Clear flexClearScrollable, flexClearEverything
            VSFlexGrid1.Rows = 2
 End Select
End If
End Sub
Private Sub ReLineGrid()
    Dim IntCounter As Integer
    IntCounter = 0
    Dim i As Integer
    With FG
        For i = .FixedRows To .Rows - 1
            If .TextMatrix(i, .ColIndex("ItemCode")) <> "" Then
                IntCounter = IntCounter + 1
                .TextMatrix(i, .ColIndex("Ser")) = IntCounter
            End If
        Next i
    End With
End Sub
Function GetQty(Optional ItemID As Double) As Double
Dim sql As String
Dim rs2 As ADODB.Recordset
Set rs2 = New ADODB.Recordset
sql = " SELECT    MAX(PartItemQty) AS Qty"
sql = sql & " From dbo.TblItemsParts"
sql = sql & " Where (ItemID = " & ItemID & ")"
rs2.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
If rs2.RecordCount > 0 Then
GetQty = IIf(IsNull(rs2("Qty").value), 0, rs2("Qty").value)
Else
GetQty = 0
End If
End Function


Sub RetriveItems(Optional ItemID As Double)
Dim sql As String
Dim i As Integer
Dim rs2  As ADODB.Recordset
Set rs2 = New ADODB.Recordset
    FG.Clear flexClearScrollable, flexClearEverything
    FG.Rows = 1
sql = " SELECT     TOP 100 PERCENT dbo.TblItemsParts.Unitid, dbo.TblItemsParts.PartItemPrice, dbo.TblItemsParts.PartItemQty, dbo.TblItemsParts.ItemID, dbo.TblItemsParts.TableID, "
sql = sql & "                      dbo.TblUnites.UnitName, dbo.TblUnites.UnitNamee, dbo.TblItemsParts.PartItemID, dbo.TblItems.ItemName, dbo.TblItems.ItemNamee, dbo.TblItems.Fullcode,"
sql = sql & "                      dbo.TblItems.barCodeNO , TblItems_1.MasterNo"
sql = sql & " FROM         dbo.TblItemsParts INNER JOIN"
sql = sql & "                      dbo.TblUnites ON dbo.TblItemsParts.Unitid = dbo.TblUnites.UnitID LEFT OUTER JOIN"
sql = sql & "                      dbo.TblItems TblItems_1 ON dbo.TblItemsParts.ItemID = TblItems_1.ItemID LEFT OUTER JOIN"
sql = sql & "                      dbo.TblItems ON dbo.TblItemsParts.PartItemID = dbo.TblItems.ItemID"
sql = sql & " Where (dbo.TblItemsParts.ItemID = " & ItemID & ")"
rs2.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
If rs2.RecordCount > 0 Then
rs2.MoveFirst
With FG
.Rows = .Rows + rs2.RecordCount
For i = 1 To .Rows - 1
.TextMatrix(i, .ColIndex("Ser")) = i
.TextMatrix(i, .ColIndex("UnitID")) = IIf(IsNull(rs2("Unitid").value), 0, rs2("Unitid").value)
.TextMatrix(i, .ColIndex("ItemID")) = IIf(IsNull(rs2("PartItemID").value), 0, rs2("PartItemID").value)
.TextMatrix(i, .ColIndex("ItemCode")) = IIf(IsNull(rs2("Fullcode").value), 0, rs2("Fullcode").value)
.TextMatrix(i, .ColIndex("Qty")) = IIf(IsNull(rs2("PartItemQty").value), 0, rs2("PartItemQty").value)
.TextMatrix(i, .ColIndex("Cost")) = IIf(IsNull(rs2("PartItemPrice").value), 0, rs2("PartItemPrice").value)
.TextMatrix(i, .ColIndex("TempCost")) = IIf(IsNull(rs2("PartItemPrice").value), 0, rs2("PartItemPrice").value)
.TextMatrix(i, .ColIndex("TempQty")) = IIf(IsNull(rs2("PartItemQty").value), 0, rs2("PartItemQty").value)
If SystemOptions.UserInterface = ArabicInterface Then
.TextMatrix(i, .ColIndex("ItemName")) = IIf(IsNull(rs2("ItemName").value), "", rs2("ItemName").value)
.TextMatrix(i, .ColIndex("UnitName")) = IIf(IsNull(rs2("UnitName").value), "", rs2("UnitName").value)
Else
.TextMatrix(i, .ColIndex("ItemName")) = IIf(IsNull(rs2("ItemNamee").value), "", rs2("ItemNamee").value)
.TextMatrix(i, .ColIndex("UnitName")) = IIf(IsNull(rs2("UnitNamee").value), "", rs2("UnitNamee").value)
End If
TxtMasterNo.Text = IIf(IsNull(rs2("MasterNo").value), "", rs2("MasterNo").value)
rs2.MoveNext
Next i
End With
End If
End Sub

Sub DivNo()
If val(TxtQtyPermix.Text) <> 0 Then
TxtNoPermix.Text = Round(val(TxtPlanQty.Text) / val(TxtQtyPermix.Text), 0)
TxtNoPermix2.Text = val(TxtPlanQty.Text) / val(TxtQtyPermix.Text)
Else
TxtNoPermix.Text = 0
TxtNoPermix2.Text = 0
End If
RelainGrid
End Sub

Private Sub Command1_Click()
If Me.TxtModFlg.Text <> "R" Then
If BasedRd(1).value = True Then
TxtPlanQty.Text = RetriveQtyPlan()
Else
TxtPlanQty.Text = RetriveQtyOrder()
End If
Me.DcbCustomer.BoundText = GetCusIDFromPlan()
DivNo
End If
End Sub

Private Sub DcbItem_Change()
DcbItem_Click (0)
End Sub

Private Sub DcbItem_Click(Area As Integer)
If val(Me.DcbItem.BoundText) <> 0 Then
Me.TxtItemCode.Text = GetItemCode(val(Me.DcbItem.BoundText))
End If
If Me.TxtModFlg.Text <> "R" Then
If val(DcbItem.BoundText) <> 0 Then
RetriveItems val(DcbItem.BoundText)
RetriveEqup val(DcbPermixerMach.BoundText), val(DcbItem.BoundText)
If val(GetQty(val(DcbItem.BoundText))) <> 0 Then
TxtMinQty.Text = Round((100 / GetQty(val(DcbItem.BoundText))) * 25, 4)
End If
End If
End If
RelainGrid
FillGrid
End Sub
Function RetriveQtyOrder() As Double
Dim rs2 As ADODB.Recordset
Dim sql As String
Set rs2 = New ADODB.Recordset
sql = " SELECT     SUM(dbo.Transaction_Details.ShowQty) AS SumCunt"
sql = sql & " FROM         dbo.Transaction_Details RIGHT OUTER JOIN"
sql = sql & "                      dbo.Transactions ON dbo.Transaction_Details.Transaction_ID = dbo.Transactions.Transaction_ID"
sql = sql & " WHERE     (dbo.Transaction_Details.Item_ID = " & val(DcbItem.BoundText) & ") AND (dbo.Transactions.Transaction_Date <= " & SQLDate(ToDate.value, True) & ") and (dbo.Transactions.Transaction_Date >= " & SQLDate(Fromdate.value, True) & ") AND"
sql = sql & "                      (dbo.Transactions.Transaction_Type = 6)"
rs2.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
If rs2.RecordCount > 0 Then
RetriveQtyOrder = IIf(IsNull(rs2("SumCunt").value), 0, rs2("SumCunt").value)
Else
RetriveQtyOrder = 0
End If
End Function
Function RetriveQtyPlan() As Double
Dim rs2 As ADODB.Recordset
Dim sql As String
Set rs2 = New ADODB.Recordset
sql = " SELECT     SUM(dbo.TbllProductionPlanDetails.Price) AS SumCunt"
sql = sql & " FROM         dbo.TbllProductionPlanDetails RIGHT OUTER JOIN"
sql = sql & "                      dbo.TbllProductionPlan ON dbo.TbllProductionPlanDetails.TbllProductionPlanD = dbo.TbllProductionPlan.TbllProductionPlanD"
sql = sql & " WHERE     (dbo.TbllProductionPlan.Todate <= " & SQLDate(ToDate.value, True) & ") AND (dbo.TbllProductionPlan.FromDate >= " & SQLDate(Fromdate.value, True) & ") AND (dbo.TbllProductionPlanDetails.ItemID = " & val(DcbItem.BoundText) & ")"
rs2.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
If rs2.RecordCount > 0 Then
RetriveQtyPlan = IIf(IsNull(rs2("SumCunt").value), 0, rs2("SumCunt").value)
Else
RetriveQtyPlan = 0
End If
End Function
Function GetCusIDFromPlan() As Double
Dim sql As String
Dim rs2 As ADODB.Recordset
Set rs2 = New ADODB.Recordset
sql = "Select CustomerId  "
sql = sql & " FROM         dbo.TbllProductionPlanDetails RIGHT OUTER JOIN"
sql = sql & "                      dbo.TbllProductionPlan ON dbo.TbllProductionPlanDetails.TbllProductionPlanD = dbo.TbllProductionPlan.TbllProductionPlanD"
sql = sql & " WHERE     (dbo.TbllProductionPlan.Todate <= " & SQLDate(ToDate.value, True) & ") AND (dbo.TbllProductionPlan.FromDate >= " & SQLDate(Fromdate.value, True) & ") AND (dbo.TbllProductionPlanDetails.ItemID = " & val(DcbItem.BoundText) & ")"
rs2.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
If rs2.RecordCount > 0 Then
GetCusIDFromPlan = IIf(IsNull(rs2("CustomerId").value), 0, rs2("CustomerId").value)
Else
GetCusIDFromPlan = 0
End If
End Function
Sub RetriveEqup(Optional FixedassetId As Double, Optional ItemID As Double)
Dim sql As String
Dim rs2 As ADODB.Recordset
Set rs2 = New ADODB.Recordset
sql = " SELECT     dbo.TblEquiCapacity.GroupID, SUM(dbo.TblEquiCapacity.Capacity) AS SumCapacity"
sql = sql & " FROM         dbo.TblEquipments LEFT OUTER JOIN"
sql = sql & "                      dbo.TblEquiCapacity ON dbo.TblEquipments.id = dbo.TblEquiCapacity.EquipID"
sql = sql & " Where (dbo.TblEquipments.FixedassetId = " & FixedassetId & ") And (dbo.TblEquiCapacity.ItemID = " & ItemID & ")"
sql = sql & " GROUP BY dbo.TblEquiCapacity.GroupID"
rs2.Open sql, Cn, adOpenKeyset, adLockOptimistic, adCmdText
If rs2.RecordCount > 0 Then
DcbGroup.BoundText = IIf(IsNull(rs2("GroupID").value), 0, rs2("GroupID").value)
Else
DcbGroup.BoundText = 0
End If
End Sub
Sub RetriveCapacity(Optional FixedassetId As Double)
Dim sql As String
Dim rs2 As ADODB.Recordset
Set rs2 = New ADODB.Recordset
sql = " SELECT    Capacites"
sql = sql & " FROM        TblEquipments "
sql = sql & " Where (FixedassetId = " & FixedassetId & ") and ChKLockeq=0 "
rs2.Open sql, Cn, adOpenKeyset, adLockOptimistic, adCmdText
If rs2.RecordCount > 0 Then
TxtQtyPermix.Text = IIf(IsNull(rs2("Capacites").value), 0, rs2("Capacites").value)
Else
TxtQtyPermix.Text = 0
End If
DivNo
End Sub

Private Sub DcbItemFinish_Change()
DcbItemFinish_Click (0)
End Sub

Private Sub DcbItemFinish_Click(Area As Integer)
Me.Text4.Text = GetItemCode(val(Me.DcbItemFinish.BoundText))
End Sub

Private Sub DcbPermixerMach_Change()
DcbPermixerMach_Click (0)
End Sub

Private Sub DcbPermixerMach_Click(Area As Integer)
If Me.TxtModFlg.Text <> "R" Then
If val(DcbPermixerMach.BoundText) <> 0 Then
RetriveEqup val(DcbPermixerMach.BoundText), val(DcbItem.BoundText)
 RetriveCapacity val(DcbPermixerMach.BoundText)
End If
End If
End Sub

Sub FillGrid()
Dim sql As String
Dim Rs1 As ADODB.Recordset
     VSFlexGrid1.Clear flexClearScrollable, flexClearEverything
    VSFlexGrid1.Rows = 1
'Sql = " select * from TblQCItems"
sql = " SELECT     dbo.TblGroupQCItems.Standrd, dbo.TblGroupQCItems.GroupID, dbo.TblGroupQCItems.GQCID, dbo.TblQCItems.name, dbo.TblQCItems.namee ,dbo.TblGroupQCItems.Comment"
sql = sql & " FROM         dbo.TblGroupQCItems LEFT OUTER JOIN"
sql = sql & "  dbo.TblQCItems ON dbo.TblGroupQCItems.GQCID = dbo.TblQCItems.qcid"
sql = sql & " where dbo.TblGroupQCItems.GroupID in(select  GroupID from TblItems where ItemID =" & val(DcbItem.BoundText) & " )"
Set Rs1 = New ADODB.Recordset
  Dim i As Integer
  Rs1.Open sql, Cn, adOpenKeyset, adLockOptimistic, adCmdText
    If Rs1.RecordCount > 0 Then
     Rs1.MoveFirst
     End If
    
     With VSFlexGrid1
                    For i = .FixedRows To Rs1.RecordCount
                   .Rows = .FixedRows + Rs1.RecordCount
                   .TextMatrix(i, .ColIndex("Ser")) = i
                   .TextMatrix(i, .ColIndex("ItemID")) = IIf(IsNull(Rs1("GQCID").value), 0, Rs1("GQCID").value)
                   .TextMatrix(i, .ColIndex("Standrd")) = IIf(IsNull(Rs1("Standrd").value), "", Rs1("Standrd").value)
                   .TextMatrix(i, .ColIndex("Comment")) = IIf(IsNull(Rs1("Comment").value), "", Rs1("Comment").value)
                   If SystemOptions.UserInterface = ArabicInterface Then
                   .TextMatrix(i, .ColIndex("Name")) = IIf(IsNull(Rs1("Name").value), "", Rs1("Name").value)
                   Else
                   .TextMatrix(i, .ColIndex("Name")) = IIf(IsNull(Rs1("NameE").value), "", Rs1("NameE").value)
                   End If
                   Rs1.MoveNext
             Next i
        End With
ErrTrap:
    End Sub

Private Sub FG_Click()
RelainGrid
End Sub
Sub RelainGrid()
Dim i As Integer
Dim SumValue As Double
Dim SumQty As Double
Dim SumCost As Double
SumValue = 0
SumQty = 0
SumCost = 0
With FG
For i = 1 To .Rows - 1
If val(.TextMatrix(i, .ColIndex("ItemID"))) <> 0 Then
If val(.TextMatrix(i, .ColIndex("Total"))) <> 0 Then
SumValue = SumValue + val(.TextMatrix(i, .ColIndex("Total")))
End If
If val(.TextMatrix(i, .ColIndex("TempQty"))) <> 0 Then
If val(TxtPlanQty.Text) <> 0 Then
.TextMatrix(i, .ColIndex("Qty")) = val(.TextMatrix(i, .ColIndex("TempQty"))) * val(TxtPlanQty.Text)
.TextMatrix(i, .ColIndex("Cost")) = val(.TextMatrix(i, .ColIndex("TempCost"))) * val(TxtPlanQty.Text)
End If
If val(TxtNoPermix2.Text) <> 0 Then
.TextMatrix(i, .ColIndex("Qty")) = Round(val(.TextMatrix(i, .ColIndex("Qty"))) / val(TxtNoPermix2.Text), 4)
.TextMatrix(i, .ColIndex("Cost")) = Round(val(.TextMatrix(i, .ColIndex("Cost"))) / val(TxtNoPermix2.Text), 4)
End If
SumCost = SumCost + val(.TextMatrix(i, .ColIndex("Cost")))
SumQty = SumQty + val(.TextMatrix(i, .ColIndex("Qty")))
End If
End If
Next i
End With
lbl(23).Caption = SumQty
lbl(25).Caption = SumValue
lbl(27).Caption = SumCost
End Sub
    Private Sub Form_Load()
    On Error GoTo ErrTrap
    Dim conection As String
    Dim My_SQL As String
           If SystemOptions.UserInterface = ArabicInterface Then
               FG.ColComboList(FG.ColIndex("Stage")) = "#1;Premixer |#2;All|#3; Extruder |#4;Grinder"
               VSFlexGrid1.ColComboList(VSFlexGrid1.ColIndex("Stage")) = "#1;Premixer |#2;All|#3; Extruder |#4;Grinder"
            ElseIf SystemOptions.UserInterface = EnglishInterface Then
               FG.ColComboList(FG.ColIndex("Stage")) = "#1;Premixer |#2;All|#3; Extruder |#4;Grinder"
               VSFlexGrid1.ColComboList(VSFlexGrid1.ColIndex("Stage")) = "#1;Premixer |#2;All|#3; Extruder |#4;Grinder"
            End If
    conection = "select * from TblBatchSheet order by  ID "
    Set BKGrndPic = New ClsBackGroundPic
    Set RsSavRec = New ADODB.Recordset
    RsSavRec.CursorLocation = adUseClient
    RsSavRec.Open conection, Cn, adOpenStatic, adLockOptimistic, adCmdText
    Me.TxtModFlg.Text = "R"
    Resize_Form Me
    Dim Dcombos As New ClsDataCombos
    Dcombos.GetBranches Me.Dcbranch
     Dcombos.GetCustomersSuppliers 1, Me.DcbCustomer
    Dcombos.GetUsers Me.DCboUserName
    Dcombos.GetItemsNamesupdate Me.DcbItem
    Dcombos.GetItemsNamesupdate Me.DcbItemFinish
    Dcombos.GetItemSGroups Me.DcbGroup
    Dcombos.GetEquipments Me.DcbExtruderMach
    Dcombos.GetEquipments Me.DcbBlenderMach
    Dcombos.GetEquipments Me.DcbPermixerMach
    Dcombos.GetEquipments Me.DcbGrinderMach
    
    BtnLast_Click

    ShowTip
    If SystemOptions.UserInterface = EnglishInterface Then
        SetInterface Me
        ChangeLang
        SwitchKeyboardLang LANG_ENGLISH
        Else
        SwitchKeyboardLang LANG_ARABIC
    End If
    If OPEN_NEW_SCREEN = True Then
        btnNew_Click
    End If
   Me.Refresh
ErrTrap:
End Sub
Function Coding() As String
Dim tempcode As String
Dim code As String
Dim no_of_digit  As Integer
Dim diffrent As Integer
Dim i As Integer
  no_of_digit = 4
 tempcode = GetMaxCode
    If Len(tempcode) < no_of_digit Then
            
            diffrent = no_of_digit - Len(tempcode)
            tempcode = ""

                For i = 1 To diffrent
                    tempcode = tempcode & "0"
                    code = code & "0"
                                   
                Next i
            Coding = tempcode & GetMaxCode
    End If
End Function
Function GetMaxID() As Double
Dim sql As String
Dim rs2 As ADODB.Recordset
Set rs2 = New ADODB.Recordset
sql = "select max(ID) as ID from TblBatchSheet where ItemID=" & val(DcbItem.BoundText) & " and id<>" & val(TxtSerial1.Text) & "  "
rs2.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
If rs2.RecordCount > 0 Then
GetMaxID = IIf(IsNull(rs2("ID").value), 0, rs2("ID").value)
Else
GetMaxID = 0
End If
End Function
Function GetBachNo() As String
Dim sql As String
Dim rs2 As ADODB.Recordset
Set rs2 = New ADODB.Recordset
sql = "select BatchNo from TblBatchSheet where  id = " & GetMaxID & "  "
rs2.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
If rs2.RecordCount > 0 Then
GetBachNo = IIf(IsNull(rs2("BatchNo").value), "", rs2("BatchNo").value)
Else
GetBachNo = ""
End If
End Function

Function GetMaxCode() As String
Dim sql As String
Dim rs2 As ADODB.Recordset
Set rs2 = New ADODB.Recordset
sql = "select max(  CAST(code AS float)  ) as Code from TblBatchSheet where YearID=" & val(year(XPDtbTrans.value) - 2000) & "  "
rs2.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
If rs2.RecordCount > 0 Then
GetMaxCode = IIf(IsNull(rs2("Code").value), 0, rs2("Code").value) + 1
Else
GetMaxCode = 1
End If
End Function
Public Sub FiLLRec()

  '  On Error GoTo ErrTrap
    Dim sql As String
    Dim ID As Double
             If Me.TxtModFlg.Text = "E" Then
                 StrSQL = "Delete From TblBatchSheetDet Where BachShetID =" & val(TxtSerial1.Text) & ""
                  Cn.Execute StrSQL, , adExecuteNoRecords
              End If
    RsSavRec.Fields("CusID").value = val(Me.DcbCustomer.BoundText)
    RsSavRec.Fields("YearID").value = (year(XPDtbTrans.value) - 2000)
    RsSavRec.Fields("Code").value = txtCode.Text
    RsSavRec.Fields("TotalCost").value = val(lbl(27).Caption)
    RsSavRec.Fields("NoPermix2").value = val(TxtNoPermix2.Text)
    RsSavRec.Fields("MinQty").value = val(TxtMinQty.Text)
    RsSavRec.Fields("TotalValue").value = val(lbl(25).Caption)
    RsSavRec.Fields("TotalQty").value = val(lbl(23).Caption)
    RsSavRec.Fields("BranchID").value = val(Me.Dcbranch.BoundText)
    RsSavRec.Fields("RecordDate").value = XPDtbTrans.value
    RsSavRec.Fields("BatchNo").value = TxtBatchNo.Text
    RsSavRec.Fields("RefNo").value = TxtRefNo.Text
    RsSavRec.Fields("ItemID").value = val(Me.DcbItem.BoundText)
    RsSavRec.Fields("ItemFInshID").value = val(DcbItemFinish.BoundText)
    RsSavRec.Fields("UserID").value = val(Me.DCboUserName.BoundText)
    RsSavRec.Fields("GroupID").value = val(Me.DcbGroup.BoundText)
    If BasedRd(1).value = True Then
    RsSavRec.Fields("BasedType").value = 1
    Else
    RsSavRec.Fields("BasedType").value = 0
    End If
    RsSavRec.Fields("FromDate").value = Fromdate.value
    RsSavRec.Fields("ToDate").value = ToDate.value
    RsSavRec.Fields("PlanQty").value = val(TxtPlanQty.Text)
    RsSavRec.Fields("QtyPermix").value = val(TxtQtyPermix.Text)
    RsSavRec.Fields("NoPermix").value = val(TxtNoPermix.Text)
    RsSavRec.Fields("PermixerMachID").value = val(DcbPermixerMach.BoundText)
    RsSavRec.Fields("GrinderMachID").value = val(DcbGrinderMach.BoundText)
    RsSavRec.Fields("ExtruderMachID").value = val(DcbExtruderMach.BoundText)
    RsSavRec.Fields("BlenderMachID").value = val(DcbBlenderMach.BoundText)
    RsSavRec.Fields("MasterNo").value = TxtMasterNo.Text
    RsSavRec.Fields("Remarks").value = TxtRemarks.Text
    RsSavRec.Fields("SpecSeries").value = TxtSpecSeries.Text
    RsSavRec.update
  
''//////////////////////////
     Set RsDevsub = New ADODB.Recordset
    StrSQL = "SELECT  *  from TblBatchSheetDet Where (1 = -1)"
    RsDevsub.Open StrSQL, Cn, adOpenKeyset, adLockOptimistic, adCmdText
    Dim i As Integer
    Dim str2 As String
    With FG
       For i = .FixedRows To .Rows - 1
       If val(.TextMatrix(i, .ColIndex("ItemID"))) <> 0 Then
       RsDevsub.AddNew
                 RsDevsub("TypeTrans").value = 0
                 RsDevsub("BachShetID").value = val(Me.TxtSerial1.Text)
                 RsDevsub("ItemID").value = IIf((.TextMatrix(i, .ColIndex("ItemID"))) = "", Null, val(.TextMatrix(i, .ColIndex("ItemID"))))
                 RsDevsub("Stage").value = IIf((val(.TextMatrix(i, .ColIndex("Stage")))) = 0, 1, (.TextMatrix(i, .ColIndex("Stage"))))
                 RsDevsub("UnitID").value = IIf((.TextMatrix(i, .ColIndex("UnitID"))) = "", Null, val(.TextMatrix(i, .ColIndex("UnitID"))))
                 RsDevsub("Qty").value = IIf((.TextMatrix(i, .ColIndex("Qty"))) = "", Null, val(.TextMatrix(i, .ColIndex("Qty"))))
                 RsDevsub("TempQty").value = IIf((.TextMatrix(i, .ColIndex("TempQty"))) = "", Null, val(.TextMatrix(i, .ColIndex("TempQty"))))
                 RsDevsub("C1").value = IIf((.TextMatrix(i, .ColIndex("C1"))) = "", Null, val((.TextMatrix(i, .ColIndex("C1")))))
                 RsDevsub("C2").value = IIf((.TextMatrix(i, .ColIndex("C2"))) = "", Null, val(.TextMatrix(i, .ColIndex("C2"))))
                 RsDevsub("C3").value = IIf((.TextMatrix(i, .ColIndex("C3"))) = "", Null, val(.TextMatrix(i, .ColIndex("C3"))))
                 RsDevsub("TempCost").value = IIf((.TextMatrix(i, .ColIndex("TempCost"))) = "", Null, val((.TextMatrix(i, .ColIndex("TempCost")))))
                 RsDevsub("Cost").value = IIf((.TextMatrix(i, .ColIndex("Cost"))) = "", Null, val((.TextMatrix(i, .ColIndex("Cost")))))
                 RsDevsub("Total").value = IIf((.TextMatrix(i, .ColIndex("Total"))) = "", Null, val(.TextMatrix(i, .ColIndex("Total"))))
       RsDevsub.update
      End If
     Next i
    End With
'''///////////////
      Set RsDevsub = New ADODB.Recordset
    StrSQL = "SELECT  *  from TblBatchSheetDet Where (1 = -1)"
    RsDevsub.Open StrSQL, Cn, adOpenKeyset, adLockOptimistic, adCmdText

    With Me.VSFlexGrid1
       For i = .FixedRows To .Rows - 1
       If val(.TextMatrix(i, .ColIndex("ItemID"))) <> 0 Then
       RsDevsub.AddNew
                RsDevsub("TypeTrans").value = 1
                RsDevsub("BachShetID").value = val(Me.TxtSerial1.Text)
                RsDevsub("Stage").value = IIf((val(.TextMatrix(i, .ColIndex("Stage")))) = 0, 1, (.TextMatrix(i, .ColIndex("Stage"))))
                RsDevsub("ItemID").value = IIf((.TextMatrix(i, .ColIndex("ItemID"))) = "", Null, val(.TextMatrix(i, .ColIndex("ItemID"))))
                RsDevsub("Standrd").value = IIf((.TextMatrix(i, .ColIndex("Standrd"))) = "", Null, (.TextMatrix(i, .ColIndex("Standrd"))))
                RsDevsub("Comment").value = IIf((.TextMatrix(i, .ColIndex("Comment"))) = "", Null, (.TextMatrix(i, .ColIndex("Comment"))))
       RsDevsub.update
      End If
     Next i
    End With
  ''///////////////////
      Select Case Me.TxtModFlg.Text
        Case "N"
            Dim Msg As String
            If SystemOptions.UserInterface = ArabicInterface Then
                Msg = "  „ Õ›Ÿ »Ì«‰«  Â–Â «·⁄„·Ì… " & CHR(13)
                Msg = Msg + "Â·  —€» ›Ì ≈÷«›… »Ì«‰«  √Œ—Ï"
            Else
               Msg = " This record alredy saved... " & CHR(13)
                Msg = Msg + " You want to enter another record?"
           End If
                If MsgBox(Msg, vbYesNo + vbQuestion + vbMsgBoxRight + vbMsgBoxRtlReading + vbDefaultButton2, App.title) = vbYes Then
              
                Me.Refresh
                FiLLTXT
                TxtModFlg = "R"
                 If SystemOptions.UserInterface = ArabicInterface Then
             Else
              
                Me.Refresh
                FiLLTXT
                TxtModFlg = "R"
                MsgBox "Changes Was Saved ... Continuation Add Data ", vbInformation + vbOKOnly + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
            End If
                Call btnNew_Click
            Else
                Me.Refresh
                FiLLTXT
                TxtModFlg = "R"
            End If
         Case "E"
            If SystemOptions.UserInterface = ArabicInterface Then
                MsgBox " „ Õ›Ÿ Â–Â «· ⁄œÌ·« ", vbInformation + vbOKOnly + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
                
                Me.Refresh
                FiLLTXT
                TxtModFlg = "R"
            Else
                MsgBox "Changes was saved", vbInformation + vbOKOnly + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
                
                Me.Refresh
                FiLLTXT
                TxtModFlg = "R"
            End If
       End Select
  Exit Sub
ErrTrap:
    If RsSavRec.EditMode <> adEditNone Then
        RsSavRec.CancelUpdate
    End If
   End Sub


' full data from database
'+++++++++++++++++++++++++++++++++++++++
Public Sub FiLLTXT()
   On Error GoTo ErrTrap
   ' TxtMinQty.Text = IIf(IsNull(RsSavRec.Fields("MinQty").value), 0, RsSavRec.Fields("MinQty").value)
   Me.DcbCustomer.BoundText = IIf(IsNull(RsSavRec.Fields("CusID").value), 0, RsSavRec.Fields("CusID").value)
    TxtMinQty.Text = IIf(IsNull(RsSavRec.Fields("MinQty").value), 0, RsSavRec.Fields("MinQty").value)
    lbl(25).Caption = IIf(IsNull(RsSavRec.Fields("TotalValue").value), 0, RsSavRec.Fields("TotalValue").value)
    lbl(27).Caption = IIf(IsNull(RsSavRec.Fields("TotalCost").value), 0, RsSavRec.Fields("TotalCost").value)
    lbl(23).Caption = IIf(IsNull(RsSavRec.Fields("TotalQty").value), 0, RsSavRec.Fields("TotalQty").value)
     txtCode.Text = IIf(IsNull(RsSavRec.Fields("Code").value), "", RsSavRec.Fields("Code").value)
    TxtSerial1.Text = IIf(IsNull(RsSavRec.Fields("ID").value), "", RsSavRec.Fields("ID").value)
    XPDtbTrans.value = IIf(IsNull(RsSavRec.Fields("RecordDate").value), Date, RsSavRec.Fields("RecordDate").value)
    Dcbranch.BoundText = IIf(IsNull(RsSavRec.Fields("BranchID").value), "", RsSavRec.Fields("BranchID").value)
    TxtBatchNo.Text = IIf(IsNull(RsSavRec.Fields("BatchNo").value), "", RsSavRec.Fields("BatchNo").value)
    TxtRefNo.Text = IIf(IsNull(RsSavRec.Fields("RefNo").value), "", RsSavRec.Fields("RefNo").value)
    Me.DcbItem.BoundText = IIf(IsNull(RsSavRec.Fields("ItemID").value), 0, RsSavRec.Fields("ItemID").value)
    Me.DcbItemFinish.BoundText = IIf(IsNull(RsSavRec.Fields("ItemFInshID").value), 0, RsSavRec.Fields("ItemFInshID").value)   ': ProgressBar1.value = 90
    Me.DcbGroup.BoundText = IIf(IsNull(RsSavRec.Fields("GroupID").value), "", RsSavRec.Fields("GroupID").value)
    Me.DCboUserName.BoundText = IIf(IsNull(RsSavRec.Fields("UserID").value), 0, RsSavRec.Fields("UserID").value)  ': ProgressBar1.value = 10
    If Not IsNull(RsSavRec.Fields("BasedType").value) Then
    If (RsSavRec.Fields("BasedType").value) = 1 Then
    BasedRd(1).value = True
    Else
    BasedRd(0).value = True
    End If
    Else
    BasedRd(0).value = True
    End If
    Fromdate.value = IIf(IsNull(RsSavRec.Fields("FromDate").value), Date, RsSavRec.Fields("FromDate").value)
    ToDate.value = IIf(IsNull(RsSavRec.Fields("ToDate").value), Date, RsSavRec.Fields("ToDate").value)
    TxtPlanQty.Text = IIf(IsNull(RsSavRec.Fields("PlanQty").value), 0, RsSavRec.Fields("PlanQty").value)
    TxtQtyPermix.Text = IIf(IsNull(RsSavRec.Fields("QtyPermix").value), 0, RsSavRec.Fields("QtyPermix").value)
    TxtNoPermix.Text = IIf(IsNull(RsSavRec.Fields("NoPermix").value), "", RsSavRec.Fields("NoPermix").value)
    TxtNoPermix2.Text = IIf(IsNull(RsSavRec.Fields("NoPermix2").value), val(TxtNoPermix.Text), RsSavRec.Fields("NoPermix2").value)
    Me.DcbPermixerMach.BoundText = IIf(IsNull(RsSavRec.Fields("PermixerMachID").value), "", RsSavRec.Fields("PermixerMachID").value)
    Me.DcbGrinderMach.BoundText = IIf(IsNull(RsSavRec.Fields("GrinderMachID").value), "", RsSavRec.Fields("GrinderMachID").value)
    Me.DcbExtruderMach.BoundText = IIf(IsNull(RsSavRec.Fields("ExtruderMachID").value), "", RsSavRec.Fields("ExtruderMachID").value)
    Me.DcbBlenderMach.BoundText = IIf(IsNull(RsSavRec.Fields("BlenderMachID").value), "", RsSavRec.Fields("BlenderMachID").value)
    TxtMasterNo.Text = IIf(IsNull(RsSavRec.Fields("MasterNo").value), "", RsSavRec.Fields("MasterNo").value)
    TxtSpecSeries.Text = IIf(IsNull(RsSavRec.Fields("SpecSeries").value), "", RsSavRec.Fields("SpecSeries").value)
    TxtRemarks.Text = IIf(IsNull(RsSavRec.Fields("Remarks").value), "", RsSavRec.Fields("Remarks").value)
     LabCurrRec.Caption = RsSavRec.AbsolutePosition ': ProgressBar1.value = 50
     LabCountRec.Caption = RsSavRec.RecordCount ': ProgressBar1.value = 60
     CheckIsPlanOrder
FullGridData
ErrTrap:
End Sub

' check before rece
'++++++++++++++++++++++++++++++++++++++++++++
Private Sub btnSave_Click()
   ' On Error GoTo ErrTrap
    Dim Msg As String
    Dim StrVacCode As String
    Dim StrVacName As String
    Dim CtrlTxt As Control

    '---------------------- check if data Vaclete -----------------------
      If Dcbranch.Text = "" And val(Dcbranch.BoundText) = 0 Then
        If SystemOptions.UserInterface = ArabicInterface Then
            MsgBox "⁄›Ê« ...«·—Ã«¡ ≈Œ Ì«— «·›—⁄", vbInformation + vbMsgBoxRtlReading + vbOKOnly + vbMsgBoxRight, App.title
            Else
            MsgBox "Please Select Branch ", vbInformation + vbOKOnly + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
         End If
            Dcbranch.SetFocus
            Exit Sub
     End If
           If DcbItem.Text = "" And val(DcbItem.BoundText) = 0 Then
        If SystemOptions.UserInterface = ArabicInterface Then
            MsgBox "⁄›Ê« ...«·—Ã«¡ ≈Œ Ì«— «·’‰›", vbInformation + vbMsgBoxRtlReading + vbOKOnly + vbMsgBoxRight, App.title
            Else
            MsgBox "Please Select Item ", vbInformation + vbOKOnly + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
         End If
            DcbItem.SetFocus
            Exit Sub
     End If
     
    If TxtBatchNo.Text = "" Then
     TxtBatchNo.Text = (year(XPDtbTrans.value) - 2000) & Coding()
     txtCode.Text = Coding()
    End If
    TxtRefNo.Text = GetBachNo
    If TxtRefNo.Text = "" Then
    TxtRefNo.Text = TxtBatchNo.Text
    End If
    ' -------------------------------------- txtmodflg type -------------------
    Select Case Me.TxtModFlg.Text
            '------------------------------ new record ----------------------------
        Case "N"
                '------------------------- save record -----------------------------
          AddNewRecored
          AddNewRec
           
        '  BtnLast_Click
        Case "E"
            '----------------------------- save edit -------------------------------
            FiLLRec
    End Select
    Exit Sub
ErrTrap:
If SystemOptions.UserInterface = ArabicInterface Then
    MsgBox "Â‰«ﬂ Œÿ√ „« ›Ì ≈œŒ«· «·»Ì«‰« ", vbOKOnly + vbMsgBoxRight, App.title
    Else
    MsgBox "Sorry Error douring insert data", vbOKOnly + vbMsgBoxRight, App.title
    End If
End Sub
' new recored
'++++++++++++++++++++++++++++++++++++
Public Sub AddNewRec()
  'On Error GoTo ErrTrap
    Dim StrRecID As String
    StrRecID = new_id("TblBatchSheet", "ID", "")
    Me.TxtSerial1.Text = StrRecID
    RsSavRec.AddNew
    RsSavRec.Fields("ID").value = IIf(StrRecID <> "", StrRecID, Null)
    FiLLRec
ErrTrap:
End Sub

 Sub FullGridData()
 On Error GoTo ErrTrap
  Dim Rs1 As ADODB.Recordset
  Dim sql As String
    FG.Clear flexClearScrollable, flexClearEverything
    FG.Rows = 1
sql = " SELECT     dbo.TblBatchSheetDet.ID, dbo.TblBatchSheetDet.BachShetID, dbo.TblBatchSheetDet.TypeTrans, dbo.TblBatchSheetDet.Stage, dbo.TblItems.ItemName,"
sql = sql + "                      dbo.TblItems.Fullcode, dbo.TblItems.ItemNamee, dbo.TblBatchSheetDet.ItemID, dbo.TblBatchSheetDet.UnitID, dbo.TblUnites.UnitName, dbo.TblUnites.UnitNamee,"
sql = sql + "                       dbo.TblBatchSheetDet.Qty , dbo.TblBatchSheetDet.c1, dbo.TblBatchSheetDet.c2, dbo.TblBatchSheetDet.C3, dbo.TblBatchSheetDet.total,dbo.TblBatchSheetDet.TempQty ,dbo.TblBatchSheetDet.TempCost,dbo.TblBatchSheetDet.Cost"
sql = sql + "  FROM         dbo.TblBatchSheetDet LEFT OUTER JOIN"
sql = sql + "                       dbo.TblUnites ON dbo.TblBatchSheetDet.UnitID = dbo.TblUnites.UnitID LEFT OUTER JOIN"
sql = sql + "                       dbo.TblItems ON dbo.TblBatchSheetDet.ItemID = dbo.TblItems.ItemID"
sql = sql + "  Where (dbo.TblBatchSheetDet.TypeTrans = 0) And (dbo.TblBatchSheetDet.BachShetID = " & TxtSerial1.Text & ")"
Set Rs1 = New ADODB.Recordset
  Rs1.Open sql, Cn, adOpenKeyset, adLockOptimistic, adCmdText
    If Rs1.RecordCount > 0 Then
     Rs1.MoveFirst
     End If
     Dim i As Integer
     With FG
              For i = .FixedRows To Rs1.RecordCount
                   .Rows = .FixedRows + Rs1.RecordCount
                   .TextMatrix(i, .ColIndex("Ser")) = i
                   .TextMatrix(i, .ColIndex("Stage")) = IIf(IsNull(Rs1("Stage").value), 1, Rs1("Stage").value)
                   .TextMatrix(i, .ColIndex("ItemCode")) = IIf(IsNull(Rs1("Fullcode").value), "", Rs1("Fullcode").value)
                   .TextMatrix(i, .ColIndex("ItemID")) = IIf(IsNull(Rs1("ItemID").value), 0, Rs1("ItemID").value)
                   .TextMatrix(i, .ColIndex("UnitID")) = IIf(IsNull(Rs1("UnitID").value), 0, Rs1("UnitID").value)
                   .TextMatrix(i, .ColIndex("Qty")) = IIf(IsNull(Rs1("Qty").value), 0, Rs1("Qty").value)
                   .TextMatrix(i, .ColIndex("TempQty")) = IIf(IsNull(Rs1("TempQty").value), .TextMatrix(i, .ColIndex("Qty")), Rs1("TempQty").value)
                   .TextMatrix(i, .ColIndex("C1")) = IIf(IsNull(Rs1("C1").value), 0, Rs1("C1").value)
                   .TextMatrix(i, .ColIndex("C2")) = IIf(IsNull(Rs1("C2").value), 0, Rs1("C2").value)
                   .TextMatrix(i, .ColIndex("C3")) = IIf(IsNull(Rs1("C3").value), 0, Rs1("C3").value)
                   .TextMatrix(i, .ColIndex("TempCost")) = IIf(IsNull(Rs1("TempCost").value), 0, Rs1("TempCost").value)
                   .TextMatrix(i, .ColIndex("Cost")) = IIf(IsNull(Rs1("Cost").value), 0, Rs1("Cost").value)
                   .TextMatrix(i, .ColIndex("Total")) = IIf(IsNull(Rs1("Total").value), 0, Rs1("Total").value)
                   If SystemOptions.UserInterface = ArabicInterface Then
                   .TextMatrix(i, .ColIndex("ItemName")) = IIf(IsNull(Rs1("ItemName").value), "", (Rs1("ItemName").value))
                   .TextMatrix(i, .ColIndex("UnitName")) = IIf(IsNull(Rs1("UnitName").value), "", (Rs1("UnitName").value))
                   Else
                   .TextMatrix(i, .ColIndex("ItemName")) = IIf(IsNull(Rs1("ItemNamee").value), "", (Rs1("ItemNamee").value))
                   .TextMatrix(i, .ColIndex("UnitName")) = IIf(IsNull(Rs1("UnitNamee").value), "", (Rs1("UnitNamee").value))
                   End If
                   Rs1.MoveNext
             Next i
        End With
' ''/////////////////////////
     VSFlexGrid1.Clear flexClearScrollable, flexClearEverything
    VSFlexGrid1.Rows = 1
sql = " SELECT     dbo.TblBatchSheetDet.ID, dbo.TblBatchSheetDet.BachShetID, dbo.TblBatchSheetDet.TypeTrans, dbo.TblBatchSheetDet.ItemID, dbo.TblQCItems.name, "
sql = sql + "                      dbo.TblQCItems.NameE , dbo.TblQCItems.comment, dbo.TblBatchSheetDet.Standrd , dbo.TblBatchSheetDet.Stage ,dbo.TblBatchSheetDet.Comment"
sql = sql + " FROM         dbo.TblBatchSheetDet LEFT OUTER JOIN"
sql = sql + "                      dbo.TblQCItems ON dbo.TblBatchSheetDet.ItemID = dbo.TblQCItems.qcid"
sql = sql + " Where (dbo.TblBatchSheetDet.TypeTrans = 1) And (dbo.TblBatchSheetDet.BachShetID = " & TxtSerial1.Text & ")"
Set Rs1 = New ADODB.Recordset
  
  Rs1.Open sql, Cn, adOpenKeyset, adLockOptimistic, adCmdText
    If Rs1.RecordCount > 0 Then
     Rs1.MoveFirst
     End If
    
     With VSFlexGrid1
                    For i = .FixedRows To Rs1.RecordCount
                   .Rows = .FixedRows + Rs1.RecordCount
                   .TextMatrix(i, .ColIndex("Ser")) = i
                   .TextMatrix(i, .ColIndex("Stage")) = IIf(IsNull(Rs1("Stage").value), 1, Rs1("Stage").value)
                   .TextMatrix(i, .ColIndex("ItemID")) = IIf(IsNull(Rs1("ItemID").value), 0, Rs1("ItemID").value)
                   .TextMatrix(i, .ColIndex("Standrd")) = IIf(IsNull(Rs1("Standrd").value), "", Rs1("Standrd").value)
                   .TextMatrix(i, .ColIndex("Comment")) = IIf(IsNull(Rs1("Comment").value), "", Rs1("Comment").value)
                   If SystemOptions.UserInterface = ArabicInterface Then
                   .TextMatrix(i, .ColIndex("Name")) = IIf(IsNull(Rs1("Name").value), "", Rs1("Name").value)
                   Else
                   .TextMatrix(i, .ColIndex("Name")) = IIf(IsNull(Rs1("NameE").value), "", Rs1("NameE").value)
                   End If
                   Rs1.MoveNext
             Next i
        End With
        Exit Sub
ErrTrap:
    End Sub
    Private Sub RemoveGridRow2()
    With Me.FG
        If .Row <= 0 Then Exit Sub
        .RemoveItem .Row
    End With
    ReLineGrid
End Sub
Private Sub RemoveGridRow()
    With Me.VSFlexGrid1
        If .Row <= 0 Then Exit Sub
        .RemoveItem .Row
    End With
    ReLineGrid
End Sub


Private Sub FromDate_Change()
If Me.TxtModFlg.Text <> "R" Then
If BasedRd(1).value = True Then
TxtPlanQty.Text = RetriveQtyPlan()
Else
TxtPlanQty.Text = RetriveQtyOrder()
End If
DivNo
Me.DcbCustomer.BoundText = GetCusIDFromPlan()
End If
End Sub

Private Sub ISButton2_Click()
print_report2
End Sub

Private Sub ISButton5_Click()
print_report
End Sub

Private Sub Text4_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        If Text4.Text = "" Then
            Me.DcbItemFinish.BoundText = ""
        Else
            Me.DcbItemFinish.BoundText = GetItemID(Trim$(Me.Text4.Text))
        End If
    End If
End Sub

Private Sub ToDate_Change()
FromDate_Change
End Sub

Private Sub TxtItemCode_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        If TxtItemCode.Text = "" Then
            Me.DcbItem.BoundText = ""
        Else
            Me.DcbItem.BoundText = GetItemID(Trim$(Me.TxtItemCode.Text))
        End If
    End If
End Sub

Private Sub TxtNoPermix_Change()
If Me.TxtModFlg.Text <> "R" Then
DivNo
End If
End Sub

Private Sub TxtPlanQty_Change()
If Me.TxtModFlg.Text <> "R" Then
DivNo
End If
End Sub

Private Sub TxtQtyPermix_Change()
If Me.TxtModFlg.Text <> "R" Then
DivNo
End If
End Sub

' change id search
Private Sub TxtSerial1_Change()
    Dim TxtMod As String
    TxtMod = TxtModFlg.Text
    TxtModFlg.Text = ""
    TxtModFlg = TxtMod
End Sub
' search for select id
Public Function FindRec(ByVal RecId As Long)
    On Error GoTo ErrTrap
    RsSavRec.find "ID=" & RecId, , adSearchForward, 1
    If Not (RsSavRec.EOF) Then
        FiLLTXT
        End If
    Exit Function
ErrTrap:
    If RsSavRec.EditMode <> adEditNone Then
        RsSavRec.CancelUpdate
        BtnUndo_Click
    End If
  End Function
  ' cancel camnd sub
  '+++++++++++++++++++++++++++++++
  Private Sub BtnCancel_Click()
    Unload Me
End Sub
' undo sub
 Private Sub BtnUndo_Click()
    FindRec val(TxtSerial1.Text)
    Me.TxtModFlg.Text = "R"
    FiLLTXT
     BtnLast_Click
End Sub
' delet sub
Private Sub btnDelete_Click()
    Dim MSGType As Integer
    Dim BlnRecExist As Boolean
    Dim StrMSG  As String
    Dim Msg As String
    On Error GoTo ErrTrap
    If DoPremis(Do_Delete, Me.Name, True) = False Then
        Exit Sub
    End If
    Dim X As Integer
    Dim i As Integer
    Dim ID As Double

    If SystemOptions.UserInterface = EnglishInterface Then
        X = MsgBox("Confirm Delete This line", vbCritical + vbYesNo)
    Else
        X = MsgBox(" √ﬂÌœ «·Õ–›", vbCritical + vbYesNo)
    End If
    If X = vbNo Then Exit Sub
     If TxtSerial1.Text = "" Then
       If SystemOptions.UserInterface = EnglishInterface Then
                X = MsgBox(" Nothing To Delet ", vbInformation + vbMsgBoxRtlReading + vbOKOnly + vbMsgBoxRight, App.title)
               Else
                X = MsgBox("⁄›Ê« ...·« ÌÊÃœ »Ì«‰«  ··Õ–›", vbInformation + vbMsgBoxRtlReading + vbOKOnly + vbMsgBoxRight, App.title)
       End If
               Else
      Dim StrSQL As String
      StrSQL = "delete from   Transactions where BatchID =" & val(TxtSerial1.Text) & ""
      Cn.Execute StrSQL
                RsSavRec.find "ID=" & val(TxtSerial1.Text), , adSearchForward, 1
                  StrSQL = "Delete From TblBatchSheetDet Where BachShetID =" & val(TxtSerial1.Text) & ""
                  Cn.Execute StrSQL, , adExecuteNoRecords
                                          RsSavRec.delete
            FG.Clear flexClearScrollable, flexClearEverything
            FG.Rows = 1
          VSFlexGrid1.Clear flexClearScrollable, flexClearEverything
      VSFlexGrid1.Rows = 1
                 If SystemOptions.UserInterface = EnglishInterface Then
                X = MsgBox(" Delete  Successfully ", vbInformation + vbMsgBoxRtlReading + vbOKOnly + vbMsgBoxRight, App.title)
               Else
                X = MsgBox(" „  ⁄„·Ì… «·Õ–› »‰Ã«Õ", vbInformation + vbMsgBoxRtlReading + vbOKOnly + vbMsgBoxRight, App.title)
               End If

     End If                       '------------------------------ Move Next ---------------------------.
        Me.Refresh
     LabCurrRec.Caption = 0
     LabCountRec.Caption = 0
       ' FillGridWithData
        BtnNext_Click
     Exit Sub
ErrTrap:
     Select Case Err.Number
        Case -2147217873, -2147467259
        If SystemOptions.UserInterface = ArabicInterface Then
            StrMSG = "⁄›Ê« ·« ÌÃÊ“ Õ–› «·”Ã· ·«— »«ÿÂ »»Ì«‰«  √Œ—Ì"
            Else
            StrMSG = "You can not delete the record"
            StrMSG = StrMSG & " Is related to with other data"
            End If
            RsSavRec.CancelUpdate
            MsgBox StrMSG, vbOKOnly + vbMsgBoxRight, App.title
           'Cn.Errors.Clear
    End Select

End Sub
' exit without save sub
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
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
               btnSave_Click
        Case vbCancel
              Cancel = True
        End Select
    End If
    Exit Sub
ErrTrap:
End Sub
Private Sub Form_Terminate()
     ' Set FrmVacancy = Nothing
End Sub
Private Sub Form_Unload(Cancel As Integer)
    On Error GoTo ErrTrap
    If RsSavRec.State = adStateOpen Then
        If Not (RsSavRec.EOF Or RsSavRec.BOF) Then
            If RsSavRec.EditMode <> adEditNone Then
                RsSavRec.CancelUpdate
            End If
        End If
        RsSavRec.Close
        Set RsSavRec = Nothing
    End If
ErrTrap:
End Sub
Private Sub Form_Activate()
    Me.ZOrder 0
End Sub
Public Sub EditRec(StrTable As String, _
                   RecId As String)
     FiLLRec
End Sub
Private Sub TxtModFlg_Change()
    If TxtModFlg.Text = "N" Then
    XPDtbTrans.Enabled = True
        
        Me.btnNew.Enabled = False
        btnModify.Enabled = False
        btnDelete.Enabled = False
        Me.btnQuery.Enabled = False
        ISButton1.Enabled = False
     '   Grid.Enabled = False
        BtnUndo.Enabled = True
        Me.btnSave.Enabled = True
        BtnUpdate.Enabled = False
        
    ElseIf TxtModFlg.Text = "R" Then
     XPDtbTrans.Enabled = False
        btnModify.Enabled = False
        btnDelete.Enabled = False
        If TxtSerial1.Text <> "" Then
            btnModify.Enabled = True
            btnDelete.Enabled = True
    End If
        BtnUpdate.Enabled = True
        Me.btnQuery.Enabled = True
        Me.btnNew.Enabled = True
        BtnUndo.Enabled = False
        Me.btnSave.Enabled = False
        ISButton1.Enabled = True
        btnNext.Enabled = True
        btnPrevious.Enabled = True
        btnFirst.Enabled = True
        btnLast.Enabled = True
   ElseIf TxtModFlg.Text = "E" Then
    XPDtbTrans.Enabled = True
        Me.btnNew.Enabled = False
        btnModify.Enabled = False
        btnDelete.Enabled = False
        Me.btnQuery.Enabled = False
        BtnUpdate.Enabled = False
        BtnUndo.Enabled = True
        Me.btnSave.Enabled = True
    '    Grid.Enabled = False
        btnNext.Enabled = False
        btnPrevious.Enabled = False
        btnFirst.Enabled = False
        btnLast.Enabled = False
    End If
End Sub

' move btowen recored
Private Sub BtnFirst_Click()
    On Error GoTo ErrTrap
    Dim Msg As String
    If Me.TxtModFlg.Text = "N" Then
        FindRec val(TxtSerial1.Text)
        Me.TxtModFlg.Text = "R"
    End If
    TxtModFlg = "R"
    If RsSavRec.RecordCount = 0 Then
        clear_all Me
        
        Exit Sub
    End If
BegnieWork:
    RsSavRec.MoveFirst
    
    FiLLTXT
    Exit Sub
ErrTrap:
    Select Case Err.Number
        Case -2147217885
        If SystemOptions.UserInterface = ArabicInterface Then
            Msg = "⁄›Ê« ·ﬁœ  „ Õ–› Â–« «·”Ã· «· «·Ï" & CHR(13)
            Msg = Msg & "„‰ ﬁ»· „” Œœ„ √Œ— ⁄·Ï «·‘»ﬂ… " & CHR(13)
            Msg = Msg & "”Ê› Ì „  ÕœÌÀ «·»Ì«‰« "
            Else
            Msg = "Sorry I've been to delete this record" & CHR(13)
            Msg = Msg & "By another user on the network " & CHR(13)
            Msg = Msg & "Data will be updated"
            End If
            MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading + vbOKOnly, App.title
            RsSavRec.Requery
            Resume BegnieWork
    End Select
End Sub
Private Sub BtnLast_Click()
    On Error GoTo ErrTrap
    Dim Msg As String
    If Me.TxtModFlg.Text = "N" Then
        FindRec val(TxtSerial1.Text)
        Me.TxtModFlg.Text = "R"
    End If
    TxtModFlg = "R"
    If RsSavRec.RecordCount = 0 Then
        clear_all Me
        
        Exit Sub
    End If
BegnieWork:
    RsSavRec.MoveLast
 
    FiLLTXT
    Exit Sub
ErrTrap:
    Select Case Err.Number
        Case -2147217885
        If SystemOptions.UserInterface = ArabicInterface Then
            Msg = "⁄›Ê« ·ﬁœ  „ Õ–› Â–« «·”Ã· «· «·Ï" & CHR(13)
            Msg = Msg & "„‰ ﬁ»· „” Œœ„ √Œ— ⁄·Ï «·‘»ﬂ… " & CHR(13)
            Msg = Msg & "”Ê› Ì „  ÕœÌÀ «·»Ì«‰« "
            Else
               Msg = "Sorry I've been to delete this record" & CHR(13)
            Msg = Msg & "By another user on the network " & CHR(13)
            Msg = Msg & "Data will be updated"
            End If
            MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading + vbOKOnly, App.title
            RsSavRec.Requery
            Resume BegnieWork
    End Select
End Sub
Private Sub btnModify_Click()
    Dim Msg As String
    If DoPremis(Do_Edit, Me.Name, True) = False Then
        Exit Sub
    End If
    On Error GoTo ErrTrap
    If TxtSerial1.Text <> "" Then
        TxtModFlg = "E"
            FG.Rows = FG.Rows + 1
             VSFlexGrid1.Rows = VSFlexGrid1.Rows + 1
        Me.DCboUserName.BoundText = user_id
        Me.Dcbranch.SetFocus
    End If
    Exit Sub
ErrTrap:
    Select Case Err.Number
        Case -2147467259
            'Could not update; currently locked.
            If SystemOptions.UserInterface = ArabicInterface Then
            Msg = "⁄›Ê«" & CHR(13)
            Msg = Msg & " ·«Ì„ﬂ‰  ⁄œÌ· Â–« «·”Ã· ›Ï «·Êﬁ  «·Õ«·Ï" & CHR(13)
            Msg = Msg & "ÕÌÀ «‰Â ﬁÌœ «· ⁄œÌ· „‰ ﬁ»· „” Œœ„ «Œ— ⁄·Ï «·‘»ﬂ…"
            Else
            Msg = "Sorry.." & CHR(13)
            Msg = Msg & " You can not edit this the record now" & CHR(13)
            Msg = Msg & "It was being edited by another user on the network"
           
            End If
            MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading + vbOKOnly, App.title
                    If RsSavRec.EditMode <> adEditNone Then
                RsSavRec.CancelUpdate
                'RsSavRec.Requery
            End If
    End Select
End Sub
Private Sub btnNew_Click()
    Dim My_SQL As String
    Dim rs As ADODB.Recordset
    If DoPremis(Do_New, Me.Name, True) = False Then
        Exit Sub
    End If
    On Error GoTo ErrTrap
    Set rs = New ADODB.Recordset
    
    clear_all Me

    TxtModFlg.Text = "N"
    FG.Clear flexClearScrollable, flexClearEverything
    FG.Rows = 2
    VSFlexGrid1.Clear flexClearScrollable, flexClearEverything
    VSFlexGrid1.Rows = 2
    Me.DCboUserName.BoundText = user_id
    Me.Dcbranch.BoundText = Current_branch
    Dcbranch.SetFocus
    XPDtbTrans.value = Date
    Fromdate.value = Date
    ToDate.value = Date
    BasedRd(0).value = True
    BasedRd_Click (0)
    FillGrid
ErrTrap:
End Sub
Private Sub BtnNext_Click()
    On Error GoTo ErrTrap
    Dim Msg As String
    If Me.TxtModFlg.Text = "N" Then
        FindRec val(TxtSerial1.Text)
        Me.TxtModFlg.Text = "R"
    End If
    TxtModFlg = "R"
    If RsSavRec.RecordCount = 0 Then
        clear_all Me
      
        Exit Sub
    End If
BegnieWork:
     If RsSavRec.EOF Then
        RsSavRec.MoveLast
    Else
        RsSavRec.MoveNext
        If RsSavRec.EOF Then
            RsSavRec.MoveLast
        End If
    End If
    
    FiLLTXT
    Exit Sub
ErrTrap:
    Select Case Err.Number
        Case -2147217885
        If SystemOptions.UserInterface = ArabicInterface Then
            Msg = "⁄›Ê« ·ﬁœ  „ Õ–› Â–« «·”Ã· «· «·Ï" & CHR(13)
            Msg = Msg & "„‰ ﬁ»· „” Œœ„ √Œ— ⁄·Ï «·‘»ﬂ… " & CHR(13)
            Msg = Msg & "”Ê› Ì „  ÕœÌÀ «·»Ì«‰« "
            Else
               Msg = "Sorry I've been to delete this record" & CHR(13)
            Msg = Msg & "By another user on the network " & CHR(13)
            Msg = Msg & "Data will be updated"
            End If
            MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading + vbOKOnly, App.title
            RsSavRec.Requery
            Resume BegnieWork
    End Select
End Sub
Private Sub BtnPrevious_Click()
    On Error GoTo ErrTrap
    Dim Msg As String
    If Me.TxtModFlg.Text = "N" Then
        FindRec val(TxtSerial1.Text)
        Me.TxtModFlg.Text = "R"
    End If
    TxtModFlg = "R"
    If RsSavRec.RecordCount = 0 Then
        clear_all Me
       
        Exit Sub
    End If
BegnieWork:
    RsSavRec.MovePrevious
    If RsSavRec.BOF Then
        RsSavRec.MoveFirst
    End If
    
    FiLLTXT
    Exit Sub
ErrTrap:
    Select Case Err.Number
        Case -2147217885
        If SystemOptions.UserInterface = ArabicInterface Then
            Msg = "⁄›Ê« ·ﬁœ  „ Õ–› Â–« «·”Ã· «· «·Ï" & CHR(13)
            Msg = Msg & "„‰ ﬁ»· „” Œœ„ √Œ— ⁄·Ï «·‘»ﬂ… " & CHR(13)
            Msg = Msg & "”Ê› Ì „  ÕœÌÀ «·»Ì«‰« "
            Else
            Msg = "Sorry I've been to delete this record" & CHR(13)
            Msg = Msg & "By another user on the network " & CHR(13)
            Msg = Msg & "Data will be updated"
            End If
            MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading + vbOKOnly, App.title
            RsSavRec.Requery
            Resume BegnieWork
    End Select
End Sub


'Information for camand
'++++++++++++++++++++++++++++++++++++++
Private Sub ShowTip()
    On Error GoTo ErrTrap
    Dim TTP As New clstooltip
    Dim Wrap As String
    Dim Msg As String
    Wrap = CHR(13) + CHR(10)
    With TTP
        .Create Me.hWnd, Me.Caption, 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 10000
        .DelayTime = 300
        Msg = "ÃœÌœ" & Wrap & "·› Õ ”Ã· ÃœÌœ " & Wrap & "≈÷€ÿ Â–« «·„› «Õ" & Wrap & "√Ê „› «Õ" & " F12 √Ê Enter"
             .AddControl btnNew, Msg, True
    End With
    With TTP
        .Create Me.hWnd, Me.Caption, 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 10000
        .DelayTime = 300
        Msg = " ⁄œÌ·" & Wrap & "· ⁄œÌ·  ”Ã· «·Õ«·Ï " & Wrap & "≈÷€ÿ Â–« «·„› «Õ" & Wrap & "√Ê „› «Õ" & " F11"
        .AddControl btnModify, Msg, True
    End With
    With TTP
        .Create Me.hWnd, Me.Caption, 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 10000
        .DelayTime = 300
        Msg = "Õ›Ÿ" & Wrap & "· ”ÃÌ· «·»Ì«‰«  œ«Œ· ﬁ«⁄œ… " & Wrap & "«·»Ì«‰«  ≈÷€ÿ Â–« «·„› «Õ" & Wrap & "√Ê „› «Õ" & " F10"
        .AddControl btnSave, Msg, True
    End With
    With TTP
        .Create Me.hWnd, Me.Caption, 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 10000
        .DelayTime = 300
        Msg = " —«Ã⁄" & Wrap & "·· —«Ã⁄ ⁄‰ «·⁄„·Ì… «·Õ«·Ì…" & Wrap & "≈÷€ÿ Â–« «·„› «Õ" & Wrap & "√Ê „› «Õ" & " F9"
        .AddControl BtnUndo, Msg, True
    End With
    With TTP
        .Create Me.hWnd, Me.Caption, 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 10000
        .DelayTime = 300
        Msg = "Õ–› «·”Ã·" & Wrap & "·Õ–› «·”Ã· «·Õ«·Ï" & Wrap & "≈÷€ÿ Â–« «·„› «Õ" & Wrap & "√Ê „› «Õ" & " F18"
        .AddControl btnDelete, Msg, True
    End With
    With TTP
        .Create Me.hWnd, Me.Caption, 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 10000
        .DelayTime = 300
        Msg = "Œ—ÊÃ" & Wrap & "·≈€·«ﬁ Â–Â «·‰«›–…" & Wrap & "≈÷€ÿ Â–« «·„› «Õ" & Wrap & "√Ê „› «Õ" & " Ctrl+x"
        .AddControl btnCancel, Msg, True
    End With
    With TTP
        .Create Me.hWnd, Me.Caption, 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 10000
        .DelayTime = 300
        Msg = "«·«Ê·" & Wrap & "··«‰ ﬁ«· «·Ï «·”Ã· «·«Ê·" & Wrap & "≈÷€ÿ Â–« «·„› «Õ" & Wrap & "√Ê „› «Õ" & " Home √Ê UpArrow"
        .AddControl btnFirst, Msg, True
    End With
    With TTP
        .Create Me.hWnd, Me.Caption, 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 10000
        .DelayTime = 300
        Msg = "«·”«»ﬁ" & Wrap & "··«‰ ﬁ«· «·Ï «·”Ã· «·”«»ﬁ" & Wrap & "≈÷€ÿ Â–« «·„› «Õ" & Wrap & "√Ê „› «Õ" & " PageUp √Ê LeftArrow"
        .AddControl btnPrevious, Msg, True
    End With
    With TTP
        .Create Me.hWnd, Me.Caption, 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 10000
        .DelayTime = 300
        Msg = "«· «·Ï" & Wrap & "··«‰ ﬁ«· «·Ï «·”Ã· «· «·Ï" & Wrap & "≈÷€ÿ Â–« «·„› «Õ" & Wrap & "√Ê „› «Õ" & " PageDown √Ê RightArrow"
        .AddControl btnNext, Msg, True
    End With
    With TTP
        .Create Me.hWnd, Me.Caption, 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 10000
        .DelayTime = 300
        Msg = "«·«ŒÌ—" & Wrap & "··«‰ ﬁ«· «·Ï «·”Ã· «·«ŒÌ—" & Wrap & "≈÷€ÿ Â–« «·„› «Õ" & Wrap & "√Ê „› «Õ" & " End √Ê DownArrow"
        .AddControl btnLast, Msg, True
    End With
ErrTrap:
End Sub
' short cut for keys
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    On Error GoTo ErrTrap
    If KeyCode = vbKeyReturn Then
        If Me.TxtModFlg.Text = "R" Then
            btnNew_Click
        Else
            SendKeys "{TAB}"
        End If
    End If
    'New ---------------------------
    If KeyCode = vbKeyF12 Then
        If btnNew.Enabled = False Then Exit Sub
        btnNew_Click
    End If
    'Edit ------------------------
    If KeyCode = vbKeyF11 Then
        If btnModify.Enabled = False Then Exit Sub
        btnModify_Click
    End If
    'save --------------------------------------------------------------------------------
    If KeyCode = vbKeyF10 Then
        If btnSave.Enabled = False Then Exit Sub
        btnSave_Click
    End If
    'undo ------------------------------------------------------------------------------
    If KeyCode = vbKeyF9 Then
        If BtnUndo.Enabled = False Then Exit Sub
        BtnUndo_Click
    End If
    'Delete ---------------------------------------------------------------------------
    If KeyCode = vbKeyF8 Then
        If btnDelete.Enabled = False Then Exit Sub
        btnDelete_Click
    End If
    'Exit ----------------------------------------------------------------------
    If Shift = 2 Then
        If KeyCode = vbKeyX Then
            If btnCancel.Enabled = False Then Exit Sub
            BtnCancel_Click
        End If
    End If
    'Moveing through Records ---------------------------------------------------------------------------
    'If TxtModFlg.Text = "R" Then
    'Move first --------------------------------------------
    If KeyCode = vbKeyUp Or KeyCode = vbKeyHome Then
        If btnFirst.Enabled = False Then Exit Sub
        BtnFirst_Click
    End If
    'Move Previous---------------------------------------------------------
    If KeyCode = vbKeyLeft Or KeyCode = vbKeyPageUp Then
        If btnPrevious.Enabled = False Then Exit Sub
        BtnPrevious_Click
    End If
    'Move Next---------------------------------------------------------
    If KeyCode = vbKeyRight Or KeyCode = vbKeyPageDown Then
        If btnNext.Enabled = False Then Exit Sub
        BtnNext_Click
    End If
    'Move Last---------------------------------------------------------
    If KeyCode = vbKeyDown Or KeyCode = vbKeyEnd Then
        If btnLast.Enabled = False Then Exit Sub
        BtnLast_Click
    End If
    'End If
    Exit Sub
ErrTrap:
End Sub
Function CheckItem() As Boolean
Dim sql As String
Dim Rs3 As ADODB.Recordset
Set Rs3 = New ADODB.Recordset
sql = " SELECT     ItemID, ItemMakingNew"
sql = sql & " From dbo.TblItems"
sql = sql & " Where (ItemID =" & val(DcbItem.BoundText) & ") And (ItemMakingNew = 1)"
Rs3.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
If Rs3.RecordCount > 0 Then
CheckItem = True
Else
CheckItem = False
End If
End Function
Sub CheckIsPlanOrder()
Dim sql As String
Dim rs2 As ADODB.Recordset
Set rs2 = New ADODB.Recordset
sql = "Select * from  Transactions where BatchID =" & val(TxtSerial1.Text) & ""
rs2.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
If rs2.RecordCount > 0 Then
 Accredit.Enabled = False
Else
Accredit.Enabled = True
End If
End Sub
Sub SavePalnOrder(Optional UnitID As Long)
Dim sql As String
Dim rs2 As ADODB.Recordset
Dim Rs3 As ADODB.Recordset
Dim Transaction_ID As Double
 Dim Sanad_No As Integer
 Dim TransSerial As String
Sanad_No = 49

        my_branch = val(Dcbranch.BoundText)

        If TransSerial = "" Then
            If Voucher_coding(val(my_branch), XPDtbTrans.value, Sanad_No, 0, , 26) = "error" Then
                MsgBox " ·« Ì„ﬂ‰ «÷«›…   Â–« «·”‰œ ·«‰ﬂ  ⁄œÌ  «·Õœ «·„”„ÊÕ »… „‰ «·”‰œ«   ": Exit Sub
            Else
                       
                If Voucher_coding(val(my_branch), XPDtbTrans.value, Sanad_No, 0, , 26) = "" Then
                    TransSerial = CStr(new_id("Transactions", "Transaction_Serial", "", True, "Transaction_Type=26"))
                Else
                    TransSerial = Voucher_coding(val(my_branch), XPDtbTrans.value, Sanad_No, 0, , 26)
                End If
            End If
        End If
        
sql = "select * from Transactions where 1=-1"
Transaction_ID = CStr(new_id("Transactions", "Transaction_ID", "", True))
Set rs2 = New ADODB.Recordset
rs2.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
rs2.AddNew
rs2("Transaction_ID").value = Transaction_ID
rs2("Transaction_Date").value = XPDtbTrans.value
rs2("Transaction_Type").value = 26
rs2("BranchID").value = val(Dcbranch.BoundText)
rs2("BatchID").value = val(TxtSerial1.Text)
rs2("UserID").value = user_id
rs2("Transaction_Serial").value = TransSerial
rs2("NoteSerial1").value = TransSerial
rs2("BatchNo").value = TxtBatchNo.Text
If val(Me.DcbCustomer.BoundText) <> 0 Then
rs2("CusID").value = val(Me.DcbCustomer.BoundText)
Else
rs2("CusID").value = 1
End If
rs2.update
Set Rs3 = New ADODB.Recordset
sql = "select * from Transaction_Details where 1=-1"
Rs3.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
Rs3.AddNew
Rs3("Transaction_ID").value = Transaction_ID
Rs3("Item_ID").value = val(DcbItem.BoundText)

        If SystemOptions.BatchCreateManyworkOrder = True Then
                        Rs3("ShowQty").value = val(TxtQtyPermix.Text)
                        Rs3("Quantity").value = val(TxtQtyPermix.Text)
        Else
                    Rs3("ShowQty").value = val(TxtQtyPermix.Text) * val(TxtNoPermix.Text)
                    Rs3("Quantity").value = val(TxtQtyPermix.Text) * val(TxtNoPermix.Text)
        End If

Rs3("showPrice").value = val(lbl(27).Caption) / val(TxtQtyPermix.Text)

Rs3("Price").value = val(lbl(27).Caption) / val(TxtQtyPermix.Text)
Rs3("UnitId").value = UnitID

Rs3.update

End Sub
Private Sub ChangeLang()
On Error GoTo ErrTrap
       Dim XPic As IPictureDisp
    Set XPic = Me.btnFirst.ButtonImage
    Set Me.btnFirst.ButtonImage = Me.btnLast.ButtonImage
    Set Me.btnLast.ButtonImage = XPic
    Set XPic = Me.btnPrevious.ButtonImage
    Set Me.btnPrevious.ButtonImage = Me.btnNext.ButtonImage
    Set Me.btnNext.ButtonImage = XPic
   ''''''''''''''''''''////
   lbl(26).Caption = "Mini.Qty"
   lbl(9).Caption = "Total Qty"
   lbl(24).Caption = "Total"
   Accredit.Caption = "Conversion To order Prod."
       Me.Caption = "Batch Sheet"
      Label1(2).Caption = Me.Caption
      Me.lbl(4).Caption = "ID"
      Me.lbl(2).Caption = "Date"
      lbl(7).Caption = "Branch"
      lbl(5).Caption = "Batch No"
      lbl(6).Caption = "Ref. No. "
   lbl(0).Caption = "Product Name"
   lbl(1).Caption = "Product Type"
   lbl(3).Caption = "Finish"
lbl(21).Caption = "Remarks"
Frame10.Caption = "Based On"
BasedRd(0).RightToLeft = False
BasedRd(1).RightToLeft = False
BasedRd(0).Caption = "Work Orders"
BasedRd(1).Caption = "Production Plan"
lbl(15).Caption = "From"
lbl(13).Caption = "To"

lbl(10).Caption = "Plan Qty"
lbl(12).Caption = "Qty Per Permix"
lbl(11).Caption = "No Of Permix"
lbl(16).Caption = "Permixer Machine"
lbl(14).Caption = "Grinder Machine"
lbl(18).Caption = "Extruder Machine"
lbl(17).Caption = "Blender Machine"

lbl(19).Caption = "Master No"
lbl(20).Caption = "Spec Series"
Cmd(0).Caption = "Delete"
Cmd(1).Caption = "Delete All"
Cmd(2).Caption = "Delete"
Cmd(3).Caption = "Delete All"
lbl(22).Caption = "Data of Specification"
ISButton2.Caption = "Print Spec."


    ISButton5.Caption = "Print"
    ISButton8.Caption = "Search"
    '''''''''''''' next
    ''''''''''''''''''''''''''''''''''''''' next
    Me.Label2(0).Caption = "Current Record"
    Me.Label2(1).Caption = "NO. Recordes"
    Me.lbl(8).Caption = "by"
    '''''''''''''''''''''''''''''''' next
    btnNew.Caption = "New"
    btnModify.Caption = "Modify"
    btnSave.Caption = "Save"
    BtnUndo.Caption = "Undo"
    BtnUpdate.Caption = "Refresh "
    ISButton1.Caption = "Print"
    btnQuery.Caption = "Search"
    btnDelete.Caption = "Delete"
    btnCancel.Caption = "Exit"
   lbl(28).Caption = "Total Cost"
  With Me.FG
  .TextMatrix(0, .ColIndex("Ser")) = "Serial"
  .TextMatrix(0, .ColIndex("ItemCode")) = "Item Code"
  .TextMatrix(0, .ColIndex("ItemName")) = "Item Name"
  .TextMatrix(0, .ColIndex("Stage")) = "Stage"
  .TextMatrix(0, .ColIndex("UnitName")) = "UOM"
  .TextMatrix(0, .ColIndex("C1")) = "C1"
  .TextMatrix(0, .ColIndex("C2")) = "C2"
  .TextMatrix(0, .ColIndex("C3")) = "C3"
  .TextMatrix(0, .ColIndex("Total")) = "Total"
  .TextMatrix(0, .ColIndex("Qty")) = "Qty"
  .TextMatrix(0, .ColIndex("Cost")) = "Cost"
  
  End With
    With Me.VSFlexGrid1
  .TextMatrix(0, .ColIndex("Ser")) = "Serial"
  .TextMatrix(0, .ColIndex("Name")) = "Description"
  .TextMatrix(0, .ColIndex("Standrd")) = "Standrd"
  .TextMatrix(0, .ColIndex("Stage")) = "Stage"
  .TextMatrix(0, .ColIndex("Comment")) = "Remarks"
  End With
ErrTrap:
End Sub
Private Sub AddNewRecored()
   Dim My_SQL As String
   Dim rs As ADODB.Recordset
  On Error GoTo ErrTrap
    Set rs = New ADODB.Recordset
   My_SQL = "TblBatchSheet"
    rs.Open My_SQL, Cn, adOpenKeyset, adLockOptimistic, adCmdTable
    If rs.RecordCount > 0 Then
        TxtSerial1.Text = rs.RecordCount + 1
    Else
        TxtSerial1.Text = 1
    End If
   rs.Close
ErrTrap:
End Sub
'+++++++++++++++++++++++++++++++++ end

Function print_report(Optional NoteSerial As String)
    Dim MySQL As String
    Dim RsData As New ADODB.Recordset
    Dim xApp As New CRAXDRT.Application
    Dim xReport As CRAXDRT.Report
    Dim CViewer As ClsReportViewer
    Dim StrReportTitle As String
    Dim StrFileName As String
    Dim Msg As String
MySQL = " SELECT     dbo.TblBatchSheet.ID, dbo.TblBatchSheet.RecordDate, dbo.TblBranchesData.branch_name, dbo.TblBranchesData.branch_namee, dbo.TblBatchSheet.BranchID, "
MySQL = MySQL & "                      dbo.TblBatchSheet.BatchNo, dbo.TblBatchSheet.RefNo, dbo.TblBatchSheet.ItemID, dbo.TblItems.ItemName, dbo.TblItems.Fullcode, dbo.TblItems.ItemNamee,"
MySQL = MySQL & "                      dbo.TblBatchSheet.ItemFInshID, TblItems_1.ItemName AS FinishItemName, TblItems_1.Fullcode AS FinishFullcode, TblItems_1.ItemNamee AS FinishItemNameE,"
MySQL = MySQL & "                      dbo.TblBatchSheet.GroupID, dbo.Groups.GroupName, dbo.Groups.Fullcode AS GroupFullcode, dbo.Groups.GroupNamee, dbo.TblBatchSheet.BasedType,"
MySQL = MySQL & "                      dbo.TblBatchSheet.FromDate, dbo.TblBatchSheet.ToDate, dbo.TblBatchSheet.PlanQty, dbo.TblBatchSheet.QtyPermix, dbo.TblBatchSheet.NoPermix,"
MySQL = MySQL & "                      dbo.TblBatchSheet.MasterNo, dbo.TblBatchSheet.SpecSeries, dbo.TblBatchSheet.Remarks, dbo.TblBatchSheet.PermixerMachID, dbo.FixedAssets.Name,"
MySQL = MySQL & "                      dbo.FixedAssets.namee, dbo.TblBatchSheet.GrinderMachID, FixedAssets_1.Name AS GrinderName, FixedAssets_1.namee AS GrinderNameE,"
MySQL = MySQL & "                      dbo.TblBatchSheet.ExtruderMachID, FixedAssets_2.Name AS ExtrudeName, FixedAssets_2.namee AS ExtrudeNameE, dbo.TblBatchSheet.BlenderMachID,"
MySQL = MySQL & "                      FixedAssets_3.Name AS BlenderName, FixedAssets_3.namee AS BlenderNameE, dbo.TblBatchSheetDet.TypeTrans, dbo.TblBatchSheetDet.Stage,"
MySQL = MySQL & "                      dbo.TblBatchSheetDet.Qty, dbo.TblBatchSheetDet.C1, dbo.TblBatchSheetDet.C2, dbo.TblBatchSheetDet.C3, dbo.TblBatchSheetDet.Total,"
MySQL = MySQL & "                      dbo.TblBatchSheetDet.Standrd, dbo.TblBatchSheetDet.UnitID, dbo.TblUnites.UnitName, dbo.TblUnites.UnitNamee, dbo.TblBatchSheetDet.ItemID AS ItemIDDet,"
MySQL = MySQL & "                      TblItems_2.ItemName AS ItemNameDet, TblItems_2.Fullcode AS FullcodeDet, TblItems_2.ItemNamee AS ItemNameeDet, dbo.TblQCItems.name AS Qname,"
MySQL = MySQL & "                      dbo.TblQCItems.namee AS QnameE , dbo.TblBatchSheetDet.Comment"
MySQL = MySQL & " FROM         dbo.TblItems TblItems_2 RIGHT OUTER JOIN"
MySQL = MySQL & "                      dbo.TblQCItems RIGHT OUTER JOIN"
MySQL = MySQL & "                      dbo.TblBatchSheetDet ON dbo.TblQCItems.qcid = dbo.TblBatchSheetDet.ItemID ON TblItems_2.ItemID = dbo.TblBatchSheetDet.ItemID LEFT OUTER JOIN"
MySQL = MySQL & "                      dbo.TblUnites ON dbo.TblBatchSheetDet.UnitID = dbo.TblUnites.UnitID RIGHT OUTER JOIN"
MySQL = MySQL & "                      dbo.TblBatchSheet ON dbo.TblBatchSheetDet.BachShetID = dbo.TblBatchSheet.ID LEFT OUTER JOIN"
MySQL = MySQL & "                      dbo.FixedAssets FixedAssets_3 ON dbo.TblBatchSheet.BlenderMachID = FixedAssets_3.id LEFT OUTER JOIN"
MySQL = MySQL & "                      dbo.FixedAssets FixedAssets_2 ON dbo.TblBatchSheet.ExtruderMachID = FixedAssets_2.id LEFT OUTER JOIN"
MySQL = MySQL & "                      dbo.FixedAssets FixedAssets_1 ON dbo.TblBatchSheet.GrinderMachID = FixedAssets_1.id LEFT OUTER JOIN"
MySQL = MySQL & "                      dbo.FixedAssets ON dbo.TblBatchSheet.PermixerMachID = dbo.FixedAssets.id LEFT OUTER JOIN"
MySQL = MySQL & "                      dbo.Groups ON dbo.TblBatchSheet.GroupID = dbo.Groups.GroupID LEFT OUTER JOIN"
MySQL = MySQL & "                      dbo.TblItems TblItems_1 ON dbo.TblBatchSheet.ItemID = TblItems_1.ItemID LEFT OUTER JOIN"
MySQL = MySQL & "                      dbo.TblItems ON dbo.TblBatchSheet.ItemID = dbo.TblItems.ItemID LEFT OUTER JOIN"
MySQL = MySQL & "                      dbo.TblBranchesData ON dbo.TblBatchSheet.BranchID = dbo.TblBranchesData.branch_id"
MySQL = MySQL & " Where (TblBatchSheet.ID = " & val(TxtSerial1.Text) & ") "

        If SystemOptions.UserInterface = ArabicInterface Then
            StrFileName = App.path & "\REPORTS\REPORTS NEW\" & "RepBatchSheet.rpt"
        Else
            StrFileName = App.path & "\REPORTS\REPORTS NEW\" & "RepBatchSheetE.rpt"
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
        ' xReport.ParameterFields(2).AddCurrentValue RPTComment_Arabic
        StrReportTitle = "" '& StrAccountName
      Else
        xReport.ParameterFields(1).AddCurrentValue cCompanyInfo.EngCompanyName   ' RPTCompany_Name_Eng
        'xReport.ParameterFields(2).AddCurrentValue RPTComment_Eng
        StrReportTitle = ""
 
    End If

    xReport.ParameterFields(3).AddCurrentValue user_name
    xReport.reporttitle = StrReportTitle
    xReport.EnableParameterPrompting = False
    xReport.ApplicationName = App.title
    xReport.ReportAuthor = App.title
    Set CViewer = New ClsReportViewer
    CViewer.FireReport xReport, WindowTarget, "", , , , StrFileName
    RsData.Close
    Set RsData = Nothing
    Screen.MousePointer = vbDefault
     hide_logo = False
 End Function
Function print_report2(Optional NoteSerial As String)
    Dim MySQL As String
    Dim RsData As New ADODB.Recordset
    Dim xApp As New CRAXDRT.Application
    Dim xReport As CRAXDRT.Report
    Dim CViewer As ClsReportViewer
    Dim StrReportTitle As String
    Dim StrFileName As String
    Dim Msg As String
MySQL = " SELECT     dbo.TblBatchSheet.ID, dbo.TblBatchSheet.RecordDate, dbo.TblBranchesData.branch_name, dbo.TblBranchesData.branch_namee, dbo.TblBatchSheet.BranchID, "
MySQL = MySQL & "                      dbo.TblBatchSheet.BatchNo, dbo.TblBatchSheet.RefNo, dbo.TblBatchSheet.ItemID, dbo.TblBatchSheet.ItemFInshID, TblItems_1.ItemName, TblItems_1.Fullcode,"
MySQL = MySQL & "                      TblItems_1.ItemNamee, dbo.TblBatchSheet.GroupID, dbo.Groups.GroupName, dbo.Groups.Fullcode AS GroupFullcode, dbo.Groups.GroupNamee,"
MySQL = MySQL & "                      dbo.TblBatchSheet.BasedType, dbo.TblBatchSheet.FromDate, dbo.TblBatchSheet.ToDate, dbo.TblBatchSheet.PlanQty, dbo.TblBatchSheet.QtyPermix,"
MySQL = MySQL & "                      dbo.TblBatchSheet.NoPermix, dbo.TblBatchSheet.MasterNo, dbo.TblBatchSheet.SpecSeries, dbo.TblBatchSheet.Remarks, dbo.TblBatchSheet.PermixerMachID,"
MySQL = MySQL & "                      FixedAssets_3.Name, FixedAssets_3.namee, dbo.TblBatchSheet.GrinderMachID, FixedAssets_1.Name AS GrinderName, FixedAssets_1.namee AS GrinderNameE,"
MySQL = MySQL & "                      dbo.TblBatchSheet.ExtruderMachID, FixedAssets_2.Name AS ExtrudeName, FixedAssets_2.namee AS ExtrudeNameE, dbo.TblBatchSheet.BlenderMachID,"
MySQL = MySQL & "                      FixedAssets_3.Name AS BlenderName, FixedAssets_3.namee AS BlenderNameE, dbo.TblBatchSheetDet.TypeTrans, dbo.TblBatchSheetDet.Stage,"
MySQL = MySQL & "                      dbo.TblBatchSheetDet.Qty, dbo.TblBatchSheetDet.C1, dbo.TblBatchSheetDet.C2, dbo.TblBatchSheetDet.C3, dbo.TblBatchSheetDet.Total,"
MySQL = MySQL & "                      dbo.TblBatchSheetDet.Standrd, dbo.TblBatchSheetDet.UnitID, dbo.TblBatchSheetDet.ItemID AS ItemIDDet, dbo.TblQCItems.name AS Qname,"
MySQL = MySQL & "                      dbo.TblQCItems.namee AS QnameE , dbo.TblBatchSheetDet.Comment"
MySQL = MySQL & " FROM         dbo.TblQCItems RIGHT OUTER JOIN"
MySQL = MySQL & "                      dbo.TblBatchSheetDet ON dbo.TblQCItems.qcid = dbo.TblBatchSheetDet.ItemID RIGHT OUTER JOIN"
MySQL = MySQL & "                      dbo.TblBatchSheet ON dbo.TblBatchSheetDet.BachShetID = dbo.TblBatchSheet.ID LEFT OUTER JOIN"
MySQL = MySQL & "                      dbo.FixedAssets FixedAssets_3 ON dbo.TblBatchSheet.BlenderMachID = FixedAssets_3.id LEFT OUTER JOIN"
MySQL = MySQL & "                      dbo.FixedAssets FixedAssets_2 ON dbo.TblBatchSheet.ExtruderMachID = FixedAssets_2.id LEFT OUTER JOIN"
MySQL = MySQL & "                      dbo.FixedAssets FixedAssets_1 ON dbo.TblBatchSheet.GrinderMachID = FixedAssets_1.id LEFT OUTER JOIN"
MySQL = MySQL & "                      dbo.FixedAssets FixedAssets_4 ON dbo.TblBatchSheet.PermixerMachID = FixedAssets_4.id LEFT OUTER JOIN"
MySQL = MySQL & "                      dbo.Groups ON dbo.TblBatchSheet.GroupID = dbo.Groups.GroupID LEFT OUTER JOIN"
MySQL = MySQL & "                      dbo.TblItems TblItems_1 ON dbo.TblBatchSheet.ItemID = TblItems_1.ItemID LEFT OUTER JOIN"
MySQL = MySQL & "                      dbo.TblBranchesData ON dbo.TblBatchSheet.BranchID = dbo.TblBranchesData.branch_id"
MySQL = MySQL & " Where (TblBatchSheet.ID = " & val(TxtSerial1.Text) & ") and (dbo.TblBatchSheetDet.TypeTrans = 1) "

        If SystemOptions.UserInterface = ArabicInterface Then
            StrFileName = App.path & "\REPORTS\REPORTS NEW\" & "RepBatchSheet2.rpt"
        Else
            StrFileName = App.path & "\REPORTS\REPORTS NEW\" & "RepBatchSheet2E.rpt"
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
        ' xReport.ParameterFields(2).AddCurrentValue RPTComment_Arabic
        StrReportTitle = "" '& StrAccountName
      Else
        xReport.ParameterFields(1).AddCurrentValue cCompanyInfo.EngCompanyName   ' RPTCompany_Name_Eng
        'xReport.ParameterFields(2).AddCurrentValue RPTComment_Eng
        StrReportTitle = ""
 
    End If

    xReport.ParameterFields(3).AddCurrentValue user_name
    xReport.reporttitle = StrReportTitle
    xReport.EnableParameterPrompting = False
    xReport.ApplicationName = App.title
    xReport.ReportAuthor = App.title
    Set CViewer = New ClsReportViewer
    CViewer.FireReport xReport, WindowTarget, "", , , , StrFileName
    RsData.Close
    Set RsData = Nothing
    Screen.MousePointer = vbDefault
     hide_logo = False
 End Function

Private Sub VSFlexGrid1_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
With Me.VSFlexGrid1
Select Case .ColKey(Col)
Case "Comment"
.ComboList = ""
Case "Name"
Cancel = True
Case "Standrd"
.ComboList = ""
Case "Ser"
Cancel = True

End Select
End With
End Sub

Private Sub XPDtbTrans_Change()
TxtBatchNo.Text = ""
End Sub
