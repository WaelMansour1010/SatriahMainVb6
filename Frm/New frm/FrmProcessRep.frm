VERSION 5.00
Object = "{C115893A-A3BF-43AF-B28D-69DB846077F3}#1.0#0"; "vsflex8u.ocx"
Object = "{0AFE7BE0-11B7-4A3E-978D-D4501E9A57FE}#1.0#0"; "c1sizer.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{FE5DCFAD-BC1D-11D2-94CF-004005455FAA}#1.4#0"; "ImpulseButton.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form FrmProcessRep 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   10065
   ClientLeft      =   1410
   ClientTop       =   2970
   ClientWidth     =   17475
   Icon            =   "FrmProcessRep.frx":0000
   LinkTopic       =   "Form3"
   MDIChild        =   -1  'True
   RightToLeft     =   -1  'True
   ScaleHeight     =   10065
   ScaleWidth      =   17475
   WindowState     =   2  'Maximized
   Begin C1SizerLibCtl.C1Elastic C1Elastic1 
      Height          =   10065
      Left            =   0
      TabIndex        =   26
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
         TabIndex        =   49
         Top             =   0
         Width           =   17505
         Begin ImpulseButton.ISButton btnLast 
            Height          =   315
            Left            =   450
            TabIndex        =   50
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
            ButtonImage     =   "FrmProcessRep.frx":6852
            ColorButton     =   16777215
            AcclimateGrayTones=   -1  'True
            DrawFocusRectangle=   0   'False
            DisabledImageExtraction=   0
            DisabledImageStyle=   1
         End
         Begin ImpulseButton.ISButton btnNext 
            Height          =   315
            Left            =   915
            TabIndex        =   51
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
            ButtonImage     =   "FrmProcessRep.frx":6BEC
            ColorButton     =   16777215
            DrawFocusRectangle=   0   'False
            DisabledImageStyle=   1
         End
         Begin ImpulseButton.ISButton btnPrevious 
            Height          =   315
            Left            =   1515
            TabIndex        =   52
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
            ButtonImage     =   "FrmProcessRep.frx":6F86
            ColorButton     =   16777215
            DrawFocusRectangle=   0   'False
            DisabledImageStyle=   1
         End
         Begin ImpulseButton.ISButton btnFirst 
            Height          =   315
            Left            =   2040
            TabIndex        =   53
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
            ButtonImage     =   "FrmProcessRep.frx":7320
            ColorButton     =   16777215
            DrawFocusRectangle=   0   'False
            DisabledImageStyle=   1
         End
         Begin VB.Image Image1 
            Height          =   615
            Left            =   13200
            Picture         =   "FrmProcessRep.frx":76BA
            Stretch         =   -1  'True
            Top             =   120
            Visible         =   0   'False
            Width           =   735
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Process Report"
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
            TabIndex        =   54
            Top             =   120
            Width           =   4080
         End
      End
      Begin C1SizerLibCtl.C1Elastic C1Elastic2 
         Height          =   735
         Left            =   0
         TabIndex        =   27
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
         Begin VB.TextBox TxtSerial1 
            Alignment       =   1  'Right Justify
            Height          =   315
            Left            =   14520
            RightToLeft     =   -1  'True
            TabIndex        =   65
            Top             =   240
            Width           =   1545
         End
         Begin VB.TextBox TxtBatchNo 
            Alignment       =   1  'Right Justify
            Height          =   315
            Left            =   120
            RightToLeft     =   -1  'True
            TabIndex        =   2
            Top             =   240
            Width           =   2265
         End
         Begin MSComCtl2.DTPicker XPDtbTrans 
            Height          =   315
            Left            =   11385
            TabIndex        =   0
            Top             =   240
            Width           =   1590
            _ExtentX        =   2805
            _ExtentY        =   556
            _Version        =   393216
            Format          =   96206849
            CurrentDate     =   38784
         End
         Begin MSDataListLib.DataCombo Dcbranch 
            Bindings        =   "FrmProcessRep.frx":8ABF
            Height          =   315
            Left            =   3960
            TabIndex        =   1
            Top             =   240
            Width           =   6240
            _ExtentX        =   11007
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
            Caption         =   "—ﬁ„ «· ’ÕÌÕ"
            Height          =   285
            Index           =   5
            Left            =   2760
            TabIndex        =   56
            Top             =   240
            Width           =   870
         End
         Begin VB.Label lbl 
            Alignment       =   2  'Center
            BackColor       =   &H00E2E9E9&
            Caption         =   "«· «—ÌŒ"
            Height          =   285
            Index           =   2
            Left            =   13275
            TabIndex        =   30
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
            TabIndex        =   29
            Top             =   240
            Width           =   930
         End
         Begin VB.Label lbl 
            Alignment       =   2  'Center
            BackColor       =   &H00E2E9E9&
            Caption         =   "«·›—⁄"
            Height          =   285
            Index           =   7
            Left            =   9960
            TabIndex        =   28
            Top             =   240
            Width           =   1620
         End
      End
      Begin C1SizerLibCtl.C1Elastic C1Elastic4 
         Height          =   1695
         Left            =   0
         TabIndex        =   31
         TabStop         =   0   'False
         Top             =   1440
         Width           =   17520
         _cx             =   30903
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
         Begin VB.TextBox TxtBatchSize 
            Alignment       =   1  'Right Justify
            Enabled         =   0   'False
            Height          =   315
            Left            =   8880
            TabIndex        =   67
            TabStop         =   0   'False
            Top             =   1200
            Width           =   2880
         End
         Begin VB.TextBox TxtNoBatch 
            Alignment       =   1  'Right Justify
            Enabled         =   0   'False
            Height          =   315
            Left            =   13125
            TabIndex        =   66
            TabStop         =   0   'False
            Top             =   1200
            Width           =   2880
         End
         Begin VB.TextBox Text4 
            Alignment       =   1  'Right Justify
            Enabled         =   0   'False
            Height          =   315
            Left            =   5835
            TabIndex        =   5
            TabStop         =   0   'False
            Top             =   840
            Width           =   1440
         End
         Begin VB.TextBox TxtRemarks 
            Alignment       =   1  'Right Justify
            Height          =   675
            Left            =   120
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   9
            TabStop         =   0   'False
            Top             =   120
            Width           =   7155
         End
         Begin VB.TextBox TxtSpecSeries 
            Alignment       =   1  'Right Justify
            Enabled         =   0   'False
            Height          =   315
            Left            =   120
            TabIndex        =   12
            TabStop         =   0   'False
            Top             =   1200
            Width           =   2880
         End
         Begin VB.TextBox TxtMasterNo 
            Alignment       =   1  'Right Justify
            Enabled         =   0   'False
            Height          =   315
            Left            =   4395
            TabIndex        =   11
            TabStop         =   0   'False
            Top             =   1200
            Width           =   2880
         End
         Begin VB.TextBox Text3 
            Alignment       =   1  'Right Justify
            Enabled         =   0   'False
            Height          =   315
            Left            =   14565
            TabIndex        =   7
            TabStop         =   0   'False
            Top             =   840
            Width           =   1440
         End
         Begin VB.TextBox TxtItemCode 
            Alignment       =   1  'Right Justify
            Enabled         =   0   'False
            Height          =   315
            Left            =   14565
            TabIndex        =   3
            TabStop         =   0   'False
            Top             =   480
            Width           =   1440
         End
         Begin MSDataListLib.DataCombo DcbItem 
            Bindings        =   "FrmProcessRep.frx":8AD4
            Height          =   315
            Left            =   8880
            TabIndex        =   4
            Top             =   480
            Width           =   5580
            _ExtentX        =   9843
            _ExtentY        =   556
            _Version        =   393216
            Enabled         =   0   'False
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
            Bindings        =   "FrmProcessRep.frx":8AE9
            Height          =   315
            Left            =   8880
            TabIndex        =   8
            Top             =   840
            Width           =   5580
            _ExtentX        =   9843
            _ExtentY        =   556
            _Version        =   393216
            Enabled         =   0   'False
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
            Bindings        =   "FrmProcessRep.frx":8AFE
            Height          =   315
            Left            =   120
            TabIndex        =   6
            Top             =   840
            Width           =   5580
            _ExtentX        =   9843
            _ExtentY        =   556
            _Version        =   393216
            Enabled         =   0   'False
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
            Bindings        =   "FrmProcessRep.frx":8B13
            Height          =   315
            Left            =   8880
            TabIndex        =   10
            Top             =   120
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
         Begin VB.Label lbl 
            Alignment       =   2  'Center
            BackColor       =   &H00E2E9E9&
            Caption         =   "ÕÃ„ «·»« ‘"
            Height          =   285
            Index           =   10
            Left            =   11760
            RightToLeft     =   -1  'True
            TabIndex        =   69
            Top             =   1200
            Width           =   1290
         End
         Begin VB.Label lbl 
            Alignment       =   2  'Center
            BackColor       =   &H00E2E9E9&
            Caption         =   "⁄œœ «·»« ‘"
            Height          =   285
            Index           =   6
            Left            =   16080
            RightToLeft     =   -1  'True
            TabIndex        =   68
            Top             =   1200
            Width           =   1290
         End
         Begin VB.Label lbl 
            Alignment       =   2  'Center
            BackColor       =   &H00E2E9E9&
            Caption         =   "„·«ÕŸ« "
            Height          =   285
            Index           =   21
            Left            =   7440
            RightToLeft     =   -1  'True
            TabIndex        =   63
            Top             =   240
            Width           =   1410
         End
         Begin VB.Label lbl 
            Alignment       =   2  'Center
            BackColor       =   &H00E2E9E9&
            Caption         =   "—ﬁ„ «· Œ’Ì’"
            Height          =   285
            Index           =   20
            Left            =   2880
            RightToLeft     =   -1  'True
            TabIndex        =   62
            Top             =   1200
            Width           =   1410
         End
         Begin VB.Label lbl 
            Alignment       =   2  'Center
            BackColor       =   &H00E2E9E9&
            Caption         =   "—ﬁ„ «· Õﬂ„"
            Height          =   285
            Index           =   19
            Left            =   7440
            RightToLeft     =   -1  'True
            TabIndex        =   61
            Top             =   1200
            Width           =   1290
         End
         Begin VB.Label lbl 
            Alignment       =   2  'Center
            BackColor       =   &H00E2E9E9&
            Caption         =   "«·„⁄œ…"
            Height          =   285
            Index           =   16
            Left            =   16080
            RightToLeft     =   -1  'True
            TabIndex        =   60
            Top             =   120
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
            TabIndex        =   59
            Top             =   840
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
            TabIndex        =   58
            Top             =   840
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
            TabIndex        =   57
            Top             =   480
            Width           =   1290
         End
      End
      Begin C1SizerLibCtl.C1Elastic C1Elastic5 
         Height          =   615
         Left            =   0
         TabIndex        =   32
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
            TabIndex        =   35
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
               TabIndex        =   39
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
               TabIndex        =   38
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
               TabIndex        =   37
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
               TabIndex        =   36
               Top             =   120
               Width           =   780
            End
         End
         Begin MSDataListLib.DataCombo DCboUserName 
            Height          =   315
            Left            =   7785
            TabIndex        =   33
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
            TabIndex        =   34
            Top             =   105
            Width           =   1485
         End
      End
      Begin C1SizerLibCtl.C1Elastic C1Elastic7 
         Height          =   615
         Left            =   0
         TabIndex        =   40
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
            Left            =   16395
            TabIndex        =   41
            ToolTipText     =   "· ”ÃÌ· »Ì«‰«  ÃœÌœ…"
            Top             =   120
            Width           =   870
            _ExtentX        =   1535
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
            ButtonImage     =   "FrmProcessRep.frx":8B28
            ColorButton     =   14871017
            DrawFocusRectangle=   0   'False
         End
         Begin ImpulseButton.ISButton btnSave 
            Height          =   330
            Left            =   14370
            TabIndex        =   42
            ToolTipText     =   "Õ›Ÿ «·»Ì«‰«  œ«Œ· ﬁ«⁄œ… «·»Ì«‰« "
            Top             =   120
            Width           =   855
            _ExtentX        =   1508
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
            ButtonImage     =   "FrmProcessRep.frx":F38A
            ColorButton     =   14871017
            DrawFocusRectangle=   0   'False
         End
         Begin ImpulseButton.ISButton btnModify 
            Height          =   330
            Left            =   15330
            TabIndex        =   43
            ToolTipText     =   "· ⁄œÌ· «·»Ì«‰«  «·Õ«·Ì…"
            Top             =   120
            Width           =   975
            _ExtentX        =   1720
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
            ButtonImage     =   "FrmProcessRep.frx":F724
            ColorButton     =   14871017
            DrawFocusRectangle=   0   'False
         End
         Begin ImpulseButton.ISButton BtnUndo 
            Height          =   330
            Left            =   13290
            TabIndex        =   44
            ToolTipText     =   "·· —«Ã⁄ ⁄‰ «·ÕœÀ Ê«·—ÃÊ⁄ «·Ï «·Ê÷⁄ «·ÿ»Ì⁄Ì"
            Top             =   120
            Width           =   855
            _ExtentX        =   1508
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
            ButtonImage     =   "FrmProcessRep.frx":15F86
            ColorButton     =   14871017
            DrawFocusRectangle=   0   'False
         End
         Begin ImpulseButton.ISButton btnDelete 
            Height          =   330
            Left            =   12345
            TabIndex        =   45
            ToolTipText     =   "Õ–› «·»Ì«‰«  «·„Õœœ…"
            Top             =   120
            Width           =   840
            _ExtentX        =   1482
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
            ButtonImage     =   "FrmProcessRep.frx":16320
            ColorButton     =   14871017
            DrawFocusRectangle=   0   'False
         End
         Begin ImpulseButton.ISButton btnCancel 
            Height          =   330
            Left            =   120
            TabIndex        =   46
            ToolTipText     =   "«·Œ—ÊÃ «·Ï  «·‰«›–… «·—∆Ì”Ì…"
            Top             =   120
            Width           =   975
            _ExtentX        =   1720
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
            ButtonImage     =   "FrmProcessRep.frx":168BA
            ColorButton     =   14871017
            DrawFocusRectangle=   0   'False
            DisabledImageStyle=   1
         End
         Begin ImpulseButton.ISButton ISButton5 
            Height          =   375
            Left            =   11400
            TabIndex        =   47
            TabStop         =   0   'False
            Top             =   120
            Width           =   855
            _ExtentX        =   1508
            _ExtentY        =   661
            ButtonStyle     =   1
            ButtonPositionImage=   1
            Caption         =   "Print Blender "
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
            ButtonImage     =   "FrmProcessRep.frx":16C54
            ColorButton     =   14871017
            DrawFocusRectangle=   0   'False
            DisabledImageStyle=   1
         End
         Begin ImpulseButton.ISButton ISButton8 
            Height          =   330
            Left            =   1305
            TabIndex        =   48
            TabStop         =   0   'False
            ToolTipText     =   "(Ctrl+F)  ··»ÕÀ ≈÷€ÿ Â–« «·„› «Õ √Ê ≈÷€ÿ "
            Top             =   120
            Visible         =   0   'False
            Width           =   840
            _ExtentX        =   1482
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
            ButtonImage     =   "FrmProcessRep.frx":1D4B6
            ColorButton     =   14871017
            DrawFocusRectangle=   0   'False
         End
         Begin ImpulseButton.ISButton ISButton2 
            Height          =   375
            Left            =   9960
            TabIndex        =   70
            TabStop         =   0   'False
            Top             =   120
            Width           =   1335
            _ExtentX        =   2355
            _ExtentY        =   661
            ButtonStyle     =   1
            ButtonPositionImage=   1
            Caption         =   "View Blender "
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
            ButtonImage     =   "FrmProcessRep.frx":1D850
            ColorButton     =   14871017
            DrawFocusRectangle=   0   'False
            DisabledImageStyle=   1
         End
         Begin ImpulseButton.ISButton ISButton3 
            Height          =   375
            Left            =   8880
            TabIndex        =   71
            TabStop         =   0   'False
            Top             =   120
            Width           =   975
            _ExtentX        =   1720
            _ExtentY        =   661
            ButtonStyle     =   1
            ButtonPositionImage=   1
            Caption         =   "Print Premixer"
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
            ButtonImage     =   "FrmProcessRep.frx":240B2
            ColorButton     =   14871017
            DrawFocusRectangle=   0   'False
            DisabledImageStyle=   1
         End
         Begin ImpulseButton.ISButton ISButton4 
            Height          =   375
            Left            =   7560
            TabIndex        =   72
            TabStop         =   0   'False
            Top             =   120
            Width           =   1335
            _ExtentX        =   2355
            _ExtentY        =   661
            ButtonStyle     =   1
            ButtonPositionImage=   1
            Caption         =   "View Premixer"
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
            ButtonImage     =   "FrmProcessRep.frx":2A914
            ColorButton     =   14871017
            DrawFocusRectangle=   0   'False
            DisabledImageStyle=   1
         End
         Begin ImpulseButton.ISButton ISButton6 
            Height          =   375
            Left            =   6600
            TabIndex        =   73
            TabStop         =   0   'False
            Top             =   120
            Width           =   855
            _ExtentX        =   1508
            _ExtentY        =   661
            ButtonStyle     =   1
            ButtonPositionImage=   1
            Caption         =   "Print Extender"
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
            ButtonImage     =   "FrmProcessRep.frx":31176
            ColorButton     =   14871017
            DrawFocusRectangle=   0   'False
            DisabledImageStyle=   1
         End
         Begin ImpulseButton.ISButton ISButton7 
            Height          =   375
            Left            =   5160
            TabIndex        =   74
            TabStop         =   0   'False
            Top             =   120
            Width           =   1335
            _ExtentX        =   2355
            _ExtentY        =   661
            ButtonStyle     =   1
            ButtonPositionImage=   1
            Caption         =   "View Extender"
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
            ButtonImage     =   "FrmProcessRep.frx":379D8
            ColorButton     =   14871017
            DrawFocusRectangle=   0   'False
            DisabledImageStyle=   1
         End
         Begin ImpulseButton.ISButton ISButton9 
            Height          =   375
            Left            =   4080
            TabIndex        =   75
            TabStop         =   0   'False
            Top             =   120
            Width           =   855
            _ExtentX        =   1508
            _ExtentY        =   661
            ButtonStyle     =   1
            ButtonPositionImage=   1
            Caption         =   "Print Grinder "
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
            ButtonImage     =   "FrmProcessRep.frx":3E23A
            ColorButton     =   14871017
            DrawFocusRectangle=   0   'False
            DisabledImageStyle=   1
         End
         Begin ImpulseButton.ISButton ISButton10 
            Height          =   375
            Left            =   2520
            TabIndex        =   76
            TabStop         =   0   'False
            Top             =   120
            Width           =   1335
            _ExtentX        =   2355
            _ExtentY        =   661
            ButtonStyle     =   1
            ButtonPositionImage=   1
            Caption         =   "View Grinder "
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
            ButtonImage     =   "FrmProcessRep.frx":44A9C
            ColorButton     =   14871017
            DrawFocusRectangle=   0   'False
            DisabledImageStyle=   1
         End
      End
      Begin C1SizerLibCtl.C1Elastic C1Elastic3 
         Height          =   5535
         Left            =   0
         TabIndex        =   55
         TabStop         =   0   'False
         Top             =   3210
         Width           =   17475
         _cx             =   30824
         _cy             =   9763
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
            TabIndex        =   13
            Top             =   5190
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
            ButtonImage     =   "FrmProcessRep.frx":4B2FE
            ColorButton     =   14871017
            DrawFocusRectangle=   0   'False
         End
         Begin ImpulseButton.ISButton Cmd 
            Height          =   225
            Index           =   1
            Left            =   14115
            TabIndex        =   14
            Top             =   5190
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
            ButtonImage     =   "FrmProcessRep.frx":4B898
            ColorButton     =   14871017
            DrawFocusRectangle=   0   'False
         End
         Begin VSFlex8UCtl.VSFlexGrid FG 
            Height          =   5115
            Left            =   0
            TabIndex        =   64
            Top             =   0
            Width           =   17310
            _cx             =   30533
            _cy             =   9022
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
            Cols            =   14
            FixedRows       =   1
            FixedCols       =   1
            RowHeightMin    =   300
            RowHeightMax    =   0
            ColWidthMin     =   0
            ColWidthMax     =   0
            ExtendLastCol   =   -1  'True
            FormatString    =   $"FrmProcessRep.frx":4BE32
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
      TabIndex        =   19
      Text            =   "modflag"
      Top             =   4200
      Visible         =   0   'False
      Width           =   465
   End
   Begin VB.ComboBox CmbType 
      BackColor       =   &H80000018&
      Height          =   315
      ItemData        =   "FrmProcessRep.frx":4BFBD
      Left            =   18360
      List            =   "FrmProcessRep.frx":4BFCD
      RightToLeft     =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   18
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
      TabIndex        =   17
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
      TabIndex        =   16
      Top             =   1680
      Visible         =   0   'False
      Width           =   3105
   End
   Begin VB.TextBox Emp_id 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   18600
      RightToLeft     =   -1  'True
      TabIndex        =   15
      Top             =   5640
      Visible         =   0   'False
      Width           =   1575
   End
   Begin MSDataListLib.DataCombo DCUser 
      CausesValidation=   0   'False
      Height          =   315
      Left            =   18720
      TabIndex        =   20
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
      TabIndex        =   21
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
            Picture         =   "FrmProcessRep.frx":4BFE6
            Key             =   "CompanyName"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmProcessRep.frx":4C380
            Key             =   "Ser"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmProcessRep.frx":4C71A
            Key             =   "Vac_Name"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmProcessRep.frx":4CAB4
            Key             =   "ShareCount"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmProcessRep.frx":4CE4E
            Key             =   "Dis_Count"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmProcessRep.frx":4D1E8
            Key             =   "Bouns"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmProcessRep.frx":4D582
            Key             =   "SharesValue"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmProcessRep.frx":4DB1C
            Key             =   "BuyValue"
         EndProperty
      EndProperty
   End
   Begin ImpulseButton.ISButton BtnUpdate 
      Height          =   330
      Left            =   18480
      TabIndex        =   22
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
      ButtonImage     =   "FrmProcessRep.frx":4DEB6
      ColorButton     =   14871017
      DrawFocusRectangle=   0   'False
   End
   Begin ImpulseButton.ISButton ISButton1 
      Height          =   405
      Left            =   18840
      TabIndex        =   24
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
      ButtonImage     =   "FrmProcessRep.frx":54718
      ColorButton     =   14871017
      DrawFocusRectangle=   0   'False
      DisabledImageStyle=   1
   End
   Begin ImpulseButton.ISButton btnQuery 
      Height          =   330
      Left            =   19800
      TabIndex        =   25
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
      ButtonImage     =   "FrmProcessRep.frx":5AF7A
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
      TabIndex        =   23
      Top             =   360
      Width           =   855
   End
End
Attribute VB_Name = "FrmProcessRep"
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



Private Sub Cmd_Click(Index As Integer)
If Me.TxtModFlg.Text <> "R" Then
Select Case Index
Case 0
RemoveGridRow2
Case 1
 FG.Clear flexClearScrollable, flexClearEverything
            FG.Rows = 2
 End Select
End If
End Sub
Private Sub ReLineGrid()
    Dim IntCounter As Integer
    IntCounter = 0
    Dim i As Integer
    With FG
        For i = .FixedRows To .Rows - 1
            If .TextMatrix(i, .ColIndex("ShortDes")) <> "" Then
                IntCounter = IntCounter + 1
                .TextMatrix(i, .ColIndex("Ser")) = IntCounter
            End If
        Next i
    End With
End Sub
Private Sub DcbItem_Change()
DcbItem_Click (0)
End Sub

Private Sub DcbItem_Click(Area As Integer)
Me.TxtItemCode.Text = GetItemCode(val(Me.DcbItem.BoundText))
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
FillGrid val(DcbPermixerMach.BoundText)
End Sub
Sub RetriveBatch(Optional BatchNo As String)
If Me.TxtModFlg.Text <> "R" Then
Dim sql As String
Dim Rs2 As ADODB.Recordset
Set Rs2 = New ADODB.Recordset
sql = " select * from TblBatchSheet  WHERE     (BatchNo = N'" & BatchNo & "')"
Rs2.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
If Rs2.RecordCount > 0 Then
DcbItem.BoundText = IIf(IsNull(Rs2("ItemID").value), 0, Rs2("ItemID").value)
DcbGroup.BoundText = IIf(IsNull(Rs2("GroupID").value), 0, Rs2("GroupID").value)
DcbItemFinish.BoundText = IIf(IsNull(Rs2("GroupID").value), 0, Rs2("GroupID").value)
TxtMasterNo.Text = IIf(IsNull(Rs2("MasterNo").value), "", Rs2("MasterNo").value)
TxtSpecSeries.Text = IIf(IsNull(Rs2("SpecSeries").value), "", Rs2("SpecSeries").value)
TxtNoBatch.Text = IIf(IsNull(Rs2("NoPermix").value), "", Rs2("NoPermix").value)
TxtBatchSize.Text = IIf(IsNull(Rs2("MinQty").value), "", Rs2("MinQty").value)
Else
TxtBatchSize.Text = ""
TxtSpecSeries.Text = ""
TxtMasterNo.Text = ""
TxtNoBatch.Text = ""
DcbItem.BoundText = 0
DcbGroup.BoundText = 0
DcbItemFinish.BoundText = 0
End If
End If
End Sub
Private Sub FG_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
With FG
Select Case .ColKey(Col)
Case "ShortDes"
Cancel = True
Case "Descrip"
Cancel = True
End Select
End With
End Sub
Sub FillGrid(Optional EquipID As Double)
If Me.TxtModFlg.Text <> "R" Then
  FG.Clear flexClearScrollable, flexClearEverything
    FG.Rows = 1
Dim i As Integer
Dim sql As String
Dim Rs2 As ADODB.Recordset
Set Rs2 = New ADODB.Recordset
sql = " SELECT     ID, EquipID, ShortDes, Descrip"
sql = sql & " From dbo.TblEquipmentsDes"
sql = sql & "  Where (EquipID = " & EquipID & ")"
Rs2.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
If Rs2.RecordCount > 0 Then
With FG
.Rows = .Rows + Rs2.RecordCount
Rs2.MoveFirst
For i = 1 To .Rows - 1
.TextMatrix(i, .ColIndex("Ser")) = i
.TextMatrix(i, .ColIndex("ShortDes")) = IIf(IsNull(Rs2("ShortDes").value), "", Rs2("ShortDes").value)
.TextMatrix(i, .ColIndex("SpecID")) = IIf(IsNull(Rs2("ID").value), 0, Rs2("ID").value)
.TextMatrix(i, .ColIndex("Descrip")) = IIf(IsNull(Rs2("Descrip").value), "", Rs2("Descrip").value)
Rs2.MoveNext
Next i
End With
End If
End If
End Sub
    Private Sub Form_Load()
    On Error GoTo ErrTrap
    Dim conection As String
    Dim My_SQL As String
    conection = "select * from TblProcessReport order by  ID "
    Set BKGrndPic = New ClsBackGroundPic
    Set RsSavRec = New ADODB.Recordset
    RsSavRec.CursorLocation = adUseClient
    RsSavRec.Open conection, Cn, adOpenStatic, adLockOptimistic, adCmdText
    Me.TxtModFlg.Text = "R"
    Resize_Form Me
    Dim Dcombos As New ClsDataCombos
    Dcombos.GetBranches Me.Dcbranch
    Dcombos.GetUsers Me.DCboUserName
    Dcombos.GetItemsNamesupdate Me.DcbItem
    Dcombos.GetItemsNamesupdate Me.DcbItemFinish
    Dcombos.GetItemSGroups Me.DcbGroup
    Dcombos.GetEquipments Me.DcbPermixerMach
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



Public Sub FiLLRec()

  '  On Error GoTo ErrTrap
    Dim sql As String
    Dim ID As Double
             If Me.TxtModFlg.Text = "E" Then
                 StrSQL = "Delete From TblProcessReportDet Where ProcessID =" & val(TxtSerial1.Text) & ""
                  Cn.Execute StrSQL, , adExecuteNoRecords
              End If
    RsSavRec.Fields("BranchID").value = val(Me.Dcbranch.BoundText)
    RsSavRec.Fields("RecordDate").value = XPDtbTrans.value
    RsSavRec.Fields("BatchNo").value = TxtBatchNo.Text
    RsSavRec.Fields("PermixerMachID").value = val(DcbPermixerMach.BoundText)
    RsSavRec.Fields("Remarks").value = TxtRemarks.Text
    RsSavRec.Fields("ItemID").value = val(Me.DcbItem.BoundText)
    RsSavRec.Fields("ItemFInshID").value = val(DcbItemFinish.BoundText)
    RsSavRec.Fields("UserID").value = val(Me.DCboUserName.BoundText)
    RsSavRec.Fields("GroupID").value = val(Me.DcbGroup.BoundText)
    RsSavRec.Fields("NoBatch").value = val(TxtNoBatch.Text)
    RsSavRec.Fields("BatchSize").value = val(TxtBatchSize.Text)
    RsSavRec.Fields("MasterNo").value = TxtMasterNo.Text
    RsSavRec.Fields("SpecSeries").value = TxtSpecSeries.Text
    RsSavRec.update
  
''//////////////////////////
     Set RsDevsub = New ADODB.Recordset
    StrSQL = "SELECT  *  from TblProcessReportDet Where (1 = -1)"
    RsDevsub.Open StrSQL, Cn, adOpenKeyset, adLockOptimistic, adCmdText
    Dim i As Integer
    With FG
       For i = .FixedRows To .Rows - 1
       If val(.TextMatrix(i, .ColIndex("SpecID"))) <> 0 Then
       RsDevsub.AddNew
                 RsDevsub("ProcessID").value = val(Me.TxtSerial1.Text)
                 RsDevsub("SpecID").value = IIf((.TextMatrix(i, .ColIndex("SpecID"))) = "", Null, val(.TextMatrix(i, .ColIndex("SpecID"))))
                 RsDevsub("C1").value = IIf((.TextMatrix(i, .ColIndex("C1"))) = "", Null, .TextMatrix(i, .ColIndex("C1")))
                 RsDevsub("C2").value = IIf((.TextMatrix(i, .ColIndex("C2"))) = "", Null, .TextMatrix(i, .ColIndex("C2")))
                 RsDevsub("C3").value = IIf((.TextMatrix(i, .ColIndex("C3"))) = "", Null, .TextMatrix(i, .ColIndex("C3")))
                 RsDevsub("C4").value = IIf((.TextMatrix(i, .ColIndex("C4"))) = "", Null, .TextMatrix(i, .ColIndex("C4")))
                 RsDevsub("C5").value = IIf((.TextMatrix(i, .ColIndex("C5"))) = "", Null, .TextMatrix(i, .ColIndex("C5")))
                 RsDevsub("C6").value = IIf((.TextMatrix(i, .ColIndex("C6"))) = "", Null, .TextMatrix(i, .ColIndex("C6")))
                 RsDevsub("C7").value = IIf((.TextMatrix(i, .ColIndex("C7"))) = "", Null, .TextMatrix(i, .ColIndex("C7")))
                 RsDevsub("C8").value = IIf((.TextMatrix(i, .ColIndex("C8"))) = "", Null, .TextMatrix(i, .ColIndex("C8")))
                 RsDevsub("C9").value = IIf((.TextMatrix(i, .ColIndex("C9"))) = "", Null, .TextMatrix(i, .ColIndex("C9")))
                 RsDevsub("C10").value = IIf((.TextMatrix(i, .ColIndex("C10"))) = "", Null, .TextMatrix(i, .ColIndex("C10")))
       RsDevsub.update
      End If
     Next i
    End With
      Select Case Me.TxtModFlg.Text
        Case "N"
            Dim Msg As String
            If SystemOptions.UserInterface = ArabicInterface Then
                Msg = "  „ Õ›Ÿ »Ì«‰«  Â–Â «·⁄„·Ì… " & Chr(13)
                Msg = Msg + "Â·  —€» ›Ì ≈÷«›… »Ì«‰«  √Œ—Ï"
            Else
               Msg = " This record alredy saved... " & Chr(13)
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
    TxtSerial1.Text = IIf(IsNull(RsSavRec.Fields("ID").value), "", RsSavRec.Fields("ID").value)
    XPDtbTrans.value = IIf(IsNull(RsSavRec.Fields("RecordDate").value), Date, RsSavRec.Fields("RecordDate").value)
    Dcbranch.BoundText = IIf(IsNull(RsSavRec.Fields("BranchID").value), "", RsSavRec.Fields("BranchID").value)
    TxtBatchNo.Text = IIf(IsNull(RsSavRec.Fields("BatchNo").value), "", RsSavRec.Fields("BatchNo").value)
    Me.DcbItem.BoundText = IIf(IsNull(RsSavRec.Fields("ItemID").value), 0, RsSavRec.Fields("ItemID").value)
    Me.DcbItemFinish.BoundText = IIf(IsNull(RsSavRec.Fields("ItemFInshID").value), 0, RsSavRec.Fields("ItemFInshID").value)   ': ProgressBar1.value = 90
    Me.DcbGroup.BoundText = IIf(IsNull(RsSavRec.Fields("GroupID").value), "", RsSavRec.Fields("GroupID").value)
    Me.DCboUserName.BoundText = IIf(IsNull(RsSavRec.Fields("UserID").value), 0, RsSavRec.Fields("UserID").value)  ': ProgressBar1.value = 10
    Me.DcbPermixerMach.BoundText = IIf(IsNull(RsSavRec.Fields("PermixerMachID").value), "", RsSavRec.Fields("PermixerMachID").value)
    TxtMasterNo.Text = IIf(IsNull(RsSavRec.Fields("MasterNo").value), "", RsSavRec.Fields("MasterNo").value)
    TxtSpecSeries.Text = IIf(IsNull(RsSavRec.Fields("SpecSeries").value), "", RsSavRec.Fields("SpecSeries").value)
    TxtRemarks.Text = IIf(IsNull(RsSavRec.Fields("Remarks").value), "", RsSavRec.Fields("Remarks").value)
    TxtNoBatch.Text = IIf(IsNull(RsSavRec.Fields("NoBatch").value), "", RsSavRec.Fields("NoBatch").value)
    TxtBatchSize.Text = IIf(IsNull(RsSavRec.Fields("BatchSize").value), "", RsSavRec.Fields("BatchSize").value)
     LabCurrRec.Caption = RsSavRec.AbsolutePosition ': ProgressBar1.value = 50
     LabCountRec.Caption = RsSavRec.RecordCount ': ProgressBar1.value = 60
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
'            DcbItem.SetFocus
            Exit Sub
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
    StrRecID = new_id("TblProcessReport", "ID", "")
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
sql = " SELECT     dbo.TblProcessReportDet.SpecID, dbo.TblProcessReportDet.ProcessID, dbo.TblProcessReportDet.C1, dbo.TblProcessReportDet.C2, dbo.TblProcessReportDet.C3, "
sql = sql + "                      dbo.TblProcessReportDet.C4, dbo.TblProcessReportDet.C5, dbo.TblProcessReportDet.C6, dbo.TblProcessReportDet.C7, dbo.TblProcessReportDet.C8,"
sql = sql + "                      dbo.TblProcessReportDet.C9 , dbo.TblProcessReportDet.C10, dbo.TblEquipmentsDes.ShortDes, dbo.TblEquipmentsDes.Descrip"
sql = sql + " FROM         dbo.TblProcessReportDet LEFT OUTER JOIN"
sql = sql + "                      dbo.TblEquipmentsDes ON dbo.TblProcessReportDet.SpecID = dbo.TblEquipmentsDes.ID"
sql = sql + " Where (dbo.TblProcessReportDet.ProcessID = " & val(TxtSerial1.Text) & ")"
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
                   .TextMatrix(i, .ColIndex("SpecID")) = IIf(IsNull(Rs1("SpecID").value), 0, Rs1("SpecID").value)
                   .TextMatrix(i, .ColIndex("ShortDes")) = IIf(IsNull(Rs1("ShortDes").value), "", Rs1("ShortDes").value)
                   .TextMatrix(i, .ColIndex("Descrip")) = IIf(IsNull(Rs1("Descrip").value), "", Rs1("Descrip").value)
                   .TextMatrix(i, .ColIndex("C1")) = IIf(IsNull(Rs1("C1").value), "", Rs1("C1").value)
                   .TextMatrix(i, .ColIndex("C2")) = IIf(IsNull(Rs1("C2").value), "", Rs1("C2").value)
                   .TextMatrix(i, .ColIndex("C3")) = IIf(IsNull(Rs1("C3").value), "", Rs1("C3").value)
                   .TextMatrix(i, .ColIndex("C4")) = IIf(IsNull(Rs1("C4").value), "", Rs1("C4").value)
                   .TextMatrix(i, .ColIndex("C5")) = IIf(IsNull(Rs1("C5").value), "", Rs1("C5").value)
                   .TextMatrix(i, .ColIndex("C6")) = IIf(IsNull(Rs1("C6").value), "", Rs1("C6").value)
                   .TextMatrix(i, .ColIndex("C7")) = IIf(IsNull(Rs1("C7").value), "", Rs1("C7").value)
                   .TextMatrix(i, .ColIndex("C8")) = IIf(IsNull(Rs1("C8").value), "", Rs1("C8").value)
                   .TextMatrix(i, .ColIndex("C9")) = IIf(IsNull(Rs1("C9").value), "", Rs1("C9").value)
                   .TextMatrix(i, .ColIndex("C10")) = IIf(IsNull(Rs1("C10").value), "", Rs1("C10").value)
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

Private Sub ISButton10_Click()
print_report2 , 3
End Sub

Private Sub ISButton2_Click()
print_report3 , 1
End Sub

Private Sub ISButton3_Click()
print_report , 0
End Sub

Private Sub ISButton4_Click()
print_report , 1
End Sub

Private Sub ISButton5_Click()
Dim i As Integer
For i = 1 To val(TxtNoBatch.Text)
print_report3 , 0
DoEvents
Next i
End Sub

Private Sub ISButton6_Click()
Dim i As Integer
For i = 1 To val(TxtNoBatch.Text)
print_report2 , 0
DoEvents
Next i
End Sub

Private Sub ISButton7_Click()
print_report2 , 1
End Sub

Private Sub ISButton9_Click()
Dim i As Integer
For i = 1 To val(TxtNoBatch.Text)
print_report2 , 2
DoEvents
Next i
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

Private Sub TxtBatchNo_Change()
If TxtBatchNo.Text <> "" Then
RetriveBatch TxtBatchNo.Text
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
      StrSQL = "delete from   TblProcessReportDet where ProcessID =" & val(TxtSerial1.Text) & ""
      Cn.Execute StrSQL
                RsSavRec.find "ID=" & val(TxtSerial1.Text), , adSearchForward, 1
                                          RsSavRec.delete
            FG.Clear flexClearScrollable, flexClearEverything
            FG.Rows = 1
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
                    StrMSG = "You will close this screen before save " & Chr(13)
                    StrMSG = StrMSG & " the new data  " & Chr(13)
                    StrMSG = StrMSG & " do you want save before exit" & Chr(13)
                    StrMSG = StrMSG & "yes" & "-" & "save the new data" & Chr(13)
                    StrMSG = StrMSG & "no" & "-" & "Don't save" & Chr(13)
                    StrMSG = StrMSG & "cancel" & "-" & "Cancel Exit" & Chr(13)
                    Else
                    StrMSG = "”Ê› Ì „ €·ﬁ «·‘«‘… Ê·„  ‰ Â „‰  ”ÃÌ·" & Chr(13)
                    StrMSG = StrMSG & " «·»Ì«‰«  «·ÃœÌœ… «·Õ«·Ì… " & Chr(13)
                    StrMSG = StrMSG & " Â·  —Ìœ «·Õ›Ÿ ﬁ»· «·Œ—ÊÃ" & Chr(13)
                    StrMSG = StrMSG & "‰⁄„" & "-" & "Ì „ Õ›Ÿ «·»Ì«‰«  «·ÃœÌœ…" & Chr(13)
                    StrMSG = StrMSG & "·«" & "-" & "·‰ Ì „ «·Õ›Ÿ" & Chr(13)
                    StrMSG = StrMSG & "≈·€«¡ «·√„—" & "-" & "≈·€«¡ ⁄„·Ì… «·Œ—ÊÃ" & Chr(13)
                        End If
                    Case "E"
                If SystemOptions.UserInterface = EnglishInterface Then
                    StrMSG = "You will close this screen before save  " & Chr(13)
                    StrMSG = StrMSG & " the Modifications  " & Chr(13)
                    StrMSG = StrMSG & " do you want save before exit" & Chr(13)
                    StrMSG = StrMSG & "yes" & "-" & "save the new data" & Chr(13)
                    StrMSG = StrMSG & "no" & "-" & "Don't save" & Chr(13)
                    StrMSG = StrMSG & "cancel" & "-" & "Cancel Exit" & Chr(13)
                    Else
                    StrMSG = "”Ê› Ì „ €·ﬁ «·‘«‘… Ê·„  ‰ Â „‰  ”ÃÌ·" & Chr(13)
                    StrMSG = StrMSG & " «· ⁄œÌ·«  «·ÃœÌœ… ⁄·Ï «·”Ã· «·Õ«·Ï " & Chr(13)
                    StrMSG = StrMSG & " Â·  —Ìœ «·Õ›Ÿ ﬁ»· «·Œ—ÊÃ" & Chr(13)
                    StrMSG = StrMSG & "‰⁄„" & "-" & "Ì „ Õ›Ÿ «· ⁄œÌ·«   «·ÃœÌœ…" & Chr(13)
                    StrMSG = StrMSG & "·«" & "-" & "·‰ Ì „ «·Õ›Ÿ" & Chr(13)
                    StrMSG = StrMSG & "≈·€«¡ «·√„—" & "-" & "≈·€«¡ ⁄„·Ì… «·Œ—ÊÃ" & Chr(13)
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
            Msg = "⁄›Ê« ·ﬁœ  „ Õ–› Â–« «·”Ã· «· «·Ï" & Chr(13)
            Msg = Msg & "„‰ ﬁ»· „” Œœ„ √Œ— ⁄·Ï «·‘»ﬂ… " & Chr(13)
            Msg = Msg & "”Ê› Ì „  ÕœÌÀ «·»Ì«‰« "
            Else
            Msg = "Sorry I've been to delete this record" & Chr(13)
            Msg = Msg & "By another user on the network " & Chr(13)
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
            Msg = "⁄›Ê« ·ﬁœ  „ Õ–› Â–« «·”Ã· «· «·Ï" & Chr(13)
            Msg = Msg & "„‰ ﬁ»· „” Œœ„ √Œ— ⁄·Ï «·‘»ﬂ… " & Chr(13)
            Msg = Msg & "”Ê› Ì „  ÕœÌÀ «·»Ì«‰« "
            Else
               Msg = "Sorry I've been to delete this record" & Chr(13)
            Msg = Msg & "By another user on the network " & Chr(13)
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
       
        Me.DCboUserName.BoundText = user_id
        Me.Dcbranch.SetFocus
    End If
    Exit Sub
ErrTrap:
    Select Case Err.Number
        Case -2147467259
            'Could not update; currently locked.
            If SystemOptions.UserInterface = ArabicInterface Then
            Msg = "⁄›Ê«" & Chr(13)
            Msg = Msg & " ·«Ì„ﬂ‰  ⁄œÌ· Â–« «·”Ã· ›Ï «·Êﬁ  «·Õ«·Ï" & Chr(13)
            Msg = Msg & "ÕÌÀ «‰Â ﬁÌœ «· ⁄œÌ· „‰ ﬁ»· „” Œœ„ «Œ— ⁄·Ï «·‘»ﬂ…"
            Else
            Msg = "Sorry.." & Chr(13)
            Msg = Msg & " You can not edit this the record now" & Chr(13)
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
    Me.DCboUserName.BoundText = user_id
    Me.Dcbranch.BoundText = Current_branch
    Dcbranch.SetFocus
    XPDtbTrans.value = Date
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
            Msg = "⁄›Ê« ·ﬁœ  „ Õ–› Â–« «·”Ã· «· «·Ï" & Chr(13)
            Msg = Msg & "„‰ ﬁ»· „” Œœ„ √Œ— ⁄·Ï «·‘»ﬂ… " & Chr(13)
            Msg = Msg & "”Ê› Ì „  ÕœÌÀ «·»Ì«‰« "
            Else
               Msg = "Sorry I've been to delete this record" & Chr(13)
            Msg = Msg & "By another user on the network " & Chr(13)
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
            Msg = "⁄›Ê« ·ﬁœ  „ Õ–› Â–« «·”Ã· «· «·Ï" & Chr(13)
            Msg = Msg & "„‰ ﬁ»· „” Œœ„ √Œ— ⁄·Ï «·‘»ﬂ… " & Chr(13)
            Msg = Msg & "”Ê› Ì „  ÕœÌÀ «·»Ì«‰« "
            Else
            Msg = "Sorry I've been to delete this record" & Chr(13)
            Msg = Msg & "By another user on the network " & Chr(13)
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
    Wrap = Chr(13) + Chr(10)
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
       Me.Caption = "Process Report"
      Label1(2).Caption = Me.Caption
      Me.lbl(4).Caption = "ID"
      Me.lbl(2).Caption = "Date"
      lbl(7).Caption = "Branch"
      lbl(5).Caption = "Batch No"
      lbl(16).Caption = "Equipment"
   lbl(0).Caption = "Product Name"
   lbl(1).Caption = "Product Type"
   lbl(3).Caption = "Finish"
   lbl(21).Caption = "Remarks"
lbl(6).Caption = "No of Batch"
lbl(10).Caption = "Batch Size"
lbl(19).Caption = "Master No"
lbl(20).Caption = "Spec Series"
Cmd(0).Caption = "Delete"
Cmd(1).Caption = "Delete All"

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
   
  With Me.FG
  .TextMatrix(0, .ColIndex("Ser")) = "Serial"
  .TextMatrix(0, .ColIndex("ShortDes")) = "Short Description "
  .TextMatrix(0, .ColIndex("Descrip")) = "Description"
  .TextMatrix(0, .ColIndex("C1")) = "1"
  .TextMatrix(0, .ColIndex("C2")) = "2"
  .TextMatrix(0, .ColIndex("C3")) = "3"
  .TextMatrix(0, .ColIndex("C4")) = "4"
  .TextMatrix(0, .ColIndex("C5")) = "5"
  .TextMatrix(0, .ColIndex("C6")) = "6"
  .TextMatrix(0, .ColIndex("C7")) = "7"
  .TextMatrix(0, .ColIndex("C8")) = "8"
  .TextMatrix(0, .ColIndex("C9")) = "9"
  .TextMatrix(0, .ColIndex("C10")) = "10"
  End With

ErrTrap:
End Sub
Private Sub AddNewRecored()
   Dim My_SQL As String
   Dim rs As ADODB.Recordset
  On Error GoTo ErrTrap
    Set rs = New ADODB.Recordset
   My_SQL = "TblProcessReport"
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
Function print_report(Optional NoteSerial As String, Optional PrintTyp As Integer = 0)
    Dim MySQL As String
    Dim RsData As New ADODB.Recordset
    Dim xApp As New CRAXDRT.Application
    Dim xReport As CRAXDRT.Report
    Dim CViewer As ClsReportViewer
    Dim StrReportTitle As String
    Dim StrFileName As String
    Dim sql As String
    Dim Msg As String
    Dim i As Integer
    Dim Rs2 As ADODB.Recordset
    Set Rs2 = New ADODB.Recordset
    Cn.Execute " Delete from TblProcessReportDet2 where ProcessID=" & val(Me.TxtSerial1.Text) & ""
    sql = "Select * from TblProcessReportDet2 where 1=-1"
    Rs2.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
    For i = 0 To val(TxtNoBatch.Text) - 1
    Rs2.AddNew
    Rs2("ProcessID").value = val(Me.TxtSerial1.Text)
    Rs2("PrimixNo").value = i
    Rs2.update
    Next i
MySQL = " SELECT        dbo.TblProcessReport.BranchID, dbo.TblProcessReport.ID, dbo.TblBranchesData.branch_name, dbo.TblBranchesData.branch_namee, dbo.TblProcessReport.RecordDate, dbo.TblProcessReport.BatchNo, "
MySQL = MySQL & "                         dbo.TblProcessReport.ItemID, TblItems_2.ItemCode, TblItems_2.ItemName, TblItems_2.ItemNamee, dbo.TblProcessReport.ItemFInshID, TblItems_1.ItemCode AS FinishItemCode,"
MySQL = MySQL & "                         TblItems_1.ItemName AS FinishItemName, TblItems_1.ItemNamee AS FinishItemNameE, dbo.TblProcessReport.Remarks, dbo.TblProcessReport.GroupID, dbo.Groups.GroupName, dbo.Groups.GroupCode,"
MySQL = MySQL & "                         dbo.Groups.GroupNamee, dbo.TblProcessReport.NoBatch, dbo.TblProcessReport.BatchSize, dbo.TblProcessReport.MasterNo, dbo.TblProcessReport.SpecSeries, dbo.TblProcessReport.PermixerMachID,"
MySQL = MySQL & "                         dbo.FixedAssets.Name , dbo.FixedAssets.NameE, dbo.TblProcessReportDet2.PrimixNo"
MySQL = MySQL & " FROM            dbo.Groups RIGHT OUTER JOIN"
MySQL = MySQL & "                         dbo.FixedAssets RIGHT OUTER JOIN"
MySQL = MySQL & "                         dbo.TblProcessReport LEFT OUTER JOIN"
MySQL = MySQL & "                         dbo.TblProcessReportDet2 ON dbo.TblProcessReport.ID = dbo.TblProcessReportDet2.ProcessID ON dbo.FixedAssets.id = dbo.TblProcessReport.PermixerMachID ON"
MySQL = MySQL & "                         dbo.Groups.GroupID = dbo.TblProcessReport.GroupID LEFT OUTER JOIN"
MySQL = MySQL & "                         dbo.TblItems AS TblItems_1 ON dbo.TblProcessReport.ItemFInshID = TblItems_1.ItemID LEFT OUTER JOIN"
MySQL = MySQL & "                         dbo.TblItems AS TblItems_2 ON dbo.TblProcessReport.ItemID = TblItems_2.ItemID LEFT OUTER JOIN"
MySQL = MySQL & "                         dbo.TblBranchesData ON dbo.TblProcessReport.BranchID = dbo.TblBranchesData.branch_id"
MySQL = MySQL & " Where (dbo.TblProcessReport.ID = " & val(TxtSerial1.Text) & ")"
        If SystemOptions.UserInterface = ArabicInterface Then
            StrFileName = App.path & "\REPORTS\REPORTS NEW\" & "RepProcessReport.rpt"
        Else
            StrFileName = App.path & "\REPORTS\REPORTS NEW\" & "RepProcessReportE.rpt"
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
    
    If PrintTyp = 0 Then
    CViewer.FireReport xReport, PrinterTarget, "", , , , StrFileName
    Else
    CViewer.FireReport xReport, WindowTarget, "", , , , StrFileName
    End If
    RsData.Close
    Set RsData = Nothing
    Screen.MousePointer = vbDefault
     hide_logo = False
 End Function
Function print_report3(Optional NoteSerial As String, Optional PrintTyp As Integer = 0)
    Dim MySQL As String
    Dim RsData As New ADODB.Recordset
    Dim xApp As New CRAXDRT.Application
    Dim xReport As CRAXDRT.Report
    Dim CViewer As ClsReportViewer
    Dim StrReportTitle As String
    Dim StrFileName As String
    Dim sql As String
    Dim Msg As String

MySQL = " SELECT        dbo.TblProcessReport.BranchID, dbo.TblProcessReport.ID, dbo.TblBranchesData.branch_name, dbo.TblBranchesData.branch_namee, dbo.TblProcessReport.RecordDate, dbo.TblProcessReport.BatchNo, "
MySQL = MySQL & "                         dbo.TblProcessReport.ItemID, TblItems_2.ItemCode, TblItems_2.ItemName, TblItems_2.ItemNamee, dbo.TblProcessReport.ItemFInshID, TblItems_1.ItemCode AS FinishItemCode,"
MySQL = MySQL & "                         TblItems_1.ItemName AS FinishItemName, TblItems_1.ItemNamee AS FinishItemNameE, dbo.TblProcessReport.Remarks, dbo.TblProcessReport.GroupID, dbo.Groups.GroupName, dbo.Groups.GroupCode,"
MySQL = MySQL & "                         dbo.Groups.GroupNamee, dbo.TblProcessReport.NoBatch, dbo.TblProcessReport.BatchSize, dbo.TblProcessReport.MasterNo, dbo.TblProcessReport.SpecSeries, dbo.TblProcessReport.PermixerMachID,"
MySQL = MySQL & "                         dbo.FixedAssets.Name , dbo.FixedAssets.NameE, dbo.TblProcessReportDet2.PrimixNo"
MySQL = MySQL & " FROM            dbo.Groups RIGHT OUTER JOIN"
MySQL = MySQL & "                         dbo.FixedAssets RIGHT OUTER JOIN"
MySQL = MySQL & "                         dbo.TblProcessReport LEFT OUTER JOIN"
MySQL = MySQL & "                         dbo.TblProcessReportDet2 ON dbo.TblProcessReport.ID = dbo.TblProcessReportDet2.ProcessID ON dbo.FixedAssets.id = dbo.TblProcessReport.PermixerMachID ON"
MySQL = MySQL & "                         dbo.Groups.GroupID = dbo.TblProcessReport.GroupID LEFT OUTER JOIN"
MySQL = MySQL & "                         dbo.TblItems AS TblItems_1 ON dbo.TblProcessReport.ItemFInshID = TblItems_1.ItemID LEFT OUTER JOIN"
MySQL = MySQL & "                         dbo.TblItems AS TblItems_2 ON dbo.TblProcessReport.ItemID = TblItems_2.ItemID LEFT OUTER JOIN"
MySQL = MySQL & "                         dbo.TblBranchesData ON dbo.TblProcessReport.BranchID = dbo.TblBranchesData.branch_id"
MySQL = MySQL & " Where (dbo.TblProcessReport.ID = " & val(TxtSerial1.Text) & ")"
        If SystemOptions.UserInterface = ArabicInterface Then
            StrFileName = App.path & "\REPORTS\REPORTS NEW\" & "RepProcessReport1.rpt"
        Else
            StrFileName = App.path & "\REPORTS\REPORTS NEW\" & "RepProcessReport1E.rpt"
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
    
    If PrintTyp = 0 Then
    CViewer.FireReport xReport, PrinterTarget, "", , , , StrFileName
    Else
    CViewer.FireReport xReport, WindowTarget, "", , , , StrFileName
    End If
    RsData.Close
    Set RsData = Nothing
    Screen.MousePointer = vbDefault
     hide_logo = False
 End Function
Function print_report2(Optional NoteSerial As String, Optional PrintTyp As Integer = 0)
    Dim MySQL As String
    Dim RsData As New ADODB.Recordset
    Dim xApp As New CRAXDRT.Application
    Dim xReport As CRAXDRT.Report
    Dim CViewer As ClsReportViewer
    Dim StrReportTitle As String
    Dim StrFileName As String
    Dim Msg As String
MySQL = " SELECT        dbo.TblProcessReport.BranchID, dbo.TblProcessReport.ID, dbo.TblBranchesData.branch_name, dbo.TblBranchesData.branch_namee, dbo.TblProcessReport.RecordDate, dbo.TblProcessReport.BatchNo, "
MySQL = MySQL & "                         dbo.TblProcessReport.ItemID, TblItems_2.ItemCode, TblItems_2.ItemName, TblItems_2.ItemNamee, dbo.TblProcessReport.ItemFInshID, TblItems_1.ItemCode AS FinishItemCode,"
MySQL = MySQL & "                         TblItems_1.ItemName AS FinishItemName, TblItems_1.ItemNamee AS FinishItemNameE, dbo.TblProcessReport.Remarks, dbo.TblProcessReport.GroupID, Groups_1.GroupName, Groups_1.GroupCode,"
MySQL = MySQL & "                         Groups_1.GroupNamee, dbo.TblProcessReport.NoBatch, dbo.TblProcessReport.BatchSize, dbo.TblProcessReport.MasterNo, dbo.TblProcessReport.SpecSeries, dbo.TblProcessReport.PermixerMachID,"
MySQL = MySQL & "                         dbo.FixedAssets.Name, dbo.FixedAssets.namee, TblItems_2.GroupID AS GroupID2, dbo.Groups.GroupName AS GroupName2, dbo.Groups.GroupNamee AS GroupName2E,"
MySQL = MySQL & "                         dbo.Groups.GroupCode AS GroupCode2E, dbo.TblGroupQCItems.Standrd, dbo.TblGroupQCItems.Comment, dbo.TblGroupQCItems.GRINDER, dbo.TblGroupQCItems.EXTENDER, dbo.TblGroupQCItems.GQCID,"
MySQL = MySQL & "                         dbo.TblQCItems.name AS QCname, dbo.TblQCItems.namee AS QCnameE"
MySQL = MySQL & " FROM            dbo.TblQCItems RIGHT OUTER JOIN"
MySQL = MySQL & "                         dbo.TblGroupQCItems ON dbo.TblQCItems.qcid = dbo.TblGroupQCItems.GQCID RIGHT OUTER JOIN"
MySQL = MySQL & "                         dbo.Groups ON dbo.TblGroupQCItems.GroupID = dbo.Groups.GroupID RIGHT OUTER JOIN"
MySQL = MySQL & "                         dbo.TblItems AS TblItems_2 ON dbo.Groups.GroupID = TblItems_2.GroupID RIGHT OUTER JOIN"
MySQL = MySQL & "                         dbo.Groups AS Groups_1 RIGHT OUTER JOIN"
MySQL = MySQL & "                         dbo.FixedAssets RIGHT OUTER JOIN"
MySQL = MySQL & "                         dbo.TblProcessReport ON dbo.FixedAssets.id = dbo.TblProcessReport.PermixerMachID ON Groups_1.GroupID = dbo.TblProcessReport.GroupID LEFT OUTER JOIN"
MySQL = MySQL & "                         dbo.TblItems AS TblItems_1 ON dbo.TblProcessReport.ItemFInshID = TblItems_1.ItemID ON TblItems_2.ItemID = dbo.TblProcessReport.ItemID LEFT OUTER JOIN"
MySQL = MySQL & "                         dbo.TblBranchesData ON dbo.TblProcessReport.BranchID = dbo.TblBranchesData.branch_id"
MySQL = MySQL & " Where (dbo.TblProcessReport.ID = " & val(TxtSerial1.Text) & ")"
If PrintTyp = 0 Or PrintTyp = 1 Then
MySQL = MySQL & " and isnull(dbo.TblGroupQCItems.Extender,0)=1 "
Else
MySQL = MySQL & " and isnull(dbo.TblGroupQCItems.GRINDER,0)=1 "
End If
If PrintTyp = 0 Or PrintTyp = 1 Then
        If SystemOptions.UserInterface = ArabicInterface Then
            StrFileName = App.path & "\REPORTS\REPORTS NEW\" & "RepProcessReport2.rpt"
        Else
            StrFileName = App.path & "\REPORTS\REPORTS NEW\" & "RepProcessReport2E.rpt"
        End If
 Else
         If SystemOptions.UserInterface = ArabicInterface Then
            StrFileName = App.path & "\REPORTS\REPORTS NEW\" & "RepProcessReport3.rpt"
        Else
            StrFileName = App.path & "\REPORTS\REPORTS NEW\" & "RepProcessReport3E.rpt"
        End If
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
    
    If PrintTyp = 0 Or PrintTyp = 2 Then
    CViewer.FireReport xReport, PrinterTarget, "", , , , StrFileName
    Else
    CViewer.FireReport xReport, WindowTarget, "", , , , StrFileName
    End If
    RsData.Close
    Set RsData = Nothing
    Screen.MousePointer = vbDefault
     hide_logo = False
 End Function


