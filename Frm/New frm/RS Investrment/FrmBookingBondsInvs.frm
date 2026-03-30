VERSION 5.00
Object = "{C115893A-A3BF-43AF-B28D-69DB846077F3}#1.0#0"; "vsflex8u.ocx"
Object = "{0AFE7BE0-11B7-4A3E-978D-D4501E9A57FE}#1.0#0"; "c1sizer.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{FE5DCFAD-BC1D-11D2-94CF-004005455FAA}#1.4#0"; "ImpulseButton.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{D155F1AE-D9A4-458C-8CEE-498CB717DB7B}#1.0#0"; "DBPix20.ocx"
Begin VB.Form FrmBookingBondsInvs 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   10065
   ClientLeft      =   1410
   ClientTop       =   2970
   ClientWidth     =   17475
   Icon            =   "FrmBookingBondsInvs.frx":0000
   LinkTopic       =   "Form3"
   MDIChild        =   -1  'True
   RightToLeft     =   -1  'True
   ScaleHeight     =   10065
   ScaleWidth      =   17475
   WindowState     =   2  'Maximized
   Begin C1SizerLibCtl.C1Elastic C1Elastic1 
      Height          =   10065
      Left            =   0
      TabIndex        =   33
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
         TabIndex        =   62
         Top             =   0
         Width           =   17508
         Begin VB.TextBox tXTRootAccount 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   3240
            RightToLeft     =   -1  'True
            TabIndex        =   64
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
            TabIndex        =   63
            Top             =   360
            Visible         =   0   'False
            Width           =   2055
         End
         Begin ImpulseButton.ISButton btnLast 
            Height          =   315
            Left            =   450
            TabIndex        =   65
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
            ButtonImage     =   "FrmBookingBondsInvs.frx":6852
            ColorButton     =   16777215
            AcclimateGrayTones=   -1  'True
            DrawFocusRectangle=   0   'False
            DisabledImageExtraction=   0
            DisabledImageStyle=   1
         End
         Begin ImpulseButton.ISButton btnNext 
            Height          =   315
            Left            =   915
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
            ButtonImage     =   "FrmBookingBondsInvs.frx":6BEC
            ColorButton     =   16777215
            DrawFocusRectangle=   0   'False
            DisabledImageStyle=   1
         End
         Begin ImpulseButton.ISButton btnPrevious 
            Height          =   315
            Left            =   1515
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
            ButtonImage     =   "FrmBookingBondsInvs.frx":6F86
            ColorButton     =   16777215
            DrawFocusRectangle=   0   'False
            DisabledImageStyle=   1
         End
         Begin ImpulseButton.ISButton btnFirst 
            Height          =   315
            Left            =   2040
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
            ButtonImage     =   "FrmBookingBondsInvs.frx":7320
            ColorButton     =   16777215
            DrawFocusRectangle=   0   'False
            DisabledImageStyle=   1
         End
         Begin VB.Image Image1 
            Height          =   615
            Left            =   13200
            Picture         =   "FrmBookingBondsInvs.frx":76BA
            Stretch         =   -1  'True
            Top             =   120
            Visible         =   0   'False
            Width           =   735
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "”‰œ ÕÃ“"
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
            Height          =   375
            Index           =   2
            Left            =   8880
            RightToLeft     =   -1  'True
            TabIndex        =   69
            Top             =   120
            Width           =   4080
         End
      End
      Begin C1SizerLibCtl.C1Elastic C1Elastic2 
         Height          =   732
         Left            =   0
         TabIndex        =   34
         TabStop         =   0   'False
         Top             =   720
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
            Alignment       =   2  'Center
            BackColor       =   &H00FFFFC0&
            Enabled         =   0   'False
            Height          =   312
            Left            =   13920
            RightToLeft     =   -1  'True
            TabIndex        =   0
            Top             =   240
            Width           =   2040
         End
         Begin MSComCtl2.DTPicker XPDtbTrans 
            Height          =   312
            Left            =   8028
            TabIndex        =   1
            Top             =   240
            Width           =   1596
            _ExtentX        =   2805
            _ExtentY        =   556
            _Version        =   393216
            Format          =   94175233
            CurrentDate     =   38784
         End
         Begin MSDataListLib.DataCombo Dcbranch 
            Bindings        =   "FrmBookingBondsInvs.frx":8ABF
            Height          =   288
            Left            =   360
            TabIndex        =   2
            Top             =   240
            Width           =   6000
            _ExtentX        =   10583
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
            Caption         =   "ﬂÊœ «·„‘—Ê⁄"
            Height          =   288
            Index           =   9
            Left            =   3240
            TabIndex        =   73
            Top             =   240
            Visible         =   0   'False
            Width           =   1608
         End
         Begin VB.Label lbl 
            Alignment       =   2  'Center
            BackColor       =   &H00E2E9E9&
            Caption         =   "«· «—ÌŒ"
            Height          =   288
            Index           =   2
            Left            =   10392
            TabIndex        =   37
            Top             =   240
            Width           =   900
         End
         Begin VB.Label lbl 
            Alignment       =   2  'Center
            BackColor       =   &H00E2E9E9&
            Caption         =   "—ﬁ„ «·Õ—ﬂ…"
            Height          =   288
            Index           =   4
            Left            =   16308
            RightToLeft     =   -1  'True
            TabIndex        =   36
            Top             =   240
            Width           =   936
         End
         Begin VB.Label lbl 
            Alignment       =   2  'Center
            BackColor       =   &H00E2E9E9&
            Caption         =   "«·›—⁄"
            Height          =   288
            Index           =   7
            Left            =   6228
            TabIndex        =   35
            Top             =   240
            Width           =   1620
         End
      End
      Begin C1SizerLibCtl.C1Elastic C1Elastic4 
         Height          =   972
         Left            =   0
         TabIndex        =   38
         TabStop         =   0   'False
         Top             =   1560
         Width           =   17520
         _cx             =   30903
         _cy             =   1720
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
         Begin VB.TextBox TxtSquareCode 
            Alignment       =   1  'Right Justify
            Height          =   312
            Left            =   5412
            RightToLeft     =   -1  'True
            TabIndex        =   9
            Top             =   480
            Width           =   948
         End
         Begin VB.TextBox TxtTelephone 
            Alignment       =   1  'Right Justify
            Height          =   312
            Left            =   8028
            RightToLeft     =   -1  'True
            TabIndex        =   4
            Top             =   120
            Width           =   2028
         End
         Begin VB.TextBox TxtNameP 
            Alignment       =   1  'Right Justify
            Height          =   312
            Left            =   11280
            RightToLeft     =   -1  'True
            TabIndex        =   3
            Top             =   120
            Width           =   4680
         End
         Begin VB.TextBox TxtSearchCode 
            Alignment       =   1  'Right Justify
            Height          =   312
            Left            =   15012
            RightToLeft     =   -1  'True
            TabIndex        =   7
            Top             =   480
            Width           =   948
         End
         Begin VB.ComboBox DcbTypePrstg 
            Height          =   288
            ItemData        =   "FrmBookingBondsInvs.frx":8AD4
            Left            =   360
            List            =   "FrmBookingBondsInvs.frx":8AD6
            RightToLeft     =   -1  'True
            TabIndex        =   6
            Top             =   120
            Width           =   2412
         End
         Begin MSDataListLib.DataCombo DcbProject 
            Bindings        =   "FrmBookingBondsInvs.frx":8AD8
            Height          =   288
            Left            =   8028
            TabIndex        =   8
            Top             =   480
            Width           =   6852
            _ExtentX        =   12091
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
         Begin MSDataListLib.DataCombo DcbBank 
            Bindings        =   "FrmBookingBondsInvs.frx":8AED
            Height          =   288
            Left            =   3960
            TabIndex        =   5
            Top             =   120
            Width           =   2400
            _ExtentX        =   4233
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
         Begin MSDataListLib.DataCombo DcbSquare 
            Bindings        =   "FrmBookingBondsInvs.frx":8B02
            Height          =   288
            Left            =   360
            TabIndex        =   10
            Top             =   480
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
            Caption         =   "ÿ—Ìﬁ… «·œ›⁄"
            Height          =   288
            Index           =   11
            Left            =   2880
            RightToLeft     =   -1  'True
            TabIndex        =   75
            Top             =   120
            Width           =   1032
         End
         Begin VB.Label lbl 
            Alignment       =   2  'Center
            BackColor       =   &H00E2E9E9&
            Caption         =   "«·„—»⁄"
            Height          =   288
            Index           =   10
            Left            =   6348
            RightToLeft     =   -1  'True
            TabIndex        =   74
            Top             =   480
            Width           =   1776
         End
         Begin VB.Label lbl 
            Alignment       =   2  'Center
            BackColor       =   &H00E2E9E9&
            Caption         =   "«·Â« ›"
            Height          =   288
            Index           =   3
            Left            =   10320
            TabIndex        =   42
            Top             =   120
            Width           =   756
         End
         Begin VB.Label lbl 
            Alignment       =   2  'Center
            BackColor       =   &H00E2E9E9&
            Caption         =   "«”„ «·Õ«Ã“"
            Height          =   288
            Index           =   1
            Left            =   16068
            TabIndex        =   41
            Top             =   120
            Width           =   1368
         End
         Begin VB.Label lbl 
            Alignment       =   2  'Center
            BackColor       =   &H00E2E9E9&
            Caption         =   "«·„‘—Ê⁄"
            Height          =   288
            Index           =   15
            Left            =   16068
            TabIndex        =   40
            Top             =   480
            Width           =   1368
         End
         Begin VB.Label lbl 
            Alignment       =   2  'Center
            BackColor       =   &H00E2E9E9&
            Caption         =   "«·»‰ﬂ"
            Height          =   288
            Index           =   0
            Left            =   6348
            RightToLeft     =   -1  'True
            TabIndex        =   39
            Top             =   120
            Width           =   1776
         End
      End
      Begin C1SizerLibCtl.C1Elastic C1Elastic5 
         Height          =   615
         Left            =   0
         TabIndex        =   43
         TabStop         =   0   'False
         Top             =   8880
         Width           =   17520
         _cx             =   30903
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
         Begin VB.CommandButton Command1 
            Caption         =   "› Õ «·„Œÿÿ"
            Height          =   372
            Left            =   5160
            RightToLeft     =   -1  'True
            TabIndex        =   88
            Top             =   120
            Width           =   1812
         End
         Begin C1SizerLibCtl.C1Elastic C1Elastic6 
            Height          =   405
            Left            =   360
            TabIndex        =   46
            TabStop         =   0   'False
            Top             =   105
            Width           =   4095
            _cx             =   7223
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
               Left            =   2985
               RightToLeft     =   -1  'True
               TabIndex        =   50
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
               TabIndex        =   49
               Top             =   120
               Width           =   975
            End
            Begin VB.Label LabCurrRec 
               Alignment       =   2  'Center
               BackColor       =   &H00E2E9E9&
               ForeColor       =   &H00800000&
               Height          =   210
               Left            =   2160
               RightToLeft     =   -1  'True
               TabIndex        =   48
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
               TabIndex        =   47
               Top             =   120
               Width           =   780
            End
         End
         Begin MSDataListLib.DataCombo DCboUserName 
            Height          =   315
            Left            =   7800
            TabIndex        =   44
            Top             =   105
            Width           =   5940
            _ExtentX        =   10478
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
            Left            =   14040
            TabIndex        =   45
            Top             =   105
            Width           =   1500
         End
      End
      Begin C1SizerLibCtl.C1Elastic C1Elastic7 
         Height          =   615
         Left            =   0
         TabIndex        =   51
         TabStop         =   0   'False
         Top             =   9480
         Width           =   17520
         _cx             =   30903
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
         Begin ImpulseButton.ISButton btnNew 
            Height          =   330
            Left            =   14760
            TabIndex        =   52
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
            ButtonImage     =   "FrmBookingBondsInvs.frx":8B17
            ColorButton     =   14871017
            DrawFocusRectangle=   0   'False
         End
         Begin ImpulseButton.ISButton btnSave 
            Height          =   330
            Left            =   10920
            TabIndex        =   53
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
            ButtonImage     =   "FrmBookingBondsInvs.frx":F379
            ColorButton     =   14871017
            DrawFocusRectangle=   0   'False
         End
         Begin ImpulseButton.ISButton btnModify 
            Height          =   330
            Left            =   13200
            TabIndex        =   54
            ToolTipText     =   "· ⁄œÌ· «·»Ì«‰«  «·Õ«·Ì…"
            Top             =   120
            Width           =   1230
            _ExtentX        =   2170
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
            ButtonImage     =   "FrmBookingBondsInvs.frx":F713
            ColorButton     =   14871017
            DrawFocusRectangle=   0   'False
         End
         Begin ImpulseButton.ISButton BtnUndo 
            Height          =   330
            Left            =   9240
            TabIndex        =   55
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
            ButtonImage     =   "FrmBookingBondsInvs.frx":15F75
            ColorButton     =   14871017
            DrawFocusRectangle=   0   'False
         End
         Begin ImpulseButton.ISButton btnDelete 
            Height          =   330
            Left            =   7560
            TabIndex        =   56
            ToolTipText     =   "Õ–› «·»Ì«‰«  «·„Õœœ…"
            Top             =   120
            Width           =   1335
            _ExtentX        =   2355
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
            ButtonImage     =   "FrmBookingBondsInvs.frx":1630F
            ColorButton     =   14871017
            DrawFocusRectangle=   0   'False
         End
         Begin ImpulseButton.ISButton btnCancel 
            Height          =   330
            Left            =   2160
            TabIndex        =   57
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
            ButtonImage     =   "FrmBookingBondsInvs.frx":168A9
            ColorButton     =   14871017
            DrawFocusRectangle=   0   'False
            DisabledImageStyle=   1
         End
         Begin ImpulseButton.ISButton ISButton5 
            Height          =   405
            Left            =   6000
            TabIndex        =   58
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
            ButtonImage     =   "FrmBookingBondsInvs.frx":16C43
            ColorButton     =   14871017
            DrawFocusRectangle=   0   'False
            DisabledImageStyle=   1
         End
         Begin ImpulseButton.ISButton ISButton8 
            Height          =   330
            Left            =   3960
            TabIndex        =   59
            TabStop         =   0   'False
            ToolTipText     =   "(Ctrl+F)  ··»ÕÀ ≈÷€ÿ Â–« «·„› «Õ √Ê ≈÷€ÿ "
            Top             =   120
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
            ButtonImage     =   "FrmBookingBondsInvs.frx":1D4A5
            ColorButton     =   14871017
            DrawFocusRectangle=   0   'False
         End
      End
      Begin C1SizerLibCtl.C1Elastic C1Elastic3 
         Height          =   6735
         Left            =   0
         TabIndex        =   60
         TabStop         =   0   'False
         Top             =   2640
         Width           =   17400
         _cx             =   30692
         _cy             =   11880
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
         Begin VB.TextBox Txt_Path_General_photo 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   120
            RightToLeft     =   -1  'True
            TabIndex        =   93
            Top             =   1440
            Visible         =   0   'False
            Width           =   495
         End
         Begin VB.CommandButton Command3 
            Caption         =   "«œ—«Ã ’Ê—…"
            Height          =   372
            Left            =   0
            RightToLeft     =   -1  'True
            TabIndex        =   91
            Top             =   5640
            Visible         =   0   'False
            Width           =   2868
         End
         Begin MSComDlg.CommonDialog CommonDialog1 
            Left            =   7200
            Top             =   5760
            _ExtentX        =   847
            _ExtentY        =   847
            _Version        =   393216
         End
         Begin VB.CommandButton Command2 
            Caption         =   " Õ„Ì· «·„Œÿÿ"
            Height          =   372
            Left            =   15972
            RightToLeft     =   -1  'True
            TabIndex        =   90
            Top             =   840
            Width           =   1080
         End
         Begin VB.TextBox Txt_path_photo 
            Alignment       =   1  'Right Justify
            Height          =   372
            Left            =   240
            RightToLeft     =   -1  'True
            TabIndex        =   89
            Top             =   840
            Width           =   15624
         End
         Begin VB.TextBox TxtPartCode 
            Alignment       =   1  'Right Justify
            Height          =   312
            Left            =   240
            RightToLeft     =   -1  'True
            TabIndex        =   86
            Top             =   120
            Width           =   1296
         End
         Begin VB.TextBox TxtHouseAreaFrom 
            Alignment       =   1  'Right Justify
            Height          =   312
            Left            =   4872
            RightToLeft     =   -1  'True
            TabIndex        =   19
            Top             =   480
            Width           =   1296
         End
         Begin VB.TextBox TxtHouseAreaTo 
            Alignment       =   1  'Right Justify
            Height          =   312
            Left            =   2616
            RightToLeft     =   -1  'True
            TabIndex        =   20
            Top             =   480
            Width           =   1308
         End
         Begin VB.TextBox TxtPriceTo 
            Alignment       =   1  'Right Justify
            Height          =   312
            Left            =   12276
            RightToLeft     =   -1  'True
            TabIndex        =   16
            Top             =   480
            Width           =   1296
         End
         Begin VB.ComboBox DcbConstStatus 
            Height          =   288
            ItemData        =   "FrmBookingBondsInvs.frx":1D83F
            Left            =   7848
            List            =   "FrmBookingBondsInvs.frx":1D841
            RightToLeft     =   -1  'True
            TabIndex        =   12
            Top             =   120
            Width           =   3264
         End
         Begin VB.TextBox TxtPriceFrom 
            Alignment       =   1  'Right Justify
            Height          =   312
            Left            =   14556
            RightToLeft     =   -1  'True
            TabIndex        =   15
            Top             =   480
            Width           =   1296
         End
         Begin VB.TextBox TxtRoomFrom 
            Alignment       =   1  'Right Justify
            Height          =   312
            Left            =   9804
            RightToLeft     =   -1  'True
            TabIndex        =   17
            Top             =   480
            Width           =   1308
         End
         Begin VB.TextBox TxtRoomTo 
            Alignment       =   1  'Right Justify
            Height          =   312
            Left            =   7848
            RightToLeft     =   -1  'True
            TabIndex        =   18
            Top             =   480
            Width           =   1308
         End
         Begin VB.TextBox TxtLanAreaTo 
            Alignment       =   1  'Right Justify
            Height          =   312
            Left            =   2616
            RightToLeft     =   -1  'True
            TabIndex        =   14
            Top             =   120
            Width           =   1308
         End
         Begin VB.TextBox TxtLanAreaFrom 
            Alignment       =   1  'Right Justify
            Height          =   312
            Left            =   4872
            RightToLeft     =   -1  'True
            TabIndex        =   13
            Top             =   120
            Width           =   1296
         End
         Begin VSFlex8UCtl.VSFlexGrid Grid 
            Height          =   1992
            Left            =   3216
            TabIndex        =   61
            Top             =   1320
            Width           =   13980
            _cx             =   24659
            _cy             =   3514
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
            Cols            =   22
            FixedRows       =   1
            FixedCols       =   1
            RowHeightMin    =   320
            RowHeightMax    =   0
            ColWidthMin     =   0
            ColWidthMax     =   0
            ExtendLastCol   =   -1  'True
            FormatString    =   $"FrmBookingBondsInvs.frx":1D843
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
            Index           =   0
            Left            =   15600
            TabIndex        =   70
            Top             =   5880
            Width           =   1530
            _ExtentX        =   2699
            _ExtentY        =   476
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
            ButtonImage     =   "FrmBookingBondsInvs.frx":1DBB5
            ColorButton     =   14871017
            DrawFocusRectangle=   0   'False
         End
         Begin ImpulseButton.ISButton Cmd 
            Height          =   270
            Index           =   1
            Left            =   13935
            TabIndex        =   71
            Top             =   5880
            Width           =   1500
            _ExtentX        =   2646
            _ExtentY        =   476
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
            ButtonImage     =   "FrmBookingBondsInvs.frx":1E14F
            ColorButton     =   14871017
            DrawFocusRectangle=   0   'False
         End
         Begin ImpulseButton.ISButton ISButton3 
            Height          =   372
            Left            =   240
            TabIndex        =   21
            Top             =   480
            Width           =   2040
            _ExtentX        =   3598
            _ExtentY        =   661
            Caption         =   "⁄—÷"
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
            ButtonImage     =   "FrmBookingBondsInvs.frx":1E6E9
            ColorButton     =   14871017
            DrawFocusRectangle=   0   'False
            DisabledImageExtraction=   0
            LowerToggledContent=   0   'False
         End
         Begin MSDataListLib.DataCombo DcbModel 
            Bindings        =   "FrmBookingBondsInvs.frx":24F4B
            Height          =   288
            Left            =   12276
            TabIndex        =   11
            Top             =   120
            Width           =   3576
            _ExtentX        =   6324
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
         Begin VSFlex8UCtl.VSFlexGrid VSFlexGrid1 
            Height          =   2475
            Left            =   3120
            TabIndex        =   85
            Top             =   3480
            Width           =   14070
            _cx             =   24818
            _cy             =   4366
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
            Cols            =   23
            FixedRows       =   1
            FixedCols       =   1
            RowHeightMin    =   320
            RowHeightMax    =   0
            ColWidthMin     =   0
            ColWidthMax     =   0
            ExtendLastCol   =   -1  'True
            FormatString    =   $"FrmBookingBondsInvs.frx":24F60
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
         Begin DBPIXLib.DBPix20 DBPix201 
            Height          =   4200
            Left            =   240
            TabIndex        =   92
            Top             =   1320
            Width           =   2865
            _Version        =   131072
            _ExtentX        =   5054
            _ExtentY        =   7408
            _StockProps     =   1
            BackColor       =   16777152
            _Image          =   "FrmBookingBondsInvs.frx":252EF
            ImageResampleWidth=   100
            ImageResampleHeight=   100
            ImageResampleMode=   1
            ImageSaveFormat =   0
            JPEGQuality     =   75
            JPEGEncoding    =   0
            JPEGColorMode   =   0
            JPEGNoRecompress=   -1  'True
            JPEGRotateWarning=   0
            PNGColorDepth   =   0
            PNGCompression  =   0
            PNGFilter       =   0
            PNGInterlace    =   1
            ImageDitherMethod=   3
            ImagePaletteMethod=   4
            ImagePreviewMode=   0   'False
            ImageKeepMetaData=   -1  'True
            UseAmbientBackcolor=   -1  'True
            ViewAsyncDecoding=   -1  'True
            ViewEnableMouseZoom=   -1  'True
            ViewInitialZoom =   0
            ViewHAlign      =   1
            ViewVAlign      =   1
            ViewMenuMode    =   0
         End
         Begin VB.Label lbl 
            Alignment       =   2  'Center
            BackColor       =   &H00E2E9E9&
            Caption         =   "ﬂÊœ «·ﬁÿ⁄…"
            Height          =   288
            Index           =   21
            Left            =   1548
            RightToLeft     =   -1  'True
            TabIndex        =   87
            Top             =   120
            Width           =   1020
         End
         Begin VB.Label lbl 
            Alignment       =   2  'Center
            BackColor       =   &H00E2E9E9&
            Caption         =   "„”«Õ… «·»‰«¡  „‰"
            Height          =   288
            Index           =   20
            Left            =   6060
            RightToLeft     =   -1  'True
            TabIndex        =   84
            Top             =   480
            Width           =   1764
         End
         Begin VB.Label lbl 
            Alignment       =   2  'Center
            BackColor       =   &H00E2E9E9&
            Caption         =   "«·Ï"
            Height          =   288
            Index           =   19
            Left            =   3456
            RightToLeft     =   -1  'True
            TabIndex        =   83
            Top             =   480
            Width           =   1740
         End
         Begin VB.Label lbl 
            Alignment       =   2  'Center
            BackColor       =   &H00E2E9E9&
            Caption         =   "«·”⁄— „‰"
            Height          =   288
            Index           =   18
            Left            =   15840
            RightToLeft     =   -1  'True
            TabIndex        =   82
            Top             =   480
            Width           =   1356
         End
         Begin VB.Label lbl 
            Alignment       =   2  'Center
            BackColor       =   &H00E2E9E9&
            Caption         =   "«·Ï"
            Height          =   288
            Index           =   17
            Left            =   12984
            RightToLeft     =   -1  'True
            TabIndex        =   81
            Top             =   480
            Width           =   1728
         End
         Begin VB.Label lbl 
            Alignment       =   2  'Center
            BackColor       =   &H00E2E9E9&
            Caption         =   "⁄œœ «·€—› „‰"
            Height          =   288
            Index           =   16
            Left            =   10848
            RightToLeft     =   -1  'True
            TabIndex        =   80
            Top             =   480
            Width           =   1728
         End
         Begin VB.Label lbl 
            Alignment       =   2  'Center
            BackColor       =   &H00E2E9E9&
            Caption         =   "«·Ï"
            Height          =   288
            Index           =   14
            Left            =   8688
            RightToLeft     =   -1  'True
            TabIndex        =   79
            Top             =   480
            Width           =   1740
         End
         Begin VB.Label lbl 
            Alignment       =   2  'Center
            BackColor       =   &H00E2E9E9&
            Caption         =   "«·Ï"
            Height          =   288
            Index           =   13
            Left            =   3456
            RightToLeft     =   -1  'True
            TabIndex        =   78
            Top             =   120
            Width           =   1740
         End
         Begin VB.Label lbl 
            Alignment       =   2  'Center
            BackColor       =   &H00E2E9E9&
            Caption         =   "„”«Õ… «·«—÷ „‰"
            Height          =   288
            Index           =   5
            Left            =   6060
            RightToLeft     =   -1  'True
            TabIndex        =   77
            Top             =   120
            Width           =   1764
         End
         Begin VB.Label lbl 
            Alignment       =   2  'Center
            BackColor       =   &H00E2E9E9&
            Caption         =   "«·Õ«·… «·«‰‘«∆Ì…"
            Height          =   288
            Index           =   12
            Left            =   10968
            TabIndex        =   76
            Top             =   120
            Width           =   1344
         End
         Begin VB.Label lbl 
            Alignment       =   2  'Center
            BackColor       =   &H00E2E9E9&
            Caption         =   "«·‰„Ê–Ã"
            Height          =   288
            Index           =   6
            Left            =   15840
            TabIndex        =   72
            Top             =   120
            Width           =   1356
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
      TabIndex        =   26
      Text            =   "modflag"
      Top             =   4200
      Visible         =   0   'False
      Width           =   465
   End
   Begin VB.ComboBox CmbType 
      BackColor       =   &H80000018&
      Height          =   315
      ItemData        =   "FrmBookingBondsInvs.frx":25307
      Left            =   18360
      List            =   "FrmBookingBondsInvs.frx":25317
      RightToLeft     =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   25
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
      TabIndex        =   24
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
      TabIndex        =   23
      Top             =   1680
      Visible         =   0   'False
      Width           =   3105
   End
   Begin VB.TextBox Emp_id 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   18600
      RightToLeft     =   -1  'True
      TabIndex        =   22
      Top             =   5640
      Visible         =   0   'False
      Width           =   1575
   End
   Begin MSDataListLib.DataCombo DCUser 
      CausesValidation=   0   'False
      Height          =   315
      Left            =   18720
      TabIndex        =   27
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
      TabIndex        =   28
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
            Picture         =   "FrmBookingBondsInvs.frx":25330
            Key             =   "CompanyName"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmBookingBondsInvs.frx":256CA
            Key             =   "Ser"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmBookingBondsInvs.frx":25A64
            Key             =   "Vac_Name"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmBookingBondsInvs.frx":25DFE
            Key             =   "ShareCount"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmBookingBondsInvs.frx":26198
            Key             =   "Dis_Count"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmBookingBondsInvs.frx":26532
            Key             =   "Bouns"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmBookingBondsInvs.frx":268CC
            Key             =   "SharesValue"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmBookingBondsInvs.frx":26E66
            Key             =   "BuyValue"
         EndProperty
      EndProperty
   End
   Begin ImpulseButton.ISButton BtnUpdate 
      Height          =   330
      Left            =   18480
      TabIndex        =   29
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
      ButtonImage     =   "FrmBookingBondsInvs.frx":27200
      ColorButton     =   14871017
      DrawFocusRectangle=   0   'False
   End
   Begin ImpulseButton.ISButton ISButton1 
      Height          =   405
      Left            =   18840
      TabIndex        =   31
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
      ButtonImage     =   "FrmBookingBondsInvs.frx":2DA62
      ColorButton     =   14871017
      DrawFocusRectangle=   0   'False
      DisabledImageStyle=   1
   End
   Begin ImpulseButton.ISButton btnQuery 
      Height          =   330
      Left            =   19800
      TabIndex        =   32
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
      ButtonImage     =   "FrmBookingBondsInvs.frx":342C4
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
      TabIndex        =   30
      Top             =   360
      Width           =   855
   End
End
Attribute VB_Name = "FrmBookingBondsInvs"
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
RemoveGridRow
Case 1
 VSFlexGrid1.Clear flexClearScrollable, flexClearEverything
            VSFlexGrid1.Rows = 2
 End Select
End If
End Sub
Private Sub ReLineGrid()
    Dim IntCounter As Integer
    IntCounter = 0
    Dim i As Integer
    With Me.Grid
        For i = .FixedRows To .Rows - 1
            If .TextMatrix(i, .ColIndex("BlockNo")) <> "" Then
                IntCounter = IntCounter + 1
                .TextMatrix(i, .ColIndex("Ser")) = IntCounter
              '  .TextMatrix(i, .ColIndex("PartCode")) = Me.TxtProjectCode.Text & Me.TxtSquareCode.Text & .TextMatrix(i, .ColIndex("BlockNo")) & .TextMatrix(i, .ColIndex("PartNo"))
              ' .TextMatrix(i, .ColIndex("Total")) = val(.TextMatrix(i, .ColIndex("AddValue"))) + val(.TextMatrix(i, .ColIndex("ValueOffice"))) + val(.TextMatrix(i, .ColIndex("MOHPrice")))
            End If
        Next i
    End With
End Sub

Private Sub Command1_Click()

If Txt_path_photo.Text = "" Then
MsgBox ("„‰ ›÷·ﬂ Õœœ «·’Ê—… «Ê·«")
Exit Sub
End If


Dim iFileNo As Integer
iFileNo = FreeFile
'StrFileName = App.path &" \"& "PathPhoto.txt"
 Open App.path & "\PathPhoto.txt" For Output As #iFileNo

'Open "D:\Developers Code\PathPhoto.txt" For Output As #iFileNo

 Print #iFileNo, Txt_path_photo.Text
 Print #iFileNo, TxtSerial1.Text

 Close #iFileNo

 Shell App.path & "\Dwaween.exe"



'MsgBox App.path
'Open "D:\Developers Code\PathPhoto.txt" For Output As #iFileNo

'Print #iFileNo, Txt_path_photo.Text
'Print #iFileNo, TxtSerial1.Text

'Close #iFileNo

'Shell "D:\Developers Code\DocumentProcessing1.exe"

End Sub

Private Sub Command2_Click()
'CommonDialog1.filter = "Pic(*.Jpg)|*.Jpg|All files (*.*)|*.*"
CommonDialog1.filter = "Pic(*.Jpg)|*.Jpg"
CommonDialog1.InitDir = App.path & "\Images"
CommonDialog1.DefaultExt = "txt"
CommonDialog1.DialogTitle = "Select File"
CommonDialog1.ShowOpen
Txt_path_photo.Text = CommonDialog1.filename
End Sub

Private Sub Command3_Click()
Dim X As String

 'If xptxtid.text = "" Then Exit Sub
    If SystemOptions.UserInterface = ArabicInterface Then
        X = MsgBox("Â·  —Ìœ ’Ê—… „‰ „·›", vbExclamation + vbYesNoCancel)
    Else
        X = MsgBox("Do you want to upload photo from file", vbExclamation + vbYesNoCancel)
    End If
    If X = vbYes Then
        DBPix201.ImageLoad
        DoEvents
        If SystemOptions.UserInterface = ArabicInterface Then
            MsgBox " „  Õ„Ì· «·’Ê—…"
        Else
            MsgBox "Photo was uploaded"
        End If
    Else

        If X = vbNo Then
            DBPix201.TWAINAcquire
            If SystemOptions.UserInterface = ArabicInterface Then
                MsgBox " „ „”Õ ÷Ê∆Ì  ··’Ê—…"
            Else
                MsgBox "Photo was scanned "
            End If
            DoEvents
        Else
            Exit Sub
        End If
    End If
End Sub

Private Sub DBPix201_Click()
On Error Resume Next
If Txt_Path_General_photo = "" Then Exit Sub
   Load FrmViewPic
   Set FrmViewPic.MainView.Picture = LoadPicture(Txt_Path_General_photo)
    
   FrmViewPic.show vbModal
    
    
End Sub

Private Sub DcbProject_Change()
On Error Resume Next
DcbProject_Click (0)
End Sub

Private Sub DcbProject_Click(Area As Integer)
On Error Resume Next

If DcbProject.BoundText <> "" Then
TxtSearchCode.Text = DcbProject.BoundText
    Dim Dcombos As New ClsDataCombos
    Dcombos.GetProjecInvestmentSquare Me.DcbSquare, DcbProject.BoundText
    RetriveProjectData DcbProject.BoundText
End If
End Sub


Private Sub DcbSquare_Change()
DcbSquare_Click (0)
End Sub

Private Sub DcbSquare_Click(Area As Integer)
Dim code As String
 GetID_CodeSqureProject code, val(DcbSquare.BoundText)
TxtSquareCode.Text = code
End Sub

    Private Sub Form_Load()
    On Error GoTo ErrTrap
    Dim conection As String
    Dim My_SQL As String
    If SystemOptions.UserInterface = ArabicInterface Then
    With DcbTypePrstg
    .Clear
    .AddItem "ﬁÌ„…"
    .AddItem "‰”»…"
    End With
    With DcbConstStatus
    .Clear
    .AddItem "·„ Ì „ «·»‰«¡"
    .AddItem " Õ  «·«‰‘«¡"
    .AddItem " „ «·»‰«¡"
    .AddItem "„ Êﬁ›"
    End With
   Else
       With DcbConstStatus
    .Clear
    .AddItem "Not Built"
    .AddItem "Under Construction"
    .AddItem "Built"
    .AddItem "Stopped"
    End With
   With DcbTypePrstg
    .Clear
    .AddItem "Value"
    .AddItem "Percentage"
    End With
   End If
      If SystemOptions.UserInterface = ArabicInterface Then
                Grid.ColComboList(Grid.ColIndex("ConstStatus")) = "#1;·„ Ì „ «·»‰«¡|#2; Õ  «·≈‰‘«¡|#3; „ «·»‰«¡|#4;„ Êﬁ›"
            ElseIf SystemOptions.UserInterface = EnglishInterface Then
               Grid.ColComboList(Grid.ColIndex("ConstStatus")) = "#1;Not Built |#2;Under Construction|#3; Built |#4;Stopped"
            End If
    conection = "select * from TblBookingBondsInvs order by  ID "
    Set BKGrndPic = New ClsBackGroundPic
    Set RsSavRec = New ADODB.Recordset
    RsSavRec.CursorLocation = adUseClient
    RsSavRec.Open conection, Cn, adOpenStatic, adLockOptimistic, adCmdText
    Me.TxtModFlg.Text = "R"
    Resize_Form Me
    Dim Dcombos As New ClsDataCombos
    Dcombos.GetBranches Me.Dcbranch
    Dcombos.GetUsers Me.DCboUserName
    Dcombos.GetProjecInvestment Me.DcbProject
    Dcombos.GetProjecInvestmentSquare Me.DcbSquare
    Dcombos.GetModelInvestment Me.DcbModel
    Dcombos.GetBanks Me.DcbBank
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
                 StrSQL = "Delete From TblBookingBondsInvsDet Where BokID =" & val(TxtSerial1.Text) & ""
                  Cn.Execute StrSQL, , adExecuteNoRecords
              End If
    RsSavRec.Fields("BrnchID").value = val(Me.Dcbranch.BoundText)
    RsSavRec.Fields("RecordDate").value = XPDtbTrans.value
    RsSavRec.Fields("Name").value = TxtNameP.Text
    RsSavRec.Fields("Telephone").value = TxtTelephone.Text
    RsSavRec.Fields("BanckID").value = val(Me.DcbBank.BoundText)
    RsSavRec.Fields("PaymentID").value = val(DcbTypePrstg.ListIndex)
    RsSavRec.Fields("UserID").value = val(Me.DCboUserName.BoundText)
    RsSavRec.Fields("ProjectID").value = Me.DcbProject.BoundText
    RsSavRec.Fields("SquareID").value = val(DcbSquare.BoundText)
    RsSavRec.Fields("path_photo").value = Txt_path_photo.Text
    
    'RsSavRec.Fields("ModelID").value = val(DcbModel.BoundText)
    'RsSavRec.Fields("RoomFrom").value = val(TxtRoomFrom.Text)
    'RsSavRec.Fields("RoomTo").value = val(TxtRoomTo.Text)
    'RsSavRec.Fields("PriceFrom").value = val(TxtPriceFrom.Text)
    'RsSavRec.Fields("PriceTo").value = val(TxtPriceTo.Text)
    'RsSavRec.Fields("LanAreaFrom").value = val(TxtLanAreaFrom.Text)
    'RsSavRec.Fields("LanAreaTo").value = val(TxtLanAreaTo.Text)
    'RsSavRec.Fields("HouseAreaFrom").value = val(TxtLanAreaFrom.Text)
    'RsSavRec.Fields("HouseAreaTo").value = val(TxtHouseAreaTo.Text)
    'RsSavRec.Fields("ConstStatus").value = val(DcbConstStatus.ListIndex)
    RsSavRec.update
  
''//////////////////////////
    Set RsDevsub = New ADODB.Recordset
    StrSQL = "SELECT  *  from TblBookingBondsInvsDet Where (1 = -1)"
    RsDevsub.Open StrSQL, Cn, adOpenKeyset, adLockOptimistic, adCmdText
    
    Dim RsMost As ADODB.Recordset
    Set RsMost = New ADODB.Recordset
    
    StrSQL = "SELECT  *  from Land_Planner Where Land_Name = '" & CommonDialog1.FileTitle & "'"
    RsMost.Open StrSQL, Cn, adOpenKeyset, adLockOptimistic, adCmdText
    
    Dim i As Integer
    Dim str2 As String
    'With Me.Grid
    '   For I = .FixedRows To .Rows - 1
    '   If .TextMatrix(I, .ColIndex("BlockNo")) <> "" Then
    '   RsDevsub.AddNew
    '            RsDevsub("TypeTrans").value = 0
    '            RsDevsub("BokID").value = val(Me.TxtSerial1.Text)
    '            RsDevsub("SquareID").value = IIf((.TextMatrix(I, .ColIndex("SquareID"))) = "", Null, val(.TextMatrix(I, .ColIndex("SquareID"))))
    '            RsDevsub("BlockNo").value = IIf((.TextMatrix(I, .ColIndex("BlockNo"))) = "", Null, (.TextMatrix(I, .ColIndex("BlockNo"))))
    '            RsDevsub("PartID").value = IIf((.TextMatrix(I, .ColIndex("PartID"))) = "", Null, val(.TextMatrix(I, .ColIndex("PartID"))))
    '            RsDevsub("ModelID").value = IIf((.TextMatrix(I, .ColIndex("ModelID"))) = "", Null, val(.TextMatrix(I, .ColIndex("ModelID"))))
    '            RsDevsub("Name").value = IIf((.TextMatrix(I, .ColIndex("Name"))) = "", Null, (.TextMatrix(I, .ColIndex("Name"))))
    '            RsDevsub("BedroomsNo").value = IIf((.TextMatrix(I, .ColIndex("BedroomsNo"))) = "", Null, val(.TextMatrix(I, .ColIndex("BedroomsNo"))))
    '            RsDevsub("ConstStatus").value = IIf((.TextMatrix(I, .ColIndex("ConstStatus"))) = "", Null, val(.TextMatrix(I, .ColIndex("ConstStatus"))))
    '            RsDevsub("MOHPrice").value = IIf((.TextMatrix(I, .ColIndex("MOHPrice"))) = "", Null, val(.TextMatrix(I, .ColIndex("MOHPrice"))))
    '            RsDevsub("LandArea").value = IIf((.TextMatrix(I, .ColIndex("LandArea"))) = "", Null, val(.TextMatrix(I, .ColIndex("LandArea"))))
    '            RsDevsub("Remarks").value = IIf((.TextMatrix(I, .ColIndex("Remarks"))) = "", Null, (.TextMatrix(I, .ColIndex("Remarks"))))
    '            RsDevsub("HouseArea").value = IIf((.TextMatrix(I, .ColIndex("HouseArea"))) = "", Null, val(.TextMatrix(I, .ColIndex("HouseArea"))))
    '            RsDevsub("Total").value = IIf((.TextMatrix(I, .ColIndex("Total"))) = "", 0, .TextMatrix(I, .ColIndex("Total")))
    '            RsDevsub("AddValue").value = IIf((.TextMatrix(I, .ColIndex("AddValue"))) = "", 0, .TextMatrix(I, .ColIndex("AddValue")))
    '            RsDevsub("ValueOffice").value = IIf((.TextMatrix(I, .ColIndex("ValueOffice"))) = "", 0, val(.TextMatrix(I, .ColIndex("ValueOffice"))))
    '   RsDevsub.update
    '  End If
    ' Next I
    'End With
'''///////////////
Dim BtIndex As String
      Set RsDevsub = New ADODB.Recordset
    StrSQL = "SELECT  *  from TblBookingBondsInvsDet Where (1 = -1)"
    RsDevsub.Open StrSQL, Cn, adOpenKeyset, adLockOptimistic, adCmdText
    
    With Me.VSFlexGrid1
       For i = .FixedRows To .Rows - 1
       If .TextMatrix(i, .ColIndex("BlockNo")) <> "" Then
       RsDevsub.AddNew
                RsDevsub("TypeTrans").value = 1
                RsDevsub("BokID").value = val(Me.TxtSerial1.Text)
                RsDevsub("SquareID").value = IIf((.TextMatrix(i, .ColIndex("SquareID"))) = "", Null, val(.TextMatrix(i, .ColIndex("SquareID"))))
                RsDevsub("BlockNo").value = IIf((.TextMatrix(i, .ColIndex("BlockNo"))) = "", Null, (.TextMatrix(i, .ColIndex("BlockNo"))))
                RsDevsub("PartID").value = IIf((.TextMatrix(i, .ColIndex("PartID"))) = "", Null, val(.TextMatrix(i, .ColIndex("PartID"))))
                RsDevsub("ModelID").value = IIf((.TextMatrix(i, .ColIndex("ModelID"))) = "", Null, val(.TextMatrix(i, .ColIndex("ModelID"))))
                RsDevsub("Name").value = IIf((.TextMatrix(i, .ColIndex("Name"))) = "", Null, (.TextMatrix(i, .ColIndex("Name"))))
                RsDevsub("BedroomsNo").value = IIf((.TextMatrix(i, .ColIndex("BedroomsNo"))) = "", Null, val(.TextMatrix(i, .ColIndex("BedroomsNo"))))
                RsDevsub("ConstStatus").value = IIf((.TextMatrix(i, .ColIndex("ConstStatus"))) = "", Null, val(.TextMatrix(i, .ColIndex("ConstStatus"))))
                RsDevsub("MOHPrice").value = IIf((.TextMatrix(i, .ColIndex("MOHPrice"))) = "", Null, val(.TextMatrix(i, .ColIndex("MOHPrice"))))
                RsDevsub("LandArea").value = IIf((.TextMatrix(i, .ColIndex("LandArea"))) = "", Null, val(.TextMatrix(i, .ColIndex("LandArea"))))
                RsDevsub("Remarks").value = IIf((.TextMatrix(i, .ColIndex("Remarks"))) = "", Null, (.TextMatrix(i, .ColIndex("Remarks"))))
                RsDevsub("HouseArea").value = IIf((.TextMatrix(i, .ColIndex("HouseArea"))) = "", Null, val(.TextMatrix(i, .ColIndex("HouseArea"))))
                RsDevsub("Total").value = IIf((.TextMatrix(i, .ColIndex("Total"))) = "", 0, .TextMatrix(i, .ColIndex("Total")))
                RsDevsub("AddValue").value = IIf((.TextMatrix(i, .ColIndex("AddValue"))) = "", 0, .TextMatrix(i, .ColIndex("AddValue")))
                RsDevsub("ValueOffice").value = IIf((.TextMatrix(i, .ColIndex("ValueOffice"))) = "", 0, val(.TextMatrix(i, .ColIndex("ValueOffice"))))
                RsDevsub("L_X").value = 417
                RsDevsub("L_Y").value = 167
                RsDevsub("S_Width").value = 75
                RsDevsub("S_Height").value = 23
                RsDevsub("B_Color").value = "color" ' IIf((.TextMatrix(i, .ColIndex("B_Color"))) = "", color, val(.TextMatrix(i, .ColIndex("B_Color"))))
                RsDevsub("B_Type").value = "B"
                RsDevsub("Status").value = 0
                If RsMost.RecordCount = 1 Then
                    RsDevsub("Pic_NameID").value = RsMost("Land_ID").value  ' IIf((.TextMatrix(i, .ColIndex("Pic_NameID"))) = "", 0, val(.TextMatrix(i, .ColIndex("Pic_NameID"))))3.
                End If
                
               
                
       RsDevsub.update
       BtIndex = "Button" & RsDevsub("ID").value
       Cn.Execute "Update TblBookingBondsInvsDet set B_Index ='" & BtIndex & "' where id =" & RsDevsub("ID").value & ""
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
    TxtSerial1.Text = IIf(IsNull(RsSavRec.Fields("ID").value), "", RsSavRec.Fields("ID").value)
    XPDtbTrans.value = IIf(IsNull(RsSavRec.Fields("RecordDate").value), Date, RsSavRec.Fields("RecordDate").value)
    Dcbranch.BoundText = IIf(IsNull(RsSavRec.Fields("BrnchID").value), "", RsSavRec.Fields("BrnchID").value)
    Me.TxtNameP.Text = IIf(IsNull(RsSavRec.Fields("Name").value), "", RsSavRec.Fields("Name").value)
    Me.TxtTelephone.Text = IIf(IsNull(RsSavRec.Fields("Telephone").value), "", RsSavRec.Fields("Telephone").value)
    Me.DcbTypePrstg.ListIndex = IIf(IsNull(RsSavRec.Fields("PaymentID").value), -1, RsSavRec.Fields("PaymentID").value)
    Me.DcbBank.BoundText = IIf(IsNull(RsSavRec.Fields("BanckID").value), 0, RsSavRec.Fields("BanckID").value)  ': ProgressBar1.value = 90
    Me.DcbProject.BoundText = IIf(IsNull(RsSavRec.Fields("ProjectID").value), "", RsSavRec.Fields("ProjectID").value)
    Me.DCboUserName.BoundText = IIf(IsNull(RsSavRec.Fields("UserID").value), 0, RsSavRec.Fields("UserID").value)  ': ProgressBar1.value = 10
    Me.DcbSquare.BoundText = IIf(IsNull(RsSavRec.Fields("SquareID").value), "", RsSavRec.Fields("SquareID").value)
    Me.DcbModel.BoundText = IIf(IsNull(RsSavRec.Fields("ModelID").value), "", RsSavRec.Fields("ModelID").value)
    TxtRoomFrom.Text = IIf(IsNull(RsSavRec.Fields("RoomFrom").value), "", RsSavRec.Fields("RoomFrom").value)
    TxtRoomTo.Text = IIf(IsNull(RsSavRec.Fields("RoomTo").value), "", RsSavRec.Fields("RoomTo").value)
    TxtPriceFrom.Text = IIf(IsNull(RsSavRec.Fields("PriceFrom").value), "", RsSavRec.Fields("PriceFrom").value)
    TxtPriceTo.Text = IIf(IsNull(RsSavRec.Fields("PriceTo").value), "", RsSavRec.Fields("PriceTo").value)
    TxtLanAreaFrom.Text = IIf(IsNull(RsSavRec.Fields("LanAreaFrom").value), "", RsSavRec.Fields("LanAreaFrom").value)
    TxtLanAreaTo.Text = IIf(IsNull(RsSavRec.Fields("LanAreaTo").value), "", RsSavRec.Fields("LanAreaTo").value)
    TxtHouseAreaFrom.Text = IIf(IsNull(RsSavRec.Fields("HouseAreaFrom").value), "", RsSavRec.Fields("HouseAreaFrom").value)
    TxtHouseAreaTo.Text = IIf(IsNull(RsSavRec.Fields("HouseAreaTo").value), "", RsSavRec.Fields("HouseAreaTo").value)
    Txt_path_photo.Text = IIf(IsNull(RsSavRec.Fields("path_photo").value), "", RsSavRec.Fields("path_photo").value)
   ' Txt_Path_General_photo.Text = IIf(IsNull(RsSavRec.Fields("Path_General_photo").value), "", RsSavRec.Fields("Path_General_photo").value)
    
    'RsSavRec.Fields("path_photo").value = Txt_path_photo.Text
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

If TxtNameP.Text = "" Then
If SystemOptions.UserInterface = ArabicInterface Then
MsgBox "Ì—ÃÏ «œŒ«·  «”„ «·Õ«Ã“ "
Else
MsgBox "Please Enter Name"
End If
TxtNameP.SetFocus
Exit Sub
End If
     If SystemOptions.UserInterface = ArabicInterface Then
     If Me.DcbProject.Text = "" Or val(DcbProject.BoundText) = 0 Then
     MsgBox "Ì—ÃÏ «œŒ«· «”„ «·„‘—Ê⁄"
     DcbProject.SetFocus
     Exit Sub
     End If
     Else
    If Me.DcbProject.Text = "" Or val(DcbProject.BoundText) = 0 Then
     MsgBox "Please Enter Name"
     DcbProject.SetFocus
     Exit Sub
     End If
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
    StrRecID = new_id("TblBookingBondsInvs", "ID", "")
    Me.TxtSerial1.Text = StrRecID
    RsSavRec.AddNew
    RsSavRec.Fields("ID").value = IIf(StrRecID <> "", StrRecID, Null)
    FiLLRec
ErrTrap:
End Sub
Sub RetriveProjectData(Optional ProjectCode As String)
Dim rs2 As ADODB.Recordset
Set rs2 = New ADODB.Recordset
Dim sql As String
sql = "select * from  TblProjecInvestment where  (ProjectCode = N'" & ProjectCode & "')"
rs2.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
If rs2.RecordCount > 0 Then
Txt_path_photo.Text = IIf(IsNull(rs2("path_photo").value), "", rs2("path_photo").value)
Txt_Path_General_photo.Text = IIf(IsNull(rs2("Path_General_photo").value), "", rs2("Path_General_photo").value)


Else
Txt_path_photo.Text = ""
Txt_Path_General_photo.Text = ""
End If







        Dim Str_Path As String
            Str_Path = Txt_Path_General_photo
 DBPix201.ImageClear
 If Str_Path = "" Then Exit Sub
        If Dir(Str_Path) <> "" Then
                DBPix201.ImageLoadFile Str_Path
                Else
   DBPix201.ImageClear

End If

End Sub
 Sub FullGridData()
 On Error GoTo ErrTrap
  Dim Rs1 As ADODB.Recordset
  Dim sql As String
    Grid.Clear flexClearScrollable, flexClearEverything
    Grid.Rows = 1
sql = " SELECT     dbo.TblBookingBondsInvsDet.ID, dbo.TblBookingBondsInvsDet.BokID, dbo.TblBookingBondsInvsDet.TypeTrans, dbo.TblBookingBondsInvsDet.SquareID, "
sql = sql + "                      dbo.TblBookingBondsInvsDet.PartID, dbo.TblProjecInvestmentDet.PartNo, dbo.TblBookingBondsInvsDet.Name AS ModelName,"
sql = sql + "                      dbo.TblBookingBondsInvsDet.ModelID, dbo.TblModelInves.Name, dbo.TblModelInves.NameE, dbo.TblBookingBondsInvsDet.BedroomsNo,"
sql = sql + "                      dbo.TblBookingBondsInvsDet.ConstStatus, dbo.TblBookingBondsInvsDet.MOHPrice, dbo.TblBookingBondsInvsDet.AddValue, dbo.TblBookingBondsInvsDet.LandArea,"
sql = sql + "                      dbo.TblBookingBondsInvsDet.HouseArea, dbo.TblBookingBondsInvsDet.ValueOffice, dbo.TblBookingBondsInvsDet.Total, dbo.TblBookingBondsInvsDet.Remarks,"
sql = sql + "                      dbo.TblBookingBondsInvsDet.BlockNo , dbo.TblProjecInvestmentDet.PartCode, dbo.TblProjecInvestment.Square, dbo.TblProjecInvestment.SquareCode"
sql = sql + " FROM         dbo.TblProjecInvestment INNER JOIN"
sql = sql + "                      dbo.TblProjecInvestmentDet ON dbo.TblProjecInvestment.ID = dbo.TblProjecInvestmentDet.ProjInvID RIGHT OUTER JOIN"
sql = sql + "                      dbo.TblBookingBondsInvsDet LEFT OUTER JOIN"
sql = sql + "                      dbo.TblModelInves ON dbo.TblBookingBondsInvsDet.ModelID = dbo.TblModelInves.ID ON"
sql = sql + "                      dbo.TblProjecInvestmentDet.ID = dbo.TblBookingBondsInvsDet.PartID"
sql = sql + " Where (dbo.TblBookingBondsInvsDet.BokID = " & val(TxtSerial1.Text) & ") And (dbo.TblBookingBondsInvsDet.TypeTrans = 0)"
Set Rs1 = New ADODB.Recordset
  
  Rs1.Open sql, Cn, adOpenKeyset, adLockOptimistic, adCmdText
    If Rs1.RecordCount > 0 Then
     Rs1.MoveFirst
     End If
     Dim i As Integer
     With Grid
                    For i = .FixedRows To Rs1.RecordCount
                   .Rows = .FixedRows + Rs1.RecordCount
             If SystemOptions.UserInterface = ArabicInterface Then
                   .TextMatrix(i, .ColIndex("Approve")) = "«÷€ÿ ··≈⁄ „«œ "
             Else
                   .TextMatrix(i, .ColIndex("Approve")) = "Approve"
             End If
                   .TextMatrix(i, .ColIndex("Ser")) = i
                   .TextMatrix(i, .ColIndex("Total")) = IIf(IsNull(Rs1("Total").value), 0, Rs1("Total").value)
                   .TextMatrix(i, .ColIndex("AddValue")) = IIf(IsNull(Rs1("AddValue").value), 0, Rs1("AddValue").value)
                   .TextMatrix(i, .ColIndex("ValueOffice")) = IIf(IsNull(Rs1("ValueOffice").value), 0, Rs1("ValueOffice").value)
                   .TextMatrix(i, .ColIndex("PartCode")) = IIf(IsNull(Rs1("PartCode").value), "", Rs1("PartCode").value)
                   .TextMatrix(i, .ColIndex("SquareID")) = IIf(IsNull(Rs1("SquareID").value), "", Rs1("SquareID").value)
                   .TextMatrix(i, .ColIndex("PartID")) = IIf(IsNull(Rs1("PartID").value), "", Rs1("PartID").value)
                   .TextMatrix(i, .ColIndex("BlockNo")) = IIf(IsNull(Rs1("BlockNo").value), "", Rs1("BlockNo").value)
                   .TextMatrix(i, .ColIndex("PartNo")) = IIf(IsNull(Rs1("PartNo").value), "", Rs1("PartNo").value)
                   .TextMatrix(i, .ColIndex("ModelID")) = IIf(IsNull(Rs1("ModelID").value), 0, Rs1("ModelID").value)
                   .TextMatrix(i, .ColIndex("LandArea")) = IIf(IsNull(Rs1("LandArea").value), 0, Rs1("LandArea").value)
                   .TextMatrix(i, .ColIndex("HouseArea")) = IIf(IsNull(Rs1("HouseArea").value), 0, Rs1("HouseArea").value)
                   .TextMatrix(i, .ColIndex("SquareID")) = IIf(IsNull(Rs1("SquareID").value), "", Rs1("SquareID").value)
                   .TextMatrix(i, .ColIndex("SquareNo")) = IIf(IsNull(Rs1("Square").value), "", Rs1("Square").value)
                   .TextMatrix(i, .ColIndex("MOHPrice")) = IIf(IsNull(Rs1("MOHPrice").value), 0, Rs1("MOHPrice").value)
                   .TextMatrix(i, .ColIndex("BedroomsNo")) = IIf(IsNull(Rs1("BedroomsNo").value), "", Rs1("BedroomsNo").value)
                   .TextMatrix(i, .ColIndex("ConstStatus")) = IIf(IsNull(Rs1("ConstStatus").value), "", Rs1("ConstStatus").value)
                 '  .TextMatrix(i, .ColIndex("StatusID")) = IIf(IsNull(Rs1("StatusID").value), "", Rs1("StatusID").value)
                   .TextMatrix(i, .ColIndex("Remarks")) = IIf(IsNull(Rs1("Remarks").value), "", Rs1("Remarks").value)
                   If SystemOptions.UserInterface = ArabicInterface Then
                   .TextMatrix(i, .ColIndex("Name")) = IIf(IsNull(Rs1("Name").value), IIf(IsNull(Rs1("ModelName").value), "", Rs1("ModelName").value), Rs1("Name").value)
                   Else
                   .TextMatrix(i, .ColIndex("Name")) = IIf(IsNull(Rs1("NameE").value), IIf(IsNull(Rs1("ModelName").value), "", Rs1("ModelName").value), Rs1("NameE").value)
                   End If
                   
                   
                   
                   Rs1.MoveNext
             Next i
        End With
 ''/////////////////////////
     VSFlexGrid1.Clear flexClearScrollable, flexClearEverything
    VSFlexGrid1.Rows = 1
sql = " SELECT     dbo.TblBookingBondsInvsDet.ID, dbo.TblBookingBondsInvsDet.BokID, dbo.TblBookingBondsInvsDet.TypeTrans, dbo.TblBookingBondsInvsDet.SquareID, "
sql = sql + "                      dbo.TblBookingBondsInvsDet.PartID, dbo.TblProjecInvestmentDet.PartNo, dbo.TblBookingBondsInvsDet.Name AS ModelName,"
sql = sql + "                      dbo.TblBookingBondsInvsDet.ModelID, dbo.TblModelInves.Name, dbo.TblModelInves.NameE, dbo.TblBookingBondsInvsDet.BedroomsNo,"
sql = sql + "                      dbo.TblBookingBondsInvsDet.ConstStatus, dbo.TblBookingBondsInvsDet.MOHPrice, dbo.TblBookingBondsInvsDet.AddValue, dbo.TblBookingBondsInvsDet.LandArea,"
sql = sql + "                      dbo.TblBookingBondsInvsDet.HouseArea, dbo.TblBookingBondsInvsDet.ValueOffice, dbo.TblBookingBondsInvsDet.Total, dbo.TblBookingBondsInvsDet.Remarks,"
sql = sql + "                      dbo.TblBookingBondsInvsDet.BlockNo , dbo.TblProjecInvestmentDet.PartCode, dbo.TblProjecInvestment.Square, dbo.TblProjecInvestment.SquareCode"
sql = sql + " FROM         dbo.TblProjecInvestment INNER JOIN"
sql = sql + "                      dbo.TblProjecInvestmentDet ON dbo.TblProjecInvestment.ID = dbo.TblProjecInvestmentDet.ProjInvID RIGHT OUTER JOIN"
sql = sql + "                      dbo.TblBookingBondsInvsDet LEFT OUTER JOIN"
sql = sql + "                      dbo.TblModelInves ON dbo.TblBookingBondsInvsDet.ModelID = dbo.TblModelInves.ID ON"
sql = sql + "                      dbo.TblProjecInvestmentDet.ID = dbo.TblBookingBondsInvsDet.PartID"
sql = sql + " Where (dbo.TblBookingBondsInvsDet.BokID = " & val(TxtSerial1.Text) & ") And (dbo.TblBookingBondsInvsDet.TypeTrans = 1)"
Set Rs1 = New ADODB.Recordset
  
  Rs1.Open sql, Cn, adOpenKeyset, adLockOptimistic, adCmdText
    If Rs1.RecordCount > 0 Then
     Rs1.MoveFirst
     End If
    
     With VSFlexGrid1
                    For i = .FixedRows To Rs1.RecordCount
                   .Rows = .FixedRows + Rs1.RecordCount
                   .TextMatrix(i, .ColIndex("Ser")) = i
                   .TextMatrix(i, .ColIndex("Total")) = IIf(IsNull(Rs1("Total").value), 0, Rs1("Total").value)
                   .TextMatrix(i, .ColIndex("AddValue")) = IIf(IsNull(Rs1("AddValue").value), 0, Rs1("AddValue").value)
                   .TextMatrix(i, .ColIndex("ValueOffice")) = IIf(IsNull(Rs1("ValueOffice").value), 0, Rs1("ValueOffice").value)
                   .TextMatrix(i, .ColIndex("PartCode")) = IIf(IsNull(Rs1("PartCode").value), "", Rs1("PartCode").value)
                   .TextMatrix(i, .ColIndex("SquareID")) = IIf(IsNull(Rs1("SquareID").value), "", Rs1("SquareID").value)
                   .TextMatrix(i, .ColIndex("PartID")) = IIf(IsNull(Rs1("PartID").value), "", Rs1("PartID").value)
                   .TextMatrix(i, .ColIndex("BlockNo")) = IIf(IsNull(Rs1("BlockNo").value), "", Rs1("BlockNo").value)
                   .TextMatrix(i, .ColIndex("PartNo")) = IIf(IsNull(Rs1("PartNo").value), "", Rs1("PartNo").value)
                   .TextMatrix(i, .ColIndex("ModelID")) = IIf(IsNull(Rs1("ModelID").value), 0, Rs1("ModelID").value)
                   .TextMatrix(i, .ColIndex("LandArea")) = IIf(IsNull(Rs1("LandArea").value), 0, Rs1("LandArea").value)
                   .TextMatrix(i, .ColIndex("HouseArea")) = IIf(IsNull(Rs1("HouseArea").value), 0, Rs1("HouseArea").value)
                   .TextMatrix(i, .ColIndex("SquareID")) = IIf(IsNull(Rs1("SquareID").value), "", Rs1("SquareID").value)
                   .TextMatrix(i, .ColIndex("SquareNo")) = IIf(IsNull(Rs1("Square").value), "", Rs1("Square").value)
                   .TextMatrix(i, .ColIndex("MOHPrice")) = IIf(IsNull(Rs1("MOHPrice").value), 0, Rs1("MOHPrice").value)
                   .TextMatrix(i, .ColIndex("BedroomsNo")) = IIf(IsNull(Rs1("BedroomsNo").value), "", Rs1("BedroomsNo").value)
                   .TextMatrix(i, .ColIndex("ConstStatus")) = IIf(IsNull(Rs1("ConstStatus").value), "", Rs1("ConstStatus").value)
                 '  .TextMatrix(i, .ColIndex("StatusID")) = IIf(IsNull(Rs1("StatusID").value), "", Rs1("StatusID").value)
                   .TextMatrix(i, .ColIndex("Remarks")) = IIf(IsNull(Rs1("Remarks").value), "", Rs1("Remarks").value)
                   If SystemOptions.UserInterface = ArabicInterface Then
                   .TextMatrix(i, .ColIndex("Name")) = IIf(IsNull(Rs1("Name").value), IIf(IsNull(Rs1("ModelName").value), "", Rs1("ModelName").value), Rs1("Name").value)
                   Else
                   .TextMatrix(i, .ColIndex("Name")) = IIf(IsNull(Rs1("NameE").value), IIf(IsNull(Rs1("ModelName").value), "", Rs1("ModelName").value), Rs1("NameE").value)
                   End If
    
    
          '   Dim Str_Path As String
          '   Str_Path = App.path & "\" & SystemOptions.ImagesPath & "\" & IIf(IsNull(Rs1("PartID").value), "", Rs1("PartID").value) & ".JPG"
    
          '   If Dir(Str_Path) <> "" Then
          '      DBPix201.ImageLoadFile (App.path & "\" & SystemOptions.ImagesPath & "\" & IIf(IsNull(Rs1("PartID").value), "", Rs1("PartID").value) & ".JPG")
          '    Else
        '        Str_Path = App.path & "\" & SystemOptions.ImagesPath & "\DefaultCar.JPG"
                'If Dir(Str_Path) <> "" Then
                '        DBPix201.ImageLoadFile (App.path & "\" & SystemOptions.ImagesPath & "\DefualtRealState.JPG")
                'Else
                '        DBPix201.ImageClear
                'End If
          '  End If
                   Rs1.MoveNext
             Next i
        End With
        Exit Sub
ErrTrap:
    End Sub


Function CountInBooking() As Double
Dim i As Integer
Dim Cout As Double
Cout = 0
With VSFlexGrid1
For i = 1 To .Rows - 1
If val(.TextMatrix(i, .ColIndex("PartID"))) <> 0 Then
Cout = Cout + 1
End If
Next i
End With
CountInBooking = Cout
End Function
Sub AddUnit(Optional Row As Long)
Dim i As Integer
Dim k As Integer
With VSFlexGrid1
If .Rows = 1 Then
k = .Rows
Else
k = .Rows - 1
End If
.Rows = .Rows + 1
i = k

.TextMatrix(i, .ColIndex("SquareID")) = Grid.TextMatrix(Row, Grid.ColIndex("SquareID"))
.TextMatrix(i, .ColIndex("SquareNo")) = Grid.TextMatrix(Row, Grid.ColIndex("SquareNo"))
.TextMatrix(i, .ColIndex("BlockNo")) = Grid.TextMatrix(Row, Grid.ColIndex("BlockNo"))
.TextMatrix(i, .ColIndex("PartID")) = Grid.TextMatrix(Row, Grid.ColIndex("PartID"))
.TextMatrix(i, .ColIndex("PartCode")) = Grid.TextMatrix(Row, Grid.ColIndex("PartCode"))
.TextMatrix(i, .ColIndex("PartNo")) = Grid.TextMatrix(Row, Grid.ColIndex("PartNo"))
.TextMatrix(i, .ColIndex("Name")) = Grid.TextMatrix(Row, Grid.ColIndex("Name"))
.TextMatrix(i, .ColIndex("ModelID")) = Grid.TextMatrix(Row, Grid.ColIndex("ModelID"))
.TextMatrix(i, .ColIndex("LandArea")) = Grid.TextMatrix(Row, Grid.ColIndex("LandArea"))
.TextMatrix(i, .ColIndex("HouseArea")) = Grid.TextMatrix(Row, Grid.ColIndex("HouseArea"))
.TextMatrix(i, .ColIndex("MOHPrice")) = Grid.TextMatrix(Row, Grid.ColIndex("MOHPrice"))
.TextMatrix(i, .ColIndex("BedroomsNo")) = Grid.TextMatrix(Row, Grid.ColIndex("BedroomsNo"))
.TextMatrix(i, .ColIndex("ConstStatus")) = Grid.TextMatrix(Row, Grid.ColIndex("ConstStatus"))
.TextMatrix(i, .ColIndex("AddValue")) = Grid.TextMatrix(Row, Grid.ColIndex("AddValue"))
.TextMatrix(i, .ColIndex("ValueOffice")) = Grid.TextMatrix(Row, Grid.ColIndex("ValueOffice"))
.TextMatrix(i, .ColIndex("Total")) = Grid.TextMatrix(Row, Grid.ColIndex("Total"))
.TextMatrix(i, .ColIndex("Remarks")) = Grid.TextMatrix(Row, Grid.ColIndex("Remarks"))
.TextMatrix(i, .ColIndex("Ser")) = i

DBPix201.ImageSaveFile (App.path & "\" & SystemOptions.ImagesPath & "\" & Grid.TextMatrix(Row, Grid.ColIndex("PartID")) & ".JPG")

 DBPix201.ImageLoadFile (App.path & "\" & SystemOptions.ImagesPath & "\DefualtRealState.JPG")
End With
End Sub
Private Sub Grid_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
With Grid
If .ColKey(Col) <> "Approve" Then
Cancel = True
End If
End With
End Sub


Private Sub RemoveGridRow()

    With Me.VSFlexGrid1

        If .Row <= 0 Then Exit Sub
        .RemoveItem .Row
    End With

    ReLineGrid
End Sub
Function GetNoBooking(Optional PartID As Double) As Double
Dim sql As String
Dim Rs3 As ADODB.Recordset
Set Rs3 = New ADODB.Recordset
sql = " SELECT     COUNT(PartID) AS CountPartID"
sql = sql & " From dbo.TblBookingBondsInvsDet"
sql = sql & " Where (TypeTrans = 1) And (BokID <> " & val(TxtSerial1.Text) & ") And (PartID = " & PartID & ")"
Rs3.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
If Rs3.RecordCount > 0 Then
GetNoBooking = IIf(IsNull(Rs3("CountPartID").value), 0, Rs3("CountPartID").value)
Else
GetNoBooking = 0
End If
End Function
Sub FillGrid()
Dim k As Integer
Dim i As Integer
Dim PartID As Double
Dim NoBooking As Double
Dim NoBookingInGrid As Double

Dim sql As String
Dim rs2 As ADODB.Recordset
Set rs2 = New ADODB.Recordset
sql = " SELECT     dbo.TblProjecInvestment.ID, dbo.TblProjecInvestment.Square, dbo.TblProjecInvestment.SquareCode, dbo.TblProjecInvestment.BrnchID, "
sql = sql & "                      dbo.TblBranchesData.branch_name, dbo.TblBranchesData.branch_namee, dbo.TblProjecInvestment.RecordDate, dbo.TblProjecInvestment.Name,"
sql = sql & "                      dbo.TblProjecInvestment.NameE, dbo.TblProjecInvestment.TypePrstg, dbo.TblProjecInvestmentDet.ID AS IDDet, dbo.TblProjecInvestmentDet.BlockNo,"
sql = sql & "                      dbo.TblProjecInvestmentDet.PartNo, dbo.TblProjecInvestmentDet.ModelName, dbo.TblProjecInvestmentDet.LandArea, dbo.TblProjecInvestmentDet.HouseArea,"
sql = sql & "                      dbo.TblProjecInvestmentDet.MarketPrice, dbo.TblProjecInvestmentDet.MOHPrice, dbo.TblProjecInvestmentDet.BedroomsNo, dbo.TblProjecInvestmentDet.ConstStatus,"
sql = sql & "                       dbo.TblProjecInvestmentDet.StatusID, dbo.TblProjecInvestmentDet.Remarks, dbo.TblProjecInvestmentDet.DeveloperCode, dbo.TblProjecInvestmentDet.PartCode,"
sql = sql & "                      dbo.TblProjecInvestmentDet.ValueOffice, dbo.TblProjecInvestment.ProjectCode, dbo.TblProjecInvestmentDet.ModelID, dbo.TblModelInves.Name AS ModelName1,"
sql = sql & "                      dbo.TblModelInves.NameE AS ModelName1E ,dbo.TblModelInves.NoBooking , dbo.TblProjecInvestmentDet.AddValue, dbo.TblProjecInvestmentDet.Total"
sql = sql & " FROM         dbo.TblModelInves RIGHT OUTER JOIN"
sql = sql & "                      dbo.TblProjecInvestmentDet ON dbo.TblModelInves.ID = dbo.TblProjecInvestmentDet.ModelID RIGHT OUTER JOIN"
sql = sql & "                      dbo.TblProjecInvestment ON dbo.TblProjecInvestmentDet.ProjInvID = dbo.TblProjecInvestment.ID LEFT OUTER JOIN"
sql = sql & "                      dbo.TblBranchesData ON dbo.TblProjecInvestment.BrnchID = dbo.TblBranchesData.branch_id"
sql = sql & " WHERE     (dbo.TblProjecInvestment.ProjectCode = N'" & DcbProject.BoundText & "')"
If val(DcbSquare.BoundText) <> 0 And DcbSquare.Text <> "" Then
sql = sql & " and dbo.TblProjecInvestment.ID=" & val(DcbSquare.BoundText) & ""
End If
If TxtPartCode.Text <> "" Then
sql = sql & " and  dbo.TblProjecInvestmentDet.PartCode like N'%" & TxtPartCode.Text & "%'"
'(NameE LIKE N'%" & Name & "%')"
End If

If val(DcbModel.BoundText) <> 0 And DcbModel.Text <> "" Then
sql = sql & " and  dbo.TblProjecInvestmentDet.ModelID=" & val(DcbModel.BoundText) & ""
End If
If val(DcbConstStatus.ListIndex) <> -1 And DcbConstStatus.Text <> "" Then
sql = sql & " and  dbo.TblProjecInvestmentDet.ConstStatus=" & val(DcbConstStatus.ListIndex) + 1 & ""
End If
If val(TxtLanAreaFrom.Text) <> 0 Then
sql = sql & " and dbo.TblProjecInvestmentDet.LandArea >=" & val(TxtLanAreaFrom.Text) & ""
End If
If val(TxtLanAreaTo.Text) <> 0 Then
sql = sql & " and dbo.TblProjecInvestmentDet.LandArea <=" & val(TxtLanAreaTo.Text) & ""
End If
If val(TxtHouseAreaFrom.Text) <> 0 Then
sql = sql & " and dbo.TblProjecInvestmentDet.HouseArea >=" & val(TxtHouseAreaFrom.Text) & ""
End If
If val(TxtHouseAreaTo.Text) <> 0 Then
sql = sql & " and dbo.TblProjecInvestmentDet.HouseArea <=" & val(TxtHouseAreaTo.Text) & ""
End If
If val(TxtRoomFrom.Text) <> 0 Then
sql = sql & " and dbo.TblProjecInvestmentDet.BedroomsNo >=" & val(TxtRoomFrom.Text) & ""
End If
If val(TxtRoomTo.Text) <> 0 Then
sql = sql & " and dbo.TblProjecInvestmentDet.BedroomsNo <=" & val(TxtRoomTo.Text) & ""
End If
If val(TxtPriceFrom.Text) <> 0 Then
sql = sql & " and  dbo.TblProjecInvestmentDet.Total >=" & val(TxtPriceFrom.Text) & ""
End If
If val(TxtPriceTo.Text) <> 0 Then
sql = sql & " and dbo.TblProjecInvestmentDet.Total <=" & val(TxtPriceTo.Text) & ""
End If
Set rs2 = New ADODB.Recordset
 Grid.Clear flexClearScrollable, flexClearEverything
            Grid.Rows = 2
rs2.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
If rs2.RecordCount > 0 Then
With Grid
.Rows = .Rows + rs2.RecordCount
rs2.MoveFirst
i = 0
For k = 1 To .Rows - 2
PartID = IIf(IsNull(rs2("IDDet").value), 0, rs2("IDDet").value)
NoBooking = IIf(IsNull(rs2("NoBooking").value), 0, rs2("NoBooking").value)
NoBookingInGrid = NoBookingUnitInGrid(PartID) + GetNoBooking(PartID)
If NoBooking > NoBookingInGrid Then
i = i + 1
       If SystemOptions.UserInterface = ArabicInterface Then
             .TextMatrix(i, .ColIndex("Approve")) = " ÕœÌœ "
             Else
             .TextMatrix(i, .ColIndex("Approve")) = "Select"
             End If
.TextMatrix(i, .ColIndex("Ser")) = i
.TextMatrix(i, .ColIndex("Squareno")) = IIf(IsNull(rs2("Square").value), "", rs2("Square").value)
.TextMatrix(i, .ColIndex("PartID")) = PartID
.TextMatrix(i, .ColIndex("SquareID")) = IIf(IsNull(rs2("ID").value), 0, rs2("ID").value)
.TextMatrix(i, .ColIndex("BlockNo")) = IIf(IsNull(rs2("BlockNo").value), "", rs2("BlockNo").value)
.TextMatrix(i, .ColIndex("PartCode")) = IIf(IsNull(rs2("PartCode").value), "", rs2("PartCode").value)
.TextMatrix(i, .ColIndex("PartNo")) = IIf(IsNull(rs2("PartNo").value), "", rs2("PartNo").value)
.TextMatrix(i, .ColIndex("ModelID")) = IIf(IsNull(rs2("ModelID").value), 0, rs2("ModelID").value)
.TextMatrix(i, .ColIndex("LandArea")) = IIf(IsNull(rs2("LandArea").value), "", rs2("LandArea").value)
.TextMatrix(i, .ColIndex("HouseArea")) = IIf(IsNull(rs2("HouseArea").value), "", rs2("HouseArea").value)
.TextMatrix(i, .ColIndex("DeveloperCode")) = IIf(IsNull(rs2("DeveloperCode").value), "", rs2("DeveloperCode").value)
.TextMatrix(i, .ColIndex("MarketPrice")) = IIf(IsNull(rs2("MarketPrice").value), "", rs2("MarketPrice").value)
.TextMatrix(i, .ColIndex("MOHPrice")) = IIf(IsNull(rs2("MOHPrice").value), "", rs2("MOHPrice").value)
.TextMatrix(i, .ColIndex("BedroomsNo")) = IIf(IsNull(rs2("BedroomsNo").value), "", rs2("BedroomsNo").value)
.TextMatrix(i, .ColIndex("ConstStatus")) = IIf(IsNull(rs2("ConstStatus").value), "", rs2("ConstStatus").value)
.TextMatrix(i, .ColIndex("AddValue")) = IIf(IsNull(rs2("AddValue").value), "", rs2("AddValue").value)

.TextMatrix(i, .ColIndex("ValueOffice")) = IIf(IsNull(rs2("ValueOffice").value), "", rs2("ValueOffice").value)
.TextMatrix(i, .ColIndex("Total")) = IIf(IsNull(rs2("Total").value), "", rs2("Total").value)
.TextMatrix(i, .ColIndex("StatusID")) = IIf(IsNull(rs2("StatusID").value), "", rs2("StatusID").value)
.TextMatrix(i, .ColIndex("Remarks")) = IIf(IsNull(rs2("Remarks").value), "", rs2("Remarks").value)
If SystemOptions.UserInterface = ArabicInterface Then
.TextMatrix(i, .ColIndex("Name")) = IIf(IsNull(rs2("ModelName1").value), IIf(IsNull(rs2("ModelName").value), "", rs2("ModelName").value), rs2("ModelName1").value)
Else
.TextMatrix(i, .ColIndex("Name")) = IIf(IsNull(rs2("ModelName1E").value), IIf(IsNull(rs2("ModelName").value), "", rs2("ModelName").value), rs2("ModelName1E").value)
End If
End If

rs2.MoveNext
Next k
End With
End If
End Sub

Private Sub Grid_CellButtonClick(ByVal Row As Long, ByVal Col As Long)
With Grid
Select Case .ColKey(Col)
Case "Approve"

If Me.TxtModFlg.Text <> "R" Then
If CountInBooking() < SystemOptions.NoBooking Then
If ChekGrid(.TextMatrix(Row, .ColIndex("PartID"))) = False Then

AddUnit Row

Else
If SystemOptions.UserInterface = ArabicInterface Then
MsgBox "·«Ì„ﬂ‰ «· ﬂ—«—"
Else
MsgBox "Can not repeat"
End If
Exit Sub
End If
Else
If SystemOptions.UserInterface = ArabicInterface Then
MsgBox "·«Ì„ﬂ‰  Ã«Ê“ Õœ «·ÿ·»"
Else
MsgBox "Can not exceed request limit"
End If
Exit Sub
End If
End If
End Select
End With
End Sub
Function NoBookingUnitInGrid(Optional PartID As Double) As Double
Dim i As Integer
Dim Cout As Double
Cout = 0
With VSFlexGrid1
For i = 1 To .Rows - 1
If val(.TextMatrix(i, .ColIndex("PartID"))) = PartID Then
Cout = Cout + 1
End If
Next i
End With
NoBookingUnitInGrid = Cout
End Function

Function ChekGrid(Optional PartID As Double) As Boolean
Dim i As Integer
With VSFlexGrid1
For i = 1 To .Rows - 1
If val(.TextMatrix(i, .ColIndex("PartID"))) = PartID Then
ChekGrid = True
Exit Function
End If
Next i
End With
ChekGrid = False
End Function
Private Sub Grid_StartEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
With Grid
Select Case .ColKey(Col)
Case "Approve"
.ColComboList(.ColIndex("Approve")) = "..."
End Select
End With
End Sub

Private Sub ISButton3_Click()
If Me.TxtModFlg.Text <> "R" Then
If DcbProject.Text = "" Or DcbProject.BoundText = "" Then
If SystemOptions.UserInterface = ArabicInterface Then
MsgBox "Ì—ÃÏ «Œ Ì«— «·„‘—Ê⁄"
Else
MsgBox "Please select project"
End If
DcbProject.SetFocus
Exit Sub
End If
FillGrid
End If
End Sub

Private Sub ISButton5_Click()
    print_report
End Sub

Private Sub ISButton8_Click()
FrmSearchinvestment.inde = 30
FrmSearchinvestment.show
End Sub

Private Sub TxtHouseAreaFrom_KeyPress(KeyAscii As Integer)
KeyAscii = KeyAscii_Num(KeyAscii, Me.TxtHouseAreaFrom.Text, 0)
End Sub

Private Sub TxtHouseAreaTo_KeyPress(KeyAscii As Integer)
KeyAscii = KeyAscii_Num(KeyAscii, Me.TxtHouseAreaTo.Text, 0)
End Sub

Private Sub TxtLanAreaFrom_KeyPress(KeyAscii As Integer)
KeyAscii = KeyAscii_Num(KeyAscii, Me.TxtLanAreaFrom.Text, 0)
End Sub

Private Sub TxtLanAreaTo_KeyPress(KeyAscii As Integer)
KeyAscii = KeyAscii_Num(KeyAscii, Me.TxtLanAreaTo.Text, 0)
End Sub

Private Sub TxtPriceFrom_KeyPress(KeyAscii As Integer)
KeyAscii = KeyAscii_Num(KeyAscii, Me.TxtPriceFrom.Text, 0)
End Sub

Private Sub TxtPriceTo_KeyPress(KeyAscii As Integer)
KeyAscii = KeyAscii_Num(KeyAscii, Me.TxtPriceTo.Text, 0)
End Sub

Private Sub TxtRoomFrom_KeyPress(KeyAscii As Integer)
KeyAscii = KeyAscii_Num(KeyAscii, Me.TxtRoomFrom.Text, 1)
End Sub

Private Sub TxtRoomTo_KeyPress(KeyAscii As Integer)
KeyAscii = KeyAscii_Num(KeyAscii, Me.TxtRoomTo.Text, 1)
End Sub

Private Sub TxtSearchCode_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyReturn Then
DcbProject.BoundText = TxtSearchCode.Text
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
                RsSavRec.find "ID=" & val(TxtSerial1.Text), , adSearchForward, 1
                  StrSQL = "Delete From TblBookingBondsInvsDet Where BokID =" & val(TxtSerial1.Text) & ""
                  Cn.Execute StrSQL, , adExecuteNoRecords
                                          RsSavRec.delete
            Grid.Clear flexClearScrollable, flexClearEverything
            Grid.Rows = 1
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
        Command2.Enabled = True
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
    Command2.Enabled = False
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
   Command2.Enabled = True
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
            Grid.Rows = Grid.Rows + 1
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
    
    Dim Str_Path As String
            Str_Path = App.path & "\" & SystemOptions.ImagesPath & "\DefualtRealState.JPG"
            If Dir(Str_Path) <> "" Then
                    DBPix201.ImageLoadFile (App.path & "\" & SystemOptions.ImagesPath & "\DefualtRealState.JPG")
            Else
                    DBPix201.ImageClear
            End If
    
    Grid.Clear flexClearScrollable, flexClearEverything
      Grid.Rows = 2
      VSFlexGrid1.Clear flexClearScrollable, flexClearEverything
      VSFlexGrid1.Rows = 2
      
    Me.DCboUserName.BoundText = user_id
    Me.Dcbranch.BoundText = Current_branch
    Dcbranch.SetFocus
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


Private Sub ChangeLang()
On Error GoTo ErrTrap
   ' form name
   ''''''''''''''''''''////
   Command1.Caption = "Open Planned"
   Command2.Caption = "Planned"
       Me.Caption = "Booking  Data"
      Label1(2).Caption = "Booking  Data"
      Me.lbl(4).Caption = "ID"
      Me.lbl(2).Caption = "Date"
      lbl(7).Caption = "Branch"
      lbl(1).Caption = "Name"
      lbl(15).Caption = "Project "
   lbl(3).Caption = "Phone"
   lbl(0).Caption = "Banck"
   lbl(10).Caption = "Square"
   lbl(10).Caption = "Model"
lbl(18).Caption = "Price From"
lbl(16).Caption = "No Room From"
lbl(17).Caption = "To"
lbl(14).Caption = "To"
lbl(13).Caption = "To"
lbl(19).Caption = "To"
lbl(5).Caption = "Land Area From"
lbl(20).Caption = "House Area From"
Cmd(0).Caption = "Delete"
Cmd(1).Caption = "Delete All"
ISButton3.Caption = "Show"
lbl(21).Caption = "Part Code"
lbl(11).Caption = "Payment"
lbl(12).Caption = "Status"
lbl(6).Caption = "Model"


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
   
  With Me.Grid
  .TextMatrix(0, .ColIndex("Approve")) = "Booking"
  .TextMatrix(0, .ColIndex("SquareNo")) = "Square No"
  .TextMatrix(0, .ColIndex("Ser")) = "Serial"
  .TextMatrix(0, .ColIndex("BlockNo")) = "Block No."
  .TextMatrix(0, .ColIndex("PartNo")) = "Land Number"
  .TextMatrix(0, .ColIndex("PartCode")) = "Code"
  .TextMatrix(0, .ColIndex("Name")) = "Model"
  .TextMatrix(0, .ColIndex("LandArea")) = "Land Area"
  .TextMatrix(0, .ColIndex("HouseArea")) = " House Area/BUA"
  .TextMatrix(0, .ColIndex("PartNo")) = "Land Number"
  .TextMatrix(0, .ColIndex("PartCode")) = "Code"
  .TextMatrix(0, .ColIndex("Name")) = "Model"
  .TextMatrix(0, .ColIndex("DeveloperCode")) = "Developer Code"
  .TextMatrix(0, .ColIndex("MarketPrice")) = "Market Price"
  .TextMatrix(0, .ColIndex("MOHPrice")) = "MOH Price"
  .TextMatrix(0, .ColIndex("BedroomsNo")) = "No.Bedrooms "
  .TextMatrix(0, .ColIndex("ConstStatus")) = " Construction Status"
  .TextMatrix(0, .ColIndex("StatusID")) = "Status"
  .TextMatrix(0, .ColIndex("Remarks")) = "Comments"
  .TextMatrix(0, .ColIndex("Total")) = "Total"
  .TextMatrix(0, .ColIndex("AddValue")) = "Added Values"
  .TextMatrix(0, .ColIndex("ValueOffice")) = "Value Business Office"
  End With
    With Me.VSFlexGrid1
  .TextMatrix(0, .ColIndex("SquareNo")) = "Square No"
  .TextMatrix(0, .ColIndex("Ser")) = "Serial"
  .TextMatrix(0, .ColIndex("BlockNo")) = "Block No."
  .TextMatrix(0, .ColIndex("PartNo")) = "Land Number"
  .TextMatrix(0, .ColIndex("PartCode")) = "Code"
  .TextMatrix(0, .ColIndex("Name")) = "Model"
   .TextMatrix(0, .ColIndex("LandArea")) = "Land Area"
  .TextMatrix(0, .ColIndex("HouseArea")) = " House Area/BUA"
  .TextMatrix(0, .ColIndex("PartNo")) = "Land Number"
  .TextMatrix(0, .ColIndex("PartCode")) = "Code"
  .TextMatrix(0, .ColIndex("Name")) = "Model"
  .TextMatrix(0, .ColIndex("DeveloperCode")) = "Developer Code"
  .TextMatrix(0, .ColIndex("MarketPrice")) = "Market Price"
  .TextMatrix(0, .ColIndex("MOHPrice")) = "MOH Price"
  .TextMatrix(0, .ColIndex("BedroomsNo")) = "No.Bedrooms "
  .TextMatrix(0, .ColIndex("ConstStatus")) = " Construction Status"
   .TextMatrix(0, .ColIndex("StatusID")) = "Status"
  .TextMatrix(0, .ColIndex("Remarks")) = "Comments"
  .TextMatrix(0, .ColIndex("Total")) = "Total"
  .TextMatrix(0, .ColIndex("AddValue")) = "Added Values"
  .TextMatrix(0, .ColIndex("ValueOffice")) = "Value Business Office"
  End With
ErrTrap:
End Sub
Private Sub AddNewRecored()
   Dim My_SQL As String
   Dim rs As ADODB.Recordset
  On Error GoTo ErrTrap
    Set rs = New ADODB.Recordset
   My_SQL = "TblBookingBondsInvs"
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
Private Sub TxtSquareCode_Change()

End Sub

Private Sub TxtSquareCode_KeyPress(KeyAscii As Integer)
Dim ID As Double
If KeyAscii = vbKeyReturn Then
GetID_CodeSqureProject TxtSquareCode.Text, ID
Me.DcbSquare.BoundText = ID
End If
End Sub

Function print_report(Optional NoteSerial As String)
    
     
    Dim MySQL As String
    Dim RsData As New ADODB.Recordset
    Dim xApp As New CRAXDRT.Application
    Dim xReport As CRAXDRT.Report
    Dim CViewer As ClsReportViewer
    Dim StrReportTitle As String
    Dim StrFileName As String
    Dim Msg As String
 

MySQL = " SELECT TblBookingBondsInvsDet.ID, TblBookingBondsInvsDet.BokID, TblBookingBondsInvsDet.TypeTrans, TblBookingBondsInvsDet.SquareID, TblBookingBondsInvsDet.PartID, TblProjecInvestmentDet.PartNo, "
MySQL = MySQL & " TblBookingBondsInvsDet.Name AS ModelName, TblBookingBondsInvsDet.ModelID, TblModelInves.Name, TblModelInves.NameE, TblBookingBondsInvsDet.BedroomsNo, TblBookingBondsInvsDet.ConstStatus, "
MySQL = MySQL & " TblBookingBondsInvsDet.MOHPrice, TblBookingBondsInvsDet.AddValue, TblBookingBondsInvsDet.LandArea, TblBookingBondsInvsDet.HouseArea, TblBookingBondsInvsDet.ValueOffice, TblBookingBondsInvsDet.Total, "
MySQL = MySQL & " TblBookingBondsInvsDet.Remarks, TblBookingBondsInvsDet.BlockNo, TblProjecInvestmentDet.PartCode, TblProjecInvestment.Square, TblProjecInvestment.SquareCode, TblBranchesData.branch_name, "
MySQL = MySQL & " TblBranchesData.branch_namee, TblProjecInvestment.Name AS ProjectName, TblProjecInvestment.NameE AS ProjectNameE, TblBookingBondsInvs.RecordDate, TblBookingBondsInvs.Name AS BookerName, "
MySQL = MySQL & " TblBookingBondsInvs.ID AS BookingID, TblBookingBondsInvs.Telephone, BanksData.BankName, BanksData.BankNamee "
MySQL = MySQL & " FROM TblProjecInvestment INNER JOIN "
MySQL = MySQL & " TblProjecInvestmentDet ON TblProjecInvestment.ID = TblProjecInvestmentDet.ProjInvID INNER JOIN "
MySQL = MySQL & " TblBranchesData ON TblProjecInvestment.BrnchID = TblBranchesData.branch_id RIGHT OUTER JOIN "
MySQL = MySQL & " TblBookingBondsInvsDet LEFT OUTER JOIN"
MySQL = MySQL & " TblModelInves ON TblBookingBondsInvsDet.ModelID = TblModelInves.ID ON TblProjecInvestmentDet.ID = TblBookingBondsInvsDet.PartID FULL OUTER JOIN "
MySQL = MySQL & " TblBookingBondsInvs ON TblBookingBondsInvsDet.BokID = TblBookingBondsInvs.ID CROSS JOIN "
MySQL = MySQL & " BanksData "
MySQL = MySQL & " Where (TblBookingBondsInvsDet.BokID = " & val(TxtSerial1.Text) & ") And (TblBookingBondsInvsDet.TypeTrans = 1) "

        If SystemOptions.UserInterface = ArabicInterface Then
            StrFileName = App.path & "\REPORTS\REPORTS NEW\" & "RepBookingUnits.rpt"
        Else
            StrFileName = App.path & "\REPORTS\REPORTS NEW\" & "RepBookingUnits.rpt"
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
        ' xReport.ParameterFields(2).AddCurrentValue RPTComment_Arabic
        StrReportTitle = "" '& StrAccountName
      Else
        xReport.ParameterFields(1).AddCurrentValue cCompanyInfo.ArabCompanyName ' RPTCompany_Name_Eng
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


Private Sub VSFlexGrid1_CellButtonClick(ByVal Row As Long, ByVal Col As Long)



With VSFlexGrid1
Select Case .ColKey(Col)
Case "Show1"


        Dim Str_Path As String
            Str_Path = App.path & "\" & SystemOptions.ImagesPath & "\" & VSFlexGrid1.TextMatrix(Row, VSFlexGrid1.ColIndex("PartID")) & ".JPG"
 
        If Dir(Str_Path) <> "" Then
                DBPix201.ImageLoadFile (App.path & "\" & SystemOptions.ImagesPath & "\" & VSFlexGrid1.TextMatrix(Row, VSFlexGrid1.ColIndex("PartID")) & ".JPG")
                Else
                Str_Path = App.path & "\" & SystemOptions.ImagesPath & "\DefualtRealState.JPG"
                If Dir(Str_Path) <> "" Then
                        DBPix201.ImageLoadFile (App.path & "\" & SystemOptions.ImagesPath & "\DefualtRealState.JPG")
                Else
                        DBPix201.ImageClear
               End If





End If
End Select
End With




End Sub

Private Sub VSFlexGrid1_StartEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
  With Me.VSFlexGrid1

        Select Case .ColKey(Col)
        Case "Show1"
        .ColComboList(.ColIndex("Show1")) = "..."
     End Select
  End With
  
End Sub
