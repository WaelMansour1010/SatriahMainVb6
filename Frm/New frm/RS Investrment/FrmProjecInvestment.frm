VERSION 5.00
Object = "{C115893A-A3BF-43AF-B28D-69DB846077F3}#1.0#0"; "vsflex8u.ocx"
Object = "{0AFE7BE0-11B7-4A3E-978D-D4501E9A57FE}#1.0#0"; "c1sizer.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{FE5DCFAD-BC1D-11D2-94CF-004005455FAA}#1.4#0"; "ImpulseButton.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#12.0#0"; "SuiteCtrls.ocx"
Object = "{D155F1AE-D9A4-458C-8CEE-498CB717DB7B}#1.0#0"; "DBPix20.ocx"
Begin VB.Form FrmProjecInvestment 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   10065
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   17790
   Icon            =   "FrmProjecInvestment.frx":0000
   LinkTopic       =   "Form3"
   MDIChild        =   -1  'True
   RightToLeft     =   -1  'True
   ScaleHeight     =   10065
   ScaleWidth      =   17790
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
      TabIndex        =   23
      TabStop         =   0   'False
      Top             =   0
      Width           =   17790
      _cx             =   31380
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
         TabIndex        =   55
         Top             =   0
         Width           =   17820
         Begin VB.TextBox tXTRootAccount 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   3240
            RightToLeft     =   -1  'True
            TabIndex        =   57
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
            TabIndex        =   56
            Top             =   360
            Visible         =   0   'False
            Width           =   2055
         End
         Begin ImpulseButton.ISButton btnLast 
            Height          =   315
            Left            =   450
            TabIndex        =   58
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
            ButtonImage     =   "FrmProjecInvestment.frx":6852
            ColorButton     =   16777215
            AcclimateGrayTones=   -1  'True
            DrawFocusRectangle=   0   'False
            DisabledImageExtraction=   0
            DisabledImageStyle=   1
         End
         Begin ImpulseButton.ISButton btnNext 
            Height          =   315
            Left            =   915
            TabIndex        =   59
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
            ButtonImage     =   "FrmProjecInvestment.frx":6BEC
            ColorButton     =   16777215
            DrawFocusRectangle=   0   'False
            DisabledImageStyle=   1
         End
         Begin ImpulseButton.ISButton btnPrevious 
            Height          =   315
            Left            =   1515
            TabIndex        =   60
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
            ButtonImage     =   "FrmProjecInvestment.frx":6F86
            ColorButton     =   16777215
            DrawFocusRectangle=   0   'False
            DisabledImageStyle=   1
         End
         Begin ImpulseButton.ISButton btnFirst 
            Height          =   315
            Left            =   2040
            TabIndex        =   61
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
            ButtonImage     =   "FrmProjecInvestment.frx":7320
            ColorButton     =   16777215
            DrawFocusRectangle=   0   'False
            DisabledImageStyle=   1
         End
         Begin VB.Image Image1 
            Height          =   615
            Left            =   13200
            Picture         =   "FrmProjecInvestment.frx":76BA
            Stretch         =   -1  'True
            Top             =   120
            Visible         =   0   'False
            Width           =   735
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "»Ì«‰«  «·„‘«—Ì⁄ «·«” À„«—Ì…"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   14.25
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
            TabIndex        =   62
            Top             =   240
            Width           =   4080
         End
      End
      Begin C1SizerLibCtl.C1Elastic C1Elastic2 
         Height          =   735
         Left            =   0
         TabIndex        =   24
         TabStop         =   0   'False
         Top             =   720
         Width           =   17850
         _cx             =   31485
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
            Height          =   315
            Left            =   14055
            RightToLeft     =   -1  'True
            TabIndex        =   0
            Top             =   240
            Width           =   2055
         End
         Begin MSComCtl2.DTPicker XPDtbTrans 
            Height          =   315
            Left            =   9525
            TabIndex        =   25
            Top             =   240
            Width           =   1605
            _ExtentX        =   2831
            _ExtentY        =   556
            _Version        =   393216
            Format          =   183238657
            CurrentDate     =   38784
         End
         Begin MSDataListLib.DataCombo Dcbranch 
            Bindings        =   "FrmProjecInvestment.frx":8ABF
            Height          =   315
            Left            =   615
            TabIndex        =   1
            Top             =   240
            Width           =   6390
            _ExtentX        =   11271
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
         Begin MSComDlg.CommonDialog CommonDialog1 
            Left            =   0
            Top             =   0
            _ExtentX        =   847
            _ExtentY        =   847
            _Version        =   393216
         End
         Begin VB.Label lbl 
            Alignment       =   2  'Center
            BackColor       =   &H00E2E9E9&
            Caption         =   "ﬂÊœ «·„‘—Ê⁄"
            Height          =   285
            Index           =   9
            Left            =   3300
            TabIndex        =   68
            Top             =   240
            Visible         =   0   'False
            Width           =   1635
         End
         Begin VB.Label lbl 
            Alignment       =   2  'Center
            BackColor       =   &H00E2E9E9&
            Caption         =   "«· «—ÌŒ"
            Height          =   285
            Index           =   2
            Left            =   11550
            TabIndex        =   28
            Top             =   240
            Width           =   885
         End
         Begin VB.Label lbl 
            Alignment       =   2  'Center
            BackColor       =   &H00E2E9E9&
            Caption         =   "—ﬁ„ «·Õ—ﬂ…"
            Height          =   285
            Index           =   4
            Left            =   16605
            RightToLeft     =   -1  'True
            TabIndex        =   27
            Top             =   240
            Width           =   945
         End
         Begin VB.Label lbl 
            Alignment       =   2  'Center
            BackColor       =   &H00E2E9E9&
            Caption         =   "«·›—⁄"
            Height          =   285
            Index           =   7
            Left            =   7095
            TabIndex        =   26
            Top             =   240
            Width           =   1635
         End
      End
      Begin C1SizerLibCtl.C1Elastic C1Elastic4 
         Height          =   975
         Left            =   0
         TabIndex        =   29
         TabStop         =   0   'False
         Top             =   1560
         Width           =   17850
         _cx             =   31485
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
         Begin VB.TextBox TxtProjectCode 
            Alignment       =   1  'Right Justify
            Height          =   315
            Left            =   15285
            RightToLeft     =   -1  'True
            TabIndex        =   69
            Top             =   120
            Width           =   930
         End
         Begin VB.TextBox txtFile 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFFF&
            Height          =   315
            Left            =   0
            RightToLeft     =   -1  'True
            TabIndex        =   66
            Top             =   0
            Visible         =   0   'False
            Width           =   1395
         End
         Begin VB.TextBox TxtNameE 
            Alignment       =   1  'Right Justify
            Height          =   315
            Left            =   510
            RightToLeft     =   -1  'True
            TabIndex        =   34
            Top             =   120
            Width           =   6495
         End
         Begin VB.TextBox TxtNameP 
            Alignment       =   1  'Right Justify
            Height          =   315
            Left            =   9525
            RightToLeft     =   -1  'True
            TabIndex        =   32
            Top             =   120
            Width           =   5655
         End
         Begin VB.TextBox TxtSearchCode 
            Alignment       =   1  'Right Justify
            Height          =   315
            Left            =   15285
            RightToLeft     =   -1  'True
            TabIndex        =   2
            Top             =   480
            Width           =   930
         End
         Begin VB.TextBox TxtPercenValue 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFFFF&
            Height          =   315
            Left            =   510
            RightToLeft     =   -1  'True
            TabIndex        =   5
            Top             =   480
            Width           =   1845
         End
         Begin VB.ComboBox DcbTypePrstg 
            Height          =   315
            ItemData        =   "FrmProjecInvestment.frx":8AD4
            Left            =   2550
            List            =   "FrmProjecInvestment.frx":8AD6
            RightToLeft     =   -1  'True
            TabIndex        =   4
            Top             =   480
            Width           =   4455
         End
         Begin MSDataListLib.DataCombo DcboEmpName 
            Bindings        =   "FrmProjecInvestment.frx":8AD8
            Height          =   315
            Left            =   9525
            TabIndex        =   3
            Top             =   480
            Width           =   5655
            _ExtentX        =   9975
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
         Begin MSComDlg.CommonDialog CD1 
            Left            =   0
            Top             =   0
            _ExtentX        =   847
            _ExtentY        =   847
            _Version        =   393216
         End
         Begin VB.Label lbl 
            Alignment       =   2  'Center
            BackColor       =   &H00E2E9E9&
            Caption         =   "«”„ «·„‘—Ê⁄ «‰Ã·Ì“Ì"
            Height          =   285
            Index           =   3
            Left            =   7215
            TabIndex        =   35
            Top             =   120
            Width           =   1515
         End
         Begin VB.Label lbl 
            Alignment       =   2  'Center
            BackColor       =   &H00E2E9E9&
            Caption         =   "«”„ «·„‘—Ê⁄ ⁄—»Ì"
            Height          =   285
            Index           =   1
            Left            =   16365
            TabIndex        =   33
            Top             =   120
            Width           =   1395
         End
         Begin VB.Label lbl 
            Alignment       =   2  'Center
            BackColor       =   &H00E2E9E9&
            Caption         =   "„œÌ— «·„‘—Ê⁄"
            Height          =   285
            Index           =   15
            Left            =   16365
            TabIndex        =   31
            Top             =   480
            Width           =   1395
         End
         Begin VB.Label lbl 
            Alignment       =   2  'Center
            BackColor       =   &H00E2E9E9&
            Caption         =   "⁄„Ê·… «œ«—… «·„‘«—Ì⁄"
            Height          =   285
            Index           =   0
            Left            =   7095
            RightToLeft     =   -1  'True
            TabIndex        =   30
            Top             =   480
            Width           =   1785
         End
      End
      Begin C1SizerLibCtl.C1Elastic C1Elastic5 
         Height          =   735
         Left            =   0
         TabIndex        =   36
         TabStop         =   0   'False
         Top             =   8640
         Width           =   17850
         _cx             =   31485
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
         Begin C1SizerLibCtl.C1Elastic C1Elastic6 
            Height          =   495
            Left            =   360
            TabIndex        =   39
            TabStop         =   0   'False
            Top             =   120
            Width           =   4185
            _cx             =   7382
            _cy             =   873
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
               Left            =   3030
               RightToLeft     =   -1  'True
               TabIndex        =   43
               Top             =   120
               Width           =   1020
            End
            Begin VB.Label Label2 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "⁄œœ «·”Ã·« :"
               Height          =   210
               Index           =   1
               Left            =   1095
               RightToLeft     =   -1  'True
               TabIndex        =   42
               Top             =   120
               Width           =   975
            End
            Begin VB.Label LabCurrRec 
               Alignment       =   2  'Center
               BackColor       =   &H00E2E9E9&
               ForeColor       =   &H00800000&
               Height          =   210
               Left            =   2205
               RightToLeft     =   -1  'True
               TabIndex        =   41
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
               TabIndex        =   40
               Top             =   120
               Width           =   780
            End
         End
         Begin MSDataListLib.DataCombo DCboUserName 
            Height          =   315
            Left            =   7965
            TabIndex        =   37
            Top             =   120
            Width           =   6030
            _ExtentX        =   10636
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
            Height          =   270
            Index           =   8
            Left            =   14295
            TabIndex        =   38
            Top             =   120
            Width           =   1875
         End
      End
      Begin C1SizerLibCtl.C1Elastic C1Elastic7 
         Height          =   615
         Left            =   0
         TabIndex        =   44
         TabStop         =   0   'False
         Top             =   9480
         Width           =   17850
         _cx             =   31485
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
            Left            =   15030
            TabIndex        =   45
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
            ButtonImage     =   "FrmProjecInvestment.frx":8AED
            ColorButton     =   14871017
            DrawFocusRectangle=   0   'False
         End
         Begin ImpulseButton.ISButton btnSave 
            Height          =   330
            Left            =   11115
            TabIndex        =   46
            ToolTipText     =   "Õ›Ÿ «·»Ì«‰«  œ«Œ· ﬁ«⁄œ… «·»Ì«‰« "
            Top             =   120
            Width           =   1320
            _ExtentX        =   2328
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
            ButtonImage     =   "FrmProjecInvestment.frx":F34F
            ColorButton     =   14871017
            DrawFocusRectangle=   0   'False
         End
         Begin ImpulseButton.ISButton btnModify 
            Height          =   330
            Left            =   13440
            TabIndex        =   47
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
            ButtonImage     =   "FrmProjecInvestment.frx":F6E9
            ColorButton     =   14871017
            DrawFocusRectangle=   0   'False
         End
         Begin ImpulseButton.ISButton BtnUndo 
            Height          =   330
            Left            =   9390
            TabIndex        =   48
            ToolTipText     =   "·· —«Ã⁄ ⁄‰ «·ÕœÀ Ê«·—ÃÊ⁄ «·Ï «·Ê÷⁄ «·ÿ»Ì⁄Ì"
            Top             =   120
            Width           =   1470
            _ExtentX        =   2593
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
            ButtonImage     =   "FrmProjecInvestment.frx":15F4B
            ColorButton     =   14871017
            DrawFocusRectangle=   0   'False
         End
         Begin ImpulseButton.ISButton btnDelete 
            Height          =   330
            Left            =   7725
            TabIndex        =   49
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
            ButtonImage     =   "FrmProjecInvestment.frx":162E5
            ColorButton     =   14871017
            DrawFocusRectangle=   0   'False
         End
         Begin ImpulseButton.ISButton btnCancel 
            Height          =   330
            Left            =   2205
            TabIndex        =   50
            ToolTipText     =   "«·Œ—ÊÃ «·Ï  «·‰«›–… «·—∆Ì”Ì…"
            Top             =   120
            Width           =   1230
            _ExtentX        =   2170
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
            ButtonImage     =   "FrmProjecInvestment.frx":1687F
            ColorButton     =   14871017
            DrawFocusRectangle=   0   'False
            DisabledImageStyle=   1
         End
         Begin ImpulseButton.ISButton ISButton5 
            Height          =   405
            Left            =   6120
            TabIndex        =   51
            TabStop         =   0   'False
            ToolTipText     =   "ÿ»«⁄… «·»Ì«‰«  "
            Top             =   120
            Visible         =   0   'False
            Width           =   1020
            _ExtentX        =   1799
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
            ButtonImage     =   "FrmProjecInvestment.frx":16C19
            ColorButton     =   14871017
            DrawFocusRectangle=   0   'False
            DisabledImageStyle=   1
         End
         Begin ImpulseButton.ISButton ISButton8 
            Height          =   330
            Left            =   4035
            TabIndex        =   52
            TabStop         =   0   'False
            ToolTipText     =   "(Ctrl+F)  ··»ÕÀ ≈÷€ÿ Â–« «·„› «Õ √Ê ≈÷€ÿ "
            Top             =   120
            Visible         =   0   'False
            Width           =   1005
            _ExtentX        =   1773
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
            ButtonImage     =   "FrmProjecInvestment.frx":1D47B
            ColorButton     =   14871017
            DrawFocusRectangle=   0   'False
         End
      End
      Begin C1SizerLibCtl.C1Elastic C1Elastic3 
         Height          =   6735
         Left            =   0
         TabIndex        =   53
         TabStop         =   0   'False
         Top             =   2640
         Width           =   17850
         _cx             =   31485
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
            Height          =   375
            Left            =   615
            RightToLeft     =   -1  'True
            TabIndex        =   74
            Top             =   600
            Visible         =   0   'False
            Width           =   1485
         End
         Begin VB.CommandButton Command3 
            Caption         =   "«œ—«Ã ’Ê—…"
            Height          =   372
            Left            =   0
            RightToLeft     =   -1  'True
            TabIndex        =   72
            Top             =   5160
            Width           =   4260
         End
         Begin VB.TextBox Txt_path_photo 
            Alignment       =   1  'Right Justify
            Enabled         =   0   'False
            Height          =   315
            Left            =   9525
            RightToLeft     =   -1  'True
            TabIndex        =   71
            Top             =   480
            Width           =   6690
         End
         Begin VB.CommandButton Command2 
            Caption         =   " Õ„Ì· «·„Œÿÿ"
            Height          =   375
            Left            =   16500
            RightToLeft     =   -1  'True
            TabIndex        =   70
            Top             =   480
            Width           =   1125
         End
         Begin VB.TextBox TxtSquare 
            Alignment       =   1  'Right Justify
            Height          =   315
            Left            =   9525
            RightToLeft     =   -1  'True
            TabIndex        =   7
            Top             =   120
            Width           =   5745
         End
         Begin VB.TextBox TxtSquareCode 
            Alignment       =   1  'Right Justify
            Height          =   315
            Left            =   15285
            RightToLeft     =   -1  'True
            TabIndex        =   6
            Top             =   120
            Width           =   930
         End
         Begin VSFlex8UCtl.VSFlexGrid Grid 
            Height          =   4755
            Left            =   4530
            TabIndex        =   54
            Top             =   840
            Width           =   13230
            _cx             =   23336
            _cy             =   8387
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
            Cols            =   20
            FixedRows       =   1
            FixedCols       =   1
            RowHeightMin    =   320
            RowHeightMax    =   0
            ColWidthMin     =   0
            ColWidthMax     =   0
            ExtendLastCol   =   -1  'True
            FormatString    =   $"FrmProjecInvestment.frx":1D815
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
            Left            =   16095
            TabIndex        =   64
            Top             =   5640
            Width           =   1590
            _ExtentX        =   2805
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
            ButtonImage     =   "FrmProjecInvestment.frx":1DB3D
            ColorButton     =   14871017
            DrawFocusRectangle=   0   'False
         End
         Begin ImpulseButton.ISButton Cmd 
            Height          =   270
            Index           =   1
            Left            =   14415
            TabIndex        =   65
            Top             =   5640
            Width           =   1515
            _ExtentX        =   2672
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
            ButtonImage     =   "FrmProjecInvestment.frx":1E0D7
            ColorButton     =   14871017
            DrawFocusRectangle=   0   'False
         End
         Begin XtremeSuiteControls.RadioButton RdTyp 
            Height          =   315
            Index           =   0
            Left            =   6120
            TabIndex        =   8
            Top             =   120
            Width           =   885
            _Version        =   786432
            _ExtentX        =   1561
            _ExtentY        =   556
            _StockProps     =   79
            Caption         =   "ÌœÊÌ"
            BackColor       =   14871017
            UseVisualStyle  =   -1  'True
            TextAlignment   =   1
            RightToLeft     =   -1  'True
         End
         Begin XtremeSuiteControls.RadioButton RdTyp 
            Height          =   315
            Index           =   1
            Left            =   5010
            TabIndex        =   9
            Top             =   120
            Width           =   855
            _Version        =   786432
            _ExtentX        =   1508
            _ExtentY        =   556
            _StockProps     =   79
            Caption         =   "„‰ „·›"
            BackColor       =   14871017
            UseVisualStyle  =   -1  'True
            TextAlignment   =   1
            RightToLeft     =   -1  'True
         End
         Begin ImpulseButton.ISButton ISButton3 
            Height          =   315
            Left            =   360
            TabIndex        =   11
            Top             =   120
            Width           =   2085
            _ExtentX        =   3678
            _ExtentY        =   556
            Caption         =   "«” Ì—«œ «·„·›"
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
            ButtonImage     =   "FrmProjecInvestment.frx":1E671
            ColorButton     =   14871017
            DrawFocusRectangle=   0   'False
            DisabledImageExtraction=   0
            LowerToggledContent=   0   'False
         End
         Begin ImpulseButton.ISButton ISButton4 
            Height          =   315
            Left            =   2640
            TabIndex        =   10
            ToolTipText     =   "«÷«›… «·»Ì«‰«  «·Ï «·œ« «"
            Top             =   120
            Width           =   2115
            _ExtentX        =   3731
            _ExtentY        =   556
            Caption         =   "Õœœ «·„”«—"
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
            ButtonImage     =   "FrmProjecInvestment.frx":24ED3
            ColorButton     =   14871017
            DrawFocusRectangle=   0   'False
            DisabledImageExtraction=   0
            LowerToggledContent=   0   'False
         End
         Begin DBPIXLib.DBPix20 DBPix201 
            Height          =   3480
            Left            =   30
            TabIndex        =   73
            Top             =   960
            Width           =   4515
            _Version        =   131072
            _ExtentX        =   7964
            _ExtentY        =   6138
            _StockProps     =   1
            BackColor       =   16777152
            _Image          =   "FrmProjecInvestment.frx":2B735
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
            Caption         =   "«·„—»⁄"
            Height          =   285
            Index           =   6
            Left            =   16365
            TabIndex        =   67
            Top             =   120
            Width           =   1395
         End
         Begin VB.Label lbl 
            Alignment       =   2  'Center
            BackColor       =   &H00E2E9E9&
            Caption         =   " ›«’Ì· "
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   178
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C00000&
            Height          =   285
            Index           =   5
            Left            =   6255
            RightToLeft     =   -1  'True
            TabIndex        =   63
            Top             =   480
            Width           =   1905
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
      TabIndex        =   16
      Text            =   "modflag"
      Top             =   4200
      Visible         =   0   'False
      Width           =   465
   End
   Begin VB.ComboBox CmbType 
      BackColor       =   &H80000018&
      Height          =   315
      ItemData        =   "FrmProjecInvestment.frx":2B74D
      Left            =   18360
      List            =   "FrmProjecInvestment.frx":2B75D
      RightToLeft     =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   15
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
      TabIndex        =   14
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
      TabIndex        =   13
      Top             =   1680
      Visible         =   0   'False
      Width           =   3105
   End
   Begin VB.TextBox Emp_id 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   18600
      RightToLeft     =   -1  'True
      TabIndex        =   12
      Top             =   5640
      Visible         =   0   'False
      Width           =   1575
   End
   Begin MSDataListLib.DataCombo DCUser 
      CausesValidation=   0   'False
      Height          =   315
      Left            =   18720
      TabIndex        =   17
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
      TabIndex        =   18
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
            Picture         =   "FrmProjecInvestment.frx":2B776
            Key             =   "CompanyName"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmProjecInvestment.frx":2BB10
            Key             =   "Ser"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmProjecInvestment.frx":2BEAA
            Key             =   "Vac_Name"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmProjecInvestment.frx":2C244
            Key             =   "ShareCount"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmProjecInvestment.frx":2C5DE
            Key             =   "Dis_Count"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmProjecInvestment.frx":2C978
            Key             =   "Bouns"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmProjecInvestment.frx":2CD12
            Key             =   "SharesValue"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmProjecInvestment.frx":2D2AC
            Key             =   "BuyValue"
         EndProperty
      EndProperty
   End
   Begin ImpulseButton.ISButton BtnUpdate 
      Height          =   330
      Left            =   18480
      TabIndex        =   19
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
      ButtonImage     =   "FrmProjecInvestment.frx":2D646
      ColorButton     =   14871017
      DrawFocusRectangle=   0   'False
   End
   Begin ImpulseButton.ISButton ISButton1 
      Height          =   405
      Left            =   18840
      TabIndex        =   21
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
      ButtonImage     =   "FrmProjecInvestment.frx":33EA8
      ColorButton     =   14871017
      DrawFocusRectangle=   0   'False
      DisabledImageStyle=   1
   End
   Begin ImpulseButton.ISButton btnQuery 
      Height          =   330
      Left            =   19800
      TabIndex        =   22
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
      ButtonImage     =   "FrmProjecInvestment.frx":3A70A
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
      TabIndex        =   20
      Top             =   360
      Width           =   855
   End
End
Attribute VB_Name = "FrmProjecInvestment"
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
 Dim ii As Long
 Public LonRow As Double
Public LngCol As Double

Private Sub Cmd_Click(index As Integer)
If Me.TxtModFlg.text <> "R" Then
Select Case index
Case 0
RemoveGridRow
Case 1
 Grid.Clear flexClearScrollable, flexClearEverything
            Grid.rows = 2
 End Select
End If
End Sub
Private Sub ReLineGrid()
    Dim IntCounter As Integer
    IntCounter = 0
    Dim i As Integer
    With Me.Grid
        For i = .FixedRows To .rows - 1
            If .TextMatrix(i, .ColIndex("BlockNo")) <> "" Then
                IntCounter = IntCounter + 1
                .TextMatrix(i, .ColIndex("Ser")) = IntCounter
          '      .TextMatrix(i, .ColIndex("PartCode")) = Me.TxtProjectCode.Text & Me.TxtSquareCode.Text & .TextMatrix(i, .ColIndex("BlockNo")) & .TextMatrix(i, .ColIndex("PartNo"))
               .TextMatrix(i, .ColIndex("Total")) = val(.TextMatrix(i, .ColIndex("AddValue"))) + val(.TextMatrix(i, .ColIndex("ValueOffice"))) + val(.TextMatrix(i, .ColIndex("MOHPrice")))
            End If
        Next i
    End With
End Sub

Private Sub Command2_Click()
CommonDialog1.filter = "Pic(*.Jpg)|*.Jpg"
CommonDialog1.InitDir = App.path & "\Images"
CommonDialog1.DefaultExt = "txt"
CommonDialog1.DialogTitle = "Select File"
CommonDialog1.ShowOpen
Txt_path_photo.text = CommonDialog1.FileName
End Sub

Private Sub Command3_Click()
CommonDialog1.filter = "Pic(*.Jpg)|*.Jpg"
CommonDialog1.InitDir = App.path & "\Images"
CommonDialog1.DefaultExt = "txt"
CommonDialog1.DialogTitle = "Select File"
CommonDialog1.ShowOpen
Txt_Path_General_photo.text = CommonDialog1.FileName

         Dim Str_Path As String
             Str_Path = Txt_Path_General_photo.text
    
             If Dir(Str_Path) <> "" Then
                DBPix201.ImageLoadFile (Txt_Path_General_photo.text)
              Else
                Str_Path = Txt_Path_General_photo.text
                If Dir(Str_Path) <> "" Then
                        DBPix201.ImageLoadFile (App.path & "\" & SystemOptions.ImagesPath & "\DefualtRealState.JPG")
                Else
                        DBPix201.ImageClear
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

Private Sub DcboEmpName_Change()
DcboEmpName_Click (0)
End Sub
Private Sub DcboEmpName_Click(Area As Integer)
 If val(DcboEmpName.BoundText) = 0 Then Exit Sub
 Dim EmpCode  As String
    GetEmployeeIDFromCode , , DcboEmpName.BoundText, EmpCode
    TxtSearchCode.text = EmpCode
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
   Else
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
    conection = "select * from TblProjecInvestment order by  ID "
    Set BKGrndPic = New ClsBackGroundPic
    Set RsSavRec = New ADODB.Recordset
    RsSavRec.CursorLocation = adUseClient
    RsSavRec.Open conection, Cn, adOpenStatic, adLockOptimistic, adCmdText
    Me.TxtModFlg.text = "R"
    Resize_Form Me
    Dim Dcombos As New ClsDataCombos
    Dcombos.GetBranches Me.dcBranch
    Dcombos.GetEmployees Me.DcboEmpName
    Dcombos.GetUsers Me.DCboUserName
    
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
             If Me.TxtModFlg.text = "E" Then
                 StrSQL = "Delete From TblProjecInvestmentDet Where ProjInvID =" & val(TxtSerial1.text) & ""
                  Cn.Execute StrSQL, , adExecuteNoRecords
              End If
      ''//////////
    RsSavRec.Fields("BrnchID").value = val(Me.dcBranch.BoundText)
    RsSavRec.Fields("RecordDate").value = XPDtbTrans.value
    RsSavRec.Fields("Name").value = TxtNameP.text
    RsSavRec.Fields("NameE").value = TxtNameE.text
    RsSavRec.Fields("path_photo").value = Txt_path_photo.text
    
    RsSavRec.Fields("Path_General_photo").value = Txt_Path_General_photo.text
    
    RsSavRec.Fields("EmpID").value = val(Me.DcboEmpName.BoundText)
    RsSavRec.Fields("TypePrstg").value = val(DcbTypePrstg.ListIndex)
    RsSavRec.Fields("UserID").value = val(Me.DCboUserName.BoundText)
    RsSavRec.Fields("PercenValue").value = val(TxtPercenValue.text)
    RsSavRec.Fields("Square").value = TxtSquare.text
    RsSavRec.Fields("SquareCode").value = TxtSquareCode.text
    RsSavRec.Fields("ProjectCode").value = TxtProjectCode.text
   If RdTyp(1).value = True Then
    RsSavRec.Fields("TypeImport").value = 1
   Else
   RsSavRec.Fields("TypeImport").value = 0
   End If
   RsSavRec.Fields("Path").value = Me.txtFile.text
    RsSavRec.update
  
''//////////////////////////
Dim ProjeInvsID As Double
      Set RsDevsub = New ADODB.Recordset
    StrSQL = "SELECT  *  from TblProjecInvestmentDet Where (1 = -1)"
    RsDevsub.Open StrSQL, Cn, adOpenKeyset, adLockOptimistic, adCmdText
    Dim i As Integer
    Dim str2 As String
    With Me.Grid
       For i = .FixedRows To .rows - 1
       If .TextMatrix(i, .ColIndex("BlockNo")) <> "" Then
       RsDevsub.AddNew
       ProjeInvsID = val(.TextMatrix(i, .ColIndex("ID")))
       If Me.Checked(ProjeInvsID) = True Then
        Else
       ProjeInvsID = 1
        maxx ProjeInvsID
       
        .TextMatrix(i, .ColIndex("ID")) = ProjeInvsID
       End If
                RsDevsub("ModelName").value = IIf((.TextMatrix(i, .ColIndex("Name"))) = "", Null, .TextMatrix(i, .ColIndex("Name")))
                RsDevsub("ProjInvID").value = val(Me.TxtSerial1.text)
                RsDevsub("ID").value = IIf((.TextMatrix(i, .ColIndex("ID"))) = "", 0, .TextMatrix(i, .ColIndex("ID")))
                RsDevsub("BlockNo").value = IIf((.TextMatrix(i, .ColIndex("BlockNo"))) = "", Null, .TextMatrix(i, .ColIndex("BlockNo")))
                RsDevsub("PartNo").value = IIf((.TextMatrix(i, .ColIndex("PartNo"))) = "", Null, .TextMatrix(i, .ColIndex("PartNo")))
                RsDevsub("ModelID").value = IIf((.TextMatrix(i, .ColIndex("ModelID"))) = "", Null, val(.TextMatrix(i, .ColIndex("ModelID"))))
                RsDevsub("LandArea").value = IIf((.TextMatrix(i, .ColIndex("LandArea"))) = "", Null, val(.TextMatrix(i, .ColIndex("LandArea"))))
                RsDevsub("HouseArea").value = IIf((.TextMatrix(i, .ColIndex("HouseArea"))) = "", Null, val(.TextMatrix(i, .ColIndex("HouseArea"))))
                RsDevsub("DeveloperCode").value = IIf((.TextMatrix(i, .ColIndex("DeveloperCode"))) = "", Null, (.TextMatrix(i, .ColIndex("DeveloperCode"))))
                RsDevsub("MarketPrice").value = IIf((.TextMatrix(i, .ColIndex("MarketPrice"))) = "", Null, val(.TextMatrix(i, .ColIndex("MarketPrice"))))
                RsDevsub("MOHPrice").value = IIf((.TextMatrix(i, .ColIndex("MOHPrice"))) = "", Null, val(.TextMatrix(i, .ColIndex("MOHPrice"))))
                RsDevsub("BedroomsNo").value = IIf((.TextMatrix(i, .ColIndex("BedroomsNo"))) = "", Null, val(.TextMatrix(i, .ColIndex("BedroomsNo"))))
                RsDevsub("ConstStatus").value = IIf((.TextMatrix(i, .ColIndex("ConstStatus"))) = "", Null, val(.TextMatrix(i, .ColIndex("ConstStatus"))))
                RsDevsub("StatusID").value = IIf((.TextMatrix(i, .ColIndex("StatusID"))) = "", Null, val(.TextMatrix(i, .ColIndex("StatusID"))))
                RsDevsub("Remarks").value = IIf((.TextMatrix(i, .ColIndex("Remarks"))) = "", Null, (.TextMatrix(i, .ColIndex("Remarks"))))
                RsDevsub("PartCode").value = IIf((.TextMatrix(i, .ColIndex("PartCode"))) = "", Null, (.TextMatrix(i, .ColIndex("PartCode"))))
                RsDevsub("Total").value = IIf((.TextMatrix(i, .ColIndex("Total"))) = "", 0, .TextMatrix(i, .ColIndex("Total")))
                RsDevsub("AddValue").value = IIf((.TextMatrix(i, .ColIndex("AddValue"))) = "", 0, .TextMatrix(i, .ColIndex("AddValue")))
                RsDevsub("ValueOffice").value = IIf((.TextMatrix(i, .ColIndex("ValueOffice"))) = "", 0, val(.TextMatrix(i, .ColIndex("ValueOffice"))))
       RsDevsub.update
      End If
     Next i
    End With
'''///////////////
  
      Select Case Me.TxtModFlg.text
        Case "N"
            Dim Msg As String
            If SystemOptions.UserInterface = ArabicInterface Then
                Msg = "  „ Õ›Ÿ »Ì«‰«  Â–Â «·⁄„·Ì… " & CHR(13)
                Msg = Msg + "Â·  —€» ›Ì ≈÷«›… »Ì«‰«  √Œ—Ï"
            Else
               Msg = " This record alredy saved... " & CHR(13)
                Msg = Msg + " You want to enter another record?"
           End If
                If MsgBox(Msg, vbYesNo + vbQuestion + vbMsgBoxRight + vbMsgBoxRtlReading + vbDefaultButton2, App.Title) = vbYes Then
              
                Me.Refresh
                FiLLTXT
                TxtModFlg = "R"
                 If SystemOptions.UserInterface = ArabicInterface Then
             Else
              
                Me.Refresh
                FiLLTXT
                TxtModFlg = "R"
                MsgBox "Changes Was Saved ... Continuation Add Data ", vbInformation + vbOKOnly + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
            End If
                Call btnNew_Click
            Else
                Me.Refresh
                FiLLTXT
                TxtModFlg = "R"
            End If
         Case "E"
            If SystemOptions.UserInterface = ArabicInterface Then
                MsgBox " „ Õ›Ÿ Â–Â «· ⁄œÌ·« ", vbInformation + vbOKOnly + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
                
                Me.Refresh
                FiLLTXT
                TxtModFlg = "R"
            Else
                MsgBox "Changes was saved", vbInformation + vbOKOnly + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
                
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
Function Checked(Optional ProjeInvsID As Double = 0) As Boolean
     Checked = False
     Dim RsDev As ADODB.Recordset
    Dim StrSQL As String
    Set RsDev = New ADODB.Recordset
  If ProjeInvsID <> 0 Then
   StrSQL = " select * from FoxySerial where ProjeInvsID=" & ProjeInvsID & ""
   RsDev.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText
If RsDev.RecordCount > 0 Then
Checked = True
Else
Checked = False
End If
End If
End Function
Sub maxx(Optional ByRef ProjeInvsID As Double = 0)
  Dim RsDev As ADODB.Recordset
  Dim StrSQL As String
  Set RsDev = New ADODB.Recordset
If ProjeInvsID <> 0 Then
   StrSQL = " select max(ProjeInvsID) as mx from FoxySerial"
   RsDev.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText
   ProjeInvsID = IIf(IsNull(RsDev("mx").value), 0, RsDev("mx").value) + 1
   Set RsDev = New ADODB.Recordset
   RsDev.Open "FoxySerial", Cn, adOpenStatic, adLockOptimistic, adCmdTable
   RsDev.AddNew
   RsDev("ProjeInvsID").value = ProjeInvsID
   RsDev.update
End If
End Sub
' full data from database
'+++++++++++++++++++++++++++++++++++++++
Public Sub FiLLTXT()
   On Error GoTo ErrTrap
    TxtSerial1.text = IIf(IsNull(RsSavRec.Fields("ID").value), "", RsSavRec.Fields("ID").value)
    XPDtbTrans.value = IIf(IsNull(RsSavRec.Fields("RecordDate").value), Date, RsSavRec.Fields("RecordDate").value)
    dcBranch.BoundText = IIf(IsNull(RsSavRec.Fields("BrnchID").value), "", RsSavRec.Fields("BrnchID").value)
    Me.TxtNameP.text = IIf(IsNull(RsSavRec.Fields("Name").value), "", RsSavRec.Fields("Name").value)
    Me.TxtNameE.text = IIf(IsNull(RsSavRec.Fields("NameE").value), "", RsSavRec.Fields("NameE").value)
    Me.DcbTypePrstg.ListIndex = IIf(IsNull(RsSavRec.Fields("TypePrstg").value), -1, RsSavRec.Fields("TypePrstg").value)
    Me.TxtPercenValue.text = IIf(IsNull(RsSavRec.Fields("PercenValue").value), 0, RsSavRec.Fields("PercenValue").value) ': ProgressBar1.value = 90
    Me.DcboEmpName.BoundText = IIf(IsNull(RsSavRec.Fields("EmpID").value), 0, RsSavRec.Fields("EmpID").value)
    Me.DCboUserName.BoundText = IIf(IsNull(RsSavRec.Fields("UserID").value), 0, RsSavRec.Fields("UserID").value)  ': ProgressBar1.value = 10
    TxtSquare.text = IIf(IsNull(RsSavRec.Fields("Square").value), "", RsSavRec.Fields("Square").value)
    TxtSquareCode.text = IIf(IsNull(RsSavRec.Fields("SquareCode").value), "", RsSavRec.Fields("SquareCode").value)
    TxtProjectCode.text = IIf(IsNull(RsSavRec.Fields("ProjectCode").value), "", RsSavRec.Fields("ProjectCode").value)
    Txt_path_photo.text = IIf(IsNull(RsSavRec.Fields("path_photo").value), "", RsSavRec.Fields("path_photo").value)
     Txt_Path_General_photo.text = IIf(IsNull(RsSavRec.Fields("Path_General_photo").value), "", RsSavRec.Fields("Path_General_photo").value)
    
    
          Dim Str_Path As String
             Str_Path = Txt_Path_General_photo.text
                DBPix201.ImageClear
                If Str_Path = "" Then GoTo xl:
             If Dir(Str_Path) <> "" Then
                DBPix201.ImageLoadFile (Txt_Path_General_photo.text)
              Else
            DBPix201.ImageClear
            End If
xl:
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
      If dcBranch.text = "" And val(dcBranch.BoundText) = 0 Then
        If SystemOptions.UserInterface = ArabicInterface Then
            MsgBox "⁄›Ê« ...«·—Ã«¡ ≈Œ Ì«— «·›—⁄", vbInformation + vbMsgBoxRtlReading + vbOKOnly + vbMsgBoxRight, App.Title
            Else
            MsgBox "Please Select Branch ", vbInformation + vbOKOnly + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
         End If
            dcBranch.SetFocus
            Exit Sub
     End If

     '      If DcbTypePrstg.Text = "" And val(DcbTypePrstg.ListIndex) = -1 Then
     '   If SystemOptions.UserInterface = ArabicInterface Then
     '       MsgBox "⁄›Ê« ...«·—Ã«¡ ≈Œ Ì«— ‰Ê⁄ ⁄„Ê·… «·«œ«—…", vbInformation + vbMsgBoxRtlReading + vbOKOnly + vbMsgBoxRight, App.title
     '       Else
     '       MsgBox "Please Select Commission ", vbInformation + vbOKOnly + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
     '    End If
     '       DcbTypePrstg.SetFocus
     '       Exit Sub
     'End If
     '         If val(TxtPercenValue.Text) = 0 Then
     '   If SystemOptions.UserInterface = ArabicInterface Then
     '       MsgBox "⁄›Ê« ...«·—Ã«¡ ≈œŒ«·  ﬁÌ„…  ⁄„Ê·… «·«œ«—…", vbInformation + vbMsgBoxRtlReading + vbOKOnly + vbMsgBoxRight, App.title
     '       Else
     '       MsgBox "Please Eneter Commission ", vbInformation + vbOKOnly + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
     '    End If
     '       TxtPercenValue.SetFocus
     '       Exit Sub
     'End If
If TxtProjectCode.text = "" Then
If SystemOptions.UserInterface = ArabicInterface Then
MsgBox "Ì—ÃÏ «œŒ«· ﬂÊœ «·„‘—Ê⁄ «Ê·«"
Else
MsgBox "Please enter the project code"
End If
TxtProjectCode.SetFocus
Exit Sub
End If
If TxtSquareCode.text = "" Then
If SystemOptions.UserInterface = ArabicInterface Then
MsgBox "Ì—ÃÏ «œŒ«· ﬂÊœ «·„—»⁄ «Ê·«"
Else
MsgBox "Please enter the square code"
End If
TxtSquareCode.SetFocus
Exit Sub
End If
           If DcboEmpName.text = "" And val(DcboEmpName.BoundText) = 0 Then
        If SystemOptions.UserInterface = ArabicInterface Then
            MsgBox "⁄›Ê« ...«·—Ã«¡ ≈Œ Ì«— „œÌ— «·„‘—Ê⁄ ", vbInformation + vbMsgBoxRtlReading + vbOKOnly + vbMsgBoxRight, App.Title
            Else
            MsgBox "Please Select Manager ", vbInformation + vbOKOnly + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
         End If
            DcboEmpName.SetFocus
            Exit Sub
     End If
     If SystemOptions.UserInterface = ArabicInterface Then
     If Me.TxtNameP.text = "" Then
     MsgBox "Ì—ÃÏ «œŒ«· «”„ «·„‘—Ê⁄"
     TxtNameP.SetFocus
     Exit Sub
     End If
     Else
    If Me.TxtNameE.text = "" Then
     MsgBox "Please Enter Name"
     TxtNameE.SetFocus
     Exit Sub
     End If
     End If
    ' -------------------------------------- txtmodflg type -------------------
    Select Case Me.TxtModFlg.text
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
    MsgBox "Â‰«ﬂ Œÿ√ „« ›Ì ≈œŒ«· «·»Ì«‰« ", vbOKOnly + vbMsgBoxRight, App.Title
    Else
    MsgBox "Sorry Error douring insert data", vbOKOnly + vbMsgBoxRight, App.Title
    End If
End Sub
' new recored
'++++++++++++++++++++++++++++++++++++
Public Sub AddNewRec()
  'On Error GoTo ErrTrap
    Dim StrRecID As String
    StrRecID = new_id("TblProjecInvestment", "ID", "")
    Me.TxtSerial1.text = StrRecID
    RsSavRec.AddNew
    RsSavRec.Fields("ID").value = IIf(StrRecID <> "", StrRecID, Null)
    FiLLRec
ErrTrap:
End Sub
Function GetIDModel(Optional Name As String) As Double
Dim sql As String
Dim Rs3 As ADODB.Recordset
Set Rs3 = New ADODB.Recordset
sql = "Select * from TblModelInves "
sql = sql & " WHERE     (Name LIKE N'%" & Name & "%') or (NameE LIKE N'%" & Name & "%')"
Rs3.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
If Rs3.RecordCount > 0 Then
GetIDModel = IIf(IsNull(Rs3("ID").value), 0, Rs3("ID").value)
Else
GetIDModel = SaveModel(Name)
End If
End Function
Function SaveModel(Optional Name As String) As Double
Dim Rs3 As ADODB.Recordset
Dim ID As Double
Dim StrSQL As String
    Set Rs3 = New ADODB.Recordset
    StrSQL = "SELECT  *  From TblModelInves where 1=-1 "
    Rs3.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText
    Rs3.AddNew
    ID = CStr(new_id("TblModelInves", "ID", "", True))
    Rs3("ID") = ID
    Rs3("Name") = Name
    Rs3("NameE") = Name
    Rs3("NoBooking") = 1
    Rs3.update
    SaveModel = ID
End Function
 Sub FullGridData()
 On Error GoTo ErrTrap
  Dim Rs1 As ADODB.Recordset
  Set Rs1 = New ADODB.Recordset
  Dim sql As String
    Grid.Clear flexClearScrollable, flexClearEverything
    Grid.rows = 1
  sql = " SELECT        dbo.TblProjecInvestmentDet.ID, dbo.TblProjecInvestmentDet.ProjInvID, dbo.TblProjecInvestmentDet.SquareNo, dbo.TblProjecInvestmentDet.BlockNo, dbo.TblProjecInvestmentDet.PartNo, "
  sql = sql + "                       dbo.TblProjecInvestmentDet.LandArea, dbo.TblProjecInvestmentDet.HouseArea, dbo.TblProjecInvestmentDet.MarketPrice, dbo.TblProjecInvestmentDet.MOHPrice, dbo.TblProjecInvestmentDet.BedroomsNo,"
  sql = sql + "                       dbo.TblProjecInvestmentDet.ConstStatus, dbo.TblProjecInvestmentDet.StatusID, dbo.TblProjecInvestmentDet.Remarks, dbo.TblProjecInvestmentDet.ModelName, dbo.TblProjecInvestmentDet.ModelID,"
  sql = sql + "                       dbo.TblModelInves.Name , dbo.TblModelInves.NameE ,dbo.TblProjecInvestmentDet.DeveloperCode ,dbo.TblProjecInvestmentDet.PartCode"
  sql = sql + " ,dbo.TblProjecInvestmentDet.Total ,dbo.TblProjecInvestmentDet.AddValue ,dbo.TblProjecInvestmentDet.ValueOffice ,dbo.TblProjecInvestmentDet.ModelName "
  sql = sql + " FROM            dbo.TblProjecInvestmentDet LEFT OUTER JOIN"
  sql = sql + "                       dbo.TblModelInves ON dbo.TblProjecInvestmentDet.ModelID = dbo.TblModelInves.ID"
  sql = sql + "      Where (dbo.TblProjecInvestmentDet.ProjInvID = " & val(TxtSerial1.text) & ")"
  
  Rs1.Open sql, Cn, adOpenKeyset, adLockOptimistic, adCmdText
    If Rs1.RecordCount > 0 Then
     Rs1.MoveFirst
     End If
     Dim i As Integer
     With Grid
                    For i = .FixedRows To Rs1.RecordCount
                   .rows = .FixedRows + Rs1.RecordCount
                   .TextMatrix(i, .ColIndex("Ser")) = i
                   .TextMatrix(i, .ColIndex("Total")) = IIf(IsNull(Rs1("Total").value), 0, Rs1("Total").value)
                   .TextMatrix(i, .ColIndex("AddValue")) = IIf(IsNull(Rs1("AddValue").value), 0, Rs1("AddValue").value)
                   .TextMatrix(i, .ColIndex("ValueOffice")) = IIf(IsNull(Rs1("ValueOffice").value), 0, Rs1("ValueOffice").value)
                   
                   .TextMatrix(i, .ColIndex("PartCode")) = IIf(IsNull(Rs1("PartCode").value), "", Rs1("PartCode").value)
                   .TextMatrix(i, .ColIndex("ID")) = IIf(IsNull(Rs1("ID").value), "", Rs1("ID").value)
                   .TextMatrix(i, .ColIndex("SquareNo")) = IIf(IsNull(Rs1("SquareNo").value), "", Rs1("SquareNo").value)
                   .TextMatrix(i, .ColIndex("BlockNo")) = IIf(IsNull(Rs1("BlockNo").value), "", Rs1("BlockNo").value)
                   .TextMatrix(i, .ColIndex("PartNo")) = IIf(IsNull(Rs1("PartNo").value), "", Rs1("PartNo").value)
                   .TextMatrix(i, .ColIndex("ModelID")) = IIf(IsNull(Rs1("ModelID").value), 0, Rs1("ModelID").value)
                   .TextMatrix(i, .ColIndex("LandArea")) = IIf(IsNull(Rs1("LandArea").value), 0, Rs1("LandArea").value)
                   .TextMatrix(i, .ColIndex("HouseArea")) = IIf(IsNull(Rs1("HouseArea").value), 0, Rs1("HouseArea").value)
                   .TextMatrix(i, .ColIndex("DeveloperCode")) = IIf(IsNull(Rs1("DeveloperCode").value), "", Rs1("DeveloperCode").value)
                   .TextMatrix(i, .ColIndex("MarketPrice")) = IIf(IsNull(Rs1("MarketPrice").value), 0, Rs1("MarketPrice").value)
                   .TextMatrix(i, .ColIndex("MOHPrice")) = IIf(IsNull(Rs1("MOHPrice").value), 0, Rs1("MOHPrice").value)
                   .TextMatrix(i, .ColIndex("BedroomsNo")) = IIf(IsNull(Rs1("BedroomsNo").value), "", Rs1("BedroomsNo").value)
                   .TextMatrix(i, .ColIndex("ConstStatus")) = IIf(IsNull(Rs1("ConstStatus").value), "", Rs1("ConstStatus").value)
                   .TextMatrix(i, .ColIndex("StatusID")) = IIf(IsNull(Rs1("StatusID").value), "", Rs1("StatusID").value)
                   .TextMatrix(i, .ColIndex("Remarks")) = IIf(IsNull(Rs1("Remarks").value), "", Rs1("Remarks").value)
                   If SystemOptions.UserInterface = ArabicInterface Then
                   .TextMatrix(i, .ColIndex("Name")) = IIf(IsNull(Rs1("Name").value), IIf(IsNull(Rs1("ModelName").value), "", Rs1("ModelName").value), Rs1("Name").value)
                   Else
                   .TextMatrix(i, .ColIndex("Name")) = IIf(IsNull(Rs1("NameE").value), IIf(IsNull(Rs1("ModelName").value), "", Rs1("ModelName").value), Rs1("NameE").value)
                   End If
                   Rs1.MoveNext
             Next i
        End With
        Exit Sub
ErrTrap:
    End Sub

Private Sub Grid_AfterEdit(ByVal row As Long, ByVal Col As Long)
Dim LngRow As Long
Dim StrAccountCode As String
 With Grid
   Select Case .ColKey(Col)
      Case "Name"
            StrAccountCode = .ComboData
                LngRow = .FindRow(StrAccountCode, .FixedRows, .ColIndex("ModelID"), False, True)
                .TextMatrix(row, .ColIndex("ModelID")) = StrAccountCode
               Case "ConstStatus"
            StrAccountCode = .ComboData
                LngRow = .FindRow(StrAccountCode, .FixedRows, .ColIndex("ConstStatus"), False, True)
                .TextMatrix(row, .ColIndex("ConstStatus")) = StrAccountCode
              If row = .rows - 1 Then
                .rows = .rows + 1
            End If
   End Select
 End With
 ReLineGrid
End Sub

Private Sub Grid_BeforeEdit(ByVal row As Long, ByVal Col As Long, Cancel As Boolean)
With Grid
Select Case .ColKey(Col)
Case "SquareNo"
.ComboList = ""
Case "BlockNo"
.ComboList = ""
Case "PartNo"
.ComboList = ""
Case "PartCode"
Cancel = True
Case "LandArea"
.ComboList = ""
Case "HouseArea"
.ComboList = ""
Case "DeveloperCode"
.ComboList = ""
Case "MarketPrice"
.ComboList = ""
Case "MOHPrice"
.ComboList = ""
Case "BedroomsNo"
.ComboList = ""
Case "Remarks"
.ComboList = ""
Case "Remarks"
.ComboList = ""
Case "ValueOffice"
.ComboList = ""
Case "AddValue"
.ComboList = ""
Case "Total"
Cancel = True

End Select
End With
End Sub

Private Sub Grid_StartEdit(ByVal row As Long, ByVal Col As Long, Cancel As Boolean)
Dim sql As String
Dim rs2 As ADODB.Recordset
Dim StrComboList As String
With Grid
Select Case .ColKey(Col)
Case "Name"
Set rs2 = New ADODB.Recordset
If SystemOptions.UserInterface = ArabicInterface Then
sql = "select id ,name from TblModelInves"
Else
sql = "select id ,namee from TblModelInves"
End If
rs2.Open sql, Cn, adOpenKeyset, adLockOptimistic, adCmdText
          If SystemOptions.UserInterface = ArabicInterface Then
                    StrComboList = .BuildComboList(rs2, "name", "id")
                Else
                    StrComboList = .BuildComboList(rs2, "namee", "id")
                End If
       
                If StrComboList <> "" Then
                    StrComboList = "|" & StrComboList
                End If
                 .ComboList = StrComboList
End Select
End With
End Sub
Private Sub RemoveGridRow()

    With Me.Grid

        If .row <= 0 Then Exit Sub
        .RemoveItem .row
    End With

    ReLineGrid
End Sub
Private Sub ISButton3_Click()
On Error Resume Next
Dim astrSplit2tems2() As String
If Me.TxtModFlg.text <> "R" Then
   Grid.Clear flexClearScrollable, flexClearEverything
      Grid.rows = 1
If txtFile.text = "" Then MsgBox "Õœœ «·„·› «Ê·«": Exit Sub
Dim ExcelObj As Object
    Dim ExcelBook As Object
    Dim ExcelSheet As Object
    Dim i As Long
    Dim EmpID As Integer
'Dim SquareNo As String
Dim BlockNo As String
Dim PartNo As String
Dim ModelID As String
Dim LandArea As String
Dim HouseArea As String
Dim DeveloperCode As String
Dim MarketPrice As String
Dim MOHPrice As String
Dim BedroomsNo As String
Dim ConstStatus As String
Dim StatusId As String
Dim Remarks As String
Dim Name As String
Dim AddValue As String
Dim ValueOffice As String
Dim ModelID1 As Double
Dim PartCode As String
    Set ExcelObj = CreateObject("Excel.Application")
    Set ExcelSheet = CreateObject("Excel.Sheet")
    ExcelObj.Workbooks.Open txtFile.text   ' App.Path & "\TrialBalance.xls"
DoEvents
    Set ExcelBook = ExcelObj.Workbooks(1)
    Set ExcelSheet = ExcelBook.Worksheets(1)
 
    With ExcelSheet
    i = 2
    Do Until .Cells(i, 3) & "" = ""
'    SquareNo = .Cells(i, 1)
    PartCode = .Cells(i, 1)
    Name = .Cells(i, 2) & "  " & .Cells(i, 5)      '‰„Ê–Ã &

    BlockNo = .Cells(i, 3)
    PartNo = .Cells(i, 4)
    ModelID = .Cells(i, 23)
    
    LandArea = .Cells(i, 25)
    LandArea = Round(val(LandArea), 0)
    
    HouseArea = .Cells(i, 26)
    HouseArea = Round(val(HouseArea), 0)
    
    BedroomsNo = .Cells(i, 27)
    MOHPrice = .Cells(i, 33)
    MOHPrice = Round(val(MOHPrice), 0)
  '  DeveloperCode = .cells(i, 6)
  '  MarketPrice = .cells(i, 7)
    
    
    ConstStatus = "NA"  '.cells(i, 8)
    'StatusId = .cells(i, 11)
    AddValue = 0 '  .cells(i, 9)
    ValueOffice = 0 ' .cells(i, 10)
   ' total = .cells(i, 11)
   
   AddValue = MOHPrice
ValueOffice = MOHPrice
 With Grid
        .rows = .rows + 1
        .TextMatrix(i - 1, .ColIndex("AddValue")) = AddValue
        .TextMatrix(i - 1, .ColIndex("ValueOffice")) = ValueOffice
         ModelID1 = GetIDModel(Trim(Name))
        .TextMatrix(i - 1, .ColIndex("ModelID")) = ModelID1
        .TextMatrix(i - 1, .ColIndex("BlockNo")) = BlockNo
        .TextMatrix(i - 1, .ColIndex("PartNo")) = PartNo
        .TextMatrix(i - 1, .ColIndex("Name")) = Name
        .TextMatrix(i - 1, .ColIndex("LandArea")) = LandArea
        .TextMatrix(i - 1, .ColIndex("HouseArea")) = HouseArea
        .TextMatrix(i - 1, .ColIndex("PartCode")) = PartCode
        
        '.TextMatrix(i - 1, .ColIndex("DeveloperCode")) = DeveloperCode
        '.TextMatrix(i - 1, .ColIndex("MarketPrice")) = MarketPrice
        .TextMatrix(i - 1, .ColIndex("MOHPrice")) = MOHPrice
        .TextMatrix(i - 1, .ColIndex("BedroomsNo")) = BedroomsNo
        If ConstStatus = " Õ  «·≈‰‘«¡" Then
        .TextMatrix(i - 1, .ColIndex("ConstStatus")) = 2
        ElseIf ConstStatus = " „ «·»‰«¡" Then
        .TextMatrix(i - 1, .ColIndex("ConstStatus")) = 3
        Else
        .TextMatrix(i - 1, .ColIndex("ConstStatus")) = 1
        End If
       ' .TextMatrix(i - 1, .ColIndex("StatusID")) = 1
       ' .TextMatrix(i - 1, .ColIndex("Remarks")) = Remarks
      ' Grid_AfterEdit i - 1, .ColIndex("ConstStatus")
 End With

 If .Cells(i, 1) & "" = "" Then Exit Sub
        i = i + 1
    Loop

    End With
     ReLineGrid
Grid.SetFocus
       ExcelObj.Workbooks.Close

    Set ExcelSheet = Nothing
    Set ExcelBook = Nothing
    Set ExcelObj = Nothing
 End If
End Sub

Private Sub ISButton4_Click()
If TxtProjectCode.text = "" Then
If SystemOptions.UserInterface = ArabicInterface Then
MsgBox "Ì—ÃÏ «œŒ«· ﬂÊœ «·„‘—Ê⁄ «Ê·«"
Else
MsgBox "Please enter the project code"
End If
TxtProjectCode.SetFocus
Exit Sub
End If
If TxtSquareCode.text = "" Then
If SystemOptions.UserInterface = ArabicInterface Then
MsgBox "Ì—ÃÏ «œŒ«· ﬂÊœ «·„—»⁄ «Ê·«"
Else
MsgBox "Please enter the square code"
End If
TxtSquareCode.SetFocus
Exit Sub
End If
      Grid.Clear flexClearScrollable, flexClearEverything
      Grid.rows = 1
CD1.ShowOpen
txtFile.text = CD1.FileName

End Sub

Private Sub RdTyp_Click(index As Integer)

ISButton4.Enabled = False
ISButton3.Enabled = False
If RdTyp(0).value = True Then

Else
ISButton4.Enabled = True
ISButton3.Enabled = True
End If
End Sub

Private Sub TxtPercenValue_KeyPress(KeyAscii As Integer)
KeyAscii = KeyAscii_Num(KeyAscii, Me.TxtPercenValue.text, 0)
End Sub

Private Sub TxtSearchCode_KeyPress(KeyAscii As Integer)
Dim EmpID As Integer

    If KeyAscii = vbKeyReturn Then
        GetEmployeeIDFromCode TxtSearchCode.text, EmpID
        DcboEmpName.BoundText = EmpID
    End If
End Sub

' change id search
Private Sub TxtSerial1_Change()
    Dim TxtMod As String
    TxtMod = TxtModFlg.text
    TxtModFlg.text = ""
    TxtModFlg = TxtMod
End Sub
' search for select id
Public Function FindRec(ByVal RecId As Long)
    On Error GoTo ErrTrap
    RsSavRec.Find "ID=" & RecId, , adSearchForward, 1
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
    FindRec val(TxtSerial1.text)
    Me.TxtModFlg.text = "R"
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
     If TxtSerial1.text = "" Then
       If SystemOptions.UserInterface = EnglishInterface Then
                X = MsgBox(" Nothing To Delet ", vbInformation + vbMsgBoxRtlReading + vbOKOnly + vbMsgBoxRight, App.Title)
               Else
                X = MsgBox("⁄›Ê« ...·« ÌÊÃœ »Ì«‰«  ··Õ–›", vbInformation + vbMsgBoxRtlReading + vbOKOnly + vbMsgBoxRight, App.Title)
       End If
               Else
      Dim StrSQL As String
                RsSavRec.Find "ID=" & val(TxtSerial1.text), , adSearchForward, 1
                  StrSQL = "Delete From TblProjecInvestmentDet Where ProjInvID =" & val(TxtSerial1.text) & ""
                  Cn.Execute StrSQL, , adExecuteNoRecords
                                          RsSavRec.delete
            Grid.Clear flexClearScrollable, flexClearEverything
            Grid.rows = 1

                 If SystemOptions.UserInterface = EnglishInterface Then
                X = MsgBox(" Delete  Successfully ", vbInformation + vbMsgBoxRtlReading + vbOKOnly + vbMsgBoxRight, App.Title)
               Else
                X = MsgBox(" „  ⁄„·Ì… «·Õ–› »‰Ã«Õ", vbInformation + vbMsgBoxRtlReading + vbOKOnly + vbMsgBoxRight, App.Title)
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
            MsgBox StrMSG, vbOKOnly + vbMsgBoxRight, App.Title
           'Cn.Errors.Clear
    End Select

End Sub
' exit without save sub
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
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
    If TxtModFlg.text = "N" Then
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
        
    ElseIf TxtModFlg.text = "R" Then
     XPDtbTrans.Enabled = False
        btnModify.Enabled = False
        btnDelete.Enabled = False
        If TxtSerial1.text <> "" Then
            btnModify.Enabled = True
            btnDelete.Enabled = True
            Command2.Enabled = False
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
   ElseIf TxtModFlg.text = "E" Then
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
    If Me.TxtModFlg.text = "N" Then
        FindRec val(TxtSerial1.text)
        Me.TxtModFlg.text = "R"
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
            MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading + vbOKOnly, App.Title
            RsSavRec.Requery
            Resume BegnieWork
    End Select
End Sub
Private Sub BtnLast_Click()
    On Error GoTo ErrTrap
    Dim Msg As String
    If Me.TxtModFlg.text = "N" Then
        FindRec val(TxtSerial1.text)
        Me.TxtModFlg.text = "R"
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
            MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading + vbOKOnly, App.Title
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
    If TxtSerial1.text <> "" Then
        TxtModFlg = "E"
            Grid.rows = Grid.rows + 1
        Me.DCboUserName.BoundText = user_id
        Me.dcBranch.SetFocus
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
            MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading + vbOKOnly, App.Title
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
RdTyp(0).value = True
RdTyp_Click (0)
    TxtModFlg.text = "N"
    Grid.Clear flexClearScrollable, flexClearEverything
      Grid.rows = 2
    Me.DCboUserName.BoundText = user_id
    Me.dcBranch.BoundText = Current_branch
    dcBranch.SetFocus
ErrTrap:
End Sub
Private Sub BtnNext_Click()
    On Error GoTo ErrTrap
    Dim Msg As String
    If Me.TxtModFlg.text = "N" Then
        FindRec val(TxtSerial1.text)
        Me.TxtModFlg.text = "R"
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
            MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading + vbOKOnly, App.Title
            RsSavRec.Requery
            Resume BegnieWork
    End Select
End Sub
Private Sub BtnPrevious_Click()
    On Error GoTo ErrTrap
    Dim Msg As String
    If Me.TxtModFlg.text = "N" Then
        FindRec val(TxtSerial1.text)
        Me.TxtModFlg.text = "R"
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
            MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading + vbOKOnly, App.Title
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
        If Me.TxtModFlg.text = "R" Then
            btnNew_Click
        Else
            Sendkeys "{TAB}"
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
   Command2.Caption = "Planned"
   lbl(5).Caption = "Details"
       Me.Caption = "Project Data"
      Label1(2).Caption = "Project Data"
      Me.lbl(4).Caption = "ID"
      Me.lbl(2).Caption = "Date"
      lbl(7).Caption = "Branch"
      lbl(9).Caption = "Project Code"
      lbl(1).Caption = "Project"
      lbl(3).Caption = "Name English"
   lbl(15).Caption = "Project Manager"
   lbl(0).Caption = "Commission"
   lbl(6).Caption = "Square"
RdTyp(0).RightToLeft = False
RdTyp(1).RightToLeft = False
RdTyp(0).Caption = "Manual"
RdTyp(1).Caption = "From File"
ISButton4.Caption = "Select Path"
ISButton3.Caption = "Import"
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
   
  With Me.Grid
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
   My_SQL = "TblProjecInvestment"
    rs.Open My_SQL, Cn, adOpenKeyset, adLockOptimistic, adCmdTable
    If rs.RecordCount > 0 Then
        TxtSerial1.text = rs.RecordCount + 1
    Else
        TxtSerial1.text = 1
    End If
   rs.Close
ErrTrap:
End Sub
'+++++++++++++++++++++++++++++++++ end
