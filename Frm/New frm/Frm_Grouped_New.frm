VERSION 5.00
Object = "{C115893A-A3BF-43AF-B28D-69DB846077F3}#1.0#0"; "vsflex8u.ocx"
Object = "{0AFE7BE0-11B7-4A3E-978D-D4501E9A57FE}#1.0#0"; "c1sizer.ocx"
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{FE5DCFAD-BC1D-11D2-94CF-004005455FAA}#1.4#0"; "ImpulseButton.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form FrmPay_Garanty_Shipment3M 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   6975
   ClientLeft      =   -15
   ClientTop       =   330
   ClientWidth     =   7920
   Icon            =   "Frm_Grouped_New.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   6975
   ScaleWidth      =   7920
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   8
      Charset         =   178
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Begin C1SizerLibCtl.C1Tab C1Tab1 
      Height          =   6930
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   7965
      _cx             =   14049
      _cy             =   12224
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial (Arabic)"
         Size            =   9.75
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Enabled         =   -1  'True
      Appearance      =   2
      MousePointer    =   0
      Version         =   801
      BackColor       =   12648447
      ForeColor       =   128
      FrontTabColor   =   14871017
      BackTabColor    =   8454143
      TabOutlineColor =   -2147483632
      FrontTabForeColor=   -2147483630
      Caption         =   "0|1|2|3|4|5"
      Align           =   0
      CurrTab         =   5
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
      Begin C1SizerLibCtl.C1Elastic C1Elastic100 
         Height          =   6510
         Left            =   -9720
         TabIndex        =   1
         TabStop         =   0   'False
         Top             =   45
         Width           =   7875
         _cx             =   13891
         _cy             =   11483
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
         Begin VB.Frame Frm2 
            BackColor       =   &H00E2E9E9&
            BorderStyle     =   0  'None
            Enabled         =   0   'False
            Height          =   1470
            Left            =   270
            RightToLeft     =   -1  'True
            TabIndex        =   2
            Top             =   4080
            Width           =   7590
            Begin VB.TextBox TxtVacName 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00FFFFFF&
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   178
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   345
               Left            =   195
               MaxLength       =   50
               RightToLeft     =   -1  'True
               TabIndex        =   6
               Top             =   465
               Width           =   5580
            End
            Begin VB.TextBox TxtUnitID 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00FFFFFF&
               Enabled         =   0   'False
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   178
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   345
               Left            =   4470
               MaxLength       =   50
               RightToLeft     =   -1  'True
               TabIndex        =   5
               Top             =   105
               Width           =   1305
            End
            Begin VB.ComboBox CmbType 
               BackColor       =   &H80000018&
               Height          =   315
               ItemData        =   "Frm_Grouped_New.frx":000C
               Left            =   2280
               List            =   "Frm_Grouped_New.frx":001C
               RightToLeft     =   -1  'True
               Style           =   2  'Dropdown List
               TabIndex        =   4
               Top             =   2550
               Visible         =   0   'False
               Width           =   1005
            End
            Begin VB.TextBox TxtVacNamee 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00FFFFFF&
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   178
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   345
               Left            =   195
               MaxLength       =   50
               RightToLeft     =   -1  'True
               TabIndex        =   3
               Top             =   840
               Width           =   5580
            End
            Begin VB.Label Label1 
               Alignment       =   2  'Center
               BackColor       =   &H00E2E9E9&
               Caption         =   "«”„ «·ÊÕœ… ⁄—»Ì"
               Height          =   255
               Index           =   0
               Left            =   6030
               RightToLeft     =   -1  'True
               TabIndex        =   9
               Top             =   450
               Width           =   1410
            End
            Begin VB.Label Label1 
               Alignment       =   2  'Center
               BackColor       =   &H00E2E9E9&
               Caption         =   "„”·”·"
               Height          =   285
               Index           =   3
               Left            =   6285
               RightToLeft     =   -1  'True
               TabIndex        =   8
               Top             =   90
               Width           =   810
            End
            Begin VB.Label Label1 
               Alignment       =   2  'Center
               BackColor       =   &H00E2E9E9&
               Caption         =   "«”„ «·ÊÕœ… «‰Ã·Ì“Ì"
               Height          =   255
               Index           =   1
               Left            =   6000
               RightToLeft     =   -1  'True
               TabIndex        =   7
               Top             =   840
               Width           =   1530
            End
         End
         Begin C1SizerLibCtl.C1Elastic EleHeader 
            Height          =   675
            Left            =   0
            TabIndex        =   10
            TabStop         =   0   'False
            Top             =   0
            Width           =   7980
            _cx             =   14076
            _cy             =   1191
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial (Arabic)"
               Size            =   21.75
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
            Caption         =   ""
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
            Begin VB.TextBox TxtModFlg 
               Alignment       =   1  'Right Justify
               Appearance      =   0  'Flat
               BackColor       =   &H0000FF00&
               Enabled         =   0   'False
               Height          =   285
               Left            =   2700
               RightToLeft     =   -1  'True
               TabIndex        =   11
               Top             =   120
               Visible         =   0   'False
               Width           =   945
            End
            Begin ImpulseButton.ISButton btnLast 
               Height          =   345
               Left            =   390
               TabIndex        =   12
               Top             =   150
               Width           =   405
               _ExtentX        =   714
               _ExtentY        =   609
               ButtonStyle     =   1
               ButtonPositionImage=   4
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
               ButtonImage     =   "Frm_Grouped_New.frx":0035
               ColorButton     =   16777215
               DrawFocusRectangle=   0   'False
               DisabledImageExtraction=   0
               DisabledImageStyle=   1
            End
            Begin ImpulseButton.ISButton btnNext 
               Height          =   345
               Left            =   855
               TabIndex        =   13
               Top             =   150
               Width           =   405
               _ExtentX        =   714
               _ExtentY        =   609
               ButtonStyle     =   1
               ButtonPositionImage=   4
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
               ButtonImage     =   "Frm_Grouped_New.frx":03CF
               ColorButton     =   16777215
               DrawFocusRectangle=   0   'False
               DisabledImageStyle=   1
            End
            Begin ImpulseButton.ISButton btnPrevious 
               Height          =   345
               Left            =   1305
               TabIndex        =   14
               Top             =   150
               Width           =   405
               _ExtentX        =   714
               _ExtentY        =   609
               ButtonStyle     =   1
               ButtonPositionImage=   4
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
               ButtonImage     =   "Frm_Grouped_New.frx":0769
               ColorButton     =   16777215
               DrawFocusRectangle=   0   'False
               DisabledImageStyle=   1
            End
            Begin ImpulseButton.ISButton btnFirst 
               Height          =   345
               Left            =   1770
               TabIndex        =   15
               Top             =   150
               Width           =   405
               _ExtentX        =   714
               _ExtentY        =   609
               ButtonStyle     =   1
               ButtonPositionImage=   4
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
               ButtonImage     =   "Frm_Grouped_New.frx":0B03
               ColorButton     =   16777215
               DrawFocusRectangle=   0   'False
               DisabledImageStyle=   1
            End
            Begin VB.Label Label1 
               Alignment       =   2  'Center
               BackStyle       =   0  'Transparent
               Caption         =   "ÊÕœ«  «·√’‰«›"
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
               Left            =   4680
               TabIndex        =   35
               Top             =   120
               Width           =   3015
            End
         End
         Begin C1SizerLibCtl.C1Elastic EltCont 
            Height          =   750
            Left            =   510
            TabIndex        =   16
            TabStop         =   0   'False
            Top             =   5580
            Width           =   6870
            _cx             =   12118
            _cy             =   1323
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
            AutoSizeChildren=   0
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
               Height          =   420
               Left            =   5805
               TabIndex        =   17
               Top             =   375
               Width           =   720
               _ExtentX        =   1270
               _ExtentY        =   741
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
               ButtonImage     =   "Frm_Grouped_New.frx":0E9D
               ColorButton     =   14871017
               DrawFocusRectangle=   0   'False
            End
            Begin ImpulseButton.ISButton btnSave 
               Height          =   420
               Left            =   4290
               TabIndex        =   18
               Top             =   375
               Width           =   720
               _ExtentX        =   1270
               _ExtentY        =   741
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
               ButtonImage     =   "Frm_Grouped_New.frx":1237
               ColorButton     =   14871017
               DrawFocusRectangle=   0   'False
            End
            Begin ImpulseButton.ISButton btnModify 
               Height          =   420
               Left            =   4935
               TabIndex        =   19
               Top             =   375
               Width           =   825
               _ExtentX        =   1455
               _ExtentY        =   741
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
               ButtonImage     =   "Frm_Grouped_New.frx":15D1
               ColorButton     =   14871017
               DrawFocusRectangle=   0   'False
            End
            Begin ImpulseButton.ISButton BtnUndo 
               Height          =   420
               Left            =   3405
               TabIndex        =   20
               Top             =   375
               Width           =   840
               _ExtentX        =   1482
               _ExtentY        =   741
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
               ButtonImage     =   "Frm_Grouped_New.frx":196B
               ColorButton     =   14871017
               DrawFocusRectangle=   0   'False
            End
            Begin ImpulseButton.ISButton btnDelete 
               Height          =   420
               Left            =   990
               TabIndex        =   21
               Top             =   375
               Width           =   690
               _ExtentX        =   1217
               _ExtentY        =   741
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
               ButtonImage     =   "Frm_Grouped_New.frx":1D05
               ColorButton     =   14871017
               DrawFocusRectangle=   0   'False
            End
            Begin ImpulseButton.ISButton btnQuery 
               Height          =   330
               Left            =   2640
               TabIndex        =   22
               TabStop         =   0   'False
               ToolTipText     =   "(Ctrl+F)  ··»ÕÀ ≈÷€ÿ Â–« «·„› «Õ √Ê ≈÷€ÿ "
               Top             =   450
               Width           =   720
               _ExtentX        =   1270
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
               ButtonImage     =   "Frm_Grouped_New.frx":229F
               ColorButton     =   14871017
               DrawFocusRectangle=   0   'False
            End
            Begin ImpulseButton.ISButton BtnUpdate 
               Height          =   330
               Left            =   5925
               TabIndex        =   23
               TabStop         =   0   'False
               ToolTipText     =   " ÕœÌÀ ﬁ«⁄œ… «·»Ì«‰« "
               Top             =   -15
               Visible         =   0   'False
               Width           =   840
               _ExtentX        =   1482
               _ExtentY        =   582
               ButtonStyle     =   1
               ButtonPositionImage=   1
               Caption         =   " ÕœÌÀ"
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
               ButtonImage     =   "Frm_Grouped_New.frx":2639
               ColorButton     =   14871017
               DrawFocusRectangle=   0   'False
            End
            Begin ImpulseButton.ISButton btnCancel 
               Height          =   420
               Left            =   75
               TabIndex        =   24
               Top             =   375
               Width           =   870
               _ExtentX        =   1535
               _ExtentY        =   741
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
               ButtonImage     =   "Frm_Grouped_New.frx":29D3
               ColorButton     =   14871017
               DrawFocusRectangle=   0   'False
               DisabledImageStyle=   1
            End
            Begin ImpulseButton.ISButton ISButton1 
               Height          =   285
               Left            =   1800
               TabIndex        =   25
               TabStop         =   0   'False
               Top             =   480
               Width           =   765
               _ExtentX        =   1349
               _ExtentY        =   503
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
               ButtonImage     =   "Frm_Grouped_New.frx":2D6D
               ColorButton     =   14871017
               DrawFocusRectangle=   0   'False
               DisabledImageStyle=   1
            End
            Begin VB.Label Label2 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "«·”Ã· «·Õ«·Ì:"
               Height          =   210
               Index           =   0
               Left            =   3555
               RightToLeft     =   -1  'True
               TabIndex        =   29
               Top             =   15
               Width           =   975
            End
            Begin VB.Label Label2 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "⁄œœ «·”Ã·« :"
               Height          =   210
               Index           =   1
               Left            =   960
               RightToLeft     =   -1  'True
               TabIndex        =   28
               Top             =   15
               Width           =   975
            End
            Begin VB.Label LabCurrRec 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Height          =   210
               Left            =   2745
               RightToLeft     =   -1  'True
               TabIndex        =   27
               Top             =   45
               Width           =   675
            End
            Begin VB.Label LabCountRec 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Height          =   210
               Left            =   270
               RightToLeft     =   -1  'True
               TabIndex        =   26
               Top             =   15
               Width           =   540
            End
         End
         Begin VSFlex8Ctl.VSFlexGrid Grid 
            Height          =   3270
            Left            =   120
            TabIndex        =   30
            Top             =   795
            Width           =   7650
            _cx             =   13494
            _cy             =   5768
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
            Rows            =   50
            Cols            =   4
            FixedRows       =   1
            FixedCols       =   1
            RowHeightMin    =   320
            RowHeightMax    =   0
            ColWidthMin     =   0
            ColWidthMax     =   0
            ExtendLastCol   =   -1  'True
            FormatString    =   $"Frm_Grouped_New.frx":3107
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
      Begin C1SizerLibCtl.C1Elastic C1Elastic1 
         Height          =   6510
         Left            =   -9420
         TabIndex        =   31
         TabStop         =   0   'False
         Top             =   45
         Width           =   7875
         _cx             =   13891
         _cy             =   11483
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
         Begin VB.TextBox XPMTxtRemark 
            Alignment       =   1  'Right Justify
            Height          =   2955
            Left            =   465
            MultiLine       =   -1  'True
            RightToLeft     =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   44
            Top             =   2070
            Width           =   6030
         End
         Begin VB.TextBox XPTxtBoxName 
            Alignment       =   1  'Right Justify
            Height          =   315
            Left            =   465
            MaxLength       =   50
            RightToLeft     =   -1  'True
            TabIndex        =   43
            Top             =   1455
            Width           =   6030
         End
         Begin VB.TextBox XPTxtBoxID 
            Alignment       =   1  'Right Justify
            Height          =   315
            Left            =   4485
            Locked          =   -1  'True
            MaxLength       =   10
            RightToLeft     =   -1  'True
            TabIndex        =   42
            Top             =   960
            Width           =   2010
         End
         Begin C1SizerLibCtl.C1Elastic C1Elastic5 
            Height          =   705
            Left            =   0
            TabIndex        =   36
            TabStop         =   0   'False
            Top             =   0
            Width           =   7875
            _cx             =   13891
            _cy             =   1244
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial (Arabic)"
               Size            =   18
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
            Caption         =   ""
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
            Begin VB.TextBox TxtModFlg1 
               Alignment       =   1  'Right Justify
               Height          =   345
               Left            =   2610
               RightToLeft     =   -1  'True
               TabIndex        =   37
               Top             =   180
               Visible         =   0   'False
               Width           =   855
            End
            Begin ImpulseButton.ISButton XPBtnMove 
               Height          =   345
               Index           =   0
               Left            =   1155
               TabIndex        =   38
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
               ButtonImage     =   "Frm_Grouped_New.frx":31B5
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
               TabIndex        =   39
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
               ButtonImage     =   "Frm_Grouped_New.frx":354F
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
               TabIndex        =   40
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
               ButtonImage     =   "Frm_Grouped_New.frx":38E9
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
               TabIndex        =   41
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
               ButtonImage     =   "Frm_Grouped_New.frx":3C83
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
            Begin VB.Label Label1 
               Alignment       =   2  'Center
               BackStyle       =   0  'Transparent
               Caption         =   "»Ì«‰«  √·Ê«‰ «·√’‰«›"
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
               Index           =   4
               Left            =   4080
               TabIndex        =   59
               Top             =   120
               Width           =   3735
            End
         End
         Begin C1SizerLibCtl.C1Elastic C1Elastic6 
            Height          =   990
            Left            =   465
            TabIndex        =   48
            TabStop         =   0   'False
            Top             =   5400
            Width           =   6855
            _cx             =   12091
            _cy             =   1746
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
            AutoSizeChildren=   0
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
               Height          =   420
               Index           =   0
               Left            =   5805
               TabIndex        =   49
               Top             =   495
               Width           =   720
               _ExtentX        =   1270
               _ExtentY        =   741
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
               ButtonImage     =   "Frm_Grouped_New.frx":401D
               ColorButton     =   14871017
               DrawFocusRectangle=   0   'False
            End
            Begin ImpulseButton.ISButton Cmd 
               Height          =   420
               Index           =   2
               Left            =   3930
               TabIndex        =   50
               Top             =   495
               Width           =   720
               _ExtentX        =   1270
               _ExtentY        =   741
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
               ButtonImage     =   "Frm_Grouped_New.frx":43B7
               ColorButton     =   14871017
               DrawFocusRectangle=   0   'False
            End
            Begin ImpulseButton.ISButton Cmd 
               Height          =   420
               Index           =   1
               Left            =   4815
               TabIndex        =   51
               Top             =   495
               Width           =   825
               _ExtentX        =   1455
               _ExtentY        =   741
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
               ButtonImage     =   "Frm_Grouped_New.frx":4751
               ColorButton     =   14871017
               DrawFocusRectangle=   0   'False
            End
            Begin ImpulseButton.ISButton Cmd 
               Height          =   420
               Index           =   3
               Left            =   2805
               TabIndex        =   52
               Top             =   495
               Width           =   840
               _ExtentX        =   1482
               _ExtentY        =   741
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
               ButtonImage     =   "Frm_Grouped_New.frx":4AEB
               ColorButton     =   14871017
               DrawFocusRectangle=   0   'False
            End
            Begin ImpulseButton.ISButton Cmd 
               Height          =   420
               Index           =   4
               Left            =   1830
               TabIndex        =   53
               Top             =   495
               Width           =   690
               _ExtentX        =   1217
               _ExtentY        =   741
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
               ButtonImage     =   "Frm_Grouped_New.frx":4E85
               ColorButton     =   14871017
               DrawFocusRectangle=   0   'False
            End
            Begin ImpulseButton.ISButton Cmd 
               Height          =   420
               Index           =   6
               Left            =   75
               TabIndex        =   54
               Top             =   495
               Width           =   870
               _ExtentX        =   1535
               _ExtentY        =   741
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
               ButtonImage     =   "Frm_Grouped_New.frx":541F
               ColorButton     =   14871017
               DrawFocusRectangle=   0   'False
               DisabledImageStyle=   1
            End
            Begin VB.Label XPTxtCount 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Height          =   210
               Left            =   270
               RightToLeft     =   -1  'True
               TabIndex        =   58
               Top             =   135
               Width           =   540
            End
            Begin VB.Label XPTxtCurrent 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Height          =   210
               Left            =   2745
               RightToLeft     =   -1  'True
               TabIndex        =   57
               Top             =   165
               Width           =   675
            End
            Begin VB.Label Label2 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "⁄œœ «·”Ã·« :"
               Height          =   210
               Index           =   3
               Left            =   960
               RightToLeft     =   -1  'True
               TabIndex        =   56
               Top             =   135
               Width           =   975
            End
            Begin VB.Label Label2 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "«·”Ã· «·Õ«·Ì:"
               Height          =   210
               Index           =   2
               Left            =   3555
               RightToLeft     =   -1  'True
               TabIndex        =   55
               Top             =   135
               Width           =   975
            End
         End
         Begin VB.Label Lbl 
            Alignment       =   2  'Center
            BackColor       =   &H00E2E9E9&
            Caption         =   "ﬂÊœ «··Ê‰"
            Height          =   285
            Index           =   0
            Left            =   6690
            RightToLeft     =   -1  'True
            TabIndex        =   47
            Top             =   975
            Width           =   1005
         End
         Begin VB.Label Lbl 
            Alignment       =   2  'Center
            BackColor       =   &H00E2E9E9&
            Caption         =   "„·«ÕŸ« "
            Height          =   315
            Index           =   1
            Left            =   6690
            RightToLeft     =   -1  'True
            TabIndex        =   46
            Top             =   2670
            Width           =   1005
         End
         Begin VB.Label Lbl 
            Alignment       =   2  'Center
            BackColor       =   &H00E2E9E9&
            Caption         =   "«”„ «··Ê‰"
            Height          =   315
            Index           =   3
            Left            =   6690
            RightToLeft     =   -1  'True
            TabIndex        =   45
            Top             =   1455
            Width           =   1005
         End
      End
      Begin C1SizerLibCtl.C1Elastic C1Elastic2 
         Height          =   6510
         Left            =   -9120
         TabIndex        =   32
         TabStop         =   0   'False
         Top             =   45
         Width           =   7875
         _cx             =   13891
         _cy             =   11483
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
         Begin VB.TextBox XPMTxtRemark2 
            Alignment       =   1  'Right Justify
            Height          =   2835
            Left            =   465
            MultiLine       =   -1  'True
            RightToLeft     =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   62
            Top             =   2220
            Width           =   5940
         End
         Begin VB.TextBox XPTxtBoxName2 
            Alignment       =   1  'Right Justify
            Height          =   315
            Left            =   465
            MaxLength       =   50
            RightToLeft     =   -1  'True
            TabIndex        =   61
            Top             =   1605
            Width           =   5940
         End
         Begin VB.TextBox XPTxtBoxID2 
            Alignment       =   1  'Right Justify
            Height          =   315
            Left            =   3750
            Locked          =   -1  'True
            MaxLength       =   10
            RightToLeft     =   -1  'True
            TabIndex        =   60
            Top             =   990
            Width           =   2655
         End
         Begin C1SizerLibCtl.C1Elastic C1Elastic7 
            Height          =   705
            Left            =   0
            TabIndex        =   63
            TabStop         =   0   'False
            Top             =   0
            Width           =   7875
            _cx             =   13891
            _cy             =   1244
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial (Arabic)"
               Size            =   18
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
            Caption         =   ""
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
            Begin VB.TextBox TxtModFlg2 
               Alignment       =   1  'Right Justify
               Height          =   345
               Left            =   3090
               RightToLeft     =   -1  'True
               TabIndex        =   64
               Top             =   180
               Visible         =   0   'False
               Width           =   855
            End
            Begin ImpulseButton.ISButton XPBtnMove2 
               Height          =   345
               Index           =   0
               Left            =   1155
               TabIndex        =   65
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
               ButtonImage     =   "Frm_Grouped_New.frx":57B9
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
            Begin ImpulseButton.ISButton XPBtnMove2 
               Height          =   345
               Index           =   2
               Left            =   90
               TabIndex        =   66
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
               ButtonImage     =   "Frm_Grouped_New.frx":5B53
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
            Begin ImpulseButton.ISButton XPBtnMove2 
               Height          =   345
               Index           =   1
               Left            =   1680
               TabIndex        =   67
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
               ButtonImage     =   "Frm_Grouped_New.frx":5EED
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
            Begin ImpulseButton.ISButton XPBtnMove2 
               Height          =   345
               Index           =   3
               Left            =   615
               TabIndex        =   68
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
               ButtonImage     =   "Frm_Grouped_New.frx":6287
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
            Begin VB.Label Label1 
               Alignment       =   2  'Center
               BackStyle       =   0  'Transparent
               Caption         =   "»Ì«‰«  „ﬁ«”«  «·√’‰«›"
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
               Index           =   5
               Left            =   4080
               TabIndex        =   83
               Top             =   120
               Width           =   3735
            End
         End
         Begin C1SizerLibCtl.C1Elastic C1Elastic8 
            Height          =   990
            Left            =   465
            TabIndex        =   72
            TabStop         =   0   'False
            Top             =   5400
            Width           =   6855
            _cx             =   12091
            _cy             =   1746
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
            AutoSizeChildren=   0
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
            Begin ImpulseButton.ISButton Cmd2 
               Height          =   420
               Index           =   0
               Left            =   5805
               TabIndex        =   73
               Top             =   495
               Width           =   720
               _ExtentX        =   1270
               _ExtentY        =   741
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
               ButtonImage     =   "Frm_Grouped_New.frx":6621
               ColorButton     =   14871017
               DrawFocusRectangle=   0   'False
            End
            Begin ImpulseButton.ISButton Cmd2 
               Height          =   420
               Index           =   2
               Left            =   3930
               TabIndex        =   74
               Top             =   495
               Width           =   720
               _ExtentX        =   1270
               _ExtentY        =   741
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
               ButtonImage     =   "Frm_Grouped_New.frx":69BB
               ColorButton     =   14871017
               DrawFocusRectangle=   0   'False
            End
            Begin ImpulseButton.ISButton Cmd2 
               Height          =   420
               Index           =   1
               Left            =   4815
               TabIndex        =   75
               Top             =   495
               Width           =   825
               _ExtentX        =   1455
               _ExtentY        =   741
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
               ButtonImage     =   "Frm_Grouped_New.frx":6D55
               ColorButton     =   14871017
               DrawFocusRectangle=   0   'False
            End
            Begin ImpulseButton.ISButton Cmd2 
               Height          =   420
               Index           =   3
               Left            =   2805
               TabIndex        =   76
               Top             =   495
               Width           =   840
               _ExtentX        =   1482
               _ExtentY        =   741
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
               ButtonImage     =   "Frm_Grouped_New.frx":70EF
               ColorButton     =   14871017
               DrawFocusRectangle=   0   'False
            End
            Begin ImpulseButton.ISButton Cmd2 
               Height          =   420
               Index           =   4
               Left            =   1830
               TabIndex        =   77
               Top             =   495
               Width           =   690
               _ExtentX        =   1217
               _ExtentY        =   741
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
               ButtonImage     =   "Frm_Grouped_New.frx":7489
               ColorButton     =   14871017
               DrawFocusRectangle=   0   'False
            End
            Begin ImpulseButton.ISButton Cmd2 
               Height          =   420
               Index           =   6
               Left            =   75
               TabIndex        =   78
               Top             =   495
               Width           =   870
               _ExtentX        =   1535
               _ExtentY        =   741
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
               ButtonImage     =   "Frm_Grouped_New.frx":7A23
               ColorButton     =   14871017
               DrawFocusRectangle=   0   'False
               DisabledImageStyle=   1
            End
            Begin VB.Label Label2 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "«·”Ã· «·Õ«·Ì:"
               Height          =   210
               Index           =   5
               Left            =   3555
               RightToLeft     =   -1  'True
               TabIndex        =   82
               Top             =   135
               Width           =   975
            End
            Begin VB.Label Label2 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "⁄œœ «·”Ã·« :"
               Height          =   210
               Index           =   4
               Left            =   960
               RightToLeft     =   -1  'True
               TabIndex        =   81
               Top             =   135
               Width           =   975
            End
            Begin VB.Label XPTxtCurrent2 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Height          =   210
               Left            =   2745
               RightToLeft     =   -1  'True
               TabIndex        =   80
               Top             =   165
               Width           =   675
            End
            Begin VB.Label XPTxtCount2 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Height          =   210
               Left            =   270
               RightToLeft     =   -1  'True
               TabIndex        =   79
               Top             =   120
               Width           =   540
            End
         End
         Begin VB.Label Lbl 
            Alignment       =   2  'Center
            BackColor       =   &H00E2E9E9&
            Caption         =   "ﬂÊœ «·„ﬁ«”"
            Height          =   285
            Index           =   5
            Left            =   6405
            RightToLeft     =   -1  'True
            TabIndex        =   71
            Top             =   1005
            Width           =   1380
         End
         Begin VB.Label Lbl 
            Alignment       =   2  'Center
            BackColor       =   &H00E2E9E9&
            Caption         =   "„·«ÕŸ« "
            Height          =   195
            Index           =   4
            Left            =   6405
            RightToLeft     =   -1  'True
            TabIndex        =   70
            Top             =   2700
            Width           =   1380
         End
         Begin VB.Label Lbl 
            Alignment       =   2  'Center
            BackColor       =   &H00E2E9E9&
            Caption         =   "«”„ «·„ﬁ«”"
            Height          =   315
            Index           =   2
            Left            =   6495
            RightToLeft     =   -1  'True
            TabIndex        =   69
            Top             =   1605
            Width           =   1200
         End
      End
      Begin C1SizerLibCtl.C1Elastic C1Elastic3 
         Height          =   6510
         Left            =   -8820
         TabIndex        =   33
         TabStop         =   0   'False
         Top             =   45
         Width           =   7875
         _cx             =   13891
         _cy             =   11483
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
         Begin VB.TextBox XPMTxtRemark3 
            Alignment       =   1  'Right Justify
            Height          =   2235
            Left            =   645
            MultiLine       =   -1  'True
            RightToLeft     =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   104
            Top             =   2670
            Width           =   5580
         End
         Begin VB.TextBox XPTxtBoxName3 
            Alignment       =   1  'Right Justify
            Height          =   315
            Left            =   645
            MaxLength       =   50
            RightToLeft     =   -1  'True
            TabIndex        =   103
            Top             =   1695
            Width           =   5580
         End
         Begin VB.TextBox XPTxtBoxID3 
            Alignment       =   1  'Right Justify
            Height          =   315
            Left            =   4305
            Locked          =   -1  'True
            MaxLength       =   10
            RightToLeft     =   -1  'True
            TabIndex        =   102
            Top             =   1200
            Width           =   1920
         End
         Begin VB.TextBox XPTxtBoxNamee 
            Alignment       =   1  'Right Justify
            Height          =   315
            Left            =   645
            MaxLength       =   50
            RightToLeft     =   -1  'True
            TabIndex        =   101
            Top             =   2130
            Width           =   5580
         End
         Begin C1SizerLibCtl.C1Elastic C1Elastic9 
            Height          =   990
            Left            =   465
            TabIndex        =   84
            TabStop         =   0   'False
            Top             =   5400
            Width           =   6855
            _cx             =   12091
            _cy             =   1746
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
            AutoSizeChildren=   0
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
            Begin ImpulseButton.ISButton Cmd3 
               Height          =   420
               Index           =   0
               Left            =   5805
               TabIndex        =   85
               Top             =   495
               Width           =   720
               _ExtentX        =   1270
               _ExtentY        =   741
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
               ButtonImage     =   "Frm_Grouped_New.frx":7DBD
               ColorButton     =   14871017
               DrawFocusRectangle=   0   'False
            End
            Begin ImpulseButton.ISButton Cmd3 
               Height          =   420
               Index           =   2
               Left            =   3930
               TabIndex        =   86
               Top             =   495
               Width           =   720
               _ExtentX        =   1270
               _ExtentY        =   741
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
               ButtonImage     =   "Frm_Grouped_New.frx":8157
               ColorButton     =   14871017
               DrawFocusRectangle=   0   'False
            End
            Begin ImpulseButton.ISButton Cmd3 
               Height          =   420
               Index           =   1
               Left            =   4815
               TabIndex        =   87
               Top             =   495
               Width           =   825
               _ExtentX        =   1455
               _ExtentY        =   741
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
               ButtonImage     =   "Frm_Grouped_New.frx":84F1
               ColorButton     =   14871017
               DrawFocusRectangle=   0   'False
            End
            Begin ImpulseButton.ISButton Cmd3 
               Height          =   420
               Index           =   3
               Left            =   2805
               TabIndex        =   88
               Top             =   495
               Width           =   840
               _ExtentX        =   1482
               _ExtentY        =   741
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
               ButtonImage     =   "Frm_Grouped_New.frx":888B
               ColorButton     =   14871017
               DrawFocusRectangle=   0   'False
            End
            Begin ImpulseButton.ISButton Cmd3 
               Height          =   420
               Index           =   4
               Left            =   1830
               TabIndex        =   89
               Top             =   495
               Width           =   690
               _ExtentX        =   1217
               _ExtentY        =   741
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
               ButtonImage     =   "Frm_Grouped_New.frx":8C25
               ColorButton     =   14871017
               DrawFocusRectangle=   0   'False
            End
            Begin ImpulseButton.ISButton Cmd3 
               Height          =   420
               Index           =   6
               Left            =   75
               TabIndex        =   90
               Top             =   495
               Width           =   870
               _ExtentX        =   1535
               _ExtentY        =   741
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
               ButtonImage     =   "Frm_Grouped_New.frx":91BF
               ColorButton     =   14871017
               DrawFocusRectangle=   0   'False
               DisabledImageStyle=   1
            End
            Begin VB.Label XPTxtCount3 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Height          =   210
               Left            =   270
               RightToLeft     =   -1  'True
               TabIndex        =   94
               Top             =   120
               Width           =   540
            End
            Begin VB.Label XPTxtCurrent3 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Height          =   210
               Left            =   2745
               RightToLeft     =   -1  'True
               TabIndex        =   93
               Top             =   165
               Width           =   675
            End
            Begin VB.Label Label2 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "⁄œœ «·”Ã·« :"
               Height          =   210
               Index           =   7
               Left            =   960
               RightToLeft     =   -1  'True
               TabIndex        =   92
               Top             =   135
               Width           =   975
            End
            Begin VB.Label Label2 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "«·”Ã· «·Õ«·Ì:"
               Height          =   210
               Index           =   6
               Left            =   3555
               RightToLeft     =   -1  'True
               TabIndex        =   91
               Top             =   135
               Width           =   975
            End
         End
         Begin C1SizerLibCtl.C1Elastic C1Elastic10 
            Height          =   705
            Left            =   0
            TabIndex        =   95
            TabStop         =   0   'False
            Top             =   0
            Width           =   7875
            _cx             =   13891
            _cy             =   1244
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial (Arabic)"
               Size            =   18
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
            Caption         =   ""
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
            Begin VB.TextBox TxtModFlg3 
               Alignment       =   1  'Right Justify
               Height          =   345
               Left            =   2730
               RightToLeft     =   -1  'True
               TabIndex        =   96
               Top             =   180
               Visible         =   0   'False
               Width           =   855
            End
            Begin ImpulseButton.ISButton XPBtnMove3 
               Height          =   345
               Index           =   0
               Left            =   1155
               TabIndex        =   97
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
               ButtonImage     =   "Frm_Grouped_New.frx":9559
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
            Begin ImpulseButton.ISButton XPBtnMove3 
               Height          =   345
               Index           =   2
               Left            =   90
               TabIndex        =   98
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
               ButtonImage     =   "Frm_Grouped_New.frx":98F3
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
            Begin ImpulseButton.ISButton XPBtnMove3 
               Height          =   345
               Index           =   1
               Left            =   1680
               TabIndex        =   99
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
               ButtonImage     =   "Frm_Grouped_New.frx":9C8D
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
            Begin ImpulseButton.ISButton XPBtnMove3 
               Height          =   345
               Index           =   3
               Left            =   615
               TabIndex        =   100
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
               ButtonImage     =   "Frm_Grouped_New.frx":A027
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
            Begin VB.Label Label1 
               Alignment       =   2  'Center
               BackStyle       =   0  'Transparent
               Caption         =   "«„«ﬂ‰ «· Œ“Ì‰"
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
               Index           =   6
               Left            =   4080
               TabIndex        =   109
               Top             =   120
               Width           =   3735
            End
         End
         Begin VB.Label Lbl 
            Alignment       =   2  'Center
            BackColor       =   &H00E2E9E9&
            Caption         =   "ﬂÊœ "
            Height          =   285
            Index           =   9
            Left            =   6405
            RightToLeft     =   -1  'True
            TabIndex        =   108
            Top             =   1200
            Width           =   1005
         End
         Begin VB.Label Lbl 
            Alignment       =   2  'Center
            BackColor       =   &H00E2E9E9&
            Caption         =   "„·«ÕŸ« "
            Height          =   195
            Index           =   8
            Left            =   6405
            RightToLeft     =   -1  'True
            TabIndex        =   107
            Top             =   3150
            Width           =   1200
         End
         Begin VB.Label Lbl 
            Alignment       =   2  'Center
            BackColor       =   &H00E2E9E9&
            Caption         =   "«·«”„ ⁄—»Ì"
            Height          =   315
            Index           =   7
            Left            =   6405
            RightToLeft     =   -1  'True
            TabIndex        =   106
            Top             =   1695
            Width           =   1200
         End
         Begin VB.Label Lbl 
            Alignment       =   2  'Center
            BackColor       =   &H00E2E9E9&
            Caption         =   "«·«”„ «‰Ã·Ì“Ì"
            Height          =   315
            Index           =   6
            Left            =   6405
            RightToLeft     =   -1  'True
            TabIndex        =   105
            Top             =   2130
            Width           =   1200
         End
      End
      Begin C1SizerLibCtl.C1Elastic C1Elastic4 
         Height          =   6510
         Left            =   -8520
         TabIndex        =   34
         TabStop         =   0   'False
         Top             =   45
         Width           =   7875
         _cx             =   13891
         _cy             =   11483
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
         Begin VB.Frame Frm24 
            BackColor       =   &H00E2E9E9&
            BorderStyle     =   0  'None
            Enabled         =   0   'False
            Height          =   1245
            Left            =   360
            RightToLeft     =   -1  'True
            TabIndex        =   121
            Top             =   4275
            Width           =   7155
            Begin VB.TextBox TxtSerial 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00FFFFFF&
               Enabled         =   0   'False
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   178
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   315
               Left            =   4290
               MaxLength       =   50
               RightToLeft     =   -1  'True
               TabIndex        =   124
               Top             =   45
               Width           =   1785
            End
            Begin VB.TextBox TxtVacName4 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00FFFFFF&
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   178
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   315
               Left            =   315
               MaxLength       =   50
               RightToLeft     =   -1  'True
               TabIndex        =   123
               Tag             =   "⁄›Ê« Ì—ÃÏ «œŒ«· √”„  «·‰Ê⁄"
               Top             =   405
               Width           =   5760
            End
            Begin VB.TextBox TxtVacNamee4 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00FFFFFF&
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   178
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   315
               Left            =   315
               MaxLength       =   50
               RightToLeft     =   -1  'True
               TabIndex        =   122
               Top             =   765
               Width           =   5760
            End
            Begin VB.Label Label1 
               Alignment       =   2  'Center
               BackColor       =   &H00E2E9E9&
               Caption         =   "ﬂÊœ  "
               Height          =   195
               Index           =   10
               Left            =   5985
               RightToLeft     =   -1  'True
               TabIndex        =   127
               Top             =   30
               Width           =   1110
            End
            Begin VB.Label Label1 
               Alignment       =   2  'Center
               BackColor       =   &H00E2E9E9&
               Caption         =   "⁄—»Ì"
               Height          =   285
               Index           =   9
               Left            =   6060
               RightToLeft     =   -1  'True
               TabIndex        =   126
               Top             =   360
               Width           =   1035
            End
            Begin VB.Label Label1 
               Alignment       =   2  'Center
               BackColor       =   &H00E2E9E9&
               Caption         =   "«‰Ã·Ì“Ì"
               Height          =   285
               Index           =   8
               Left            =   5985
               RightToLeft     =   -1  'True
               TabIndex        =   125
               Top             =   720
               Width           =   1170
            End
         End
         Begin VB.Frame FraHeader 
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   0  'None
            Height          =   540
            Left            =   0
            RightToLeft     =   -1  'True
            TabIndex        =   110
            Top             =   0
            Width           =   7875
            Begin VB.Frame Frmo2 
               BackColor       =   &H00E0E0E0&
               BorderStyle     =   0  'None
               Enabled         =   0   'False
               Height          =   375
               Left            =   540
               RightToLeft     =   -1  'True
               TabIndex        =   113
               Top             =   450
               Visible         =   0   'False
               Width           =   3105
               Begin MSDataListLib.DataCombo DCUser 
                  CausesValidation=   0   'False
                  Height          =   315
                  Left            =   -255
                  TabIndex        =   114
                  Tag             =   "„‰ ›÷·ﬂ √œŒ· —ﬁ„ «·ﬁ÷Ì…"
                  Top             =   15
                  Width           =   2340
                  _ExtentX        =   4128
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
                  Left            =   2160
                  RightToLeft     =   -1  'True
                  TabIndex        =   115
                  Top             =   45
                  Width           =   855
               End
            End
            Begin VB.TextBox TxtModFlg4 
               Alignment       =   1  'Right Justify
               Appearance      =   0  'Flat
               BackColor       =   &H0000FF00&
               Enabled         =   0   'False
               Height          =   285
               Left            =   2100
               RightToLeft     =   -1  'True
               TabIndex        =   112
               Text            =   "modflag"
               Top             =   -150
               Visible         =   0   'False
               Width           =   465
            End
            Begin VB.TextBox TxtVac_ID 
               Alignment       =   1  'Right Justify
               Appearance      =   0  'Flat
               BackColor       =   &H000000FF&
               Height          =   240
               Left            =   2790
               RightToLeft     =   -1  'True
               TabIndex        =   111
               Top             =   -90
               Visible         =   0   'False
               Width           =   945
            End
            Begin MSComctlLib.ImageList GrdImageList 
               Left            =   3840
               Top             =   480
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
                     Picture         =   "Frm_Grouped_New.frx":A3C1
                     Key             =   "CompanyName"
                  EndProperty
                  BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                     Picture         =   "Frm_Grouped_New.frx":A75B
                     Key             =   "Ser"
                  EndProperty
                  BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                     Picture         =   "Frm_Grouped_New.frx":AAF5
                     Key             =   "Vac_Name"
                  EndProperty
                  BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                     Picture         =   "Frm_Grouped_New.frx":AE8F
                     Key             =   "ShareCount"
                  EndProperty
                  BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                     Picture         =   "Frm_Grouped_New.frx":B229
                     Key             =   "Dis_Count"
                  EndProperty
                  BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                     Picture         =   "Frm_Grouped_New.frx":B5C3
                     Key             =   "Bouns"
                  EndProperty
                  BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                     Picture         =   "Frm_Grouped_New.frx":B95D
                     Key             =   "SharesValue"
                  EndProperty
                  BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                     Picture         =   "Frm_Grouped_New.frx":BEF7
                     Key             =   "BuyValue"
                  EndProperty
               EndProperty
            End
            Begin ImpulseButton.ISButton btnLast4 
               Height          =   315
               Left            =   90
               TabIndex        =   116
               Top             =   150
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
               ButtonImage     =   "Frm_Grouped_New.frx":C291
               ColorButton     =   16777215
               AcclimateGrayTones=   -1  'True
               DrawFocusRectangle=   0   'False
               DisabledImageExtraction=   0
               DisabledImageStyle=   1
            End
            Begin ImpulseButton.ISButton btnNext4 
               Height          =   315
               Left            =   555
               TabIndex        =   117
               Top             =   150
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
               ButtonImage     =   "Frm_Grouped_New.frx":C62B
               ColorButton     =   16777215
               DrawFocusRectangle=   0   'False
               DisabledImageStyle=   1
            End
            Begin ImpulseButton.ISButton btnPrevious4 
               Height          =   315
               Left            =   1155
               TabIndex        =   118
               Top             =   150
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
               ButtonImage     =   "Frm_Grouped_New.frx":C9C5
               ColorButton     =   16777215
               DrawFocusRectangle=   0   'False
               DisabledImageStyle=   1
            End
            Begin ImpulseButton.ISButton btnFirst4 
               Height          =   315
               Left            =   1620
               TabIndex        =   119
               Top             =   150
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
               ButtonImage     =   "Frm_Grouped_New.frx":CD5F
               ColorButton     =   16777215
               DrawFocusRectangle=   0   'False
               DisabledImageStyle=   1
            End
            Begin VB.Label Label1 
               Alignment       =   2  'Center
               BackStyle       =   0  'Transparent
               Caption         =   " ⁄—Ì› «·„Ê«’›« "
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
               Index           =   7
               Left            =   4095
               RightToLeft     =   -1  'True
               TabIndex        =   120
               Top             =   90
               Width           =   3720
            End
         End
         Begin C1SizerLibCtl.C1Elastic C1Elastic11 
            Height          =   900
            Left            =   915
            TabIndex        =   128
            TabStop         =   0   'False
            Top             =   5535
            Width           =   6135
            _cx             =   10821
            _cy             =   1587
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
            AutoSizeChildren=   0
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
            Begin ImpulseButton.ISButton btnNew4 
               Height          =   330
               Left            =   4575
               TabIndex        =   129
               Top             =   555
               Width           =   750
               _ExtentX        =   1323
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
               ButtonImage     =   "Frm_Grouped_New.frx":D0F9
               ColorButton     =   14871017
               DrawFocusRectangle=   0   'False
            End
            Begin ImpulseButton.ISButton btnSave4 
               Height          =   330
               Left            =   3030
               TabIndex        =   130
               Top             =   555
               Width           =   750
               _ExtentX        =   1323
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
               ButtonImage     =   "Frm_Grouped_New.frx":D493
               ColorButton     =   14871017
               DrawFocusRectangle=   0   'False
            End
            Begin ImpulseButton.ISButton btnModify4 
               Height          =   330
               Left            =   3795
               TabIndex        =   131
               Top             =   555
               Width           =   750
               _ExtentX        =   1323
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
               ButtonImage     =   "Frm_Grouped_New.frx":D82D
               ColorButton     =   14871017
               DrawFocusRectangle=   0   'False
            End
            Begin ImpulseButton.ISButton BtnUndo4 
               Height          =   330
               Left            =   2265
               TabIndex        =   132
               Top             =   555
               Width           =   750
               _ExtentX        =   1323
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
               ButtonImage     =   "Frm_Grouped_New.frx":DBC7
               ColorButton     =   14871017
               DrawFocusRectangle=   0   'False
            End
            Begin ImpulseButton.ISButton btnDelete4 
               Height          =   330
               Left            =   1500
               TabIndex        =   133
               Top             =   555
               Width           =   750
               _ExtentX        =   1323
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
               ButtonImage     =   "Frm_Grouped_New.frx":DF61
               ColorButton     =   14871017
               DrawFocusRectangle=   0   'False
            End
            Begin ImpulseButton.ISButton btnQuery4 
               Height          =   330
               Left            =   5040
               TabIndex        =   134
               TabStop         =   0   'False
               ToolTipText     =   "(Ctrl+F)  ··»ÕÀ ≈÷€ÿ Â–« «·„› «Õ √Ê ≈÷€ÿ "
               Top             =   90
               Visible         =   0   'False
               Width           =   840
               _ExtentX        =   1482
               _ExtentY        =   582
               ButtonStyle     =   1
               ButtonPositionImage=   1
               Caption         =   "»ÕÀ"
               BackColor       =   14737632
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
               ButtonImage     =   "Frm_Grouped_New.frx":E4FB
               ColorButton     =   14737632
               DrawFocusRectangle=   0   'False
            End
            Begin ImpulseButton.ISButton BtnUpdate4 
               Height          =   330
               Left            =   3765
               TabIndex        =   135
               TabStop         =   0   'False
               ToolTipText     =   " ÕœÌÀ ﬁ«⁄œ… «·»Ì«‰« "
               Top             =   105
               Visible         =   0   'False
               Width           =   840
               _ExtentX        =   1482
               _ExtentY        =   582
               ButtonStyle     =   1
               ButtonPositionImage=   1
               Caption         =   " ÕœÌÀ"
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
               ButtonImage     =   "Frm_Grouped_New.frx":E895
               ColorButton     =   14871017
               DrawFocusRectangle=   0   'False
            End
            Begin ImpulseButton.ISButton BtnPrint 
               Height          =   285
               Left            =   4725
               TabIndex        =   136
               TabStop         =   0   'False
               Top             =   150
               Visible         =   0   'False
               Width           =   285
               _ExtentX        =   503
               _ExtentY        =   503
               ButtonStyle     =   1
               ButtonPositionImage=   2
               Caption         =   ""
               BackColor       =   14871017
               FontSize        =   14.25
               FontName        =   "Arial"
               FontBold        =   -1  'True
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Arial"
                  Size            =   14.25
                  Charset         =   178
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ButtonImage     =   "Frm_Grouped_New.frx":EC2F
               ColorButton     =   14871017
               DrawFocusRectangle=   0   'False
               DisabledImageStyle=   1
            End
            Begin ImpulseButton.ISButton btnCancel4 
               Height          =   330
               Left            =   705
               TabIndex        =   137
               Top             =   555
               Width           =   750
               _ExtentX        =   1323
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
               ButtonImage     =   "Frm_Grouped_New.frx":EFC9
               ColorButton     =   14871017
               DrawFocusRectangle=   0   'False
               DisabledImageStyle=   1
            End
            Begin VB.Label LabCountRec4 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Height          =   210
               Left            =   240
               RightToLeft     =   -1  'True
               TabIndex        =   141
               Top             =   225
               Width           =   540
            End
            Begin VB.Label LabCurrRec4 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Height          =   210
               Left            =   1800
               RightToLeft     =   -1  'True
               TabIndex        =   140
               Top             =   240
               Width           =   675
            End
            Begin VB.Label Label2 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "⁄œœ «·”Ã·« :"
               Height          =   210
               Index           =   9
               Left            =   810
               RightToLeft     =   -1  'True
               TabIndex        =   139
               Top             =   225
               Width           =   975
            End
            Begin VB.Label Label2 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "«·”Ã· «·Õ«·Ì:"
               Height          =   210
               Index           =   8
               Left            =   2505
               RightToLeft     =   -1  'True
               TabIndex        =   138
               Top             =   225
               Width           =   975
            End
         End
         Begin VSFlex8Ctl.VSFlexGrid Grid4 
            Height          =   3555
            Left            =   0
            TabIndex        =   142
            Top             =   570
            Width           =   7875
            _cx             =   13891
            _cy             =   6271
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
            Rows            =   50
            Cols            =   4
            FixedRows       =   1
            FixedCols       =   1
            RowHeightMin    =   320
            RowHeightMax    =   0
            ColWidthMin     =   0
            ColWidthMax     =   0
            ExtendLastCol   =   0   'False
            FormatString    =   $"Frm_Grouped_New.frx":F363
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
      End
      Begin C1SizerLibCtl.C1Elastic C1Elastic12 
         Height          =   6510
         Left            =   45
         TabIndex        =   143
         TabStop         =   0   'False
         Top             =   45
         Width           =   7875
         _cx             =   13891
         _cy             =   11483
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
         Begin VB.Frame Frame3 
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   0  'None
            Height          =   540
            Left            =   0
            RightToLeft     =   -1  'True
            TabIndex        =   154
            Top             =   0
            Width           =   7875
            Begin VB.TextBox TxtVac_ID5 
               Alignment       =   1  'Right Justify
               Appearance      =   0  'Flat
               BackColor       =   &H000000FF&
               Height          =   285
               Left            =   2430
               RightToLeft     =   -1  'True
               TabIndex        =   159
               Top             =   -210
               Visible         =   0   'False
               Width           =   1065
            End
            Begin VB.TextBox TxtModFlg5 
               Alignment       =   1  'Right Justify
               Appearance      =   0  'Flat
               BackColor       =   &H0000FF00&
               Enabled         =   0   'False
               Height          =   285
               Left            =   3780
               RightToLeft     =   -1  'True
               TabIndex        =   158
               Text            =   "modflag"
               Top             =   450
               Visible         =   0   'False
               Width           =   465
            End
            Begin VB.Frame Frame4 
               BackColor       =   &H00E0E0E0&
               BorderStyle     =   0  'None
               Enabled         =   0   'False
               Height          =   375
               Left            =   540
               RightToLeft     =   -1  'True
               TabIndex        =   155
               Top             =   450
               Visible         =   0   'False
               Width           =   3105
               Begin MSDataListLib.DataCombo DCUser5 
                  CausesValidation=   0   'False
                  Height          =   315
                  Left            =   -255
                  TabIndex        =   156
                  Tag             =   "„‰ ›÷·ﬂ √œŒ· —ﬁ„ «·ﬁ÷Ì…"
                  Top             =   15
                  Width           =   2340
                  _ExtentX        =   4128
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
                  Index           =   16
                  Left            =   2160
                  RightToLeft     =   -1  'True
                  TabIndex        =   157
                  Top             =   45
                  Width           =   855
               End
            End
            Begin MSComctlLib.ImageList GrdImageList5 
               Left            =   3720
               Top             =   -480
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
                     Picture         =   "Frm_Grouped_New.frx":F3EB
                     Key             =   "CompanyName"
                  EndProperty
                  BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                     Picture         =   "Frm_Grouped_New.frx":F785
                     Key             =   "Ser"
                  EndProperty
                  BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                     Picture         =   "Frm_Grouped_New.frx":FB1F
                     Key             =   "Vac_Name"
                  EndProperty
                  BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                     Picture         =   "Frm_Grouped_New.frx":FEB9
                     Key             =   "ShareCount"
                  EndProperty
                  BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                     Picture         =   "Frm_Grouped_New.frx":10253
                     Key             =   "Dis_Count"
                  EndProperty
                  BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                     Picture         =   "Frm_Grouped_New.frx":105ED
                     Key             =   "Bouns"
                  EndProperty
                  BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                     Picture         =   "Frm_Grouped_New.frx":10987
                     Key             =   "SharesValue"
                  EndProperty
                  BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                     Picture         =   "Frm_Grouped_New.frx":10F21
                     Key             =   "BuyValue"
                  EndProperty
               EndProperty
            End
            Begin ImpulseButton.ISButton btnLast5 
               Height          =   315
               Left            =   90
               TabIndex        =   160
               Top             =   150
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
               ButtonImage     =   "Frm_Grouped_New.frx":112BB
               ColorButton     =   16777215
               AcclimateGrayTones=   -1  'True
               DrawFocusRectangle=   0   'False
               DisabledImageExtraction=   0
               DisabledImageStyle=   1
            End
            Begin ImpulseButton.ISButton btnNext5 
               Height          =   315
               Left            =   555
               TabIndex        =   161
               Top             =   150
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
               ButtonImage     =   "Frm_Grouped_New.frx":11655
               ColorButton     =   16777215
               DrawFocusRectangle=   0   'False
               DisabledImageStyle=   1
            End
            Begin ImpulseButton.ISButton btnPrevious5 
               Height          =   315
               Left            =   1155
               TabIndex        =   162
               Top             =   150
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
               ButtonImage     =   "Frm_Grouped_New.frx":119EF
               ColorButton     =   16777215
               DrawFocusRectangle=   0   'False
               DisabledImageStyle=   1
            End
            Begin ImpulseButton.ISButton btnFirst5 
               Height          =   315
               Left            =   1620
               TabIndex        =   163
               Top             =   150
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
               ButtonImage     =   "Frm_Grouped_New.frx":11D89
               ColorButton     =   16777215
               DrawFocusRectangle=   0   'False
               DisabledImageStyle=   1
            End
            Begin VB.Label Label1 
               Alignment       =   2  'Center
               BackStyle       =   0  'Transparent
               Caption         =   "⁄‰«’— «· ﬂ«·Ì› «·’‰«⁄ÌÂ"
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
               Index           =   17
               Left            =   4215
               RightToLeft     =   -1  'True
               TabIndex        =   164
               Top             =   90
               Width           =   3630
            End
         End
         Begin VB.Frame Frm25 
            BackColor       =   &H00E2E9E9&
            BorderStyle     =   0  'None
            Caption         =   "”"
            Enabled         =   0   'False
            Height          =   1485
            Left            =   465
            RightToLeft     =   -1  'True
            TabIndex        =   144
            Top             =   4005
            Width           =   7050
            Begin VB.TextBox TxtVacName5 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00FFFFFF&
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   178
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   315
               Left            =   240
               MaxLength       =   50
               RightToLeft     =   -1  'True
               TabIndex        =   149
               Tag             =   "⁄›Ê« Ì—ÃÏ «œŒ«· √”„ «·„Õ«›Ÿ…"
               Top             =   390
               Width           =   4890
            End
            Begin VB.TextBox TxtSerial5 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00FFFFFF&
               Enabled         =   0   'False
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   178
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   315
               Left            =   4065
               MaxLength       =   50
               RightToLeft     =   -1  'True
               TabIndex        =   148
               Top             =   30
               Width           =   1065
            End
            Begin VB.ComboBox Combo1 
               BackColor       =   &H80000018&
               Height          =   315
               ItemData        =   "Frm_Grouped_New.frx":12123
               Left            =   2280
               List            =   "Frm_Grouped_New.frx":12133
               RightToLeft     =   -1  'True
               Style           =   2  'Dropdown List
               TabIndex        =   147
               Top             =   3150
               Visible         =   0   'False
               Width           =   1005
            End
            Begin VB.TextBox TxtVacNamee5 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00FFFFFF&
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   178
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   315
               Left            =   240
               MaxLength       =   50
               RightToLeft     =   -1  'True
               TabIndex        =   145
               Tag             =   "⁄›Ê« Ì—ÃÏ «œŒ«· √”„ «·„Õ«›Ÿ…"
               Top             =   720
               Width           =   4890
            End
            Begin MSDataListLib.DataCombo DcboExpensesID 
               Height          =   315
               Left            =   240
               TabIndex        =   146
               Tag             =   "«Œ — «·œÊ·… „‰ ›÷·ﬂ"
               Top             =   1110
               Width           =   4890
               _ExtentX        =   8625
               _ExtentY        =   556
               _Version        =   393216
               Text            =   ""
               RightToLeft     =   -1  'True
            End
            Begin VB.Label Label1 
               Alignment       =   2  'Center
               BackColor       =   &H00E2E9E9&
               Caption         =   "«”„ «·⁄‰’— ⁄—»Ì"
               Height          =   285
               Index           =   15
               Left            =   5220
               RightToLeft     =   -1  'True
               TabIndex        =   153
               Top             =   390
               Width           =   1650
            End
            Begin VB.Label Label1 
               Alignment       =   2  'Center
               BackColor       =   &H00E2E9E9&
               Caption         =   "ﬂÊœ «·⁄‰’—"
               Height          =   195
               Index           =   14
               Left            =   5565
               RightToLeft     =   -1  'True
               TabIndex        =   152
               Top             =   30
               Width           =   990
            End
            Begin VB.Label Label1 
               Alignment       =   2  'Center
               BackColor       =   &H00E2E9E9&
               Caption         =   "«”„ «·„’—Ê›"
               Height          =   285
               Index           =   12
               Left            =   5400
               RightToLeft     =   -1  'True
               TabIndex        =   151
               Top             =   1110
               Width           =   1290
            End
            Begin VB.Label Label1 
               Alignment       =   2  'Center
               BackColor       =   &H00E2E9E9&
               Caption         =   "«”„ «·⁄‰’— «‰Ã·Ì“Ì"
               Height          =   285
               Index           =   11
               Left            =   5280
               RightToLeft     =   -1  'True
               TabIndex        =   150
               Top             =   720
               Width           =   1530
            End
         End
         Begin C1SizerLibCtl.C1Elastic C1Elastic13 
            Height          =   900
            Left            =   825
            TabIndex        =   165
            TabStop         =   0   'False
            Top             =   5505
            Width           =   5580
            _cx             =   9842
            _cy             =   1587
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
            AutoSizeChildren=   0
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
            Begin ImpulseButton.ISButton btnNew5 
               Height          =   330
               Left            =   4575
               TabIndex        =   166
               Top             =   555
               Width           =   750
               _ExtentX        =   1323
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
               ButtonImage     =   "Frm_Grouped_New.frx":1214C
               ColorButton     =   14871017
               DrawFocusRectangle=   0   'False
            End
            Begin ImpulseButton.ISButton btnSave5 
               Height          =   330
               Left            =   3030
               TabIndex        =   167
               Top             =   555
               Width           =   750
               _ExtentX        =   1323
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
               ButtonImage     =   "Frm_Grouped_New.frx":124E6
               ColorButton     =   14871017
               DrawFocusRectangle=   0   'False
            End
            Begin ImpulseButton.ISButton btnModify5 
               Height          =   330
               Left            =   3795
               TabIndex        =   168
               Top             =   555
               Width           =   750
               _ExtentX        =   1323
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
               ButtonImage     =   "Frm_Grouped_New.frx":12880
               ColorButton     =   14871017
               DrawFocusRectangle=   0   'False
            End
            Begin ImpulseButton.ISButton BtnUndo5 
               Height          =   330
               Left            =   2265
               TabIndex        =   169
               Top             =   555
               Width           =   750
               _ExtentX        =   1323
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
               ButtonImage     =   "Frm_Grouped_New.frx":12C1A
               ColorButton     =   14871017
               DrawFocusRectangle=   0   'False
            End
            Begin ImpulseButton.ISButton btnDelete5 
               Height          =   330
               Left            =   1500
               TabIndex        =   170
               Top             =   555
               Width           =   750
               _ExtentX        =   1323
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
               ButtonImage     =   "Frm_Grouped_New.frx":12FB4
               ColorButton     =   14871017
               DrawFocusRectangle=   0   'False
            End
            Begin ImpulseButton.ISButton btnQuery5 
               Height          =   330
               Left            =   5160
               TabIndex        =   171
               TabStop         =   0   'False
               ToolTipText     =   "(Ctrl+F)  ··»ÕÀ ≈÷€ÿ Â–« «·„› «Õ √Ê ≈÷€ÿ "
               Top             =   90
               Visible         =   0   'False
               Width           =   840
               _ExtentX        =   1482
               _ExtentY        =   582
               ButtonStyle     =   1
               ButtonPositionImage=   1
               Caption         =   "»ÕÀ"
               BackColor       =   14737632
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
               ButtonImage     =   "Frm_Grouped_New.frx":1354E
               ColorButton     =   14737632
               DrawFocusRectangle=   0   'False
            End
            Begin ImpulseButton.ISButton BtnUpdate5 
               Height          =   330
               Left            =   3765
               TabIndex        =   172
               TabStop         =   0   'False
               ToolTipText     =   " ÕœÌÀ ﬁ«⁄œ… «·»Ì«‰« "
               Top             =   105
               Visible         =   0   'False
               Width           =   840
               _ExtentX        =   1482
               _ExtentY        =   582
               ButtonStyle     =   1
               ButtonPositionImage=   1
               Caption         =   " ÕœÌÀ"
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
               ButtonImage     =   "Frm_Grouped_New.frx":138E8
               ColorButton     =   14871017
               DrawFocusRectangle=   0   'False
            End
            Begin ImpulseButton.ISButton ISButton21 
               Height          =   285
               Left            =   4725
               TabIndex        =   173
               TabStop         =   0   'False
               Top             =   150
               Visible         =   0   'False
               Width           =   285
               _ExtentX        =   503
               _ExtentY        =   503
               ButtonStyle     =   1
               ButtonPositionImage=   2
               Caption         =   ""
               BackColor       =   14871017
               FontSize        =   14.25
               FontName        =   "Arial"
               FontBold        =   -1  'True
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Arial"
                  Size            =   14.25
                  Charset         =   178
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ButtonImage     =   "Frm_Grouped_New.frx":13C82
               ColorButton     =   14871017
               DrawFocusRectangle=   0   'False
               DisabledImageStyle=   1
            End
            Begin ImpulseButton.ISButton btnCancel5 
               Height          =   330
               Left            =   705
               TabIndex        =   174
               Top             =   555
               Width           =   750
               _ExtentX        =   1323
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
               ButtonImage     =   "Frm_Grouped_New.frx":1401C
               ColorButton     =   14871017
               DrawFocusRectangle=   0   'False
               DisabledImageStyle=   1
            End
            Begin VB.Label Label2 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "«·”Ã· «·Õ«·Ì:"
               Height          =   210
               Index           =   11
               Left            =   2505
               RightToLeft     =   -1  'True
               TabIndex        =   178
               Top             =   225
               Width           =   975
            End
            Begin VB.Label Label2 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "⁄œœ «·”Ã·« :"
               Height          =   210
               Index           =   10
               Left            =   810
               RightToLeft     =   -1  'True
               TabIndex        =   177
               Top             =   225
               Width           =   975
            End
            Begin VB.Label LabCurrRec5 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Height          =   210
               Left            =   1800
               RightToLeft     =   -1  'True
               TabIndex        =   176
               Top             =   240
               Width           =   675
            End
            Begin VB.Label LabCountRec5 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Height          =   210
               Left            =   240
               RightToLeft     =   -1  'True
               TabIndex        =   175
               Top             =   225
               Width           =   540
            End
         End
         Begin VSFlex8UCtl.VSFlexGrid Grid5 
            Height          =   3285
            Left            =   0
            TabIndex        =   179
            Top             =   600
            Width           =   7875
            _cx             =   13891
            _cy             =   5794
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
            Rows            =   50
            Cols            =   5
            FixedRows       =   1
            FixedCols       =   1
            RowHeightMin    =   320
            RowHeightMax    =   0
            ColWidthMin     =   0
            ColWidthMax     =   0
            ExtendLastCol   =   -1  'True
            FormatString    =   $"Frm_Grouped_New.frx":143B6
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
   End
End
Attribute VB_Name = "FrmPay_Garanty_Shipment3M"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public SendForm As Integer
'#####################################################################################################################################################
Dim RsSavRec As ADODB.Recordset
Dim RecId As String
Dim II As Long
'#####################################################################################################################################################
Dim RS As ADODB.Recordset
Dim TTP As clstooltip
'#####################################################################################################################################################
Dim Rs2 As ADODB.Recordset
Dim TTP2 As clstooltip
'#####################################################################################################################################################
Dim Rs3 As ADODB.Recordset
Dim TTP3 As clstooltip
'#####################################################################################################################################################
Dim RsSavRec4 As ADODB.Recordset
Dim BKGrndPic4 As ClsBackGroundPic
Dim RecId4 As String
'Dim II4 As Long
'#####################################################################################################################################################
Dim RsSavRec5 As ADODB.Recordset
Dim BKGrndPic5 As ClsBackGroundPic
Dim RecId5 As String
'Dim II5 As Long
Dim cSearch  As clsDCboSearch
'#####################################################################################################################################################

Private Sub ChangeLang()

    Label1(3).Caption = "Code"
    Label1(0).Caption = "Name Ar"
    Label1(1).Caption = "Name Eng"
    ISButton1.Caption = "Prient"
    btnQuery.Caption = "Search"
    With Grid
        .TextMatrix(0, .ColIndex("UnitID")) = "Unit Code"
        .TextMatrix(0, .ColIndex("UnitName")) = "Unit Name AR"
        .TextMatrix(0, .ColIndex("UnitNameE")) = "Unit Name Eng"
        Label1(2).Caption = "Unit  Data"
        btnNew.Caption = "New"
        btnModify.Caption = "Modify"
        btnSave.Caption = "Save"
        BtnUndo.Caption = "Undo"
        btnDelete.Caption = "Delete"
        btnCancel.Caption = "Exit"
        Label2(0).Caption = "Current Record"
        Label2(1).Caption = "NO Of Record"
    End With

    Dim XPic As IPictureDisp
    Set XPic = Me.btnFirst.ButtonImage
    Set Me.btnFirst.ButtonImage = Me.btnLast.ButtonImage
    Set Me.btnLast.ButtonImage = XPic
    Set XPic = Me.btnPrevious.ButtonImage
    Set Me.btnPrevious.ButtonImage = Me.btnNext.ButtonImage
    Set Me.btnNext.ButtonImage = XPic
'#####################################################################################################################################################
    Dim XPic2 As IPictureDisp
    Set XPic2 = Me.XPBtnMove(1).ButtonImage
    Set Me.XPBtnMove(1).ButtonImage = Me.XPBtnMove(2).ButtonImage
    Set Me.XPBtnMove(2).ButtonImage = XPic2
    Set XPic2 = Me.XPBtnMove(0).ButtonImage
    Set Me.XPBtnMove(0).ButtonImage = Me.XPBtnMove(3).ButtonImage
    Set Me.XPBtnMove(3).ButtonImage = XPic2

    Label1(4).Caption = "Color Data"
    Lbl(0).Caption = "Color Code"
    Lbl(3).Caption = "color  Name"
    Lbl(1).Caption = "Remarks"
    Label2(2).Caption = "Current Record"
    Label2(3).Caption = "NO. Recordes"
    Me.Cmd(0).Caption = "New"
    Me.Cmd(1).Caption = "Edit"
    Me.Cmd(2).Caption = "Save"
    Me.Cmd(3).Caption = "Undo"
    Me.Cmd(4).Caption = "Delete"
    Me.Cmd(6).Caption = "Exit"
'#####################################################################################################################################################
    Set XPic = Me.XPBtnMove2(1).ButtonImage
    Set Me.XPBtnMove2(1).ButtonImage = Me.XPBtnMove2(2).ButtonImage
    Set Me.XPBtnMove2(2).ButtonImage = XPic
    Set XPic = Me.XPBtnMove2(0).ButtonImage
    Set Me.XPBtnMove2(0).ButtonImage = Me.XPBtnMove2(3).ButtonImage
    Set Me.XPBtnMove2(3).ButtonImage = XPic
    
    Label1(5).Caption = "Size Data"
    Lbl(5).Caption = "Size Code"
    Lbl(2).Caption = "Size  Name"
    Label2(4).Caption = "Remarks"
    Label2(5).Caption = "Current Record"
    Lbl(4).Caption = "NO. Recordes"
    Me.Cmd2(0).Caption = "New"
    Me.Cmd2(1).Caption = "Edit"
    Me.Cmd2(2).Caption = "Save"
    Me.Cmd2(3).Caption = "Undo"
    Me.Cmd2(4).Caption = "Delete"
    Me.Cmd2(6).Caption = "Exit"
'#####################################################################################################################################################
    Set XPic = Me.XPBtnMove3(1).ButtonImage
    Set Me.XPBtnMove3(1).ButtonImage = Me.XPBtnMove3(2).ButtonImage
    Set Me.XPBtnMove3(2).ButtonImage = XPic
    Set XPic = Me.XPBtnMove3(0).ButtonImage
    Set Me.XPBtnMove3(0).ButtonImage = Me.XPBtnMove3(3).ButtonImage
    Set Me.XPBtnMove3(3).ButtonImage = XPic

    Label1(6).Caption = "Stores Locations"
    Lbl(9).Caption = " Code"
    Lbl(7).Caption = " Name Ar"
    Lbl(6).Caption = " Name Eng"
    Lbl(8).Caption = "Remarks"
    Label2(6).Caption = "Current Record"
    Label2(7).Caption = "NO. Recordes"
    Me.Cmd3(0).Caption = "New"
    Me.Cmd3(1).Caption = "Edit"
    Me.Cmd3(2).Caption = "Save"
    Me.Cmd3(3).Caption = "Undo"
    Me.Cmd3(4).Caption = "Delete"
    Me.Cmd3(6).Caption = "Exit"
'#####################################################################################################################################################
    Set XPic = Me.btnFirst4.ButtonImage
    Set Me.btnFirst4.ButtonImage = Me.btnLast4.ButtonImage
    Set Me.btnLast4.ButtonImage = XPic
    Set XPic = Me.btnPrevious4.ButtonImage
    Set Me.btnPrevious4.ButtonImage = Me.btnNext4.ButtonImage
    Set Me.btnNext4.ButtonImage = XPic

    Label1(7).Caption = "Items Specifications"
    Label1(10).Caption = "Code"
    Label1(9).Caption = "Name AR"
    Label1(8).Caption = "Name ENG"

    Label2(8).Caption = "Current Record"
    Label2(9).Caption = "NO. Recordes"

    btnNew4.Caption = "New"
    btnModify4.Caption = "Modify"
    btnSave4.Caption = "Save"
    BtnUndo4.Caption = "Undo"
    btnDelete4.Caption = "Delete"
    btnCancel4.Caption = "Exit"

    With Me.Grid4
        .TextMatrix(0, .ColIndex("Ser")) = "Ser"
        .TextMatrix(0, .ColIndex("id")) = "Id"
        .TextMatrix(0, .ColIndex("name")) = "Name AR"
        .TextMatrix(0, .ColIndex("namee")) = "Name ENG"
    End With
'######################################################################################################################################################
    Set XPic = Me.btnFirst5.ButtonImage
    Set Me.btnFirst5.ButtonImage = Me.btnLast5.ButtonImage
    Set Me.btnLast5.ButtonImage = XPic
    Set XPic = Me.btnPrevious5.ButtonImage
    Set Me.btnPrevious5.ButtonImage = Me.btnNext5.ButtonImage
    Set Me.btnNext5.ButtonImage = XPic

    Label1(17).Caption = "Production Cost component "
    With Me.Grid5
        .TextMatrix(0, .ColIndex("Ser")) = "Ser"
        .TextMatrix(0, .ColIndex("TBLProductionElementsId")) = "Element ID"
        .TextMatrix(0, .ColIndex("Name")) = " Name A"
        .TextMatrix(0, .ColIndex("NameE")) = " Name E"
        .TextMatrix(0, .ColIndex("ExpensesID")) = "Expenses Name"
    End With
    
    Label1(14).Caption = "ID"
    Label1(15).Caption = "Name AR"
    Label1(11).Caption = "Name En"
    Label1(12).Caption = "Expenses Name"
    Label2(11).Caption = "Curr. Rec."
    Label2(10).Caption = "Rec. Count."
    btnNew5.Caption = "New"
    btnModify5.Caption = "Modify"
    btnSave5.Caption = "Save"
    BtnUndo5.Caption = "Undo"
    btnDelete5.Caption = "Delete"
    btnCancel5.Caption = "Exit"
End Sub
Private Sub Form_Load()

    On Error GoTo ErrTrap
    
    Me.left = (mdifrmmain.Width - Me.Width) / 2
    Me.top = (mdifrmmain.Height - Me.Height) / 2 - 500

    If SystemOptions.UserInterface = EnglishInterface Then
        SetInterface Me
        ChangeLang
    End If
    
    If SendForm = 0 Then
        Dim cGrdBack As New ClsBackGroundPic
        Set Me.Grid.WallPaper = cGrdBack.Picture
        Dim i As Integer
        Dim My_SQL As String
        Dim Dcombos As ClsDataCombos
    
        ScreenNameArabic = "  «·ÊÕœ«  «·„” Œœ„… ›Ï «·»—‰«„Ã "
        ScreenNameEnglish = " Units Data  "
        RegisterLogInOut Me.name, ScreenNameArabic, ScreenNameEnglish, "1"
    
        My_SQL = "TblUnites"
        Set RsSavRec = New ADODB.Recordset
        RsSavRec.CursorLocation = adUseClient
        RsSavRec.Open My_SQL, Cn, adOpenKeyset, adLockPessimistic, adCmdTableDirect
        Me.TxtModFlg.Text = "R"
         Resize_Form Me
        FillGridWithData
        With Me.Grid
            For i = 0 To .Cols - 1
                .Cell(flexcpPictureAlignment, 0, i) = flexPicAlignRightCenter
            Next
            .ExtendLastCol = True
        End With
        BtnFirst_Click
'#####################################################################################################################################################
    ElseIf SendForm = 1 Then
        ScreenNameArabic = "‘«‘… «·Ê«‰ «·«’‰«›  "
        ScreenNameEnglish = "  Items Color Data "
        RegisterLogInOut Me.name, ScreenNameArabic, ScreenNameEnglish, "1"
    
        Resize_Form Me
        Set RS = New ADODB.Recordset
        RS.Open "TblItemsColors", Cn, adOpenStatic, adLockOptimistic, adCmdTable
        Me.TxtModFlg1.Text = "R"
        Retrive
'#####################################################################################################################################################
    ElseIf SendForm = 2 Then
        ScreenNameArabic = "»Ì«‰«  „ﬁ«”«  «·√’‰«›   "
        ScreenNameEnglish = "  Items Size"
        RegisterLogInOut Me.name, ScreenNameArabic, ScreenNameEnglish, "1"

        Set Cmd2(0).ButtonImage = mdifrmmain.ImgLstTree.ListImages("New").Picture
        Set Cmd2(1).ButtonImage = mdifrmmain.ImgLstTree.ListImages("Edit").Picture
        Set Cmd2(2).ButtonImage = mdifrmmain.ImgLstTree.ListImages("save").Picture
        Set Cmd2(3).ButtonImage = mdifrmmain.ImgLstTree.ListImages("Undo").Picture
        Set Cmd2(4).ButtonImage = mdifrmmain.ImgLstTree.ListImages("Del").Picture
        Set Cmd2(6).ButtonImage = mdifrmmain.ImgLstTree.ListImages("Exit").Picture
        Resize_Form Me
        Set Rs2 = New ADODB.Recordset
        Rs2.Open "TblItemsSizes", Cn, adOpenStatic, adLockOptimistic, adCmdTable
        Me.TxtModFlg2.Text = "R"
        Retrive2
'#####################################################################################################################################################
    ElseIf SendForm = 3 Then
        ScreenNameArabic = " «‰Ê«⁄ ›—“ «·«’‰«›  "
        ScreenNameEnglish = " Items Class "
        RegisterLogInOut Me.name, ScreenNameArabic, ScreenNameEnglish, "1"

        Set Cmd3(0).ButtonImage = mdifrmmain.ImgLstTree.ListImages("New").Picture
        Set Cmd3(1).ButtonImage = mdifrmmain.ImgLstTree.ListImages("Edit").Picture
        Set Cmd3(2).ButtonImage = mdifrmmain.ImgLstTree.ListImages("save").Picture
        Set Cmd3(3).ButtonImage = mdifrmmain.ImgLstTree.ListImages("Undo").Picture
        Set Cmd3(4).ButtonImage = mdifrmmain.ImgLstTree.ListImages("Del").Picture
        Set Cmd3(6).ButtonImage = mdifrmmain.ImgLstTree.ListImages("Exit").Picture
        Resize_Form Me
        Set Rs3 = New ADODB.Recordset
        Rs3.Open "TblstoresLocations", Cn, adOpenStatic, adLockOptimistic, adCmdTable
        Me.TxtModFlg3.Text = "R"
        Retrive3
'#####################################################################################################################################################
    ElseIf SendForm = 4 Then
        My_SQL = "TblSpecification"
        Set BKGrndPic4 = New ClsBackGroundPic
        Set RsSavRec4 = New ADODB.Recordset
        RsSavRec4.CursorLocation = adUseClient
        RsSavRec4.Open My_SQL, Cn, adOpenKeyset, adLockPessimistic, adCmdTableDirect
        Me.TxtModFlg4.Text = "R"
        Resize_Form Me
        My_SQL = "select UserID,UserName From tblUsers "
        fill_combo DCUser, My_SQL

        FillGrid4WithData

        With Me.Grid4
            .Cell(flexcpPicture, 0, .ColIndex("name")) = Me.GrdImageList.ListImages("Vac_Name").ExtractIcon
            .Cell(flexcpPicture, 0, .ColIndex("Ser")) = Me.GrdImageList.ListImages("Ser").ExtractIcon
            For i = 0 To .Cols - 1
                .Cell(flexcpPictureAlignment, 0, i) = flexPicAlignRightCenter
            Next
            .ExtendLastCol = True
            .WallPaper = BKGrndPic4.Picture
            .RowHeight(-1) = 300
        End With
        btnFirst4_Click
'######################################################################################################################################################
    ElseIf SendForm = 5 Then
        ScreenNameArabic = " ⁄‰«’— «· ﬂ«·Ì› «·’‰«⁄ÌÂ "
        ScreenNameEnglish = "  Production Cost Elements "
        RegisterLogInOut Me.name, ScreenNameArabic, ScreenNameEnglish, "1"
        My_SQL = "TBLProductionElements"
        Set BKGrndPic5 = New ClsBackGroundPic
        Set RsSavRec5 = New ADODB.Recordset
        RsSavRec5.CursorLocation = adUseClient
        RsSavRec5.Open My_SQL, Cn, adOpenKeyset, adLockPessimistic, adCmdTableDirect
        Me.TxtModFlg5.Text = "R"
        Resize_Form Me
        My_SQL = "select UserID,UserName From tblUsers "
        fill_combo DCUser5, My_SQL
        Set Dcombos = New ClsDataCombos
        Dcombos.GetExpensesNames Me.DcboExpensesID
        Set cSearch = New clsDCboSearch
        Set cSearch.Client = Me.DcboExpensesID
        ModFgLib.LinkFgColWithDataCombo Grid5, Grid5.ColIndex("ExpensesID"), Me.DcboExpensesID
        FillGrid5WithData
        With Me.Grid5
            .Cell(flexcpPicture, 0, .ColIndex("Name")) = Me.GrdImageList5.ListImages("Vac_Name").ExtractIcon
            .Cell(flexcpPicture, 0, .ColIndex("NameE")) = Me.GrdImageList5.ListImages("Vac_Name").ExtractIcon
            .Cell(flexcpPicture, 0, .ColIndex("Ser")) = Me.GrdImageList5.ListImages("Ser").ExtractIcon
            For i = 0 To .Cols - 1
                .Cell(flexcpPictureAlignment, 0, i) = flexPicAlignRightCenter
            Next
            .ExtendLastCol = True
            .WallPaper = BKGrndPic5.Picture
            .RowHeight(-1) = 300
        End With
        btnFirst5_Click
        
'######################################################################################################################################################
    End If
        C1Tab1.TabVisible(SendForm) = True
    C1Tab1.CurrTab = SendForm
    
ErrTrap:
End Sub
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)

    Dim IntResult As String
    Dim StrMSG As String
    
    On Error GoTo ErrTrap

    If SendForm = 0 Then
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
        End If
    ElseIf SendForm = 1 Then
        If Me.TxtModFlg1.Text <> "R" Then
            Select Case Me.TxtModFlg1.Text
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
        End If
    ElseIf SendForm = 2 Then
        If Me.TxtModFlg2.Text <> "R" Then
            Select Case Me.TxtModFlg2.Text
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
        End If
    ElseIf SendForm = 3 Then
        If Me.TxtModFlg3.Text <> "R" Then
            Select Case Me.TxtModFlg3.Text
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
        End If
    ElseIf SendForm = 4 Then
        If Me.TxtModFlg4.Text <> "R" Then
            Select Case Me.TxtModFlg4.Text
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
        End If
    ElseIf SendForm = 5 Then
        If Me.TxtModFlg5.Text <> "R" Then
            Select Case Me.TxtModFlg5.Text
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
        End If
    'ElseIf SendForm = 6 Then
    '    If Me.TxtModFlg6.Text <> "R" Then
    '        Select Case Me.TxtModFlg6.Text
    '            Case "N"
    '                If SystemOptions.UserInterface = EnglishInterface Then
    '                    StrMSG = "You will close this screen before save " & Chr(13)
    '                    StrMSG = StrMSG & " the new data  " & Chr(13)
    '                    StrMSG = StrMSG & " do you want save before exit" & Chr(13)
    '                    StrMSG = StrMSG & "yes" & "-" & "save the new data" & Chr(13)
    '                    StrMSG = StrMSG & "no" & "-" & "Don't save" & Chr(13)
    '                    StrMSG = StrMSG & "cancel" & "-" & "Cancel Exit" & Chr(13)
    '                Else
    '                    StrMSG = "”Ê› Ì „ €·ﬁ «·‘«‘… Ê·„  ‰ Â „‰  ”ÃÌ·" & Chr(13)
    '                    StrMSG = StrMSG & " «·»Ì«‰«  «·ÃœÌœ… «·Õ«·Ì… " & Chr(13)
    '                    StrMSG = StrMSG & " Â·  —Ìœ «·Õ›Ÿ ﬁ»· «·Œ—ÊÃ" & Chr(13)
    '                    StrMSG = StrMSG & "‰⁄„" & "-" & "Ì „ Õ›Ÿ «·»Ì«‰«  «·ÃœÌœ…" & Chr(13)
    '                    StrMSG = StrMSG & "·«" & "-" & "·‰ Ì „ «·Õ›Ÿ" & Chr(13)
    '                    StrMSG = StrMSG & "≈·€«¡ «·√„—" & "-" & "≈·€«¡ ⁄„·Ì… «·Œ—ÊÃ" & Chr(13)
    '                End If
    '            Case "E"
    '                If SystemOptions.UserInterface = EnglishInterface Then
    '                    StrMSG = "You will close this screen before save  " & Chr(13)
    '                    StrMSG = StrMSG & " the Modifications  " & Chr(13)
    '                    StrMSG = StrMSG & " do you want save before exit" & Chr(13)
    '                    StrMSG = StrMSG & "yes" & "-" & "save the new data" & Chr(13)
    '                    StrMSG = StrMSG & "no" & "-" & "Don't save" & Chr(13)
    '                    StrMSG = StrMSG & "cancel" & "-" & "Cancel Exit" & Chr(13)
    '                Else
    '                    StrMSG = "”Ê› Ì „ €·ﬁ «·‘«‘… Ê·„  ‰ Â „‰  ”ÃÌ·" & Chr(13)
    '                    StrMSG = StrMSG & " «· ⁄œÌ·«  «·ÃœÌœ… ⁄·Ï «·”Ã· «·Õ«·Ï " & Chr(13)
    '                    StrMSG = StrMSG & " Â·  —Ìœ «·Õ›Ÿ ﬁ»· «·Œ—ÊÃ" & Chr(13)
    '                    StrMSG = StrMSG & "‰⁄„" & "-" & "Ì „ Õ›Ÿ «· ⁄œÌ·«   «·ÃœÌœ…" & Chr(13)
    '                    StrMSG = StrMSG & "·«" & "-" & "·‰ Ì „ «·Õ›Ÿ" & Chr(13)
    '                    StrMSG = StrMSG & "≈·€«¡ «·√„—" & "-" & "≈·€«¡ ⁄„·Ì… «·Œ—ÊÃ" & Chr(13)
    '                End If
    '        End Select
    '    End If
    'ElseIf SendForm = 7 Then
    '    If Me.TxtModFlg7.Text <> "R" Then
    '        Select Case Me.TxtModFlg7.Text
    '            Case "N"
    '                If SystemOptions.UserInterface = EnglishInterface Then
    '                    StrMSG = "You will close this screen before save " & Chr(13)
    '                    StrMSG = StrMSG & " the new data  " & Chr(13)
    '                    StrMSG = StrMSG & " do you want save before exit" & Chr(13)
    '                    StrMSG = StrMSG & "yes" & "-" & "save the new data" & Chr(13)
    '                    StrMSG = StrMSG & "no" & "-" & "Don't save" & Chr(13)
    '                    StrMSG = StrMSG & "cancel" & "-" & "Cancel Exit" & Chr(13)
    '                Else
    '                    StrMSG = "”Ê› Ì „ €·ﬁ «·‘«‘… Ê·„  ‰ Â „‰  ”ÃÌ·" & Chr(13)
    '                    StrMSG = StrMSG & " «·»Ì«‰«  «·ÃœÌœ… «·Õ«·Ì… " & Chr(13)
    '                    StrMSG = StrMSG & " Â·  —Ìœ «·Õ›Ÿ ﬁ»· «·Œ—ÊÃ" & Chr(13)
    '                    StrMSG = StrMSG & "‰⁄„" & "-" & "Ì „ Õ›Ÿ «·»Ì«‰«  «·ÃœÌœ…" & Chr(13)
    '                    StrMSG = StrMSG & "·«" & "-" & "·‰ Ì „ «·Õ›Ÿ" & Chr(13)
    '                    StrMSG = StrMSG & "≈·€«¡ «·√„—" & "-" & "≈·€«¡ ⁄„·Ì… «·Œ—ÊÃ" & Chr(13)
    '                End If
    '            Case "E"
    '                If SystemOptions.UserInterface = EnglishInterface Then
    '                    StrMSG = "You will close this screen before save  " & Chr(13)
    '                    StrMSG = StrMSG & " the Modifications  " & Chr(13)
    '                    StrMSG = StrMSG & " do you want save before exit" & Chr(13)
    '                    StrMSG = StrMSG & "yes" & "-" & "save the new data" & Chr(13)
    '                    StrMSG = StrMSG & "no" & "-" & "Don't save" & Chr(13)
    '                    StrMSG = StrMSG & "cancel" & "-" & "Cancel Exit" & Chr(13)
    '                Else
    '                    StrMSG = "”Ê› Ì „ €·ﬁ «·‘«‘… Ê·„  ‰ Â „‰  ”ÃÌ·" & Chr(13)
    '                    StrMSG = StrMSG & " «· ⁄œÌ·«  «·ÃœÌœ… ⁄·Ï «·”Ã· «·Õ«·Ï " & Chr(13)
    '                    StrMSG = StrMSG & " Â·  —Ìœ «·Õ›Ÿ ﬁ»· «·Œ—ÊÃ" & Chr(13)
    '                    StrMSG = StrMSG & "‰⁄„" & "-" & "Ì „ Õ›Ÿ «· ⁄œÌ·«   «·ÃœÌœ…" & Chr(13)
    '                    StrMSG = StrMSG & "·«" & "-" & "·‰ Ì „ «·Õ›Ÿ" & Chr(13)
    '                    StrMSG = StrMSG & "≈·€«¡ «·√„—" & "-" & "≈·€«¡ ⁄„·Ì… «·Œ—ÊÃ" & Chr(13)
    '                End If
    '        End Select
    '    End If
    End If
    
    If StrMSG <> "" Then
        IntResult = MsgBox(StrMSG, vbMsgBoxRight + vbYesNoCancel + vbMsgBoxRtlReading + vbQuestion, App.title)
        Select Case IntResult
            Case vbYes
                Cancel = True
                Select Case SendForm
                    Case 0
                        btnSave_Click
                    Case 1
                        SaveData
                    Case 2
                        SaveData2
                    Case 3
                        SaveData3
                    Case 4
                        btnSave4_Click
                    Case 5
                        btnSave5_Click
                    Case 6
                        'btnSave6_Click
                    Case 7
                        'btnSave7_Click
                End Select
            Case vbCancel
                Cancel = True
        End Select
    End If
    
    
    Exit Sub
ErrTrap:
End Sub
Private Sub Form_Unload(Cancel As Integer)

    On Error GoTo ErrTrap
    
'#####################################################################################################################################################
    RegisterLogInOut Me.name, ScreenNameArabic, ScreenNameEnglish
    If RsSavRec.State = adStateOpen Then
        If Not (RsSavRec.EOF Or RsSavRec.BOF) Then
            If RsSavRec.EditMode <> adEditNone Then
                RsSavRec.CancelUpdate
            End If
        End If
        RsSavRec.Close
        Set RsSavRec = Nothing
    End If
'#####################################################################################################################################################
    RegisterLogInOut Me.name, ScreenNameArabic, ScreenNameEnglish
    If RS.State = adStateOpen Then
        If Not (RS.EOF Or RS.BOF) Then
            If RS.EditMode <> adEditNone Then
                RS.CancelUpdate
            End If
        End If
        RS.Close
    End If
    Set RS = Nothing
    Set TTP = Nothing
'#####################################################################################################################################################
    RegisterLogInOut Me.name, ScreenNameArabic, ScreenNameEnglish
    If Rs2.State = adStateOpen Then
        If Not (Rs2.EOF Or Rs2.BOF) Then
            If Rs2.EditMode <> adEditNone Then
                Rs2.CancelUpdate
            End If
        End If
        Rs2.Close
    End If
    Set Rs2 = Nothing
    Set TTP2 = Nothing
'#####################################################################################################################################################
    RegisterLogInOut Me.name, ScreenNameArabic, ScreenNameEnglish
    If Rs3.State = adStateOpen Then
        If Not (Rs3.EOF Or Rs3.BOF) Then
            If Rs3.EditMode <> adEditNone Then
                Rs3.CancelUpdate
            End If
        End If
        Rs3.Close
    End If
    Set Rs3 = Nothing
    Set TTP3 = Nothing
'#####################################################################################################################################################
    If RsSavRec4.State = adStateOpen Then
        If Not (RsSavRec4.EOF Or RsSavRec4.BOF) Then
            If RsSavRec4.EditMode <> adEditNone Then
                RsSavRec4.CancelUpdate
            End If
        End If
        RsSavRec4.Close
        Set RsSavRec4 = Nothing
    End If
'#####################################################################################################################################################
    RegisterLogInOut Me.name, ScreenNameArabic, ScreenNameEnglish
    If RsSavRec5.State = adStateOpen Then
        If Not (RsSavRec5.EOF Or RsSavRec5.BOF) Then
            If RsSavRec5.EditMode <> adEditNone Then
                RsSavRec5.CancelUpdate
            End If
        End If
        RsSavRec5.Close
        Set RsSavRec5 = Nothing
    End If
    Set cSearch = Nothing
'######################################################################################################################################################
Exit Sub
ErrTrap:
End Sub
'#####################################################################################################################################################
'#####################################################################################################################################################
'#####################################################################################################################################################
Private Sub btnQuery_Click()
    Load FrmSearchUnit
    FrmSearchUnit.show
End Sub
Function print_report(Optional NoteSerial As String, Optional X As Integer)
    
    Dim MySQL As String
    Dim RsData As New ADODB.Recordset
    Dim xApp As New CRAXDRT.Application
    Dim xReport As CRAXDRT.Report
    Dim CViewer As ClsReportViewer
    Dim StrReportTitle As String
    Dim StrFileName As String
    Dim Msg As String
   
    MySQL = "  SELECT     UnitID, UnitName, UnitNamee"
    MySQL = MySQL & " From dbo.TblUnites"

    If SystemOptions.UserInterface = ArabicInterface Then
        StrFileName = App.path & "\REPORTS\REPORTS NEW\" & "repUnit.rpt"
    Else
        StrFileName = App.path & "\REPORTS\REPORTS NEW\" & "repUnit.rpt"
    End If
    
    If Dir(StrFileName) = "" Then
        Screen.MousePointer = vbDefault
        Exit Function
    End If

    Set RsData = New ADODB.Recordset
    RsData.Open MySQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

    If RsData.BOF Or RsData.EOF Then
        If SystemOptions.UserInterface = ArabicInterface Then
            Msg = "·« ÊÃœ »Ì«‰«  ··⁄—÷"
        Else
            Msg = "There's no data to show"
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
        StrReportTitle = "" '& StrAccountName
    Else
        xReport.ParameterFields(1).AddCurrentValue cCompanyInfo.ArabCompanyName ' RPTCompany_Name_Eng
        xReport.ParameterFields(4).AddCurrentValue get_branch_name(val(my_branch))
        StrReportTitle = ""
    End If

    xReport.ParameterFields(3).AddCurrentValue user_name
    
    Dim Total As String
    Dim dif As String
    Dim totl As Double

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
Private Sub BtnCancel_Click()
    Unload Me
End Sub
Private Sub btnDelete_Click()

    On Error GoTo ErrTrap

    Dim MSGType As Integer
    Dim BlnRecExist As Boolean
    Dim StrMSG  As String
    Dim Msg As String
    
    If TxtUnitID.Text <> "" Then
        If UnitsHaveTransactions(val(TxtUnitID.Text)) = True Then
            If SystemOptions.UserInterface = ArabicInterface Then
                Msg = " ·« Ì„ﬂ‰ Õ–› Â–… «·ÊÕœ… ·ÊÃÊœ ⁄„·Ì«  „— »ÿÂ »Â« "
            Else
                Msg = " Can't Modify Unit - Unit Have Transaction "
            End If

            MsgBox Msg, vbCritical
            Exit Sub
        End If
        If SystemOptions.UserInterface = ArabicInterface Then
            MSGType = MsgBox("Â·  —€» ›Ì Õ–› Â–« «·”Ã·", vbYesNo + vbMsgBoxRight, App.title)
        Else
            MSGType = MsgBox("Delete This Record", vbYesNo + vbMsgBoxRight, App.title)
        End If
        If MSGType = vbYes Then
            RsSavRec.Find "UnitID=" & val(TxtUnitID.Text), , adSearchForward, 1
            If Not (RsSavRec.BOF Or RsSavRec.EOF) Then
                CuurentLogdata ("D")
                RsSavRec.Delete
                If SystemOptions.UserInterface = ArabicInterface Then
                    MsgBox " „  ⁄„·Ì… «·Õ–› »‰Ã«Õ", vbOKOnly + vbMsgBoxRight, App.title
                Else
                    MsgBox "Delete Success...", vbOKOnly + vbMsgBoxRight, App.title
                End If
                FillGridWithData
                BtnNext_Click
            End If
        End If
    End If
    Exit Sub
ErrTrap:
    Select Case Err.Number
        Case -2147217873, -2147467259
            If SystemOptions.UserInterface = ArabicInterface Then
                StrMSG = "⁄›Ê« ·« ÌÃÊ“ Õ–› «·”Ã· ·«— »«ÿÂ »»Ì«‰«  √Œ—Ì"
            Else
                StrMSG = "Sorry .. can't Delete this record , Reason : Data integrity"
            End If
            RsSavRec.CancelUpdate
            MsgBox StrMSG, vbOKOnly + vbMsgBoxRight, App.title
            Cn.Errors.Clear
    End Select
End Sub
Private Sub BtnFirst_Click()

    On Error GoTo ErrTrap

    Dim Msg As String

    If Me.TxtModFlg.Text = "N" Then
        FindRec val(Me.TxtUnitID.Text)
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
                Msg = "Sorry.. this record Already Deleted" & Chr(13)
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
        FindRec val(Me.TxtUnitID.Text)
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
                Msg = "Sorry.. this record Already Deleted" & Chr(13)
            End If
            MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading + vbOKOnly, App.title
            RsSavRec.Requery
            Resume BegnieWork
    End Select
End Sub
Private Sub btnModify_Click()

    On Error GoTo ErrTrap

    Dim Msg As String

    If TxtUnitID.Text <> "" Then
        TxtModFlg = "E"
        Frm2.Enabled = True
        Me.TxtVacName.SetFocus
        CuurentLogdata
    End If

    Exit Sub
ErrTrap:
    Select Case Err.Number
        Case -2147467259
            If SystemOptions.UserInterface = ArabicInterface Then
                Msg = "⁄›Ê«" & Chr(13)
                Msg = Msg & " ·«Ì„ﬂ‰  ⁄œÌ· Â–« «·”Ã· ›Ï «·Êﬁ  «·Õ«·Ï" & Chr(13)
                Msg = Msg & "ÕÌÀ «‰Â ﬁÌœ «· ⁄œÌ· „‰ ﬁ»· „” Œœ„ «Œ— ⁄·Ï «·‘»ﬂ…"
            Else
                Msg = "Sorry" & Chr(13)
                Msg = Msg & " Can't Edit this record now - Another user work with it now" & Chr(13)
            End If
            MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading + vbOKOnly, App.title
            If RsSavRec.EditMode <> adEditNone Then
                RsSavRec.CancelUpdate
            End If
    End Select
End Sub
Private Sub btnNew_Click()

    On Error GoTo ErrTrap
    
    Dim My_SQL As String

    Dim RS As ADODB.Recordset
    Set RS = New ADODB.Recordset
    Frm2.Enabled = True
    clear_all Me
    TxtModFlg.Text = "N"
    My_SQL = "TblUnites"
    RS.Open My_SQL, Cn, adOpenKeyset, adLockOptimistic, adCmdTable
    If RS.RecordCount > 0 Then
        TxtUnitID.Text = RS.RecordCount + 1
    Else
        TxtUnitID.Text = 1
    End If
    RS.Close
    CmbType.ListIndex = 0
    TxtVacName.SetFocus
ErrTrap:
End Sub
Private Sub BtnNext_Click()

    On Error GoTo ErrTrap
    
    Dim Msg As String

    If Me.TxtModFlg.Text = "N" Then
        FindRec val(Me.TxtUnitID.Text)
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
                Msg = "Sorry , this Recored was deleted by other user on the network" & Chr(13)
                Msg = Msg & "Date will be updated now" & Chr(13)
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
        FindRec val(Me.TxtUnitID.Text)
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
                Msg = "Sorry , this Recored was deleted by other user on the network" & Chr(13)
                Msg = Msg & "Date will be updated now" & Chr(13)
            End If
            MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading + vbOKOnly, App.title
            RsSavRec.Requery
            Resume BegnieWork
    End Select
End Sub
Private Sub btnSave_Click()

    On Error GoTo ErrTrap
    
    Dim Msg As String
    Dim StrVacCode As String
    Dim StrVacName As String
    Dim CtrlTxt As Control

    If Trim(Me.TxtVacName.Text) = "" Then
        If SystemOptions.UserInterface = ArabicInterface Then
            Msg = "ÌÃ» ﬂ «»… «”„ «·ÊÕœ… ...!!!"
        Else
            Msg = "Please Enter The name"
        End If
        MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
        TxtVacName.SetFocus
        Exit Sub
    End If
    StrVacName = IsRecExist("TblUnites", "UnitName", Trim(TxtVacName.Text), "UnitName", "UnitID<>'" & Trim(TxtUnitID.Text) & "'")

    If StrVacName <> "" Then
        If SystemOptions.UserInterface = ArabicInterface Then
            Msg = "·ﬁœ ”»ﬁ  ”ÃÌ· Â–Â «·ÊÕœ… „‰ ﬁ»·"
        Else
            Msg = "this Unit Already Exist"
        End If

        MsgBox Msg, vbOKOnly + vbMsgBoxRight, App.title
        TxtVacName.SetFocus
        Exit Sub
    End If
    Select Case Me.TxtModFlg.Text
        Case "N"
            AddNewRec
            BtnLast_Click

        Case "E"
            FiLLRec
    End Select
    Exit Sub
ErrTrap:
    If SystemOptions.UserInterface = ArabicInterface Then
        MsgBox "Â‰«ﬂ Œÿ√ „« ›Ì ≈œŒ«· «·»Ì«‰« ", vbOKOnly + vbMsgBoxRight, App.title
    Else
        MsgBox "Error in Enterd data", vbOKOnly + vbMsgBoxRight, App.title
    End If
End Sub
Private Sub BtnUndo_Click()
    FindRec val(TxtUnitID.Text)
    Me.TxtModFlg.Text = "R"
End Sub
Private Sub BtnUpdate_Click()

    On Error GoTo ErrTrap
    
    Dim Msg As String
    Dim FristCount As Long
    Dim LastCount As Long
    
    FristCount = RsSavRec.RecordCount
    RsSavRec.Requery
    LastCount = RsSavRec.RecordCount
    BtnUndo_Click
    If SystemOptions.UserInterface = ArabicInterface Then
        If FristCount = LastCount Then
            Msg = "·«  ÊÃœ »Ì«‰«  ÃœÌœ…"
        Else
            Msg = "⁄œœ «·”Ã·«  ﬁ»· «· ÕœÌÀ" & vbCrLf & FristCount & vbCrLf & "⁄œœ «·”Ã·«  »⁄œ «· ÕœÌÀ" & vbCrLf & LastCount
            If LastCount > FristCount Then
                Msg = Msg + vbCrLf & "⁄œœ «·”Ã·«  «·ÃœÌœ…" & vbCrLf & LastCount - FristCount
            Else
                Msg = Msg + vbCrLf & "⁄œœ «·”Ã·«  «·„Õ–Ê›…" & vbCrLf & FristCount - LastCount
            End If
        End If
    Else
    If FristCount = LastCount Then
            Msg = "No new data"
        Else
            Msg = "No. of records before the update" & vbCrLf & FristCount & vbCrLf & "⁄œœ «·”Ã·«  »⁄œ «· ÕœÌÀ" & vbCrLf & LastCount
        
            If LastCount > FristCount Then
                Msg = Msg + vbCrLf & "No. of new records" & vbCrLf & LastCount - FristCount
            Else
                Msg = Msg + vbCrLf & "No. of deleted records" & vbCrLf & FristCount - LastCount
            End If
        End If
    End If
    MsgBox Msg, vbOKOnly + vbMsgBoxRight + vbInformation, App.title
ErrTrap:
End Sub
Public Sub AddNewRec()

    On Error GoTo ErrTrap
    
    Dim StrRecID As String
    
    StrRecID = new_id("TblUnites", "UnitID", "")
    RsSavRec.AddNew
    RsSavRec.Fields("UnitID").value = IIf(StrRecID <> "", StrRecID, Null)
    FiLLRec
ErrTrap:
End Sub
Public Sub FiLLRec()

    On Error GoTo ErrTrap

    RsSavRec.Fields("UnitName").value = IIf(TxtVacName.Text <> "", Trim(TxtVacName.Text), Null)
    RsSavRec.Fields("UnitNamee").value = IIf(TxtVacNamee.Text <> "", Trim(TxtVacNamee.Text), Null)

    RsSavRec.update
    CuurentLogdata
    If SystemOptions.UserInterface = ArabicInterface Then
        MsgBox " „  ⁄„·Ì… «·Õ›Ÿ »‰Ã«Õ", vbOKOnly + vbMsgBoxRight, App.title
    Else
        MsgBox "Saved Successfully", vbOKOnly + vbMsgBoxRight, App.title
    End If
    FillGridWithData
    TxtModFlg = "R"
    Exit Sub
ErrTrap:
    If RsSavRec.EditMode <> adEditNone Then
        RsSavRec.CancelUpdate
    End If
End Sub
Public Sub FiLLTXT()

    On Error GoTo ErrTrap
    
    Dim i As Integer
    
    Frm2.Enabled = False
    TxtUnitID.Text = IIf(IsNull(RsSavRec.Fields("UnitID").value), "", RsSavRec.Fields("UnitID").value)
    TxtVacName.Text = IIf(IsNull(RsSavRec.Fields("UnitName").value), "", RsSavRec.Fields("UnitName").value)
    TxtVacNamee.Text = IIf(IsNull(RsSavRec.Fields("UnitNamee").value), "", RsSavRec.Fields("UnitNamee").value)
    LabCurrRec.Caption = RsSavRec.AbsolutePosition
    LabCountRec.Caption = RsSavRec.RecordCount
    With Grid
        For i = 1 To .Rows - 1
            If Trim(TxtUnitID.Text) = .TextMatrix(i, .ColIndex("UnitID")) Then
                .Row = i
                Exit Sub
            End If
        Next
    End With
ErrTrap:
End Sub
Public Sub EditRec(StrTable As String, RecId As String)
    FiLLRec
End Sub
Private Sub Grid_EnterCell()
    On Error GoTo ErrTrap
    FindRec val(Me.Grid.TextMatrix(Me.Grid.Row, Me.Grid.ColIndex("UnitID")))
ErrTrap:
End Sub
Private Sub TxtDis_Count_KeyPress(KeyAscii As Integer)
    KeyAscii = DataFormat(CurOnly, KeyAscii)
End Sub
Private Sub ISButton1_Click()
    print_report
End Sub
Private Sub TxtUnitID_Change()
    Dim TxtMod As String
    TxtMod = TxtModFlg.Text
    TxtModFlg.Text = ""
    TxtModFlg = TxtMod
End Sub
Public Function FindRec(ByVal RecId As Long)

    On Error GoTo ErrTrap
    
    RsSavRec.Find "UnitID=" & RecId, , adSearchForward, 1

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
Private Sub TxtModFlg_Change()
    If TxtModFlg.Text = "N" Then
        Frm2.Enabled = True
        Me.btnNew.Enabled = False
        btnModify.Enabled = False
        btnDelete.Enabled = False
        Me.btnQuery.Enabled = False
        Grid.Enabled = False
        BtnUndo.Enabled = True
        Me.btnSave.Enabled = True
        BtnUpdate.Enabled = False
    ElseIf TxtModFlg.Text = "R" Then
        Frm2.Enabled = False
        Grid.Enabled = True
        btnModify.Enabled = False
        btnDelete.Enabled = False
        If TxtUnitID.Text <> "" Then
            btnModify.Enabled = True
            btnDelete.Enabled = True
        End If
        BtnUpdate.Enabled = True
        Me.btnQuery.Enabled = True
        Me.btnNew.Enabled = True
        BtnUndo.Enabled = False
        Me.btnSave.Enabled = False
        btnNext.Enabled = True
        btnPrevious.Enabled = True
        btnFirst.Enabled = True
        btnLast.Enabled = True
    ElseIf TxtModFlg.Text = "E" Then
        Frm2.Enabled = True
        Me.btnNew.Enabled = False
        btnModify.Enabled = False
        btnDelete.Enabled = False
        Me.btnQuery.Enabled = False
        BtnUpdate.Enabled = False
        BtnUndo.Enabled = True
        Me.btnSave.Enabled = True
        Grid.Enabled = False
        btnNext.Enabled = False
        btnPrevious.Enabled = False
        btnFirst.Enabled = False
        btnLast.Enabled = False
    End If

End Sub
Public Sub FillGridWithData()

    On Error GoTo ErrTrap

    Dim i As Integer
    Dim RS As ADODB.Recordset
    Dim My_SQL As String

    Set RS = New ADODB.Recordset
    My_SQL = "select * From TblUnites order by UnitID"
    RS.Open My_SQL, Cn, adOpenKeyset, adLockOptimistic, adCmdText
    With Me.Grid
        .Rows = 2
        .Clear flexClearScrollable
        If RS.RecordCount > 0 Then
            .Rows = RS.RecordCount + 1
            RS.MoveFirst
            For i = 1 To .Rows - 1
                .TextMatrix(i, .ColIndex("Ser")) = i
                .TextMatrix(i, .ColIndex("UnitName")) = IIf(IsNull(RS.Fields("UnitName").value), "", RS.Fields("UnitName").value)
                .TextMatrix(i, .ColIndex("UnitNamee")) = IIf(IsNull(RS.Fields("UnitNamee").value), "", RS.Fields("UnitNamee").value)
                .TextMatrix(i, .ColIndex("UnitID")) = IIf(IsNull(RS.Fields("UnitID").value), "", RS.Fields("UnitID").value)
                RS.MoveNext
            Next
            RS.Close
        End If
        .RowHeight(-1) = 300
    End With
ErrTrap:
End Sub
Function CuurentLogdata(Optional Currentmode As String)
    LogTextA = "    ‘«‘… " & ScreenNameArabic & Chr(13) & " —ﬁ„ «·ÊÕœ…   " & TxtUnitID.Text & Chr(13) & "  «”„ «·ÊÕœ… " & TxtVacName.Text
    LogTextE = "    Screen  " & ScreenNameEnglish & Chr(13) & " Unit No   " & TxtUnitID.Text & Chr(13) & " Unit Name" & TxtVacNamee.Text
    If Currentmode <> "D" Then
        AddToLogFile CInt(user_id), 0, Date, Time, LogTextA, LogTextE, Me.name, Me.TxtModFlg
    Else
        AddToLogFile CInt(user_id), 0, Date, Time, LogTextA, LogTextE, Me.name, "D"
    End If
End Function
Private Sub TxtVacName_GotFocus()
    SwitchKeyboardLang LANG_ARABIC
End Sub

Private Sub TxtVacNamee_GotFocus()
    SwitchKeyboardLang LANG_ENGLISH
End Sub
'#####################################################################################################################################################
'#####################################################################################################################################################
'#####################################################################################################################################################

Private Sub Cmd_Click(Index As Integer)

    On Error GoTo ErrTrap

    Select Case Index
        Case 0
            TxtModFlg1.Text = "N"
            clear_all Me
            XPTxtBoxID.Text = CStr(new_id("TblItemsColors", "ColorID", "", True))
            XPTxtBoxName.SetFocus
        Case 1
            TxtModFlg1.Text = "E"
            CuurentLogdata1
        Case 2
            SaveData
        Case 3
            Call Undo
        Case 4
            Del_Company
        Case 5
        Case 6
            Unload Me
    End Select
    Exit Sub
ErrTrap:
End Sub
Private Sub CmdHelp_Click()
    SystemOptions.SysHelp.HHTopicID = Me.HelpContextID
    SystemOptions.SysHelp.HHDisplayTopicID Me.hWnd
End Sub

Function CuurentLogdata1(Optional Currentmode As String)
    LogTextA = "    ‘«‘… " & ScreenNameArabic & Chr(13) & " —ﬁ„ «··Ê‰   " & XPTxtBoxID.Text & Chr(13) & "  «”„ «··Ê‰ " & XPTxtBoxName.Text & Chr(13) & "  „·«ÕŸ«  " & XPMTxtRemark.Text
    LogTextE = "    Screen  " & ScreenNameEnglish & Chr(13) & " Color No   " & XPTxtBoxID.Text & Chr(13) & " Color Name" & XPTxtBoxName.Text & Chr(13) & "  Remarks " & XPMTxtRemark.Text
    If Currentmode <> "D" Then
        AddToLogFile CInt(user_id), 0, Date, Time, LogTextA, LogTextE, Me.name, Me.TxtModFlg1
    Else
        AddToLogFile CInt(user_id), 0, Date, Time, LogTextA, LogTextE, Me.name, "D"
    End If
End Function
Private Sub TxtModFlg1_Change()

    On Error GoTo ErrTrap

    Select Case Me.TxtModFlg1.Text
        Case "R"
            Me.Cmd(2).Enabled = False
            Me.Cmd(3).Enabled = False
            Me.Cmd(0).Enabled = True
            Me.Cmd(1).Enabled = True
            Me.Cmd(4).Enabled = True
            Me.XPBtnMove(0).Enabled = True
            Me.XPBtnMove(1).Enabled = True
            Me.XPBtnMove(2).Enabled = True
            Me.XPBtnMove(3).Enabled = True
            Me.XPTxtBoxID.locked = True
            Me.XPTxtBoxName.locked = True
            Me.XPMTxtRemark.locked = True
            If RS.RecordCount < 1 Then
                Me.XPBtnMove(0).Enabled = False
                Me.XPBtnMove(1).Enabled = False
                Me.XPBtnMove(2).Enabled = False
                Me.XPBtnMove(3).Enabled = False
                Me.Cmd(1).Enabled = False
                Me.Cmd(4).Enabled = False
            End If
        Case "N"
            Me.Cmd(2).Enabled = True
            Me.Cmd(3).Enabled = True
            Me.Cmd(0).Enabled = False
            Me.Cmd(1).Enabled = False
            Me.Cmd(4).Enabled = False
            Me.XPTxtBoxID.locked = True
            Me.XPTxtBoxName.locked = False
            Me.XPMTxtRemark.locked = False
        Case "E"
            Me.Cmd(2).Enabled = True
            Me.Cmd(3).Enabled = True
            Me.Cmd(0).Enabled = False
            Me.Cmd(1).Enabled = False
            Me.Cmd(4).Enabled = False
            Me.XPBtnMove(0).Enabled = False
            Me.XPBtnMove(1).Enabled = False
            Me.XPBtnMove(2).Enabled = False
            Me.XPBtnMove(3).Enabled = False
            Me.XPTxtBoxID.locked = True
            Me.XPTxtBoxName.locked = False
            Me.XPMTxtRemark.locked = False
    End Select
    Exit Sub
ErrTrap:
End Sub
Public Sub Retrive(Optional Lngid As Long = 0)

    On Error GoTo ErrTrap

    If RS.RecordCount < 1 Then
        XPTxtCurrent.Caption = 0
        XPTxtCount.Caption = 0
        Exit Sub
    End If
    XPTxtBoxID.Text = IIf(IsNull(RS("ColorID").value), "", val(RS("ColorID").value))
    XPTxtBoxName.Text = IIf(IsNull(RS("ColorName").value), "", Trim(RS("ColorName").value))
    XPMTxtRemark.Text = IIf(IsNull(RS("ColorComment").value), "", Trim(RS("ColorComment").value))
    XPTxtCurrent.Caption = RS.AbsolutePosition
    XPTxtCount.Caption = RS.RecordCount
    Exit Sub
ErrTrap:
End Sub
Private Sub XPBtnMove_Click(Index As Integer)

    On Error GoTo ErrTrap

    If Me.TxtModFlg1.Text = "N" Then
        clear_all Me
        Me.TxtModFlg1.Text = "R"
        XPBtnMove_Click (1)
    End If
    Select Case Index
        Case 0
            If Not (RS.EOF Or RS.BOF) Then
                RS.MovePrevious

                If RS.BOF Then RS.MoveFirst
            End If
        Case 1
            If Not (RS.EOF Or RS.BOF) Then
                RS.MoveFirst
            End If
        Case 2
            If Not (RS.EOF Or RS.BOF) Then
                RS.MoveLast
            End If
        Case 3
            If Not (RS.EOF Or RS.BOF) Then
                RS.MoveNext
                If RS.EOF Then RS.MoveLast
            End If
    End Select
    Retrive
    Exit Sub
ErrTrap:
End Sub
Private Sub SaveData()
    Dim Msg As String
    Dim StrSQL As String
    Dim RsTemp As New ADODB.Recordset
    Dim RsTempM As New ADODB.Recordset
    Dim BeginTrans As Boolean
    
    On Error GoTo ErrTrap
    
    If Me.TxtModFlg1.Text <> "R" Then
        If XPTxtBoxName.Text = "" Then
            If SystemOptions.UserInterface = ArabicInterface Then
                MsgBox "„‰ ›÷·ﬂ √œŒ· «”„ «··Ê‰ ", vbOKOnly + vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
            Else
                MsgBox "Please enter the name", vbOKOnly + vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
            End If
            XPTxtBoxName.SetFocus
            Exit Sub
        End If
        Select Case Me.TxtModFlg1.Text
            Case "N"
                StrSQL = "select * from  TblItemsColors where ColorName ='" & Trim(XPTxtBoxName.Text) & "'"
                RsTemp.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText
                If RsTemp.RecordCount > 0 Then
                    If SystemOptions.UserInterface = ArabicInterface Then
                        Msg = "Â‰«ﬂ ·Ê‰ „”Ã· „”»ﬁ« »Â–« «·«”„" & Chr(13)
                        Msg = Msg + "»—Ã«¡ «· √ﬂœ „‰ «·«”„ «·’ÕÌÕ " & Chr(13)
                        Msg = Msg + "√Ê  €ÌÌ— √Ê  „ÌÌ“ «”„ «··Ê‰"
                    Else
                        Msg = "This record already exists"
                    End If
                    MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
                    XPTxtBoxName.SetFocus
                    Exit Sub
                End If
            Case "E"
                StrSQL = "select * from  TblItemsColors where ColorName='" & Trim(XPTxtBoxName.Text) & "'"
                RsTemp.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText
                If RsTemp.RecordCount > 0 Then
                    If RsTemp("ColorID").value <> val(XPTxtBoxID.Text) Then
                        If SystemOptions.UserInterface = ArabicInterface Then
                            Msg = "Â‰«ﬂ ·Ê‰  „”Ã· „”»ﬁ« »Â–« «·«”„" & Chr(13)
                            Msg = Msg + "»—Ã«¡ «· √ﬂœ „‰ «·«”„ «·’ÕÌÕ " & Chr(13)
                            Msg = Msg + "√Ê  €ÌÌ— √Ê  „ÌÌ“ «”„ «··Ê‰"
                        Else
                            Msg = "This record already exists"
                        End If
                        MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
                        XPTxtBoxName.SetFocus
                        Exit Sub
                    End If
                End If
        End Select
        Cn.BeginTrans
        BeginTrans = True
        Select Case Me.TxtModFlg1.Text
            Case "N"
                RS.AddNew
                RS("ColorID").value = val(XPTxtBoxID.Text)
            Case "E"
                If RS("ColorID").value <> val(Me.XPTxtBoxID.Text) Then
                    RS.Find "ColorID=" & val(Me.XPTxtBoxID.Text), , adSearchForward, 1

                    If RS.EOF Or RS.EOF Then
                        Exit Sub
                    End If
                End If
        End Select
        RS("ColorName").value = Trim(XPTxtBoxName.Text)
        RS("ColorComment").value = IIf(XPMTxtRemark.Text = "", Null, Trim(XPMTxtRemark.Text))
        RS.update
        Cn.CommitTrans
        BeginTrans = False
        XPTxtCurrent.Caption = RS.AbsolutePosition
        XPTxtCount.Caption = RS.RecordCount
        CuurentLogdata1
        Select Case Me.TxtModFlg1.Text
            Case "N"
                If SystemOptions.UserInterface = ArabicInterface Then
                    Msg = "  „ Õ›Ÿ »Ì«‰«  Â–« «··Ê‰" & Chr(13)
                    Msg = Msg + "Â·  —€» ›Ì ≈÷«›… »Ì«‰«  √Œ—Ì"
                Else
                    Msg = "Recored saved successfully , do you want to add another recored"
                End If
                If MsgBox(Msg, vbYesNo + vbQuestion + vbMsgBoxRight + vbMsgBoxRtlReading + vbDefaultButton2, App.title) = vbYes Then
                    Cmd_Click (0)
                    Exit Sub
                End If
            Case "E"
                If SystemOptions.UserInterface = ArabicInterface Then
                    MsgBox " „ Õ›Ÿ Â–Â «· ⁄œÌ·« ", vbInformation + vbOKOnly + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
                Else
                    MsgBox "Record edited successfully", vbInformation + vbOKOnly + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
                End If
        End Select
        TxtModFlg1.Text = "R"
    End If

    Exit Sub
ErrTrap:
    If BeginTrans = True Then
        BeginTrans = False
        Cn.RollbackTrans
    End If
    If Err.Number = -2147217900 Then
        If SystemOptions.UserInterface = ArabicInterface Then
            Msg = "·« Ì„ﬂ‰ Õ›Ÿ Â–Â «·»Ì«‰«  " & Chr(13)
            Msg = Msg + "·ﬁœ  „ «œŒ«· ﬁÌ„ €Ì— ’«·Õ… " & Chr(13)
            Msg = Msg + " √ﬂœ „‰ œﬁ… «·»Ì«‰«  Ê√⁄œ «·„Õ«Ê·…"
        Else
            Msg = "Data Can't be saved " & Chr(13)
            Msg = Msg & "Invalid data values was entered" & Chr(13)
            Msg = Msg & "Please make sure of the entered data and try again"
        End If
        MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
        Exit Sub
    End If
    If SystemOptions.UserInterface = ArabicInterface Then
        Msg = "⁄›Ê«...ÕœÀ Œÿ√ „« √À‰«¡ Õ›Ÿ Â–Â «·»Ì«‰«  " & Chr(13)
    Else
        Msg = "Sorry , somthing went wrong while saving data" & Chr(13)
    End If
    MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
End Sub
Private Sub Undo()

    On Error GoTo ErrTrap

    Select Case TxtModFlg1.Text
        Case "N"
            clear_all Me
            Me.TxtModFlg1.Text = "R"
            XPBtnMove_Click (1)
        Case "E"
            RS.Find "BoxID='" & val(XPTxtBoxID.Text) & "'", , adSearchForward, adBookmarkFirst
            If RS.EOF Or RS.BOF Then
                Me.TxtModFlg1.Text = "R"
                Exit Sub
            End If
            Retrive
            Me.TxtModFlg1.Text = "R"
    End Select
    Exit Sub
ErrTrap:
End Sub
Private Sub Del_Company()
    Dim Msg As String
    Dim StrSQL As String
    Dim RsTemp As New ADODB.Recordset
    On Error GoTo ErrTrap
    If XPTxtBoxID.Text <> "" Then
        If val(Me.XPTxtBoxID.Text) = 1 Then
            If SystemOptions.UserInterface = ArabicInterface Then
                Msg = "·«Ì„ﬂ‰ Õ–› Â–« «·”Ã·...!!!"
            Else
                Msg = "sorry, this record cannot be deleted "
            End If
            MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
            Exit Sub
        End If
        StrSQL = "select * from Transaction_Details where ColorID=" & Trim(XPTxtBoxID.Text)
        RsTemp.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText
        If Not (RsTemp.EOF Or RsTemp.BOF) Then
            If SystemOptions.UserInterface = ArabicInterface Then
                Msg = "·« Ì„ﬂ‰ Õ–› »Ì«‰«  Â–« «··Ê‰" & Chr(13)
                Msg = Msg + "Â‰«ﬂ »⁄÷ «·⁄„·Ì«  „— »ÿ… »Â–« «··Ê‰"
            Else
                Msg = "sorry, this record cannot be deleted due to data integration"
            End If
            MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
            Exit Sub
        End If
        If SystemOptions.UserInterface = ArabicInterface Then
            Msg = "”Ì „ Õ–› »Ì«‰«  «··Ê‰ —ﬁ„ " & Chr(13)
            Msg = Msg + (XPTxtBoxID.Text) & Chr(13)
            Msg = Msg + " Â·  —€» ›Ì Õ–› Â–Â «·»Ì«‰« ø"
        Else
            Msg = "Are you sure you want to delete this record"
        End If
        If MsgBox(Msg, vbYesNo + vbQuestion + vbMsgBoxRight + vbMsgBoxRtlReading, App.title) = vbYes Then
            If Not RS.RecordCount < 1 Then
                CuurentLogdata1 ("D")
                RS.Delete
                RS.MoveFirst
                If RS.RecordCount < 1 Then
                    clear_all Me
                    TxtModFlg1_Change
                    XPTxtCurrent.Caption = 0
                    XPTxtCount.Caption = 0
                Else
                    Retrive
                End If
            End If
        End If
    Else
        clear_all Me
        If SystemOptions.UserInterface = ArabicInterface Then
            Msg = "Â–Â «·⁄„·Ì… €Ì— „ «Õ… ÕÌÀ √‰Â ·«ÌÊÃœ √Ì ”Ã·« "
        Else
            Msg = "This operation is not available due to lack of records"
        End If
        MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
        TxtModFlg1_Change
        Exit Sub
    End If
    TxtModFlg1_Change
    Exit Sub
ErrTrap:
    If Err.Number = -2147217887 Then
        If SystemOptions.UserInterface = ArabicInterface Then
            Msg = "·«Ì„ﬂ‰ Õ–› Â–« «·”Ã· · ﬂ«„· «·»Ì«‰«  " & Chr(13) & " ÊÃœ »Ì«‰«  „— »ÿ… »Â–« «··Ê‰ "
            Msg = Msg & Chr(13) & Err.description
        Else
            Msg = "sorry, this record cannot be deleted due to data integration"
        End If
        MsgBox Msg, vbMsgBoxRight + vbMsgBoxRtlReading + vbExclamation, App.title
        RS.CancelUpdate
    End If

End Sub
'2
'#####################################################################################################################################################
'#####################################################################################################################################################
'#####################################################################################################################################################
Private Sub Cmd2_Click(Index As Integer)

    'On Error GoTo ErrTrap

    Select Case Index
        Case 0
            TxtModFlg2.Text = "N"
            clear_all Me
            XPTxtBoxID2.Text = CStr(new_id("TblItemsSizes", "SizeId", "", True))
            XPTxtBoxName2.SetFocus
        Case 1
            TxtModFlg2.Text = "E"
            CuurentLogdata2
        Case 2
            SaveData2
        Case 3
            Call Undo2
        Case 4
            Del_Company2
        Case 5
        Case 6
            Unload Me
    End Select
    Exit Sub
ErrTrap:
End Sub
Private Sub CmdHelp2_Click()
    SystemOptions.SysHelp.HHTopicID = Me.HelpContextID
    SystemOptions.SysHelp.HHDisplayTopicID Me.hWnd
End Sub
Function CuurentLogdata2(Optional Currentmode As String)
    LogTextA = "    ‘«‘… " & ScreenNameArabic & Chr(13) & " —ﬁ„ «·„ﬁ«”   " & XPTxtBoxID2.Text & Chr(13) & "  «”„ «·„ﬁ«” " & XPTxtBoxName2.Text & Chr(13) & "  „·«ÕŸ«  " & XPMTxtRemark2.Text
    LogTextE = "    Screen  " & ScreenNameEnglish & Chr(13) & " Size No   " & XPTxtBoxID2.Text & Chr(13) & " Size Name" & XPTxtBoxName2.Text & Chr(13) & "  Remarks " & XPMTxtRemark2.Text
    If Currentmode <> "D" Then
        AddToLogFile CInt(user_id), 0, Date, Time, LogTextA, LogTextE, Me.name, Me.TxtModFlg2
    Else
        AddToLogFile CInt(user_id), 0, Date, Time, LogTextA, LogTextE, Me.name, "D"
    End If
End Function

Private Sub TxtModFlg2_Change()

    On Error GoTo ErrTrap
    
    Select Case Me.TxtModFlg2.Text
        Case "R"
            Me.Cmd2(2).Enabled = False
            Me.Cmd2(3).Enabled = False
            Me.Cmd2(0).Enabled = True
            Me.Cmd2(1).Enabled = True
            Me.Cmd2(4).Enabled = True
            Me.XPBtnMove2(0).Enabled = True
            Me.XPBtnMove2(1).Enabled = True
            Me.XPBtnMove2(2).Enabled = True
            Me.XPBtnMove2(3).Enabled = True
            Me.XPTxtBoxID2.locked = True
            Me.XPTxtBoxName2.locked = True
            Me.XPMTxtRemark2.locked = True
            If Rs2.RecordCount < 1 Then
                Me.XPBtnMove2(0).Enabled = False
                Me.XPBtnMove2(1).Enabled = False
                Me.XPBtnMove2(2).Enabled = False
                Me.XPBtnMove2(3).Enabled = False
                Me.Cmd2(1).Enabled = False
                Me.Cmd2(4).Enabled = False
            End If
        Case "N"
            Me.Cmd2(2).Enabled = True
            Me.Cmd2(3).Enabled = True
            Me.Cmd2(0).Enabled = False
            Me.Cmd2(1).Enabled = False
            Me.Cmd2(4).Enabled = False
            Me.XPTxtBoxID2.locked = True
            Me.XPTxtBoxName2.locked = False
            Me.XPMTxtRemark2.locked = False
        Case "E"
            Me.Cmd2(2).Enabled = True
            Me.Cmd2(3).Enabled = True
            Me.Cmd2(0).Enabled = False
            Me.Cmd2(1).Enabled = False
            Me.Cmd2(4).Enabled = False
            Me.XPBtnMove2(0).Enabled = False
            Me.XPBtnMove2(1).Enabled = False
            Me.XPBtnMove2(2).Enabled = False
            Me.XPBtnMove2(3).Enabled = False
            Me.XPTxtBoxID2.locked = True
            Me.XPTxtBoxName2.locked = False
            Me.XPMTxtRemark2.locked = False
    End Select

    Exit Sub
ErrTrap:
End Sub
Public Sub Retrive2(Optional Lngid As Long = 0)

    On Error GoTo ErrTrap

    If Rs2.RecordCount < 1 Then
        XPTxtCurrent2.Caption = 0
        XPTxtCount2.Caption = 0
        Exit Sub
    End If
    XPTxtBoxID2.Text = IIf(IsNull(Rs2("SizeId").value), "", val(Rs2("SizeId").value))
    XPTxtBoxName2.Text = IIf(IsNull(Rs2("SizeName").value), "", Trim(Rs2("SizeName").value))
    XPMTxtRemark2.Text = IIf(IsNull(Rs2("SizeComment").value), "", Trim(Rs2("SizeComment").value))
    XPTxtCurrent2.Caption = Rs2.AbsolutePosition
    XPTxtCount2.Caption = Rs2.RecordCount
    Exit Sub
ErrTrap:
End Sub
Private Sub XPBtnMove2_Click(Index As Integer)

    On Error GoTo ErrTrap

    If Me.TxtModFlg2.Text = "N" Then
        clear_all Me
        Me.TxtModFlg2.Text = "R"
        XPBtnMove_Click (1)
    End If
    Select Case Index
        Case 0
            If Not (Rs2.EOF Or Rs2.BOF) Then
                Rs2.MovePrevious

                If Rs2.BOF Then Rs2.MoveFirst
            End If
        Case 1
            If Not (Rs2.EOF Or Rs2.BOF) Then
                Rs2.MoveFirst
            End If
        Case 2
            If Not (Rs2.EOF Or Rs2.BOF) Then
                Rs2.MoveLast
            End If
        Case 3
            If Not (Rs2.EOF Or Rs2.BOF) Then
                Rs2.MoveNext
                If Rs2.EOF Then Rs2.MoveLast
            End If
    End Select
    Retrive2
    Exit Sub
ErrTrap:
End Sub
Private Sub SaveData2()

    Dim Msg As String
    Dim Strs2QL As String
    Dim rs2Temp As New ADODB.Recordset
    Dim rs2TempM As New ADODB.Recordset
    Dim BeginTrans As Boolean
    
    On Error GoTo ErrTrap

    If Me.TxtModFlg2.Text <> "R" Then
        If XPTxtBoxName2.Text = "" Then
            If SystemOptions.UserInterface = ArabicInterface Then
                MsgBox "„‰ ›÷·ﬂ √œŒ· «”„ «·„ﬁ«” ", vbOKOnly + vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
            Else
                MsgBox "Please enter the name", vbOKOnly + vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
            End If
            XPTxtBoxName2.SetFocus
            Exit Sub
        End If
        Select Case Me.TxtModFlg2.Text
            Case "N"
                Strs2QL = "select * from  TblItemsSizes where SizeName ='" & Trim(XPTxtBoxName2.Text) & "'"
                rs2Temp.Open Strs2QL, Cn, adOpenStatic, adLockOptimistic, adCmdText
                If rs2Temp.RecordCount > 0 Then
                    If SystemOptions.UserInterface = ArabicInterface Then
                        Msg = "Â‰«ﬂ „ﬁ«” „”Ã· „”»ﬁ« »Â–« «·«”„" & Chr(13)
                        Msg = Msg + "»—Ã«¡ «· √ﬂœ „‰ «·«”„ «·’ÕÌÕ " & Chr(13)
                        Msg = Msg + "√Ê  €ÌÌ— √Ê  „ÌÌ“ «”„ «·„ﬁ«”"
                    Else
                        Msg = "This record already exists"
                    End If
                    MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
                    XPTxtBoxName2.SetFocus
                    Exit Sub
                End If
            Case "E"
                Strs2QL = "select * from  TblItemsSizes where SizeName='" & Trim(XPTxtBoxName2.Text) & "'"
                rs2Temp.Open Strs2QL, Cn, adOpenStatic, adLockOptimistic, adCmdText
                If rs2Temp.RecordCount > 0 Then
                    If rs2Temp("SizeId").value <> val(XPTxtBoxID2.Text) Then
                        If SystemOptions.UserInterface = ArabicInterface Then
                            Msg = "Â‰«ﬂ „ﬁ«”  „”Ã· „”»ﬁ« »Â–« «·«”„" & Chr(13)
                            Msg = Msg + "»—Ã«¡ «· √ﬂœ „‰ «·«”„ «·’ÕÌÕ " & Chr(13)
                            Msg = Msg + "√Ê  €ÌÌ— √Ê  „ÌÌ“ «”„ «·„ﬁ«”"
                        Else
                            Msg = "This record already exists"
                        End If
                        MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
                        XPTxtBoxName2.SetFocus
                        Exit Sub
                    End If
                End If
        End Select
        Cn.BeginTrans
        BeginTrans = True
        Select Case Me.TxtModFlg2.Text
            Case "N"
                Rs2.AddNew
                Rs2("SizeId").value = val(XPTxtBoxID2.Text)
            Case "E"
                If Rs2("SizeId").value <> val(Me.XPTxtBoxID2.Text) Then
                    Rs2.Find "SizeId=" & val(Me.XPTxtBoxID2.Text), , adSearchForward, 1
                    If Rs2.EOF Or Rs2.EOF Then
                        Exit Sub
                    End If
                End If
        End Select
        Rs2("SizeName").value = Trim(XPTxtBoxName2.Text)
        Rs2("SizeComment").value = IIf(XPMTxtRemark2.Text = "", Null, Trim(XPMTxtRemark2.Text))
        Rs2.update
        Cn.CommitTrans
        BeginTrans = False
        XPTxtCurrent2.Caption = Rs2.AbsolutePosition
        XPTxtCount2.Caption = Rs2.RecordCount
        CuurentLogdata2
        Select Case Me.TxtModFlg2.Text
            Case "N"
                If SystemOptions.UserInterface = ArabicInterface Then
                    Msg = "  „ Õ›Ÿ »Ì«‰«  Â–« «·„ﬁ«”" & Chr(13)
                    Msg = Msg + "Â·  —€» ›Ì ≈÷«›… »Ì«‰«  √Œ—Ì"
                Else
                    Msg = "Recored saved successfully , do you want to add another recored"
                End If
                If MsgBox(Msg, vbYesNo + vbQuestion + vbMsgBoxRight + vbMsgBoxRtlReading + vbDefaultButton2, App.title) = vbYes Then
                    Cmd_Click (0)
                    Exit Sub
                End If
            Case "E"
                If SystemOptions.UserInterface = ArabicInterface Then
                    MsgBox " „ Õ›Ÿ Â–Â «· ⁄œÌ·« ", vbInformation + vbOKOnly + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
                Else
                    MsgBox "Record Edited successfully", vbInformation + vbOKOnly + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
                End If
        End Select
        TxtModFlg2.Text = "R"
    End If

    Exit Sub
ErrTrap:
    If BeginTrans = True Then
        BeginTrans = False
        Cn.RollbackTrans
    End If
    If Err.Number = -2147217900 Then
        If SystemOptions.UserInterface = ArabicInterface Then
            Msg = "·« Ì„ﬂ‰ Õ›Ÿ Â–Â «·»Ì«‰«  " & Chr(13)
            Msg = Msg + "·ﬁœ  „ «œŒ«· ﬁÌ„ €Ì— ’«·Õ… " & Chr(13)
            Msg = Msg + " √ﬂœ „‰ œﬁ… «·»Ì«‰«  Ê√⁄œ «·„Õ«Ê·…"
        Else
            Msg = "Data Can't be saved " & Chr(13)
            Msg = Msg & "Invalid data values was entered" & Chr(13)
            Msg = Msg & "Please make sure of the entered data and try again"
        End If
        MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
        Exit Sub
    End If
    If SystemOptions.UserInterface = ArabicInterface Then
        Msg = "⁄›Ê«...ÕœÀ Œÿ√ „« √À‰«¡ Õ›Ÿ Â–Â «·»Ì«‰«  " & Chr(13)
    Else
        Msg = "Sorry , somthing went wrong while saving data" & Chr(13)
    End If
    MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
End Sub
Private Sub Undo2()

    On Error GoTo ErrTrap

    Select Case TxtModFlg2.Text
        Case "N"
            clear_all Me
            Me.TxtModFlg2.Text = "R"
            XPBtnMove2_Click (1)
        Case "E"
            Rs2.Find "SizeId='" & val(XPTxtBoxID2.Text) & "'", , adSearchForward, adBookmarkFirst
            If Rs2.EOF Or Rs2.BOF Then
                Me.TxtModFlg2.Text = "R"
                Exit Sub
            End If
            Retrive2
            Me.TxtModFlg2.Text = "R"
    End Select

    Exit Sub
ErrTrap:
End Sub
Private Sub Del_Company2()

    Dim Msg As String
    Dim Strs2QL As String
    Dim rs2Temp As New ADODB.Recordset
    
    On Error GoTo ErrTrap

    If XPTxtBoxID2.Text <> "" Then
        If val(Me.XPTxtBoxID2.Text) = 1 Then
            If SystemOptions.UserInterface = ArabicInterface Then
                Msg = "·«Ì„ﬂ‰ Õ–› Â–« «·”Ã·...!!!"
            Else
                Msg = "sorry, this record cannot be deleted "
            End If
            MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
            Exit Sub
        End If
        Strs2QL = "select * from Transaction_Details where ItemSize=" & Trim(XPTxtBoxID2.Text)
        rs2Temp.Open Strs2QL, Cn, adOpenStatic, adLockOptimistic, adCmdText
        If Not (rs2Temp.EOF Or rs2Temp.BOF) Then
            If SystemOptions.UserInterface = ArabicInterface Then
                Msg = "·« Ì„ﬂ‰ Õ–› »Ì«‰«  Â–« «··Ê‰" & Chr(13)
                Msg = Msg + "Â‰«ﬂ »⁄÷ «·⁄„·Ì«  „— »ÿ… »Â–« «··Ê‰"
            Else
                Msg = "sorry, this record cannot be deleted due to data integration"
            End If
            MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
            Exit Sub
        End If
        If SystemOptions.UserInterface = ArabicInterface Then
            Msg = "”Ì „ Õ–› »Ì«‰«  «··Ê‰ —ﬁ„ " & Chr(13)
            Msg = Msg + (XPTxtBoxID2.Text) & Chr(13)
            Msg = Msg + " Â·  —€» ›Ì Õ–› Â–Â «·»Ì«‰« ø"
        Else
            Msg = "Are you sure you want to delete this record"
        End If
        If MsgBox(Msg, vbYesNo + vbQuestion + vbMsgBoxRight + vbMsgBoxRtlReading, App.title) = vbYes Then
            If Not Rs2.RecordCount < 1 Then
                Rs2.Delete
                CuurentLogdata2 ("D")
                Rs2.MoveFirst
                If Rs2.RecordCount < 1 Then
                    clear_all Me
                    TxtModFlg2_Change
                    XPTxtCurrent2.Caption = 0
                    XPTxtCount2.Caption = 0
                Else
                    Retrive2
                End If
            End If
        End If
    Else
        clear_all Me
        If SystemOptions.UserInterface = ArabicInterface Then
            Msg = "Â–Â «·⁄„·Ì… €Ì— „ «Õ… ÕÌÀ √‰Â ·«ÌÊÃœ √Ì ”Ã·« "
        Else
            Msg = "This operation is not available due to lack of records"
        End If
        MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
        TxtModFlg2_Change
        Exit Sub
    End If
    TxtModFlg2_Change
    Exit Sub
ErrTrap:
    If Err.Number = -2147217887 Then
        If SystemOptions.UserInterface = ArabicInterface Then
            Msg = "·«Ì„ﬂ‰ Õ–› Â–« «·”Ã· · ﬂ«„· «·»Ì«‰«  " & Chr(13) & " ÊÃœ »Ì«‰«  „— »ÿ… »Â–« «··Ê‰ "
            Msg = Msg & Chr(13) & Err.description
        Else
            Msg = "sorry, this record cannot be deleted due to data integration"
        End If
        MsgBox Msg, vbMsgBoxRight + vbMsgBoxRtlReading + vbExclamation, App.title
        Rs2.CancelUpdate
    End If
End Sub
'3
'#####################################################################################################################################################
'#####################################################################################################################################################
'#####################################################################################################################################################
Private Sub Cmd3_Click(Index As Integer)

    On Error GoTo ErrTrap

    Select Case Index
        Case 0
            TxtModFlg3.Text = "N"
            clear_all Me
            XPTxtBoxName3.SetFocus
        Case 1
            TxtModFlg3.Text = "E"
            CuurentLogdata3
        Case 2
            SaveData3
        Case 3
            Call Undo3
        Case 4
            Del_Company3
        Case 5
        Case 6
            Unload Me
    End Select
    Exit Sub
ErrTrap:
End Sub
Function CuurentLogdata3(Optional Currentmode As String)

    LogTextA = "    ‘«‘… " & ScreenNameArabic & Chr(13) & " —ﬁ„ «·›—“   " & XPTxtBoxID3.Text & Chr(13) & "  «”„ «·›—“ " & XPTxtBoxName3.Text & Chr(13) & "  „·«ÕŸ«  " & XPMTxtRemark3.Text
    LogTextE = "    Screen  " & ScreenNameEnglish & Chr(13) & " Class No   " & XPTxtBoxID3.Text & Chr(13) & " Class Name" & XPTxtBoxName3.Text & Chr(13) & "  Remarks " & XPMTxtRemark3.Text
    If Currentmode <> "D" Then
        AddToLogFile CInt(user_id), 0, Date, Time, LogTextA, LogTextE, Me.name, Me.TxtModFlg3
    Else
        AddToLogFile CInt(user_id), 0, Date, Time, LogTextA, LogTextE, Me.name, "D"
    End If
End Function

Private Sub TxtModFlg3_Change()

    On Error GoTo ErrTrap

    Select Case Me.TxtModFlg3.Text
        Case "R"
            Me.Cmd3(2).Enabled = False
            Me.Cmd3(3).Enabled = False
            Me.Cmd3(0).Enabled = True
            Me.Cmd3(1).Enabled = True
            Me.Cmd3(4).Enabled = True
            Me.XPBtnMove3(0).Enabled = True
            Me.XPBtnMove3(1).Enabled = True
            Me.XPBtnMove3(2).Enabled = True
            Me.XPBtnMove3(3).Enabled = True
            Me.XPTxtBoxID3.locked = True
            Me.XPTxtBoxName3.locked = True
            Me.XPMTxtRemark3.locked = True
            If Rs3.RecordCount < 1 Then
                Me.XPBtnMove3(0).Enabled = False
                Me.XPBtnMove3(1).Enabled = False
                Me.XPBtnMove3(2).Enabled = False
                Me.XPBtnMove3(3).Enabled = False
                Me.Cmd3(1).Enabled = False
                Me.Cmd3(4).Enabled = False
            End If
        Case "N"
            Me.Cmd3(2).Enabled = True
            Me.Cmd3(3).Enabled = True
            Me.Cmd3(0).Enabled = False
            Me.Cmd3(1).Enabled = False
            Me.Cmd3(4).Enabled = False
            Me.XPTxtBoxID3.locked = True
            Me.XPTxtBoxName3.locked = False
            Me.XPMTxtRemark3.locked = False
        Case "E"
            Me.Cmd3(2).Enabled = True
            Me.Cmd3(3).Enabled = True
            Me.Cmd3(0).Enabled = False
            Me.Cmd3(1).Enabled = False
            Me.Cmd3(4).Enabled = False
            Me.XPBtnMove3(0).Enabled = False
            Me.XPBtnMove3(1).Enabled = False
            Me.XPBtnMove3(2).Enabled = False
            Me.XPBtnMove3(3).Enabled = False
            Me.XPTxtBoxID3.locked = True
            Me.XPTxtBoxName3.locked = False
            Me.XPMTxtRemark3.locked = False
    End Select
    Exit Sub
ErrTrap:
End Sub
Public Sub Retrive3(Optional Lngid As Long = 0)

    On Error GoTo ErrTrap

    If Rs3.RecordCount < 1 Then
        XPTxtCurrent3.Caption = 0
        XPTxtCount3.Caption = 0
        Exit Sub
    End If

    XPTxtBoxID3.Text = IIf(IsNull(Rs3("Locid").value), "", val(Rs3("Locid").value))
    XPTxtBoxName3.Text = IIf(IsNull(Rs3("name").value), "", Trim(Rs3("name").value))
    XPTxtBoxNamee.Text = IIf(IsNull(Rs3("namee").value), "", Trim(Rs3("namee").value))
    XPMTxtRemark3.Text = IIf(IsNull(Rs3("Comment").value), "", Trim(Rs3("Comment").value))
    XPTxtCurrent3.Caption = Rs3.AbsolutePosition
    XPTxtCount3.Caption = Rs3.RecordCount
    Exit Sub
ErrTrap:
End Sub
Private Sub XPBtnMove3_Click(Index As Integer)

    On Error GoTo ErrTrap

    If Me.TxtModFlg3.Text = "N" Then
        clear_all Me
        Me.TxtModFlg3.Text = "R"
        XPBtnMove3_Click (1)
    End If

    Select Case Index
        Case 0
            If Not (Rs3.EOF Or Rs3.BOF) Then
                Rs3.MovePrevious
                If Rs3.BOF Then Rs3.MoveFirst
            End If
        Case 1
            If Not (Rs3.EOF Or Rs3.BOF) Then
                Rs3.MoveFirst
            End If
        Case 2
            If Not (Rs3.EOF Or Rs3.BOF) Then
                Rs3.MoveLast
            End If
        Case 3
            If Not (Rs3.EOF Or Rs3.BOF) Then
                Rs3.MoveNext
                If Rs3.EOF Then Rs3.MoveLast
            End If
    End Select
    Retrive3
    Exit Sub
ErrTrap:
End Sub
Private Sub SaveData3()

    Dim Msg As String
    Dim Strs3QL As String
    Dim rs3Temp As New ADODB.Recordset
    Dim rs3TempM As New ADODB.Recordset
    Dim BeginTrans As Boolean
    
    On Error GoTo ErrTrap

    If Me.TxtModFlg3.Text <> "R" Then
        If XPTxtBoxName3.Text = "" Then
            If SystemOptions.UserInterface = ArabicInterface Then
                MsgBox "„‰ ›÷·ﬂ √œŒ· «”„ «·„ﬂ«‰ ", vbOKOnly + vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
            Else
                MsgBox "Please enter the name", vbOKOnly + vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
            End If
            XPTxtBoxName3.SetFocus
            Exit Sub
        End If
        Select Case Me.TxtModFlg3.Text
            Case "N"
                Strs3QL = "select * from  TblstoresLocations where name ='" & Trim(XPTxtBoxName3.Text) & "'"
                rs3Temp.Open Strs3QL, Cn, adOpenStatic, adLockOptimistic, adCmdText
                If rs3Temp.RecordCount > 0 Then
                    If SystemOptions.UserInterface = ArabicInterface Then
                        Msg = "Â‰«ﬂ „ﬂ«‰   „”Ã· „”»ﬁ« »Â–« «·«”„" & Chr(13)
                        Msg = Msg + "»—Ã«¡ «· √ﬂœ „‰ «·«”„ «·’ÕÌÕ " & Chr(13)
                        Msg = Msg + "√Ê  €ÌÌ— √Ê  „ÌÌ“ «”„ «·„ﬁ«”"
                    Else
                        Msg = "This record already exists"
                    End If
                    MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
                    XPTxtBoxName3.SetFocus
                    Exit Sub
                End If
            Case "E"
                Strs3QL = "select * from  TblstoresLocations where name='" & Trim(XPTxtBoxName3.Text) & "'"
                rs3Temp.Open Strs3QL, Cn, adOpenStatic, adLockOptimistic, adCmdText
                If rs3Temp.RecordCount > 0 Then
                    If rs3Temp("Locid").value <> val(XPTxtBoxID3.Text) Then
                        If SystemOptions.UserInterface = ArabicInterface Then
                            Msg = "Â‰«ﬂ „ﬂ«‰  „”Ã· „”»ﬁ« »Â–« «·«”„" & Chr(13)
                            Msg = Msg + "»—Ã«¡ «· √ﬂœ „‰ «·«”„ «·’ÕÌÕ " & Chr(13)
                            Msg = Msg + "√Ê  €ÌÌ— √Ê  „ÌÌ“ «”„ «·„ﬁ«”"
                        Else
                            Msg = "This record already exists"
                        End If
                        MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
                        XPTxtBoxName3.SetFocus
                        Exit Sub
                    End If
                End If
        End Select
        Cn.BeginTrans
        BeginTrans = True
        Select Case Me.TxtModFlg3.Text
            Case "N"
                Rs3.AddNew
                XPTxtBoxID3.Text = CStr(new_id("TblstoresLocations", "Locid", "", True))
                Rs3("Locid").value = val(XPTxtBoxID3.Text)
            Case "E"
                If Rs3("Locid").value <> val(Me.XPTxtBoxID3.Text) Then
                    Rs3.Find "Locid=" & val(Me.XPTxtBoxID3.Text), , adSearchForward, 1
                    If Rs3.EOF Or Rs3.EOF Then
                        Exit Sub
                    End If
                End If
        End Select
        Rs3("name").value = Trim(XPTxtBoxName3.Text)
        Rs3("namee").value = Trim(XPTxtBoxNamee.Text)
        Rs3("Comment").value = IIf(XPMTxtRemark3.Text = "", Null, Trim(XPMTxtRemark3.Text))
        Rs3.update
        Cn.CommitTrans
        BeginTrans = False
        XPTxtCurrent3.Caption = Rs3.AbsolutePosition
        XPTxtCount3.Caption = Rs3.RecordCount
        CuurentLogdata3
        Select Case Me.TxtModFlg3.Text
            Case "N"
                If SystemOptions.UserInterface = ArabicInterface Then
                    Msg = "  „ Õ›Ÿ »Ì«‰«  Â–« «·„ﬁ«”" & Chr(13)
                    Msg = Msg + "Â·  —€» ›Ì ≈÷«›… »Ì«‰«  √Œ—Ì"
                Else
                    Msg = "Recored saved successfully , do you want to add another recored"
                End If
                If MsgBox(Msg, vbYesNo + vbQuestion + vbMsgBoxRight + vbMsgBoxRtlReading + vbDefaultButton2, App.title) = vbYes Then
                    Cmd3_Click (0)
                    Exit Sub
                End If
            Case "E"
                If SystemOptions.UserInterface = ArabicInterface Then
                    MsgBox " „ Õ›Ÿ Â–Â «· ⁄œÌ·« ", vbInformation + vbOKOnly + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
                Else
                    MsgBox "Record saved successfully", vbInformation + vbOKOnly + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
                End If
        End Select
        TxtModFlg3.Text = "R"
    End If
    Exit Sub
ErrTrap:
    If BeginTrans = True Then
        BeginTrans = False
        Cn.RollbackTrans
    End If
    If Err.Number = -2147217900 Then
        If SystemOptions.UserInterface = ArabicInterface Then
            Msg = "·« Ì„ﬂ‰ Õ›Ÿ Â–Â «·»Ì«‰«  " & Chr(13)
            Msg = Msg + "·ﬁœ  „ «œŒ«· ﬁÌ„ €Ì— ’«·Õ… " & Chr(13)
            Msg = Msg + " √ﬂœ „‰ œﬁ… «·»Ì«‰«  Ê√⁄œ «·„Õ«Ê·…"
        Else
            Msg = "Data Can't be saved " & Chr(13)
            Msg = Msg & "Invalid data values was entered" & Chr(13)
            Msg = Msg & "Please make sure of the entered data and try again"
        End If
        MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
        Exit Sub
    End If
    If SystemOptions.UserInterface = ArabicInterface Then
        Msg = "⁄›Ê«...ÕœÀ Œÿ√ „« √À‰«¡ Õ›Ÿ Â–Â «·»Ì«‰«  " & Chr(13)
    Else
        Msg = "Sorry , somthing went wrong while saving data" & Chr(13)
    End If
    MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
End Sub
Private Sub Undo3()

    On Error GoTo ErrTrap

    Select Case TxtModFlg3.Text
        Case "N"
            clear_all Me
            Me.TxtModFlg3.Text = "R"
            XPBtnMove3_Click (1)
        Case "E"
            Rs3.Find "BoxID='" & val(XPTxtBoxID3.Text) & "'", , adSearchForward, adBookmarkFirst
            If Rs3.EOF Or Rs3.BOF Then
                Me.TxtModFlg3.Text = "R"
                Exit Sub
            End If
            Retrive3
            Me.TxtModFlg3.Text = "R"
    End Select
    Exit Sub
ErrTrap:
End Sub
Private Sub Del_Company3()

    Dim Msg As String
    Dim Strs3QL As String
    Dim rs3Temp As New ADODB.Recordset
    
    On Error GoTo ErrTrap
    
    If XPTxtBoxID3.Text <> "" Then
        'If val(Me.XPTxtBoxID3.Text) = 1 Then
            'If SystemOptions.UserInterface = ArabicInterface Then
                'Msg = "·«Ì„ﬂ‰ Õ–› Â–« «·”Ã·...!!!"
            'Else
                'Msg = "sorry, this record cannot be deleted"
            'End If
            'MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
            'Exit Sub
        'End If
        
        'Strs3QL = "select * from Transaction_Details where Locid=" & Trim(XPTxtBoxID3.Text)
        'rs3Temp.Open Strs3QL, Cn, adOpenStatic, adLockOptimistic, adCmdText
        'If Not (rs3Temp.EOF Or rs3Temp.BOF) Then
        '    Msg = "·« Ì„ﬂ‰ Õ–› »Ì«‰«  Â–« «··Ê‰" & Chr(13)
        '    Msg = Msg + "Â‰«ﬂ »⁄÷ «·⁄„·Ì«  „— »ÿ… »Â–« «··Ê‰"
        '    MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
        '    Exit Sub
        'End If
        If SystemOptions.UserInterface = ArabicInterface Then
            Msg = "”Ì „ Õ–› »Ì«‰«  «··Ê‰ —ﬁ„ " & Chr(13)
            Msg = Msg + (XPTxtBoxID3.Text) & Chr(13)
            Msg = Msg + " Â·  —€» ›Ì Õ–› Â–Â «·»Ì«‰« ø"
        Else
            Msg = "Are you sure you want to delete this record"
        End If
        If MsgBox(Msg, vbYesNo + vbQuestion + vbMsgBoxRight + vbMsgBoxRtlReading, App.title) = vbYes Then
            If Not Rs3.RecordCount < 1 Then
                CuurentLogdata3 ("D")
                Rs3.Delete
                Rs3.MoveFirst

                If Rs3.RecordCount < 1 Then
                    clear_all Me
                    TxtModFlg3_Change
                    XPTxtCurrent3.Caption = 0
                    XPTxtCount3.Caption = 0
                Else
                    Retrive3
                End If
            End If
        End If
    Else
        clear_all Me
        If SystemOptions.UserInterface = ArabicInterface Then
            Msg = "Â–Â «·⁄„·Ì… €Ì— „ «Õ… ÕÌÀ √‰Â ·«ÌÊÃœ √Ì ”Ã·« "
        Else
            Msg = "This operation is not available due to lack of records"
        End If
        MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
        TxtModFlg3_Change
        Exit Sub
    End If
    TxtModFlg3_Change
    Exit Sub
ErrTrap:
    If Err.Number = -2147217887 Then
        If SystemOptions.UserInterface = ArabicInterface Then
            Msg = "·«Ì„ﬂ‰ Õ–› Â–« «·”Ã· · ﬂ«„· «·»Ì«‰«  " & Chr(13) & " ÊÃœ »Ì«‰«  „— »ÿ… »Â–« «··Ê‰ "
            Msg = Msg & Chr(13) & Err.description
        Else
            Msg = "sorry, this record cannot be deleted due to data integration"
        End If
        MsgBox Msg, vbMsgBoxRight + vbMsgBoxRtlReading + vbExclamation, App.title
        Rs3.CancelUpdate
    End If
End Sub
'4
'######################################################################################################################################################
'######################################################################################################################################################
'######################################################################################################################################################

Private Sub btnCancel4_Click()
    Unload Me
End Sub
Private Sub btnDelete4_Click()

    Dim MSGType As Integer
    Dim BlnRecExist As Boolean
    Dim StrMSG  As String
    Dim Msg As String
    
    On Error GoTo ErrTrap

    If DoPremis(Do_Delete, Me.name, True) = False Then
        Exit Sub
    End If
    If SystemOptions.UserInterface = ArabicInterface Then
        MSGType = MsgBox("Â·  —€» ›Ì Õ–› Â–« «·”Ã·", vbQuestion + vbMsgBoxRtlReading + vbYesNo + vbMsgBoxRight, App.title)
    Else
        MSGType = MsgBox("Do you want to delete this record", vbQuestion + vbMsgBoxRtlReading + vbYesNo + vbMsgBoxRight, App.title)
    End If
    If MSGType = vbYes Then
        RsSavRec4.Find "id=" & val(TxtVac_ID.Text), , adSearchForward, 1
        RsSavRec4.Delete
        If SystemOptions.UserInterface = ArabicInterface Then
            MsgBox " „  ⁄„·Ì… «·Õ–› »‰Ã«Õ", vbInformation + vbMsgBoxRtlReading + vbOKOnly + vbMsgBoxRight, App.title
        Else
            MsgBox "Record deleted successfully", vbInformation + vbOKOnly + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
        End If
        FillGrid4WithData
        btnNext4_Click
    End If
    Exit Sub
ErrTrap:
    Select Case Err.Number
        Case -2147217873, -2147467259
            If SystemOptions.UserInterface = ArabicInterface Then
                StrMSG = "⁄›Ê« ·« ÌÃÊ“ Õ–› «·”Ã· ·«— »«ÿÂ »»Ì«‰«  √Œ—Ì"
            Else
                StrMSG = "sorry, this record cannot be deleted due to data integration"
            End If
            RsSavRec4.CancelUpdate
            MsgBox StrMSG, vbOKOnly + vbMsgBoxRight, App.title
            Cn.Errors.Clear
    End Select
End Sub
Private Sub btnFirst4_Click()

    On Error GoTo ErrTrap

    Dim Msg As String

    If Me.TxtModFlg4.Text = "N" Then
        FindRec4 val(TxtVac_ID.Text)
        Me.TxtModFlg4.Text = "R"
    End If

    TxtModFlg4 = "R"

    If RsSavRec4.RecordCount = 0 Then
        clear_all Me
        Exit Sub
    End If
BegnieWork:
    RsSavRec4.MoveFirst
    FiLLTXT4
    Exit Sub
ErrTrap:
    Select Case Err.Number
        Case -2147217885
            If SystemOptions.UserInterface = ArabicInterface Then
                Msg = "⁄›Ê« ·ﬁœ  „ Õ–› Â–« «·”Ã· «· «·Ï" & Chr(13)
                Msg = Msg & "„‰ ﬁ»· „” Œœ„ √Œ— ⁄·Ï «·‘»ﬂ… " & Chr(13)
                Msg = Msg & "”Ê› Ì „  ÕœÌÀ «·»Ì«‰« "
            Else
                Msg = "Sorry , this Recored was deleted by other user on the network" & Chr(13)
                Msg = Msg & "Date will be updated now" & Chr(13)
            End If
            MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading + vbOKOnly, App.title
            RsSavRec4.Requery
            Resume BegnieWork
    End Select
End Sub
Private Sub btnLast4_Click()

    On Error GoTo ErrTrap

    Dim Msg As String
    
    If Me.TxtModFlg4.Text = "N" Then
        FindRec4 val(TxtVac_ID.Text)
        Me.TxtModFlg4.Text = "R"
    End If
    TxtModFlg4 = "R"
    If RsSavRec4.RecordCount = 0 Then
        clear_all Me
        Exit Sub
    End If
BegnieWork:
    RsSavRec4.MoveLast
    FiLLTXT4
    Exit Sub
ErrTrap:
    Select Case Err.Number
        Case -2147217885
            If SystemOptions.UserInterface = ArabicInterface Then
                Msg = "⁄›Ê« ·ﬁœ  „ Õ–› Â–« «·”Ã· «· «·Ï" & Chr(13)
                Msg = Msg & "„‰ ﬁ»· „” Œœ„ √Œ— ⁄·Ï «·‘»ﬂ… " & Chr(13)
                Msg = Msg & "”Ê› Ì „  ÕœÌÀ «·»Ì«‰« "
            Else
                Msg = "Sorry , this Recored was deleted by other user on the network" & Chr(13)
                Msg = Msg & "Date will be updated now" & Chr(13)
            End If
            MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading + vbOKOnly, App.title
            RsSavRec4.Requery
            Resume BegnieWork
    End Select

End Sub
Private Sub btnModify4_Click()

    Dim Msg As String

    If DoPremis(Do_Edit, Me.name, True) = False Then
        Exit Sub
    End If

    On Error GoTo ErrTrap

    If TxtVac_ID.Text <> "" Then
        TxtModFlg4 = "E"
        Frm24.Enabled = True
        Me.TxtVacName4.SetFocus
    End If

    Exit Sub
ErrTrap:

    Select Case Err.Number
        Case -2147467259
            If SystemOptions.UserInterface = ArabicInterface Then
                Msg = "⁄›Ê«" & Chr(13)
                Msg = Msg & " ·«Ì„ﬂ‰  ⁄œÌ· Â–« «·”Ã· ›Ï «·Êﬁ  «·Õ«·Ï" & Chr(13)
                Msg = Msg & "ÕÌÀ «‰Â ﬁÌœ «· ⁄œÌ· „‰ ﬁ»· „” Œœ„ «Œ— ⁄·Ï «·‘»ﬂ…"
            Else
                Msg = "Sorry , this Recored was deleted by other user on the network" & Chr(13)
                Msg = Msg & "Date will be updated now" & Chr(13)
            End If
            MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading + vbOKOnly, App.title
        
            If RsSavRec4.EditMode <> adEditNone Then
                RsSavRec4.CancelUpdate
            End If
    End Select
End Sub
Private Sub btnNew4_Click()

    Dim My_SQL As String
    Dim RS As ADODB.Recordset

    If DoPremis(Do_New, Me.name, True) = False Then
        Exit Sub
    End If

    On Error GoTo ErrTrap
    Set RS = New ADODB.Recordset
    Frm24.Enabled = True
    clear_all Me
    TxtModFlg4.Text = "N"

    My_SQL = "TblSpecification"
    RS.Open My_SQL, Cn, adOpenKeyset, adLockOptimistic, adCmdTable

    If RS.RecordCount > 0 Then
        TxtSerial.Text = RS.RecordCount + 1
    Else
        TxtSerial.Text = 1
    End If

    RS.Close
   
    TxtVacName4.SetFocus
ErrTrap:
End Sub
Private Sub btnNext4_Click()

    On Error GoTo ErrTrap
    
    Dim Msg As String

    If Me.TxtModFlg4.Text = "N" Then
        FindRec4 val(TxtVac_ID.Text)
        Me.TxtModFlg4.Text = "R"
    End If
    TxtModFlg4 = "R"
    If RsSavRec4.RecordCount = 0 Then
        clear_all Me
        Exit Sub
    End If
BegnieWork:
    If RsSavRec4.EOF Then
        RsSavRec4.MoveLast
    Else
        RsSavRec4.MoveNext
        If RsSavRec4.EOF Then
            RsSavRec4.MoveLast
        End If
    End If
    FiLLTXT4
    Exit Sub
ErrTrap:
    Select Case Err.Number
        Case -2147217885
            If SystemOptions.UserInterface = ArabicInterface Then
                Msg = "⁄›Ê« ·ﬁœ  „ Õ–› Â–« «·”Ã· «· «·Ï" & Chr(13)
                Msg = Msg & "„‰ ﬁ»· „” Œœ„ √Œ— ⁄·Ï «·‘»ﬂ… " & Chr(13)
                Msg = Msg & "”Ê› Ì „  ÕœÌÀ «·»Ì«‰« "
            Else
                Msg = "Sorry , this Recored was deleted by other user on the network" & Chr(13)
                Msg = Msg & "Date will be updated now" & Chr(13)
            End If
            MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading + vbOKOnly, App.title
            RsSavRec4.Requery
            Resume BegnieWork
    End Select
End Sub
Private Sub btnPrevious4_Click()

    On Error GoTo ErrTrap
    
    Dim Msg As String
    If Me.TxtModFlg4.Text = "N" Then
        FindRec4 val(TxtVac_ID.Text)
        Me.TxtModFlg4.Text = "R"
    End If
    TxtModFlg4 = "R"
    If RsSavRec4.RecordCount = 0 Then
        clear_all Me
        Exit Sub
    End If

BegnieWork:
    RsSavRec4.MovePrevious
    If RsSavRec4.BOF Then
        RsSavRec4.MoveFirst
    End If
    FiLLTXT4
    Exit Sub
ErrTrap:
    Select Case Err.Number
        Case -2147217885
            If SystemOptions.UserInterface = ArabicInterface Then
                Msg = "⁄›Ê« ·ﬁœ  „ Õ–› Â–« «·”Ã· «· «·Ï" & Chr(13)
                Msg = Msg & "„‰ ﬁ»· „” Œœ„ √Œ— ⁄·Ï «·‘»ﬂ… " & Chr(13)
                Msg = Msg & "”Ê› Ì „  ÕœÌÀ «·»Ì«‰« "
            Else
                Msg = "Sorry , this Recored was deleted by other user on the network" & Chr(13)
                Msg = Msg & "Date will be updated now" & Chr(13)
            End If
            MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading + vbOKOnly, App.title
            RsSavRec4.Requery
            Resume BegnieWork
    End Select
End Sub
Private Sub btnSave4_Click()

    On Error GoTo ErrTrap
    
    Dim Msg As String
    Dim StrVacCode As String
    Dim StrVacName As String
    Dim CtrlTxt As Control
    For Each CtrlTxt In Me.Controls

        If TypeOf CtrlTxt Is TextBox Or TypeOf CtrlTxt Is ComboBox Then
            If CtrlTxt.Text = "" And CtrlTxt.Tag <> "" And CtrlTxt.Enabled = True Then
                MsgBox CtrlTxt.Tag, vbOKOnly + vbMsgBoxRight, App.title
                CtrlTxt.SetFocus
                Exit Sub
            End If
        End If
    Next
    StrVacName = IsRecExist("TblSpecification", "nameˆ", Trim(TxtVacName4.Text), "name", "ID<>'" & Trim(TxtVac_ID.Text) & "'")

    If StrVacName <> "" Then
        If SystemOptions.UserInterface = ArabicInterface Then
            Msg = "·ﬁœ ”»ﬁ  ”ÃÌ· Â–« «·‰Ê⁄ „‰ ﬁ»·"
        Else
            Msg = "This record already exists"
        End If
        MsgBox Msg, vbOKOnly + vbMsgBoxRight, App.title
        TxtVacName4.SetFocus
    
        Exit Sub

    End If
    Select Case Me.TxtModFlg4.Text
        Case "N"
            AddNewRec4
            btnLast4_Click
        Case "E"
            FiLLRec4
    End Select
    Exit Sub
ErrTrap:
    If SystemOptions.UserInterface = ArabicInterface Then
        MsgBox "Â‰«ﬂ Œÿ√ „« ›Ì ≈œŒ«· «·»Ì«‰« ", vbOKOnly + vbMsgBoxRight, App.title
    Else
        MsgBox "Something went wrong while inserting data", vbOKOnly + vbMsgBoxRight, App.title
    End If
End Sub
 
Private Sub BtnUndo4_Click()
    FindRec4 val(TxtVac_ID.Text)
    Me.TxtModFlg4.Text = "R"
End Sub

Private Sub BtnUpdate4_Click()

    On Error GoTo ErrTrap
    
    Dim Msg As String
    Dim FristCount As Long
    Dim LastCount As Long
    
    FristCount = RsSavRec4.RecordCount
    RsSavRec4.Requery
    LastCount = RsSavRec4.RecordCount
    BtnUndo4_Click
    If SystemOptions.UserInterface = ArabicInterface Then
        If FristCount = LastCount Then
            Msg = "·«  ÊÃœ »Ì«‰«  ÃœÌœ…"
        Else
            Msg = "⁄œœ «·”Ã·«  ﬁ»· «· ÕœÌÀ" & vbCrLf & FristCount & vbCrLf & "⁄œœ «·”Ã·«  »⁄œ «· ÕœÌÀ" & vbCrLf & LastCount
            If LastCount > FristCount Then
                Msg = Msg + vbCrLf & "⁄œœ «·”Ã·«  «·ÃœÌœ…" & vbCrLf & LastCount - FristCount
            Else
                Msg = Msg + vbCrLf & "⁄œœ «·”Ã·«  «·„Õ–Ê›…" & vbCrLf & FristCount - LastCount
            End If
        End If
    Else
        If FristCount = LastCount Then
            Msg = "No new data"
        Else
            Msg = "No. of records before the update" & vbCrLf & FristCount & vbCrLf & "No. of records after update" & vbCrLf & LastCount
        
            If LastCount > FristCount Then
                Msg = Msg + vbCrLf & "No. of new records" & vbCrLf & LastCount - FristCount
            Else
                Msg = Msg + vbCrLf & "No. of deleted records" & vbCrLf & FristCount - LastCount
            End If
        End If
    End If
    MsgBox Msg, vbOKOnly + vbMsgBoxRight + vbInformation, App.title
ErrTrap:
End Sub

Public Sub AddNewRec4()

    On Error GoTo ErrTrap
    
    Dim StrRecId4 As String
    
    StrRecId4 = new_id("TblSpecification", "id", "")
    RsSavRec4.AddNew
    RsSavRec4.Fields("id").value = IIf(StrRecId4 <> "", StrRecId4, Null)
    FiLLRec4
ErrTrap:
End Sub
Public Sub FiLLRec4()

    On Error GoTo ErrTrap

    RsSavRec4.Fields("name").value = IIf(TxtVacName4.Text <> "", Trim(TxtVacName4.Text), Null)
    RsSavRec4.Fields("namee").value = IIf(TxtVacNamee4.Text <> "", Trim(TxtVacNamee4.Text), Null)
    RsSavRec4.update
    If SystemOptions.UserInterface = ArabicInterface Then
        MsgBox " „  ⁄„·Ì… «·Õ›Ÿ »‰Ã«Õ", vbInformation + vbOKOnly + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
    Else
        MsgBox "Record saved successfully", vbInformation + vbOKOnly + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
    End If
    FillGrid4WithData
    TxtModFlg4 = "R"
    Exit Sub
ErrTrap:
    If RsSavRec4.EditMode <> adEditNone Then
        RsSavRec4.CancelUpdate
    End If
End Sub
Public Sub FiLLTXT4()

    On Error GoTo ErrTrap
    
    Dim i As Integer
    
    Frm24.Enabled = False
    TxtVac_ID.Text = IIf(IsNull(RsSavRec4.Fields("id").value), "", RsSavRec4.Fields("id").value)
    TxtVacName4.Text = IIf(IsNull(RsSavRec4.Fields("name").value), "", RsSavRec4.Fields("name").value)
    TxtVacNamee4.Text = IIf(IsNull(RsSavRec4.Fields("namee").value), "", RsSavRec4.Fields("namee").value)
    LabCurrRec4.Caption = RsSavRec4.AbsolutePosition
    LabCountRec4.Caption = RsSavRec4.RecordCount
    With Grid4
        For i = 1 To .Rows - 1
            If Trim(TxtVac_ID.Text) = .TextMatrix(i, .ColIndex("id")) Then
                TxtSerial.Text = .TextMatrix(i, .ColIndex("Ser"))
                .Row = i
                Exit Sub
            End If
        Next
    End With
ErrTrap:
End Sub

Public Sub EditRec4(StrTable As String, RecId4 As String)
    FiLLRec4
End Sub
Private Sub Grid4_EnterCell()

    On Error GoTo ErrTrap
    
    FindRec4 val(Me.Grid4.TextMatrix(Me.Grid4.Row, Me.Grid4.ColIndex("id")))
ErrTrap:
End Sub
Private Sub TxtVac_ID_Change()

    Dim TxtMod As String
    
    TxtMod = TxtModFlg4.Text
    TxtModFlg4.Text = ""
    TxtModFlg4 = TxtMod
End Sub
Public Function FindRec4(ByVal RecId4 As Long)

    On Error GoTo ErrTrap
    
    RsSavRec4.Find "id=" & RecId4, , adSearchForward, 1
    If Not (RsSavRec4.EOF) Then
        FiLLTXT4
    End If
    Exit Function
ErrTrap:
    If RsSavRec4.EditMode <> adEditNone Then
        RsSavRec4.CancelUpdate
        BtnUndo4_Click
    End If
End Function
Private Sub TxtModFlg4_Change()
    If TxtModFlg4.Text = "N" Then
        Frm24.Enabled = True
        Me.btnNew4.Enabled = False
        btnModify4.Enabled = False
        btnDelete4.Enabled = False
        Me.btnQuery4.Enabled = False
        Grid4.Enabled = False
        BtnUndo4.Enabled = True
        Me.btnSave4.Enabled = True
        BtnUpdate4.Enabled = False
    ElseIf TxtModFlg4.Text = "R" Then
        Frm24.Enabled = False
        Grid4.Enabled = True
        btnModify4.Enabled = False
        btnDelete4.Enabled = False
        If TxtVac_ID.Text <> "" Then
            btnModify4.Enabled = True
            btnDelete4.Enabled = True
        End If
        BtnUpdate4.Enabled = True
        Me.btnQuery4.Enabled = True
        Me.btnNew4.Enabled = True
        BtnUndo4.Enabled = False
        Me.btnSave4.Enabled = False
        btnNext4.Enabled = True
        btnPrevious4.Enabled = True
        btnFirst4.Enabled = True
        btnLast4.Enabled = True
    ElseIf TxtModFlg4.Text = "E" Then
        Frm24.Enabled = True
        Me.btnNew4.Enabled = False
        btnModify4.Enabled = False
        btnDelete4.Enabled = False
        Me.btnQuery4.Enabled = False
        BtnUpdate4.Enabled = False
        BtnUndo4.Enabled = True
        Me.btnSave4.Enabled = True
        Grid4.Enabled = False
        btnNext4.Enabled = False
        btnPrevious4.Enabled = False
        btnFirst4.Enabled = False
        btnLast4.Enabled = False
    End If
End Sub
Public Sub FillGrid4WithData()

    On Error GoTo ErrTrap

    Dim i As Integer
    Dim RS As ADODB.Recordset
    Dim My_SQL As String

    Set RS = New ADODB.Recordset
    My_SQL = "select * From TblSpecification order by id"
    RS.Open My_SQL, Cn, adOpenKeyset, adLockReadOnly, adCmdText

    With Me.Grid4
        .Rows = 2
        .Clear flexClearScrollable
        If RS.RecordCount > 0 Then
            .Rows = RS.RecordCount + 1
            RS.MoveFirst
            For i = 1 To .Rows - 1
                .TextMatrix(i, .ColIndex("Ser")) = i
                .TextMatrix(i, .ColIndex("name")) = IIf(IsNull(RS.Fields("name").value), "", RS.Fields("name").value)
                .TextMatrix(i, .ColIndex("namee")) = IIf(IsNull(RS.Fields("namee").value), "", RS.Fields("namee").value)
                .TextMatrix(i, .ColIndex("id")) = IIf(IsNull(RS.Fields("id").value), "", RS.Fields("id").value)
                RS.MoveNext
            Next
            RS.Close
        End If
        .RowHeight(-1) = 300
    End With
ErrTrap:
End Sub
Private Function CheckDelCountry(Lngid As Long) As Boolean

    Dim RS As ADODB.Recordset
    Dim StrSQL As String
    
    StrSQL = "Select * From TblEmployee Where id=" & Lngid & ""
    Set RS = New ADODB.Recordset
    RS.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText
    If Not (RS.BOF Or RS.EOF) Then
        CheckDelCountry = False
    Else
        CheckDelCountry = True
    End If
    RS.Close
    Set RS = Nothing
End Function
Private Sub TxtVacName4_GotFocus()
    SwitchKeyboardLang LANG_ARABIC
End Sub
Private Sub TxtVacNamee4_GotFocus()
    SwitchKeyboardLang LANG_ENGLISH
End Sub
'5
'######################################################################################################################################################
'######################################################################################################################################################
'######################################################################################################################################################

 

Private Sub btnCancel5_Click()
    Unload Me
End Sub

Private Sub btnDelete5_Click()

    Dim MSGType As Integer
    Dim BlnRecExist As Boolean
    Dim StrMSG  As String
    Dim Msg As String
    
    On Error GoTo ErrTrap
    
    If DoPremis(Do_Delete, Me.name, True) = False Then
        Exit Sub
    End If
    If TxtVac_ID5.Text <> "" Then
        If SystemOptions.UserInterface = ArabicInterface Then
            MSGType = MsgBox("Â·  —€» ›Ì Õ–› Â–« «·”Ã·", vbQuestion + vbMsgBoxRtlReading + vbYesNo + vbMsgBoxRight, App.title)
        Else
            MSGType = MsgBox("Do you want to delete this record", vbQuestion + vbMsgBoxRtlReading + vbYesNo + vbMsgBoxRight, App.title)
        End If
        If MSGType = vbYes Then
            RsSavRec5.Find "TBLProductionElementsId=" & val(TxtVac_ID5.Text), , adSearchForward, 1
            CuurentLogdata5 ("D")
            RsSavRec5.Delete
            If SystemOptions.UserInterface = ArabicInterface Then
                MsgBox " „  ⁄„·Ì… «·Õ–› »‰Ã«Õ", vbInformation + vbMsgBoxRtlReading + vbOKOnly + vbMsgBoxRight, App.title
            Else
                MsgBox "Record deleted successfully", vbInformation + vbOKOnly + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
            End If
            FillGrid5WithData
            btnNext5_Click
        End If
    End If
    Exit Sub
ErrTrap:
    Select Case Err.Number
        Case -2147217873, -2147467259
            If SystemOptions.UserInterface = ArabicInterface Then
                StrMSG = "⁄›Ê« ·« ÌÃÊ“ Õ–› «·”Ã· ·«— »«ÿÂ »»Ì«‰«  √Œ—Ì"
            Else
                StrMSG = "sorry, this record cannot be deleted due to data integration"
            End If
            RsSavRec5.CancelUpdate
            MsgBox StrMSG, vbOKOnly + vbMsgBoxRight, App.title
            Cn.Errors.Clear
    End Select
End Sub
Private Sub btnFirst5_Click()

    On Error GoTo ErrTrap

    Dim Msg As String

    If Me.TxtModFlg5.Text = "N" Then
        FindRec5 val(TxtVac_ID5.Text)
        Me.TxtModFlg5.Text = "R"
    End If
    TxtModFlg5 = "R"
    If RsSavRec5.RecordCount = 0 Then
        clear_all Me
        Exit Sub
    End If
BegnieWork:
    RsSavRec5.MoveFirst
    FiLLTXT5
    Exit Sub
ErrTrap:
    Select Case Err.Number
        Case -2147217885
            If SystemOptions.UserInterface = ArabicInterface Then
                Msg = "⁄›Ê« ·ﬁœ  „ Õ–› Â–« «·”Ã· «· «·Ï" & Chr(13)
                Msg = Msg & "„‰ ﬁ»· „” Œœ„ √Œ— ⁄·Ï «·‘»ﬂ… " & Chr(13)
                Msg = Msg & "”Ê› Ì „  ÕœÌÀ «·»Ì«‰« "
            Else
                Msg = "Sorry , this Recored was deleted by other user on the network" & Chr(13)
                Msg = Msg & "Date will be updated now" & Chr(13)
            End If
            MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading + vbOKOnly, App.title
            RsSavRec5.Requery
            Resume BegnieWork
    End Select
End Sub
Private Sub btnLast5_Click()

    On Error GoTo ErrTrap

    Dim Msg As String

    If Me.TxtModFlg5.Text = "N" Then
        FindRec5 val(TxtVac_ID5.Text)
        Me.TxtModFlg5.Text = "R"
    End If
    TxtModFlg5 = "R"
    If RsSavRec5.RecordCount = 0 Then
        clear_all Me
        Exit Sub
    End If
BegnieWork:
    RsSavRec5.MoveLast
    FiLLTXT5
    Exit Sub
ErrTrap:
    Select Case Err.Number
        Case -2147217885
            If SystemOptions.UserInterface = ArabicInterface Then
                Msg = "⁄›Ê« ·ﬁœ  „ Õ–› Â–« «·”Ã· «· «·Ï" & Chr(13)
                Msg = Msg & "„‰ ﬁ»· „” Œœ„ √Œ— ⁄·Ï «·‘»ﬂ… " & Chr(13)
                Msg = Msg & "”Ê› Ì „  ÕœÌÀ «·»Ì«‰« "
            Else
                Msg = "Sorry , this Recored was deleted by other user on the network" & Chr(13)
                Msg = Msg & "Date will be updated now" & Chr(13)
            End If
            MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading + vbOKOnly, App.title
            RsSavRec5.Requery
            Resume BegnieWork
    End Select
End Sub
Private Sub btnModify5_Click()

    Dim Msg As String

    If DoPremis(Do_Edit, Me.name, True) = False Then
        Exit Sub
    End If

    On Error GoTo ErrTrap

    If TxtVac_ID5.Text <> "" Then
        TxtModFlg5 = "E"
        Frm25.Enabled = True
        Me.TxtVacName5.SetFocus
        CuurentLogdata5
    End If
    Exit Sub
ErrTrap:
    Select Case Err.Number
        Case -2147467259
            If SystemOptions.UserInterface = ArabicInterface Then
                Msg = "⁄›Ê«" & Chr(13)
                Msg = Msg & " ·«Ì„ﬂ‰  ⁄œÌ· Â–« «·”Ã· ›Ï «·Êﬁ  «·Õ«·Ï" & Chr(13)
                Msg = Msg & "ÕÌÀ «‰Â ﬁÌœ «· ⁄œÌ· „‰ ﬁ»· „” Œœ„ «Œ— ⁄·Ï «·‘»ﬂ…"
            Else
                Msg = "Sorry , this Recored was deleted by other user on the network" & Chr(13)
                Msg = Msg & "Date will be updated now" & Chr(13)
            End If
            MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading + vbOKOnly, App.title
            If RsSavRec5.EditMode <> adEditNone Then
                RsSavRec5.CancelUpdate
            End If
    End Select
End Sub
Private Sub btnNew5_Click()

    Dim My_SQL As String
    Dim RS As ADODB.Recordset

    If DoPremis(Do_New, Me.name, True) = False Then
        Exit Sub
    End If

    On Error GoTo ErrTrap
    
    Set RS = New ADODB.Recordset
    Frm25.Enabled = True
    Me.TxtVac_ID5.Text = ""
    Me.TxtVacName5.Text = ""
    Me.TxtVacNamee5.Text = ""
    Me.DcboExpensesID.BoundText = ""
    TxtModFlg5.Text = "N"

    My_SQL = "TBLProductionElements"
    RS.Open My_SQL, Cn, adOpenKeyset, adLockOptimistic, adCmdTable

    If RS.RecordCount > 0 Then
        TxtSerial5.Text = RS.RecordCount + 1
    Else
        TxtSerial5.Text = 1
    End If
    RS.Close
    CmbType.ListIndex = 0
    TxtVacName5.SetFocus
ErrTrap:
End Sub
Private Sub btnNext5_Click()

    On Error GoTo ErrTrap
    
    Dim Msg As String

    If Me.TxtModFlg5.Text = "N" Then
        FindRec5 val(TxtVac_ID5.Text)
        Me.TxtModFlg5.Text = "R"
    End If
    TxtModFlg5 = "R"
    If RsSavRec5.RecordCount = 0 Then
        clear_all Me
        Exit Sub
    End If
BegnieWork:
    If RsSavRec5.EOF Then
        RsSavRec5.MoveLast
    Else
        RsSavRec5.MoveNext
        If RsSavRec5.EOF Then
            RsSavRec5.MoveLast
        End If
    End If
    FiLLTXT5
    Exit Sub
ErrTrap:
    Select Case Err.Number
        Case -2147217885
            If SystemOptions.UserInterface = ArabicInterface Then
                Msg = "⁄›Ê« ·ﬁœ  „ Õ–› Â–« «·”Ã· «· «·Ï" & Chr(13)
                Msg = Msg & "„‰ ﬁ»· „” Œœ„ √Œ— ⁄·Ï «·‘»ﬂ… " & Chr(13)
                Msg = Msg & "”Ê› Ì „  ÕœÌÀ «·»Ì«‰« "
            Else
                Msg = "Sorry , this Recored was deleted by other user on the network" & Chr(13)
                Msg = Msg & "Date will be updated now" & Chr(13)
            End If
            MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading + vbOKOnly, App.title
            RsSavRec5.Requery
            Resume BegnieWork
    End Select
End Sub
Private Sub btnPrevious5_Click()

    On Error GoTo ErrTrap
    
    Dim Msg As String
    
    If Me.TxtModFlg5.Text = "N" Then
        FindRec5 val(TxtVac_ID5.Text)
        Me.TxtModFlg5.Text = "R"
    End If
    TxtModFlg5 = "R"
    If RsSavRec5.RecordCount = 0 Then
        clear_all Me
        Exit Sub
    End If
BegnieWork:
    RsSavRec5.MovePrevious
    If RsSavRec5.BOF Then
        RsSavRec5.MoveFirst
    End If
    FiLLTXT5
    Exit Sub
ErrTrap:
    Select Case Err.Number
        Case -2147217885
            If SystemOptions.UserInterface = ArabicInterface Then
                Msg = "⁄›Ê« ·ﬁœ  „ Õ–› Â–« «·”Ã· «· «·Ï" & Chr(13)
                Msg = Msg & "„‰ ﬁ»· „” Œœ„ √Œ— ⁄·Ï «·‘»ﬂ… " & Chr(13)
                Msg = Msg & "”Ê› Ì „  ÕœÌÀ «·»Ì«‰« "
            Else
                Msg = "Sorry , this Recored was deleted by other user on the network" & Chr(13)
                Msg = Msg & "Date will be updated now" & Chr(13)
            End If
            MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading + vbOKOnly, App.title
            RsSavRec5.Requery
            Resume BegnieWork
    End Select
End Sub

Private Sub btnSave5_Click()

    On Error GoTo ErrTrap
    
    Dim Msg As String
    Dim StrVacCode As String
    Dim StrVacName As String
    Dim CtrlTxt As Control

    For Each CtrlTxt In Me.Controls
        If TypeOf CtrlTxt Is TextBox Or TypeOf CtrlTxt Is ComboBox Then
            If CtrlTxt.Text = "" And CtrlTxt.Tag <> "" And CtrlTxt.Enabled = True Then
                MsgBox CtrlTxt.Tag, vbOKOnly + vbMsgBoxRight, App.title
                CtrlTxt.SetFocus
                Exit Sub
            End If
        End If
    Next
    StrVacName = IsRecExist("TBLProductionElements", "Name", Trim(TxtVacName5.Text), "Name", "Vac_ID<>'" & Trim(TxtVac_ID5.Text) & "'")
    If StrVacName <> "" Then
        If SystemOptions.UserInterface = ArabicInterface Then
            Msg = "·ﬁœ ”»ﬁ  ”ÃÌ· Â–« «·‰Ê⁄ „‰ ﬁ»·"
        Else
            Msg = "This record already exists"
        End If
        MsgBox Msg, vbOKOnly + vbMsgBoxRight, App.title
        TxtVacName5.SetFocus
        Exit Sub
    End If
    Select Case Me.TxtModFlg5.Text
        Case "N"
            AddNewRec5
            btnLast5_Click

        Case "E"
            FiLLRec5
    End Select
    Exit Sub
ErrTrap:
    If SystemOptions.UserInterface = ArabicInterface Then
        MsgBox "Â‰«ﬂ Œÿ√ „« ›Ì ≈œŒ«· «·»Ì«‰« ", vbOKOnly + vbMsgBoxRight, App.title
    Else
        MsgBox "Record saved successfully", vbInformation + vbOKOnly + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
    End If
End Sub
 
Private Sub BtnUndo5_Click()
    FindRec5 val(TxtVac_ID5.Text)
    Me.TxtModFlg5.Text = "R"
End Sub
Private Sub BtnUpdate5_Click()

    On Error GoTo ErrTrap
    
    Dim Msg As String
    Dim FristCount As Long
    Dim LastCount As Long
    
    FristCount = RsSavRec5.RecordCount
    RsSavRec5.Requery
    LastCount = RsSavRec5.RecordCount
    BtnUndo5_Click
    If SystemOptions.UserInterface = ArabicInterface Then
        If FristCount = LastCount Then
            Msg = "·«  ÊÃœ »Ì«‰«  ÃœÌœ…"
        Else
            Msg = "⁄œœ «·”Ã·«  ﬁ»· «· ÕœÌÀ" & vbCrLf & FristCount & vbCrLf & "⁄œœ «·”Ã·«  »⁄œ «· ÕœÌÀ" & vbCrLf & LastCount
            If LastCount > FristCount Then
                Msg = Msg + vbCrLf & "⁄œœ «·”Ã·«  «·ÃœÌœ…" & vbCrLf & LastCount - FristCount
            Else
                Msg = Msg + vbCrLf & "⁄œœ «·”Ã·«  «·„Õ–Ê›…" & vbCrLf & FristCount - LastCount
            End If
        End If
    Else
        If FristCount = LastCount Then
            Msg = "No new data"
        Else
            Msg = "No. of records before the update" & vbCrLf & FristCount & vbCrLf & "No. of records after update" & vbCrLf & LastCount
        
            If LastCount > FristCount Then
                Msg = Msg + vbCrLf & "No. of new records" & vbCrLf & LastCount - FristCount
            Else
                Msg = Msg + vbCrLf & "No. of deleted records" & vbCrLf & FristCount - LastCount
            End If
        End If
    End If
    MsgBox Msg, vbOKOnly + vbMsgBoxRight + vbInformation, App.title
ErrTrap:
End Sub
Private Sub DcboExpensesID_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF5 Then
        Dim Dcombos As ClsDataCombos
        Set Dcombos = New ClsDataCombos
        Dcombos.GetExpensesNames Me.DcboExpensesID
    End If
End Sub
Public Sub AddNewRec5()

    On Error GoTo ErrTrap
    
    Dim StrRecId5 As String
    
    StrRecId5 = new_id("TBLProductionElements", "TBLProductionElementsId", "")
    RsSavRec5.AddNew
    RsSavRec5.Fields("TBLProductionElementsId").value = IIf(StrRecId5 <> "", StrRecId5, Null)
    FiLLRec5
ErrTrap:
End Sub
Function CuurentLogdata5(Optional Currentmode As String)
   
    LogTextA = "    ‘«‘… " & ScreenNameArabic & Chr(13) & " ﬂÊœ   " & TxtSerial5.Text & Chr(13) & "  «”„ «·⁄‰’— ⁄—»Ì " & TxtVacName5.Text & Chr(13) & "  «”„ «·„’—Ê› " & DcboExpensesID
    LogTextE = "    Screen  " & ScreenNameEnglish & Chr(13) & " Code   " & TxtSerial5.Text & Chr(13) & "Element English Name" & TxtVacNamee5.Text & Chr(13) & " Expenses Namae " & DcboExpensesID
    If Currentmode <> "D" Then
        AddToLogFile CInt(user_id), 0, Date, Time, LogTextA, LogTextE, Me.name, Me.TxtModFlg5
    Else
        AddToLogFile CInt(user_id), 0, Date, Time, LogTextA, LogTextE, Me.name, "D"
    End If
End Function
Public Sub FiLLRec5()

    On Error GoTo ErrTrap

    RsSavRec5.Fields("Name").value = IIf(TxtVacName5.Text <> "", Trim(TxtVacName5.Text), Null)
    RsSavRec5.Fields("Namee").value = IIf(TxtVacNamee5.Text <> "", Trim(TxtVacNamee5.Text), Null)
    RsSavRec5.Fields("ExpensesID").value = IIf(DcboExpensesID.BoundText <> 0, val(DcboExpensesID.BoundText), Null)
    RsSavRec5.update
    CuurentLogdata5
    If SystemOptions.UserInterface = ArabicInterface Then
        MsgBox " „  ⁄„·Ì… «·Õ›Ÿ »‰Ã«Õ", vbInformation + vbOKOnly + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
    Else
        MsgBox "Saves Successfully", vbInformation + vbOKOnly + vbMsgBoxRight + vbMsgBoxRtlReading, App.title

    End If
    FillGrid5WithData
    TxtModFlg5 = "R"
    Exit Sub
ErrTrap:
    If RsSavRec5.EditMode <> adEditNone Then
        RsSavRec5.CancelUpdate
    End If
End Sub
Public Sub FiLLTXT5()

    On Error GoTo ErrTrap
    
    Dim i As Integer
    
    Frm25.Enabled = False
    TxtVac_ID5.Text = IIf(IsNull(RsSavRec5.Fields("TBLProductionElementsId").value), "", RsSavRec5.Fields("TBLProductionElementsId").value)
    TxtVacName5.Text = IIf(IsNull(RsSavRec5.Fields("Name").value), "", RsSavRec5.Fields("Name").value)
    TxtVacNamee5.Text = IIf(IsNull(RsSavRec5.Fields("Namee").value), "", RsSavRec5.Fields("Namee").value)
    Me.DcboExpensesID.BoundText = IIf(IsNull(RsSavRec5.Fields("ExpensesID").value), "", RsSavRec5.Fields("ExpensesID").value)
    LabCurrRec5.Caption = RsSavRec5.AbsolutePosition
    LabCountRec5.Caption = RsSavRec5.RecordCount
    With Grid5
        For i = 1 To .Rows - 1
            If Trim(TxtVac_ID5.Text) = .TextMatrix(i, .ColIndex("TBLProductionElementsId")) Then
                TxtSerial5.Text = .TextMatrix(i, .ColIndex("Ser"))
                .Row = i
                Exit Sub
            End If
        Next
    End With
ErrTrap:
End Sub
Public Sub EditRec5(StrTable As String, RecId5 As String)
    FiLLRec5
End Sub
Private Sub Grid5_EnterCell()

    On Error GoTo ErrTrap
    
    FindRec5 val(Me.Grid5.TextMatrix(Me.Grid5.Row, Me.Grid5.ColIndex("TBLProductionElementsId")))
ErrTrap:
End Sub
Private Sub TxtVac_ID5_Change()

    Dim TxtMod As String
    
    TxtMod = TxtModFlg5.Text
    TxtModFlg5.Text = ""
    TxtModFlg5 = TxtMod
End Sub
Public Function FindRec5(ByVal RecId5 As Long)

    On Error GoTo ErrTrap
    
    RsSavRec5.Find "TBLProductionElementsId=" & RecId5, , adSearchForward, 1
    If Not (RsSavRec5.EOF) Then
        FiLLTXT5
    End If
    Exit Function
ErrTrap:
    If RsSavRec5.EditMode <> adEditNone Then
        RsSavRec5.CancelUpdate
        BtnUndo5_Click
    End If
End Function
Private Sub TxtModFlg5_Change()
    If TxtModFlg5.Text = "N" Then
        Frm25.Enabled = True
        Me.btnNew5.Enabled = False
        btnModify5.Enabled = False
        btnDelete5.Enabled = False
        Me.btnQuery5.Enabled = False
        Grid5.Enabled = False
        BtnUndo5.Enabled = True
        Me.btnSave5.Enabled = True
        BtnUpdate5.Enabled = False
    ElseIf TxtModFlg5.Text = "R" Then
        Frm25.Enabled = False
        Grid5.Enabled = True
        btnModify5.Enabled = False
        btnDelete5.Enabled = False
        If TxtVac_ID5.Text <> "" Then
            btnModify5.Enabled = True
            btnDelete5.Enabled = True
        End If
        BtnUpdate5.Enabled = True
        Me.btnQuery5.Enabled = True
        Me.btnNew5.Enabled = True
        BtnUndo5.Enabled = False
        Me.btnSave5.Enabled = False
        btnNext5.Enabled = True
        btnPrevious5.Enabled = True
        btnFirst5.Enabled = True
        btnLast5.Enabled = True
    ElseIf TxtModFlg5.Text = "E" Then
        Frm25.Enabled = True
        Me.btnNew5.Enabled = False
        btnModify5.Enabled = False
        btnDelete5.Enabled = False
        Me.btnQuery5.Enabled = False
        BtnUpdate5.Enabled = False
        BtnUndo5.Enabled = True
        Me.btnSave5.Enabled = True
        Grid5.Enabled = False
        btnNext5.Enabled = False
        btnPrevious5.Enabled = False
        btnFirst5.Enabled = False
        btnLast5.Enabled = False
    End If
End Sub
Public Sub FillGrid5WithData()

    On Error GoTo ErrTrap

    Dim i As Integer
    Dim RS As ADODB.Recordset
    Dim My_SQL As String
    Set RS = New ADODB.Recordset
    
    My_SQL = "select * From TBLProductionElements order by TBLProductionElementsId"
    RS.Open My_SQL, Cn, adOpenKeyset, adLockReadOnly, adCmdText
    With Me.Grid5
        .Rows = 2
        .Clear flexClearScrollable
        If RS.RecordCount > 0 Then
            .Rows = RS.RecordCount + 1
            RS.MoveFirst
            For i = 1 To .Rows - 1
                .TextMatrix(i, .ColIndex("Ser")) = i
                .TextMatrix(i, .ColIndex("Name")) = IIf(IsNull(RS.Fields("Name").value), "", RS.Fields("Name").value)
                .TextMatrix(i, .ColIndex("Namee")) = IIf(IsNull(RS.Fields("Namee").value), "", RS.Fields("Namee").value)
                .TextMatrix(i, .ColIndex("TBLProductionElementsId")) = IIf(IsNull(RS.Fields("TBLProductionElementsId").value), "", RS.Fields("TBLProductionElementsId").value)
                .TextMatrix(i, .ColIndex("ExpensesID")) = IIf(IsNull(RS.Fields("ExpensesID").value), "", RS.Fields("ExpensesID").value)
                RS.MoveNext
            Next
            RS.Close
        End If
        .RowHeight(-1) = 300
    End With
ErrTrap:
End Sub
Private Function CheckDelCountry5(LngExpensesID As Long) As Boolean

    Dim RS As ADODB.Recordset
    Dim StrSQL As String
    
    StrSQL = "Select * From TblEmployee Where TBLProductionElementsId=" & LngExpensesID & ""
    Set RS = New ADODB.Recordset
    RS.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText
    If Not (RS.BOF Or RS.EOF) Then
        CheckDelCountry5 = False
    Else
        CheckDelCountry5 = True
    End If
    RS.Close
    Set RS = Nothing
End Function
Private Sub TxtVacName5_GotFocus()
    SwitchKeyboardLang LANG_ARABIC
End Sub
Private Sub TxtVacNamee5_GotFocus()
    SwitchKeyboardLang LANG_ENGLISH
End Sub
