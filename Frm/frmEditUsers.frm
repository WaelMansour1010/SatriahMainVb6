VERSION 5.00
Object = "{C115893A-A3BF-43AF-B28D-69DB846077F3}#1.0#0"; "vsflex8u.ocx"
Object = "{0AFE7BE0-11B7-4A3E-978D-D4501E9A57FE}#1.0#0"; "c1sizer.ocx"
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{FE5DCFAD-BC1D-11D2-94CF-004005455FAA}#1.4#0"; "ImpulseButton.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form FrmEditUsers 
   ClientHeight    =   10950
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   19095
   Icon            =   "frmEditUsers.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MDIChild        =   -1  'True
   RightToLeft     =   -1  'True
   ScaleHeight     =   10950
   ScaleWidth      =   19095
   WindowState     =   2  'Maximized
   Begin C1SizerLibCtl.C1Tab TabMain 
      Height          =   10950
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   19095
      _cx             =   33681
      _cy             =   19315
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
      Caption         =   ""
      Align           =   5
      CurrTab         =   11
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
      Begin C1SizerLibCtl.C1Elastic ELe 
         Height          =   10605
         Index           =   1
         Left            =   30
         TabIndex        =   1
         TabStop         =   0   'False
         Top             =   30
         Width           =   19020
         _cx             =   33549
         _cy             =   18706
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
         Begin VB.TextBox txtCreditLimitSalesMan 
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
            IMEMode         =   3  'DISABLE
            Left            =   1905
            MaxLength       =   50
            RightToLeft     =   -1  'True
            TabIndex        =   91
            Top             =   9930
            Width           =   1665
         End
         Begin VB.TextBox TXtReportName2 
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
            IMEMode         =   3  'DISABLE
            Left            =   12795
            MaxLength       =   50
            RightToLeft     =   -1  'True
            TabIndex        =   89
            Top             =   9840
            Width           =   1770
         End
         Begin VB.TextBox TXtReportName1 
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
            IMEMode         =   3  'DISABLE
            Left            =   15885
            MaxLength       =   50
            RightToLeft     =   -1  'True
            TabIndex        =   87
            Top             =   10200
            Width           =   1485
         End
         Begin VB.TextBox TXtReportName 
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
            IMEMode         =   3  'DISABLE
            Left            =   15885
            MaxLength       =   50
            RightToLeft     =   -1  'True
            TabIndex        =   85
            Top             =   9840
            Width           =   1485
         End
         Begin VB.Frame Frm2 
            BackColor       =   &H00E2E9E9&
            BorderStyle     =   0  'None
            Enabled         =   0   'False
            Height          =   360
            Left            =   -600
            RightToLeft     =   -1  'True
            TabIndex        =   13
            Top             =   10440
            Visible         =   0   'False
            Width           =   8100
            Begin VB.ComboBox CmbType 
               BackColor       =   &H80000018&
               Height          =   315
               ItemData        =   "frmEditUsers.frx":058A
               Left            =   2280
               List            =   "frmEditUsers.frx":059A
               RightToLeft     =   -1  'True
               Style           =   2  'Dropdown List
               TabIndex        =   15
               Top             =   2670
               Visible         =   0   'False
               Width           =   1005
            End
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
               Left            =   15585
               MaxLength       =   50
               RightToLeft     =   -1  'True
               TabIndex        =   14
               Top             =   990
               Visible         =   0   'False
               Width           =   1410
            End
            Begin MSDataListLib.DataCombo DCSalesRepGroups 
               Height          =   315
               Left            =   120
               TabIndex        =   16
               Tag             =   "⁄›Ê« Ì—ÃÏ «œŒ«· √”„ «·„Ã„Ê⁄Â"
               Top             =   -360
               Width           =   3075
               _ExtentX        =   5424
               _ExtentY        =   556
               _Version        =   393216
               Text            =   ""
               RightToLeft     =   -1  'True
            End
            Begin MSDataListLib.DataCombo DCJob 
               Height          =   315
               Left            =   120
               TabIndex        =   17
               Tag             =   "⁄›Ê« Ì—ÃÏ «œŒ«·  «·ÊŸÌ€…"
               Top             =   -600
               Visible         =   0   'False
               Width           =   3075
               _ExtentX        =   5424
               _ExtentY        =   556
               _Version        =   393216
               Text            =   ""
               RightToLeft     =   -1  'True
            End
         End
         Begin VB.Frame FraHeader 
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   0  'None
            Height          =   525
            Left            =   150
            RightToLeft     =   -1  'True
            TabIndex        =   2
            Top             =   -30
            Width           =   18840
            Begin VB.Frame Frmo2 
               BackColor       =   &H00E0E0E0&
               BorderStyle     =   0  'None
               Enabled         =   0   'False
               Height          =   375
               Left            =   540
               RightToLeft     =   -1  'True
               TabIndex        =   5
               Top             =   450
               Visible         =   0   'False
               Width           =   3105
               Begin MSDataListLib.DataCombo DCUser 
                  CausesValidation=   0   'False
                  Height          =   315
                  Left            =   -255
                  TabIndex        =   6
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
                  TabIndex        =   7
                  Top             =   45
                  Width           =   855
               End
            End
            Begin VB.TextBox TxtModFlg 
               Alignment       =   1  'Right Justify
               Appearance      =   0  'Flat
               BackColor       =   &H0000FF00&
               Enabled         =   0   'False
               Height          =   285
               Left            =   2580
               RightToLeft     =   -1  'True
               TabIndex        =   4
               Text            =   "modflag"
               Top             =   90
               Visible         =   0   'False
               Width           =   465
            End
            Begin VB.TextBox TxtVac_ID 
               Alignment       =   1  'Right Justify
               Appearance      =   0  'Flat
               BackColor       =   &H000000FF&
               Height          =   285
               Left            =   3030
               RightToLeft     =   -1  'True
               TabIndex        =   3
               Top             =   510
               Visible         =   0   'False
               Width           =   1065
            End
            Begin MSComctlLib.ImageList GrdImageList 
               Left            =   3120
               Top             =   0
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
                     Picture         =   "frmEditUsers.frx":05B3
                     Key             =   "CompanyName"
                  EndProperty
                  BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                     Picture         =   "frmEditUsers.frx":094D
                     Key             =   "Ser"
                  EndProperty
                  BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                     Picture         =   "frmEditUsers.frx":0CE7
                     Key             =   "Vac_Name"
                  EndProperty
                  BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                     Picture         =   "frmEditUsers.frx":1081
                     Key             =   "ShareCount"
                  EndProperty
                  BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                     Picture         =   "frmEditUsers.frx":141B
                     Key             =   "Dis_Count"
                  EndProperty
                  BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                     Picture         =   "frmEditUsers.frx":17B5
                     Key             =   "Bouns"
                  EndProperty
                  BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                     Picture         =   "frmEditUsers.frx":1B4F
                     Key             =   "SharesValue"
                  EndProperty
                  BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                     Picture         =   "frmEditUsers.frx":20E9
                     Key             =   "BuyValue"
                  EndProperty
               EndProperty
            End
            Begin ImpulseButton.ISButton btnLast 
               Height          =   315
               Left            =   90
               TabIndex        =   8
               Top             =   30
               Width           =   405
               _ExtentX        =   714
               _ExtentY        =   556
               ButtonStyle     =   1
               ButtonPositionImage=   1
               Caption         =   ""
               BackColor       =   14871017
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
               ButtonImage     =   "frmEditUsers.frx":2483
               ColorButton     =   14871017
               AcclimateGrayTones=   -1  'True
               DrawFocusRectangle=   0   'False
               DisabledImageExtraction=   0
               DisabledImageStyle=   1
            End
            Begin ImpulseButton.ISButton btnNext 
               Height          =   315
               Left            =   555
               TabIndex        =   9
               Top             =   30
               Width           =   405
               _ExtentX        =   714
               _ExtentY        =   556
               ButtonStyle     =   1
               ButtonPositionImage=   1
               Caption         =   ""
               BackColor       =   14871017
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
               ButtonImage     =   "frmEditUsers.frx":281D
               ColorButton     =   14871017
               DrawFocusRectangle=   0   'False
               DisabledImageStyle=   1
            End
            Begin ImpulseButton.ISButton btnPrevious 
               Height          =   315
               Left            =   1155
               TabIndex        =   10
               Top             =   30
               Width           =   405
               _ExtentX        =   714
               _ExtentY        =   556
               ButtonStyle     =   1
               ButtonPositionImage=   1
               Caption         =   ""
               BackColor       =   14871017
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
               ButtonImage     =   "frmEditUsers.frx":2BB7
               ColorButton     =   14871017
               DrawFocusRectangle=   0   'False
               DisabledImageStyle=   1
            End
            Begin ImpulseButton.ISButton btnFirst 
               Height          =   315
               Left            =   1620
               TabIndex        =   11
               Top             =   30
               Width           =   405
               _ExtentX        =   714
               _ExtentY        =   556
               ButtonStyle     =   1
               ButtonPositionImage=   1
               Caption         =   ""
               BackColor       =   14871017
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
               ButtonImage     =   "frmEditUsers.frx":2F51
               ColorButton     =   14871017
               DrawFocusRectangle=   0   'False
               DisabledImageStyle=   1
            End
            Begin VB.Label Label1 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "»Ì«‰«  «·„” Œœ„Ì‰"
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
               Left            =   6360
               RightToLeft     =   -1  'True
               TabIndex        =   12
               Top             =   120
               Width           =   3255
            End
         End
         Begin C1SizerLibCtl.C1Elastic EltCont 
            Height          =   855
            Left            =   6120
            TabIndex        =   18
            TabStop         =   0   'False
            Top             =   9750
            Width           =   6750
            _cx             =   11906
            _cy             =   1508
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
               Height          =   330
               Left            =   5295
               TabIndex        =   19
               Top             =   435
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
               ButtonImage     =   "frmEditUsers.frx":32EB
               ColorButton     =   14871017
               DrawFocusRectangle=   0   'False
            End
            Begin ImpulseButton.ISButton btnSave 
               Height          =   330
               Left            =   3510
               TabIndex        =   20
               Top             =   435
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
               ButtonImage     =   "frmEditUsers.frx":3685
               ColorButton     =   14871017
               DrawFocusRectangle=   0   'False
            End
            Begin ImpulseButton.ISButton btnModify 
               Height          =   330
               Left            =   4395
               TabIndex        =   21
               Top             =   435
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
               ButtonImage     =   "frmEditUsers.frx":3A1F
               ColorButton     =   14871017
               DrawFocusRectangle=   0   'False
            End
            Begin ImpulseButton.ISButton BtnUndo 
               Height          =   330
               Left            =   2745
               TabIndex        =   22
               Top             =   435
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
               ButtonImage     =   "frmEditUsers.frx":3DB9
               ColorButton     =   14871017
               DrawFocusRectangle=   0   'False
            End
            Begin ImpulseButton.ISButton btnDelete 
               Height          =   330
               Left            =   1020
               TabIndex        =   23
               Top             =   435
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
               ButtonImage     =   "frmEditUsers.frx":4153
               ColorButton     =   14871017
               DrawFocusRectangle=   0   'False
            End
            Begin ImpulseButton.ISButton btnQuery 
               Height          =   330
               Left            =   1800
               TabIndex        =   24
               TabStop         =   0   'False
               ToolTipText     =   "(Ctrl+F)  ··»ÕÀ ≈÷€ÿ Â–« «·„› «Õ √Ê ≈÷€ÿ "
               Top             =   450
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
               ButtonImage     =   "frmEditUsers.frx":46ED
               ColorButton     =   14871017
               DrawFocusRectangle=   0   'False
            End
            Begin ImpulseButton.ISButton BtnUpdate 
               Height          =   330
               Left            =   5010
               TabIndex        =   25
               TabStop         =   0   'False
               ToolTipText     =   " ÕœÌÀ ﬁ«⁄œ… «·»Ì«‰« "
               Top             =   60
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
               ButtonImage     =   "frmEditUsers.frx":4A87
               ColorButton     =   14871017
               DrawFocusRectangle=   0   'False
            End
            Begin ImpulseButton.ISButton BtnPrint 
               Height          =   285
               Left            =   3960
               TabIndex        =   26
               TabStop         =   0   'False
               Top             =   90
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
               ButtonImage     =   "frmEditUsers.frx":4E21
               ColorButton     =   14871017
               DrawFocusRectangle=   0   'False
               DisabledImageStyle=   1
            End
            Begin ImpulseButton.ISButton btnCancel 
               Height          =   330
               Left            =   225
               TabIndex        =   27
               Top             =   435
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
               ButtonImage     =   "frmEditUsers.frx":51BB
               ColorButton     =   14871017
               DrawFocusRectangle=   0   'False
               DisabledImageStyle=   1
            End
            Begin VB.Label LabCountRec 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Height          =   210
               Left            =   720
               RightToLeft     =   -1  'True
               TabIndex        =   31
               Top             =   45
               Width           =   540
            End
            Begin VB.Label LabCurrRec 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Height          =   210
               Left            =   2280
               RightToLeft     =   -1  'True
               TabIndex        =   30
               Top             =   0
               Width           =   675
            End
            Begin VB.Label Label2 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "⁄œœ «·”Ã·« :"
               Height          =   450
               Index           =   1
               Left            =   930
               RightToLeft     =   -1  'True
               TabIndex        =   29
               Top             =   45
               Width           =   1215
            End
            Begin VB.Label Label2 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "«·”Ã· «·Õ«·Ì:"
               Height          =   210
               Index           =   0
               Left            =   2895
               RightToLeft     =   -1  'True
               TabIndex        =   28
               Top             =   75
               Width           =   975
            End
         End
         Begin VSFlex8UCtl.VSFlexGrid Grid 
            Height          =   1620
            Left            =   30
            TabIndex        =   32
            Top             =   525
            Width           =   5835
            _cx             =   10292
            _cy             =   2857
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
            Cols            =   13
            FixedRows       =   1
            FixedCols       =   1
            RowHeightMin    =   320
            RowHeightMax    =   0
            ColWidthMin     =   0
            ColWidthMax     =   0
            ExtendLastCol   =   -1  'True
            FormatString    =   $"frmEditUsers.frx":5555
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
         Begin C1SizerLibCtl.C1Elastic ELe 
            Height          =   600
            Index           =   6
            Left            =   30
            TabIndex        =   33
            TabStop         =   0   'False
            Top             =   2160
            Width           =   5835
            _cx             =   10292
            _cy             =   1058
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
            ForeColor       =   192
            FloodColor      =   6553600
            ForeColorDisabled=   -2147483631
            Caption         =   "«· ÊﬁÌ⁄ «·«·ﬂ —Ê‰Ì"
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
            Begin Dynamic_Byte.NewViewBox ImgPic 
               Height          =   270
               Left            =   120
               TabIndex        =   34
               ToolTipText     =   "≈÷€ÿ ⁄·Ï «·’Ê—… „— Ì‰ ·· ﬂ»Ì—"
               Top             =   210
               Width           =   3030
               _ExtentX        =   5345
               _ExtentY        =   476
            End
            Begin ImpulseButton.ISButton CmdPic 
               Height          =   240
               Index           =   0
               Left            =   4680
               TabIndex        =   35
               Top             =   240
               Width           =   1260
               _ExtentX        =   2223
               _ExtentY        =   423
               ButtonStyle     =   1
               ButtonPositionImage=   1
               Caption         =   "≈÷«›… ’Ê—…"
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
               ButtonImage     =   "frmEditUsers.frx":5730
               ColorButton     =   14871017
               Alignment       =   1
               DrawFocusRectangle=   0   'False
            End
            Begin ImpulseButton.ISButton CmdPic 
               Height          =   225
               Index           =   1
               Left            =   3360
               TabIndex        =   36
               Top             =   240
               Width           =   1380
               _ExtentX        =   2434
               _ExtentY        =   397
               ButtonStyle     =   1
               ButtonPositionImage=   1
               Caption         =   "Õ–› «·’Ê—…"
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
               ButtonImage     =   "frmEditUsers.frx":5ACA
               ColorButton     =   14871017
               Alignment       =   1
               DrawFocusRectangle=   0   'False
            End
         End
         Begin VSFlex8Ctl.VSFlexGrid FG 
            Height          =   1620
            Left            =   120
            TabIndex        =   37
            Top             =   3435
            Width           =   5265
            _cx             =   9287
            _cy             =   2857
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
            Rows            =   1
            Cols            =   4
            FixedRows       =   1
            FixedCols       =   0
            RowHeightMin    =   300
            RowHeightMax    =   0
            ColWidthMin     =   0
            ColWidthMax     =   0
            ExtendLastCol   =   0   'False
            FormatString    =   $"frmEditUsers.frx":5E64
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
         End
         Begin C1SizerLibCtl.C1Elastic ELe 
            Height          =   1725
            Index           =   0
            Left            =   12210
            TabIndex        =   38
            TabStop         =   0   'False
            Top             =   3360
            Width           =   6240
            _cx             =   11007
            _cy             =   3043
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
            Begin VB.ListBox ListGroupAll 
               Height          =   1035
               ItemData        =   "frmEditUsers.frx":5F28
               Left            =   3450
               List            =   "frmEditUsers.frx":5F2F
               RightToLeft     =   -1  'True
               TabIndex        =   41
               Top             =   240
               Width           =   2310
            End
            Begin VB.ListBox ListGroupSelected 
               BackColor       =   &H0080FFFF&
               Height          =   1035
               ItemData        =   "frmEditUsers.frx":5F41
               Left            =   165
               List            =   "frmEditUsers.frx":5F48
               RightToLeft     =   -1  'True
               TabIndex        =   40
               Top             =   240
               Width           =   2850
            End
            Begin VSFlex8UCtl.VSFlexGrid FgItems 
               Height          =   1710
               Index           =   0
               Left            =   8595
               TabIndex        =   39
               Top             =   150
               Width           =   6240
               _cx             =   11007
               _cy             =   3016
               Appearance      =   2
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
               RowHeightMin    =   300
               RowHeightMax    =   0
               ColWidthMin     =   0
               ColWidthMax     =   0
               ExtendLastCol   =   -1  'True
               FormatString    =   $"frmEditUsers.frx":5F5F
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
            Begin VB.Label Label32 
               Alignment       =   2  'Center
               BackStyle       =   0  'Transparent
               Caption         =   "«·›—Ê⁄ «·„— »ÿ…"
               Height          =   225
               Index           =   0
               Left            =   5355
               RightToLeft     =   -1  'True
               TabIndex        =   59
               Top             =   30
               Width           =   735
            End
            Begin VB.Label Label8 
               Alignment       =   2  'Center
               BackStyle       =   0  'Transparent
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
               Height          =   270
               Left            =   3135
               RightToLeft     =   -1  'True
               TabIndex        =   47
               Top             =   240
               Width           =   240
            End
            Begin VB.Label Label7 
               Alignment       =   2  'Center
               BackStyle       =   0  'Transparent
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
               Height          =   360
               Left            =   3135
               RightToLeft     =   -1  'True
               TabIndex        =   46
               Top             =   495
               Width           =   240
            End
            Begin VB.Label Label6 
               Alignment       =   2  'Center
               BackStyle       =   0  'Transparent
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
               Height          =   330
               Left            =   3135
               RightToLeft     =   -1  'True
               TabIndex        =   45
               Top             =   840
               Width           =   240
            End
            Begin VB.Label Label5 
               Alignment       =   2  'Center
               BackStyle       =   0  'Transparent
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
               Height          =   480
               Left            =   3135
               RightToLeft     =   -1  'True
               TabIndex        =   44
               Top             =   1095
               Width           =   240
            End
            Begin VB.Label Label11 
               Alignment       =   2  'Center
               BackStyle       =   0  'Transparent
               Caption         =   "Õœœ«·›—Ê⁄"
               Height          =   225
               Left            =   4485
               RightToLeft     =   -1  'True
               TabIndex        =   43
               Top             =   30
               Width           =   690
            End
            Begin VB.Label Label12 
               Alignment       =   2  'Center
               BackStyle       =   0  'Transparent
               Caption         =   "«·›—Ê⁄ «·„Õœœ…"
               Height          =   225
               Left            =   2355
               RightToLeft     =   -1  'True
               TabIndex        =   42
               Top             =   30
               Width           =   660
            End
         End
         Begin C1SizerLibCtl.C1Elastic ELe 
            Height          =   4560
            Index           =   2
            Left            =   5505
            TabIndex        =   48
            TabStop         =   0   'False
            Top             =   5115
            Width           =   6645
            _cx             =   11721
            _cy             =   8043
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
            Begin VB.CommandButton cmdReloadList 
               Caption         =   "«·€«¡ «·„Õœœ"
               Height          =   195
               Index           =   1
               Left            =   3480
               RightToLeft     =   -1  'True
               TabIndex        =   52
               Top             =   390
               Width           =   2955
            End
            Begin VB.ListBox ListStoreSelected 
               BackColor       =   &H0080FFFF&
               Height          =   3375
               ItemData        =   "frmEditUsers.frx":601F
               Left            =   240
               List            =   "frmEditUsers.frx":6026
               MultiSelect     =   1  'Simple
               RightToLeft     =   -1  'True
               TabIndex        =   51
               Top             =   645
               Width           =   2745
            End
            Begin VB.ListBox ListStoreall 
               Height          =   3375
               ItemData        =   "frmEditUsers.frx":603D
               Left            =   3660
               List            =   "frmEditUsers.frx":6044
               MultiSelect     =   1  'Simple
               RightToLeft     =   -1  'True
               TabIndex        =   50
               Top             =   645
               Width           =   2775
            End
            Begin VSFlex8UCtl.VSFlexGrid FgItems 
               Height          =   4980
               Index           =   1
               Left            =   9285
               TabIndex        =   49
               Top             =   495
               Width           =   6555
               _cx             =   11562
               _cy             =   8784
               Appearance      =   2
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
               RowHeightMin    =   300
               RowHeightMax    =   0
               ColWidthMin     =   0
               ColWidthMax     =   0
               ExtendLastCol   =   -1  'True
               FormatString    =   $"frmEditUsers.frx":6056
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
            Begin VB.Label Label32 
               Alignment       =   2  'Center
               BackStyle       =   0  'Transparent
               Caption         =   "«·„Œ«“‰ «·„— »ÿ…"
               Height          =   285
               Index           =   1
               Left            =   5100
               RightToLeft     =   -1  'True
               TabIndex        =   60
               Top             =   30
               Width           =   1575
            End
            Begin VB.Label Label4 
               Alignment       =   2  'Center
               BackStyle       =   0  'Transparent
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
               Height          =   780
               Left            =   3105
               RightToLeft     =   -1  'True
               TabIndex        =   58
               Top             =   2880
               Width           =   495
            End
            Begin VB.Label Label3 
               Alignment       =   2  'Center
               BackStyle       =   0  'Transparent
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
               Height          =   615
               Left            =   3105
               RightToLeft     =   -1  'True
               TabIndex        =   57
               Top             =   2475
               Width           =   495
            End
            Begin VB.Label Label22 
               Alignment       =   2  'Center
               BackStyle       =   0  'Transparent
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
               Height          =   630
               Left            =   3105
               RightToLeft     =   -1  'True
               TabIndex        =   56
               Top             =   1890
               Width           =   495
            End
            Begin VB.Label LblSelect 
               Alignment       =   2  'Center
               BackStyle       =   0  'Transparent
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
               Height          =   585
               Left            =   3105
               RightToLeft     =   -1  'True
               TabIndex        =   55
               Top             =   1305
               Width           =   495
            End
            Begin VB.Label Label9 
               Alignment       =   2  'Center
               BackStyle       =   0  'Transparent
               Caption         =   "Õœœ «·„Œ«“‰"
               Height          =   720
               Left            =   4155
               RightToLeft     =   -1  'True
               TabIndex        =   54
               Top             =   75
               Visible         =   0   'False
               Width           =   1020
            End
            Begin VB.Label Label10 
               Alignment       =   2  'Center
               BackStyle       =   0  'Transparent
               Caption         =   "«·„Œ«“‰ «·„Õœœ…"
               Height          =   465
               Left            =   1275
               RightToLeft     =   -1  'True
               TabIndex        =   53
               Top             =   30
               Width           =   1515
            End
         End
         Begin C1SizerLibCtl.C1Elastic ELe 
            Height          =   4500
            Index           =   3
            Left            =   12210
            TabIndex        =   61
            TabStop         =   0   'False
            Top             =   5235
            Width           =   6240
            _cx             =   11007
            _cy             =   7938
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
            Begin VB.CommandButton cmdReloadList 
               Caption         =   "«·€«¡ «·„Õœœ"
               Height          =   210
               Index           =   2
               Left            =   3480
               RightToLeft     =   -1  'True
               TabIndex        =   66
               Top             =   300
               Width           =   2745
            End
            Begin VB.ListBox ListBoxesAll 
               Height          =   3375
               ItemData        =   "frmEditUsers.frx":6116
               Left            =   3480
               List            =   "frmEditUsers.frx":611D
               MultiSelect     =   1  'Simple
               RightToLeft     =   -1  'True
               TabIndex        =   65
               Top             =   600
               Width           =   2610
            End
            Begin VB.ListBox ListBoxesSelected 
               BackColor       =   &H0080FFFF&
               Height          =   3375
               ItemData        =   "frmEditUsers.frx":612F
               Left            =   150
               List            =   "frmEditUsers.frx":6136
               RightToLeft     =   -1  'True
               TabIndex        =   64
               Top             =   600
               Width           =   2865
            End
            Begin VSFlex8UCtl.VSFlexGrid FgItems 
               Height          =   4830
               Index           =   2
               Left            =   8310
               TabIndex        =   62
               Top             =   525
               Width           =   6195
               _cx             =   10927
               _cy             =   8520
               Appearance      =   2
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
               RowHeightMin    =   300
               RowHeightMax    =   0
               ColWidthMin     =   0
               ColWidthMax     =   0
               ExtendLastCol   =   -1  'True
               FormatString    =   $"frmEditUsers.frx":614D
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
            Begin VB.Label Label18 
               Alignment       =   2  'Center
               BackStyle       =   0  'Transparent
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
               Height          =   285
               Left            =   2970
               RightToLeft     =   -1  'True
               TabIndex        =   72
               Top             =   1170
               Width           =   510
            End
            Begin VB.Label Label17 
               Alignment       =   2  'Center
               BackStyle       =   0  'Transparent
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
               Height          =   630
               Left            =   2970
               RightToLeft     =   -1  'True
               TabIndex        =   71
               Top             =   1890
               Width           =   510
            End
            Begin VB.Label Label16 
               Alignment       =   2  'Center
               BackStyle       =   0  'Transparent
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
               Height          =   510
               Left            =   2970
               RightToLeft     =   -1  'True
               TabIndex        =   70
               Top             =   2505
               Width           =   510
            End
            Begin VB.Label Label15 
               Alignment       =   2  'Center
               BackStyle       =   0  'Transparent
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
               Height          =   615
               Left            =   2970
               RightToLeft     =   -1  'True
               TabIndex        =   69
               Top             =   3000
               Width           =   510
            End
            Begin VB.Label Label14 
               Alignment       =   2  'Center
               BackStyle       =   0  'Transparent
               Caption         =   "Õœœ«·Œ“‰"
               Height          =   585
               Left            =   3960
               RightToLeft     =   -1  'True
               TabIndex        =   68
               Top             =   60
               Visible         =   0   'False
               Width           =   780
            End
            Begin VB.Label Label13 
               Alignment       =   2  'Center
               BackStyle       =   0  'Transparent
               Caption         =   "«·Œ“‰ «·„Õœœ…"
               Height          =   585
               Left            =   915
               RightToLeft     =   -1  'True
               TabIndex        =   67
               Top             =   60
               Width           =   1005
            End
            Begin VB.Label Label32 
               Alignment       =   2  'Center
               BackStyle       =   0  'Transparent
               Caption         =   "«·Œ“‰ «·„— »ÿ…"
               Height          =   225
               Index           =   2
               Left            =   5070
               RightToLeft     =   -1  'True
               TabIndex        =   63
               Top             =   60
               Width           =   1125
            End
         End
         Begin C1SizerLibCtl.C1Elastic ELe 
            Height          =   4515
            Index           =   4
            Left            =   -150
            TabIndex        =   73
            TabStop         =   0   'False
            Top             =   5115
            Width           =   5655
            _cx             =   9975
            _cy             =   7964
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
            Begin VB.CommandButton cmdReloadList 
               Caption         =   "«·€«¡ «·„Õœœ"
               Height          =   195
               Index           =   3
               Left            =   2985
               RightToLeft     =   -1  'True
               TabIndex        =   78
               Top             =   360
               Width           =   2655
            End
            Begin VB.ListBox ListAccountSelect 
               BackColor       =   &H0080FFFF&
               Height          =   3375
               ItemData        =   "frmEditUsers.frx":620D
               Left            =   150
               List            =   "frmEditUsers.frx":6214
               MultiSelect     =   1  'Simple
               RightToLeft     =   -1  'True
               TabIndex        =   77
               Top             =   570
               Width           =   2460
            End
            Begin VB.ListBox ListAllAccount 
               Height          =   3375
               ItemData        =   "frmEditUsers.frx":622B
               Left            =   3195
               List            =   "frmEditUsers.frx":6232
               MultiSelect     =   1  'Simple
               RightToLeft     =   -1  'True
               TabIndex        =   76
               Top             =   570
               Width           =   2310
            End
            Begin VSFlex8UCtl.VSFlexGrid FgItems 
               Height          =   4725
               Index           =   3
               Left            =   7485
               TabIndex        =   74
               Top             =   525
               Width           =   5535
               _cx             =   9763
               _cy             =   8334
               Appearance      =   2
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
               RowHeightMin    =   300
               RowHeightMax    =   0
               ColWidthMin     =   0
               ColWidthMax     =   0
               ExtendLastCol   =   -1  'True
               FormatString    =   $"frmEditUsers.frx":6244
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
            Begin VB.Label Label25 
               Alignment       =   2  'Center
               BackStyle       =   0  'Transparent
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
               Height          =   630
               Left            =   2460
               RightToLeft     =   -1  'True
               TabIndex        =   84
               Top             =   3105
               Width           =   870
            End
            Begin VB.Label Label24 
               Alignment       =   2  'Center
               BackStyle       =   0  'Transparent
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
               Height          =   555
               Left            =   2460
               RightToLeft     =   -1  'True
               TabIndex        =   83
               Top             =   2565
               Width           =   870
            End
            Begin VB.Label Label23 
               Alignment       =   2  'Center
               BackStyle       =   0  'Transparent
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
               Height          =   360
               Left            =   2460
               RightToLeft     =   -1  'True
               TabIndex        =   82
               Top             =   1830
               Width           =   870
            End
            Begin VB.Label Label21 
               Alignment       =   2  'Center
               BackStyle       =   0  'Transparent
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
               Height          =   645
               Left            =   2460
               RightToLeft     =   -1  'True
               TabIndex        =   81
               Top             =   1200
               Width           =   870
            End
            Begin VB.Label Label20 
               Alignment       =   2  'Center
               BackStyle       =   0  'Transparent
               Caption         =   "Õœœ «·Õ”«»« "
               Height          =   630
               Left            =   3960
               RightToLeft     =   -1  'True
               TabIndex        =   80
               Top             =   180
               Width           =   1305
            End
            Begin VB.Label Label19 
               Alignment       =   2  'Center
               BackStyle       =   0  'Transparent
               Caption         =   "«·Õ”«»«  «·„Õœœ…"
               Height          =   630
               Left            =   390
               RightToLeft     =   -1  'True
               TabIndex        =   79
               Top             =   60
               Width           =   1575
            End
            Begin VB.Label Label32 
               Alignment       =   2  'Center
               BackStyle       =   0  'Transparent
               Caption         =   "«·Õ”«»«  «·„— »ÿ…"
               Height          =   630
               Index           =   3
               Left            =   4230
               RightToLeft     =   -1  'True
               TabIndex        =   75
               Top             =   -15
               Width           =   1395
            End
         End
         Begin C1SizerLibCtl.C1Elastic ELe 
            Height          =   2565
            Index           =   5
            Left            =   6015
            TabIndex        =   93
            TabStop         =   0   'False
            Top             =   600
            Width           =   12495
            _cx             =   22040
            _cy             =   4524
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
            Begin VB.CheckBox chkHidLowering 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "«ŸÂ«— «·«⁄„œ… ﬂ«„·… ›Ì   ‰»ÌÂ«  «·«‰ «Ã"
               Height          =   195
               Left            =   615
               RightToLeft     =   -1  'True
               TabIndex        =   125
               Top             =   1830
               Width           =   3255
            End
            Begin VB.CheckBox isDeactivatedchk 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "≈Ìﬁ«› «·„” Œœ„"
               Height          =   195
               Left            =   0
               RightToLeft     =   -1  'True
               TabIndex        =   124
               Top             =   720
               Width           =   1905
            End
            Begin VB.CheckBox chkNextLogin 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   " €ÌÌ— ﬂ·„… «·„—Ê— ⁄‰œ «·œŒÊ·"
               Height          =   195
               Left            =   2085
               RightToLeft     =   -1  'True
               TabIndex        =   123
               Top             =   720
               Width           =   2295
            End
            Begin VB.ComboBox CboPriv 
               BackColor       =   &H00C0FFFF&
               Height          =   315
               ItemData        =   "frmEditUsers.frx":6304
               Left            =   90
               List            =   "frmEditUsers.frx":630E
               RightToLeft     =   -1  'True
               TabIndex        =   122
               Top             =   0
               Width           =   2445
            End
            Begin VB.CheckBox AllowSelectEmp 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "«Œ Ì«— «·„‰œÊ» ›Ì «·ﬁ»÷"
               Height          =   195
               Left            =   855
               RightToLeft     =   -1  'True
               TabIndex        =   121
               Top             =   2040
               Width           =   3015
            End
            Begin VB.TextBox TxtSearchCode1 
               Alignment       =   2  'Center
               Height          =   345
               Left            =   6540
               TabIndex        =   100
               Top             =   1560
               Width           =   540
            End
            Begin VB.TextBox TxtSearchCode 
               Alignment       =   2  'Center
               Height          =   345
               Left            =   6540
               TabIndex        =   99
               Top             =   1200
               Width           =   540
            End
            Begin VB.TextBox XPTxtUserName 
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
               Left            =   8565
               MaxLength       =   50
               RightToLeft     =   -1  'True
               TabIndex        =   98
               Top             =   480
               Width           =   2025
            End
            Begin VB.TextBox XPTxtPassConfirm 
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
               IMEMode         =   3  'DISABLE
               Left            =   4380
               MaxLength       =   50
               PasswordChar    =   "#"
               RightToLeft     =   -1  'True
               TabIndex        =   97
               Tag             =   "⁄›Ê« Ì—ÃÏ   ‰”»… «·Œ’„"
               Top             =   870
               Width           =   2700
            End
            Begin VB.TextBox TXTCode 
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
               Left            =   8565
               MaxLength       =   50
               RightToLeft     =   -1  'True
               TabIndex        =   96
               Top             =   120
               Width           =   2025
            End
            Begin VB.TextBox TxtPassWord 
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
               IMEMode         =   3  'DISABLE
               Left            =   8565
               MaxLength       =   50
               PasswordChar    =   "#"
               RightToLeft     =   -1  'True
               TabIndex        =   95
               Tag             =   "⁄›Ê« Ì—ÃÏ   ‰”»… «·Œ’„"
               Top             =   870
               Width           =   2025
            End
            Begin VSFlex8UCtl.VSFlexGrid FgItems 
               Height          =   2550
               Index           =   4
               Left            =   17130
               TabIndex        =   94
               Top             =   255
               Width           =   12465
               _cx             =   21987
               _cy             =   4498
               Appearance      =   2
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
               RowHeightMin    =   300
               RowHeightMax    =   0
               ColWidthMin     =   0
               ColWidthMax     =   0
               ExtendLastCol   =   -1  'True
               FormatString    =   $"frmEditUsers.frx":632A
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
            Begin MSDataListLib.DataCombo DCEmP 
               Height          =   315
               Left            =   4380
               TabIndex        =   101
               Tag             =   "⁄›Ê« Ì—ÃÏ «œŒ«· «Œ Ì«— «·„‰œÊ»"
               Top             =   120
               Width           =   2700
               _ExtentX        =   4763
               _ExtentY        =   556
               _Version        =   393216
               Text            =   ""
               RightToLeft     =   -1  'True
            End
            Begin MSDataListLib.DataCombo DcBranches 
               Height          =   315
               Left            =   4380
               TabIndex        =   102
               Tag             =   "⁄›Ê« Ì—ÃÏ «œŒ«· √”„ «·›—⁄"
               Top             =   480
               Width           =   2700
               _ExtentX        =   4763
               _ExtentY        =   556
               _Version        =   393216
               Text            =   ""
               RightToLeft     =   -1  'True
            End
            Begin MSDataListLib.DataCombo DCStore 
               Height          =   315
               Left            =   8565
               TabIndex        =   103
               Tag             =   "⁄›Ê« Ì—ÃÏ «œŒ«· √”„ «·›—⁄"
               Top             =   1200
               Width           =   2025
               _ExtentX        =   3572
               _ExtentY        =   556
               _Version        =   393216
               Text            =   ""
               RightToLeft     =   -1  'True
            End
            Begin MSDataListLib.DataCombo DBCboClientName 
               Height          =   315
               Left            =   4380
               TabIndex        =   104
               Top             =   1200
               Width           =   2070
               _ExtentX        =   3651
               _ExtentY        =   556
               _Version        =   393216
               ListField       =   "6"
               BoundColumn     =   ""
               Text            =   ""
               RightToLeft     =   -1  'True
            End
            Begin MSDataListLib.DataCombo DCStore1 
               Height          =   315
               Left            =   8565
               TabIndex        =   105
               Tag             =   "⁄›Ê« Ì—ÃÏ «œŒ«· √”„ «·›—⁄"
               Top             =   1560
               Width           =   2025
               _ExtentX        =   3572
               _ExtentY        =   556
               _Version        =   393216
               Text            =   ""
               RightToLeft     =   -1  'True
            End
            Begin MSDataListLib.DataCombo DBCboClientName1 
               Height          =   315
               Left            =   4380
               TabIndex        =   106
               Top             =   1560
               Width           =   2070
               _ExtentX        =   3651
               _ExtentY        =   556
               _Version        =   393216
               ListField       =   "6"
               BoundColumn     =   ""
               Text            =   ""
               RightToLeft     =   -1  'True
            End
            Begin MSDataListLib.DataCombo DCStore2 
               Height          =   315
               Left            =   8565
               TabIndex        =   107
               Tag             =   "⁄›Ê« Ì—ÃÏ «œŒ«· √”„ «·›—⁄"
               Top             =   1890
               Width           =   2025
               _ExtentX        =   3572
               _ExtentY        =   556
               _Version        =   393216
               Text            =   ""
               RightToLeft     =   -1  'True
            End
            Begin MSDataListLib.DataCombo DCStore3 
               Height          =   315
               Left            =   4380
               TabIndex        =   108
               Tag             =   "⁄›Ê« Ì—ÃÏ «œŒ«· √”„ «·›—⁄"
               Top             =   1890
               Width           =   2700
               _ExtentX        =   4763
               _ExtentY        =   556
               _Version        =   393216
               Text            =   ""
               RightToLeft     =   -1  'True
            End
            Begin MSComDlg.CommonDialog cdg 
               Left            =   4680
               Top             =   240
               _ExtentX        =   847
               _ExtentY        =   847
               _Version        =   393216
            End
            Begin MSDataListLib.DataCombo DCBoxes 
               Height          =   315
               Left            =   90
               TabIndex        =   126
               Tag             =   "⁄›Ê« Ì—ÃÏ «œŒ«· √”„ «·›—⁄"
               Top             =   960
               Width           =   2445
               _ExtentX        =   4313
               _ExtentY        =   556
               _Version        =   393216
               Text            =   ""
               RightToLeft     =   -1  'True
            End
            Begin MSDataListLib.DataCombo Dbanks 
               Height          =   315
               Left            =   90
               TabIndex        =   127
               Tag             =   "⁄›Ê« Ì—ÃÏ «œŒ«· √”„ «·›—⁄"
               Top             =   360
               Width           =   2445
               _ExtentX        =   4313
               _ExtentY        =   556
               _Version        =   393216
               Text            =   ""
               RightToLeft     =   -1  'True
            End
            Begin MSDataListLib.DataCombo DCBoxes1 
               Height          =   315
               Left            =   90
               TabIndex        =   128
               Tag             =   "⁄›Ê« Ì—ÃÏ «œŒ«· √”„ «·›—⁄"
               Top             =   1320
               Width           =   2445
               _ExtentX        =   4313
               _ExtentY        =   556
               _Version        =   393216
               Text            =   ""
               RightToLeft     =   -1  'True
            End
            Begin VB.Label Label1 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "«·Œ“Ì‰… «·«› —«÷Ì… ··‘—«¡"
               Height          =   285
               Index           =   14
               Left            =   2595
               RightToLeft     =   -1  'True
               TabIndex        =   132
               Top             =   1440
               Width           =   1680
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "’·«ÕÌ« "
               Height          =   330
               Index           =   4
               Left            =   3330
               RightToLeft     =   -1  'True
               TabIndex        =   131
               Top             =   30
               Width           =   570
            End
            Begin VB.Label Label1 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "«·»‰ﬂ «·«› —«÷Ì"
               Height          =   285
               Index           =   6
               Left            =   2850
               RightToLeft     =   -1  'True
               TabIndex        =   130
               Top             =   360
               Width           =   1095
            End
            Begin VB.Label Label1 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "«·Œ“Ì‰… «·«› —«÷Ì… ··Ì⁄"
               Height          =   285
               Index           =   4
               Left            =   2745
               RightToLeft     =   -1  'True
               TabIndex        =   129
               Top             =   1080
               Width           =   1530
            End
            Begin VB.Label Label1 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "„Œ“‰ «” ·«„ «·„Ê«œ «·Œ«„"
               Height          =   195
               Index           =   16
               Left            =   7155
               RightToLeft     =   -1  'True
               TabIndex        =   120
               Top             =   1950
               Width           =   1350
            End
            Begin VB.Label Label1 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "„Œ“‰ ’—› «·„Ê«œ «·Œ«„"
               Height          =   195
               Index           =   15
               Left            =   10650
               RightToLeft     =   -1  'True
               TabIndex        =   119
               Top             =   1950
               Width           =   1305
            End
            Begin VB.Label Label1 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "„Œ“‰ «·‘—«¡ «·«› —«÷Ì"
               Height          =   195
               Index           =   12
               Left            =   10650
               RightToLeft     =   -1  'True
               TabIndex        =   118
               Top             =   1620
               Width           =   1275
            End
            Begin VB.Label Label1 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "«·„Ê—œ «·«› —«÷Ì"
               Height          =   285
               Index           =   11
               Left            =   7230
               RightToLeft     =   -1  'True
               TabIndex        =   117
               Top             =   1620
               Width           =   1290
            End
            Begin VB.Label Label1 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "«·⁄„Ì· «·«› —«÷Ì"
               Height          =   285
               Index           =   10
               Left            =   7230
               RightToLeft     =   -1  'True
               TabIndex        =   116
               Top             =   1260
               Width           =   1290
            End
            Begin VB.Label Label1 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "«”„ «·„” Œœ„"
               Height          =   195
               Index           =   9
               Left            =   10650
               RightToLeft     =   -1  'True
               TabIndex        =   115
               Top             =   540
               Width           =   1275
            End
            Begin VB.Label Label1 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   " √ﬂÌœ ﬂ·„… «·”—"
               Height          =   285
               Index           =   8
               Left            =   7230
               RightToLeft     =   -1  'True
               TabIndex        =   114
               Top             =   900
               Width           =   1170
            End
            Begin VB.Label Label1 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "„Œ“‰ «·»Ì⁄ «·«› —«÷Ì"
               Height          =   195
               Index           =   7
               Left            =   10650
               RightToLeft     =   -1  'True
               TabIndex        =   113
               Top             =   1260
               Width           =   1275
            End
            Begin VB.Label Label1 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "«·›—⁄"
               Height          =   285
               Index           =   5
               Left            =   7335
               RightToLeft     =   -1  'True
               TabIndex        =   112
               Top             =   480
               Width           =   945
            End
            Begin VB.Label Label1 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "ﬂ·„… «·”—"
               Height          =   195
               Index           =   0
               Left            =   10650
               RightToLeft     =   -1  'True
               TabIndex        =   111
               Top             =   900
               Width           =   1275
            End
            Begin VB.Label Label1 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "ﬂÊœ «·„” Œœ„"
               Height          =   195
               Index           =   3
               Left            =   10650
               RightToLeft     =   -1  'True
               TabIndex        =   110
               Top             =   150
               Width           =   1275
            End
            Begin VB.Label Label1 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "«”„ «·„ÊŸ›"
               Height          =   285
               Index           =   1
               Left            =   7590
               RightToLeft     =   -1  'True
               TabIndex        =   109
               Top             =   150
               Width           =   930
            End
         End
         Begin C1SizerLibCtl.C1Elastic ELe 
            Height          =   1845
            Index           =   7
            Left            =   5880
            TabIndex        =   133
            TabStop         =   0   'False
            Top             =   3240
            Width           =   6120
            _cx             =   10795
            _cy             =   3254
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
            Begin VB.ListBox ListProductLineAll 
               Height          =   1230
               ItemData        =   "frmEditUsers.frx":63EA
               Left            =   3450
               List            =   "frmEditUsers.frx":63F1
               RightToLeft     =   -1  'True
               TabIndex        =   138
               Top             =   330
               Width           =   2565
            End
            Begin VB.ListBox ListProductLineSelected 
               BackColor       =   &H0080FFFF&
               Height          =   1230
               ItemData        =   "frmEditUsers.frx":6403
               Left            =   0
               List            =   "frmEditUsers.frx":640A
               RightToLeft     =   -1  'True
               TabIndex        =   137
               Top             =   360
               Width           =   2775
            End
            Begin VSFlex8UCtl.VSFlexGrid FgItems 
               Height          =   1830
               Index           =   5
               Left            =   8430
               TabIndex        =   134
               Top             =   165
               Width           =   6120
               _cx             =   10795
               _cy             =   3228
               Appearance      =   2
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
               RowHeightMin    =   300
               RowHeightMax    =   0
               ColWidthMin     =   0
               ColWidthMax     =   0
               ExtendLastCol   =   -1  'True
               FormatString    =   $"frmEditUsers.frx":6421
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
            Begin VB.Label Label26 
               Alignment       =   2  'Center
               BackStyle       =   0  'Transparent
               Caption         =   "«·ŒÿÊÿ «·„Õœœ…"
               Height          =   255
               Left            =   480
               RightToLeft     =   -1  'True
               TabIndex        =   144
               Top             =   0
               Width           =   1515
            End
            Begin VB.Label Label27 
               Alignment       =   2  'Center
               BackStyle       =   0  'Transparent
               Caption         =   "Õœœ ŒÿÊÿ «·«‰ «Ã"
               Height          =   255
               Left            =   4440
               RightToLeft     =   -1  'True
               TabIndex        =   143
               Top             =   120
               Width           =   1515
            End
            Begin VB.Label Label28 
               Alignment       =   2  'Center
               BackStyle       =   0  'Transparent
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
               Height          =   255
               Left            =   2910
               RightToLeft     =   -1  'True
               TabIndex        =   142
               Top             =   330
               Width           =   585
            End
            Begin VB.Label Label29 
               Alignment       =   2  'Center
               BackStyle       =   0  'Transparent
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
               Height          =   255
               Left            =   2910
               RightToLeft     =   -1  'True
               TabIndex        =   141
               Top             =   570
               Width           =   585
            End
            Begin VB.Label Label30 
               Alignment       =   2  'Center
               BackStyle       =   0  'Transparent
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
               Height          =   255
               Left            =   2910
               RightToLeft     =   -1  'True
               TabIndex        =   140
               Top             =   930
               Width           =   585
            End
            Begin VB.Label Label31 
               Alignment       =   2  'Center
               BackStyle       =   0  'Transparent
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
               Height          =   255
               Left            =   2910
               RightToLeft     =   -1  'True
               TabIndex        =   139
               Top             =   1170
               Width           =   585
            End
         End
         Begin C1SizerLibCtl.C1Elastic ELe 
            Height          =   1845
            Index           =   8
            Left            =   13560
            TabIndex        =   135
            TabStop         =   0   'False
            Top             =   2940
            Visible         =   0   'False
            Width           =   5400
            _cx             =   9525
            _cy             =   3254
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
            Begin VSFlex8UCtl.VSFlexGrid FgItems 
               Height          =   1830
               Index           =   6
               Left            =   7440
               TabIndex        =   136
               Top             =   165
               Width           =   5400
               _cx             =   9525
               _cy             =   3228
               Appearance      =   2
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
               RowHeightMin    =   300
               RowHeightMax    =   0
               ColWidthMin     =   0
               ColWidthMax     =   0
               ExtendLastCol   =   -1  'True
               FormatString    =   $"frmEditUsers.frx":64E1
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
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "«·Õœ «·«∆ „«‰Ï ··„‰œÊ»"
            Height          =   405
            Index           =   20
            Left            =   3585
            RightToLeft     =   -1  'True
            TabIndex        =   92
            Top             =   9930
            Width           =   1680
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "ÿ»«⁄Â „—œÊœ«  «·„»Ì⁄« "
            Height          =   405
            Index           =   19
            Left            =   14235
            RightToLeft     =   -1  'True
            TabIndex        =   90
            Top             =   9840
            Width           =   1530
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "ÿ»«⁄Â ‰ﬁÿÂ «·»Ì⁄"
            Height          =   285
            Index           =   18
            Left            =   17565
            RightToLeft     =   -1  'True
            TabIndex        =   88
            Top             =   10200
            Width           =   1125
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "ÿ»«⁄Â ›« Ê—… «·»Ì⁄"
            Height          =   285
            Index           =   17
            Left            =   17565
            RightToLeft     =   -1  'True
            TabIndex        =   86
            Top             =   9840
            Width           =   1365
         End
      End
   End
End
Attribute VB_Name = "FrmEditUsers"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim RsSavRec As ADODB.Recordset
Dim BKGrndPic As ClsBackGroundPic
Dim RecId As String
Dim ii As Long
Dim cSearch  As clsDCboSearch
Dim RsTemp As New ADODB.Recordset

Private Sub BtnCancel_Click()
    Unload Me
End Sub
Function CuurentLogdata(Optional Currentmode As String)
     LogTextA = "    ‘«‘… " & ScreenNameArabic & CHR(13) & "ﬂÊœ  «·„” Œœ„ " & TXTCode.text & CHR(13) & "   «”„ «·„” Œœ„  " & XPTxtUserName.text & CHR(13) & "   «·›—⁄ " & DcBranches.text
        LogTexte = "  Screen  " & ScreenNameEnglish & CHR(13) & " User Code " & TXTCode.text & CHR(13) & "   User Name  " & XPTxtUserName.text & CHR(13) & "   Branch " & DcBranches
       If Currentmode <> "D" Then
        AddToLogFile CInt(user_id), , Date, Time, LogTextA, LogTexte, Me.Name, Me.TxtModFlg, , , TXTCode, TXTCode
    Else
        AddToLogFile CInt(user_id), , Date, Time, LogTextA, LogTexte, Me.Name, "D", , , TXTCode, TXTCode
    End If
End Function
Private Sub btnDelete_Click()

    Dim MSGType As Integer
    Dim BlnRecExist As Boolean
    Dim StrMSG  As String
    Dim Msg As String
    
    On Error GoTo ErrTrap

    If DoPremis(Do_Delete, Me.Name, True) = False Then
        Exit Sub
    End If
 If val(TxtVac_ID.text) = 1 Then
MSGType = MsgBox("·«Ì„ﬂ‰   Õ–› Â–« «·”Ã·", vbCritical + vbMsgBoxRtlReading + vbMsgBoxRight, App.Title)
Exit Sub

End If

    If TxtVac_ID.text <> "" Then
    If SystemOptions.UserInterface = ArabicInterface Then
        MSGType = MsgBox("Â·  —€» ›Ì Õ–› Â–« «·”Ã·", vbQuestion + vbMsgBoxRtlReading + vbYesNo + vbMsgBoxRight, App.Title)
     Else
        MSGType = MsgBox("Confirm Delete", vbQuestion + vbMsgBoxRtlReading + vbYesNo + vbMsgBoxRight, App.Title)
   End If
        If MSGType = vbYes Then
           Cn.Execute "Delete from TblUsersStores where userid = " & val(TxtVac_ID.text) & ""
           Cn.Execute "Delete from TblUsersBranches where userid = " & val(TxtVac_ID.text) & ""
           Cn.Execute "Delete from TblUsersBoxes where userid = " & val(TxtVac_ID.text) & ""
           Cn.Execute "Delete from TblUserAccount where UserID = " & val(TxtVac_ID.text) & ""
           Cn.Execute "Delete from TblUsersProductLine where UserID = " & val(TxtVac_ID.text) & ""
            RsSavRec.Find "userid=" & val(TxtVac_ID.text), , adSearchForward, 1
            CuurentLogdata ("D")
            Dim StrSQL As String
            StrSQL = "Delete From TblUsersStores Where userid=" & RsSavRec("UserID").value
            Cn.Execute StrSQL, , adExecuteNoRecords
            
            
            StrSQL = "Delete From TblUsersProductLine Where userid=" & RsSavRec("UserID").value
            Cn.Execute StrSQL, , adExecuteNoRecords
                       
            
            StrSQL = "Delete From TblUsersBranches Where userid=" & RsSavRec("UserID").value
            Cn.Execute StrSQL, , adExecuteNoRecords
            
            Set Me.ImgPic.Picture = Nothing
            RsSavRec.delete
            If SystemOptions.UserInterface = ArabicInterface Then
               MsgBox " „  ⁄„·Ì… «·Õ–› »‰Ã«Õ", vbInformation + vbMsgBoxRtlReading + vbOKOnly + vbMsgBoxRight, App.Title
            Else
               MsgBox "Delete Successfully", vbInformation + vbMsgBoxRtlReading + vbOKOnly + vbMsgBoxRight, App.Title
            End If
            '------------------------------ Move Next ---------------------------.
            FillGridWithData
            BtnNext_Click
        End If
    End If

    Exit Sub
ErrTrap:
    Select Case Err.Number
        Case -2147217873, -2147467259
            'StrMSG = "⁄›Ê« ·« ÌÃÊ“ Õ–› «·”Ã· ·«— »«ÿÂ »»Ì«‰«  √Œ—Ì"
           If SystemOptions.UserInterface = ArabicInterface Then
            StrMSG = "⁄›Ê« ·« ÌÃÊ“ Õ–› «·”Ã· ·«— »«ÿÂ »»Ì«‰«  √Œ—Ì"
            Else
            StrMSG = "Sorry ... This record can not be deleted because it is linked to other data"
            End If
            RsSavRec.CancelUpdate
            MsgBox StrMSG, vbOKOnly + vbMsgBoxRight, App.Title
            'clear the ConnectiOn Errors
            Cn.Errors.Clear
    End Select
End Sub
Private Sub BtnFirst_Click()

    On Error GoTo ErrTrap

    Dim Msg As String

    If Me.TxtModFlg.text = "N" Then
        FindRec val(TxtVac_ID.text)
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
           ' Msg = "⁄›Ê« ·ﬁœ  „ Õ–› Â–« «·”Ã· «· «·Ï" & Chr(13)
           ' Msg = Msg & "„‰ ﬁ»· „” Œœ„ √Œ— ⁄·Ï «·‘»ﬂ… " & Chr(13)
           ' Msg = Msg & "”Ê› Ì „  ÕœÌÀ «·»Ì«‰« "
      If SystemOptions.UserInterface = ArabicInterface Then
            Msg = "⁄›Ê« ·ﬁœ  „ Õ–› Â–« «·”Ã· «· «·Ï" & CHR(13)
            Msg = Msg & "„‰ ﬁ»· „” Œœ„ √Œ— ⁄·Ï «·‘»ﬂ… " & CHR(13)
            Msg = Msg & "”Ê› Ì „  ÕœÌÀ «·»Ì«‰« "
     Else
           Msg = "Sorry..The following record has been deleted" & CHR(13)
           Msg = Msg & "By another user on the network " & CHR(13)
           Msg = Msg & "The data will be updated " & CHR(13)
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
        FindRec val(TxtVac_ID.text)
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
          '  Msg = "⁄›Ê« ·ﬁœ  „ Õ–› Â–« «·”Ã· «· «·Ï" & Chr(13)
          '  Msg = Msg & "„‰ ﬁ»· „” Œœ„ √Œ— ⁄·Ï «·‘»ﬂ… " & Chr(13)
          '  Msg = Msg & "”Ê› Ì „  ÕœÌÀ «·»Ì«‰« "
    If SystemOptions.UserInterface = ArabicInterface Then
            Msg = "⁄›Ê« ·ﬁœ  „ Õ–› Â–« «·”Ã· «· «·Ï" & CHR(13)
            Msg = Msg & "„‰ ﬁ»· „” Œœ„ √Œ— ⁄·Ï «·‘»ﬂ… " & CHR(13)
            Msg = Msg & "”Ê› Ì „  ÕœÌÀ «·»Ì«‰« "
     Else
           Msg = "Sorry..The following record has been deleted" & CHR(13)
           Msg = Msg & "By another user on the network " & CHR(13)
           Msg = Msg & "The data will be updated " & CHR(13)
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

    If TxtVac_ID.text <> "" Then
        TxtModFlg = "E"
        Frm2.Enabled = True
        'Me.TXTDiscounts.SetFocus
        CuurentLogdata
    End If

    Exit Sub
ErrTrap:

    Select Case Err.Number

        Case -2147467259
            'Could not update; currently locked.
          '  Msg = "⁄›Ê«" & Chr(13)
          '  Msg = Msg & " ·«Ì„ﬂ‰  ⁄œÌ· Â–« «·”Ã· ›Ï «·Êﬁ  «·Õ«·Ï" & Chr(13)
          '  Msg = Msg & "ÕÌÀ «‰Â ﬁÌœ «· ⁄œÌ· „‰ ﬁ»· „” Œœ„ «Œ— ⁄·Ï «·‘»ﬂ…"
            If SystemOptions.UserInterface = ArabicInterface Then
            Msg = "⁄›Ê«" & CHR(13)
            Msg = Msg & " ·«Ì„ﬂ‰  ⁄œÌ· Â–« «·”Ã· ›Ï «·Êﬁ  «·Õ«·Ï" & CHR(13)
            Msg = Msg & "ÕÌÀ «‰Â ﬁÌœ «· ⁄œÌ· „‰ ﬁ»· „” Œœ„ «Œ— ⁄·Ï «·‘»ﬂ…"
           Else
            Msg = "Sorry..." & CHR(13)
            Msg = Msg & " This record can not be edited at this time" & CHR(13)
            Msg = Msg & "Because it was modified by another user on the network"
         
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
    Frm2.Enabled = True
    '-----------------------------------
    Me.TxtVac_ID.text = ""
 
    Me.DcBranches.BoundText = ""
    Me.DCEmP.BoundText = ""
    Me.DCJob.BoundText = ""
    Me.DCSalesRepGroups.BoundText = ""
    
    clear_all Me
    FillGridWithData
    CboPriv.ListIndex = 0
    '-----------------------------------
    TxtModFlg.text = "N"

    My_SQL = "TBLSalesRepData"
    rs.Open My_SQL, Cn, adOpenKeyset, adLockOptimistic, adCmdTable

    If rs.RecordCount > 0 Then
        TxtSerial.text = rs.RecordCount + 1
    Else
        TxtSerial.text = 1
    End If

    rs.Close
    CmbType.ListIndex = 0

    ListGroupSelected.Clear
    ListBoxesSelected.Clear
    ListStoreSelected.Clear
    
 
ErrTrap:
End Sub

Private Sub BtnNext_Click()
    On Error GoTo ErrTrap
    Dim Msg As String

    If Me.TxtModFlg.text = "N" Then
        FindRec val(TxtVac_ID.text)
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
            'Msg = "⁄›Ê« ·ﬁœ  „ Õ–› Â–« «·”Ã· «· «·Ï" & Chr(13)
            'Msg = Msg & "„‰ ﬁ»· „” Œœ„ √Œ— ⁄·Ï «·‘»ﬂ… " & Chr(13)
            'Msg = Msg & "”Ê› Ì „  ÕœÌÀ «·»Ì«‰« "
    If SystemOptions.UserInterface = ArabicInterface Then
            Msg = "⁄›Ê« ·ﬁœ  „ Õ–› Â–« «·”Ã· «· «·Ï" & CHR(13)
            Msg = Msg & "„‰ ﬁ»· „” Œœ„ √Œ— ⁄·Ï «·‘»ﬂ… " & CHR(13)
            Msg = Msg & "”Ê› Ì „  ÕœÌÀ «·»Ì«‰« "
           Else
           Msg = "Sorry..The following record has been deleted" & CHR(13)
           Msg = Msg & "By another user on the network " & CHR(13)
           Msg = Msg & "The data will be updated " & CHR(13)
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
        FindRec val(TxtVac_ID.text)
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
          '  Msg = "⁄›Ê« ·ﬁœ  „ Õ–› Â–« «·”Ã· «· «·Ï" & Chr(13)
          '  Msg = Msg & "„‰ ﬁ»· „” Œœ„ √Œ— ⁄·Ï «·‘»ﬂ… " & Chr(13)
          '  Msg = Msg & "”Ê› Ì „  ÕœÌÀ «·»Ì«‰« "
         If SystemOptions.UserInterface = ArabicInterface Then
            Msg = "⁄›Ê« ·ﬁœ  „ Õ–› Â–« «·”Ã· «· «·Ï" & CHR(13)
            Msg = Msg & "„‰ ﬁ»· „” Œœ„ √Œ— ⁄·Ï «·‘»ﬂ… " & CHR(13)
            Msg = Msg & "”Ê› Ì „  ÕœÌÀ «·»Ì«‰« "
           Else
           Msg = "Sorry..The following record has been deleted" & CHR(13)
           Msg = Msg & "By another user on the network " & CHR(13)
           Msg = Msg & "The data will be updated " & CHR(13)
       End If
            MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading + vbOKOnly, App.Title
            RsSavRec.Requery
            Resume BegnieWork
    End Select

End Sub

Private Sub btnQuery_Click()
FrmUserSearch.show
FrmUserSearch.lblSearchtype = 0

End Sub

Private Sub btnSave_Click()
    On Error GoTo ErrTrap
    Dim Msg As String
    Dim StrVacCode As String
    Dim StrVacName As String
    Dim CtrlTxt As Control
    '---------------------- check if data Vaclete -----------------------
    If CboPriv.ListIndex = -1 Then
    
        If SystemOptions.UserInterface = EnglishInterface Then
            Msg = "Specify Privligies"
        Else
              Msg = "Õœœ «·’·«ÕÌ« "
        End If
        
        
        MsgBox Msg, vbOKOnly + vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
        CboPriv.SetFocus
        Exit Sub
    End If
 
    If Trim(DcBranches.BoundText) = "" Then
        If SystemOptions.UserInterface = EnglishInterface Then
            Msg = "Specify Branch"
        Else
            Msg = "Õœœ «·›—⁄ «·«› —«÷Ì  "
        End If

        MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
        DcBranches.SetFocus
        Sendkeys "{F4}"
        Screen.MousePointer = vbDefault
        Exit Sub
    End If
 
    If Trim(Me.DCEmP.BoundText) = "" Then
        If SystemOptions.UserInterface = EnglishInterface Then
            Msg = "Specify Employee"
        Else
            Msg = "Õœœ «·„ÊŸ›    "
        End If

        MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
        DCEmP.SetFocus
        Sendkeys "{F4}"
        Screen.MousePointer = vbDefault
        Exit Sub
    End If
 
    If XPTxtUserName.text = "" Then
    
         If SystemOptions.UserInterface = EnglishInterface Then
            Msg = "Specify User"
        Else
           Msg = "√œŒ· «”„ «·„” Œœ„"
        End If
        
        
        MsgBox Msg, vbOKOnly + vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
        XPTxtUserName.SetFocus
        Exit Sub
    End If

 '   If TxtPassWord.text = "" Then
 '       Msg = "√œŒ· ﬂ·„… «·„—Ê—"
 '       MsgBox Msg, vbOKOnly + vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
 '       TxtPassWord.SetFocus
 '       Exit Sub
 '   End If
 '
 '   If XPTxtPassConfirm.text = "" Then
 '       Msg = "√œŒ·  √ﬂÌœ ﬂ·„… «·„—Ê—"
 '       MsgBox Msg, vbOKOnly + vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
 '       XPTxtPassConfirm.SetFocus
 '       Exit Sub
 '   End If
Dim StrSQL As String
    If StrComp(TxtPassWord.text, XPTxtPassConfirm.text, vbTextCompare) <> 0 Then
            If SystemOptions.UserInterface = EnglishInterface Then
                Msg = "Passwords not matched"
            Else
                Msg = "ﬂ·„… «·„—Ê— Ê √ﬂÌœ ﬂ·„… «·„—Ê— " & CHR(13)
                Msg = Msg + "€Ì— „ ÿ«»ﬁ Ì‰"
             End If
        
        MsgBox Msg, vbOKOnly + vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
        XPTxtPassConfirm.SetFocus
        Exit Sub
    End If

    StrSQL = "select * From TblUsers where UserName='" & Trim(XPTxtUserName.text) & "'" & " and UserID<>" & val(TxtVac_ID.text)
    Set RsTemp = New ADODB.Recordset
    RsTemp.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText

    If Not (RsTemp.EOF Or RsTemp.BOF) Then
    If SystemOptions.UserInterface = EnglishInterface Then
    Msg = "Another user already Exist with the same name"
    Else
        Msg = "ÌÊÃœ „” Œœ„ „”Ã· „”»ﬁ« »Â–« «·«”„" & CHR(13)
        Msg = Msg + "»—Ã«¡ «· √ﬂœ „‰ «·«”„ «·’ÕÌÕ" & CHR(13)
        Msg = Msg + "√Ê  €ÌÌ— √Ê  „ÌÌ“ «”„ «·„” Œœ„"
    End If
        MsgBox Msg, vbOKOnly + vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
        XPTxtUserName.SetFocus
        RsTemp.Close
        Exit Sub
    End If

 
    '------------------------------ check if Empcode exist ----------------------
 
    ' -------------------------------------- txtmodflg type -------------------
    Select Case Me.TxtModFlg.text

            '------------------------------ new record ----------------------------
        Case "N"
            '------------------------- save record -----------------------------
            AddNewRec
            BtnLast_Click

        Case "E"
            '----------------------------- save edit -------------------------------
            'RsEmployee("userid").value = RsSavRec("UserID").value
            StrSQL = "Delete From TblUsersStores Where userid=" & RsSavRec("UserID").value
            Cn.Execute StrSQL, , adExecuteNoRecords
            
            StrSQL = "Delete From TblUsersProductLine Where userid=" & RsSavRec("UserID").value
            Cn.Execute StrSQL, , adExecuteNoRecords
            
            StrSQL = "Delete From TblUsersBranches Where userid=" & RsSavRec("UserID").value
            Cn.Execute StrSQL, , adExecuteNoRecords
            
            
            FiLLRec
    End Select

    Exit Sub
ErrTrap:
If SystemOptions.UserInterface = EnglishInterface Then
MsgBox "error during saving", vbOKOnly + vbMsgBoxRight, App.Title
Else
    MsgBox "Â‰«ﬂ Œÿ√ „« ›Ì ≈œŒ«· «·»Ì«‰« ", vbOKOnly + vbMsgBoxRight, App.Title
End If
End Sub
 
Private Sub BtnUndo_Click()
    FindRec val(TxtVac_ID.text)
    Me.TxtModFlg.text = "R"
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
        Msg = "Number of records before update" & vbCrLf & FristCount & vbCrLf & "Number of records after  update" & vbCrLf & LastCount
        
        If LastCount > FristCount Then
            Msg = Msg + vbCrLf & "Number of new records" & vbCrLf & LastCount - FristCount
        Else
            Msg = Msg + vbCrLf & "Number of records deleted" & vbCrLf & FristCount - LastCount
        End If
    End If
End If
    MsgBox Msg, vbOKOnly + vbMsgBoxRight + vbInformation, App.Title
ErrTrap:
End Sub

Private Sub CmdPic_Click(index As Integer)
On Error GoTo ErrTrap
    Select Case index

        Case 0

            With cdg
               
                .CancelError = False
                .DialogTitle = " ≈Œ Ì«— ’Ê—…"
                'Set The Filter to show pictures only
                .filter = "Bitmap (*.bmp)|*.bmp|JPEG(*.JPG,*.JPEG,*.JPE,*.JFIF)|*.jpg;*.jpeg;*.jpe;*.jfif|"  ' choose formats to include
          
                .ShowOpen

                If .FileName <> "" Then
                    Set Me.ImgPic.Picture = LoadPicture(.FileName)
                End If

            End With

        Case 1
            Set Me.ImgPic.Picture = Nothing
    End Select
    Exit Sub
ErrTrap:
If SystemOptions.UserInterface = ArabicInterface Then
        MsgBox " ÕÃ„ «·’Ê—… €Ì— „œ⁄Ê„", vbCritical
Else
MsgBox " image Size Not Siutable, vbCritical"
End If


End Sub

Private Sub cmdReloadList_Click(index As Integer)
FillMylist CLng(index)
End Sub

Private Sub DBCboClientName_Change()
Dim Fullcode As String
    GetCustomersDetail val(DBCboClientName.BoundText), , Fullcode ', 1, DepitIntervalID, DepitInterval, , creditlocked
    TxtSearchCode.text = Fullcode

End Sub

Private Sub ImgPic_DblClick()
  Load FrmViewPic
    Set FrmViewPic.MainView.Picture = ImgPic.Picture
    FrmViewPic.show vbModal
End Sub

Private Sub Label15_Click()
'    If ListBoxesSelected.ListIndex > -1 Then
'        ListBoxesSelected.RemoveItem ListBoxesSelected.ListIndex
'    End If
'
'
Dim i As Long

For i = 0 To ListBoxesSelected.ListCount - 1
    If i > ListBoxesSelected.ListCount - 1 Then Exit For
    If ListBoxesSelected.Selected(i) Then
        ListBoxesSelected.RemoveItem i
        'ListStoreSelected.ListIndex
        i = i - 1
    End If
Next

End Sub

Private Sub Label16_Click()
    ListBoxesSelected.Clear
End Sub
Private Sub Label17_Click()
    Dim i As Integer
    
    ListBoxesSelected.Clear

    For i = 0 To ListBoxesAll.ListCount - 1
        ListBoxesSelected.AddItem ListBoxesAll.List(i)
        ListBoxesSelected.ItemData(i) = ListBoxesAll.ItemData(i)
    Next i

End Sub

Private Sub Label18_Click()
'    If ListBoxesAll.ListIndex = -1 Then Exit Sub
'    ListBoxesSelected.AddItem ListBoxesAll.List(ListBoxesAll.ListIndex)
'    ListBoxesSelected.ItemData(ListBoxesSelected.NewIndex) = ListBoxesAll.ItemData(ListBoxesAll.ListIndex)

    Dim i As Long
    
    For i = 0 To ListBoxesAll.ListCount - 1
        If ListBoxesAll.Selected(i) Then
            ListBoxesSelected.AddItem ListBoxesAll.List(i)
            ListBoxesSelected.ItemData(ListBoxesSelected.NewIndex) = ListBoxesAll.ItemData(i)
            
        End If
    Next
            

End Sub

Private Sub Label21_Click()
'    If ListAllAccount.ListIndex = -1 Then Exit Sub
'    ListAccountSelect.AddItem ListAllAccount.List(ListAllAccount.ListIndex)
'    ListAccountSelect.ItemData(ListAccountSelect.NewIndex) = ListAllAccount.ItemData(ListAllAccount.ListIndex)
'
    
            If ListStoreall.ListIndex = -1 Then Exit Sub
'    ListStoreSelected.AddItem ListStoreall.List(ListStoreall.ListIndex)
'    ListStoreSelected.ItemData(ListStoreSelected.NewIndex) = ListStoreall.ItemData(ListStoreall.ListIndex)
'
    
Dim i As Long

For i = 0 To ListAllAccount.ListCount - 1
    If ListAllAccount.Selected(i) Then
        ListAccountSelect.AddItem ListAllAccount.List(i)
        ListAccountSelect.ItemData(ListAccountSelect.NewIndex) = ListAllAccount.ItemData(i)
        
    End If
Next

'ItemData (i)

End Sub

Private Sub Label23_Click()
    Dim i As Integer
    ListAccountSelect.Clear
    For i = 0 To ListAllAccount.ListCount - 1
        ListAccountSelect.AddItem ListAllAccount.List(i)
        ListAccountSelect.ItemData(i) = ListAllAccount.ItemData(i)
    Next i
End Sub

Private Sub Label24_Click()
 ListAccountSelect.Clear
End Sub

Private Sub Label25_Click()
 
        

Dim i As Long

For i = 0 To ListAccountSelect.ListCount - 1
    If i > ListAccountSelect.ListCount - 1 Then Exit For
    If ListAccountSelect.Selected(i) Then
        ListAccountSelect.RemoveItem i
        'ListStoreSelected.ListIndex
        i = i - 1
    End If
Next

End Sub

Private Sub ListAllAccount_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyF3 Then
                      Account_search.show
                     Account_search.case_id = 78912

                   End If
End Sub

Private Sub ListBoxesAll_KeyUp(KeyCode As Integer, Shift As Integer)
 If KeyCode = vbKeyF3 Then
        FrmExpensesSearch.Indx = 2
        FrmExpensesSearch.RetrunType = 986
        FrmExpensesSearch.show
    End If
End Sub

Private Sub ListStoreall_KeyUp(KeyCode As Integer, Shift As Integer)
 If KeyCode = vbKeyF3 Then
        FrmStoreSearch.mIndex = 1
        Set FrmStoreSearch.RetrunFrm = Me
        FrmStoreSearch.show
    End If
End Sub

Private Sub TxtSearchCode_KeyPress(KeyAscii As Integer)
    Dim CUSTID As Integer
    
    If KeyAscii = vbKeyReturn Then
        GetCustomersDetail CUSTID, , TxtSearchCode.text, 1
        DBCboClientName.BoundText = CUSTID
    End If
End Sub
Private Sub dcEmp_Change()
    If Me.TxtModFlg = "N" Or Me.TxtModFlg = "E" Then
        If (Me.DCEmP.BoundText) = "" Then Exit Sub
        Me.TXTCode.text = get_EMPLOYEE_Data(val(Me.DCEmP.BoundText), "Fullcode")
        'DCEmp.text = DCEmp.text
    End If
End Sub
Private Sub Dcemp_Click(Area As Integer)
    dcEmp_Change
End Sub
Private Sub DCEmP_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF3 Then
        FrmEmployeeSearch.lbltype = 2911
        Set FrmEmployeeSearch.RetrunFrm = Me
        FrmEmployeeSearch.show
    End If
End Sub
Private Sub Form_Load()

    On Error GoTo ErrTrap
    
    Dim i As Integer
    Dim My_SQL As String
    Dim Dcombos As ClsDataCombos

    My_SQL = "select * from TblUsers where 1=1"
                      If user_id <> 1 Then
     My_SQL = My_SQL & "  and      branchid in(" & Current_branchSql & ")"
        End If
        
        
    Set BKGrndPic = New ClsBackGroundPic
    Set RsSavRec = New ADODB.Recordset
    RsSavRec.CursorLocation = adUseClient

    'RsSavRec.Open My_SQL, Cn, adOpenKeyset, adLockPessimistic, adCmdTableDirect
    RsSavRec.Open My_SQL, Cn, adOpenStatic, adLockOptimistic, adCmdText
    
    Me.TxtModFlg.text = "R"
    Resize_Form Me
    'load tblUsers -----------------------------------------------
    My_SQL = "select UserID,UserName From tblUsers "
    fill_combo DCUser, My_SQL
    Set Dcombos = New ClsDataCombos
    If user_id = 1 Then
    Dcombos.GetBranches Me.DcBranches
    Dcombos.GetEmployees Me.DCEmP
    Dcombos.GetCustomersSuppliers 1, Me.DBCboClientName
    Dcombos.GetCustomersSuppliers 2, Me.DBCboClientName1
    Dcombos.GetStores Me.DCStore
    Dcombos.GetStores Me.DCStore1
    Dcombos.GetStores Me.DCStore3
    Dcombos.GetStores Me.DCStore2
    Dcombos.GetBoxes Me.DCBoxes
    Dcombos.GetBoxes Me.DCBoxes1
    Dcombos.GetBanks Me.Dbanks
Else
    Dcombos.GetBranches Me.DcBranches, True
    Dcombos.GetEmployees Me.DCEmP, , , , True
    Dcombos.GetCustomersSuppliers 1, Me.DBCboClientName
    Dcombos.GetCustomersSuppliers 2, Me.DBCboClientName1
    Dcombos.GetStores Me.DCStore
    Dcombos.GetStores Me.DCStore1
    Dcombos.GetStores Me.DCStore3
    Dcombos.GetStores Me.DCStore2
    Dcombos.GetBoxes Me.DCBoxes
    Dcombos.GetBoxes Me.DCBoxes1
    Dcombos.GetBanks Me.Dbanks

End If

    Set cSearch = New clsDCboSearch
    Set cSearch.Client = Me.DCEmP
    Set cSearch.Client = Me.DcBranches
    Set cSearch.Client = Me.DCStore
    Set cSearch.Client = Me.DCBoxes
    Set cSearch.Client = Me.DCBoxes1
    Set cSearch.Client = Me.Dbanks


    ModFgLib.LinkFgColWithDataCombo Grid, Grid.ColIndex("EmpName"), Me.DCEmP
    ModFgLib.LinkFgColWithDataCombo Grid, Grid.ColIndex("BranchId"), Me.DcBranches
    
    ModFgLib.LinkFgColWithDataCombo Grid, Grid.ColIndex("StoreID"), Me.DCStore
    ModFgLib.LinkFgColWithDataCombo Grid, Grid.ColIndex("BoxID"), Me.DCBoxes
    'ModFgLib.LinkFgColWithDataCombo Grid, Grid.ColIndex("BoxID"), Me.DCBoxes
    ModFgLib.LinkFgColWithDataCombo Grid, Grid.ColIndex("BankID"), Me.Dbanks

    FillGridWithData

    With Me.Grid
        .Cell(flexcpPicture, 0, .ColIndex("DiscountValue")) = Me.GrdImageList.ListImages("Vac_Name").ExtractIcon
        .Cell(flexcpPicture, 0, .ColIndex("Ser")) = Me.GrdImageList.ListImages("Ser").ExtractIcon
        For i = 0 To .Cols - 1
            .Cell(flexcpPictureAlignment, 0, i) = flexPicAlignRightCenter
        Next
   
        .ExtendLastCol = True
        .WallPaper = BKGrndPic.Picture
        .RowHeight(-1) = 300
    End With

    BtnFirst_Click
    ShowTip

    If SystemOptions.UserInterface = EnglishInterface Then
        SetInterface Me
        ChangeLang
    End If

    FillMylist

    If OPEN_NEW_SCREEN = True Then
        btnNew_Click
    End If

ErrTrap:
End Sub
Private Sub ChangeLang()
    Dim XPic As IPictureDisp
    Set XPic = Me.btnFirst.ButtonImage
    Set Me.btnFirst.ButtonImage = Me.btnLast.ButtonImage
    Set Me.btnLast.ButtonImage = XPic
    Set XPic = Me.btnPrevious.ButtonImage
    Set Me.btnPrevious.ButtonImage = Me.btnNext.ButtonImage
    Set Me.btnNext.ButtonImage = XPic
    Label20.Caption = "All Accounts"
    Label19.Caption = "Selected Accounts"
    btnQuery.Caption = "Search"
    ELe(6).Caption = "Electronic Signature"
    Me.Caption = "Users Data"
    Me.Label1(2).Caption = Me.Caption
    Label1(3).Caption = "Code"
    Label1(1).Caption = "Name"
    lbl(4).Caption = "Privligies"
    Label1(9).Caption = "User Name"
    Label1(5).Caption = "Branch"
    Label1(4).Caption = "Box for Sale"
    Label1(14).Caption = "Box for Purchase"
    Label1(5).Caption = "Branch"
    Label1(0).Caption = "Password"
    Label1(8).Caption = "Re. password"
    Label1(7).Caption = "Sale Store"
    Label1(12).Caption = "Store Purchase"
    Label1(10).Caption = "Default Client"
    Label1(11).Caption = "Default Supplier"
    CmdPic(0).Caption = "Add Picture"
    CmdPic(1).Caption = "Delete Picture"
    Label1(6).Caption = "Bank"
    chkNextLogin.Caption = "Change password at login"
    
    chkHidLowering.Caption = "Hide the subtraction of output alerts"
    AllowSelectEmp.Caption = "Select Employee"
'Frame1.Caption = "Selected Boxes"
'Frame2.Caption = "Selected Accounts"
'    Frame11.Caption = "Selected Branch"
'    Frame10.Caption = "Selected Stores"
'    Label11.Caption = "All Branch"
'    Label12.Caption = "Selected Branch"

    Label9.Caption = "All Stores"
    Label10.Caption = "Selected Stores"

    Label2(0).Caption = "Current Record"
    Label2(1).Caption = "NO. Recordes"

    btnNew.Caption = "New"
    btnModify.Caption = "Modify"
    btnSave.Caption = "Save"
    BtnUndo.Caption = "Undo"
    btnDelete.Caption = "Delete"
    btnCancel.Caption = "Exit"
 
    With Me.Grid
        .TextMatrix(0, .ColIndex("ser")) = "Ser"
        .TextMatrix(0, .ColIndex("EmpCode")) = "Code"
        .TextMatrix(0, .ColIndex("EmpName")) = "Emp Name"
        .TextMatrix(0, .ColIndex("JobID")) = "Job"
        .TextMatrix(0, .ColIndex("groupid")) = "Group"
        .TextMatrix(0, .ColIndex("BranchId")) = "Branch"
        .TextMatrix(0, .ColIndex("discountvalue")) = "Discount%"
        .TextMatrix(0, .ColIndex("UserName")) = "UserName"
        .TextMatrix(0, .ColIndex("StoreId")) = "Store Name"
        .TextMatrix(0, .ColIndex("boxId")) = "Box Name"
        .TextMatrix(0, .ColIndex("BankID")) = "Bank"
    End With
    
    '######### khaled was here ############
    isDeactivatedchk.Caption = "Deactivate User"
    Label14.Caption = "All Boxes"
    Label13.Caption = "Selected Boxes"

End Sub

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
    'Set cSearchDCombo = Nothing
    'Set BKGrndPic = Nothing
    Set FrmVacancy = Nothing
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

    Set cSearch = Nothing
ErrTrap:
End Sub
Private Sub Label5_Click()
    If ListGroupSelected.ListIndex > -1 Then
        ListGroupSelected.RemoveItem ListGroupSelected.ListIndex
    End If
End Sub
Private Sub Label6_Click()
    ListGroupSelected.Clear
End Sub
Private Sub Label7_Click()
    Dim i As Integer
    ListGroupSelected.Clear

    For i = 0 To ListGroupAll.ListCount - 1
        ListGroupSelected.AddItem ListGroupAll.List(i)
        ListGroupSelected.ItemData(i) = ListGroupAll.ItemData(i)
    Next i
End Sub
Private Sub Label8_Click()
    If ListGroupAll.ListIndex = -1 Then Exit Sub
    ListGroupSelected.AddItem ListGroupAll.List(ListGroupAll.ListIndex)
    ListGroupSelected.ItemData(ListGroupSelected.NewIndex) = ListGroupAll.ItemData(ListGroupAll.ListIndex)
End Sub
Private Sub LblSelect_Click()
'    If ListStoreall.ListIndex = -1 Then Exit Sub
'    ListStoreSelected.AddItem ListStoreall.List(ListStoreall.ListIndex)
'    ListStoreSelected.ItemData(ListStoreSelected.NewIndex) = ListStoreall.ItemData(ListStoreall.ListIndex)
'
    
        If ListStoreall.ListIndex = -1 Then Exit Sub
'    ListStoreSelected.AddItem ListStoreall.List(ListStoreall.ListIndex)
'    ListStoreSelected.ItemData(ListStoreSelected.NewIndex) = ListStoreall.ItemData(ListStoreall.ListIndex)
'
    
Dim i As Long

For i = 0 To ListStoreall.ListCount - 1
    If ListStoreall.Selected(i) Then
        ListStoreSelected.AddItem ListStoreall.List(i)
        ListStoreSelected.ItemData(ListStoreSelected.NewIndex) = ListStoreall.ItemData(i)
        
    End If
Next

'ItemData (i)
End Sub
Private Sub Label22_Click()
    Dim i As Integer
    ListStoreSelected.Clear
    For i = 0 To ListStoreall.ListCount - 1
        ListStoreSelected.AddItem ListStoreall.List(i)
        ListStoreSelected.ItemData(i) = ListStoreall.ItemData(i)
    Next i
End Sub
Private Sub Label3_Click()
    ListStoreSelected.Clear
End Sub
Private Sub Label4_Click()
'    If ListStoreSelected.ListIndex > -1 Then
'        ListStoreSelected.RemoveItem ListStoreSelected.ListIndex
'    End If
    

Dim i As Long

For i = 0 To ListStoreSelected.ListCount - 1
    If i > ListStoreSelected.ListCount - 1 Then Exit For
    If ListStoreSelected.Selected(i) Then
        ListStoreSelected.RemoveItem i
        'ListStoreSelected.ListIndex
        i = i - 1
    End If
Next
    
    
End Sub
Function createlistString(mylist As ListBox, Optional ByRef Listitems As String)
    Dim i As Integer
    Dim str As String
    str = "0"
    Listitems = ""
    For i = 0 To mylist.ListCount - 1
        str = str & "," & mylist.ItemData(i)
        Listitems = Listitems & "," & mylist.List(i)
    Next i
    createlistString = str
End Function


Function FillMylist(Optional ByVal mIndexd As Long = 0)

    Dim sql As String
    Dim rs As New ADODB.Recordset
    Dim Balance As Double
    Dim i As Integer
    
    If mIndexd = 1 Or mIndexd = 0 Then
        sql = " SELECT * from  TblStore "
    
                       If user_id <> 1 Then
     sql = sql & "  where      branchid in(" & Current_branchSql & ")"
        End If
        
        If SystemOptions.UserInterface = ArabicInterface Then
            sql = sql & " order by  StoreName"
        Else
            sql = sql & " order by  StoreNamee"
        End If

        
        rs.Open sql, Cn, adOpenStatic, adLockReadOnly, adCmdText
    
        ListStoreall.Clear
        'ListStoreSelected.Clear
    
        If rs.RecordCount > 0 Then
            For i = 1 To rs.RecordCount
                If SystemOptions.UserInterface = ArabicInterface Then
                    ListStoreall.AddItem IIf(IsNull(rs("StoreName").value), "", rs("StoreName").value)
                Else
                    ListStoreall.AddItem IIf(IsNull(rs("StoreNamee").value), "", rs("StoreNamee").value)
                End If
    
                ListStoreall.ItemData(ListStoreall.NewIndex) = rs("StoreID").value
                rs.MoveNext
            Next i
        End If
    
        rs.Close
    End If
    If mIndexd = 0 Then
        sql = " SELECT * from  TblBranchesData "
     
           If user_id <> 1 Then
     sql = sql & "  where      branch_id in(" & Current_branchSql & ")"
        End If
        
        
        If SystemOptions.UserInterface = ArabicInterface Then
            sql = sql & " order by  branch_name"
        Else
            sql = sql & " order by  branch_namee"
        End If
     

        rs.Open sql, Cn, adOpenStatic, adLockReadOnly, adCmdText
    
        ListGroupAll.Clear
        'ListGroupSelected.Clear
    
        If rs.RecordCount > 0 Then
            For i = 1 To rs.RecordCount
                If SystemOptions.UserInterface = ArabicInterface Then
                    ListGroupAll.AddItem IIf(IsNull(rs("branch_name").value), "", rs("branch_name").value)
                Else
                    ListGroupAll.AddItem IIf(IsNull(rs("branch_namee").value), "", rs("branch_namee").value)
                End If
    
                ListGroupAll.ItemData(ListGroupAll.NewIndex) = rs("branch_id").value
                rs.MoveNext
            Next i
        End If
        rs.Close
    End If
'    sql = "select* from TblBoxesData where Type = 0 "
If mIndexd = 2 Or mIndexd = 0 Then
        sql = "select* from TblBoxesData    "
                           If user_id <> 1 Then
     sql = sql & "  where      branchid in(" & Current_branchSql & ")"
        End If
        
        ' sql = "select* from TblBoxesData where  "
        If SystemOptions.UserInterface = ArabicInterface Then
            sql = sql & " order by  BoxName"
        Else
            sql = sql & " order by  BoxNameE"
        End If
     
    
        
        rs.Open sql, Cn, adOpenStatic, adLockReadOnly, adCmdText
    
        ListBoxesAll.Clear
        
        If rs.RecordCount > 0 Then
            For i = 1 To rs.RecordCount
                If SystemOptions.UserInterface = ArabicInterface Then
                    ListBoxesAll.AddItem IIf(IsNull(rs("BoxName").value), "", rs("BoxName").value)
                Else
                    ListBoxesAll.AddItem IIf(IsNull(rs("BoxNameE").value), "", rs("BoxNameE").value)
                End If
    
                ListBoxesAll.ItemData(ListBoxesAll.NewIndex) = rs("BoxID").value
                rs.MoveNext
            Next i
        End If
        rs.Close
      ''/////Account
   End If
   If mIndexd = 3 Or mIndexd = 0 Then
        sql = " SELECT * from  ACCOUNTS "
        sql = sql & " where   last_account=0"
        If SystemOptions.UserInterface = ArabicInterface Then
            sql = sql & " order by  Account_Name"
        Else
            sql = sql & " order by  Account_NameEng"
        End If
     
        rs.Open sql, Cn, adOpenStatic, adLockReadOnly, adCmdText
    
        ListAllAccount.Clear
        'ListGroupSelected.Clear
    
        If rs.RecordCount > 0 Then
            For i = 1 To rs.RecordCount
                If SystemOptions.UserInterface = ArabicInterface Then
                    ListAllAccount.AddItem IIf(IsNull(rs("Account_Name").value), "", rs("Account_Name").value)
                Else
                    ListAllAccount.AddItem IIf(IsNull(rs("Account_NameEng").value), "", rs("Account_NameEng").value)
                End If
    
                ListAllAccount.ItemData(ListAllAccount.NewIndex) = rs("Account_ID").value
                rs.MoveNext
            Next i
        End If
        rs.Close
        
    End If
    If mIndexd = 0 Then
      sql = "select * from TblProductLine "
        ' sql = "select* from TblBoxesData where  "
       
        sql = sql & " order by  Name"
        
     
        rs.Open sql, Cn, adOpenStatic, adLockReadOnly, adCmdText
    
        ListProductLineAll.Clear
        
        If rs.RecordCount > 0 Then
            For i = 1 To rs.RecordCount
                ListProductLineAll.AddItem IIf(IsNull(rs("Name").value), "", rs("Name").value)
    
                ListProductLineAll.ItemData(ListProductLineAll.NewIndex) = rs("ID").value
                rs.MoveNext
            Next i
        End If
    End If
End Function
Private Sub Form_Activate()
    Me.ZOrder 0
End Sub
Public Sub AddNewRec()

    On Error GoTo ErrTrap
    
    Dim StrRecID As String
    
    'StrRecID = new_id("TBLSalesRepData", "id", "")
    
    RsSavRec.AddNew
    RsSavRec("UserID").value = CStr(new_id("TblUsers", "UserID", "", True))
    'RsSavRec.Fields("id").value = IIf(StrRecID <> "", StrRecID, Null)
    FiLLRec
ErrTrap:
End Sub
Public Sub FiLLRec()

    On Error GoTo ErrTrap

    RsSavRec.Fields("PassWord").value = IIf((TxtPassWord.text) <> "", (TxtPassWord.text), "")
    RsSavRec.Fields("EmpID").value = IIf(val(Me.DCEmP.BoundText) <> 0, val(Me.DCEmP.BoundText), Null)
    RsSavRec.Fields("BranchId").value = IIf(val(Me.DcBranches.BoundText) <> 0, val(Me.DcBranches.BoundText), Null)
    RsSavRec.Fields("BoxID").value = IIf(val(Me.DCBoxes.BoundText) <> 0, val(Me.DCBoxes.BoundText), Null)
    RsSavRec.Fields("BoxID1").value = IIf(val(Me.DCBoxes1.BoundText) <> 0, val(Me.DCBoxes1.BoundText), Null)
    
    RsSavRec.Fields("Reportname").value = IIf(TXtReportName.text <> "", Trim(TXtReportName.text), Null)
    RsSavRec.Fields("Reportname1").value = IIf(TXtReportName1.text <> "", Trim(TXtReportName1.text), Null)
    RsSavRec.Fields("Reportname2").value = IIf(TXtReportName2.text <> "", Trim(TXtReportName2.text), Null)
    
    RsSavRec.Fields("CreditLimitSalesMan").value = IIf(txtCreditLimitSalesMan.text <> "", Trim(txtCreditLimitSalesMan.text), Null)



    RsSavRec.Fields("BankID").value = IIf(val(Me.Dbanks.BoundText) <> 0, val(Me.Dbanks.BoundText), Null)
    RsSavRec.Fields("StoreID").value = IIf(val(Me.DCStore.BoundText) <> 0, val(Me.DCStore.BoundText), Null)
    RsSavRec.Fields("Custid").value = IIf(val(Me.DBCboClientName.BoundText) <> 0, val(Me.DBCboClientName.BoundText), Null)
    
    RsSavRec.Fields("StoreID1").value = IIf(val(Me.DCStore1.BoundText) <> 0, val(Me.DCStore1.BoundText), Null)
    RsSavRec.Fields("StoreID3").value = IIf(val(Me.DCStore3.BoundText) <> 0, val(Me.DCStore3.BoundText), Null)
    RsSavRec.Fields("StoreID2").value = IIf(val(Me.DCStore2.BoundText) <> 0, val(Me.DCStore2.BoundText), Null)
    RsSavRec.Fields("Custid1").value = IIf(val(Me.DBCboClientName1.BoundText) <> 0, val(Me.DBCboClientName1.BoundText), Null)
    
    RsSavRec("UserName").value = Trim(XPTxtUserName.text)
    If ImgPic.Picture = 0 Then
        RsSavRec("UserSign").value = Null
    Else
        If SavePictureToDB(ImgPic, RsSavRec, "UserSign") = False Then
            GoTo ErrTrap
        End If
    End If

    If Me.CboPriv.ListIndex = 0 Then
        RsSavRec("UserType").value = 2
    Else
        RsSavRec("InvPrices").value = 1
        RsSavRec("InvPrices1").value = 1
        RsSavRec("InvPrices2").value = 1
        
        RsSavRec("ShowInvProfit").value = 1
        RsSavRec("AllowOverMax").value = 1

        RsSavRec("FullPremis").value = 1
        RsSavRec("UserType").value = 0
    End If
 
    RsSavRec("PassConfirm").value = Trim(XPTxtPassConfirm.text)
   
    RsSavRec("IsActive").value = 1
    
 
    If chkNextLogin.value = vbChecked Then
        RsSavRec("ChangePW").value = 1
    Else
        RsSavRec("ChangePW").value = 0
    End If
    
    If chkHidLowering.value = vbChecked Then
        RsSavRec("HidLowering").value = 1
    Else
        RsSavRec("HidLowering").value = 0
    End If
        
        If AllowSelectEmp.value = vbChecked Then
        RsSavRec("AllowSelectEmp").value = 1
    Else
        RsSavRec("AllowSelectEmp").value = 0
    End If
    
    '########## Khaled's was here #################
    If isDeactivatedchk.value = vbChecked Then
        RsSavRec("isDeactivated").value = 1
    Else
        RsSavRec("isDeactivated").value = 0
    End If
    '###############################################
 
    
    'RsSavRec.Fields("JobID").value = IIf(Me.DCJob.BoundText <> 0, Val(Me.DCJob.BoundText), Null)

    RsSavRec.update
    Dim UsrID As Double
   UsrID = IIf(IsNull(RsSavRec("UserID").value), 0, RsSavRec("UserID").value)
    If Me.TxtModFlg.text = "E" Then
    Cn.Execute "Delete from TblUsersStores where userid = " & UsrID & ""
    Cn.Execute "Delete from TblUsersBranches where userid = " & UsrID & ""
    Cn.Execute "Delete from TblUsersBoxes where userid = " & UsrID & ""
    Cn.Execute "Delete from TblUserAccount where UserID = " & UsrID & ""
    Cn.Execute "Delete from TblUsersProductLine where UserID = " & UsrID & ""
    
    End If
    Dim i As Integer
    Dim RsEmployee As New ADODB.Recordset
    
        If ListStoreSelected.ListCount <> 0 Then
        Set RsEmployee = New ADODB.Recordset
            RsEmployee.Open "TblUsersStores", Cn, adOpenKeyset, adLockOptimistic, adCmdTable
            For i = 0 To ListStoreSelected.ListCount - 1
                RsEmployee.AddNew
                RsEmployee("storeId").value = ListStoreSelected.ItemData(i)
                RsEmployee("userid").value = UsrID
                RsEmployee.update
            Next i
            RsEmployee.Close
            Set RsEmployee = Nothing
        End If

        If ListGroupSelected.ListCount <> 0 Then
        Set RsEmployee = New ADODB.Recordset
            RsEmployee.Open "TblUsersBranches", Cn, adOpenKeyset, adLockOptimistic, adCmdTable
                For i = 0 To ListGroupSelected.ListCount - 1
                    RsEmployee.AddNew
                    RsEmployee("BranchID").value = ListGroupSelected.ItemData(i)
                    RsEmployee("userid").value = UsrID
                    RsEmployee.update
            Next i
            RsEmployee.Close
            Set RsEmployee = Nothing
        End If
        
        If ListBoxesSelected.ListCount <> 0 Then
        Set RsEmployee = New ADODB.Recordset
            RsEmployee.Open "TblUsersBoxes", Cn, adOpenKeyset, adLockOptimistic, adCmdTable
                For i = 0 To ListBoxesSelected.ListCount - 1
                    RsEmployee.AddNew
                    RsEmployee("BoxId").value = ListBoxesSelected.ItemData(i)
                    RsEmployee("userid").value = UsrID
                    RsEmployee.update
            Next i
            RsEmployee.Close
            Set RsEmployee = Nothing
        End If
        
    If ListProductLineSelected.ListCount <> 0 Then
        Set RsEmployee = New ADODB.Recordset
            RsEmployee.Open "TblUsersProductLine", Cn, adOpenKeyset, adLockOptimistic, adCmdTable
                For i = 0 To ListProductLineSelected.ListCount - 1
                    RsEmployee.AddNew
                    RsEmployee("ProductLineId").value = ListProductLineSelected.ItemData(i)
                    RsEmployee("userid").value = UsrID
                    'RsEmployee("ShowAlarm").value = FG.ValueMatrix(i, FG.ColIndex("ShowAlarm"))
                    RsEmployee("TypeLine").value = 0
                    
                    RsEmployee.update
            Next i
            RsEmployee.Close
            Set RsEmployee = Nothing
        End If
        Dim sql As String
        
        sql = "Select * from TblUsersProductLine Where  TypeLine = 1 "
        
        saveGrid sql, FG, "ShowAlarm", "", "userId", UsrID, "TypeLine", 1
        
        
         If ListAccountSelect.ListCount <> 0 Then
         sql = "select * from TblUserAccount   where 1=-1"
        Set RsEmployee = New ADODB.Recordset
            RsEmployee.Open sql, Cn, adOpenKeyset, adLockOptimistic, adCmdText
                For i = 0 To ListAccountSelect.ListCount - 1
                    RsEmployee.AddNew
                    RsEmployee("Account_ID").value = ListAccountSelect.ItemData(i)
                    RsEmployee("UserID").value = UsrID
                    RsEmployee.update
            Next i
            RsEmployee.Close
            Set RsEmployee = Nothing
        End If
        
        CuurentLogdata
        
        If SystemOptions.UserInterface = ArabicInterface Then
            MsgBox " „  ⁄„·Ì… «·Õ›Ÿ »‰Ã«Õ", vbInformation + vbOKOnly + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
        Else
            MsgBox "Saved Successfully", vbInformation + vbOKOnly + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
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
    TxtVac_ID.text = IIf(IsNull(RsSavRec.Fields("UserID").value), "", RsSavRec.Fields("UserID").value)
    TxtPassWord.text = IIf(IsNull(RsSavRec.Fields("PassWord").value), "", RsSavRec.Fields("PassWord").value)
    
    TXtReportName.text = IIf(IsNull(RsSavRec.Fields("ReportName").value), "", RsSavRec.Fields("ReportName").value)
    TXtReportName1.text = IIf(IsNull(RsSavRec.Fields("ReportName1").value), "", RsSavRec.Fields("ReportName1").value)
    TXtReportName2.text = IIf(IsNull(RsSavRec.Fields("ReportName2").value), "", RsSavRec.Fields("ReportName2").value)
    
    txtCreditLimitSalesMan.text = IIf(IsNull(RsSavRec.Fields("CreditLimitSalesMan").value), "", RsSavRec.Fields("CreditLimitSalesMan").value)
    
    XPTxtPassConfirm.text = IIf(IsNull(RsSavRec.Fields("PassConfirm").value), "", RsSavRec.Fields("PassConfirm").value)
    XPTxtUserName.text = IIf(IsNull(RsSavRec.Fields("UserName").value), 0, RsSavRec.Fields("UserName").value)
    If Not IsNull(RsSavRec("UserType").value) Then
        If RsSavRec("UserType").value = 2 Then
            CboPriv.ListIndex = 0
        Else
            CboPriv.ListIndex = 1
        End If
    End If
      
    If Not IsNull(RsSavRec("ChangePW").value) Then
        If RsSavRec("ChangePW").value = 0 Then
            chkNextLogin.value = vbUnchecked
        Else
            chkNextLogin.value = vbChecked
        End If
    Else
        chkNextLogin.value = vbUnchecked
    End If
    
   
    If Not IsNull(RsSavRec("HidLowering").value) Then
        If RsSavRec("HidLowering").value = 0 Then
            chkHidLowering.value = vbUnchecked
        Else
            chkHidLowering.value = vbChecked
        End If
    Else
        chkHidLowering.value = vbUnchecked
    End If
    
       If Not IsNull(RsSavRec("AllowSelectEmp").value) Then
        If RsSavRec("AllowSelectEmp").value = 0 Then
            AllowSelectEmp.value = vbUnchecked
        Else
            AllowSelectEmp.value = vbChecked
        End If
    Else
        AllowSelectEmp.value = vbUnchecked
    End If
      
    
    
    '################# khaled was here #####################
    If Not IsNull(RsSavRec("isDeactivated").value) Then
        If RsSavRec("isDeactivated").value = 0 Then
            isDeactivatedchk.value = vbUnchecked
        Else
            isDeactivatedchk.value = vbChecked
        End If
    Else
        isDeactivatedchk.value = vbUnchecked
    End If
    '#######################################################
       
    Me.DCEmP.BoundText = IIf(IsNull(RsSavRec.Fields("EmpID").value), "", RsSavRec.Fields("EmpID").value)
    TXTCode.text = get_EMPLOYEE_Data(val(Me.DCEmP.BoundText), "fullcode")
    Me.DcBranches.BoundText = IIf(IsNull(RsSavRec.Fields("BranchId").value), "", RsSavRec.Fields("BranchId").value)
    Me.DCStore.BoundText = IIf(IsNull(RsSavRec.Fields("StoreID").value), "", RsSavRec.Fields("StoreID").value)
    Me.DCStore1.BoundText = IIf(IsNull(RsSavRec.Fields("StoreID1").value), "", RsSavRec.Fields("StoreID1").value)
    Me.DCStore3.BoundText = IIf(IsNull(RsSavRec.Fields("StoreID3").value), "", RsSavRec.Fields("StoreID3").value)
        Me.DCStore2.BoundText = IIf(IsNull(RsSavRec.Fields("StoreID2").value), "", RsSavRec.Fields("StoreID2").value)
    Me.DCBoxes1.BoundText = IIf(IsNull(RsSavRec.Fields("BoxID1").value), "", RsSavRec.Fields("BoxID1").value)
    Me.DCBoxes.BoundText = IIf(IsNull(RsSavRec.Fields("BoxID").value), "", RsSavRec.Fields("BoxID").value)
    Me.Dbanks.BoundText = IIf(IsNull(RsSavRec.Fields("BankID").value), "", RsSavRec.Fields("BankID").value)
    Me.DBCboClientName.BoundText = IIf(IsNull(RsSavRec.Fields("Custid").value), "", RsSavRec.Fields("Custid").value)
    Me.DBCboClientName1.BoundText = IIf(IsNull(RsSavRec.Fields("Custid1").value), "", RsSavRec.Fields("Custid1").value)
    If Not IsNull(RsSavRec("UserSign").value) Then
        If LenB(RsSavRec("UserSign")) Then
            LoadPictureFromDB ImgPic, RsSavRec, "UserSign"
        Else
            Set ImgPic.Picture = Nothing
        End If
    Else
        Set ImgPic.Picture = Nothing
    End If
    

'********************************************************************
     
    ListStoreSelected.Clear

    Dim RsEmployee As ADODB.Recordset
    Dim StrSQL As String
 
    StrSQL = " SELECT     TOP 100 PERCENT dbo.TblStore.StoreName, dbo.TblStore.StoreNamee, dbo.TblStore.StoreID"
    StrSQL = StrSQL & "  FROM         dbo.TblUsersStores INNER JOIN"
    StrSQL = StrSQL & "   dbo.TblStore ON dbo.TblUsersStores.StoreID = dbo.TblStore.StoreID"
    StrSQL = StrSQL & "  Where (dbo.TblUsersStores.UserID = " & val(TxtVac_ID.text) & ")"
    StrSQL = StrSQL & "  ORDER BY dbo.TblUsersStores.id"

    Set RsEmployee = New ADODB.Recordset

    RsEmployee.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

    If Not (RsEmployee.BOF Or RsEmployee.EOF) Then
        For i = 0 To RsEmployee.RecordCount - 1
            If SystemOptions.UserInterface = ArabicInterface Then
                ListStoreSelected.AddItem IIf(IsNull(RsEmployee("StoreName").value), "", RsEmployee("StoreName").value)
            Else
                ListStoreSelected.AddItem IIf(IsNull(RsEmployee("StoreNameE").value), "", RsEmployee("StoreNameE").value)
            End If
            ListStoreSelected.ItemData(i) = val(IIf(IsNull(RsEmployee("StoreID").value), 0, (RsEmployee("StoreID").value)))
            RsEmployee.MoveNext
        Next i

        RsEmployee.Close
        Set RsEmployee = Nothing
    End If


'*********************************************************************************
     
    ListGroupSelected.Clear

    StrSQL = " SELECT     TOP 100 PERCENT dbo.TblBranchesData.branch_id, dbo.TblBranchesData.branch_name, dbo.TblBranchesData.branch_namee"
    StrSQL = StrSQL & " FROM         dbo.TblUsersBranches INNER JOIN"
    StrSQL = StrSQL & " dbo.TblBranchesData ON dbo.TblUsersBranches.BranchID = dbo.TblBranchesData.branch_id"
    StrSQL = StrSQL & " Where (dbo.TblUsersBranches.UserID = " & val(TxtVac_ID.text) & ")"
    StrSQL = StrSQL & " ORDER BY dbo.TblUsersBranches.id"

    Set RsEmployee = New ADODB.Recordset

    RsEmployee.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

    If Not (RsEmployee.BOF Or RsEmployee.EOF) Then
        For i = 0 To RsEmployee.RecordCount - 1
            If SystemOptions.UserInterface = ArabicInterface Then
                ListGroupSelected.AddItem IIf(IsNull(RsEmployee("branch_name").value), "", RsEmployee("branch_name").value)
            Else
                ListGroupSelected.AddItem IIf(IsNull(RsEmployee("branch_nameE").value), "", RsEmployee("branch_nameE").value)
            End If
            ListGroupSelected.ItemData(i) = val(IIf(IsNull(RsEmployee("branch_id").value), 0, (RsEmployee("branch_id").value)))
                
            RsEmployee.MoveNext
        Next i

        RsEmployee.Close
        Set RsEmployee = Nothing
    End If
'*********************************************************************************
    ListBoxesSelected.Clear
    
    StrSQL = "SELECT TblUsersBoxes.id, TblBoxesData.BoxName, TblUsersBoxes.BoxId, TblUsersBoxes.userid, TblBoxesData.BoxNameE"
    StrSQL = StrSQL & " FROM TblUsersBoxes INNER JOIN"
    StrSQL = StrSQL & " TblBoxesData ON TblUsersBoxes.BoxId = TblBoxesData.BoxID"
    StrSQL = StrSQL & " Where (TblUsersBoxes.UserID = " & val(TxtVac_ID.text) & ")"

    Set RsEmployee = New ADODB.Recordset

    RsEmployee.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

    If Not (RsEmployee.BOF Or RsEmployee.EOF) Then
        For i = 0 To RsEmployee.RecordCount - 1
            If SystemOptions.UserInterface = ArabicInterface Then
                ListBoxesSelected.AddItem IIf(IsNull(RsEmployee("BoxName").value), "", RsEmployee("BoxName").value)
            Else
                ListBoxesSelected.AddItem IIf(IsNull(RsEmployee("BoxNameE").value), "", RsEmployee("BoxNameE").value)
            End If
            ListBoxesSelected.ItemData(i) = val(IIf(IsNull(RsEmployee("BoxId").value), 0, (RsEmployee("BoxId").value)))
                
            RsEmployee.MoveNext
        Next i

        RsEmployee.Close
        Set RsEmployee = Nothing
    End If
    
  ListProductLineSelected.Clear
    
    StrSQL = "SELECT TblUsersProductLine.id,TblUsersProductLine.ShowAlarm, TblProductLine.Name, TblUsersProductLine.ProductLineId, TblUsersProductLine.userid "
    StrSQL = StrSQL & " FROM TblUsersProductLine INNER JOIN"
    StrSQL = StrSQL & " TblProductLine ON TblUsersProductLine.ProductLineId = TblProductLine.ID"
    StrSQL = StrSQL & " Where (TblUsersProductLine.UserID = " & val(TxtVac_ID.text) & ") and IsNull( TypeLine,0) = 0"

    Set RsEmployee = New ADODB.Recordset

    RsEmployee.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText
    
    StrSQL = "SELECT TblProductLine.id,TblProductLine.id as ProductLineID ,TblUsersProductLine.ShowAlarm, TblProductLine.Name,  TblUsersProductLine.userid "
    StrSQL = StrSQL & " FROM TblUsersProductLine RIGHT outer JOIN "
    StrSQL = StrSQL & " TblProductLine ON TblUsersProductLine.ProductLineId = TblProductLine.ID and (TblUsersProductLine.UserID = " & val(TxtVac_ID.text) & ") and TypeLine = 1"
    
    loadgrid StrSQL, FG, True, False
    
    If Not (RsEmployee.BOF Or RsEmployee.EOF) Then
        For i = 0 To RsEmployee.RecordCount - 1
            
                ListProductLineSelected.AddItem IIf(IsNull(RsEmployee("Name").value), "", RsEmployee("Name").value)
            
            ListProductLineSelected.ItemData(i) = val(IIf(IsNull(RsEmployee("ProductLineId").value), 0, (RsEmployee("ProductLineId").value)))
                
            RsEmployee.MoveNext
        Next i

        RsEmployee.Close
        Set RsEmployee = Nothing
    End If
        
'''//////////////
    ListAccountSelect.Clear
    
    StrSQL = " SELECT     dbo.TblUserAccount.UserID, dbo.TblUserAccount.Account_ID, dbo.ACCOUNTS.Account_Name, dbo.ACCOUNTS.Account_NameEng"
    StrSQL = StrSQL & " FROM         dbo.TblUserAccount LEFT OUTER JOIN"
    StrSQL = StrSQL & "                  dbo.ACCOUNTS ON dbo.TblUserAccount.Account_ID = dbo.ACCOUNTS.Account_ID"
    StrSQL = StrSQL & "     Where (dbo.TblUserAccount.UserID = " & val(TxtVac_ID.text) & ")"
    Set RsEmployee = New ADODB.Recordset

    RsEmployee.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

    If Not (RsEmployee.BOF Or RsEmployee.EOF) Then
        For i = 0 To RsEmployee.RecordCount - 1
            If SystemOptions.UserInterface = ArabicInterface Then
                ListAccountSelect.AddItem IIf(IsNull(RsEmployee("Account_Name").value), "", RsEmployee("Account_Name").value)
            Else
                ListAccountSelect.AddItem IIf(IsNull(RsEmployee("Account_NameEng").value), "", RsEmployee("Account_NameEng").value)
            End If
            ListAccountSelect.ItemData(i) = val(IIf(IsNull(RsEmployee("Account_ID").value), 0, (RsEmployee("Account_ID").value)))
                
            RsEmployee.MoveNext
        Next i

        RsEmployee.Close
        Set RsEmployee = Nothing
    End If
'*********************************************************************************
    'Me.DCJob.BoundText = IIf(IsNull(RsSavRec.Fields("JobID").value), "", RsSavRec.Fields("JobID").value)

    LabCurrRec.Caption = RsSavRec.AbsolutePosition
    LabCountRec.Caption = RsSavRec.RecordCount
    With Grid
        For i = 1 To .rows - 1
            If Trim(TxtVac_ID.text) = .TextMatrix(i, .ColIndex("UserID")) Then
                TxtSerial.text = .TextMatrix(i, .ColIndex("Ser"))
                .row = i
                Exit Sub
            End If
        Next
    End With
ErrTrap:

End Sub
Public Sub EditRec(StrTable As String, RecId As String)
    'My_SQL = "select * From " & StrTable & " where "
    'RsSavRec.Open My_SQL, cn, adOpenStatic, adLockReadOnly, adCmdTableDirect
    FiLLRec
End Sub

Private Sub Grid_EnterCell()
    On Error GoTo ErrTrap
    FindRec val(Me.Grid.TextMatrix(Me.Grid.row, Me.Grid.ColIndex("UserID")))
ErrTrap:
End Sub

Private Sub txtCode_KeyPress(KeyAscii As Integer)
If Me.TxtModFlg = "N" Or Me.TxtModFlg = "E" Then
    DCEmP.BoundText = GeTEmpIDByEmpCode(TXTCode.text)
End If
End Sub
Private Sub TxtPassWord_DblClick()
'     If user_id = 1 Then
'     MsgBox txtPassword
'     End If
End Sub
Private Sub TxtSearchCode1_KeyPress(KeyAscii As Integer)
    Dim CUSTID As Integer
    If KeyAscii = vbKeyReturn Then
        GetCustomersDetail CUSTID, , TxtSearchCode.text, 1
        DBCboClientName.BoundText = CUSTID
    End If
End Sub
Private Sub TxtVac_ID_Change()
    Dim TxtMod As String
    TxtMod = TxtModFlg.text
    TxtModFlg.text = ""
    TxtModFlg = TxtMod
End Sub
Public Function FindRec(ByVal RecId As Long)
    On Error GoTo ErrTrap
    RsSavRec.Find "UserID=" & RecId, , adSearchForward, 1

    If Not (RsSavRec.EOF) Then
        FiLLTXT
    End If

    Exit Function
ErrTrap:

    If RsSavRec.EditMode <> adEditNone Then
        RsSavRec.CancelUpdate
        BtnUndo_Click
    End If

    'RsSavRec.Filter = adFilterNone
End Function

'Private Sub TxtVacCode_KeyPress(KeyAscii As Integer)
'KeyAscii = DataFormat(ChrOnly, KeyAscii)
'End Sub

Private Sub TxtModFlg_Change()

    If TxtModFlg.text = "N" Then
        Frm2.Enabled = True
        Me.btnNew.Enabled = False
        btnModify.Enabled = False
        btnDelete.Enabled = False
        Me.btnQuery.Enabled = False
        Grid.Enabled = False
        BtnUndo.Enabled = True
        Me.btnSave.Enabled = True
        BtnUpdate.Enabled = False
        'btnNext.Enabled = False
        'btnPrevious.Enabled = False
        'btnFirst.Enabled = False
        'btnLast.Enabled = False
        ListGroupAll.Enabled = True
        ListStoreall.Enabled = True
        ListBoxesAll.Enabled = True
        ListAllAccount.Enabled = True
        ListProductLineAll.Enabled = True
    ElseIf TxtModFlg.text = "R" Then
        ListAllAccount.Enabled = False
        ListProductLineAll.Enabled = False
        Frm2.Enabled = False
        Grid.Enabled = True
        btnModify.Enabled = False
        btnDelete.Enabled = False
        If TxtVac_ID.text <> "" Then
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
        ListGroupAll.Enabled = False
        ListStoreall.Enabled = False
        ListBoxesAll.Enabled = False
    ElseIf TxtModFlg.text = "E" Then
        ListAllAccount.Enabled = True
                ListProductLineAll.Enabled = True
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
        ListGroupAll.Enabled = True
        ListStoreall.Enabled = True
        ListBoxesAll.Enabled = True
    End If

End Sub
Public Sub FillGridWithData()

    On Error GoTo ErrTrap

    Dim i As Integer
    Dim rs As ADODB.Recordset
    Dim My_SQL As String

    Set rs = New ADODB.Recordset
    My_SQL = "select * From TblUsers order by userid"
    rs.Open My_SQL, Cn, adOpenKeyset, adLockReadOnly, adCmdText

    With Me.Grid
        .rows = 2
        .Clear flexClearScrollable
        If rs.RecordCount > 0 Then
            .rows = rs.RecordCount + 1
            rs.MoveFirst
            For i = 1 To .rows - 1
                .TextMatrix(i, .ColIndex("Ser")) = i
                .TextMatrix(i, .ColIndex("UserID")) = IIf(IsNull(rs.Fields("UserID").value), "", rs.Fields("UserID").value)
                .TextMatrix(i, .ColIndex("EmpID")) = IIf(IsNull(rs.Fields("EmpID").value), "", rs.Fields("EmpID").value)
                .TextMatrix(i, .ColIndex("EmpCode")) = get_EMPLOYEE_Data(val(.TextMatrix(i, .ColIndex("EmpID"))), "fullcode")
                .TextMatrix(i, .ColIndex("UserName")) = IIf(IsNull(rs.Fields("UserName").value), "", rs.Fields("UserName").value)
                .TextMatrix(i, .ColIndex("EmpName")) = IIf(IsNull(rs.Fields("EmpID").value), "", rs.Fields("EmpID").value)
                .TextMatrix(i, .ColIndex("BranchId")) = IIf(IsNull(rs.Fields("BranchId").value), "", rs.Fields("BranchId").value)
                .TextMatrix(i, .ColIndex("StoreID")) = IIf(IsNull(rs.Fields("StoreID").value), "", rs.Fields("StoreID").value)
                .TextMatrix(i, .ColIndex("BoxID")) = IIf(IsNull(rs.Fields("BoxID").value), "", rs.Fields("BoxID").value)
                .TextMatrix(i, .ColIndex("BankID")) = IIf(IsNull(rs.Fields("BankID").value), "", rs.Fields("BankID").value)
                rs.MoveNext
            Next i
            rs.Close
        End If
        .AutoSize 0, .Cols - 1, False
        .RowHeight(-1) = 300
    End With

ErrTrap:
End Sub
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
Private Sub Form_KeyDown(KeyCode As Integer, _
                         Shift As Integer)
    On Error GoTo ErrTrap

    If KeyCode = vbKeyReturn Then
        If Me.TxtModFlg.text = "R" Then
            btnNew_Click
        Else
            Sendkeys "{TAB}"
        End If
    End If
    'New -------------------------
    If KeyCode = vbKeyF12 Then
        If btnNew.Enabled = False Then Exit Sub
        btnNew_Click
    End If
    'Edit ------------------------
    If KeyCode = vbKeyF11 Then
        If btnModify.Enabled = False Then Exit Sub
        btnModify_Click
    End If
    'save ------------------------
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

    Exit Sub
ErrTrap:
End Sub
Private Function CheckDelCountry(Lngid As Long) As Boolean
    'Dim Rs As ADODB.Recordset
    'Dim StrSQL As String
    'StrSQL = "Select * From TblEmployee Where GovernmentID=" & Lngid & ""
    'Set Rs = New ADODB.Recordset
    'Rs.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText
    'If Not (Rs.BOF Or Rs.EOF) Then
    '    CheckDelCountry = False
    'Else
    '    CheckDelCountry = True
    'End If
    'Rs.Close
    'Set Rs = Nothing
End Function

Private Sub Label28_Click()
    If ListProductLineAll.ListIndex = -1 Then Exit Sub
    ListProductLineSelected.AddItem ListProductLineAll.List(ListProductLineAll.ListIndex)
    ListProductLineSelected.ItemData(ListProductLineSelected.NewIndex) = ListProductLineAll.ItemData(ListProductLineAll.ListIndex)
'    FG.Rows = ListProductLineSelected.ListCount + 1
'    FG.TextMatrix(FG.Rows - 1, FG.ColIndex("Name")) = ListProductLineAll.List(ListProductLineAll.ListIndex)
'    FG.TextMatrix(FG.Rows - 1, FG.ColIndex("ProductLineID")) = ListProductLineAll.ItemData(ListProductLineAll.ListIndex)
End Sub

Private Sub Label29_Click()
    Dim i As Integer
    ListProductLineSelected.Clear
'    FG.Rows = 1
'    FG.Rows = ListProductLineSelected.ListCount + 1
    For i = 0 To ListProductLineAll.ListCount - 1
        ListProductLineSelected.AddItem ListProductLineAll.List(i)
        ListProductLineSelected.ItemData(i) = ListProductLineAll.ItemData(i)
'        FG.TextMatrix(i + 1, FG.ColIndex("Name")) = ListProductLineAll.List(ListProductLineAll.ListIndex)
'        FG.TextMatrix(i + 1, FG.ColIndex("ProductLineID")) = ListProductLineAll.ItemData(ListProductLineAll.ListIndex)
        
    Next i

End Sub

Private Sub Label30_Click()
 ListProductLineSelected.Clear
' FG.Rows = 1
End Sub

Private Sub Label31_Click()
    If ListProductLineSelected.ListIndex > -1 Then
      ListProductLineSelected.RemoveItem ListProductLineSelected.ListIndex
        'FG.RemoveItem
    End If

End Sub





