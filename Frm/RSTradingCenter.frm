VERSION 5.00
Object = "{C115893A-A3BF-43AF-B28D-69DB846077F3}#1.0#0"; "vsflex8u.ocx"
Object = "{0AFE7BE0-11B7-4A3E-978D-D4501E9A57FE}#1.0#0"; "c1sizer.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{FE5DCFAD-BC1D-11D2-94CF-004005455FAA}#1.4#0"; "ImpulseButton.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#12.0#0"; "SuiteCtrls.ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{49003D3A-66CD-11D7-A449-E937BE2D9041}#1.0#0"; "ALLBUTTONS.ocx"
Begin VB.Form RSAkar 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "»Ì«‰«  «·⁄ﬁ«—« "
   ClientHeight    =   9450
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   18225
   Icon            =   "RSTradingCenter.frx":0000
   LinkTopic       =   "Form2"
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   RightToLeft     =   -1  'True
   ScaleHeight     =   9450
   ScaleWidth      =   18225
   WindowState     =   2  'Maximized
   Begin VB.Frame frm2 
      Caption         =   "Frame5"
      Height          =   375
      Left            =   12720
      RightToLeft     =   -1  'True
      TabIndex        =   0
      Top             =   120
      Visible         =   0   'False
      Width           =   375
   End
   Begin C1SizerLibCtl.C1Elastic ELe 
      Height          =   9450
      Index           =   15
      Left            =   0
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   0
      Width           =   18225
      _cx             =   32147
      _cy             =   16669
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
         Height          =   555
         Left            =   0
         RightToLeft     =   -1  'True
         TabIndex        =   4
         Top             =   0
         Width           =   18285
         Begin VB.Frame Frmo2 
            BackColor       =   &H00E0E0E0&
            BorderStyle     =   0  'None
            Enabled         =   0   'False
            Height          =   375
            Left            =   2460
            RightToLeft     =   -1  'True
            TabIndex        =   7
            Top             =   90
            Visible         =   0   'False
            Width           =   3105
            Begin MSDataListLib.DataCombo DCUser 
               CausesValidation=   0   'False
               Height          =   315
               Left            =   -255
               TabIndex        =   8
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
               TabIndex        =   9
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
            Left            =   8820
            RightToLeft     =   -1  'True
            TabIndex        =   6
            Text            =   "modflag"
            Top             =   90
            Visible         =   0   'False
            Width           =   465
         End
         Begin VB.TextBox TxtVac_ID 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H000000FF&
            Height          =   240
            Left            =   7350
            RightToLeft     =   -1  'True
            TabIndex        =   5
            Top             =   150
            Visible         =   0   'False
            Width           =   945
         End
         Begin MSComctlLib.ImageList GrdImageList 
            Left            =   10200
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
                  Picture         =   "RSTradingCenter.frx":57E2
                  Key             =   "CompanyName"
               EndProperty
               BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "RSTradingCenter.frx":5B7C
                  Key             =   "Ser"
               EndProperty
               BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "RSTradingCenter.frx":5F16
                  Key             =   "Vac_Name"
               EndProperty
               BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "RSTradingCenter.frx":62B0
                  Key             =   "ShareCount"
               EndProperty
               BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "RSTradingCenter.frx":664A
                  Key             =   "Dis_Count"
               EndProperty
               BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "RSTradingCenter.frx":69E4
                  Key             =   "Bouns"
               EndProperty
               BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "RSTradingCenter.frx":6D7E
                  Key             =   "SharesValue"
               EndProperty
               BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "RSTradingCenter.frx":7318
                  Key             =   "BuyValue"
               EndProperty
            EndProperty
         End
         Begin ImpulseButton.ISButton btnLast 
            Height          =   315
            Left            =   90
            TabIndex        =   10
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
            ButtonImage     =   "RSTradingCenter.frx":76B2
            ColorButton     =   16777215
            AcclimateGrayTones=   -1  'True
            DrawFocusRectangle=   0   'False
            DisabledImageExtraction=   0
            DisabledImageStyle=   1
         End
         Begin ImpulseButton.ISButton btnNext 
            Height          =   315
            Left            =   555
            TabIndex        =   11
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
            ButtonImage     =   "RSTradingCenter.frx":7A4C
            ColorButton     =   16777215
            DrawFocusRectangle=   0   'False
            DisabledImageStyle=   1
         End
         Begin ImpulseButton.ISButton btnPrevious 
            Height          =   315
            Left            =   1155
            TabIndex        =   12
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
            ButtonImage     =   "RSTradingCenter.frx":7DE6
            ColorButton     =   16777215
            DrawFocusRectangle=   0   'False
            DisabledImageStyle=   1
         End
         Begin ImpulseButton.ISButton btnFirst 
            Height          =   315
            Left            =   1620
            TabIndex        =   13
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
            ButtonImage     =   "RSTradingCenter.frx":8180
            ColorButton     =   16777215
            DrawFocusRectangle=   0   'False
            DisabledImageStyle=   1
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "»Ì«‰«  «·⁄ﬁ«—« "
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   18
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   615
            Index           =   2
            Left            =   14655
            RightToLeft     =   -1  'True
            TabIndex        =   15
            Top             =   90
            Width           =   3270
         End
         Begin VB.Label lblcaption 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   18
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Left            =   1920
            RightToLeft     =   -1  'True
            TabIndex        =   14
            Top             =   90
            Width           =   13695
         End
         Begin VB.Image ImgFavorites 
            Height          =   390
            Left            =   5760
            Picture         =   "RSTradingCenter.frx":851A
            Stretch         =   -1  'True
            Top             =   0
            Width           =   525
         End
      End
      Begin C1SizerLibCtl.C1Tab XPTab301 
         Height          =   8250
         Left            =   0
         TabIndex        =   16
         Top             =   600
         Width           =   18300
         _cx             =   32279
         _cy             =   14552
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
         ForeColor       =   0
         FrontTabColor   =   14871017
         BackTabColor    =   12648447
         TabOutlineColor =   -2147483632
         FrontTabForeColor=   16711680
         Caption         =   "»Ì«‰«  «·⁄ﬁ«—« |»Ì«‰«   ›’Ì·ÌÂ ··⁄ﬁ«—|»Ì«‰«  «·„·ﬂÌ…|„—›ﬁ«  |«· Œ›Ì÷"
         Align           =   0
         CurrTab         =   4
         FirstTab        =   0
         Style           =   3
         Position        =   0
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
         Picture(0)      =   "RSTradingCenter.frx":C182
         Begin VB.TextBox TxtNoteID 
            Alignment       =   1  'Right Justify
            Height          =   7785
            Left            =   18945
            RightToLeft     =   -1  'True
            TabIndex        =   318
            Top             =   420
            Visible         =   0   'False
            Width           =   18210
         End
         Begin C1SizerLibCtl.C1Elastic ELe 
            Height          =   7785
            Index           =   1
            Left            =   -19755
            TabIndex        =   17
            TabStop         =   0   'False
            Top             =   420
            Width           =   18210
            _cx             =   32120
            _cy             =   13732
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
            AutoSizeChildren=   7
            BorderWidth     =   6
            ChildSpacing    =   4
            Splitter        =   0   'False
            FloodDirection  =   0
            FloodPercent    =   0
            CaptionPos      =   0
            WordWrap        =   -1  'True
            MaxChildSize    =   0
            MinChildSize    =   0
            TagWidth        =   0
            TagPosition     =   2
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
            Begin VB.TextBox TxtAccoup 
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
               Left            =   120
               MaxLength       =   50
               RightToLeft     =   -1  'True
               TabIndex        =   230
               Top             =   960
               Width           =   2130
            End
            Begin VB.TextBox TxAssetscode 
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
               Left            =   8445
               MaxLength       =   50
               RightToLeft     =   -1  'True
               TabIndex        =   228
               Top             =   960
               Width           =   1170
            End
            Begin VB.TextBox TxtCodeSales 
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
               Left            =   3735
               MaxLength       =   50
               RightToLeft     =   -1  'True
               TabIndex        =   40
               Top             =   600
               Width           =   1050
            End
            Begin VB.TextBox txtstreetname 
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
               Left            =   7605
               MaxLength       =   50
               RightToLeft     =   -1  'True
               TabIndex        =   39
               Tag             =   "«œŒ· «”„ «·‘«—⁄"
               Top             =   1665
               Width           =   2025
            End
            Begin VB.TextBox TxtAqarid 
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
               Left            =   8205
               MaxLength       =   50
               RightToLeft     =   -1  'True
               TabIndex        =   38
               Top             =   240
               Width           =   1425
            End
            Begin VB.TextBox txtaqarname 
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
               Left            =   5670
               MaxLength       =   50
               RightToLeft     =   -1  'True
               TabIndex        =   37
               Tag             =   "⁄›Ê« Ì—ÃÏ «œŒ«· √”„ «·⁄ﬁ«—"
               Top             =   600
               Width           =   3945
            End
            Begin VB.TextBox txtaqarNo 
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
               Left            =   6270
               MaxLength       =   50
               RightToLeft     =   -1  'True
               TabIndex        =   36
               Top             =   240
               Width           =   1185
            End
            Begin VB.TextBox txtaqarage 
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
               Left            =   7605
               MaxLength       =   50
               RightToLeft     =   -1  'True
               TabIndex        =   35
               Top             =   2025
               Width           =   2025
            End
            Begin VB.ComboBox dcmaintenancetypeid 
               Height          =   315
               ItemData        =   "RSTradingCenter.frx":C51C
               Left            =   7605
               List            =   "RSTradingCenter.frx":C526
               RightToLeft     =   -1  'True
               TabIndex        =   34
               Top             =   2385
               Width           =   2025
            End
            Begin VB.TextBox txtEntryCount 
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
               Left            =   120
               MaxLength       =   50
               RightToLeft     =   -1  'True
               TabIndex        =   33
               Top             =   2385
               Width           =   2160
            End
            Begin VB.ComboBox cbostatusid 
               Height          =   315
               ItemData        =   "RSTradingCenter.frx":C53E
               Left            =   5190
               List            =   "RSTradingCenter.frx":C54E
               RightToLeft     =   -1  'True
               TabIndex        =   32
               Top             =   2385
               Width           =   855
            End
            Begin VB.TextBox txtcurrentPrice 
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
               Left            =   3855
               MaxLength       =   50
               RightToLeft     =   -1  'True
               TabIndex        =   31
               Top             =   2025
               Width           =   2160
            End
            Begin VB.TextBox txtlastrentvalue 
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
               Left            =   120
               MaxLength       =   50
               RightToLeft     =   -1  'True
               TabIndex        =   30
               Top             =   2025
               Width           =   2160
            End
            Begin VB.TextBox TxtPrice 
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
               Left            =   120
               MaxLength       =   50
               RightToLeft     =   -1  'True
               TabIndex        =   29
               Top             =   3465
               Width           =   2160
            End
            Begin VB.TextBox TxtRat 
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
               Left            =   3855
               MaxLength       =   50
               RightToLeft     =   -1  'True
               TabIndex        =   28
               Top             =   3465
               Width           =   2160
            End
            Begin VB.TextBox txttotallength 
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
               Left            =   3855
               MaxLength       =   50
               RightToLeft     =   -1  'True
               TabIndex        =   27
               Top             =   3105
               Width           =   2160
            End
            Begin VB.TextBox txtnoofparking 
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
               Left            =   120
               MaxLength       =   50
               RightToLeft     =   -1  'True
               TabIndex        =   26
               Top             =   2745
               Width           =   2160
            End
            Begin VB.TextBox txtmeterRentvalue 
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
               Left            =   120
               MaxLength       =   50
               RightToLeft     =   -1  'True
               TabIndex        =   25
               Top             =   3105
               Width           =   2160
            End
            Begin VB.TextBox txtfloorcount 
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
               Left            =   7605
               MaxLength       =   50
               RightToLeft     =   -1  'True
               TabIndex        =   24
               Top             =   2745
               Width           =   2025
            End
            Begin VB.ComboBox cbointerfaceid 
               Height          =   315
               ItemData        =   "RSTradingCenter.frx":C576
               Left            =   7605
               List            =   "RSTradingCenter.frx":C580
               RightToLeft     =   -1  'True
               TabIndex        =   23
               Top             =   3105
               Width           =   2025
            End
            Begin VB.TextBox txtnoofapartement 
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
               Left            =   7605
               MaxLength       =   50
               RightToLeft     =   -1  'True
               TabIndex        =   22
               Top             =   3465
               Width           =   2025
            End
            Begin VB.TextBox txtnoofoffices 
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
               Left            =   3855
               MaxLength       =   50
               RightToLeft     =   -1  'True
               TabIndex        =   21
               Top             =   2745
               Width           =   2160
            End
            Begin VB.CommandButton Command3 
               Caption         =   "⁄—÷"
               Height          =   315
               Left            =   120
               RightToLeft     =   -1  'True
               TabIndex        =   20
               Top             =   7035
               Width           =   825
            End
            Begin VB.TextBox txtgooglemap 
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
               Left            =   1080
               MaxLength       =   50
               RightToLeft     =   -1  'True
               TabIndex        =   19
               Top             =   7035
               Width           =   4095
            End
            Begin VB.TextBox txtlocation 
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
               Height          =   780
               Left            =   120
               MaxLength       =   50
               MultiLine       =   -1  'True
               RightToLeft     =   -1  'True
               ScrollBars      =   2  'Vertical
               TabIndex        =   18
               Top             =   6090
               Width           =   5055
            End
            Begin VSFlex8UCtl.VSFlexGrid Grid 
               Height          =   1965
               Left            =   0
               TabIndex        =   41
               Top             =   -3585
               Width           =   11985
               _cx             =   21140
               _cy             =   3466
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
               Cols            =   6
               FixedRows       =   1
               FixedCols       =   1
               RowHeightMin    =   320
               RowHeightMax    =   0
               ColWidthMin     =   0
               ColWidthMax     =   0
               ExtendLastCol   =   -1  'True
               FormatString    =   $"RSTradingCenter.frx":C590
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
            Begin MSDataListLib.DataCombo DcboCountryID2 
               Height          =   315
               Left            =   7605
               TabIndex        =   42
               Tag             =   "⁄›Ê« Ì—ÃÏ «œŒ«· √”„ «·œÊ·…"
               Top             =   1320
               Width           =   2025
               _ExtentX        =   3572
               _ExtentY        =   556
               _Version        =   393216
               Text            =   ""
               RightToLeft     =   -1  'True
            End
            Begin MSDataListLib.DataCombo DcboGovernmentID 
               Height          =   315
               Left            =   3765
               TabIndex        =   43
               Tag             =   "⁄›Ê« Ì—ÃÏ «œŒ«· √”„ «·„œÌ‰…"
               Top             =   1320
               Width           =   2250
               _ExtentX        =   3969
               _ExtentY        =   556
               _Version        =   393216
               Text            =   ""
               RightToLeft     =   -1  'True
            End
            Begin MSDataListLib.DataCombo DcboCityID 
               Height          =   315
               Left            =   120
               TabIndex        =   44
               Tag             =   "⁄›Ê« Ì—ÃÏ «œŒ«· √”„ «·ÕÌ"
               Top             =   1320
               Width           =   2130
               _ExtentX        =   3757
               _ExtentY        =   556
               _Version        =   393216
               Text            =   ""
               RightToLeft     =   -1  'True
            End
            Begin MSDataListLib.DataCombo dcschemeid 
               Height          =   315
               Left            =   150
               TabIndex        =   45
               Tag             =   " "
               Top             =   1665
               Width           =   5865
               _ExtentX        =   10345
               _ExtentY        =   556
               _Version        =   393216
               Text            =   ""
               RightToLeft     =   -1  'True
            End
            Begin MSDataListLib.DataCombo dcaqartypeid 
               Height          =   315
               Left            =   3975
               TabIndex        =   46
               Tag             =   "⁄›Ê« Ì—ÃÏ «œŒ«· ‰Ê⁄ «·⁄ﬁ«—"
               Top             =   240
               Width           =   1410
               _ExtentX        =   2487
               _ExtentY        =   556
               _Version        =   393216
               Text            =   ""
               RightToLeft     =   -1  'True
            End
            Begin MSDataListLib.DataCombo dcBranch 
               Height          =   315
               Left            =   120
               TabIndex        =   47
               Top             =   240
               Width           =   3060
               _ExtentX        =   5398
               _ExtentY        =   556
               _Version        =   393216
               BackColor       =   16777215
               Text            =   ""
               RightToLeft     =   -1  'True
            End
            Begin MSDataListLib.DataCombo DcbSales 
               Height          =   315
               Left            =   120
               TabIndex        =   48
               Top             =   600
               Width           =   3540
               _ExtentX        =   6244
               _ExtentY        =   556
               _Version        =   393216
               BackColor       =   16777215
               Text            =   ""
               RightToLeft     =   -1  'True
            End
            Begin MSComCtl2.DTPicker dpstatusdate 
               Height          =   315
               Left            =   3855
               TabIndex        =   49
               TabStop         =   0   'False
               Top             =   2385
               Width           =   1350
               _ExtentX        =   2381
               _ExtentY        =   556
               _Version        =   393216
               Enabled         =   0   'False
               CalendarBackColor=   12648447
               CalendarTitleBackColor=   10383715
               CustomFormat    =   "yyyy/M/d"
               Format          =   141295619
               CurrentDate     =   37140
            End
            Begin ALLButtonS.ALLButton Command1 
               Height          =   375
               Index           =   8
               Left            =   12000
               TabIndex        =   50
               Top             =   5850
               Width           =   5925
               _ExtentX        =   10451
               _ExtentY        =   661
               BTYPE           =   3
               TX              =   "«·„—›ﬁ« "
               ENAB            =   -1  'True
               BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   178
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               COLTYPE         =   2
               FOCUSR          =   -1  'True
               BCOL            =   8421376
               BCOLO           =   16777152
               FCOL            =   16777215
               FCOLO           =   16777215
               MCOL            =   12632256
               MPTR            =   1
               MICON           =   "RSTradingCenter.frx":C689
               UMCOL           =   -1  'True
               SOFT            =   0   'False
               PICPOS          =   0
               NGREY           =   0   'False
               FX              =   0
               HAND            =   0   'False
               CHECK           =   0   'False
               VALUE           =   0   'False
            End
            Begin C1SizerLibCtl.C1Elastic ELe 
               Height          =   1080
               Index           =   18
               Left            =   0
               TabIndex        =   51
               TabStop         =   0   'False
               Top             =   3825
               Width           =   6405
               _cx             =   11298
               _cy             =   1905
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
               Begin VB.TextBox TxtStreet 
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
                  Left            =   360
                  MaxLength       =   50
                  RightToLeft     =   -1  'True
                  TabIndex        =   55
                  Top             =   705
                  Width           =   2160
               End
               Begin VB.TextBox TxtPart 
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
                  Left            =   3855
                  MaxLength       =   50
                  RightToLeft     =   -1  'True
                  TabIndex        =   54
                  Top             =   705
                  Width           =   1440
               End
               Begin VB.TextBox TxtBlock 
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
                  Height          =   300
                  Left            =   360
                  MaxLength       =   50
                  RightToLeft     =   -1  'True
                  TabIndex        =   53
                  Top             =   360
                  Width           =   2160
               End
               Begin VB.TextBox TxtUnit 
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
                  Height          =   300
                  Left            =   3855
                  MaxLength       =   50
                  RightToLeft     =   -1  'True
                  TabIndex        =   52
                  Top             =   360
                  Width           =   1440
               End
               Begin VB.Label Label18 
                  Alignment       =   1  'Right Justify
                  Caption         =   "⁄œœ «·‘Ê«—⁄"
                  Height          =   255
                  Left            =   2415
                  RightToLeft     =   -1  'True
                  TabIndex        =   60
                  Top             =   705
                  Width           =   1215
               End
               Begin VB.Label Label17 
                  Alignment       =   1  'Right Justify
                  Caption         =   "—ﬁ„ «··ÊÕ…"
                  Height          =   255
                  Left            =   5070
                  RightToLeft     =   -1  'True
                  TabIndex        =   59
                  Top             =   705
                  Width           =   1215
               End
               Begin VB.Label Label16 
                  Alignment       =   1  'Right Justify
                  Caption         =   "—ﬁ„ «·»·Êﬂ"
                  Height          =   240
                  Left            =   2415
                  RightToLeft     =   -1  'True
                  TabIndex        =   58
                  Top             =   360
                  Width           =   1215
               End
               Begin VB.Label Label15 
                  Alignment       =   1  'Right Justify
                  Caption         =   "—ﬁ„ «·Ã«œ…"
                  Height          =   240
                  Left            =   5070
                  RightToLeft     =   -1  'True
                  TabIndex        =   57
                  Top             =   360
                  Width           =   1215
               End
               Begin VB.Label Label24 
                  Alignment       =   1  'Right Justify
                  Caption         =   "Œ«’… »«·«—«÷Ì"
                  Height          =   255
                  Left            =   5070
                  RightToLeft     =   -1  'True
                  TabIndex        =   56
                  Top             =   0
                  Width           =   1215
               End
            End
            Begin C1SizerLibCtl.C1Elastic ELe 
               Height          =   1080
               Index           =   4
               Left            =   6510
               TabIndex        =   61
               TabStop         =   0   'False
               Top             =   3825
               Width           =   5445
               _cx             =   9604
               _cy             =   1905
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
               Begin VB.TextBox txtmetersalevalue 
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
                  Height          =   300
                  Left            =   120
                  MaxLength       =   50
                  RightToLeft     =   -1  'True
                  TabIndex        =   64
                  Top             =   360
                  Width           =   1440
               End
               Begin VB.TextBox TxtPriceHad 
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
                  Height          =   300
                  Left            =   2655
                  MaxLength       =   50
                  RightToLeft     =   -1  'True
                  TabIndex        =   63
                  Top             =   360
                  Width           =   1440
               End
               Begin VB.TextBox TxtPriceSom 
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
                  Left            =   2655
                  MaxLength       =   50
                  RightToLeft     =   -1  'True
                  TabIndex        =   62
                  Top             =   705
                  Width           =   1440
               End
               Begin VB.Label Label1 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  BackStyle       =   0  'Transparent
                  Caption         =   "ﬁÌ„… »Ì⁄ „2"
                  Height          =   270
                  Index           =   18
                  Left            =   1575
                  RightToLeft     =   -1  'True
                  TabIndex        =   68
                  Top             =   360
                  Width           =   930
               End
               Begin VB.Label Label12 
                  Alignment       =   1  'Right Justify
                  Caption         =   "”⁄— «·Õœ"
                  Height          =   240
                  Left            =   4350
                  RightToLeft     =   -1  'True
                  TabIndex        =   67
                  Top             =   360
                  Width           =   855
               End
               Begin VB.Label Label10 
                  Alignment       =   1  'Right Justify
                  Caption         =   "”⁄— «·”Ê„"
                  Height          =   255
                  Left            =   4350
                  RightToLeft     =   -1  'True
                  TabIndex        =   66
                  Top             =   705
                  Width           =   855
               End
               Begin VB.Label Label1 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  BackStyle       =   0  'Transparent
                  Caption         =   "»Ì«‰«  «·»Ì⁄"
                  Height          =   285
                  Index           =   71
                  Left            =   4350
                  RightToLeft     =   -1  'True
                  TabIndex        =   65
                  Top             =   0
                  Width           =   930
               End
            End
            Begin C1SizerLibCtl.C1Elastic ELe 
               Height          =   1680
               Index           =   5
               Left            =   6270
               TabIndex        =   69
               TabStop         =   0   'False
               Top             =   6090
               Width           =   5685
               _cx             =   10028
               _cy             =   2963
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
               Begin VSFlex8UCtl.VSFlexGrid VSFlexGrid1 
                  Height          =   1350
                  Left            =   120
                  TabIndex        =   70
                  Top             =   240
                  Width           =   5475
                  _cx             =   9657
                  _cy             =   2381
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
                  Cols            =   9
                  FixedRows       =   1
                  FixedCols       =   1
                  RowHeightMin    =   320
                  RowHeightMax    =   0
                  ColWidthMin     =   0
                  ColWidthMax     =   0
                  ExtendLastCol   =   -1  'True
                  FormatString    =   $"RSTradingCenter.frx":C6A5
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
               Begin VB.Label Label1 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  BackStyle       =   0  'Transparent
                  Caption         =   "»Ì«‰«  «·Ê”Ìÿ"
                  Height          =   285
                  Index           =   72
                  Left            =   4590
                  RightToLeft     =   -1  'True
                  TabIndex        =   71
                  Top             =   0
                  Width           =   1050
               End
            End
            Begin C1SizerLibCtl.C1Elastic ELe 
               Height          =   1095
               Index           =   6
               Left            =   0
               TabIndex        =   72
               TabStop         =   0   'False
               Top             =   4890
               Width           =   11955
               _cx             =   21087
               _cy             =   1931
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
               Begin VB.TextBox TxteastWriiten 
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
                  Left            =   3375
                  MaxLength       =   50
                  RightToLeft     =   -1  'True
                  TabIndex        =   80
                  Top             =   720
                  Width           =   1440
               End
               Begin VB.TextBox TxtPriceSomW 
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
                  Left            =   5790
                  MaxLength       =   50
                  RightToLeft     =   -1  'True
                  TabIndex        =   79
                  Top             =   720
                  Width           =   1425
               End
               Begin VB.TextBox TxtPriceHadW 
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
                  Left            =   8205
                  MaxLength       =   50
                  RightToLeft     =   -1  'True
                  TabIndex        =   78
                  Top             =   720
                  Width           =   1425
               End
               Begin VB.TextBox txtWestlength 
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
                  Left            =   1200
                  MaxLength       =   50
                  RightToLeft     =   -1  'True
                  TabIndex        =   77
                  Top             =   360
                  Width           =   1440
               End
               Begin VB.TextBox txtSouthlength 
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
                  Left            =   5790
                  MaxLength       =   50
                  RightToLeft     =   -1  'True
                  TabIndex        =   76
                  Top             =   360
                  Width           =   1425
               End
               Begin VB.TextBox txteastlength 
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
                  Left            =   3375
                  MaxLength       =   50
                  RightToLeft     =   -1  'True
                  TabIndex        =   75
                  Top             =   360
                  Width           =   1440
               End
               Begin VB.TextBox txtnorthlength 
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
                  Left            =   8205
                  MaxLength       =   50
                  RightToLeft     =   -1  'True
                  TabIndex        =   74
                  Top             =   360
                  Width           =   1425
               End
               Begin VB.TextBox TxtwestWriiten 
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
                  Left            =   1200
                  MaxLength       =   50
                  RightToLeft     =   -1  'True
                  TabIndex        =   73
                  Top             =   720
                  Width           =   1440
               End
               Begin VB.Label Label23 
                  Alignment       =   1  'Right Justify
                  Caption         =   "‘„«·"
                  Height          =   255
                  Left            =   9645
                  RightToLeft     =   -1  'True
                  TabIndex        =   91
                  Top             =   720
                  Width           =   870
               End
               Begin VB.Label Label22 
                  Alignment       =   1  'Right Justify
                  Caption         =   "‘—ﬁ"
                  Height          =   255
                  Left            =   4455
                  RightToLeft     =   -1  'True
                  TabIndex        =   90
                  Top             =   720
                  Width           =   870
               End
               Begin VB.Label Label21 
                  Alignment       =   1  'Right Justify
                  Caption         =   "Ã‰Ê»"
                  Height          =   255
                  Left            =   6990
                  RightToLeft     =   -1  'True
                  TabIndex        =   89
                  Top             =   720
                  Width           =   870
               End
               Begin VB.Label Label20 
                  Alignment       =   1  'Right Justify
                  Caption         =   "€—»"
                  Height          =   255
                  Left            =   1935
                  RightToLeft     =   -1  'True
                  TabIndex        =   88
                  Top             =   720
                  Width           =   1215
               End
               Begin VB.Label Label19 
                  Alignment       =   1  'Right Justify
                  Caption         =   "«·ÕœÊœ ﬂ «»Â"
                  Height          =   255
                  Left            =   10500
                  RightToLeft     =   -1  'True
                  TabIndex        =   87
                  Top             =   720
                  Width           =   1215
               End
               Begin VB.Label Label7 
                  Alignment       =   1  'Right Justify
                  Caption         =   "«·ÕœÊœ «—ﬁ«„"
                  Height          =   255
                  Left            =   10500
                  RightToLeft     =   -1  'True
                  TabIndex        =   86
                  Top             =   360
                  Width           =   1215
               End
               Begin VB.Label Label6 
                  Alignment       =   1  'Right Justify
                  Caption         =   "€—»"
                  Height          =   255
                  Left            =   1935
                  RightToLeft     =   -1  'True
                  TabIndex        =   85
                  Top             =   360
                  Width           =   1215
               End
               Begin VB.Label Label5 
                  Alignment       =   1  'Right Justify
                  Caption         =   "Ã‰Ê»"
                  Height          =   255
                  Index           =   0
                  Left            =   6990
                  RightToLeft     =   -1  'True
                  TabIndex        =   84
                  Top             =   360
                  Width           =   870
               End
               Begin VB.Label Label4 
                  Alignment       =   1  'Right Justify
                  Caption         =   "‘—ﬁ"
                  Height          =   255
                  Left            =   4455
                  RightToLeft     =   -1  'True
                  TabIndex        =   83
                  Top             =   360
                  Width           =   870
               End
               Begin VB.Label Label3 
                  Alignment       =   1  'Right Justify
                  Caption         =   "‘„«·"
                  Height          =   255
                  Left            =   9645
                  RightToLeft     =   -1  'True
                  TabIndex        =   82
                  Top             =   360
                  Width           =   870
               End
               Begin VB.Label Label25 
                  Alignment       =   1  'Right Justify
                  Caption         =   "«·ÕœÊœ "
                  Height          =   255
                  Left            =   10620
                  RightToLeft     =   -1  'True
                  TabIndex        =   81
                  Top             =   0
                  Width           =   1215
               End
            End
            Begin C1SizerLibCtl.C1Elastic ELe 
               Height          =   5625
               Index           =   7
               Left            =   12060
               TabIndex        =   92
               TabStop         =   0   'False
               Top             =   120
               Width           =   6045
               _cx             =   10663
               _cy             =   9922
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
               Begin VB.Label lblCompanyname 
                  Alignment       =   2  'Center
                  BackStyle       =   0  'Transparent
                  Caption         =   "«·”« —Ì…"
                  BeginProperty Font 
                     Name            =   "Arial"
                     Size            =   27.75
                     Charset         =   178
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H00008000&
                  Height          =   5265
                  Left            =   240
                  TabIndex        =   93
                  Top             =   2625
                  Width           =   5565
               End
               Begin VB.Image Image1 
                  Height          =   2295
                  Left            =   120
                  Picture         =   "RSTradingCenter.frx":C81D
                  Stretch         =   -1  'True
                  Top             =   120
                  Width           =   5850
               End
            End
            Begin MSDataListLib.DataCombo DcFixedAssets 
               Height          =   315
               Left            =   3750
               TabIndex        =   227
               Tag             =   "Õœœ «”„ «·„⁄œ…"
               Top             =   960
               Width           =   4665
               _ExtentX        =   8229
               _ExtentY        =   556
               _Version        =   393216
               Text            =   ""
               RightToLeft     =   -1  'True
            End
            Begin VB.Label Label1 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               BackStyle       =   0  'Transparent
               Caption         =   "„Ã„⁄ «·«Â·«ﬂ"
               Height          =   285
               Index           =   78
               Left            =   2415
               RightToLeft     =   -1  'True
               TabIndex        =   231
               Top             =   960
               Width           =   1170
            End
            Begin VB.Label Label1 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               BackStyle       =   0  'Transparent
               Caption         =   "«”„ «·«’·"
               Height          =   285
               Index           =   77
               Left            =   9885
               RightToLeft     =   -1  'True
               TabIndex        =   229
               Top             =   960
               Width           =   1905
            End
            Begin VB.Label Label2 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Height          =   375
               Index           =   2
               Left            =   4065
               RightToLeft     =   -1  'True
               TabIndex        =   122
               Top             =   1080
               Width           =   7605
            End
            Begin VB.Label Label1 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               BackStyle       =   0  'Transparent
               Caption         =   "«·„Õ’·"
               Height          =   285
               Index           =   22
               Left            =   4710
               RightToLeft     =   -1  'True
               TabIndex        =   121
               Top             =   600
               Width           =   810
            End
            Begin VB.Label Label9 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "«·›—⁄"
               ForeColor       =   &H00000000&
               Height          =   375
               Left            =   3360
               TabIndex        =   120
               Top             =   240
               Width           =   450
            End
            Begin VB.Label Label1 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               BackStyle       =   0  'Transparent
               Caption         =   "‰Ê⁄ «·⁄ﬁ«—"
               Height          =   285
               Index           =   16
               Left            =   5070
               RightToLeft     =   -1  'True
               TabIndex        =   119
               Top             =   240
               Width           =   1050
            End
            Begin VB.Label Label1 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               BackStyle       =   0  'Transparent
               Caption         =   "«”„ «·„Œÿÿ"
               Height          =   285
               Index           =   15
               Left            =   6030
               RightToLeft     =   -1  'True
               TabIndex        =   118
               Top             =   1665
               Width           =   1410
            End
            Begin VB.Label Label1 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               BackStyle       =   0  'Transparent
               Caption         =   "„”·”· «·⁄ﬁ«—"
               Height          =   195
               Index           =   3
               Left            =   10800
               RightToLeft     =   -1  'True
               TabIndex        =   117
               Top             =   240
               Width           =   990
            End
            Begin VB.Label Label1 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               BackStyle       =   0  'Transparent
               Caption         =   "«”„ «·⁄ﬁ«—"
               Height          =   285
               Index           =   1
               Left            =   9885
               RightToLeft     =   -1  'True
               TabIndex        =   116
               Top             =   600
               Width           =   1905
            End
            Begin VB.Label Label1 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               BackStyle       =   0  'Transparent
               Caption         =   "«”„ «·œÊ·Â"
               Height          =   270
               Index           =   0
               Left            =   9885
               RightToLeft     =   -1  'True
               TabIndex        =   115
               Top             =   1320
               Width           =   1905
            End
            Begin VB.Label Label1 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               BackStyle       =   0  'Transparent
               Caption         =   "«”„ «·„œÌ‰Â"
               Height          =   285
               Index           =   4
               Left            =   6030
               RightToLeft     =   -1  'True
               TabIndex        =   114
               Top             =   1320
               Width           =   1410
            End
            Begin VB.Label Label1 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               BackStyle       =   0  'Transparent
               Caption         =   "«”„ «·ÕÌ"
               Height          =   270
               Index           =   5
               Left            =   2535
               RightToLeft     =   -1  'True
               TabIndex        =   113
               Top             =   1320
               Width           =   1050
            End
            Begin VB.Label Label1 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               BackStyle       =   0  'Transparent
               Caption         =   "«”„ «·‘«—⁄"
               Height          =   285
               Index           =   6
               Left            =   10740
               RightToLeft     =   -1  'True
               TabIndex        =   112
               Top             =   1680
               Width           =   1050
            End
            Begin VB.Label Label1 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               BackStyle       =   0  'Transparent
               Caption         =   "—ﬁ„ «·⁄ﬁ«—"
               Height          =   285
               Index           =   7
               Left            =   7110
               RightToLeft     =   -1  'True
               TabIndex        =   111
               Top             =   240
               Width           =   1065
            End
            Begin VB.Label Label1 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               BackStyle       =   0  'Transparent
               Caption         =   "«·⁄„— «·“„‰Ì"
               Height          =   285
               Index           =   37
               Left            =   10485
               RightToLeft     =   -1  'True
               TabIndex        =   110
               Top             =   2025
               Width           =   1305
            End
            Begin VB.Label Label1 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               BackStyle       =   0  'Transparent
               Caption         =   "‰Ê⁄ «·’Ì«‰…"
               Height          =   285
               Index           =   35
               Left            =   10725
               RightToLeft     =   -1  'True
               TabIndex        =   109
               Top             =   2385
               Width           =   1065
            End
            Begin VB.Label Label1 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               BackStyle       =   0  'Transparent
               Caption         =   "«·Õ«·Â"
               Height          =   285
               Index           =   20
               Left            =   6030
               RightToLeft     =   -1  'True
               TabIndex        =   108
               Top             =   2385
               Width           =   1410
            End
            Begin VB.Label Label1 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               BackStyle       =   0  'Transparent
               Caption         =   "⁄œœ «·„œ«Œ·"
               Height          =   285
               Index           =   8
               Left            =   2655
               RightToLeft     =   -1  'True
               TabIndex        =   107
               Top             =   2385
               Width           =   1050
            End
            Begin VB.Label Label1 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               BackStyle       =   0  'Transparent
               Caption         =   "”⁄— «·⁄ﬁ«— «·Õ«·Ì"
               Height          =   285
               Index           =   11
               Left            =   6000
               RightToLeft     =   -1  'True
               TabIndex        =   106
               Top             =   2025
               Width           =   1410
            End
            Begin VB.Label Label1 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               BackStyle       =   0  'Transparent
               Caption         =   "«Œ— ﬁÌ„Â «ÌÃ«—ÌÂ"
               Height          =   285
               Index           =   12
               Left            =   2295
               RightToLeft     =   -1  'True
               TabIndex        =   105
               Top             =   2025
               Width           =   1410
            End
            Begin VB.Label Label14 
               Alignment       =   1  'Right Justify
               Caption         =   "„»·€ „ﬁÿÊ⁄"
               Height          =   255
               Left            =   2850
               RightToLeft     =   -1  'True
               TabIndex        =   104
               Top             =   3585
               Width           =   855
            End
            Begin VB.Label Label13 
               Alignment       =   1  'Right Justify
               Caption         =   "‰”»… «·⁄„Ê·…"
               Height          =   285
               Left            =   6030
               RightToLeft     =   -1  'True
               TabIndex        =   103
               Top             =   3465
               Width           =   1410
            End
            Begin VB.Label Label8 
               Alignment       =   1  'Right Justify
               Caption         =   "«·„”«ÕÂ «·«Ã„«·ÌÂ"
               Height          =   285
               Left            =   6030
               RightToLeft     =   -1  'True
               TabIndex        =   102
               Top             =   3225
               Width           =   1410
            End
            Begin VB.Label Label1 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               BackStyle       =   0  'Transparent
               Caption         =   "⁄œœ «·„Ê«ﬁ›"
               Height          =   285
               Index           =   34
               Left            =   2655
               RightToLeft     =   -1  'True
               TabIndex        =   101
               Top             =   2745
               Width           =   1050
            End
            Begin VB.Label Label1 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               BackStyle       =   0  'Transparent
               Caption         =   "ﬁÌ„… «ÌÃ«— „2"
               Height          =   285
               Index           =   17
               Left            =   2655
               RightToLeft     =   -1  'True
               TabIndex        =   100
               Top             =   3105
               Width           =   1050
            End
            Begin VB.Label Label1 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               BackStyle       =   0  'Transparent
               Caption         =   "⁄œœ «·ÿÊ«»ﬁ"
               Height          =   285
               Index           =   21
               Left            =   10740
               RightToLeft     =   -1  'True
               TabIndex        =   99
               Top             =   2745
               Width           =   1050
            End
            Begin VB.Label Label1 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               BackStyle       =   0  'Transparent
               Caption         =   "«·Ê«ÃÂ…"
               Height          =   285
               Index           =   19
               Left            =   10365
               RightToLeft     =   -1  'True
               TabIndex        =   98
               Top             =   3105
               Width           =   1425
            End
            Begin VB.Label Label1 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               BackStyle       =   0  'Transparent
               Caption         =   "⁄œœ «·ÊÕœ«  «·”ﬂ‰Ì…"
               Height          =   285
               Index           =   9
               Left            =   10365
               RightToLeft     =   -1  'True
               TabIndex        =   97
               Top             =   3465
               Width           =   1425
            End
            Begin VB.Label Label1 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               BackStyle       =   0  'Transparent
               Caption         =   "⁄œœ «·ÊÕœ«  «· Ã«—ÌÂ"
               Height          =   285
               Index           =   10
               Left            =   6030
               RightToLeft     =   -1  'True
               TabIndex        =   96
               Top             =   2745
               Width           =   1410
            End
            Begin VB.Label Label1 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               BackStyle       =   0  'Transparent
               Caption         =   "„Êﬁ⁄ ÃÊÃ·"
               Height          =   285
               Index           =   33
               Left            =   5190
               RightToLeft     =   -1  'True
               TabIndex        =   95
               Top             =   7155
               Width           =   930
            End
            Begin VB.Label Label1 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               BackStyle       =   0  'Transparent
               Caption         =   "Ê’› «·„Êﬁ⁄"
               Height          =   285
               Index           =   29
               Left            =   5190
               RightToLeft     =   -1  'True
               TabIndex        =   94
               Top             =   6090
               Width           =   1050
            End
         End
         Begin C1SizerLibCtl.C1Elastic ELe 
            Height          =   7785
            Index           =   0
            Left            =   -19455
            TabIndex        =   123
            TabStop         =   0   'False
            Top             =   420
            Width           =   18210
            _cx             =   32120
            _cy             =   13732
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
            Begin VB.CheckBox ChkOrder 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   " — Ì» ÿ»ﬁ« ·—ﬁ„ «·ÊÕœ…"
               ForeColor       =   &H000000FF&
               Height          =   195
               Left            =   360
               RightToLeft     =   -1  'True
               TabIndex        =   232
               Top             =   5520
               Width           =   1935
            End
            Begin VSFlex8UCtl.VSFlexGrid UnitsGrid 
               Height          =   3525
               Left            =   0
               TabIndex        =   124
               Top             =   1920
               Width           =   18135
               _cx             =   31988
               _cy             =   6218
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
               Rows            =   1
               Cols            =   30
               FixedRows       =   1
               FixedCols       =   1
               RowHeightMin    =   320
               RowHeightMax    =   0
               ColWidthMin     =   0
               ColWidthMax     =   0
               ExtendLastCol   =   -1  'True
               FormatString    =   $"RSTradingCenter.frx":1CEDB
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
            Begin VSFlex8UCtl.VSFlexGrid VSFlexGrid2 
               Height          =   1485
               Left            =   0
               TabIndex        =   125
               Top             =   6090
               Width           =   18135
               _cx             =   31988
               _cy             =   2619
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
               Cols            =   9
               FixedRows       =   1
               FixedCols       =   1
               RowHeightMin    =   320
               RowHeightMax    =   0
               ColWidthMin     =   0
               ColWidthMax     =   0
               ExtendLastCol   =   -1  'True
               FormatString    =   $"RSTradingCenter.frx":1D382
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
               Height          =   390
               Index           =   21
               Left            =   16635
               TabIndex        =   126
               Top             =   5370
               Visible         =   0   'False
               Width           =   1425
               _ExtentX        =   2514
               _ExtentY        =   688
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
               ButtonImage     =   "RSTradingCenter.frx":1D504
               ColorButton     =   14871017
               DrawFocusRectangle=   0   'False
            End
            Begin ImpulseButton.ISButton Cmd 
               Height          =   390
               Index           =   10
               Left            =   15195
               TabIndex        =   127
               Top             =   5370
               Visible         =   0   'False
               Width           =   1410
               _ExtentX        =   2487
               _ExtentY        =   688
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
               ButtonImage     =   "RSTradingCenter.frx":1DA9E
               ColorButton     =   14871017
               DrawFocusRectangle=   0   'False
            End
            Begin C1SizerLibCtl.C1Elastic ELe 
               Height          =   2055
               Index           =   8
               Left            =   0
               TabIndex        =   128
               TabStop         =   0   'False
               Top             =   120
               Width           =   18225
               _cx             =   32147
               _cy             =   3625
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
               Begin VB.CheckBox chkIsTax 
                  Alignment       =   1  'Right Justify
                  Caption         =   "ÌŒ÷⁄ ··÷—Ì»…"
                  Height          =   255
                  Left            =   7530
                  RightToLeft     =   -1  'True
                  TabIndex        =   256
                  Top             =   1470
                  Width           =   1335
               End
               Begin VB.TextBox DataBaseUnitNio 
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
                  Left            =   15960
                  MaxLength       =   50
                  RightToLeft     =   -1  'True
                  TabIndex        =   253
                  Top             =   0
                  Visible         =   0   'False
                  Width           =   825
               End
               Begin VB.ComboBox DcbTyped 
                  BackColor       =   &H00FFFFC0&
                  Height          =   315
                  ItemData        =   "RSTradingCenter.frx":1E038
                  Left            =   9000
                  List            =   "RSTradingCenter.frx":1E042
                  RightToLeft     =   -1  'True
                  TabIndex        =   251
                  Top             =   1440
                  Width           =   2955
               End
               Begin VB.TextBox RecGID 
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
                  Left            =   4980
                  MaxLength       =   50
                  RightToLeft     =   -1  'True
                  TabIndex        =   249
                  Top             =   1560
                  Visible         =   0   'False
                  Width           =   825
               End
               Begin VB.TextBox UnitElc 
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
                  TabIndex        =   248
                  Top             =   720
                  Width           =   1785
               End
               Begin VB.TextBox Disc 
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
                  TabIndex        =   245
                  Top             =   1080
                  Width           =   7545
               End
               Begin VB.TextBox BathNo 
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
                  Left            =   3600
                  MaxLength       =   50
                  RightToLeft     =   -1  'True
                  TabIndex        =   237
                  Top             =   720
                  Width           =   1065
               End
               Begin VB.TextBox RentValue 
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
                  Left            =   3600
                  MaxLength       =   50
                  RightToLeft     =   -1  'True
                  TabIndex        =   236
                  Top             =   360
                  Width           =   1065
               End
               Begin VB.TextBox UnitID 
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
                  Left            =   15960
                  MaxLength       =   50
                  RightToLeft     =   -1  'True
                  TabIndex        =   233
                  Top             =   360
                  Width           =   825
               End
               Begin VB.TextBox TxtMiniRentValue 
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
                  TabIndex        =   142
                  Top             =   360
                  Width           =   1785
               End
               Begin VB.TextBox TxtTo 
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
                  Left            =   14760
                  MaxLength       =   50
                  RightToLeft     =   -1  'True
                  TabIndex        =   141
                  Top             =   1440
                  Width           =   705
               End
               Begin VB.TextBox TxtFrom 
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
                  Left            =   15930
                  MaxLength       =   50
                  RightToLeft     =   -1  'True
                  TabIndex        =   140
                  Top             =   1440
                  Width           =   705
               End
               Begin VB.TextBox TxtKitchn 
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
                  Left            =   11520
                  MaxLength       =   50
                  RightToLeft     =   -1  'True
                  TabIndex        =   139
                  Top             =   720
                  Width           =   705
               End
               Begin VB.TextBox TxtACCountÚSpleat 
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
                  Left            =   5880
                  MaxLength       =   50
                  RightToLeft     =   -1  'True
                  TabIndex        =   138
                  Top             =   720
                  Width           =   1905
               End
               Begin VB.TextBox TxtLenght 
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
                  Left            =   7440
                  MaxLength       =   50
                  RightToLeft     =   -1  'True
                  TabIndex        =   137
                  Top             =   360
                  Width           =   705
               End
               Begin VB.TextBox TxtACCount 
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
                  Left            =   9000
                  MaxLength       =   50
                  RightToLeft     =   -1  'True
                  TabIndex        =   136
                  Top             =   720
                  Width           =   1305
               End
               Begin VB.TextBox TxtFloors 
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
                  Left            =   11520
                  MaxLength       =   50
                  RightToLeft     =   -1  'True
                  TabIndex        =   135
                  Top             =   360
                  Width           =   705
               End
               Begin VB.TextBox TxtLoung 
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
                  Left            =   13080
                  MaxLength       =   50
                  RightToLeft     =   -1  'True
                  TabIndex        =   134
                  Top             =   720
                  Width           =   1755
               End
               Begin VB.TextBox TxtRooms 
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
                  Left            =   15960
                  MaxLength       =   50
                  RightToLeft     =   -1  'True
                  TabIndex        =   133
                  Top             =   720
                  Width           =   825
               End
               Begin VB.TextBox txtMeterPrice 
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
                  Left            =   5880
                  MaxLength       =   50
                  RightToLeft     =   -1  'True
                  TabIndex        =   132
                  Top             =   360
                  Width           =   705
               End
               Begin VB.ComboBox cBORENTTYPE 
                  BackColor       =   &H00FFFFC0&
                  Height          =   315
                  ItemData        =   "RSTradingCenter.frx":1E053
                  Left            =   9000
                  List            =   "RSTradingCenter.frx":1E05D
                  RightToLeft     =   -1  'True
                  TabIndex        =   131
                  Top             =   360
                  Width           =   1305
               End
               Begin VB.TextBox TxtCount 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00FFFFC0&
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
                  Left            =   13080
                  MaxLength       =   50
                  RightToLeft     =   -1  'True
                  TabIndex        =   130
                  Top             =   1440
                  Width           =   585
               End
               Begin VB.ComboBox cBORENTTYPE1 
                  Height          =   315
                  ItemData        =   "RSTradingCenter.frx":1E07B
                  Left            =   3420
                  List            =   "RSTradingCenter.frx":1E085
                  RightToLeft     =   -1  'True
                  TabIndex        =   129
                  Top             =   1560
                  Visible         =   0   'False
                  Width           =   1455
               End
               Begin ImpulseButton.ISButton Cmd 
                  Height          =   315
                  Index           =   20
                  Left            =   2400
                  TabIndex        =   143
                  Top             =   1440
                  Width           =   720
                  _ExtentX        =   1270
                  _ExtentY        =   556
                  ButtonStyle     =   1
                  ButtonPositionImage=   1
                  Caption         =   "≈÷«›…"
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "MS Sans Serif"
                     Size            =   8.25
                     Charset         =   178
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ButtonImage     =   "RSTradingCenter.frx":1E0A0
                  DrawFocusRectangle=   0   'False
               End
               Begin MSDataListLib.DataCombo DCAkarUnit 
                  Height          =   315
                  Left            =   13080
                  TabIndex        =   144
                  Tag             =   "⁄›Ê« Ì—ÃÏ «œŒ«· ‰Ê⁄ «·⁄ﬁ«—"
                  Top             =   360
                  Width           =   1755
                  _ExtentX        =   3096
                  _ExtentY        =   556
                  _Version        =   393216
                  BackColor       =   16777152
                  Text            =   ""
                  RightToLeft     =   -1  'True
               End
               Begin XtremeSuiteControls.CheckBox FerntChk 
                  Height          =   240
                  Left            =   16320
                  TabIndex        =   240
                  Top             =   1080
                  Width           =   900
                  _Version        =   786432
                  _ExtentX        =   1587
                  _ExtentY        =   423
                  _StockProps     =   79
                  Caption         =   "«· √ÀÌÀ"
                  UseVisualStyle  =   -1  'True
                  TextAlignment   =   1
                  RightToLeft     =   -1  'True
               End
               Begin MSDataListLib.DataCombo UnitStatus 
                  Height          =   315
                  Left            =   13080
                  TabIndex        =   241
                  Tag             =   "⁄›Ê« Ì—ÃÏ «œŒ«· ‰Ê⁄ «·⁄ﬁ«—"
                  Top             =   1080
                  Width           =   1755
                  _ExtentX        =   3096
                  _ExtentY        =   556
                  _Version        =   393216
                  BackColor       =   16777152
                  Text            =   ""
                  RightToLeft     =   -1  'True
               End
               Begin MSDataListLib.DataCombo RenterDC 
                  Height          =   315
                  Left            =   9000
                  TabIndex        =   243
                  Tag             =   "⁄›Ê« Ì—ÃÏ «œŒ«· ‰Ê⁄ «·⁄ﬁ«—"
                  Top             =   1080
                  Width           =   2955
                  _ExtentX        =   5212
                  _ExtentY        =   556
                  _Version        =   393216
                  BackColor       =   16777152
                  Text            =   ""
                  RightToLeft     =   -1  'True
               End
               Begin ImpulseButton.ISButton CmdDelete 
                  Height          =   315
                  Left            =   360
                  TabIndex        =   247
                  Top             =   1440
                  Width           =   840
                  _ExtentX        =   1482
                  _ExtentY        =   556
                  ButtonStyle     =   1
                  ButtonPositionImage=   1
                  Caption         =   "Õ–›"
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "MS Sans Serif"
                     Size            =   8.25
                     Charset         =   178
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ButtonImage     =   "RSTradingCenter.frx":1E43A
                  DrawFocusRectangle=   0   'False
               End
               Begin ImpulseButton.ISButton Editbtn 
                  Height          =   315
                  Left            =   1320
                  TabIndex        =   250
                  Top             =   1440
                  Width           =   840
                  _ExtentX        =   1482
                  _ExtentY        =   556
                  ButtonStyle     =   1
                  ButtonPositionImage=   1
                  Caption         =   " ⁄œÌ·"
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "MS Sans Serif"
                     Size            =   8.25
                     Charset         =   178
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ButtonImage     =   "RSTradingCenter.frx":24C9C
                  DrawFocusRectangle=   0   'False
               End
               Begin VB.Label Label1 
                  Alignment       =   2  'Center
                  BackColor       =   &H00E2E9E9&
                  BackStyle       =   0  'Transparent
                  Caption         =   "«·‰Ê⁄"
                  Height          =   285
                  Index           =   86
                  Left            =   12120
                  RightToLeft     =   -1  'True
                  TabIndex        =   252
                  Top             =   1440
                  Width           =   930
               End
               Begin VB.Label Label1 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  BackStyle       =   0  'Transparent
                  Caption         =   "Ê’› «·ÊÕœ…"
                  Height          =   285
                  Index           =   85
                  Left            =   7800
                  RightToLeft     =   -1  'True
                  TabIndex        =   246
                  Top             =   1080
                  Width           =   1050
               End
               Begin VB.Label Label1 
                  Alignment       =   2  'Center
                  BackColor       =   &H00E2E9E9&
                  BackStyle       =   0  'Transparent
                  Caption         =   "«”„ «·„” √Ã—"
                  Height          =   285
                  Index           =   84
                  Left            =   12000
                  RightToLeft     =   -1  'True
                  TabIndex        =   244
                  Top             =   1080
                  Width           =   1050
               End
               Begin VB.Label Label1 
                  Alignment       =   2  'Center
                  BackColor       =   &H00E2E9E9&
                  BackStyle       =   0  'Transparent
                  Caption         =   "«·Õ«·…"
                  Height          =   285
                  Index           =   83
                  Left            =   15000
                  RightToLeft     =   -1  'True
                  TabIndex        =   242
                  Top             =   1080
                  Width           =   930
               End
               Begin VB.Label Label1 
                  Alignment       =   2  'Center
                  BackColor       =   &H00E2E9E9&
                  BackStyle       =   0  'Transparent
                  Caption         =   "—ﬁ„ Õ ﬂÂ—»«¡ «·ÊÕœ…"
                  Height          =   285
                  Index           =   82
                  Left            =   2040
                  RightToLeft     =   -1  'True
                  TabIndex        =   239
                  Top             =   720
                  Width           =   1530
               End
               Begin VB.Label Label1 
                  Alignment       =   2  'Center
                  BackColor       =   &H00E2E9E9&
                  BackStyle       =   0  'Transparent
                  Caption         =   "œÊ—«  „Ì«…"
                  Height          =   285
                  Index           =   81
                  Left            =   4800
                  RightToLeft     =   -1  'True
                  TabIndex        =   238
                  Top             =   720
                  Width           =   930
               End
               Begin VB.Label Label1 
                  Alignment       =   2  'Center
                  BackColor       =   &H00E2E9E9&
                  BackStyle       =   0  'Transparent
                  Caption         =   "«·ﬁÌ„… «· √ÃÌ—Ì…"
                  Height          =   285
                  Index           =   80
                  Left            =   4680
                  RightToLeft     =   -1  'True
                  TabIndex        =   235
                  Top             =   360
                  Width           =   1170
               End
               Begin VB.Label Label1 
                  Alignment       =   2  'Center
                  BackColor       =   &H00E2E9E9&
                  BackStyle       =   0  'Transparent
                  Caption         =   "—ﬁ„ «·ÊÕœ…"
                  Height          =   285
                  Index           =   79
                  Left            =   16920
                  RightToLeft     =   -1  'True
                  TabIndex        =   234
                  Top             =   360
                  Width           =   810
               End
               Begin VB.Label Label1 
                  Alignment       =   2  'Center
                  BackColor       =   &H00E2E9E9&
                  BackStyle       =   0  'Transparent
                  Caption         =   "«ﬁ· ﬁÌ„…  √ÃÌ—Ì…"
                  Height          =   285
                  Index           =   70
                  Left            =   2040
                  RightToLeft     =   -1  'True
                  TabIndex        =   160
                  Top             =   360
                  Width           =   1530
               End
               Begin VB.Label Label1 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  BackStyle       =   0  'Transparent
                  Caption         =   "«·Ï"
                  Height          =   285
                  Index           =   69
                  Left            =   15240
                  RightToLeft     =   -1  'True
                  TabIndex        =   159
                  Top             =   1440
                  Width           =   570
               End
               Begin VB.Label Label1 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  BackStyle       =   0  'Transparent
                  Caption         =   "„‰"
                  Height          =   285
                  Index           =   68
                  Left            =   16680
                  RightToLeft     =   -1  'True
                  TabIndex        =   158
                  Top             =   1440
                  Width           =   210
               End
               Begin VB.Label Label1 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  BackStyle       =   0  'Transparent
                  Caption         =   " ”·”· «·ÊÕœ… "
                  Height          =   285
                  Index           =   67
                  Left            =   16920
                  RightToLeft     =   -1  'True
                  TabIndex        =   157
                  Top             =   1440
                  Width           =   1050
               End
               Begin VB.Label Label1 
                  Alignment       =   2  'Center
                  BackColor       =   &H00E2E9E9&
                  BackStyle       =   0  'Transparent
                  Caption         =   "„ÿ«»Œ"
                  Height          =   285
                  Index           =   66
                  Left            =   12360
                  RightToLeft     =   -1  'True
                  TabIndex        =   156
                  Top             =   720
                  Width           =   690
               End
               Begin VB.Label Label1 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  BackStyle       =   0  'Transparent
                  Caption         =   "„ﬂÌ›«  ”»·Ì "
                  Height          =   285
                  Index           =   64
                  Left            =   7800
                  RightToLeft     =   -1  'True
                  TabIndex        =   155
                  Top             =   720
                  Width           =   1050
               End
               Begin VB.Label Label1 
                  Alignment       =   2  'Center
                  BackColor       =   &H00E2E9E9&
                  BackStyle       =   0  'Transparent
                  Caption         =   "«·„”«Õ…"
                  Height          =   285
                  Index           =   44
                  Left            =   8280
                  RightToLeft     =   -1  'True
                  TabIndex        =   154
                  Top             =   360
                  Width           =   570
               End
               Begin VB.Label Label1 
                  Alignment       =   2  'Center
                  BackColor       =   &H00E2E9E9&
                  BackStyle       =   0  'Transparent
                  Caption         =   "„ﬂÌ›«  ‘»«ﬂ"
                  Height          =   285
                  Index           =   41
                  Left            =   10440
                  RightToLeft     =   -1  'True
                  TabIndex        =   153
                  Top             =   720
                  Width           =   930
               End
               Begin VB.Label Label1 
                  Alignment       =   2  'Center
                  BackColor       =   &H00E2E9E9&
                  BackStyle       =   0  'Transparent
                  Caption         =   "«·ÿ«»ﬁ"
                  Height          =   285
                  Index           =   43
                  Left            =   12240
                  RightToLeft     =   -1  'True
                  TabIndex        =   152
                  Top             =   360
                  Width           =   810
               End
               Begin VB.Label Label1 
                  Alignment       =   2  'Center
                  BackColor       =   &H00E2E9E9&
                  BackStyle       =   0  'Transparent
                  Caption         =   "⁄œœ «·’«·« "
                  Height          =   285
                  Index           =   39
                  Left            =   15000
                  RightToLeft     =   -1  'True
                  TabIndex        =   151
                  Top             =   720
                  Width           =   930
               End
               Begin VB.Label Label1 
                  Alignment       =   2  'Center
                  BackColor       =   &H00E2E9E9&
                  BackStyle       =   0  'Transparent
                  Caption         =   "⁄œœ «·€—›"
                  Height          =   285
                  Index           =   28
                  Left            =   16920
                  RightToLeft     =   -1  'True
                  TabIndex        =   150
                  Top             =   720
                  Width           =   810
               End
               Begin VB.Label Label1 
                  Alignment       =   2  'Center
                  BackColor       =   &H00E2E9E9&
                  BackStyle       =   0  'Transparent
                  Caption         =   "”⁄— «·„ —"
                  Height          =   285
                  Index           =   27
                  Left            =   6600
                  RightToLeft     =   -1  'True
                  TabIndex        =   149
                  Top             =   360
                  Width           =   810
               End
               Begin VB.Label Label1 
                  Alignment       =   2  'Center
                  BackColor       =   &H00E2E9E9&
                  BackStyle       =   0  'Transparent
                  Caption         =   "ÿ—Ìﬁ… «· √ÃÌ—"
                  Height          =   285
                  Index           =   26
                  Left            =   10440
                  RightToLeft     =   -1  'True
                  TabIndex        =   148
                  Top             =   360
                  Width           =   930
               End
               Begin VB.Label Label1 
                  Alignment       =   2  'Center
                  BackColor       =   &H00E2E9E9&
                  BackStyle       =   0  'Transparent
                  Caption         =   "‰Ê⁄ «·ÊÕœ…"
                  Height          =   285
                  Index           =   24
                  Left            =   15000
                  RightToLeft     =   -1  'True
                  TabIndex        =   147
                  Top             =   360
                  Width           =   930
               End
               Begin VB.Label Label1 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  BackStyle       =   0  'Transparent
                  Caption         =   "«·⁄œœ"
                  Height          =   285
                  Index           =   23
                  Left            =   13440
                  RightToLeft     =   -1  'True
                  TabIndex        =   146
                  Top             =   1440
                  Width           =   570
               End
               Begin VB.Label Label1 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  BackStyle       =   0  'Transparent
                  Caption         =   "»Ì«‰«   ›’Ì·ÌÂ"
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   9.75
                     Charset         =   178
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H00800000&
                  Height          =   285
                  Index           =   73
                  Left            =   16560
                  RightToLeft     =   -1  'True
                  TabIndex        =   145
                  Top             =   0
                  Width           =   1410
               End
            End
            Begin VB.Label Label1 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Height          =   375
               Index           =   14
               Left            =   4590
               RightToLeft     =   -1  'True
               TabIndex        =   162
               Top             =   840
               Width           =   7605
            End
            Begin VB.Label Label1 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               BackStyle       =   0  'Transparent
               Caption         =   "»Ì«‰«  «·„’«⁄œ"
               ForeColor       =   &H00FF0000&
               Height          =   285
               Index           =   38
               Left            =   11100
               RightToLeft     =   -1  'True
               TabIndex        =   161
               Top             =   5610
               Width           =   1050
            End
         End
         Begin C1SizerLibCtl.C1Elastic ELe 
            Height          =   7785
            Index           =   2
            Left            =   -19155
            TabIndex        =   163
            TabStop         =   0   'False
            Top             =   420
            Width           =   18210
            _cx             =   32120
            _cy             =   13732
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
            Begin VB.Frame Frame1 
               Height          =   855
               Left            =   120
               RightToLeft     =   -1  'True
               TabIndex        =   292
               Top             =   3180
               Width           =   10545
               Begin VB.TextBox txtDisountAmount 
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
                  Left            =   90
                  MaxLength       =   50
                  RightToLeft     =   -1  'True
                  TabIndex        =   304
                  Top             =   150
                  Width           =   1575
               End
               Begin VB.TextBox txtPlanned 
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
                  Left            =   7110
                  MaxLength       =   50
                  RightToLeft     =   -1  'True
                  TabIndex        =   295
                  Top             =   450
                  Width           =   2355
               End
               Begin VB.TextBox txtPlotNo 
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
                  Left            =   7110
                  MaxLength       =   50
                  RightToLeft     =   -1  'True
                  TabIndex        =   293
                  Top             =   150
                  Width           =   2385
               End
               Begin Dynamic_Byte.NourHijriCal txtFromPlanneddateH 
                  Height          =   255
                  Left            =   4620
                  TabIndex        =   297
                  Top             =   150
                  Width           =   1695
                  _ExtentX        =   2990
                  _ExtentY        =   450
               End
               Begin MSComCtl2.DTPicker txtFromPlanneddate 
                  Height          =   270
                  Left            =   2940
                  TabIndex        =   298
                  TabStop         =   0   'False
                  Top             =   150
                  Width           =   1695
                  _ExtentX        =   2990
                  _ExtentY        =   476
                  _Version        =   393216
                  CalendarBackColor=   12648447
                  CalendarTitleBackColor=   10383715
                  CustomFormat    =   "yyyy/M/d"
                  Format          =   146472963
                  CurrentDate     =   37140
               End
               Begin Dynamic_Byte.NourHijriCal txtToPlanneddateH 
                  Height          =   255
                  Left            =   4620
                  TabIndex        =   300
                  Top             =   450
                  Width           =   1695
                  _ExtentX        =   2990
                  _ExtentY        =   450
               End
               Begin MSComCtl2.DTPicker txtToPlanneddate 
                  Height          =   270
                  Left            =   2940
                  TabIndex        =   301
                  TabStop         =   0   'False
                  Top             =   450
                  Width           =   1695
                  _ExtentX        =   2990
                  _ExtentY        =   476
                  _Version        =   393216
                  CalendarBackColor=   12648447
                  CalendarTitleBackColor=   10383715
                  CustomFormat    =   "yyyy/M/d"
                  Format          =   146472963
                  CurrentDate     =   37140
               End
               Begin VB.Label Label1 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  BackStyle       =   0  'Transparent
                  Caption         =   "„»·€ «· Œ›Ì÷"
                  Height          =   285
                  Index           =   92
                  Left            =   1740
                  RightToLeft     =   -1  'True
                  TabIndex        =   303
                  Top             =   180
                  Width           =   960
               End
               Begin VB.Label Label1 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  BackStyle       =   0  'Transparent
                  Caption         =   "«·Ï  «—ÌŒ"
                  Height          =   285
                  Index           =   91
                  Left            =   6030
                  RightToLeft     =   -1  'True
                  TabIndex        =   302
                  Top             =   510
                  Width           =   1050
               End
               Begin VB.Label Label1 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  BackStyle       =   0  'Transparent
                  Caption         =   "„‰  «—ÌŒ"
                  Height          =   285
                  Index           =   89
                  Left            =   6000
                  RightToLeft     =   -1  'True
                  TabIndex        =   299
                  Top             =   150
                  Width           =   1050
               End
               Begin VB.Label Label1 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  BackStyle       =   0  'Transparent
                  Caption         =   "«·„Œÿÿ"
                  Height          =   285
                  Index           =   88
                  Left            =   9300
                  RightToLeft     =   -1  'True
                  TabIndex        =   296
                  Top             =   510
                  Width           =   1050
               End
               Begin VB.Label Label1 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  BackStyle       =   0  'Transparent
                  Caption         =   "—ﬁ„ «·ﬁÿ⁄…"
                  Height          =   285
                  Index           =   87
                  Left            =   9300
                  RightToLeft     =   -1  'True
                  TabIndex        =   294
                  Top             =   150
                  Width           =   1050
               End
            End
            Begin C1SizerLibCtl.C1Elastic ELe 
               Height          =   1095
               Index           =   9
               Left            =   120
               TabIndex        =   164
               TabStop         =   0   'False
               Top             =   30
               Width           =   10515
               _cx             =   18547
               _cy             =   1931
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
               Begin VB.OptionButton ComResid 
                  Alignment       =   1  'Right Justify
                  Caption         =   "€Ì— Œ«÷⁄"
                  ForeColor       =   &H00FF0000&
                  Height          =   195
                  Index           =   0
                  Left            =   1200
                  RightToLeft     =   -1  'True
                  TabIndex        =   336
                  Top             =   720
                  Width           =   1095
               End
               Begin VB.OptionButton ComResid 
                  Alignment       =   1  'Right Justify
                  Caption         =   "Œ«÷⁄"
                  ForeColor       =   &H00FF0000&
                  Height          =   195
                  Index           =   1
                  Left            =   240
                  RightToLeft     =   -1  'True
                  TabIndex        =   335
                  Top             =   720
                  Width           =   975
               End
               Begin VB.TextBox Text1 
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
                  Index           =   10
                  Left            =   7680
                  MaxLength       =   50
                  RightToLeft     =   -1  'True
                  TabIndex        =   166
                  Top             =   360
                  Width           =   1545
               End
               Begin VB.TextBox txtsuckno 
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
                  Left            =   7080
                  MaxLength       =   50
                  RightToLeft     =   -1  'True
                  TabIndex        =   165
                  Top             =   720
                  Width           =   2145
               End
               Begin MSDataListLib.DataCombo dcsupplier 
                  Height          =   315
                  Left            =   2640
                  TabIndex        =   167
                  Tag             =   "⁄›Ê« Ì—ÃÏ «œŒ«· √”„ «·ÕÌ"
                  Top             =   360
                  Width           =   4995
                  _ExtentX        =   8811
                  _ExtentY        =   556
                  _Version        =   393216
                  Text            =   ""
                  RightToLeft     =   -1  'True
               End
               Begin Dynamic_Byte.NourHijriCal txtsuckdateH 
                  Height          =   255
                  Left            =   4560
                  TabIndex        =   168
                  Top             =   720
                  Width           =   1695
                  _ExtentX        =   2990
                  _ExtentY        =   450
               End
               Begin MSComCtl2.DTPicker dpsuckdate 
                  Height          =   270
                  Left            =   2640
                  TabIndex        =   169
                  TabStop         =   0   'False
                  Top             =   720
                  Width           =   1695
                  _ExtentX        =   2990
                  _ExtentY        =   476
                  _Version        =   393216
                  CalendarBackColor=   12648447
                  CalendarTitleBackColor=   10383715
                  CustomFormat    =   "yyyy/M/d"
                  Format          =   146538499
                  CurrentDate     =   37140
               End
               Begin MSDataListLib.DataCombo DcbAccount 
                  Height          =   315
                  Left            =   0
                  TabIndex        =   257
                  Tag             =   "⁄›Ê« Ì—ÃÏ «œŒ«· √”„ «·ÕÌ"
                  Top             =   120
                  Visible         =   0   'False
                  Width           =   4995
                  _ExtentX        =   8811
                  _ExtentY        =   556
                  _Version        =   393216
                  Text            =   ""
                  RightToLeft     =   -1  'True
               End
               Begin ImpulseButton.ISButton CmdCus 
                  Height          =   345
                  Left            =   2160
                  TabIndex        =   327
                  Top             =   360
                  Width           =   435
                  _ExtentX        =   767
                  _ExtentY        =   609
                  ButtonStyle     =   1
                  ButtonPositionImage=   1
                  Caption         =   "..."
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "MS Sans Serif"
                     Size            =   8.25
                     Charset         =   178
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ButtonImage     =   "RSTradingCenter.frx":25036
                  DrawFocusRectangle=   0   'False
               End
               Begin VB.Label Label1 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  BackStyle       =   0  'Transparent
                  Caption         =   " «—ÌŒÂ"
                  Height          =   285
                  Index           =   32
                  Left            =   5880
                  RightToLeft     =   -1  'True
                  TabIndex        =   173
                  Top             =   720
                  Width           =   1050
               End
               Begin VB.Label Label1 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  BackStyle       =   0  'Transparent
                  Caption         =   "—ﬁ„ «·’ﬂ"
                  Height          =   285
                  Index           =   31
                  Left            =   9120
                  RightToLeft     =   -1  'True
                  TabIndex        =   172
                  Top             =   720
                  Width           =   1050
               End
               Begin VB.Label Label1 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  BackStyle       =   0  'Transparent
                  Caption         =   "«·„«·ﬂ"
                  Height          =   285
                  Index           =   30
                  Left            =   8280
                  RightToLeft     =   -1  'True
                  TabIndex        =   171
                  Top             =   360
                  Width           =   1890
               End
               Begin VB.Label Label1 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  BackStyle       =   0  'Transparent
                  Caption         =   "»Ì«‰«  «·„·ﬂÌ…"
                  ForeColor       =   &H00800000&
                  Height          =   285
                  Index           =   42
                  Left            =   8400
                  RightToLeft     =   -1  'True
                  TabIndex        =   170
                  Top             =   0
                  Width           =   1890
               End
            End
            Begin C1SizerLibCtl.C1Elastic ELe 
               Height          =   2100
               Index           =   10
               Left            =   120
               TabIndex        =   174
               TabStop         =   0   'False
               Top             =   1110
               Width           =   10515
               _cx             =   18547
               _cy             =   3704
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
               Begin VB.OptionButton RdRTypeDate 
                  Alignment       =   1  'Right Justify
                  Caption         =   "ÂÃ—Ì"
                  ForeColor       =   &H00FF0000&
                  Height          =   195
                  Index           =   0
                  Left            =   1440
                  RightToLeft     =   -1  'True
                  TabIndex        =   338
                  Top             =   120
                  Width           =   735
               End
               Begin VB.OptionButton RdRTypeDate 
                  Alignment       =   1  'Right Justify
                  Caption         =   "„Ì·«œÌ"
                  ForeColor       =   &H00FF0000&
                  Height          =   195
                  Index           =   1
                  Left            =   360
                  RightToLeft     =   -1  'True
                  TabIndex        =   337
                  Top             =   120
                  Width           =   855
               End
               Begin VB.TextBox TxtProvide 
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
                  Left            =   120
                  RightToLeft     =   -1  'True
                  TabIndex        =   184
                  Top             =   1740
                  Width           =   9225
               End
               Begin VB.TextBox TxtRemarks 
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
                  Height          =   285
                  Left            =   120
                  RightToLeft     =   -1  'True
                  TabIndex        =   183
                  Top             =   1410
                  Width           =   4905
               End
               Begin VB.TextBox TxtBanckName 
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
                  Height          =   285
                  Left            =   6600
                  RightToLeft     =   -1  'True
                  TabIndex        =   182
                  Top             =   1410
                  Width           =   2745
               End
               Begin VB.TextBox TxtAcountBank 
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
                  Left            =   120
                  MaxLength       =   50
                  RightToLeft     =   -1  'True
                  TabIndex        =   181
                  Top             =   1050
                  Width           =   4905
               End
               Begin VB.TextBox TxtEmail 
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
                  Left            =   6600
                  MaxLength       =   50
                  RightToLeft     =   -1  'True
                  TabIndex        =   180
                  Top             =   1050
                  Width           =   2745
               End
               Begin VB.TextBox TxtFaxAg 
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
                  Left            =   120
                  MaxLength       =   50
                  RightToLeft     =   -1  'True
                  TabIndex        =   179
                  Top             =   690
                  Width           =   2745
               End
               Begin VB.TextBox TxtMobile 
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
                  Left            =   4080
                  MaxLength       =   50
                  RightToLeft     =   -1  'True
                  TabIndex        =   178
                  Top             =   690
                  Width           =   2025
               End
               Begin VB.TextBox TxtTel 
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
                  Left            =   6600
                  MaxLength       =   50
                  RightToLeft     =   -1  'True
                  TabIndex        =   177
                  Top             =   690
                  Width           =   2745
               End
               Begin VB.TextBox TxtagencyNo 
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
                  Height          =   285
                  Left            =   120
                  MaxLength       =   50
                  RightToLeft     =   -1  'True
                  TabIndex        =   176
                  Top             =   360
                  Width           =   2745
               End
               Begin VB.TextBox txtauthorizationname 
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
                  Height          =   285
                  Left            =   4080
                  MaxLength       =   50
                  RightToLeft     =   -1  'True
                  TabIndex        =   175
                  Top             =   360
                  Width           =   5265
               End
               Begin VB.Label Label1 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  BackStyle       =   0  'Transparent
                  Caption         =   "«·‘—Êÿ"
                  Height          =   285
                  Index           =   63
                  Left            =   9360
                  RightToLeft     =   -1  'True
                  TabIndex        =   195
                  Top             =   1740
                  Width           =   930
               End
               Begin VB.Label Label1 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  BackStyle       =   0  'Transparent
                  Caption         =   "„·«ÕŸ« "
                  Height          =   255
                  Index           =   62
                  Left            =   4920
                  RightToLeft     =   -1  'True
                  TabIndex        =   194
                  Top             =   1410
                  Width           =   1530
               End
               Begin VB.Label Label1 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  BackStyle       =   0  'Transparent
                  Caption         =   "«”„ «·»‰ﬂ"
                  Height          =   255
                  Index           =   57
                  Left            =   9240
                  RightToLeft     =   -1  'True
                  TabIndex        =   193
                  Top             =   1410
                  Width           =   1050
               End
               Begin VB.Label Label1 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  BackStyle       =   0  'Transparent
                  Caption         =   "—ﬁ„ «·Õ”«» «·»‰ﬂÌ"
                  Height          =   285
                  Index           =   52
                  Left            =   4920
                  RightToLeft     =   -1  'True
                  TabIndex        =   192
                  Top             =   1050
                  Width           =   1530
               End
               Begin VB.Label Label1 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  BackStyle       =   0  'Transparent
                  Caption         =   "«Ì„Ì·"
                  Height          =   285
                  Index           =   51
                  Left            =   9240
                  RightToLeft     =   -1  'True
                  TabIndex        =   191
                  Top             =   1050
                  Width           =   1050
               End
               Begin VB.Label TxtFax 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  BackStyle       =   0  'Transparent
                  Caption         =   "›«ﬂ”"
                  Height          =   285
                  Index           =   50
                  Left            =   2760
                  RightToLeft     =   -1  'True
                  TabIndex        =   190
                  Top             =   690
                  Width           =   1050
               End
               Begin VB.Label Label1 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  BackStyle       =   0  'Transparent
                  Caption         =   "—ﬁ„ «·Êﬂ«·Â"
                  Height          =   255
                  Index           =   49
                  Left            =   2760
                  RightToLeft     =   -1  'True
                  TabIndex        =   189
                  Top             =   360
                  Width           =   1050
               End
               Begin VB.Label Label1 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  BackStyle       =   0  'Transparent
                  Caption         =   "«·ÊﬂÌ·"
                  Height          =   255
                  Index           =   48
                  Left            =   8400
                  RightToLeft     =   -1  'True
                  TabIndex        =   188
                  Top             =   360
                  Width           =   1890
               End
               Begin VB.Label Label1 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  BackStyle       =   0  'Transparent
                  Caption         =   " ·›Ê‰"
                  Height          =   285
                  Index           =   47
                  Left            =   9240
                  RightToLeft     =   -1  'True
                  TabIndex        =   187
                  Top             =   690
                  Width           =   1050
               End
               Begin VB.Label Label1 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  BackStyle       =   0  'Transparent
                  Caption         =   "ÃÊ«·"
                  Height          =   285
                  Index           =   45
                  Left            =   5520
                  RightToLeft     =   -1  'True
                  TabIndex        =   186
                  Top             =   690
                  Width           =   1050
               End
               Begin VB.Label Label1 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  BackStyle       =   0  'Transparent
                  Caption         =   "»Ì«‰«  «·ÊﬂÌ·"
                  ForeColor       =   &H00800000&
                  Height          =   285
                  Index           =   74
                  Left            =   8400
                  RightToLeft     =   -1  'True
                  TabIndex        =   185
                  Top             =   0
                  Width           =   1890
               End
            End
            Begin C1SizerLibCtl.C1Elastic ELe 
               Height          =   3615
               Index           =   11
               Left            =   150
               TabIndex        =   196
               TabStop         =   0   'False
               Top             =   4020
               Width           =   10515
               _cx             =   18547
               _cy             =   6376
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
               Begin VB.TextBox txtManulaVat 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00C0FFFF&
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
                  Left            =   2280
                  MaxLength       =   50
                  RightToLeft     =   -1  'True
                  TabIndex        =   339
                  Top             =   0
                  Visible         =   0   'False
                  Width           =   495
               End
               Begin VB.TextBox TxtFATYou 
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
                  Left            =   2850
                  Locked          =   -1  'True
                  MaxLength       =   50
                  RightToLeft     =   -1  'True
                  TabIndex        =   332
                  Top             =   0
                  Width           =   495
               End
               Begin VB.TextBox TxtFATValue 
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
                  Left            =   0
                  Locked          =   -1  'True
                  MaxLength       =   50
                  RightToLeft     =   -1  'True
                  TabIndex        =   331
                  Top             =   0
                  Width           =   825
               End
               Begin VB.TextBox TxtContValueWithout 
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
                  Left            =   4320
                  MaxLength       =   50
                  RightToLeft     =   -1  'True
                  TabIndex        =   329
                  Top             =   0
                  Width           =   1545
               End
               Begin VB.TextBox NOOFYears 
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
                  Height          =   285
                  Left            =   7560
                  MaxLength       =   50
                  RightToLeft     =   -1  'True
                  TabIndex        =   305
                  Top             =   0
                  Width           =   825
               End
               Begin VB.TextBox TxtContValue 
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
                  Left            =   6570
                  MaxLength       =   50
                  RightToLeft     =   -1  'True
                  TabIndex        =   259
                  Top             =   390
                  Width           =   2745
               End
               Begin VB.TextBox TxtValYearIncrease 
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
                  Left            =   90
                  MaxLength       =   50
                  RightToLeft     =   -1  'True
                  TabIndex        =   258
                  Top             =   390
                  Width           =   825
               End
               Begin Dynamic_Byte.NourHijriCal DateHCont 
                  Height          =   255
                  Left            =   3810
                  TabIndex        =   260
                  Top             =   390
                  Width           =   1695
                  _ExtentX        =   2990
                  _ExtentY        =   450
               End
               Begin MSComCtl2.DTPicker DateCont 
                  Height          =   270
                  Left            =   2010
                  TabIndex        =   261
                  TabStop         =   0   'False
                  Top             =   390
                  Width           =   1695
                  _ExtentX        =   2990
                  _ExtentY        =   476
                  _Version        =   393216
                  CalendarBackColor=   12648447
                  CalendarTitleBackColor=   10383715
                  CustomFormat    =   "yyyy/M/d"
                  Format          =   146538499
                  CurrentDate     =   37140
               End
               Begin MSComCtl2.DTPicker FromCotDate 
                  Height          =   285
                  Left            =   6570
                  TabIndex        =   262
                  TabStop         =   0   'False
                  Top             =   750
                  Width           =   1335
                  _ExtentX        =   2355
                  _ExtentY        =   503
                  _Version        =   393216
                  CalendarBackColor=   12648447
                  CalendarTitleBackColor=   10383715
                  CustomFormat    =   "yyyy/M/d"
                  Format          =   146538499
                  CurrentDate     =   37140
               End
               Begin Dynamic_Byte.NourHijriCal FromCotDateH 
                  Height          =   270
                  Left            =   7980
                  TabIndex        =   263
                  Top             =   870
                  Visible         =   0   'False
                  Width           =   1335
                  _ExtentX        =   2355
                  _ExtentY        =   476
               End
               Begin Dynamic_Byte.NourHijriCal ToCotDateH 
                  Height          =   270
                  Left            =   3810
                  TabIndex        =   264
                  Top             =   750
                  Width           =   1695
                  _ExtentX        =   2990
                  _ExtentY        =   476
               End
               Begin MSComCtl2.DTPicker ToCotDate 
                  Height          =   285
                  Left            =   2010
                  TabIndex        =   265
                  TabStop         =   0   'False
                  Top             =   750
                  Width           =   1695
                  _ExtentX        =   2990
                  _ExtentY        =   503
                  _Version        =   393216
                  CalendarBackColor=   12648447
                  CalendarTitleBackColor=   10383715
                  CustomFormat    =   "yyyy/M/d"
                  Format          =   146538499
                  CurrentDate     =   37140
               End
               Begin C1SizerLibCtl.C1Elastic ELe 
                  Height          =   2565
                  Index           =   12
                  Left            =   0
                  TabIndex        =   266
                  TabStop         =   0   'False
                  Top             =   1170
                  Width           =   10335
                  _cx             =   18230
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
                  BackColor       =   -2147483633
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
                  Begin VB.TextBox TxtPaymentCount 
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
                     Left            =   7440
                     MaxLength       =   50
                     RightToLeft     =   -1  'True
                     TabIndex        =   272
                     Top             =   390
                     Width           =   1065
                  End
                  Begin VB.TextBox TxtPeriods 
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
                     Height          =   375
                     Left            =   7440
                     MaxLength       =   50
                     RightToLeft     =   -1  'True
                     TabIndex        =   271
                     Top             =   765
                     Width           =   1065
                  End
                  Begin VB.ComboBox DcbPeriodsID 
                     Height          =   315
                     ItemData        =   "RSTradingCenter.frx":253D0
                     Left            =   6240
                     List            =   "RSTradingCenter.frx":253DD
                     RightToLeft     =   -1  'True
                     TabIndex        =   270
                     Top             =   765
                     Width           =   1095
                  End
                  Begin VB.TextBox TxtPriodAlow 
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
                     Height          =   375
                     Left            =   3240
                     MaxLength       =   50
                     RightToLeft     =   -1  'True
                     TabIndex        =   269
                     Top             =   765
                     Width           =   1065
                  End
                  Begin VB.ComboBox DcbPeriodsAlowID 
                     Height          =   315
                     ItemData        =   "RSTradingCenter.frx":253F0
                     Left            =   2040
                     List            =   "RSTradingCenter.frx":253FD
                     RightToLeft     =   -1  'True
                     TabIndex        =   268
                     Top             =   765
                     Width           =   1095
                  End
                  Begin VB.CheckBox RSOutSupplier 
                     Alignment       =   1  'Right Justify
                     Caption         =   "√„·«ﬂ €Ì—"
                     Height          =   255
                     Left            =   1710
                     RightToLeft     =   -1  'True
                     TabIndex        =   267
                     Top             =   0
                     Width           =   1335
                  End
                  Begin VSFlex8UCtl.VSFlexGrid GridInstallments 
                     Height          =   1440
                     Left            =   150
                     TabIndex        =   273
                     Top             =   1140
                     Width           =   10125
                     _cx             =   17859
                     _cy             =   2540
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
                     Cols            =   40
                     FixedRows       =   1
                     FixedCols       =   1
                     RowHeightMin    =   320
                     RowHeightMax    =   0
                     ColWidthMin     =   0
                     ColWidthMax     =   0
                     ExtendLastCol   =   -1  'True
                     FormatString    =   $"RSTradingCenter.frx":25410
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
                  Begin MSComCtl2.DTPicker FristPaymentDate 
                     Height          =   255
                     Left            =   3360
                     TabIndex        =   274
                     TabStop         =   0   'False
                     Top             =   390
                     Width           =   1335
                     _ExtentX        =   2355
                     _ExtentY        =   450
                     _Version        =   393216
                     CalendarBackColor=   12648447
                     CalendarTitleBackColor=   10383715
                     CustomFormat    =   "yyyy/M/d"
                     Format          =   140574723
                     CurrentDate     =   41640
                  End
                  Begin Dynamic_Byte.NourHijriCal FirstInstallDateH 
                     Height          =   255
                     Left            =   4800
                     TabIndex        =   275
                     Top             =   390
                     Width           =   1335
                     _ExtentX        =   2355
                     _ExtentY        =   450
                  End
                  Begin ImpulseButton.ISButton Cmd 
                     Height          =   360
                     Index           =   0
                     Left            =   360
                     TabIndex        =   276
                     Top             =   765
                     Width           =   720
                     _ExtentX        =   1270
                     _ExtentY        =   635
                     ButtonStyle     =   1
                     ButtonPositionImage=   1
                     Caption         =   "≈÷«›…"
                     BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                        Name            =   "MS Sans Serif"
                        Size            =   8.25
                        Charset         =   178
                        Weight          =   400
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     ButtonImage     =   "RSTradingCenter.frx":25AEE
                     DrawFocusRectangle=   0   'False
                  End
                  Begin C1SizerLibCtl.C1Elastic ELe 
                     Height          =   615
                     Index           =   13
                     Left            =   120
                     TabIndex        =   277
                     TabStop         =   0   'False
                     Top             =   120
                     Width           =   3015
                     _cx             =   5318
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
                     BackColor       =   -2147483633
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
                     Begin VB.OptionButton Rd 
                        Alignment       =   1  'Right Justify
                        Caption         =   "⁄„Ê·…"
                        Enabled         =   0   'False
                        Height          =   435
                        Left            =   1920
                        RightToLeft     =   -1  'True
                        TabIndex        =   279
                        Top             =   120
                        Width           =   855
                     End
                     Begin VB.TextBox TxtKickbacks 
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
                        Height          =   270
                        Left            =   120
                        MaxLength       =   50
                        RightToLeft     =   -1  'True
                        TabIndex        =   278
                        Top             =   285
                        Width           =   1545
                     End
                  End
                  Begin VB.Shape Shape1 
                     BorderColor     =   &H000000FF&
                     BorderWidth     =   2
                     FillColor       =   &H000000FF&
                     Height          =   615
                     Left            =   3240
                     Top             =   120
                     Width           =   4095
                  End
                  Begin VB.Label Label1 
                     Alignment       =   1  'Right Justify
                     BackColor       =   &H00E2E9E9&
                     BackStyle       =   0  'Transparent
                     Caption         =   "»Ì«‰«  «·œ›⁄« "
                     ForeColor       =   &H00800000&
                     Height          =   285
                     Index           =   76
                     Left            =   8040
                     RightToLeft     =   -1  'True
                     TabIndex        =   284
                     Top             =   0
                     Width           =   2250
                  End
                  Begin VB.Label Label1 
                     Alignment       =   1  'Right Justify
                     BackColor       =   &H00E2E9E9&
                     BackStyle       =   0  'Transparent
                     Caption         =   "⁄œœ «·œ›⁄« "
                     Height          =   255
                     Index           =   58
                     Left            =   9120
                     RightToLeft     =   -1  'True
                     TabIndex        =   283
                     Top             =   390
                     Width           =   930
                  End
                  Begin VB.Label Label1 
                     Alignment       =   1  'Right Justify
                     BackColor       =   &H00E2E9E9&
                     BackStyle       =   0  'Transparent
                     Caption         =   " «—ÌŒ «Ê· œ›⁄Â"
                     Height          =   255
                     Index           =   59
                     Left            =   6120
                     RightToLeft     =   -1  'True
                     TabIndex        =   282
                     Top             =   390
                     Width           =   1170
                  End
                  Begin VB.Label Label1 
                     Alignment       =   1  'Right Justify
                     BackColor       =   &H00E2E9E9&
                     BackStyle       =   0  'Transparent
                     Caption         =   "«·› —Â »Ì‰ «·œ›⁄« "
                     Height          =   375
                     Index           =   60
                     Left            =   8640
                     RightToLeft     =   -1  'True
                     TabIndex        =   281
                     Top             =   765
                     Width           =   1410
                  End
                  Begin VB.Label Label1 
                     Alignment       =   1  'Right Justify
                     BackColor       =   &H00E2E9E9&
                     BackStyle       =   0  'Transparent
                     Caption         =   "› —  «·”„«Õ ··œ›⁄« "
                     Height          =   375
                     Index           =   61
                     Left            =   4320
                     RightToLeft     =   -1  'True
                     TabIndex        =   280
                     Top             =   765
                     Width           =   1650
                  End
               End
               Begin VB.Label Label1 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  BackStyle       =   0  'Transparent
                  Caption         =   "ﬁÌ„… «·÷—Ì»…"
                  Height          =   285
                  Index           =   100
                  Left            =   840
                  RightToLeft     =   -1  'True
                  TabIndex        =   334
                  Top             =   0
                  Width           =   930
               End
               Begin VB.Label Label1 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  BackStyle       =   0  'Transparent
                  Caption         =   "‰”»… «·÷—Ì»…"
                  Height          =   285
                  Index           =   99
                  Left            =   3360
                  RightToLeft     =   -1  'True
                  TabIndex        =   333
                  Top             =   0
                  Width           =   930
               End
               Begin VB.Label Label1 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  BackStyle       =   0  'Transparent
                  Caption         =   "ﬁÌ„… «· ⁄«ﬁœ »œÊ‰ ÷—Ì»Â"
                  Height          =   285
                  Index           =   96
                  Left            =   5640
                  RightToLeft     =   -1  'True
                  TabIndex        =   330
                  Top             =   0
                  Width           =   1890
               End
               Begin VB.Label Label1 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  BackStyle       =   0  'Transparent
                  Caption         =   "⁄œœ «·”‰Ê« "
                  Height          =   285
                  Index           =   94
                  Left            =   8400
                  RightToLeft     =   -1  'True
                  TabIndex        =   306
                  Top             =   0
                  Width           =   930
               End
               Begin VB.Label Label1 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  BackStyle       =   0  'Transparent
                  Caption         =   "»Ì«‰«  «·« ›«ﬁÌ…"
                  ForeColor       =   &H00800000&
                  Height          =   285
                  Index           =   75
                  Left            =   8370
                  RightToLeft     =   -1  'True
                  TabIndex        =   291
                  Top             =   30
                  Width           =   1890
               End
               Begin VB.Label Label1 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  BackStyle       =   0  'Transparent
                  Caption         =   "„‰"
                  Height          =   300
                  Index           =   56
                  Left            =   9330
                  RightToLeft     =   -1  'True
                  TabIndex        =   290
                  Top             =   690
                  Width           =   450
               End
               Begin VB.Label Label1 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  BackStyle       =   0  'Transparent
                  Caption         =   "Ì»œ¡"
                  Height          =   300
                  Index           =   50
                  Left            =   9720
                  RightToLeft     =   -1  'True
                  TabIndex        =   289
                  Top             =   750
                  Width           =   570
               End
               Begin VB.Label Label1 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  BackStyle       =   0  'Transparent
                  Caption         =   "ﬁÌ„… «· ⁄«ﬁœ"
                  Height          =   285
                  Index           =   54
                  Left            =   8370
                  RightToLeft     =   -1  'True
                  TabIndex        =   288
                  Top             =   390
                  Width           =   1890
               End
               Begin VB.Label Label1 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  BackStyle       =   0  'Transparent
                  Caption         =   " «—ÌŒ «· ⁄«ﬁœ"
                  Height          =   285
                  Index           =   53
                  Left            =   5490
                  RightToLeft     =   -1  'True
                  TabIndex        =   287
                  Top             =   390
                  Width           =   1050
               End
               Begin VB.Label Label1 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  BackStyle       =   0  'Transparent
                  Caption         =   "Ì‰ ÂÌ » «—ÌŒ"
                  Height          =   300
                  Index           =   55
                  Left            =   5490
                  RightToLeft     =   -1  'True
                  TabIndex        =   286
                  Top             =   750
                  Width           =   1050
               End
               Begin VB.Label Label1 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  BackStyle       =   0  'Transparent
                  Caption         =   "«·“Ì«œÂ «·”‰ÊÌÂ"
                  Height          =   285
                  Index           =   65
                  Left            =   90
                  RightToLeft     =   -1  'True
                  TabIndex        =   285
                  Top             =   390
                  Width           =   1890
               End
            End
            Begin C1SizerLibCtl.C1Elastic ELe 
               Height          =   7650
               Index           =   14
               Left            =   10605
               TabIndex        =   197
               TabStop         =   0   'False
               Top             =   120
               Width           =   7620
               _cx             =   13441
               _cy             =   13494
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
               Begin C1SizerLibCtl.C1Elastic ELe 
                  Height          =   2895
                  Index           =   16
                  Left            =   120
                  TabIndex        =   307
                  TabStop         =   0   'False
                  Top             =   4440
                  Width           =   5175
                  _cx             =   9128
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
                  Begin VB.CommandButton Command6 
                     Caption         =   "⁄—÷ ”‰œ«  «·„œ›Ê⁄«  «·Œ«’Â »«·⁄ﬁœ"
                     Height          =   255
                     Left            =   0
                     RightToLeft     =   -1  'True
                     TabIndex        =   341
                     Top             =   1920
                     Width           =   2895
                  End
                  Begin VB.TextBox TxtTotalValue 
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
                     Left            =   120
                     Locked          =   -1  'True
                     MaxLength       =   50
                     RightToLeft     =   -1  'True
                     TabIndex        =   328
                     Top             =   1200
                     Width           =   1425
                  End
                  Begin VB.CommandButton Command8 
                     Caption         =   "ﬂ‘› Õ”«»"
                     Height          =   255
                     Left            =   120
                     RightToLeft     =   -1  'True
                     TabIndex        =   322
                     Top             =   2400
                     Width           =   960
                  End
                  Begin VB.CommandButton cmdPaymant 
                     Caption         =   "”‰œ ’—› „œ›Ê⁄« "
                     Height          =   255
                     Left            =   3120
                     RightToLeft     =   -1  'True
                     TabIndex        =   320
                     Top             =   1920
                     Width           =   1785
                  End
                  Begin VB.CheckBox CheckGLYearly 
                     Alignment       =   1  'Right Justify
                     Caption         =   "«·ﬁÌœ ”‰ÊÌ ›ﬁÿ"
                     Height          =   255
                     Left            =   840
                     RightToLeft     =   -1  'True
                     TabIndex        =   319
                     Top             =   240
                     Visible         =   0   'False
                     Width           =   1815
                  End
                  Begin VB.TextBox TxtNoteSerial1 
                     Alignment       =   1  'Right Justify
                     Height          =   285
                     Left            =   240
                     Locked          =   -1  'True
                     RightToLeft     =   -1  'True
                     TabIndex        =   316
                     Top             =   480
                     Visible         =   0   'False
                     Width           =   2415
                  End
                  Begin VB.TextBox Text2 
                     Alignment       =   1  'Right Justify
                     Height          =   285
                     Index           =   1
                     Left            =   360
                     RightToLeft     =   -1  'True
                     TabIndex        =   315
                     Top             =   -120
                     Visible         =   0   'False
                     Width           =   375
                  End
                  Begin VB.TextBox TxtNoteSerial 
                     Alignment       =   1  'Right Justify
                     Height          =   285
                     Left            =   240
                     Locked          =   -1  'True
                     RightToLeft     =   -1  'True
                     TabIndex        =   311
                     Top             =   840
                     Width           =   2415
                  End
                  Begin VB.CommandButton Command9 
                     Caption         =   "ÿ»«⁄Â «·ﬁÌœ"
                     Height          =   255
                     Left            =   2880
                     RightToLeft     =   -1  'True
                     TabIndex        =   310
                     Top             =   1560
                     Width           =   975
                  End
                  Begin VB.CommandButton Command11 
                     Caption         =   "Õ–› «·ﬁÌœ"
                     Height          =   255
                     Left            =   1980
                     RightToLeft     =   -1  'True
                     TabIndex        =   309
                     Top             =   1560
                     Width           =   855
                  End
                  Begin VB.CommandButton Command12 
                     Caption         =   "≈‰‘«¡ «·ﬁÌœ"
                     Height          =   255
                     Left            =   3960
                     RightToLeft     =   -1  'True
                     TabIndex        =   308
                     Top             =   1560
                     Width           =   975
                  End
                  Begin MSDataListLib.DataCombo AccountVat 
                     Bindings        =   "RSTradingCenter.frx":25E88
                     Height          =   315
                     Left            =   0
                     TabIndex        =   312
                     Top             =   -240
                     Visible         =   0   'False
                     Width           =   3450
                     _ExtentX        =   6085
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
                  Begin MSComCtl2.DTPicker ToDate 
                     Height          =   270
                     Left            =   1125
                     TabIndex        =   323
                     TabStop         =   0   'False
                     Top             =   2400
                     Width           =   1290
                     _ExtentX        =   2275
                     _ExtentY        =   476
                     _Version        =   393216
                     CalendarBackColor=   12648447
                     CalendarTitleBackColor=   10383715
                     Format          =   146604035
                     CurrentDate     =   41640
                  End
                  Begin MSComCtl2.DTPicker FrmDate 
                     Height          =   270
                     Left            =   2835
                     TabIndex        =   324
                     TabStop         =   0   'False
                     Top             =   2400
                     Width           =   1335
                     _ExtentX        =   2355
                     _ExtentY        =   476
                     _Version        =   393216
                     CalendarBackColor=   12648447
                     CalendarTitleBackColor=   10383715
                     Format          =   146604035
                     CurrentDate     =   41640
                  End
                  Begin VB.Label Label1 
                     Alignment       =   1  'Right Justify
                     BackColor       =   &H00E2E9E9&
                     BackStyle       =   0  'Transparent
                     Caption         =   "«·Ï"
                     Height          =   195
                     Index           =   90
                     Left            =   2475
                     RightToLeft     =   -1  'True
                     TabIndex        =   326
                     Top             =   2400
                     Width           =   270
                  End
                  Begin VB.Label Label1 
                     Alignment       =   1  'Right Justify
                     BackColor       =   &H00E2E9E9&
                     BackStyle       =   0  'Transparent
                     Caption         =   "„‰"
                     Height          =   195
                     Index           =   93
                     Left            =   4155
                     RightToLeft     =   -1  'True
                     TabIndex        =   325
                     Top             =   2400
                     Width           =   270
                  End
                  Begin VB.Label Label1 
                     Alignment       =   1  'Right Justify
                     BackColor       =   &H00E2E9E9&
                     BackStyle       =   0  'Transparent
                     Caption         =   "—ﬁ„ «·⁄ﬁœ"
                     Height          =   195
                     Index           =   95
                     Left            =   2760
                     RightToLeft     =   -1  'True
                     TabIndex        =   317
                     Top             =   480
                     Visible         =   0   'False
                     Width           =   990
                  End
                  Begin VB.Label Label1 
                     Alignment       =   1  'Right Justify
                     BackColor       =   &H00E2E9E9&
                     BackStyle       =   0  'Transparent
                     Caption         =   "»Ì«‰«  „Õ«”»Ì…"
                     ForeColor       =   &H00FF0000&
                     Height          =   285
                     Index           =   98
                     Left            =   2280
                     RightToLeft     =   -1  'True
                     TabIndex        =   314
                     Top             =   120
                     Width           =   1890
                  End
                  Begin VB.Label Label1 
                     Alignment       =   1  'Right Justify
                     BackColor       =   &H00E2E9E9&
                     BackStyle       =   0  'Transparent
                     Caption         =   "—ﬁ„ «·ﬁÌœ"
                     Height          =   195
                     Index           =   97
                     Left            =   2760
                     RightToLeft     =   -1  'True
                     TabIndex        =   313
                     Top             =   840
                     Width           =   990
                  End
               End
               Begin VB.Image Image2 
                  Height          =   3135
                  Left            =   120
                  Picture         =   "RSTradingCenter.frx":25E9D
                  Stretch         =   -1  'True
                  Top             =   120
                  Width           =   7380
               End
               Begin VB.Label Label11 
                  Alignment       =   2  'Center
                  BackStyle       =   0  'Transparent
                  Caption         =   "«·”« —Ì…"
                  BeginProperty Font 
                     Name            =   "Arial"
                     Size            =   27.75
                     Charset         =   178
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H00008000&
                  Height          =   5295
                  Left            =   240
                  TabIndex        =   198
                  Top             =   3600
                  Width           =   6615
               End
            End
            Begin VB.Label Label1 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Height          =   375
               Index           =   40
               Left            =   4590
               RightToLeft     =   -1  'True
               TabIndex        =   199
               Top             =   840
               Width           =   7605
            End
         End
         Begin C1SizerLibCtl.C1Elastic ELe 
            Height          =   7785
            Index           =   3
            Left            =   -18855
            TabIndex        =   200
            TabStop         =   0   'False
            Top             =   420
            Width           =   18210
            _cx             =   32120
            _cy             =   13732
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
            Begin VB.Frame Frame6 
               Height          =   3615
               Left            =   11880
               TabIndex        =   201
               Top             =   360
               Visible         =   0   'False
               Width           =   3615
               Begin VB.Frame Frame7 
                  Height          =   2295
                  Left            =   120
                  TabIndex        =   207
                  Top             =   1080
                  Width           =   1335
                  Begin ALLButtonS.ALLButton Command1 
                     Height          =   375
                     Index           =   3
                     Left            =   120
                     TabIndex        =   208
                     Top             =   240
                     Width           =   1095
                     _ExtentX        =   1931
                     _ExtentY        =   661
                     BTYPE           =   3
                     TX              =   " €ÌÌ— ’Ê—…"
                     ENAB            =   -1  'True
                     BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                        Name            =   "MS Sans Serif"
                        Size            =   8.25
                        Charset         =   178
                        Weight          =   400
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     COLTYPE         =   2
                     FOCUSR          =   -1  'True
                     BCOL            =   255
                     BCOLO           =   255
                     FCOL            =   16777215
                     FCOLO           =   16777215
                     MCOL            =   12632256
                     MPTR            =   1
                     MICON           =   "RSTradingCenter.frx":3655B
                     UMCOL           =   -1  'True
                     SOFT            =   0   'False
                     PICPOS          =   0
                     NGREY           =   0   'False
                     FX              =   0
                     HAND            =   0   'False
                     CHECK           =   0   'False
                     VALUE           =   0   'False
                  End
                  Begin ALLButtonS.ALLButton Command1 
                     Height          =   375
                     Index           =   4
                     Left            =   120
                     TabIndex        =   209
                     Top             =   720
                     Width           =   1095
                     _ExtentX        =   1931
                     _ExtentY        =   661
                     BTYPE           =   3
                     TX              =   " ﬂ»Ì—"
                     ENAB            =   -1  'True
                     BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                        Name            =   "MS Sans Serif"
                        Size            =   8.25
                        Charset         =   178
                        Weight          =   400
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     COLTYPE         =   2
                     FOCUSR          =   -1  'True
                     BCOL            =   8438015
                     BCOLO           =   8438015
                     FCOL            =   0
                     FCOLO           =   0
                     MCOL            =   12632256
                     MPTR            =   1
                     MICON           =   "RSTradingCenter.frx":36577
                     UMCOL           =   -1  'True
                     SOFT            =   0   'False
                     PICPOS          =   0
                     NGREY           =   0   'False
                     FX              =   0
                     HAND            =   0   'False
                     CHECK           =   0   'False
                     VALUE           =   0   'False
                  End
                  Begin ALLButtonS.ALLButton Command1 
                     Height          =   375
                     Index           =   5
                     Left            =   120
                     TabIndex        =   210
                     Top             =   1200
                     Width           =   1095
                     _ExtentX        =   1931
                     _ExtentY        =   661
                     BTYPE           =   3
                     TX              =   " ’€Ì—"
                     ENAB            =   -1  'True
                     BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                        Name            =   "MS Sans Serif"
                        Size            =   8.25
                        Charset         =   178
                        Weight          =   400
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     COLTYPE         =   2
                     FOCUSR          =   -1  'True
                     BCOL            =   8438015
                     BCOLO           =   8438015
                     FCOL            =   0
                     FCOLO           =   0
                     MCOL            =   12632256
                     MPTR            =   1
                     MICON           =   "RSTradingCenter.frx":36593
                     UMCOL           =   -1  'True
                     SOFT            =   0   'False
                     PICPOS          =   0
                     NGREY           =   0   'False
                     FX              =   0
                     HAND            =   0   'False
                     CHECK           =   0   'False
                     VALUE           =   0   'False
                  End
                  Begin ALLButtonS.ALLButton Command1 
                     Height          =   375
                     Index           =   6
                     Left            =   120
                     TabIndex        =   211
                     Top             =   1680
                     Width           =   1095
                     _ExtentX        =   1931
                     _ExtentY        =   661
                     BTYPE           =   3
                     TX              =   "œÊ—«‰"
                     ENAB            =   -1  'True
                     BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                        Name            =   "MS Sans Serif"
                        Size            =   8.25
                        Charset         =   178
                        Weight          =   400
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     COLTYPE         =   2
                     FOCUSR          =   -1  'True
                     BCOL            =   8438015
                     BCOLO           =   8438015
                     FCOL            =   0
                     FCOLO           =   0
                     MCOL            =   12632256
                     MPTR            =   1
                     MICON           =   "RSTradingCenter.frx":365AF
                     UMCOL           =   -1  'True
                     SOFT            =   0   'False
                     PICPOS          =   0
                     NGREY           =   0   'False
                     FX              =   0
                     HAND            =   0   'False
                     CHECK           =   0   'False
                     VALUE           =   0   'False
                  End
               End
               Begin VB.Frame Frame8 
                  Height          =   2295
                  Left            =   1800
                  TabIndex        =   202
                  Top             =   1080
                  Width           =   1335
                  Begin ALLButtonS.ALLButton Command1 
                     Height          =   375
                     Index           =   0
                     Left            =   120
                     TabIndex        =   203
                     Top             =   240
                     Width           =   1095
                     _ExtentX        =   1931
                     _ExtentY        =   661
                     BTYPE           =   3
                     TX              =   "ÃœÌœ"
                     ENAB            =   -1  'True
                     BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                        Name            =   "MS Sans Serif"
                        Size            =   8.25
                        Charset         =   178
                        Weight          =   400
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     COLTYPE         =   2
                     FOCUSR          =   -1  'True
                     BCOL            =   16711680
                     BCOLO           =   16711680
                     FCOL            =   16777215
                     FCOLO           =   16777215
                     MCOL            =   12632256
                     MPTR            =   1
                     MICON           =   "RSTradingCenter.frx":365CB
                     UMCOL           =   -1  'True
                     SOFT            =   0   'False
                     PICPOS          =   0
                     NGREY           =   0   'False
                     FX              =   0
                     HAND            =   0   'False
                     CHECK           =   0   'False
                     VALUE           =   0   'False
                  End
                  Begin ALLButtonS.ALLButton Command1 
                     Height          =   375
                     Index           =   1
                     Left            =   120
                     TabIndex        =   204
                     Top             =   720
                     Width           =   1095
                     _ExtentX        =   1931
                     _ExtentY        =   661
                     BTYPE           =   3
                     TX              =   "Õ›Ÿ"
                     ENAB            =   -1  'True
                     BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                        Name            =   "MS Sans Serif"
                        Size            =   8.25
                        Charset         =   178
                        Weight          =   400
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     COLTYPE         =   2
                     FOCUSR          =   -1  'True
                     BCOL            =   16711680
                     BCOLO           =   16711680
                     FCOL            =   16777215
                     FCOLO           =   16777215
                     MCOL            =   12632256
                     MPTR            =   1
                     MICON           =   "RSTradingCenter.frx":365E7
                     UMCOL           =   -1  'True
                     SOFT            =   0   'False
                     PICPOS          =   0
                     NGREY           =   0   'False
                     FX              =   0
                     HAND            =   0   'False
                     CHECK           =   0   'False
                     VALUE           =   0   'False
                  End
                  Begin ALLButtonS.ALLButton Command1 
                     Height          =   375
                     Index           =   2
                     Left            =   120
                     TabIndex        =   205
                     Top             =   1200
                     Width           =   1095
                     _ExtentX        =   1931
                     _ExtentY        =   661
                     BTYPE           =   3
                     TX              =   "Õ–›"
                     ENAB            =   -1  'True
                     BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                        Name            =   "MS Sans Serif"
                        Size            =   8.25
                        Charset         =   178
                        Weight          =   400
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     COLTYPE         =   2
                     FOCUSR          =   -1  'True
                     BCOL            =   255
                     BCOLO           =   8421631
                     FCOL            =   16777215
                     FCOLO           =   16777215
                     MCOL            =   12632256
                     MPTR            =   1
                     MICON           =   "RSTradingCenter.frx":36603
                     UMCOL           =   -1  'True
                     SOFT            =   0   'False
                     PICPOS          =   0
                     NGREY           =   0   'False
                     FX              =   0
                     HAND            =   0   'False
                     CHECK           =   0   'False
                     VALUE           =   0   'False
                  End
                  Begin ALLButtonS.ALLButton Command1 
                     Height          =   375
                     Index           =   7
                     Left            =   120
                     TabIndex        =   206
                     Top             =   1680
                     Width           =   1095
                     _ExtentX        =   1931
                     _ExtentY        =   661
                     BTYPE           =   3
                     TX              =   "ÿ»«⁄…"
                     ENAB            =   -1  'True
                     BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                        Name            =   "MS Sans Serif"
                        Size            =   8.25
                        Charset         =   178
                        Weight          =   400
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     COLTYPE         =   2
                     FOCUSR          =   -1  'True
                     BCOL            =   8454016
                     BCOLO           =   8454016
                     FCOL            =   0
                     FCOLO           =   0
                     MCOL            =   12632256
                     MPTR            =   1
                     MICON           =   "RSTradingCenter.frx":3661F
                     UMCOL           =   -1  'True
                     SOFT            =   0   'False
                     PICPOS          =   0
                     NGREY           =   0   'False
                     FX              =   0
                     HAND            =   0   'False
                     CHECK           =   0   'False
                     VALUE           =   0   'False
                  End
               End
               Begin MSAdodcLib.Adodc Adodc1 
                  Height          =   375
                  Left            =   240
                  Top             =   480
                  Width           =   3015
                  _ExtentX        =   5318
                  _ExtentY        =   661
                  ConnectMode     =   0
                  CursorLocation  =   3
                  IsolationLevel  =   -1
                  ConnectionTimeout=   15
                  CommandTimeout  =   30
                  CursorType      =   3
                  LockType        =   3
                  CommandType     =   8
                  CursorOptions   =   0
                  CacheSize       =   50
                  MaxRecords      =   0
                  BOFAction       =   0
                  EOFAction       =   0
                  ConnectStringType=   1
                  Appearance      =   1
                  BackColor       =   -2147483643
                  ForeColor       =   -2147483640
                  Orientation     =   0
                  Enabled         =   -1
                  Connect         =   ""
                  OLEDBString     =   ""
                  OLEDBFile       =   ""
                  DataSourceName  =   ""
                  OtherAttributes =   ""
                  UserName        =   ""
                  Password        =   ""
                  RecordSource    =   ""
                  Caption         =   "  Õ—Ìﬂ «·’Ê—"
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "MS Sans Serif"
                     Size            =   8.25
                     Charset         =   178
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  _Version        =   393216
               End
            End
            Begin VSFlex8UCtl.VSFlexGrid grdLoc 
               Height          =   4080
               Left            =   120
               TabIndex        =   254
               Top             =   1320
               Width           =   11325
               _cx             =   19976
               _cy             =   7197
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
               Cols            =   4
               FixedRows       =   1
               FixedCols       =   1
               RowHeightMin    =   300
               RowHeightMax    =   0
               ColWidthMin     =   0
               ColWidthMax     =   0
               ExtendLastCol   =   -1  'True
               FormatString    =   $"RSTradingCenter.frx":3663B
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
            Begin VB.Label Label26 
               Alignment       =   1  'Right Justify
               Caption         =   " ⁄œÌ·«  ”«»ﬁ… ⁄·Ì «—ﬁ«„ «·ÊÕœ« "
               Height          =   615
               Left            =   7920
               RightToLeft     =   -1  'True
               TabIndex        =   255
               Top             =   480
               Width           =   3135
            End
            Begin VB.Label Label1 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Height          =   375
               Index           =   36
               Left            =   4560
               RightToLeft     =   -1  'True
               TabIndex        =   212
               Top             =   840
               Width           =   7575
            End
         End
         Begin C1SizerLibCtl.C1Elastic ELe 
            Height          =   7785
            Index           =   17
            Left            =   45
            TabIndex        =   342
            TabStop         =   0   'False
            Top             =   420
            Width           =   18210
            _cx             =   32120
            _cy             =   13732
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
            Begin VB.TextBox XPTxtDiscountVal 
               Alignment       =   1  'Right Justify
               Height          =   345
               Left            =   12060
               TabIndex        =   384
               Top             =   2580
               Width           =   2235
            End
            Begin VB.ComboBox XPCboDiscountType 
               Height          =   315
               Left            =   15405
               Style           =   2  'Dropdown List
               TabIndex        =   383
               Top             =   2625
               Width           =   1245
            End
            Begin VB.TextBox txtSupCode 
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
               Height          =   270
               Left            =   15540
               MaxLength       =   50
               RightToLeft     =   -1  'True
               TabIndex        =   375
               Top             =   2130
               Width           =   825
            End
            Begin VB.TextBox TxtSearch 
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
               Left            =   15510
               MaxLength       =   50
               RightToLeft     =   -1  'True
               TabIndex        =   372
               Top             =   1110
               Width           =   825
            End
            Begin VB.TextBox txtNoteSerial11 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00FFFFC0&
               Height          =   315
               Left            =   14685
               Locked          =   -1  'True
               RightToLeft     =   -1  'True
               TabIndex        =   368
               Top             =   630
               Width           =   1680
            End
            Begin VB.Frame Fra_Header 
               BackColor       =   &H00FFFFFF&
               BorderStyle     =   0  'None
               Height          =   570
               Index           =   0
               Left            =   0
               RightToLeft     =   -1  'True
               TabIndex        =   357
               Top             =   0
               Width           =   20790
               Begin VB.TextBox Text2 
                  Alignment       =   1  'Right Justify
                  Appearance      =   0  'Flat
                  BackColor       =   &H000000FF&
                  Height          =   240
                  Index           =   0
                  Left            =   3030
                  RightToLeft     =   -1  'True
                  TabIndex        =   362
                  Top             =   510
                  Visible         =   0   'False
                  Width           =   945
               End
               Begin VB.TextBox TxtModFlg2 
                  Alignment       =   1  'Right Justify
                  Appearance      =   0  'Flat
                  BackColor       =   &H0000FF00&
                  Enabled         =   0   'False
                  Height          =   285
                  Index           =   1
                  Left            =   2580
                  RightToLeft     =   -1  'True
                  TabIndex        =   361
                  Text            =   "modflag"
                  Top             =   90
                  Visible         =   0   'False
                  Width           =   465
               End
               Begin VB.Frame Frame2 
                  BackColor       =   &H00E0E0E0&
                  BorderStyle     =   0  'None
                  Enabled         =   0   'False
                  Height          =   375
                  Index           =   0
                  Left            =   540
                  RightToLeft     =   -1  'True
                  TabIndex        =   358
                  Top             =   450
                  Visible         =   0   'False
                  Width           =   3105
                  Begin MSDataListLib.DataCombo DataCombo4 
                     CausesValidation=   0   'False
                     Height          =   315
                     Index           =   0
                     Left            =   -255
                     TabIndex        =   359
                     Tag             =   "„‰ ›÷·ﬂ √œŒ· —ﬁ„ «·ﬁ÷Ì…"
                     Top             =   -15
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
                     Index           =   101
                     Left            =   2160
                     RightToLeft     =   -1  'True
                     TabIndex        =   360
                     Top             =   45
                     Width           =   855
                  End
               End
               Begin MSComctlLib.ImageList GrdImageList2 
                  Index           =   0
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
                        Picture         =   "RSTradingCenter.frx":366F2
                        Key             =   "CompanyName"
                     EndProperty
                     BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                        Picture         =   "RSTradingCenter.frx":36A8C
                        Key             =   "Ser"
                     EndProperty
                     BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                        Picture         =   "RSTradingCenter.frx":36E26
                        Key             =   "Vac_Name"
                     EndProperty
                     BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                        Picture         =   "RSTradingCenter.frx":371C0
                        Key             =   "ShareCount"
                     EndProperty
                     BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                        Picture         =   "RSTradingCenter.frx":3755A
                        Key             =   "Dis_Count"
                     EndProperty
                     BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                        Picture         =   "RSTradingCenter.frx":378F4
                        Key             =   "Bouns"
                     EndProperty
                     BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                        Picture         =   "RSTradingCenter.frx":37C8E
                        Key             =   "SharesValue"
                     EndProperty
                     BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                        Picture         =   "RSTradingCenter.frx":38228
                        Key             =   "BuyValue"
                     EndProperty
                  EndProperty
               End
               Begin ImpulseButton.ISButton btn_Last 
                  Height          =   315
                  Index           =   1
                  Left            =   90
                  TabIndex        =   363
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
                  ButtonImage     =   "RSTradingCenter.frx":385C2
                  ColorButton     =   14871017
                  AcclimateGrayTones=   -1  'True
                  DrawFocusRectangle=   0   'False
                  DisabledImageExtraction=   0
                  DisabledImageStyle=   1
               End
               Begin ImpulseButton.ISButton btn_Next 
                  Height          =   315
                  Index           =   1
                  Left            =   555
                  TabIndex        =   364
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
                  ButtonImage     =   "RSTradingCenter.frx":3895C
                  ColorButton     =   14871017
                  DrawFocusRectangle=   0   'False
                  DisabledImageStyle=   1
               End
               Begin ImpulseButton.ISButton btn_Previous 
                  Height          =   315
                  Index           =   1
                  Left            =   1155
                  TabIndex        =   365
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
                  ButtonImage     =   "RSTradingCenter.frx":38CF6
                  ColorButton     =   14871017
                  DrawFocusRectangle=   0   'False
                  DisabledImageStyle=   1
               End
               Begin ImpulseButton.ISButton btn_First 
                  Height          =   315
                  Index           =   1
                  Left            =   1620
                  TabIndex        =   366
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
                  ButtonImage     =   "RSTradingCenter.frx":39090
                  ColorButton     =   14871017
                  DrawFocusRectangle=   0   'False
                  DisabledImageStyle=   1
               End
               Begin VB.Label Label1 
                  Alignment       =   1  'Right Justify
                  BackStyle       =   0  'Transparent
                  Caption         =   "«· Œ›Ì÷"
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
                  Index           =   102
                  Left            =   11340
                  RightToLeft     =   -1  'True
                  TabIndex        =   367
                  Top             =   90
                  Width           =   2640
               End
               Begin VB.Image Image3 
                  Height          =   390
                  Left            =   7560
                  Picture         =   "RSTradingCenter.frx":3942A
                  Stretch         =   -1  'True
                  Top             =   0
                  Width           =   525
               End
            End
            Begin MSDataListLib.DataCombo DCboUserName 
               Height          =   315
               Index           =   1
               Left            =   11400
               TabIndex        =   343
               Top             =   6630
               Width           =   3345
               _ExtentX        =   5900
               _ExtentY        =   556
               _Version        =   393216
               Enabled         =   0   'False
               Text            =   ""
               RightToLeft     =   -1  'True
            End
            Begin ImpulseButton.ISButton btn_New 
               Height          =   285
               Index           =   1
               Left            =   11970
               TabIndex        =   344
               Top             =   7320
               Width           =   975
               _ExtentX        =   1720
               _ExtentY        =   503
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
               ButtonImage     =   "RSTradingCenter.frx":3D092
               ColorButton     =   14871017
               DrawFocusRectangle=   0   'False
            End
            Begin ImpulseButton.ISButton btn_Save 
               Height          =   315
               Index           =   1
               Left            =   9735
               TabIndex        =   345
               Top             =   7290
               Width           =   975
               _ExtentX        =   1720
               _ExtentY        =   556
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
               ButtonImage     =   "RSTradingCenter.frx":3D42C
               ColorButton     =   14871017
               DrawFocusRectangle=   0   'False
            End
            Begin ImpulseButton.ISButton btn_Modify 
               Height          =   225
               Index           =   1
               Left            =   10710
               TabIndex        =   346
               Top             =   7320
               Width           =   960
               _ExtentX        =   1693
               _ExtentY        =   397
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
               ButtonImage     =   "RSTradingCenter.frx":3D7C6
               ColorButton     =   14871017
               DrawFocusRectangle=   0   'False
            End
            Begin ImpulseButton.ISButton Btn_Undo 
               Height          =   225
               Index           =   1
               Left            =   8760
               TabIndex        =   347
               Top             =   7320
               Width           =   975
               _ExtentX        =   1720
               _ExtentY        =   397
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
               ButtonImage     =   "RSTradingCenter.frx":3DB60
               ColorButton     =   14871017
               DrawFocusRectangle=   0   'False
            End
            Begin ImpulseButton.ISButton btn_Delete 
               Height          =   315
               Index           =   1
               Left            =   7905
               TabIndex        =   348
               Top             =   7290
               Width           =   855
               _ExtentX        =   1508
               _ExtentY        =   556
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
               ButtonImage     =   "RSTradingCenter.frx":3DEFA
               ColorButton     =   14871017
               DrawFocusRectangle=   0   'False
            End
            Begin ImpulseButton.ISButton Btn_Update 
               Height          =   345
               Index           =   1
               Left            =   9045
               TabIndex        =   349
               TabStop         =   0   'False
               ToolTipText     =   " ÕœÌÀ ﬁ«⁄œ… «·»Ì«‰« "
               Top             =   6960
               Visible         =   0   'False
               Width           =   1110
               _ExtentX        =   1958
               _ExtentY        =   609
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
               ButtonImage     =   "RSTradingCenter.frx":3E494
               ColorButton     =   14871017
               DrawFocusRectangle=   0   'False
            End
            Begin ImpulseButton.ISButton btn_Cancel 
               Height          =   315
               Index           =   1
               Left            =   3600
               TabIndex        =   350
               Top             =   7260
               Width           =   960
               _ExtentX        =   1693
               _ExtentY        =   556
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
               ButtonImage     =   "RSTradingCenter.frx":3E82E
               ColorButton     =   14871017
               DrawFocusRectangle=   0   'False
               DisabledImageStyle=   1
            End
            Begin ImpulseButton.ISButton Btn_Print 
               Height          =   360
               Index           =   1
               Left            =   6255
               TabIndex        =   351
               TabStop         =   0   'False
               ToolTipText     =   "ÿ»«⁄… «·»Ì«‰«  "
               Top             =   7230
               Width           =   1380
               _ExtentX        =   2434
               _ExtentY        =   635
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
               ButtonImage     =   "RSTradingCenter.frx":3EBC8
               ColorButton     =   14871017
               DrawFocusRectangle=   0   'False
               DisabledImageStyle=   1
            End
            Begin ImpulseButton.ISButton btn_Query 
               Height          =   390
               Index           =   1
               Left            =   4710
               TabIndex        =   352
               TabStop         =   0   'False
               ToolTipText     =   "(Ctrl+F)  ··»ÕÀ ≈÷€ÿ Â–« «·„› «Õ √Ê ≈÷€ÿ "
               Top             =   7200
               Width           =   1395
               _ExtentX        =   2461
               _ExtentY        =   688
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
               ButtonImage     =   "RSTradingCenter.frx":4542A
               ColorButton     =   14871017
               DrawFocusRectangle=   0   'False
            End
            Begin MSComCtl2.DTPicker XPDtbTrans 
               Height          =   300
               Left            =   11070
               TabIndex        =   369
               Top             =   630
               Width           =   1800
               _ExtentX        =   3175
               _ExtentY        =   529
               _Version        =   393216
               Format          =   146604033
               CurrentDate     =   38784
            End
            Begin MSDataListLib.DataCombo DcbIqara 
               Height          =   315
               Left            =   11910
               TabIndex        =   373
               Tag             =   "⁄›Ê« Ì—ÃÏ «Œ Ì«—√”„ «·⁄ﬁ«—"
               Top             =   1110
               Width           =   3555
               _ExtentX        =   6271
               _ExtentY        =   556
               _Version        =   393216
               Text            =   ""
               RightToLeft     =   -1  'True
            End
            Begin MSDataListLib.DataCombo DcbUnitNo 
               Height          =   315
               Left            =   11940
               TabIndex        =   376
               Tag             =   "⁄›Ê« Ì—ÃÏ «œŒ«· √”„ «·ÕÌ"
               Top             =   1650
               Visible         =   0   'False
               Width           =   2025
               _ExtentX        =   3572
               _ExtentY        =   556
               _Version        =   393216
               Style           =   2
               Text            =   ""
               RightToLeft     =   -1  'True
            End
            Begin MSDataListLib.DataCombo DcbUnitType 
               Height          =   315
               Left            =   15060
               TabIndex        =   377
               Tag             =   "⁄›Ê« Ì—ÃÏ «œŒ«· √”„ «·ÕÌ"
               Top             =   1650
               Visible         =   0   'False
               Width           =   1305
               _ExtentX        =   2302
               _ExtentY        =   556
               _Version        =   393216
               Style           =   2
               Text            =   ""
               RightToLeft     =   -1  'True
            End
            Begin MSDataListLib.DataCombo dcsupplier2 
               Height          =   315
               Left            =   11940
               TabIndex        =   378
               Tag             =   "⁄›Ê« Ì—ÃÏ «Œ Ì«—√”„ «·„«·ﬂ"
               Top             =   2130
               Width           =   3075
               _ExtentX        =   5424
               _ExtentY        =   556
               _Version        =   393216
               Enabled         =   0   'False
               Text            =   ""
               RightToLeft     =   -1  'True
            End
            Begin VSFlex8UCtl.VSFlexGrid GridInstallments2 
               Height          =   3120
               Left            =   570
               TabIndex        =   382
               Top             =   3240
               Width           =   17025
               _cx             =   30030
               _cy             =   5503
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
               Cols            =   42
               FixedRows       =   1
               FixedCols       =   1
               RowHeightMin    =   320
               RowHeightMax    =   0
               ColWidthMin     =   0
               ColWidthMax     =   0
               ExtendLastCol   =   -1  'True
               FormatString    =   $"RSTradingCenter.frx":457C4
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
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "ﬁÌ„…"
               Height          =   270
               Index           =   8
               Left            =   14220
               TabIndex        =   386
               Top             =   2670
               Width           =   1080
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "‰Ê⁄ «· Œ›Ì÷"
               Height          =   255
               Index           =   10
               Left            =   16575
               TabIndex        =   385
               Top             =   2625
               Width           =   1080
            End
            Begin VB.Label Label1 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               BackStyle       =   0  'Transparent
               Caption         =   " «·„«·ﬂ"
               Height          =   165
               Index           =   106
               Left            =   16500
               RightToLeft     =   -1  'True
               TabIndex        =   381
               Top             =   2130
               Width           =   810
            End
            Begin VB.Label Label1 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               BackStyle       =   0  'Transparent
               Caption         =   "‰Ê⁄ «·ÊÕœ…"
               Height          =   285
               Index           =   105
               Left            =   16320
               RightToLeft     =   -1  'True
               TabIndex        =   380
               Top             =   1650
               Visible         =   0   'False
               Width           =   990
            End
            Begin VB.Label Label1 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               BackStyle       =   0  'Transparent
               Caption         =   "—ﬁ„ «·ÊÕœ…"
               ForeColor       =   &H00FF0000&
               Height          =   285
               Index           =   104
               Left            =   13995
               RightToLeft     =   -1  'True
               TabIndex        =   379
               Top             =   1650
               Visible         =   0   'False
               Width           =   870
            End
            Begin VB.Label Label1 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               BackStyle       =   0  'Transparent
               Caption         =   "«·⁄ﬁ«—"
               Height          =   195
               Index           =   103
               Left            =   16215
               RightToLeft     =   -1  'True
               TabIndex        =   374
               Top             =   1110
               Width           =   990
            End
            Begin VB.Label lbl 
               Alignment       =   2  'Center
               BackColor       =   &H00E2E9E9&
               Caption         =   " «—ÌŒ «·”‰œ"
               Height          =   270
               Index           =   2
               Left            =   13140
               TabIndex        =   371
               Top             =   630
               Width           =   1410
            End
            Begin VB.Label Label5 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "—ﬁ„ «·”‰œ"
               Height          =   255
               Index           =   1
               Left            =   16650
               TabIndex        =   370
               Top             =   600
               Width           =   975
            End
            Begin VB.Label Label2 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "«·”Ã· «·Õ«·Ì:"
               Height          =   225
               Index           =   4
               Left            =   3735
               RightToLeft     =   -1  'True
               TabIndex        =   356
               Top             =   6945
               Width           =   1245
            End
            Begin VB.Label Label2 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "⁄œœ «·”Ã·« :"
               Height          =   225
               Index           =   3
               Left            =   1500
               RightToLeft     =   -1  'True
               TabIndex        =   355
               Top             =   6945
               Width           =   1395
            End
            Begin VB.Label LabCurr_Rec 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Height          =   225
               Index           =   1
               Left            =   2895
               RightToLeft     =   -1  'True
               TabIndex        =   354
               Top             =   6960
               Width           =   705
            End
            Begin VB.Label LabCount_Rec 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Height          =   225
               Index           =   1
               Left            =   810
               RightToLeft     =   -1  'True
               TabIndex        =   353
               Top             =   6960
               Width           =   555
            End
         End
      End
      Begin C1SizerLibCtl.C1Elastic EltCont 
         Height          =   1020
         Left            =   8205
         TabIndex        =   213
         TabStop         =   0   'False
         Top             =   8355
         Width           =   7230
         _cx             =   12753
         _cy             =   1799
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
            Left            =   5910
            TabIndex        =   214
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
            ButtonImage     =   "RSTradingCenter.frx":45EF3
            ColorButton     =   14871017
            DrawFocusRectangle=   0   'False
         End
         Begin ImpulseButton.ISButton btnSave 
            Height          =   330
            Left            =   4365
            TabIndex        =   215
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
            ButtonImage     =   "RSTradingCenter.frx":4628D
            ColorButton     =   14871017
            DrawFocusRectangle=   0   'False
         End
         Begin ImpulseButton.ISButton btnModify 
            Height          =   330
            Left            =   5175
            TabIndex        =   216
            Top             =   555
            Width           =   765
            _ExtentX        =   1349
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
            ButtonImage     =   "RSTradingCenter.frx":46627
            ColorButton     =   14871017
            DrawFocusRectangle=   0   'False
         End
         Begin ImpulseButton.ISButton BtnUndo 
            Height          =   330
            Left            =   3600
            TabIndex        =   217
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
            ButtonImage     =   "RSTradingCenter.frx":469C1
            ColorButton     =   14871017
            DrawFocusRectangle=   0   'False
         End
         Begin ImpulseButton.ISButton btnDelete 
            Height          =   330
            Left            =   2835
            TabIndex        =   218
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
            ButtonImage     =   "RSTradingCenter.frx":46D5B
            ColorButton     =   14871017
            DrawFocusRectangle=   0   'False
         End
         Begin ImpulseButton.ISButton btnQuery 
            Height          =   330
            Left            =   1935
            TabIndex        =   219
            TabStop         =   0   'False
            ToolTipText     =   "(Ctrl+F)  ··»ÕÀ ≈÷€ÿ Â–« «·„› «Õ √Ê ≈÷€ÿ "
            Top             =   570
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
            ButtonImage     =   "RSTradingCenter.frx":472F5
            ColorButton     =   14871017
            DrawFocusRectangle=   0   'False
         End
         Begin ImpulseButton.ISButton BtnUpdate 
            Height          =   330
            Left            =   6075
            TabIndex        =   220
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
            ButtonImage     =   "RSTradingCenter.frx":4768F
            ColorButton     =   14871017
            DrawFocusRectangle=   0   'False
         End
         Begin ImpulseButton.ISButton BtnPrint 
            Height          =   405
            Left            =   1125
            TabIndex        =   221
            TabStop         =   0   'False
            Top             =   570
            Width           =   780
            _ExtentX        =   1376
            _ExtentY        =   714
            ButtonStyle     =   1
            ButtonPositionImage=   1
            Caption         =   "ÿ»«⁄Â"
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
            ButtonImage     =   "RSTradingCenter.frx":47A29
            ColorButton     =   14871017
            DrawFocusRectangle=   0   'False
            DisabledImageStyle=   1
         End
         Begin ImpulseButton.ISButton btnCancel 
            Height          =   330
            Left            =   225
            TabIndex        =   222
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
            ButtonImage     =   "RSTradingCenter.frx":47DC3
            ColorButton     =   14871017
            DrawFocusRectangle=   0   'False
            DisabledImageStyle=   1
         End
      End
      Begin ImpulseButton.ISButton cmdPrint2 
         Height          =   405
         Left            =   4320
         TabIndex        =   321
         TabStop         =   0   'False
         Top             =   8880
         Width           =   2220
         _ExtentX        =   3916
         _ExtentY        =   714
         ButtonStyle     =   1
         ButtonPositionImage=   1
         Caption         =   "ÿ»«⁄Â ⁄ﬁœ  «·„«·ﬂ"
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
         ButtonImage     =   "RSTradingCenter.frx":4815D
         ColorButton     =   14871017
         DrawFocusRectangle=   0   'False
         DisabledImageStyle=   1
      End
      Begin ImpulseButton.ISButton Cmd 
         Height          =   375
         Index           =   11
         Left            =   6960
         TabIndex        =   340
         Top             =   8880
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   661
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
         ColorButton     =   14871017
         ColorHighlight  =   16777215
         ColorHoverText  =   16711680
         ColorShadow     =   -2147483637
         ColorOutline    =   0
         DrawFocusRectangle=   0   'False
         ColorToggledHoverText=   16711680
         ColorTextShadow =   -2147483637
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "«·”Ã· «·Õ«·Ì:"
         Height          =   210
         Index           =   0
         Left            =   2400
         RightToLeft     =   -1  'True
         TabIndex        =   226
         Top             =   9075
         Width           =   975
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "⁄œœ «·”Ã·« :"
         Height          =   210
         Index           =   1
         Left            =   690
         RightToLeft     =   -1  'True
         TabIndex        =   225
         Top             =   9075
         Width           =   990
      End
      Begin VB.Label LabCurrRec 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Height          =   210
         Left            =   1695
         RightToLeft     =   -1  'True
         TabIndex        =   224
         Top             =   9090
         Width           =   675
      End
      Begin VB.Label LabCountRec 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Height          =   210
         Left            =   120
         RightToLeft     =   -1  'True
         TabIndex        =   223
         Top             =   9075
         Width           =   540
      End
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      BackStyle       =   0  'Transparent
      Caption         =   "›«ﬂ”"
      Height          =   285
      Index           =   46
      Left            =   2400
      RightToLeft     =   -1  'True
      TabIndex        =   2
      Top             =   3240
      Width           =   1050
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      BackStyle       =   0  'Transparent
      Caption         =   "«·ÊﬂÌ·"
      Height          =   285
      Index           =   25
      Left            =   0
      RightToLeft     =   -1  'True
      TabIndex        =   1
      Top             =   840
      Width           =   1050
   End
End
Attribute VB_Name = "RSAkar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim RsSavRec As ADODB.Recordset
Dim BKGrndPic As ClsBackGroundPic
Dim RecId As String
Dim II As Long
Dim cSearch  As clsDCboSearch
Dim hijriorJerojian As Integer
Dim Account_Code_dynamic As String
Dim Account_Code_dynamic167 As String
Dim OwnerAccount As String
Dim vatPercetage As Double
Dim vaTAccount As String
Public mIndex  As Long




Private Sub lbl_MouseMove(Index As Integer, _
                          Button As Integer, _
                          Shift As Integer, _
                          X As Single, _
                          Y As Single)

    If val(lbl(Index).Caption) <> 0 Then
        lbl(Index).ToolTipText = WriteNo(lbl(Index).Caption, 0, True)
    End If

End Sub

Private Sub XPCboDiscountType_Change()
    XPCboDiscountType_Click
End Sub

Private Sub XPCboDiscountType_Click()
    On Error GoTo ErrTrap

    If XPCboDiscountType.ListIndex = 0 Or XPCboDiscountType.ListIndex = 3 Or XPCboDiscountType.ListIndex = -1 Then
    
        XPTxtDiscountVal.Enabled = False
        XPTxtDiscountVal.Text = ""
    Else
    
        XPTxtDiscountVal.Enabled = True
        XPTxtDiscountVal.Text = ""
    End If

    If Me.TxtModFlg.Text = "N" Or Me.TxtModFlg.Text = "E" Then
        If FG.TextMatrix(1, FG.ColIndex("Code")) <> "" Then
            NewGrid.Calculate 1, , , True
        End If
    End If

    Me.lbl(55).Visible = (Me.XPCboDiscountType.ListIndex = 2)

    'Me.lbl(21).Visible = (Me.XPCboDiscountType.ListIndex = 2)
    If XPCboDiscountType.ListIndex = 0 Then
        lbl(8).Visible = False
        XPTxtDiscountVal.Visible = False
        lbl(8).Visible = False
    Else
        lbl(8).Visible = True
        XPTxtDiscountVal.Visible = True
        lbl(8).Visible = True
    End If

    Exit Sub
ErrTrap:
End Sub

Private Sub XPTxtDiscountVal_Change()
    Dim Msg As String
    On Error GoTo ErrTrap

    If Me.TxtModFlg.Text = "N" Or Me.TxtModFlg.Text = "E" Then
        NewGrid.Calculate 1, , , True
    End If

    Exit Sub
ErrTrap:
End Sub

Private Sub XPTxtDiscountVal_KeyPress(KeyAscii As Integer)
    KeyAscii = KeyAscii_Num(KeyAscii, Me.XPTxtDiscountVal.Text, 0)
End Sub

Private Sub DcbUnitNo_Change()
Dim Dcombos As ClsDataCombos
Dim idd As Integer
   Set Dcombos = New ClsDataCombos

If val(DcbUnitType.BoundText) > 0 Then
idd = val(DcbUnitNo.BoundText)
Dim meterPrice As Double
Dim lengh As Double
Dim customerid As Integer
Dim rentType As Integer
Dim ElectAccount As String
Dim MiniRentValue As Double
Dim Typed As Integer
 Me.TxtRemarks = GetIqarUnitData(idd, , meterPrice, lengh, customerid, rentType, , , , , , ElectAccount, MiniRentValue, Typed)
 TxtElectAccount.Text = ElectAccount
 DcbRentType.ListIndex = IIf(rentType < 0, 0, rentType - 1)
 TxtMeterValue.Text = meterPrice
 TxtMeterCount.Text = lengh
 TxtMiniRentValue.Text = MiniRentValue
 If Typed = 1 Then
 ComResid(1).value = True
 Else
 ComResid(0).value = True
 End If
 ReLineGrid
 ' dcCustomer.BoundText = customerid
  
End If

If Me.TxtModFlg.Text = "N" Or Me.TxtModFlg.Text = "E" Then
    Dim s As String
    s = "Select * from TblIqarDiscountTrans2 Where UnitNo = " & val(DcbUnitNo.BoundText) & " and unittype = " & val(DcbUnitType.BoundText)
    s = s & " and Iqar = " & val(DcbIqara.BoundText) '& " and BranchID = " & val(Dcbranch.BoundText)
    Dim rsDummy As New ADODB.Recordset
    rsDummy.Open s, Cn, adOpenStatic, adLockReadOnly
    If Not rsDummy.EOF Then
        txtDiscountPercent.Text = rsDummy!DiscountPercent & ""
        txtDiscountPercent.Tag = rsDummy!DiscountPercent & ""
    End If
End If
End Sub

Private Sub DcbUnitNo_Click(Area As Integer)
DcbUnitNo_Change
End Sub

Private Sub DcbUnitNo_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyF5 Then
ReloadCombos
End If
End Sub

Private Sub DcbUnitType_Change()
ReloadUonit
End Sub


Private Sub DcbUnitType_Click(Area As Integer)
DcbUnitType_Change
End Sub

Private Sub DcbUnitType_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyF5 Then
ReloadCombos
End If
End Sub

Private Sub dcsupplier_Change()
dcsupplier_Click (0)
End Sub

Private Sub dcsupplier2_Click(Area As Integer)
   If val(dcsupplier2.BoundText) = 0 Then Exit Sub

    Dim EmpCode  As String
 
    GetTblCustemersCode , , dcsupplier2.BoundText, EmpCode
    Me.Text1(0).Text = EmpCode
    'ClculteVAT
End Sub

Private Sub txtSupCode_KeyPress(KeyAscii As Integer)
 Dim EmpID As Integer

    If KeyAscii = vbKeyReturn Then
        GetTblCustemersCode txtSupCode.Text, EmpID, , , 57
        dcsupplier2.BoundText = EmpID
    End If
End Sub


Private Sub DcbIqara_Change()
DcbUnitType_Change
DcbIqara_Click (0)
Calculte
End Sub

Private Sub DcbIqara_Click(Area As Integer)
      If val(DcbIqara.BoundText) = 0 Then dcsupplier.BoundText = 0: Exit Sub

    Dim EmpCode  As String
 Dim ownerid As Double
    GetIqarCode , , DcbIqara.BoundText, EmpCode, ownerid
    
    Me.TxtSearch.Text = EmpCode
    dcsupplier.BoundText = ownerid
    Calculte
    'DcbUnitType_Change
End Sub

Private Sub DcbIqara_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyF3 Then

Load FrmAqarSearch
FrmAqarSearch.m_RetrunType = 1
FrmAqarSearch.show


End If


If KeyCode = vbKeyF5 Then
ReloadCombos
End If

End Sub

Private Sub TxtSearch_KeyPress(KeyAscii As Integer)
  Dim EmpID As Double
'GetTblCustemersCode
    If KeyAscii = vbKeyReturn Then
        GetIqarCode TxtSearch.Text, EmpID
        DcbIqara.BoundText = EmpID
        DcbIqara_Click (0)
    End If
End Sub



Private Sub CmdCus_Click()


If Me.TxtModFlg.Text <> "R" Then
RSOwner.Index = 1
Load RSOwner
RSOwner.show
End If

End Sub

Private Sub cmdPaymant_Click()
    If checkApility("FrmCashing1") = False Then
                Exit Sub
            End If

Load FrmPayments2
FrmPayments2.show
FrmPayments2.newrecord
FrmPayments2.DCboCashType.ListIndex = 8
FrmPayments2.DBCboClientName.BoundText = dcsupplier.BoundText
FrmPayments2.DcbIqara2.BoundText = val(TxtAqarid)
FrmPayments2.DBCboClientName_Change
 'FrmPayments2.TxtContNo.Text = val(TxtContNo.Text)
 ' FrmPayments2.TxtContractNo.Text = (txtNoteSerial1.Text)
  

End Sub

Private Sub cmdPrint2_Click()


     
    Dim MySQL As String
    Dim RsData As New ADODB.Recordset
    Dim xApp As New CRAXDRT.Application
    Dim xReport As CRAXDRT.Report
    Dim CViewer As ClsReportViewer
    Dim StrReportTitle As String
    Dim StrFileName As String
      Dim StrSQL As String
    Dim Msg As String

    
    MySQL = "SELECT   isnull(dbo.TblAqar.NOOFYears,1)  as NOOFYears, TblAqar.TypeDate,   dbo.TblAqar.Aqarid, dbo.TblAqar.aqarNo, dbo.TblAqar.aqartypeid, dbo.tblAkarType.name, dbo.tblAkarType.namee, dbo.TblAqar.CountryID, "
    MySQL = MySQL & "  tblaqar.FromPlanneddate , tblaqar.FromPlanneddateH, tblaqar.ToPlanneddate, tblaqar.ToPlanneddateH,tblaqar.DateHCont,tblaqar.DateCont,tblaqar.ContValue,tblaqar.Telephone,"
    MySQL = MySQL & "   tblaqar.PlotNo , tblaqar.Planned, tblaqar.PlotNo, tblaqar.DisountAmount,"
    
    MySQL = MySQL & "                       dbo.TblCountriesData.CountryName, dbo.TblAqar.cityid, dbo.TblCountriesGovernments.GovernmentName, dbo.TblAqar.heyid,"
    MySQL = MySQL & "                       dbo.TblCountriesGovernmentsCities.CityName, dbo.TblAqar.streetname, dbo.TblAqar.schemeid, dbo.tblSchemes.name AS SchemeName,"
    MySQL = MySQL & "                       dbo.tblSchemes.namee AS SchemeNameE, dbo.TblAqar.aqarage, dbo.TblAqar.currentPrice, dbo.TblAqar.lastrentvalue, dbo.TblAqar.maintenancetypeid,"
    MySQL = MySQL & "                       dbo.TblAqar.StatusId, dbo.TblAqar.EntryCount, dbo.TblAqar.floorcount, dbo.TblAqar.noofoffices, dbo.TblAqar.noofparking, dbo.TblAqar.interfaceid,"
    MySQL = MySQL & "                       dbo.TblAqar.noofapartement, dbo.TblAqar.totallength, dbo.TblAqar.meterRentvalue, dbo.TblAqar.Rate, dbo.TblAqar.Price, dbo.TblAqar.ownerid,"
    MySQL = MySQL & "                       TblCustemers_1.CusName, TblCustemers_1.CusNamee, dbo.TblAqar.Location, dbo.TblAqar.aqarname, dbo.TblAqar.northlength, dbo.TblAqar.eastlength,"
    MySQL = MySQL & "                       dbo.TblAqar.Southlength, dbo.TblAqar.Westlength, dbo.TblAqar.metersalevalue, dbo.TblAqar.GoogleMap, dbo.TblAqar.suckno, dbo.TblAqar.authorizationname,"
    MySQL = MySQL & "                       dbo.TblAqar.suckdateH, dbo.TblAqar.suckdate, dbo.TblAqar.statusdate, dbo.TblAqar.PriceHadW, dbo.TblAqar.PriceSomW, dbo.TblAqar.StreetNo, dbo.TblAqar.Part,"
    MySQL = MySQL & "                       dbo.TblAqar.UnitNo, TblAkarUnit_1.name AS UnitName, TblAkarUnit_1.namee AS UnitNamee, dbo.TblAqar.Block, dbo.TblAqar.PriceSom, dbo.TblAqar.PriceHad,"
    MySQL = MySQL & "                       dbo.TblAqarDetai.Aqarid AS AqaridD, dbo.TblAqarDetai.length, dbo.TblAqarDetai.namerentType, dbo.TblAqarDetai.unittype, TblAkarUnit_1.name AS nameD,"
    MySQL = MySQL & "                       TblAkarUnit_1.namee AS nameeD, dbo.TblAqarDetai.customerid, TblCustemers_1.CusName AS CusNameD, TblCustemers_1.CusNamee AS CusNameeD,"
    MySQL = MySQL & "                       TblCustemers_1.Fullcode, dbo.TblAqarDetai.rentType, dbo.TblAqarDetai.meterPrice, dbo.TblAqarDetai.roomscount, dbo.TblAqarDetai.WCcount,"
    MySQL = MySQL & "                       dbo.TblAqarDetai.kithchencount, dbo.TblAqarDetai.RentValue, dbo.TblAqarDetai.haveFurniture,dbo.TblAqarDetai.isTax, dbo.TblAqarDetai.unitdesc, dbo.TblAqarDetai.unitno AS unitnoD,"
    MySQL = MySQL & "                       dbo.TblAqarDetai2.Aqarid AS AqaridD2, dbo.TblAqarDetai2.MainCo, dbo.TblAqarDetai2.Elevatortype, dbo.TblAqarDetai2.BuildNo, dbo.TblAqarDetai2.company,"
    MySQL = MySQL & "                       dbo.TblAqarDetai2.ElevatorNo, dbo.TblAqarDetai2.MaintStrDate, dbo.TblAqarDetai2.MaintEndDate, dbo.TblAqarDetai3.Aqarid AS AqaridD3, dbo.TblAqarDetai3.Waset,"
    MySQL = MySQL & "                       dbo.TblAqarDetai3.Rate AS RateD, dbo.TblAqar.eastWriiten, dbo.TblAqar.westWriiten, dbo.TblAqarDetai2.Id AS idd2, dbo.TblAqarDetai3.Id AS idd3,"
    MySQL = MySQL & "                       dbo.TblAqarDetai.Id AS IdD, dbo.TblAqarDetai.Floor, dbo.TblAqarDetai.LoungeCount, dbo.TblAqarDetai.ACCount, dbo.TblAqarDetai.UnitElectric,"
    MySQL = MySQL & "                       dbo.TblAqarDetai.ACCountspleat, dbo.TblAqarDetai.electric, dbo.TblAqarDetai.Status, dbo.TblAqarDetai.Services, dbo.TblAqarDetai.Water,"
    MySQL = MySQL & "                       TblCustemers_2.CusName AS CusNameOwen, TblCustemers_2.CusNamee AS CusNameOwenE"
    MySQL = MySQL & "  FROM         dbo.TblAqarDetai3 RIGHT OUTER JOIN"
    MySQL = MySQL & "                       dbo.TblAqar ON dbo.TblAqarDetai3.Aqarid = dbo.TblAqar.Aqarid LEFT OUTER JOIN"
    MySQL = MySQL & "                       dbo.TblAqarDetai2 ON dbo.TblAqar.Aqarid = dbo.TblAqarDetai2.Aqarid LEFT OUTER JOIN"
    MySQL = MySQL & "                       dbo.TblAqarDetai LEFT OUTER JOIN"
    MySQL = MySQL & "                       dbo.TblCustemers TblCustemers_1 ON dbo.TblAqarDetai.customerid = TblCustemers_1.CusID LEFT OUTER JOIN"
    MySQL = MySQL & "                       dbo.TblAkarUnit TblAkarUnit_1 ON dbo.TblAqarDetai.unittype = TblAkarUnit_1.id ON dbo.TblAqar.Aqarid = dbo.TblAqarDetai.Aqarid LEFT OUTER JOIN"
    MySQL = MySQL & "                       dbo.TblAkarUnit TblAkarUnit_2 ON dbo.TblAqar.UnitNo = TblAkarUnit_2.id LEFT OUTER JOIN"
    MySQL = MySQL & "                       dbo.TblCustemers TblCustemers_2 ON dbo.TblAqar.ownerid = TblCustemers_2.CusID LEFT OUTER JOIN"
    MySQL = MySQL & "                       dbo.tblSchemes ON dbo.TblAqar.schemeid = dbo.tblSchemes.id LEFT OUTER JOIN"
    MySQL = MySQL & "                       dbo.TblCountriesGovernments INNER JOIN"
    MySQL = MySQL & "                       dbo.TblCountriesGovernmentsCities ON dbo.TblCountriesGovernments.GovernmentID = dbo.TblCountriesGovernmentsCities.GovernmentID INNER JOIN"
    MySQL = MySQL & "                       dbo.TblCountriesData ON dbo.TblCountriesGovernments.CountryID = dbo.TblCountriesData.CountryID ON"
    MySQL = MySQL & "                       dbo.TblAqar.heyid = dbo.TblCountriesGovernmentsCities.CityID AND dbo.TblAqar.cityid = dbo.TblCountriesGovernments.GovernmentID AND"
    MySQL = MySQL & "                       dbo.TblAqar.CountryID = dbo.TblCountriesData.CountryID LEFT OUTER JOIN"
    MySQL = MySQL & "                       dbo.tblAkarType ON dbo.TblAqar.aqartypeid = dbo.tblAkarType.id"
    MySQL = MySQL & "  Where (dbo.TblAqar.Aqarid =" & val(Me.TxtAqarid.Text) & " )"
StrSQL = MySQL
    If SystemOptions.UserInterface = ArabicInterface Then
        StrFileName = App.path & "\REPORTS\REPORTS NEW\" & "ReportIqar3.rpt"
    Else
        StrFileName = App.path & "\REPORTS\REPORTS NEW\" & "ReportIqar3.rpt"
    End If

    If Dir(StrFileName) = "" Then
        'GetMsgs 139, vbExclamation
        Screen.MousePointer = vbDefault
        Exit Sub
    End If

    Set RsData = New ADODB.Recordset
    RsData.Open MySQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

    If RsData.BOF Or RsData.EOF Then
        'GetMsgs 138, vbExclamation
        Msg = "?CE??I E?C?CE ?????"
        MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
        RsData.Close
        Set RsData = Nothing
        Screen.MousePointer = vbDefault
        Exit Sub
    End If

    Screen.MousePointer = vbArrowHourglass
    Set xReport = xApp.OpenReport(StrFileName)
    xReport.Database.SetDataSource RsData

    
    MySQL = ""
    MySQL = " SELECT  * from TblAqrOwin"
    MySQL = MySQL & " WHERE     (AqrID =" & val(Me.TxtAqarid.Text) & ")"
    
    
    
  
      Set RsData = New ADODB.Recordset
      
       
      RsData.Open MySQL, Cn, adOpenStatic, adLockReadOnly, adCmdText
      xReport.OpenSubreport("aa").Database.SetDataSource RsData
  
 
    'rs1.Open My_SQL, Cn, adOpenStatic, adLockReadOnly, adCmdText
    
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
    xReport.reporttitle = StrReportTitle
    xReport.EnableParameterPrompting = False
    xReport.ApplicationName = App.title
    xReport.ReportAuthor = App.title
    Set CViewer = New ClsReportViewer
    CViewer.FireReport xReport, WindowTarget, "", , , , StrFileName, , StrSQL

    RsData.Close
    Set RsData = Nothing
    Screen.MousePointer = vbDefault





End Sub

Private Sub Command11_Click()
Dim StrSQL As String
Dim des As String

 

'If checkAllocations(val(TxtContNo), des) = True Then
'MsgBox "·« Ì„ﬂ‰ «· ⁄œÌ·  ·ÊÃÊœ Õ—ﬂ«  «À»«  «” Õﬁ«ﬁ ⁄·Ì Â–« «·⁄ﬁœ ÊÂÌ ﬂ«· «·Ì " & CHR(13) & des
'Exit Sub
'End If





       StrSQL = "Delete From DOUBLE_ENTREY_VOUCHERS Where Notes_ID=" & val(Me.TXTNoteID.Text)
       Cn.Execute StrSQL, , adExecuteNoRecords
        StrSQL = "Delete From Notes Where NoteID=" & val(Me.TXTNoteID.Text)
       Cn.Execute StrSQL, , adExecuteNoRecords
       Cn.Execute "Update tblaqar set NoteID=null,NoteSerial=null where Aqarid=" & val(Me.TxtAqarid.Text) & " "
       
TxtNoteSerial.Text = ""
TXTNoteID.Text = 0
MsgBox " „ Õ–› «·ﬁÌœ"
RsSavRec.Resync adAffectCurrent

End Sub
Function CheckAcconts() As Boolean
CheckAcconts = False
    Account_Code_dynamic167 = get_account_code_branch(167, my_branch)
     
     If val(TxtContValue) <> 0 Then
          
           If Account_Code_dynamic167 = "NO branch" Then
                          MsgBox "·„ Ì „ «‰‘«¡ «·›—⁄", vbCritical
                        Exit Function
            End If

            If Account_Code_dynamic167 = "NO account" Then
                             MsgBox "»—Ã«¡ —»ÿ Õ”«» „’—Ê›«  «·«ÌÃ«—«   «·„” Õﬁ… ··„·«ﬂ", vbCritical
                
                        Exit Function
             End If
            
             
     End If
     OwnerAccount = ""
          OwnerAccount = GetMyAccountCode("TblCustemers", "CusID", val(Me.dcsupplier.BoundText))
     

            If OwnerAccount = "" Then
                             MsgBox "Õ”«» «·„«·ﬂ €Ì— „⁄—›", vbCritical
                
                       Exit Function
             End If
             
      CheckAcconts = True
      
End Function
Function createVoucher()
 
Dim NoteID As Long
Dim NoteDate As Date
Dim NoteSerial As String
Dim Notevalue As Double
Dim des As String
des = "⁄ﬁœ «ÌÃ«— —ﬁ„ " & TxtNoteSerial & " „‰ «·„«·ﬂ " & dcsupplier.Text
 des = des & "   «·⁄ﬁ«— " & txtaqarname
If RdRTypeDate(0).value = True Then
des = des & "   «·› —… „‰  " & FromCotDateH.value & " «·Ì " & ToCotDateH.value
Else
des = des & " «·„Ê«›ﬁ " & FromCotDate.value & " «·Ì " & ToCotDate.value
End If

des = des & " " & Me.TxtRemarks.Text

Dim tablename As String
Dim Filedname As String
Dim ContNo As Long
Dim VATValue As Double
Dim sql As String
tablename = "tblaqar"
Filedname = "Aqarid"
ContNo = val(Me.TxtAqarid)
Notevalue = val(TxtContValue)
Dim i As Double
 'If CheckGLYearly.value = vbChecked Then
 'Notevalue = Notevalue / IIf(val(NOOFYears) = 0, 1, NOOFYears)
 'End If
 
 
 
   With Me.GridInstallments
      Notevalue = val(.TextMatrix(1, .ColIndex("valuewithout")))
     VATValue = val(.TextMatrix(1, .ColIndex("VatValue")))
End With

If Notevalue > 0 Then

                    If Me.TxtModFlg = "N" Then
                    
                          CreateNotes NoteID, (FristPaymentDate.value), val(dcBranch.BoundText), 29802, Notevalue + VATValue, NoteSerial, val(TxtAqarid), tablename, Filedname, ContNo, des, ToHijriDate(FirstInstallDateH.value)  'RecorddateH.value
                                  TXTNoteID.Text = NoteID
                                         TxtNoteSerial.Text = NoteSerial
                         Else
                                     If TXTNoteID.Text = "" Or TxtNoteSerial.Text = "" Then
                                CreateNotes NoteID, (DateCont.value), val(dcBranch.BoundText), 29802, Notevalue + VATValue, NoteSerial, val(TxtAqarid), tablename, Filedname, ContNo, des, ToHijriDate(DateHCont.value)
                                                     TXTNoteID.Text = NoteID
                                                    TxtNoteSerial.Text = NoteSerial
                                       Else
                                                     sql = "update notes  set Note_Value=" & Notevalue + VATValue & ",note_value_by_characters='" & WriteNo(val(Notevalue + VATValue), 0, True) & "'"
                                                    sql = sql & ",NoteSerial1='" & val(TxtAqarid) & "'"
                                                       sql = sql & " where NoteID=" & val(TXTNoteID.Text)
                                                       Cn.Execute sql
                                                   
                                     End If
                           
                    End If
 
      PercentgValueAddedAccount_Transec FristPaymentDate.value, 52, 0, vaTAccount, vatPercetage
                             AccountVat.BoundText = vaTAccount

CREATE_VOUCHER_GE val(TXTNoteID.Text), val(dcBranch.BoundText), user_id, val(Notevalue), Account_Code_dynamic167, OwnerAccount, des, FristPaymentDate.value, vaTAccount, , VATValue
RsSavRec.Resync adAffectCurrent
MsgBox " „ «‰‘«¡ «·ﬁÌœ", vbInformation
End If
End Function




Sub ClculteVAT(Optional Manulavat As Double = 0)
If Me.TxtModFlg.Text <> "R" Then
 
TxtFATYou.Text = 0


              If ComResid(1).value = True Then
                
       VBA.Calendar = vbCalGreg
                            PercentgValueAddedAccount_Transec FristPaymentDate.value, 52, 0, vaTAccount, vatPercetage
                             AccountVat.BoundText = vaTAccount
               
                           
                      Else
                       vatPercetage = 0
                End If


       TxtFATYou = vatPercetage
If val(Me.txtManulaVat) <> 0 Then
TxtFATYou = txtManulaVat
End If
                             
                             
                             TxtFATValue = val(TxtContValueWithout) * TxtFATYou / 100
                             
                             TxtContValue = val(TxtFATValue) + val(TxtContValueWithout)
                             
                             
End If


End Sub



Private Sub Command6_Click()


If Me.TxtModFlg.Text = "R" Then
Unload FrmSanadatOFContract
Load FrmSanadatOFContract
FrmSanadatOFContract.Indx = 2
FrmSanadatOFContract.Label1(0).Caption = TxtAqarid.Text
FrmSanadatOFContract.TxtNotID.Text = val(TxtAqarid.Text)
FrmSanadatOFContract.TxtContNo.Text = val(TxtAqarid.Text)
FrmSanadatOFContract.show
End If

End Sub

Private Sub ComResid_Click(Index As Integer)
txtManulaVat.Text = 0
ClculteVAT
End Sub

Private Sub RdRTypeDate_Click(Index As Integer)
 hideme
CalcContractIntervalAuto
End Sub
Function CalcContractIntervalAuto()
If Me.TxtModFlg = "R" Then Exit Function
If RdRTypeDate(0).value = True Then 'ÂÃ—Ì
  VBA.Calendar = vbCalHijri
 
       VBA.Calendar = vbCalGreg
 
       hijriorJerojian = 0
  Else '„Ì·«œÌ
   
       hijriorJerojian = 1
End If

End Function


Private Sub Command12_Click()
    
   If DoPremis(Do_Edit, Me.Name, True) = False Then
      Exit Sub
    End If

    On Error GoTo ErrTrap
 
    If ChekClodePeriod(DateCont.value) = True Then
        If SystemOptions.UserInterface = ArabicInterface Then
            MsgBox "Ì—ÃÏ  €ÌÌ—  «—ÌŒ «·Õ—ﬂ… ·«‰ Â–Â «·› —… „€·ﬁ…"
        Else
            MsgBox "Please Change Date Becouse This is Period is Closed"
        End If
        Exit Sub
    End If
    
    
    

If TxtNoteSerial.Text <> "" Then
If SystemOptions.UserInterface = ArabicInterface Then
MsgBox "Ì—ÃÏ Õ–› «·ﬁÌœ «Ê·«"
Exit Sub
Else
MsgBox "Please Delete JE"
End If
Exit Sub
End If

   
    Dim StrSQL As String
    
 

    StrSQL = "Delete From DOUBLE_ENTREY_VOUCHERS Where Notes_ID=" & val(Me.TXTNoteID.Text)
    Cn.Execute StrSQL, , adExecuteNoRecords
    If CheckAcconts = False Then Exit Sub
    createVoucher
     TxtModFlg = "R"



  '  SendMessage 1
    Exit Sub
ErrTrap:
Dim Msg As String
    Select Case Err.Number
        Case -2147467259
            'Could not update; currently locked.
            Msg = "⁄›Ê«" & CHR(13)
            Msg = Msg & " ·«Ì„ﬂ‰  ⁄œÌ· Â–« «·”Ã· ›Ï «·Êﬁ  «·Õ«·Ï" & CHR(13)
            Msg = Msg & "ÕÌÀ «‰Â ﬁÌœ «· ⁄œÌ· „‰ ﬁ»· „” Œœ„ «Œ— ⁄·Ï «·‘»ﬂ…"
            MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading + vbOKOnly, App.title
        
            If RsSavRec.EditMode <> adEditNone Then
                RsSavRec.CancelUpdate
                'RsSavRec.Requery
            End If
    End Select


End Sub

Private Sub Command8_Click()
Dim StrTempAccountCode As String
                   StrTempAccountCode = GetMyAccountCode("TblCustemers", "CusID", val(Me.dcsupplier.BoundText))
 
            ShowReport StrTempAccountCode, dcsupplier.Text, FrmDate.value, todate.value

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
    
    MySQL = "SELECT   dbo.TblAqar.NOOFYears,     dbo.TblAqar.TypeDate,      dbo.TblAqar.Aqarid, dbo.TblAqar.aqarNo, dbo.TblAqar.aqartypeid, dbo.tblAkarType.name, dbo.tblAkarType.namee, dbo.TblAqar.CountryID, "
    MySQL = MySQL & "                       dbo.TblCountriesData.CountryName, dbo.TblAqar.cityid, dbo.TblCountriesGovernments.GovernmentName, dbo.TblAqar.heyid,"
    MySQL = MySQL & "                       dbo.TblCountriesGovernmentsCities.CityName, dbo.TblAqar.streetname, dbo.TblAqar.schemeid, dbo.tblSchemes.name AS SchemeName,"
    MySQL = MySQL & "                       dbo.tblSchemes.namee AS SchemeNameE, dbo.TblAqar.aqarage, dbo.TblAqar.currentPrice, dbo.TblAqar.lastrentvalue, dbo.TblAqar.maintenancetypeid,"
    MySQL = MySQL & "                       dbo.TblAqar.StatusId, dbo.TblAqar.EntryCount, dbo.TblAqar.floorcount, dbo.TblAqar.noofoffices, dbo.TblAqar.noofparking, dbo.TblAqar.interfaceid,"
    MySQL = MySQL & "                       dbo.TblAqar.noofapartement, dbo.TblAqar.totallength, dbo.TblAqar.meterRentvalue, dbo.TblAqar.Rate, dbo.TblAqar.Price, dbo.TblAqar.ownerid,"
    MySQL = MySQL & "                       TblCustemers_1.CusName, TblCustemers_1.CusNamee, dbo.TblAqar.Location, dbo.TblAqar.aqarname, dbo.TblAqar.northlength, dbo.TblAqar.eastlength,"
    MySQL = MySQL & "                       dbo.TblAqar.Southlength, dbo.TblAqar.Westlength, dbo.TblAqar.metersalevalue, dbo.TblAqar.GoogleMap, dbo.TblAqar.suckno, dbo.TblAqar.authorizationname,"
    MySQL = MySQL & "                       dbo.TblAqar.suckdateH, dbo.TblAqar.suckdate, dbo.TblAqar.statusdate, dbo.TblAqar.PriceHadW, dbo.TblAqar.PriceSomW, dbo.TblAqar.StreetNo, dbo.TblAqar.Part,"
    MySQL = MySQL & "                       dbo.TblAqar.UnitNo, TblAkarUnit_1.name AS UnitName, TblAkarUnit_1.namee AS UnitNamee, dbo.TblAqar.Block, dbo.TblAqar.PriceSom, dbo.TblAqar.PriceHad,"
    MySQL = MySQL & "                       dbo.TblAqarDetai.Aqarid AS AqaridD, dbo.TblAqarDetai.length, dbo.TblAqarDetai.namerentType, dbo.TblAqarDetai.unittype, TblAkarUnit_1.name AS nameD,"
    MySQL = MySQL & "                       TblAkarUnit_1.namee AS nameeD, dbo.TblAqarDetai.customerid, TblCustemers_1.CusName AS CusNameD, TblCustemers_1.CusNamee AS CusNameeD,"
    MySQL = MySQL & "                       TblCustemers_1.Fullcode, dbo.TblAqarDetai.rentType, dbo.TblAqarDetai.meterPrice, dbo.TblAqarDetai.roomscount, dbo.TblAqarDetai.WCcount,"
    MySQL = MySQL & "                       dbo.TblAqarDetai.kithchencount, dbo.TblAqarDetai.RentValue, dbo.TblAqarDetai.haveFurniture,dbo.TblAqarDetai.isTax, dbo.TblAqarDetai.unitdesc, dbo.TblAqarDetai.unitno AS unitnoD,"
    MySQL = MySQL & "                       dbo.TblAqarDetai2.Aqarid AS AqaridD2, dbo.TblAqarDetai2.MainCo, dbo.TblAqarDetai2.Elevatortype, dbo.TblAqarDetai2.BuildNo, dbo.TblAqarDetai2.company,"
    MySQL = MySQL & "                       dbo.TblAqarDetai2.ElevatorNo, dbo.TblAqarDetai2.MaintStrDate, dbo.TblAqarDetai2.MaintEndDate, dbo.TblAqarDetai3.Aqarid AS AqaridD3, dbo.TblAqarDetai3.Waset,"
    MySQL = MySQL & "                       dbo.TblAqarDetai3.Rate AS RateD, dbo.TblAqar.eastWriiten, dbo.TblAqar.westWriiten, dbo.TblAqarDetai2.Id AS idd2, dbo.TblAqarDetai3.Id AS idd3,"
    MySQL = MySQL & "                       dbo.TblAqarDetai.Id AS IdD, dbo.TblAqarDetai.Floor, dbo.TblAqarDetai.LoungeCount, dbo.TblAqarDetai.ACCount, dbo.TblAqarDetai.UnitElectric,"
    MySQL = MySQL & "                       dbo.TblAqarDetai.ACCountspleat, dbo.TblAqarDetai.electric, dbo.TblAqarDetai.Status, dbo.TblAqarDetai.Services, dbo.TblAqarDetai.Water,"
    MySQL = MySQL & "                       TblCustemers_2.CusName AS CusNameOwen, TblCustemers_2.CusNamee AS CusNameOwenE"
    MySQL = MySQL & "  FROM         dbo.TblAqarDetai3 RIGHT OUTER JOIN"
    MySQL = MySQL & "                       dbo.TblAqar ON dbo.TblAqarDetai3.Aqarid = dbo.TblAqar.Aqarid LEFT OUTER JOIN"
    MySQL = MySQL & "                       dbo.TblAqarDetai2 ON dbo.TblAqar.Aqarid = dbo.TblAqarDetai2.Aqarid LEFT OUTER JOIN"
    MySQL = MySQL & "                       dbo.TblAqarDetai LEFT OUTER JOIN"
    MySQL = MySQL & "                       dbo.TblCustemers TblCustemers_1 ON dbo.TblAqarDetai.customerid = TblCustemers_1.CusID LEFT OUTER JOIN"
    MySQL = MySQL & "                       dbo.TblAkarUnit TblAkarUnit_1 ON dbo.TblAqarDetai.unittype = TblAkarUnit_1.id ON dbo.TblAqar.Aqarid = dbo.TblAqarDetai.Aqarid LEFT OUTER JOIN"
    MySQL = MySQL & "                       dbo.TblAkarUnit TblAkarUnit_2 ON dbo.TblAqar.UnitNo = TblAkarUnit_2.id LEFT OUTER JOIN"
    MySQL = MySQL & "                       dbo.TblCustemers TblCustemers_2 ON dbo.TblAqar.ownerid = TblCustemers_2.CusID LEFT OUTER JOIN"
    MySQL = MySQL & "                       dbo.tblSchemes ON dbo.TblAqar.schemeid = dbo.tblSchemes.id LEFT OUTER JOIN"
    MySQL = MySQL & "                       dbo.TblCountriesGovernments INNER JOIN"
    MySQL = MySQL & "                       dbo.TblCountriesGovernmentsCities ON dbo.TblCountriesGovernments.GovernmentID = dbo.TblCountriesGovernmentsCities.GovernmentID INNER JOIN"
    MySQL = MySQL & "                       dbo.TblCountriesData ON dbo.TblCountriesGovernments.CountryID = dbo.TblCountriesData.CountryID ON"
    MySQL = MySQL & "                       dbo.TblAqar.heyid = dbo.TblCountriesGovernmentsCities.CityID AND dbo.TblAqar.cityid = dbo.TblCountriesGovernments.GovernmentID AND"
    MySQL = MySQL & "                       dbo.TblAqar.CountryID = dbo.TblCountriesData.CountryID LEFT OUTER JOIN"
    MySQL = MySQL & "                       dbo.tblAkarType ON dbo.TblAqar.aqartypeid = dbo.tblAkarType.id"
    MySQL = MySQL & "  Where (dbo.TblAqar.Aqarid =" & val(Me.TxtAqarid.Text) & " )"

    If SystemOptions.UserInterface = ArabicInterface Then
        StrFileName = App.path & "\REPORTS\REPORTS NEW\" & "ReportIqar.rpt"
    Else
        StrFileName = App.path & "\REPORTS\REPORTS NEW\" & "ReportIqar.rpt"
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
        Msg = "?CE??I E?C?CE ?????"
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
    xReport.reporttitle = StrReportTitle
    xReport.EnableParameterPrompting = False
    xReport.ApplicationName = App.title
    xReport.ReportAuthor = App.title
    Set CViewer = New ClsReportViewer
    CViewer.FireReport xReport, WindowTarget, "", , , , StrFileName, , MySQL

    RsData.Close
    Set RsData = Nothing
    Screen.MousePointer = vbDefault
End Function



Private Sub BtnCancel_Click()
    Unload Me
End Sub
Private Sub btnDelete_Click()

    Dim MSGType As Integer
    Dim BlnRecExist As Boolean
    Dim StrMSG  As String
    Dim Msg As String
    Dim bo As Boolean
    
    On Error GoTo ErrTrap

        If TxtNoteSerial.Text <> "" Then
                 If SystemOptions.UserInterface = ArabicInterface Then
                         MsgBox "Ì—ÃÏ Õ–› «·ﬁÌœ «Ê·«"
                Else
                          MsgBox "Please Delete JE"
                End If
            CuurentLogdata "E"
        Exit Sub
        End If
        
    CheCotIqar bo
    If bo = True Then
        MsgBox "·«Ì„ﬂ‰ Õ–› Â–« «·⁄ﬁ«— ·«‰Â „— »ÿ »⁄ﬁœ"
        Exit Sub
    Else
        If DoPremis(Do_Delete, Me.Name, True) = False Then
            Exit Sub
        End If
        Dim StrSQL As String
        If TxtAqarid.Text <> "" Then
            MSGType = MsgBox("Â·  —€» ›Ì Õ–› Â–« «·”Ã·", vbQuestion + vbMsgBoxRtlReading + vbYesNo + vbMsgBoxRight, App.title)
            If MSGType = vbYes Then
                RsSavRec.Find "Aqarid=" & val(TxtAqarid.Text), , adSearchForward, 1
                RsSavRec.delete
                StrSQL = "Delete From TblAqarDetai Where Aqarid=" & val(Me.TxtAqarid.Text)
                Cn.Execute StrSQL, , adExecuteNoRecords
                StrSQL = "Delete From TblAqarDetai2 Where Aqarid=" & val(Me.TxtAqarid.Text)
                Cn.Execute StrSQL, , adExecuteNoRecords
                StrSQL = "Delete From TblAqarDetai3 Where Aqarid=" & val(Me.TxtAqarid.Text)
                Cn.Execute StrSQL, , adExecuteNoRecords
                StrSQL = "Delete From TblAqrOwin Where AqrID=" & val(Me.TxtAqarid.Text)
                Cn.Execute StrSQL, , adExecuteNoRecords
                If SystemOptions.OpenAccountAqar = True Then
                    If DcbAccount.BoundText <> "" Then
                   Cn.Execute " delete from Accounts where    Account_Code='" & DcbAccount.BoundText & "'"
                   End If
               End If
        
 
        
                MsgBox " „  ⁄„·Ì… «·Õ–› »‰Ã«Õ", vbInformation + vbMsgBoxRtlReading + vbOKOnly + vbMsgBoxRight, App.title
                CuurentLogdata ("D")
                FillGridWithData
                BtnNext_Click
            End If
        End If
    End If
    Exit Sub
ErrTrap:
    Select Case Err.Number
        Case -2147217873, -2147467259
            StrMSG = "⁄›Ê« ·« ÌÃÊ“ Õ–› «·”Ã· ·«— »«ÿÂ »»Ì«‰«  √Œ—Ì"
            RsSavRec.CancelUpdate
            MsgBox StrMSG, vbOKOnly + vbMsgBoxRight, App.title
            Cn.Errors.Clear
    End Select
End Sub
Private Sub BtnFirst_Click()

    On Error GoTo ErrTrap

    Dim Msg As String

    If Me.TxtModFlg.Text = "N" Then
        FindRec val(TxtAqarid.Text)
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
            Msg = "⁄›Ê« ·ﬁœ  „ Õ–› Â–« «·”Ã· «· «·Ï" & CHR(13)
            Msg = Msg & "„‰ ﬁ»· „” Œœ„ √Œ— ⁄·Ï «·‘»ﬂ… " & CHR(13)
            Msg = Msg & "”Ê› Ì „  ÕœÌÀ «·»Ì«‰« "
            MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading + vbOKOnly, App.title
            RsSavRec.Requery
            Resume BegnieWork
    End Select
End Sub
Private Sub BtnLast_Click()

    On Error GoTo ErrTrap

    Dim Msg As String

    If Me.TxtModFlg.Text = "N" Then
        FindRec val(TxtAqarid.Text)
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
    hideme
    Exit Sub
ErrTrap:
    Select Case Err.Number
        Case -2147217885
            Msg = "⁄›Ê« ·ﬁœ  „ Õ–› Â–« «·”Ã· «· «·Ï" & CHR(13)
            Msg = Msg & "„‰ ﬁ»· „” Œœ„ √Œ— ⁄·Ï «·‘»ﬂ… " & CHR(13)
            Msg = Msg & "”Ê› Ì „  ÕœÌÀ «·»Ì«‰« "
            MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading + vbOKOnly, App.title
            RsSavRec.Requery
            Resume BegnieWork
    End Select
End Sub

Private Sub btnModify_Click()

    On Error GoTo ErrTrap
    Dim Msg As String
    Dim bo As Boolean
    
        If TxtNoteSerial.Text <> "" Then
                 If SystemOptions.UserInterface = ArabicInterface Then
                         MsgBox "Ì—ÃÏ Õ–› «·ﬁÌœ «Ê·«"
                Else
                          MsgBox "Please Delete JE"
                End If
            CuurentLogdata "E"
        Exit Sub
        End If
        
    CheCotIqar bo
    bo = False


    If bo = True Then
        MsgBox "·«Ì„ﬂ‰  ⁄œÌ· Â–« «·⁄ﬁ«— ·«‰Â „— »ÿ »⁄ﬁœ"
        Exit Sub
    Else
        If DoPremis(Do_Edit, Me.Name, True) = False Then
            Exit Sub
        End If

        If TxtAqarid.Text <> "" Then
            TxtModFlg = "E"
            VSFlexGrid2.Rows = VSFlexGrid2.Rows + 1
            VSFlexGrid1.Rows = VSFlexGrid1.Rows + 1
            CuurentLogdata
            Frm2.Enabled = True
            Me.txtaqarname.SetFocus
        End If
    End If
    Exit Sub
ErrTrap:
    Select Case Err.Number
        Case -2147467259
            'Could not update; currently locked.
            Msg = "⁄›Ê«" & CHR(13)
            Msg = Msg & " ·«Ì„ﬂ‰  ⁄œÌ· Â–« «·”Ã· ›Ï «·Êﬁ  «·Õ«·Ï" & CHR(13)
            Msg = Msg & "ÕÌÀ «‰Â ﬁÌœ «· ⁄œÌ· „‰ ﬁ»· „” Œœ„ «Œ— ⁄·Ï «·‘»ﬂ…"
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
    Frm2.Enabled = True
    clear_all Me

    UnitsGrid.Clear flexClearScrollable, flexClearEverything
    UnitsGrid.Rows = 1
    UnitsGrid.Enabled = True

    VSFlexGrid1.Clear flexClearScrollable, flexClearEverything
    VSFlexGrid1.Rows = 2
    VSFlexGrid1.Enabled = True
    VSFlexGrid2.Clear flexClearScrollable, flexClearEverything
    VSFlexGrid2.Rows = 2
    VSFlexGrid2.Enabled = True
    Me.dcBranch.BoundText = Current_branch
    dpstatusdate.value = Date
    dpsuckdate.value = Date
    txtsuckdateH.value = ToHijriDate(Date)
    
    
'    TxtFromtxFromPlanneddate.value = Date
    'TxtFromtxFromPlanneddateh.value = ToHijriDate(TxtFromtxFromPlanneddate)
   txtToPlanneddate.value = Date
    txtToPlanneddateH.value = ToHijriDate(txtToPlanneddate)
        
    dpstatusdate.value = Date
    dcmaintenancetypeid.ListIndex = 0
    cbointerfaceid.ListIndex = 0
    DcboCountryID2.BoundText = 1
    cBORENTTYPE.ListIndex = 0
    TxtModFlg.Text = "N"
    XPTab301.CurrTab = 0
   RdRTypeDate(0).value = True
hideme
ErrTrap:
End Sub
Function hideme()
If RdRTypeDate(0).value = True Then ' ÂÃ—Ì
FromCotDateH.Visible = True
dpsuckdate.Visible = False
txtFromPlanneddate.Visible = False
txtToPlanneddate.Visible = False
DateCont.Visible = False
ToCotDate.Visible = False
FristPaymentDate.Visible = False
FirstInstallDateH.Visible = True
FromCotDate.Visible = False
txtsuckdateH.Visible = True
txtFromPlanneddateH.Visible = True
txtToPlanneddateH.Visible = True
DateHCont.Visible = True
ToCotDateH.Visible = True
 

With GridInstallments
       .ColHidden(.ColIndex("RecDate")) = True
       .ColHidden(.ColIndex("RecDateH")) = False
       
End With

Else
FromCotDateH.Visible = False
FromCotDate.Visible = True
dpsuckdate.Visible = True
txtFromPlanneddate.Visible = True
txtToPlanneddate.Visible = True
DateCont.Visible = True
ToCotDate.Visible = True
 FristPaymentDate.Visible = True
FirstInstallDateH.Visible = False

txtsuckdateH.Visible = False
txtFromPlanneddateH.Visible = False
txtToPlanneddateH.Visible = False
DateHCont.Visible = False
ToCotDateH.Visible = False
FirstInstallDateH.Visible = False

With GridInstallments
       .ColHidden(.ColIndex("RecDate")) = False
       .ColHidden(.ColIndex("RecDateH")) = True
       
End With
End If

End Function
Private Sub BtnNext_Click()

    On Error GoTo ErrTrap
    
    Dim Msg As String

    If Me.TxtModFlg.Text = "N" Then
        FindRec val(TxtAqarid.Text)
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
            Msg = "⁄›Ê« ·ﬁœ  „ Õ–› Â–« «·”Ã· «· «·Ï" & CHR(13)
            Msg = Msg & "„‰ ﬁ»· „” Œœ„ √Œ— ⁄·Ï «·‘»ﬂ… " & CHR(13)
            Msg = Msg & "”Ê› Ì „  ÕœÌÀ «·»Ì«‰« "
            MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading + vbOKOnly, App.title
            RsSavRec.Requery
            Resume BegnieWork
    End Select
End Sub
Private Sub BtnPrevious_Click()

    On Error GoTo ErrTrap
    
    Dim Msg As String

    If Me.TxtModFlg.Text = "N" Then
        FindRec val(TxtAqarid.Text)
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
            Msg = "⁄›Ê« ·ﬁœ  „ Õ–› Â–« «·”Ã· «· «·Ï" & CHR(13)
            Msg = Msg & "„‰ ﬁ»· „” Œœ„ √Œ— ⁄·Ï «·‘»ﬂ… " & CHR(13)
            Msg = Msg & "”Ê› Ì „  ÕœÌÀ «·»Ì«‰« "
            MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading + vbOKOnly, App.title
            RsSavRec.Requery
            Resume BegnieWork
    End Select
End Sub
Private Sub BtnPrint_Click()
    print_report
End Sub
Private Sub btnQuery_Click()
    Load FrmAqarSearch
    FrmAqarSearch.show
End Sub
Private Sub btnSave_Click()

       VBA.Calendar = vbCalGreg
   On Error GoTo ErrTrap
    btnSave.Enabled = False
    Dim Msg As String
    Dim StrVacCode As String
    Dim StrVacName As String
    Dim CtrlTxt As Control

    For Each CtrlTxt In Me.Controls
        If TypeOf CtrlTxt Is TextBox Or TypeOf CtrlTxt Is ComboBox Then
            If CtrlTxt.Text = "" And CtrlTxt.Tag <> "" And CtrlTxt.Enabled = True Then
                MsgBox CtrlTxt.Tag, vbOKOnly + vbMsgBoxRight, App.title
               CtrlTxt.SetFocus
               btnSave.Enabled = True
                Exit Sub
            End If
        End If
    Next

    StrVacName = IsRecExist("tblaqar", "aqarname", Trim(txtaqarname.Text), "aqarname", "Aqarid<>'" & Trim(TxtAqarid.Text) & "'")

    If StrVacName <> "" Then
        Msg = "·ﬁœ ”»ﬁ  ”ÃÌ· «”„ «·⁄ﬁ«— „‰ ﬁ»·"
        MsgBox Msg, vbOKOnly + vbMsgBoxRight, App.title
        txtaqarname.SetFocus
        btnSave.Enabled = True
        Exit Sub
    End If

    If txtaqarname.Text = "" Then
        MsgBox "Ì—ÃÏ ≈œŒ«· «”„ «·⁄ﬁ«—"
        btnSave.Enabled = True
        Exit Sub
    End If
    
    If val(dcsupplier.BoundText) = 0 Then
        MsgBox "Ì—ÃÏ ≈Œ Ì«— «·„«·ﬂ"
        btnSave.Enabled = True
        Exit Sub
    End If
    
    If val(dcBranch.BoundText) = 0 Then
        MsgBox "Ì—ÃÏ ≈Œ Ì«— «·›—⁄"
        
        dcBranch.SetFocus
        btnSave.Enabled = True
        Exit Sub
    End If
         If SystemOptions.OpenAccountAqar = True Then
          Account_Code_dynamic = GetMyAccountCode("TblCustemers", "CusID", val(Me.dcsupplier.BoundText), "AccountAccountAqar")
          If Account_Code_dynamic = "" Then
          MsgBox "Â–« «·„«·ﬂ ·Ì” ·Â Õ”«» Œ«’ »«·⁄ﬁ«—« "
          Exit Sub
          End If
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
    MsgBox "Â‰«ﬂ Œÿ√ „« ›Ì ≈œŒ«· «·»Ì«‰« ", vbOKOnly + vbMsgBoxRight, App.title
    btnSave.Enabled = True
    
End Sub
Private Sub BtnUndo_Click()
    FindRec val(TxtAqarid.Text)
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
    MsgBox Msg, vbOKOnly + vbMsgBoxRight + vbInformation, App.title
ErrTrap:
End Sub
Private Sub LoadDataCombos(Optional BolExceptCountries As Boolean = False, Optional BolExceptGovern As Boolean = False, Optional BolExceptCities As Boolean = False)

    Dim Dcombo As New ClsDataCombos
    Dcombo.GetCountriesNames Me.DcboCountryID2
    If BolExceptGovern = False Then
        Dcombo.getCountriesGovernments Me.DcboGovernmentID, val(Me.DcboCountryID2.BoundText)
    End If
    If BolExceptCities = False Then
        Dcombo.GetCountriesGovernCities Me.DcboCityID, val(Me.DcboCountryID2.BoundText), val(Me.DcboGovernmentID.BoundText)

    End If
End Sub
Private Sub ChkOrder_Click()
    FillGridWithData
End Sub
Private Sub Cmd_Click(Index As Integer)
    Select Case Index
    Case 11
      If DoPremis(Do_Attach, Me.Name, True) = False Then
                Exit Sub
            End If
            ShowAttachments TxtAqarid.Text, "23022021"
            
        Case 20
            fillgride
        Case 21
            If CheckUnitContract(val(UnitsGrid.TextMatrix(UnitsGrid.Row, UnitsGrid.ColIndex("id")))) = False And CheckUnitContractMerg(val(UnitsGrid.TextMatrix(UnitsGrid.Row, UnitsGrid.ColIndex("id")))) = False Then
                RemoveGridRow
            Else
                MsgBox " «·ÊÕœ… —ﬁ„ " & UnitsGrid.TextMatrix(UnitsGrid.Row, UnitsGrid.ColIndex("unitno")) & "   ·Â« ⁄ﬁœ Ê·« Ì„ﬂ‰ Õ–›Â«"
                Exit Sub
            End If

        Case 0
            Calculations
        Case 10
            Dim X As Integer
            Dim i As Integer
            With UnitsGrid
                For i = 1 To .Rows - 1
                    If CheckUnitContract(val(UnitsGrid.TextMatrix(i, UnitsGrid.ColIndex("id")))) = True Or CheckUnitContractMerg(val(UnitsGrid.TextMatrix(i, UnitsGrid.ColIndex("id")))) = True Then
                        MsgBox " «·ÊÕœ… —ﬁ„ " & UnitsGrid.TextMatrix(i, UnitsGrid.ColIndex("unitno")) & "   ·Â« ⁄ﬁœ Ê·« Ì„ﬂ‰ Õ–›Â«"
                        Exit Sub
                    End If
                Next i
            End With
            X = MsgBox("Õ–› ﬂ· «·”ÿÊ—", vbCritical + vbYesNo)
            If X = vbYes Then
                UnitsGrid.Clear flexClearScrollable, flexClearEverything
                UnitsGrid.Rows = 1
            End If
    End Select
End Sub
Private Sub RemoveGridRow()
    With Me.UnitsGrid
        If .Row <= 0 Then Exit Sub
        .RemoveItem .Row
    End With
    ReLineGrid
End Sub
Sub fillgride()
    Dim StrSQL As String
    Dim Msg As String
    Dim i As Integer
    Dim j As Integer
    Dim RsDetails As ADODB.Recordset
    Set RsDetails = New ADODB.Recordset
    If DCAkarUnit.BoundText = "" Then
            If SystemOptions.UserInterface = ArabicInterface Then
                Msg = "ÌÃ»       «Œ Ì«— ‰Ê⁄ «·ÊÕœÂ  ...!!!"
            Else
                Msg = "must Specify Type Unit ...!!!"
            End If

            MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
            Exit Sub
        End If
     If UnitID.Text = "" And val(txtFrom.Text) = 0 Then
     If SystemOptions.UserInterface = ArabicInterface Then
     MsgBox "Ì—ÃÏ «œŒ«·  ”·”· «·ÊÕœ…"
     Else
     MsgBox "Please Eneter Unit No"
     End If
     txtFrom.SetFocus
     Exit Sub
     End If
    j = val(Me.txtFrom.Text)

    StrSQL = "SELECT     *  from dbo.TblAqarDetai Where (1 = -1)"


    
    RsDetails.Open StrSQL, Cn, adOpenKeyset, adLockOptimistic, adCmdText
    If val(TxtCount.Text) = 0 Then
    TxtCount.Text = 1
    End If
    For i = 1 To val(TxtCount.Text)
    RsDetails.AddNew
    RsDetails("id").value = new_id("TblAqarDetai", "id", "")
    If j <> 0 Then
     RsDetails("unitno").value = j
    j = j + 1
    Else
    RsDetails("unitno").value = UnitID.Text
    End If
    
    RsDetails("Aqarid").value = val(TxtAqarid.Text)
    RsDetails("MiniRentValue").value = val(TxtMiniRentValue.Text)
    RsDetails("UnitElectric").value = (UnitElc.Text)
    RsDetails("namerentType").value = cBORENTTYPE.Text
    RsDetails("unittype").value = val(DCAkarUnit.BoundText)
    RsDetails("rentType").value = val(cBORENTTYPE.ListIndex) + 1
    RsDetails("length").value = TxtLenght.Text
    RsDetails("meterPrice").value = val(txtMeterPrice.Text)
    RsDetails("roomscount").value = val(TxtRooms.Text)
    RsDetails("WCcount").value = val(BathNo.Text)
    RsDetails("kithchencount").value = val(TxtKitchn.Text)
    If FerntChk = xtpChecked Then
        RsDetails("haveFurniture").value = -1
    Else
        RsDetails("haveFurniture").value = 0
    End If
    
    
    If chkIsTax = xtpChecked Then
        RsDetails("IsTax").value = -1
    Else
        RsDetails("IsTax").value = 0
    End If
    
    
    
    RsDetails("Status") = val(UnitStatus.BoundText)
    RsDetails("unitdesc").value = Disc.Text
    RsDetails("RentValue").value = val(RentValue.Text)
    RsDetails("customerid").value = val(RenterDC.BoundText)
    RsDetails("Floor").value = TxtFloors.Text
    RsDetails("LoungeCount").value = val(TxtLoung.Text)
    RsDetails("ACCount").value = val(TxtAccount.Text)
    RsDetails("ACCountspleat").value = val(TxtACCountÚSpleat.Text)
    RsDetails("Typed").value = val(DcbTyped.ListIndex) + 1
    RsDetails.update
   Next i
   MsgBox " „ «·Õ›Ÿ"
    FillGridWithData
    ReLineGrid
End Sub
Private Sub Command1_Click(Index As Integer)

    On Error Resume Next
    
    If DoPremis(Do_Attach, Me.Name, True) = False Then
        Exit Sub
    End If
    
    If TxtAqarid.Text = "" Then
        If SystemOptions.UserInterface = ArabicInterface Then
            MsgBox "·«»œ „‰ «Õ Ì«— «·⁄ﬁ«— «Ê·«": Exit Sub
        Else
            MsgBox "Select Voucher Firstly": Exit Sub
        End If
    End If
    Unload imaged
    imaged.show
    If SystemOptions.UserInterface = EnglishInterface Then
        imaged.Label9.Caption = "Aqar #"
        imaged.Caption = "Aqar Attachment"
        imaged.txtopeation_type = "„—›ﬁ«  «·⁄ﬁ«—« "
        imaged.SUBJECT_NO = TxtAqarid.Text
        imaged.Label6.Caption = "Voucher #"
    Else
        imaged.Label9.Caption = "„—›ﬁ«    ⁄ﬁ«—  —ﬁ„"
        imaged.Caption = "„—›ﬁ«  «·⁄ﬁ«—  "
        imaged.txtopeation_type = "„—›ﬁ«  «·⁄ﬁ«—"
        imaged.SUBJECT_NO = TxtAqarid.Text
        imaged.Label6.Caption = "—ﬁ„  «·⁄ﬁ«—"
    End If
    imaged.Adodc1.CommandType = adCmdText
    imaged.Adodc1.RecordSource = "SELECT * FROM subjects_images WHERE operation_type = '„—›ﬁ«  «·⁄ﬁ«—' and subject_no='" & TxtAqarid.Text & "'"
    imaged.Adodc1.Refresh
    If imaged.Adodc1.Recordset.RecordCount > 0 Then
        imaged.DBPix201.Visible = True
    Else
        imaged.DBPix201.Visible = False
    End If
End Sub

Private Sub Command9_Click()
     '  ShowGL_cc Me.TxtNoteSerial.Text, , 200
       
       ShowGL_cc TxtNoteSerial.Text, , 200, val(Me.TXTNoteID.Text)
       
End Sub

Private Sub DateCont_Change()
    If Me.TxtModFlg.Text <> "R" Then
        DateHCont.value = ToHijriDate(DateCont.value)
        
 
   
    End If
End Sub
Private Sub DateHCont_LostFocus()
    If Me.TxtModFlg.Text <> "R" Then
        VBA.Calendar = vbCalGreg
        DateCont.value = ToGregorianDate(DateHCont.value)
    End If
End Sub
Private Sub DCAkarUnit_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF5 Then
        loadcombo
    End If
End Sub
Private Sub dcaqartypeid_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF5 Then
        loadcombo
    End If
End Sub
Private Sub DcboCityID_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF5 Then
        loadcombo
    End If
End Sub
Private Sub DcboCountryID2_Change()
    Dim Dcombos As ClsDataCombos
    
    Set Dcombos = New ClsDataCombos
    Dcombos.getCountriesGovernments Me.DcboGovernmentID, val(Me.DcboCountryID2.BoundText)
End Sub
Private Sub DcboCountryID2_Click(Area As Integer)
    DcboCountryID2_Change
End Sub
Private Sub DcboCountryID2_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF5 Then
        loadcombo
    End If
End Sub
Private Sub DcboGovernmentID_Change()
    LoadDataCombos False, True, False
End Sub
Private Sub DcboGovernmentID_Click(Area As Integer)
    DcboGovernmentID_Change
End Sub
Private Sub DcboGovernmentID_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF5 Then
        loadcombo
    End If
End Sub
Private Sub dcBranch_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF5 Then
        loadcombo
    End If
End Sub
Private Sub DcbSales_Change()
    If val(DcbSales.BoundText) = 0 Then Me.TxtCodeSales.Text = "": Exit Sub
    Me.TxtCodeSales.Text = get_EMPLOYEE_Data(val(Me.DcbSales.BoundText), "Fullcode")
End Sub
Private Sub DcbSales_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF5 Then
        loadcombo
    End If
End Sub
Private Sub DcFixedAssets_Change()
    DcFixedAssets_Click (0)
End Sub
Function GetValueAssest() As Double
    Dim sql As String
    Dim Rs3 As ADODB.Recordset

    sql = " SELECT     AccDepreciation, id"
    sql = sql & " From dbo.FixedAssets"
    sql = sql & " Where (ID = " & val(DcFixedAssets.BoundText) & ")"
    Set Rs3 = New ADODB.Recordset
    Rs3.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
    If Rs3.RecordCount > 0 Then
        GetValueAssest = IIf(IsNull(Rs3("AccDepreciation").value), 0, Rs3("AccDepreciation").value)
    Else
        GetValueAssest = 0
    End If
End Function
Private Sub DcFixedAssets_Click(Area As Integer)
    Dim AsseCode1 As String

    GetAsseteCode_ID val(DcFixedAssets.BoundText), AsseCode1, 0
    TxAssetscode.Text = AsseCode1
End Sub
Private Sub dcmaintenancetypeid_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF5 Then
        loadcombo
    End If
End Sub
Private Sub dcschemeid_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF5 Then
        loadcombo
    End If
End Sub
Private Sub dcsupplier_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF5 Then
        loadcombo
    End If
End Sub
Private Sub dpsuckdate_Change()
    If Me.TxtModFlg.Text <> "R" Then
        txtsuckdateH.value = ToHijriDate(dpsuckdate.value)
    End If
End Sub

Private Sub dcsupplier_Click(Area As Integer)
    If val(dcsupplier.BoundText) = 0 Then Exit Sub
    Dim EmpCode  As String
    GetTblCustemersCode , , dcsupplier.BoundText, EmpCode
    Me.Text1(10).Text = EmpCode
End Sub

Private Sub Editbtn_Click()
    Dim haveFurniture As Integer
    Dim isTax As Integer
    Dim sql As String
    If SystemOptions.AllowChangeUnitIqar = False Then
    If (CheckUnitContract(val(RecGID.Text)) = True Or CheckUnitContractMerg(val(RecGID.Text)) = True) Then
            If SystemOptions.UserInterface = ArabicInterface Then
                MsgBox "·«Ì„ﬂ‰ «· ⁄œÌ· ⁄·Ï »Ì«‰«  Â–Â «·ÊÕœ… ·«— »«ÿÂ« »⁄ﬁœ "
            Else
                MsgBox "This Unit data cannot be edited for being integrated with a contract record"
            End If
         Exit Sub
     End If
     End If
            If FerntChk = xtpChecked Then
            haveFurniture = -1
            Else
            haveFurniture = 0
            End If
            
            If chkIsTax = xtpChecked Then
                isTax = 1
            Else
                isTax = 0
            End If
                        
            
            
            sql = " insert into  UpdatesoldunitNo (oldunitNo,NewnitNo,TransDate,UserId,Aqarid )"
            sql = sql & " values('" & DataBaseUnitNio & "','" & UnitID.Text & "'," & SQLDate(Now, True) & "," & user_id & "," & val(TxtAqarid.Text) & ")"

 Cn.Execute sql

            sql = "Update TblAqarDetai set MiniRentValue=" & val(TxtMiniRentValue.Text) & " ,UnitElectric='" & UnitElc.Text & "'"
            sql = sql & " , namerentType='" & cBORENTTYPE.Text & "' "
           sql = sql & " ,  unitno='" & UnitID.Text & "'"
            sql = sql & " , unittype=" & val(DCAkarUnit.BoundText) & ",rentType=" & (val(cBORENTTYPE.ListIndex) + 1) & ""
            sql = sql & " , length='" & val(TxtLenght.Text) & "' ,meterPrice=" & val(txtMeterPrice.Text) & " "
            sql = sql & " , roomscount=" & val(TxtRooms.Text) & " ,WCcount=" & val(BathNo.Text) & "       "
            sql = sql & " ,kithchencount=" & val(TxtKitchn.Text) & ",haveFurniture=" & haveFurniture & ",isTax = " & isTax
            sql = sql & " ,Status=" & val(UnitStatus.BoundText) & ",RentValue=" & val(RentValue.Text) & " "
            sql = sql & " , unitdesc='" & Disc.Text & "' ,Floor='" & TxtFloors.Text & "' "
            sql = sql & " , customerid=" & val(RenterDC.BoundText) & " ,LoungeCount=" & val(TxtLoung.Text) & " "
            sql = sql & " ,ACCount=" & val(TxtAccount.Text) & " ,Typed=" & val(DcbTyped.ListIndex) + 1 & "  ,ACCountspleat=" & val(TxtACCountÚSpleat.Text) & " "
            sql = sql & " where Aqarid =" & val(TxtAqarid.Text) & " and ID=" & val(RecGID) & " "
            Cn.Execute sql
            
            If SystemOptions.UserInterface = ArabicInterface Then
                MsgBox " „ «· ⁄œÌ· »‰Ã«Õ"
            Else
                MsgBox "Recored edited successfully"
            End If
                FillGridWithData
                ReLineGrid
        
  
End Sub

Private Sub FirstInstallDateH_GotFocus()
    hijriorJerojian = 0
End Sub
Private Sub FirstInstallDateH_LostFocus()
    If Me.TxtModFlg.Text <> "R" Then
        VBA.Calendar = vbCalGreg
        FristPaymentDate.value = ToGregorianDate(FirstInstallDateH.value)
    End If
End Sub
Private Sub Form_Load()

    On Error GoTo ErrTrap
 
    Dim i As Integer
    Dim My_SQL As String
    Dim ScreenNameArabic As String
    Dim ScreenNameEnglish As String
    ScreenNameArabic = " »Ì«‰«  «·⁄ﬁ«— "
    ScreenNameEnglish = " Real Estate"
    ReloadCombos
    

      If SystemOptions.AllowEditVaTManulay = True Then
txtManulaVat.Enabled = True
txtManulaVat.Visible = True
Else
txtManulaVat.Enabled = False
txtManulaVat.Text = 0
txtManulaVat.Visible = False
End If
XPTab301.TabVisible(4) = False
If mIndex = 1 Then
    RSAkar.Caption = "«· Œ›Ì÷"
    XPTab301.TabVisible(0) = False
    XPTab301.TabVisible(1) = False
    XPTab301.TabVisible(2) = False
    XPTab301.TabVisible(3) = False
    XPTab301.TabVisible(4) = True
    EltCont.Visible = False
    btnLast.Visible = False
    btnNext.Visible = False
    btnPrevious.Visible = False
    btnFirst.Visible = False
            With XPCboDiscountType
            .Clear
            .AddItem "·«ÌÊÃœ  Œ›Ì÷"
            .AddItem " Œ›Ì÷ »ﬁÌ„…"
            .AddItem " Œ›Ì÷ »‰”»…"
        End With

End If

    If SystemOptions.UserInterface = ArabicInterface Then
        GridInstallments.ColComboList(GridInstallments.ColIndex("DMY")) = "#1; ÌÊ„|#2; ‘Â—|#3; ”‰Â"
        UnitsGrid.ColComboList(UnitsGrid.ColIndex("Typed")) = "#1; ”ﬂ‰Ì|#2;  Ã«—Ì"
    ElseIf SystemOptions.UserInterface = EnglishInterface Then
        GridInstallments.ColComboList(GridInstallments.ColIndex("DMY")) = "#1;Day |#2;Month |#3; Year"
        UnitsGrid.ColComboList(UnitsGrid.ColIndex("Typed")) = "#1;Residential |#2;Commercial "
    End If
FrmDate.value = Date
todate.value = Date
    Dim cOptions As ClsCompanyInfo
    Set cOptions = New ClsCompanyInfo
    If SystemOptions.UserInterface = ArabicInterface Then
        lblCompanyname.Caption = cOptions.ArabCompanyName & CHR(13) & CurrentBranchName
    Else
        lblCompanyname.Caption = cOptions.EngCompanyName & CHR(13) & CurrentBranchNameE
    End If

    Set BKGrndPic = New ClsBackGroundPic
    Set RsSavRec = New ADODB.Recordset
    RsSavRec.CursorLocation = adUseClient
    My_SQL = " select * from tblaqar "
        If SystemOptions.usertype <> UserAdminAll Then
        My_SQL = My_SQL & " where   BranchId=" & Current_branch
    End If
       
    RsSavRec.Open My_SQL, Cn, adOpenStatic, adLockOptimistic, adCmdText
    Me.TxtModFlg.Text = "R"
    Resize_Form Me

    BtnFirst_Click
    ShowTip
    LoadDataCombos
    With UnitsGrid
        If SystemOptions.UserInterface = ArabicInterface Then
            .Cell(flexcpPictureAlignment, 0, 0, .Rows - 1, .Cols - 1) = flexPicAlignRightCenter
            .ColComboList(.ColIndex("rentType")) = "#1;«Ã„«·Ì «·ÊÕœ…|#2;»«·„ —"
        Else
            .ColComboList(.ColIndex("rentType")) = "#1;Totals|#2;By Meter"
        End If
    End With
    loadcombo

    If SystemOptions.UserInterface = EnglishInterface Then
        SetInterface Me
        ChangeLang
    End If

    If OPEN_NEW_SCREEN = True Then
        btnNew_Click
    End If

ErrTrap:
End Sub
Public Sub loadcombo()
    Dim Dcombos As ClsDataCombos
    Dim My_SQL As String
    Set Dcombos = New ClsDataCombos

    Dcombos.GetAccountingCodes Me.DcbAccount, True, False
    Dcombos.GetBranches dcBranch
    My_SQL = "select UserID,UserName From tblUsers "
    fill_combo DCUser, My_SQL
    
    
    Dcombos.get«hay Me.DcboCityID
    Dcombos.GetCustomersSuppliers 57, Me.dcsupplier
    Dcombos.GetSalesRepData Me.DcbSales
    Dcombos.getSchemes Me.dcschemeid
    Dcombos.getAkarType Me.dcaqartypeid
    Dcombos.getAkarUnit Me.DCAkarUnit
    Dcombos.GetFixedAssets Me.DcFixedAssets
    
    My_SQL = "Select * from TblCustemers where type = 56"
    fill_combo RenterDC, My_SQL
    
    My_SQL = "select * from TblRentStatus"
    fill_combo UnitStatus, My_SQL
End Sub
Private Sub ChangeLang()

    Dim XPic As IPictureDisp
    Set XPic = Me.btnFirst.ButtonImage
    Set Me.btnFirst.ButtonImage = Me.btnLast.ButtonImage
    Set Me.btnLast.ButtonImage = XPic
    Set XPic = Me.btnPrevious.ButtonImage
    Set Me.btnPrevious.ButtonImage = Me.btnNext.ButtonImage
    Set Me.btnNext.ButtonImage = XPic

    Me.Caption = "tblaqar Data"
    Me.Label1(2).Caption = Me.Caption
    Label1(3).Caption = "Code"
    Label1(0).Caption = "Name"
    Label1(1).Caption = "Neighborhood"
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
        .TextMatrix(0, .ColIndex("Aqarid")) = "Id"
        .TextMatrix(0, .ColIndex("CityName")) = "Name"
        .TextMatrix(0, .ColIndex("GovernmentID")) = "Neighborhood"
    End With
End Sub
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

Private Sub Form_Activate()
    Me.ZOrder 0
End Sub
Private Sub ReLineGrid()
    Dim i As Integer
    Dim IntCounter As Integer
    
    With VSFlexGrid1
        For i = .FixedRows To .Rows - 1
            If .TextMatrix(i, .ColIndex("Waset")) <> "" Then
                IntCounter = IntCounter + 1
                .TextMatrix(i, .ColIndex("Ser")) = IntCounter
            End If
        Next i
    End With
    With UnitsGrid
        For i = .FixedRows To .Rows - 1
            If .TextMatrix(i, .ColIndex("unittype")) <> "" Then
                IntCounter = IntCounter + 1
                .TextMatrix(i, .ColIndex("Ser")) = IntCounter
                If val(.TextMatrix(i, .ColIndex("rentType"))) = 1 Then
                    .TextMatrix(i, .ColIndex("RentValue")) = val(.TextMatrix(i, .ColIndex("length"))) * val(.TextMatrix(i, .ColIndex("meterPrice")))
                End If
            End If
        Next i
    End With
    With VSFlexGrid2
        For i = .FixedRows To .Rows - 1
            If .TextMatrix(i, .ColIndex("ElevatorNo")) <> "" Then
                IntCounter = IntCounter + 1
                .TextMatrix(i, .ColIndex("Ser")) = IntCounter
            End If
        Next i
    End With
End Sub
Public Sub AddNewRec()

    On Error GoTo ErrTrap
    
    Dim StrRecID As String
    
    StrRecID = new_id("tblaqar", "Aqarid", "")
    RsSavRec.AddNew
    RsSavRec.Fields("Aqarid").value = IIf(StrRecID <> "", StrRecID, Null)
    TxtAqarid.Text = StrRecID
    FiLLRec
ErrTrap:
End Sub
Public Sub FiLLRec()
   ' On Error GoTo ErrTrap
    
    Dim RsDetails3 As ADODB.Recordset
    Dim RsDetails1 As ADODB.Recordset
    Dim RsDetails2 As ADODB.Recordset
    Dim RsDetails As ADODB.Recordset
    Dim i As Integer
    Dim Msg As String
    Dim StrSQL As String
    
    If Me.TxtModFlg.Text = "E" Then
     '  StrSQL = "Delete From TblAqarDetai Where Aqarid=" & val(Me.TxtAqarid.Text)
     '   Cn.Execute StrSQL, , adExecuteNoRecords
        StrSQL = "Delete From TblAqarDetai2 Where Aqarid=" & val(Me.TxtAqarid.Text)
        Cn.Execute StrSQL, , adExecuteNoRecords
        StrSQL = "Delete From TblAqarDetai3 Where Aqarid=" & val(Me.TxtAqarid.Text)
        Cn.Execute StrSQL, , adExecuteNoRecords
        StrSQL = "Delete From TblAqrOwin Where AqrID=" & val(Me.TxtAqarid.Text)
        Cn.Execute StrSQL, , adExecuteNoRecords
    End If


             If ComResid(1).value = True Then
        RsSavRec.Fields("ComResid").value = 1
        Else
        RsSavRec.Fields("ComResid").value = 0
        End If
        
        
   
   RsSavRec("TxtContValueWithout").value = IIf(val(Me.TxtContValueWithout.Text) = 0, 0, val(TxtContValueWithout.Text))
   RsSavRec("TxtFATYou").value = IIf(val(Me.TxtFATYou.Text) = 0, 0, val(TxtFATYou.Text))
   RsSavRec("TxtFATValue").value = IIf(val(Me.TxtFATValue.Text) = 0, 0, val(TxtFATValue.Text))
   

    RsSavRec("SalesEmp").value = IIf(Me.DcbSales.BoundText = "", 0, val(DcbSales.BoundText))
    RsSavRec("BranchId").value = IIf(Me.dcBranch.BoundText = "", 0, val(dcBranch.BoundText))
    RsSavRec.Fields("Aqarid").value = IIf(Me.TxtAqarid.Text <> "", val(TxtAqarid.Text), Null)
    RsSavRec.Fields("AqarName").value = IIf(Me.txtaqarname.Text <> "", Trim(txtaqarname.Text), Null)
    RsSavRec.Fields("eastWriiten").value = IIf(Me.TxteastWriiten.Text <> "", Trim(TxteastWriiten.Text), Null)
    RsSavRec.Fields("westWriiten").value = IIf(Me.TxtwestWriiten.Text <> "", Trim(TxtwestWriiten.Text), Null)
    RsSavRec.Fields("Rate").value = IIf(Me.TxtRat.Text <> "", Trim(TxtRat.Text), Null)
    RsSavRec.Fields("UnitNo").value = IIf(Me.TxtUnit.Text <> "", Trim(TxtUnit.Text), Null)
    RsSavRec.Fields("Block").value = IIf(Me.TxtBlock.Text <> "", Trim(TxtBlock.Text), Null)
    RsSavRec.Fields("Part").value = IIf(Me.TxtPart.Text <> "", Trim(TxtPart.Text), Null)
    RsSavRec.Fields("StreetNo").value = IIf(Me.TxtStreet.Text <> "", Trim(TxtStreet.Text), Null)
    RsSavRec.Fields("PriceSomW").value = IIf(Me.TxtPriceSomW.Text <> "", Trim(TxtPriceSomW.Text), Null)
    RsSavRec.Fields("PriceHadW").value = IIf(Me.TxtPriceHadW.Text <> "", Trim(TxtPriceHadW.Text), Null)
    RsSavRec.Fields("Price").value = IIf(Me.TxtPrice.Text <> "", Trim(TxtPrice.Text), Null)
    RsSavRec.Fields("PriceSom").value = IIf(Me.TxtPriceSom.Text <> "", Trim(TxtPriceSom.Text), Null)
    RsSavRec.Fields("PriceHad").value = IIf(Me.TxtPriceHad.Text <> "", Trim(TxtPriceHad.Text), Null)
    RsSavRec.Fields("FixedID").value = IIf(val(DcFixedAssets.BoundText) <> 0, val(DcFixedAssets.BoundText), Null)
    
    
         If RdRTypeDate(1).value = True Then
             RsSavRec("TypeDate").value = 1
        Else
             RsSavRec("TypeDate").value = 0
        End If
    
    If RSOutSupplier.value = vbChecked Then
        Rd.value = True
    Else
        Rd.value = False
    End If
    
    If Rd.value = True Then
        RsSavRec.Fields("TypAmola").value = 1
        RsSavRec.Fields("AmolaValus").value = IIf(TxtKickbacks.Text <> "", val((TxtKickbacks.Text)), Null)
    Else
        TxtKickbacks.Text = 0
        RsSavRec.Fields("TypAmola").value = 0
        RsSavRec.Fields("AmolaValus").value = 0
    End If
    RsSavRec.Fields("ValYearIncrease").value = IIf(Me.TxtValYearIncrease.Text <> "", Trim(TxtValYearIncrease.Text), Null)
    RsSavRec.Fields("NOOFYears").value = val(NOOFYears.Text)
    
    RsSavRec.Fields("Provide").value = IIf(TxtProvide.Text <> "", Trim(TxtProvide.Text), Null)
    RsSavRec.Fields("Remarks").value = IIf(TxtRemarks.Text <> "", Trim(TxtRemarks.Text), Null)
    RsSavRec.Fields("BanckName").value = IIf(TxtBanckName.Text <> "", Trim(TxtBanckName.Text), Null)
    RsSavRec.Fields("AgemcyNo").value = IIf(TxtagencyNo.Text <> "", Trim(TxtagencyNo.Text), Null)
    RsSavRec.Fields("Telephone").value = IIf(txtTel.Text <> "", Trim(txtTel.Text), Null)
    RsSavRec.Fields("Mobile").value = IIf(TxtMobile.Text <> "", Trim(TxtMobile.Text), Null)
    RsSavRec.Fields("Email").value = IIf(TxtEmail.Text <> "", Trim(TxtEmail.Text), Null)
    RsSavRec.Fields("Fax").value = IIf(TxtFaxAg.Text <> "", Trim(TxtFaxAg.Text), Null)
    RsSavRec.Fields("AccountBank").value = IIf(TxtAcountBank.Text <> "", Trim(TxtAcountBank.Text), Null)
    RsSavRec.Fields("ContValue").value = val(IIf(TxtContValue.Text <> "", Trim(TxtContValue.Text), 0))
    RsSavRec.Fields("PaymentNo").value = val(IIf(TxtPaymentCount.Text <> "", Trim(TxtPaymentCount.Text), 0))
    RsSavRec.Fields("Priod").value = val(IIf(TxtPeriods.Text <> "", Trim(TxtPeriods.Text), 0))
    RsSavRec.Fields("PriodAlow").value = val(IIf(TxtPriodAlow.Text <> "", Trim(TxtPriodAlow.Text), 0))
    RsSavRec.Fields("PriodDMY").value = IIf(DcbPeriodsID.ListIndex <> -1, val(DcbPeriodsID.ListIndex), Null)
    RsSavRec.Fields("PriodAlowDMY").value = IIf(DcbPeriodsAlowID.ListIndex <> -1, val(DcbPeriodsAlowID.ListIndex), Null)
    RsSavRec.Fields("DateHCont").value = DateHCont.value
    RsSavRec.Fields("DateCont").value = DateCont.value
    RsSavRec.Fields("FromCotDateH").value = FromCotDateH.value
    RsSavRec.Fields("FromCotDate").value = FromCotDate.value
    RsSavRec.Fields("ToCotDateH").value = ToCotDateH.value
    RsSavRec.Fields("ToCotDate").value = ToCotDate.value
    RsSavRec.Fields("FristPaymentDate").value = FristPaymentDate.value
    RsSavRec.Fields("FirstInstallDateH").value = FirstInstallDateH.value
    RsSavRec.Fields("countryid").value = IIf(DcboCountryID2.BoundText <> "", val(DcboCountryID2.BoundText), Null)
    RsSavRec.Fields("cityid").value = IIf(DcboGovernmentID.BoundText <> "", val(DcboGovernmentID.BoundText), Null)
    RsSavRec.Fields("heyid").value = IIf(DcboCityID.BoundText <> "", val(DcboCityID.BoundText), Null)
    RsSavRec.Fields("streetname").value = IIf(txtstreetname.Text <> "", Trim(txtstreetname.Text), Null)
    RsSavRec.Fields("schemeid").value = IIf(dcschemeid.BoundText <> "", val(dcschemeid.BoundText), Null)
    RsSavRec.Fields("aqartypeid").value = IIf(dcaqartypeid.BoundText <> "", val(dcaqartypeid.BoundText), Null)
    RsSavRec.Fields("aqarNo").value = IIf(txtaqarNo.Text <> "", Trim(txtaqarNo.Text), Null)
    RsSavRec.Fields("location").value = IIf(txtlocation.Text <> "", Trim(txtlocation.Text), Null)
    RsSavRec.Fields("aqarage").value = IIf(txtaqarage.Text <> "", val(txtaqarage.Text), Null)
    RsSavRec.Fields("floorcount").value = IIf(txtfloorcount.Text <> "", val(txtfloorcount.Text), Null)
    RsSavRec.Fields("currentPrice").value = IIf(txtcurrentPrice.Text <> "", val(txtcurrentPrice.Text), Null)
    RsSavRec.Fields("lastrentvalue").value = IIf(txtlastrentvalue.Text <> "", val(txtlastrentvalue.Text), Null)
    RsSavRec.Fields("interfaceid").value = IIf(cbointerfaceid.ListIndex <> -1, val(cbointerfaceid.ListIndex), Null)
    RsSavRec.Fields("statusid").value = IIf(cbostatusid.ListIndex <> -1, val(cbostatusid.ListIndex), Null)
    RsSavRec.Fields("maintenancetypeid").value = IIf(dcmaintenancetypeid.ListIndex <> -1, val(dcmaintenancetypeid.ListIndex), Null)
    RsSavRec.Fields("statusdate").value = dpstatusdate.value
    RsSavRec.Fields("EntryCount").value = IIf(txtEntryCount.Text <> "", val(txtEntryCount.Text), Null)
    RsSavRec.Fields("noofapartement").value = IIf(txtnoofapartement.Text <> "", val(txtnoofapartement.Text), Null)
    RsSavRec.Fields("noofoffices").value = IIf(txtnoofoffices.Text <> "", val(txtnoofoffices.Text), Null)
    RsSavRec.Fields("noofparking").value = IIf(txtnoofparking.Text <> "", val(txtnoofparking.Text), Null)
    RsSavRec.Fields("northlength").value = IIf(txtnorthlength.Text <> "", val(txtnorthlength.Text), Null)
    RsSavRec.Fields("Southlength").value = IIf(txtSouthlength.Text <> "", val(txtSouthlength.Text), Null)
    RsSavRec.Fields("eastlength").value = IIf(txteastlength.Text <> "", val(txteastlength.Text), Null)
    RsSavRec.Fields("Westlength").value = IIf(txtWestlength.Text <> "", val(txtWestlength.Text), Null)
    RsSavRec.Fields("totallength").value = IIf(txttotallength.Text <> "", val(txttotallength.Text), Null)
    RsSavRec.Fields("meterRentvalue").value = IIf(txtmeterRentvalue.Text <> "", val(txtmeterRentvalue.Text), Null)
    RsSavRec.Fields("metersalevalue").value = IIf(txtmetersalevalue.Text <> "", val(txtmetersalevalue.Text), Null)
    RsSavRec.Fields("googlemap").value = IIf(txtgooglemap.Text <> "", Trim(txtgooglemap.Text), Null)
    RsSavRec.Fields("ownerid").value = IIf(dcsupplier.BoundText <> "", val(dcsupplier.BoundText), Null)
    RsSavRec.Fields("suckno").value = IIf(txtsuckno.Text <> "", Trim(txtsuckno.Text), Null)
    RsSavRec.Fields("suckdateH").value = txtsuckdateH.value
    RsSavRec.Fields("suckdate").value = dpsuckdate.value
    
    RsSavRec.Fields("FromPlanneddateH").value = txtFromPlanneddateH.value
    RsSavRec.Fields("FromPlanneddate").value = txtFromPlanneddate.value
    
    RsSavRec.Fields("ToPlanneddateH").value = txtToPlanneddateH.value
    RsSavRec.Fields("ToPlanneddate").value = txtToPlanneddate.value
    
    RsSavRec.Fields("PlotNo").value = IIf(txtPlotNo.Text <> "", val(txtPlotNo.Text), Null)
    RsSavRec.Fields("Planned").value = IIf(txtPlanned.Text <> "", val(txtPlanned.Text), Null)
    RsSavRec.Fields("DisountAmount").value = IIf(txtDisountAmount.Text <> "", val(txtDisountAmount.Text), Null)
    
    
    RsSavRec.Fields("authorizationname").value = IIf(txtauthorizationname.Text <> "", Trim(txtauthorizationname.Text), Null)
    ''//19 08 2015
    RsSavRec.Fields("FromNo").value = IIf(Me.txtFrom.Text <> "", txtFrom.Text, Null)
    RsSavRec.Fields("ToNo").value = IIf(Me.txtTo.Text <> "", txtTo.Text, Null)
    ''////////////
         If SystemOptions.OpenAccountAqar = True Then
         ' Account_Code_dynamic = GetMyAccountCode("TblCustemers", "CusID", val(Me.dcsupplier.BoundText), "AccountAccountAqar")
         
                 If Me.TxtModFlg.Text = "N" Then
                    RsSavRec("AccounCode").value = ModAccounts.AddNewAccount(Account_Code_dynamic, Trim$(Me.txtaqarname.Text), True, False, txtaqarname.Text, , , , , , , , , , 1, 1, 1, 0, 0)
                 Else
                        If Not IsNull(RsSavRec("AccounCode").value) And Not (RsSavRec("AccounCode").value) = "" Then
                            ModAccounts.EditAccount RsSavRec("AccounCode").value, Me.txtaqarname.Text, Me.txtaqarname.Text, , , , , , , , , 1, 1, 1, 0, 0, , , , True
                        Else
                            RsSavRec("AccounCode").value = ModAccounts.AddNewAccount(Account_Code_dynamic, Trim$(Me.txtaqarname.Text), True, False, txtaqarname.Text, , , , , , , , , , 1, 1, 1, 0, 0)
                        End If
                        
                End If
         End If

    RsSavRec.update
  If Me.TxtModFlg.Text <> "E" Then
    Set RsDetails = New ADODB.Recordset
    StrSQL = "SELECT     *  from dbo.TblAqarDetai Where (1 = -1)"
    RsDetails.Open StrSQL, Cn, adOpenKeyset, adLockOptimistic, adCmdText
    
    If UnitsGrid.Rows > 1 Then
        With UnitsGrid
            For i = .FixedRows To .Rows - 1
                If .TextMatrix(i, .ColIndex("nameunittype")) <> "" Then
                    RsDetails.AddNew
                    RsDetails("Aqarid").value = val(TxtAqarid.Text)
                    If val(.TextMatrix(i, .ColIndex("id"))) <> 0 Then
                        RsDetails("id").value = val(.TextMatrix(i, .ColIndex("id")))
                    Else
                        RsDetails("id").value = new_id("TblAqarDetai", "id", "")
                    End If
                    RsDetails("MiniRentValue").value = val((.TextMatrix(i, .ColIndex("MiniRentValue"))))
                    RsDetails("UnitElectric").value = (.TextMatrix(i, .ColIndex("UnitElectric")))
                    RsDetails("unitno").value = (.TextMatrix(i, .ColIndex("unitno")))
                    RsDetails("namerentType").value = (.TextMatrix(i, .ColIndex("namerentType")))
                    RsDetails("unittype").value = val(.TextMatrix(i, .ColIndex("unittype")))
                    RsDetails("rentType").value = val(.TextMatrix(i, .ColIndex("rentType")))
                    RsDetails("length").value = .TextMatrix(i, .ColIndex("length"))
                    RsDetails("meterPrice").value = val(.TextMatrix(i, .ColIndex("meterPrice")))
                    RsDetails("roomscount").value = val(.TextMatrix(i, .ColIndex("roomscount")))
                    RsDetails("WCcount").value = val(.TextMatrix(i, .ColIndex("WCcount")))
                    RsDetails("kithchencount").value = val(.TextMatrix(i, .ColIndex("kithchencount")))
                    If .Cell(flexcpChecked, i, .ColIndex("haveFurniture")) = flexChecked Then
                        RsDetails("haveFurniture").value = -1
                    Else
                        RsDetails("haveFurniture").value = 0
                    End If
                    If .Cell(flexcpChecked, i, .ColIndex("IsTax")) = flexChecked Then
                        RsDetails("IsTax").value = -1
                    Else
                        RsDetails("IsTax").value = 0
                    End If
                        
                    RsDetails("Status") = val(.TextMatrix(i, .ColIndex("StatusId")))
                    RsDetails("unitdesc").value = .TextMatrix(i, .ColIndex("unitdesc"))
                    RsDetails("RentValue").value = val(.TextMatrix(i, .ColIndex("RentValue")))
                    RsDetails("customerid").value = val(.TextMatrix(i, .ColIndex("customerid")))
                    RsDetails("Floor").value = (.TextMatrix(i, .ColIndex("Floor")))
                    RsDetails("LoungeCount").value = val(.TextMatrix(i, .ColIndex("LoungeCount")))
                    RsDetails("ACCount").value = val(.TextMatrix(i, .ColIndex("ACCount")))
                    RsDetails("ACCountspleat").value = val(.TextMatrix(i, .ColIndex("ACCountspleat")))
                    RsDetails.update
                End If
            Next i
        End With
    End If
End If
    Set RsDetails1 = New ADODB.Recordset
    StrSQL = "SELECT     *  from dbo.TblAqarDetai2 Where (1 = -1)"
    RsDetails1.Open StrSQL, Cn, adOpenKeyset, adLockOptimistic, adCmdText

    If VSFlexGrid2.Rows > 1 Then
        With VSFlexGrid2
            For i = .FixedRows To .Rows - 1
                If .TextMatrix(i, .ColIndex("ElevatorNo")) <> "" Then
                    RsDetails1.AddNew
                    RsDetails1("Aqarid").value = val(TxtAqarid.Text)
                    RsDetails1("ElevatorNo").value = .TextMatrix(i, .ColIndex("ElevatorNo"))
                    RsDetails1("Elevatortype").value = .TextMatrix(i, .ColIndex("Elevatortype"))
                    RsDetails1("company").value = .TextMatrix(i, .ColIndex("company"))
                    RsDetails1("BuildNo").value = .TextMatrix(i, .ColIndex("BuildNo"))
                    RsDetails1("MainCo").value = .TextMatrix(i, .ColIndex("MainCo"))
                    RsDetails1("MaintStrDate").value = IIf(IsDate(.TextMatrix(i, .ColIndex("MaintStrDate"))), .TextMatrix(i, .ColIndex("MaintStrDate")), Date)
                    RsDetails1("MaintEndDate").value = IIf(IsDate(.TextMatrix(i, .ColIndex("MaintEndDate"))), .TextMatrix(i, .ColIndex("MaintEndDate")), Date)  '.TextMatrix(i, .ColIndex("MaintEndDate"))
                    RsDetails1.update
                End If
           Next i
        End With
    End If
    
    Set RsDetails2 = New ADODB.Recordset
    StrSQL = "SELECT     *  from dbo.TblAqarDetai3 Where (1 = -1)"
    RsDetails2.Open StrSQL, Cn, adOpenKeyset, adLockOptimistic, adCmdText
    
    If VSFlexGrid1.Rows > 1 Then
        With VSFlexGrid1
            For i = .FixedRows To .Rows - 1
                If .TextMatrix(i, .ColIndex("Waset")) <> "" Then
                    RsDetails2.AddNew
                    RsDetails2("Aqarid").value = val(TxtAqarid.Text)
                    RsDetails2("Waset").value = .TextMatrix(i, .ColIndex("Waset"))
                    RsDetails2("Rate").value = .TextMatrix(i, .ColIndex("Rate"))
                    RsDetails2.update
                End If
            Next i
        End With
    End If
    
    Set RsDetails3 = New ADODB.Recordset
    StrSQL = "SELECT     *  from dbo.TblAqrOwin Where (1 = -1)"
    RsDetails3.Open StrSQL, Cn, adOpenKeyset, adLockOptimistic, adCmdText

    If GridInstallments.Rows > 1 Then
        With GridInstallments
            For i = .FixedRows To .Rows - 1
                If .TextMatrix(i, .ColIndex("RecDate")) <> "" Then
                    RsDetails3.AddNew
                    RsDetails3("AqrID").value = val(TxtAqarid.Text)
                    RsDetails3("RecDate").value = IIf(IsDate(.TextMatrix(i, .ColIndex("RecDate"))), .TextMatrix(i, .ColIndex("RecDate")), Date)
                    RsDetails3("AllowDate").value = IIf(IsDate(.TextMatrix(i, .ColIndex("AllowDate"))), .TextMatrix(i, .ColIndex("AllowDate")), Date)
                    RsDetails3("value").value = val(.TextMatrix(i, .ColIndex("value")))
                    RsDetails3("RecDateH").value = .TextMatrix(i, .ColIndex("RecDateH"))
                    RsDetails3("Cont").value = val(.TextMatrix(i, .ColIndex("Cont")))
                    RsDetails3("PaymentNo").value = val(.TextMatrix(i, .ColIndex("PaymentNo")))
                    RsDetails3("AllowDateH").value = .TextMatrix(i, .ColIndex("AllowDateH"))
                    RsDetails3("DMY").value = val(.TextMatrix(i, .ColIndex("DMY")))
                    RsDetails3("valuewithout").value = val(.TextMatrix(i, .ColIndex("valuewithout")))
                    
                    RsDetails3("ValueAfterDiscount").value = val(.TextMatrix(i, .ColIndex("ValueAfterDiscount")))
                    RsDetails3("Discount").value = val(.TextMatrix(i, .ColIndex("Discount")))
                    
                    RsDetails3("VatPerc").value = val(.TextMatrix(i, .ColIndex("VatPerc")))
                    RsDetails3("VATValue").value = val(.TextMatrix(i, .ColIndex("VATValue")))
                    
                    


                    RsDetails3.update
                End If
            Next i
        End With
    End If
        
    MsgBox " „  ⁄„·Ì… «·Õ›Ÿ »‰Ã«Õ", vbInformation + vbOKOnly + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
    FiLLTXT
    'FillGridWithData
    CuurentLogdata
    TxtModFlg = "R"

    Exit Sub
ErrTrap:
    If RsSavRec.EditMode <> adEditNone Then
        RsSavRec.CancelUpdate
    End If
End Sub
Sub UpdateRow(Optional i As Long, Optional ID As Double)

    Dim haveFurniture As Integer
    Dim isTax As Integer
    Dim sql As String
    
    With Me.UnitsGrid
        If .Cell(flexcpChecked, i, .ColIndex("haveFurniture")) = flexChecked Then
          haveFurniture = -1
        Else
          haveFurniture = 0
        End If
        
        If .Cell(flexcpChecked, i, .ColIndex("isTax")) = flexChecked Then
          isTax = 1
        Else
          isTax = 0
        End If

        sql = "Update TblAqarDetai set MiniRentValue=" & val((.TextMatrix(i, .ColIndex("MiniRentValue")))) & " ,UnitElectric='" & .TextMatrix(i, .ColIndex("UnitElectric")) & "'"
        sql = sql & " , unitno='" & .TextMatrix(i, .ColIndex("unitno")) & "', namerentType='" & .TextMatrix(i, .ColIndex("namerentType")) & "' "
        sql = sql & ", unittype=" & val(.TextMatrix(i, .ColIndex("unittype"))) & ",rentType=" & val(.TextMatrix(i, .ColIndex("rentType"))) & ""
        sql = sql & ", length='" & .TextMatrix(i, .ColIndex("length")) & "' ,meterPrice=" & val(.TextMatrix(i, .ColIndex("meterPrice"))) & "          "
        sql = sql & ", roomscount=" & val(.TextMatrix(i, .ColIndex("roomscount"))) & " ,WCcount=" & val(.TextMatrix(i, .ColIndex("WCcount"))) & "       "
        sql = sql & ",kithchencount=" & val(.TextMatrix(i, .ColIndex("kithchencount"))) & ",haveFurniture=" & haveFurniture & ",isTax = " & isTax
        sql = sql & ",Status=" & val(.TextMatrix(i, .ColIndex("StatusId"))) & ",RentValue=" & val(.TextMatrix(i, .ColIndex("RentValue"))) & "          "
        sql = sql & ", unitdesc='" & .TextMatrix(i, .ColIndex("unitdesc")) & "' ,Floor='" & .TextMatrix(i, .ColIndex("Floor")) & "' "
        sql = sql & ", customerid=" & val(.TextMatrix(i, .ColIndex("customerid"))) & " ,LoungeCount=" & val(.TextMatrix(i, .ColIndex("LoungeCount"))) & "          "
        sql = sql & ",ACCount=" & val(.TextMatrix(i, .ColIndex("ACCount"))) & "  ,ACCountspleat=" & val(.TextMatrix(i, .ColIndex("ACCountspleat"))) & "         "
        sql = sql & " where Aqarid =" & val(TxtAqarid.Text) & " and ID=" & ID & " "
        Cn.Execute sql
        MsgBox " „ «· ⁄œÌ· »‰Ã«Õ"
    End With
End Sub
Private Sub Calculations(Optional WithMsg As Boolean = True)

    On Error GoTo ErrTrap
    
    Dim i  As Integer
    Dim IntNoOFQast As Integer
    Dim FirstDate As Date
    Dim PreDate As Date
    Dim NewDate As Date
    Dim NewDate2 As Date
    Dim NewDateH2 As String
    Dim PreDateH As String
    Dim DateInterva2l As String
    Dim DateInterval As String
    Dim NewDateH As String
    Dim DateNumber As Integer
    Dim DateNumber2 As Integer
    Dim Msg As String

    If TxtPaymentCount.Text = "" Then
        Msg = "ÌÃ» ≈œŒ«· ⁄œœ «·√ﬁ”«ÿ"
        If WithMsg = True Then
            MsgBox Msg, vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
            TxtPaymentCount.SetFocus
        End If
        Exit Sub
    End If
  

    If DcbPeriodsID.ListIndex = -1 Then
        Msg = "ÌÃ» ≈œŒ«·   «·› —… »Ì‰ «·«ﬁ”«ÿ"
        If WithMsg = True Then
            MsgBox Msg, vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
            DcbPeriodsID.SetFocus
        End If
        Exit Sub
    End If
    
    If Not IsNumeric(TxtPaymentCount.Text) Then
        Msg = " ⁄œœ «·√ﬁ”«ÿ ÌÃ» √‰ ÌﬂÊ‰ ﬁÌ„… —ﬁ„Ì…"
        If WithMsg = True Then
            MsgBox Msg, vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
            TxtPaymentCount.SetFocus
        End If
        Exit Sub
    End If

  '  SngAllValue = val(TxtContValue) '+ val(watervalue) + val(Electricity) + val(TxtPhone) + val(TxtEnternet)
     IntNoOFQast = val(TxtPaymentCount)

    If DcbPeriodsID.ListIndex = 0 Then
        DateInterval = "d"
    ElseIf DcbPeriodsID.ListIndex = 1 Then
        DateInterval = "M"
    ElseIf DcbPeriodsID.ListIndex = 2 Then
        DateInterval = "yyyy"
    End If
    
    If DcbPeriodsAlowID.ListIndex = 0 Then
        DateInterva2l = "d"
    ElseIf DcbPeriodsAlowID.ListIndex = 1 Then
        DateInterva2l = "M"
    ElseIf DcbPeriodsAlowID.ListIndex = 2 Then
        DateInterva2l = "yyyy"
    End If
    
    NewDate = FristPaymentDate.value
    NewDateH = FirstInstallDateH.value
     
    DateNumber = val(TxtPeriods.Text)
    DateNumber2 = val(TxtPriodAlow.Text)

  
  
    With Me.GridInstallments
        .Clear flexClearScrollable, flexClearEverything
        .Rows = .FixedRows + IntNoOFQast

        For i = 1 To IntNoOFQast
            DoEvents
            .TextMatrix(i, .ColIndex("Ser")) = i
            .TextMatrix(i, .ColIndex("RecDateH")) = hijriorJerojian
            .TextMatrix(i, .ColIndex("VatPerc")) = val(TxtFATYou.Text)
            .TextMatrix(i, .ColIndex("value")) = val(TxtContValue.Text) / IntNoOFQast
 
           .TextMatrix(i, .ColIndex("valuewithout")) = .TextMatrix(i, .ColIndex("value")) / (1 + val(TxtFATYou.Text) / 100)
           
           .TextMatrix(i, .ColIndex("VatValue")) = val(.TextMatrix(i, .ColIndex("valuewithout"))) * val(TxtFATYou.Text) / 100
         .TextMatrix(i, .ColIndex("VatPerc")) = val(TxtFATYou.Text)
            



            .TextMatrix(i, .ColIndex("PaymentNo")) = i
            .TextMatrix(i, .ColIndex("DMY")) = val(DcbPeriodsAlowID.ListIndex) + 1
            .TextMatrix(i, .ColIndex("Cont")) = val(TxtPriodAlow.Text)
            If i = 1 Then
              VBA.Calendar = vbCalGreg
                NewDate = FristPaymentDate
                        VBA.Calendar = vbCalHijri
                        
                NewDateH = FirstInstallDateH.value
                   VBA.Calendar = vbCalGreg
            Else
                PreDate = CDate(Trim(.TextMatrix(i - 1, .ColIndex("RecDate"))))
                If hijriorJerojian = 1 Then 'jorijan
                                                VBA.Calendar = vbCalGreg
                    NewDate = DateAdd(DateInterval, DateNumber, PreDate)
                    NewDateH = ToHijriDate(NewDate)
                End If
                
                PreDateH = (Trim(.TextMatrix(i - 1, .ColIndex("RecDateH"))))
     
                If hijriorJerojian = 0 Then 'hijri
                                                VBA.Calendar = vbCalHijri
                    NewDateH = (DateAdd(DateInterval, DateNumber, PreDateH))
                    NewDate = ToGregorianDate(NewDateH)
                End If
            End If
            
           .TextMatrix(i, .ColIndex("RecDate")) = Format(NewDate, "yyyy/M/d")
            .TextMatrix(i, .ColIndex("RecDateH")) = Format(NewDateH, "yyyy/M/d")
            PreDate = CDate(Trim(.TextMatrix(i, .ColIndex("RecDate"))))
            NewDate2 = DateAdd(DateInterva2l, DateNumber2, PreDate)
            NewDateH2 = ToHijriDate(NewDate2)
            .TextMatrix(i, .ColIndex("AllowDate")) = Format(NewDate2, "yyyy/M/d")
            .TextMatrix(i, .ColIndex("AllowDateH")) = Format(NewDateH2, "yyyy/M/d")
        Next i

        .AutoSize 1, .Cols - 1, False
    End With
    ReLineGrid
    Exit Sub
ErrTrap:
End Sub
Public Sub FiLLTXT()

    On Error GoTo ErrTrap
    
    Dim i As Integer
    loadcombo
    TxtAqarid.Text = IIf(IsNull(RsSavRec.Fields("Aqarid").value), "", RsSavRec.Fields("Aqarid").value)
    If Not IsNull(RsSavRec.Fields("TypAmola").value) Then
        If (RsSavRec.Fields("TypAmola").value) = 1 Then
            RSOutSupplier.value = vbChecked
            Rd.value = True
        Else
            RSOutSupplier.value = vbUnchecked
            Rd.value = False
        End If
    Else
        RSOutSupplier.value = vbUnchecked
        Rd.value = False
    End If
    DcbAccount.BoundText = IIf(IsNull(RsSavRec("AccounCode").value), "", RsSavRec("AccounCode").value)

        
        
            If Not IsNull(RsSavRec("TypeDate").value) Then
                If RsSavRec("TypeDate").value = 1 Then
                         RdRTypeDate(1).value = True
                Else
                           RdRTypeDate(0).value = True
                End If
    Else
                          RdRTypeDate(0).value = True
    End If
    

     
         Me.TxtNoteSerial.Text = IIf(IsNull(RsSavRec.Fields("NoteSerial").value), "", RsSavRec.Fields("NoteSerial").value)
         
       Me.TxtContValueWithout.Text = IIf(IsNull(RsSavRec.Fields("TxtContValueWithout").value), 0, RsSavRec.Fields("TxtContValueWithout").value)
       Me.TxtFATYou.Text = IIf(IsNull(RsSavRec.Fields("TxtFATYou").value), 0, RsSavRec.Fields("TxtFATYou").value)
       Me.TxtFATValue.Text = IIf(IsNull(RsSavRec.Fields("TxtFATValue").value), 0, RsSavRec.Fields("TxtFATValue").value)
       txtManulaVat.Text = IIf(IsNull(RsSavRec("TxtFATYou").value), 0, (RsSavRec("TxtFATYou").value))

                
 
Me.TXTNoteID.Text = IIf(IsNull(RsSavRec.Fields("NoteID").value), "", RsSavRec.Fields("NoteID").value)
     


        
        
           If Not IsNull(RsSavRec.Fields("ComResid").value) Then
                    If RsSavRec.Fields("ComResid").value = 1 Then
                                  ComResid(1).value = True
                    Else
                              ComResid(0).value = True
                    End If
   Else
                ComResid(0).value = True
   End If
   
    Me.TxtKickbacks.Text = IIf(IsNull(RsSavRec.Fields("AmolaValus").value), 0, RsSavRec.Fields("AmolaValus").value)
    Me.DcbSales.BoundText = IIf(IsNull(RsSavRec("SalesEmp").value), "", RsSavRec("SalesEmp").value)
    dcBranch.BoundText = IIf(IsNull(RsSavRec("BranchId").value), "", RsSavRec("BranchId").value)
    Me.TxtRat.Text = IIf(IsNull(RsSavRec.Fields("Rate").value), "", RsSavRec.Fields("Rate").value)
    Me.TxtUnit.Text = IIf(IsNull(RsSavRec.Fields("UnitNo").value), "", RsSavRec.Fields("UnitNo").value)
    Me.TxtBlock.Text = IIf(IsNull(RsSavRec.Fields("Block").value), "", RsSavRec.Fields("Block").value)
    Me.DcFixedAssets.BoundText = IIf(IsNull(RsSavRec.Fields("FixedID").value), "", RsSavRec.Fields("FixedID").value)
    Me.TxtBanckName.Text = IIf(IsNull(RsSavRec.Fields("BanckName").value), "", RsSavRec.Fields("BanckName").value)
    Me.TxtProvide.Text = IIf(IsNull(RsSavRec.Fields("Provide").value), "", RsSavRec.Fields("Provide").value)
    Me.TxtRemarks.Text = IIf(IsNull(RsSavRec.Fields("Remarks").value), "", RsSavRec.Fields("Remarks").value)
    Me.TxteastWriiten.Text = IIf(IsNull(RsSavRec.Fields("eastWriiten").value), "", RsSavRec.Fields("eastWriiten").value)
    Me.TxtwestWriiten.Text = IIf(IsNull(RsSavRec.Fields("westWriiten").value), "", RsSavRec.Fields("westWriiten").value)
    Me.TxtPart.Text = IIf(IsNull(RsSavRec.Fields("Part").value), "", RsSavRec.Fields("Part").value)
    Me.TxtValYearIncrease.Text = IIf(IsNull(RsSavRec.Fields("ValYearIncrease").value), "", RsSavRec.Fields("ValYearIncrease").value)
     Me.NOOFYears.Text = IIf(IsNull(RsSavRec.Fields("NOOFYears").value), 0, RsSavRec.Fields("NOOFYears").value)
     
    
    
    Me.TxtStreet.Text = IIf(IsNull(RsSavRec.Fields("StreetNo").value), "", RsSavRec.Fields("StreetNo").value)
    Me.TxtPriceSomW.Text = IIf(IsNull(RsSavRec.Fields("PriceSomW").value), "", RsSavRec.Fields("PriceSomW").value)
    Me.TxtPriceHadW.Text = IIf(IsNull(RsSavRec.Fields("PriceHadW").value), "", RsSavRec.Fields("PriceHadW").value)
    Me.TxtPrice.Text = IIf(IsNull(RsSavRec.Fields("Price").value), "", RsSavRec.Fields("Price").value)
    TxtPriceSom.Text = IIf(IsNull(RsSavRec.Fields("PriceSom").value), "", RsSavRec.Fields("PriceSom").value)
    Me.TxtPriceHad.Text = IIf(IsNull(RsSavRec.Fields("PriceHad").value), "", RsSavRec.Fields("PriceHad").value)
    txtaqarNo.Text = IIf(IsNull(RsSavRec.Fields("aqarNo").value), "", RsSavRec.Fields("aqarNo").value)
    Me.dcaqartypeid.BoundText = IIf(IsNull(RsSavRec.Fields("aqartypeid").value), "", RsSavRec.Fields("aqartypeid").value)
    '''///salah
    TxtAqarid.Text = IIf(IsNull(RsSavRec.Fields("Aqarid").value), "", RsSavRec.Fields("Aqarid").value)
    ''////
    TxtagencyNo.Text = IIf(IsNull(RsSavRec.Fields("AgemcyNo").value), "", RsSavRec.Fields("AgemcyNo").value)
    txtTel.Text = IIf(IsNull(RsSavRec.Fields("Telephone").value), "", RsSavRec.Fields("Telephone").value)
    TxtMobile.Text = IIf(IsNull(RsSavRec.Fields("Mobile").value), "", RsSavRec.Fields("Mobile").value)
    TxtEmail.Text = IIf(IsNull(RsSavRec.Fields("Email").value), "", RsSavRec.Fields("Email").value)
    TxtFaxAg.Text = IIf(IsNull(RsSavRec.Fields("Fax").value), "", RsSavRec.Fields("Fax").value)
    TxtAcountBank.Text = IIf(IsNull(RsSavRec.Fields("AccountBank").value), "", RsSavRec.Fields("AccountBank").value)
    TxtContValue.Text = val(IIf(IsNull(RsSavRec.Fields("ContValue").value), 0, RsSavRec.Fields("ContValue").value))
    TxtPaymentCount.Text = IIf(IsNull(RsSavRec.Fields("PaymentNo").value), "", RsSavRec.Fields("PaymentNo").value)
    TxtPeriods.Text = val(IIf(IsNull(RsSavRec.Fields("Priod").value), "", RsSavRec.Fields("Priod").value))
    TxtPriodAlow.Text = val(IIf(IsNull(RsSavRec.Fields("PriodAlow").value), "", RsSavRec.Fields("PriodAlow").value))
    Me.DcbPeriodsID.ListIndex = IIf(IsNull(RsSavRec.Fields("PriodDMY").value), -1, RsSavRec.Fields("PriodDMY").value)
    Me.DcbPeriodsAlowID.ListIndex = IIf(IsNull(RsSavRec.Fields("PriodAlowDMY").value), -1, RsSavRec.Fields("PriodAlowDMY").value)
    DateCont.value = IIf(IsNull(RsSavRec.Fields("DateCont").value), Date, RsSavRec.Fields("DateCont").value)
    DateHCont.value = IIf(IsNull(RsSavRec.Fields("DateHCont").value), ToHijriDate(Date), RsSavRec.Fields("DateHCont").value)
    FirstInstallDateH.value = IIf(IsNull(RsSavRec.Fields("FirstInstallDateH").value), ToHijriDate(Date), RsSavRec.Fields("FirstInstallDateH").value)
    FristPaymentDate.value = IIf(IsNull(RsSavRec.Fields("FristPaymentDate").value), Date, RsSavRec.Fields("FristPaymentDate").value)
    FromCotDate.value = IIf(IsNull(RsSavRec.Fields("FromCotDate").value), Date, RsSavRec.Fields("FromCotDate").value)
    FromCotDateH.value = IIf(IsNull(RsSavRec.Fields("FromCotDateH").value), ToHijriDate(Date), RsSavRec.Fields("FromCotDateH").value)
    ToCotDate.value = IIf(IsNull(RsSavRec.Fields("ToCotDate").value), Date, RsSavRec.Fields("ToCotDate").value)
    ToCotDateH.value = IIf(IsNull(RsSavRec.Fields("ToCotDateH").value), ToHijriDate(Date), RsSavRec.Fields("ToCotDateH").value)
    txtaqarname.Text = IIf(IsNull(RsSavRec.Fields("aqarname").value), "", RsSavRec.Fields("aqarname").value)
    Me.DcboCountryID2.BoundText = IIf(IsNull(RsSavRec.Fields("countryid").value), "", RsSavRec.Fields("countryid").value)
    Me.DcboGovernmentID.BoundText = IIf(IsNull(RsSavRec.Fields("cityid").value), "", RsSavRec.Fields("cityid").value)
    Me.DcboCityID.BoundText = IIf(IsNull(RsSavRec.Fields("heyid").value), "", RsSavRec.Fields("heyid").value)
    Me.dcschemeid.BoundText = IIf(IsNull(RsSavRec.Fields("schemeid").value), "", RsSavRec.Fields("schemeid").value)
    txtstreetname.Text = IIf(IsNull(RsSavRec.Fields("streetname").value), "", RsSavRec.Fields("streetname").value)
    txtlocation.Text = IIf(IsNull(RsSavRec.Fields("location").value), "", RsSavRec.Fields("location").value)
    txtaqarage.Text = IIf(IsNull(RsSavRec.Fields("aqarage").value), 0, RsSavRec.Fields("aqarage").value)
    Me.cbostatusid.ListIndex = IIf(IsNull(RsSavRec.Fields("statusid").value), -1, RsSavRec.Fields("statusid").value)
    Me.dcmaintenancetypeid.ListIndex = IIf(IsNull(RsSavRec.Fields("maintenancetypeid").value), -1, RsSavRec.Fields("maintenancetypeid").value)
    txtfloorcount.Text = IIf(IsNull(RsSavRec.Fields("floorcount").value), 0, RsSavRec.Fields("floorcount").value)
    txtcurrentPrice.Text = IIf(IsNull(RsSavRec.Fields("currentPrice").value), 0, RsSavRec.Fields("currentPrice").value)
    txtlastrentvalue.Text = IIf(IsNull(RsSavRec.Fields("lastrentvalue").value), 0, RsSavRec.Fields("lastrentvalue").value)
    Me.cbointerfaceid.ListIndex = IIf(IsNull(RsSavRec.Fields("interfaceid").value), -1, RsSavRec.Fields("interfaceid").value)
    txtnoofapartement.Text = IIf(IsNull(RsSavRec.Fields("noofapartement").value), 0, RsSavRec.Fields("noofapartement").value)
    txtnoofoffices.Text = IIf(IsNull(RsSavRec.Fields("noofoffices").value), 0, RsSavRec.Fields("noofoffices").value)
    txtnoofparking.Text = IIf(IsNull(RsSavRec.Fields("noofparking").value), 0, RsSavRec.Fields("noofparking").value)
    txtEntryCount.Text = IIf(IsNull(RsSavRec.Fields("EntryCount").value), 0, RsSavRec.Fields("EntryCount").value)
    txtnorthlength.Text = IIf(IsNull(RsSavRec.Fields("northlength").value), 0, RsSavRec.Fields("northlength").value)
    txtSouthlength.Text = IIf(IsNull(RsSavRec.Fields("Southlength").value), 0, RsSavRec.Fields("Southlength").value)
    txteastlength.Text = IIf(IsNull(RsSavRec.Fields("eastlength").value), 0, RsSavRec.Fields("eastlength").value)
    txtWestlength.Text = IIf(IsNull(RsSavRec.Fields("Westlength").value), 0, RsSavRec.Fields("Westlength").value)
    txttotallength.Text = IIf(IsNull(RsSavRec.Fields("totallength").value), 0, RsSavRec.Fields("totallength").value)
    txtmeterRentvalue.Text = IIf(IsNull(RsSavRec.Fields("meterRentvalue").value), 0, RsSavRec.Fields("meterRentvalue").value)
    txtmetersalevalue.Text = IIf(IsNull(RsSavRec.Fields("metersalevalue").value), 0, RsSavRec.Fields("metersalevalue").value)
    txtgooglemap.Text = IIf(IsNull(RsSavRec.Fields("googlemap").value), "", RsSavRec.Fields("googlemap").value)
    Me.dcsupplier.BoundText = IIf(IsNull(RsSavRec.Fields("ownerid").value), "", RsSavRec.Fields("ownerid").value)
    txtsuckno.Text = IIf(IsNull(RsSavRec.Fields("suckno").value), "", RsSavRec.Fields("suckno").value)
    txtauthorizationname.Text = IIf(IsNull(RsSavRec.Fields("authorizationname").value), "", RsSavRec.Fields("authorizationname").value)
    dpsuckdate.value = IIf(IsNull(RsSavRec.Fields("suckdate").value), Date, RsSavRec.Fields("suckdate").value)
    txtsuckdateH.value = IIf(IsNull(RsSavRec.Fields("suckdateH").value), ToHijriDate(Date), RsSavRec.Fields("suckdateH").value)
    
    

     txtFromPlanneddate.value = IIf(IsNull(RsSavRec.Fields("FromPlanneddate").value), Date, RsSavRec.Fields("FromPlanneddate").value)
    txtFromPlanneddateH.value = IIf(IsNull(RsSavRec.Fields("FromPlanneddateH").value), ToHijriDate(Date), RsSavRec.Fields("FromPlanneddateH").value)
    
    
     txtToPlanneddate.value = IIf(IsNull(RsSavRec.Fields("ToPlanneddate").value), Date, RsSavRec.Fields("ToPlanneddate").value)
    txtToPlanneddateH.value = IIf(IsNull(RsSavRec.Fields("ToPlanneddateH").value), ToHijriDate(Date), RsSavRec.Fields("ToPlanneddateH").value)
    
    txtPlotNo.Text = IIf(IsNull(RsSavRec.Fields("PlotNo").value), "", RsSavRec.Fields("PlotNo").value)
    txtPlanned.Text = IIf(IsNull(RsSavRec.Fields("Planned").value), "", RsSavRec.Fields("Planned").value)
    txtDisountAmount.Text = IIf(IsNull(RsSavRec.Fields("DisountAmount").value), "", RsSavRec.Fields("DisountAmount").value)
    
     
    
    
    ''// 19 08 2015
  '  Me.txtFrom.Text = IIf(IsNull(RsSavRec.Fields("FromNo").value), "", RsSavRec.Fields("FromNo").value)
  '  Me.txtto.Text = IIf(IsNull(RsSavRec.Fields("ToNo").value), "", RsSavRec.Fields("ToNo").value)
    LabCurrRec.Caption = RsSavRec.AbsolutePosition
    LabCountRec.Caption = RsSavRec.RecordCount
    TxtAccoup.Text = GetValueAssest
    With Grid
        FillGridWithData
        For i = 1 To .Rows - 1
            If Trim(TxtAqarid.Text) = .TextMatrix(i, .ColIndex("Aqarid")) Then
                TxtAqarid.Text = .TextMatrix(i, .ColIndex("Ser"))
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
Private Sub FristPaymentDate_Change()
    If Me.TxtModFlg.Text <> "R" Then
         FirstInstallDateH.value = ToHijriDate(FristPaymentDate.value)
    End If
End Sub
Private Sub FristPaymentDate_GotFocus()
    hijriorJerojian = 1
End Sub
Private Sub FromCotDate_Change()
    If Me.TxtModFlg.Text <> "R" Then
        FromCotDateH.value = ToHijriDate(FromCotDate.value)
    End If
End Sub
Private Sub FromCotDateH_LostFocus()
    VBA.Calendar = vbCalGreg
    FromCotDate.value = ToGregorianDate(FromCotDateH.value)
End Sub
Private Sub Grid_EnterCell()

    On Error GoTo ErrTrap
    
    FindRec val(Me.Grid.TextMatrix(Me.Grid.Row, Me.Grid.ColIndex("Aqarid")))
ErrTrap:
End Sub
Private Sub GridInstallments_AfterEdit(ByVal Row As Long, ByVal Col As Long)
    Dim PreDate As Date
    Dim NewDate2 As Date
    Dim NewDateH2 As String
    Dim PreDateH As String
    Dim DateInterva2l As String
    Dim DateNumber2 As Integer
      
    With GridInstallments
    
    Select Case .ColKey(Col)
Case "NetWater"
 
Case "NetElectric"
.TextMatrix(Row, .ColIndex("Electric")) = .TextMatrix(Row, .ColIndex("NetElectric"))
End Select


        'If .TextMatrix(Row, .ColIndex("RecDate")) <> "" Then
        '    If val((.TextMatrix(Row, .ColIndex("DMY")))) = 1 Then
        '        DateInterva2l = "d"
        '    ElseIf val((.TextMatrix(Row, .ColIndex("DMY")))) = 2 Then
        '        DateInterva2l = "M"
        '    ElseIf val((.TextMatrix(Row, .ColIndex("DMY")))) = 3 Then
        '        DateInterva2l = "yyyy"
        '    End If
        '    DateNumber2 = val((.TextMatrix(Row, .ColIndex("Cont"))))
        '    PreDate = CDate(Trim(.TextMatrix(Row, .ColIndex("RecDate"))))
        '    NewDate2 = DateAdd(DateInterva2l, DateNumber2, PreDate)
        '    NewDateH2 = ToHijriDate(NewDate2)
        '    .TextMatrix(Row, .ColIndex("AllowDate")) = Format(NewDate2, "yyyy/M/d")
        '    .TextMatrix(Row, .ColIndex("AllowDateH")) = Format(NewDateH2, "yyyy/M/d")
        'End If
        
        
    End With
End Sub
Private Sub ImgFavorites_Click()
    AddTofaforites Me.Name, Me.Caption, Me.Caption
End Sub

Private Sub Rd_Click()
    If Me.Rd.value = True Then
        TxtKickbacks.Enabled = True
    Else
        TxtKickbacks.Enabled = False
    End If
End Sub

 

Private Sub Text1_KeyPress(Index As Integer, KeyAscii As Integer)
    If Index = 10 Then
        Dim EmpID As Integer
        If KeyAscii = vbKeyReturn Then
            GetTblCustemersCode Text1(10).Text, EmpID, , , 57
            dcsupplier.BoundText = EmpID
        End If
    End If
End Sub

Private Sub txtManulaVat_Change()

If Me.TxtModFlg.Text <> "R" Then
TxtFATYou.Text = txtManulaVat.Text
ClculteVAT

End If
End Sub

Private Sub ToCotDate_Change()
    If Me.TxtModFlg.Text <> "R" Then
        ToCotDateH.value = ToHijriDate(ToCotDate.value)
    End If
End Sub
Private Sub ToCotDateH_LostFocus()
    If Me.TxtModFlg.Text <> "R" Then
        VBA.Calendar = vbCalGreg
        ToCotDate.value = ToGregorianDate(ToCotDateH.value)
    End If
End Sub
Sub GetAsseteCode_ID(Optional ByRef ID As Double = 0, Optional ByRef Fullcode As String = "", Optional Typ As Integer = 0)

    Dim sql As String
    Dim Rs7 As ADODB.Recordset
    Set Rs7 = New ADODB.Recordset

    If Typ = 0 Then
        sql = "select Fullcode  from FixedAssets where id=" & ID & " "
        Rs7.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
        If Rs7.RecordCount > 0 Then
            Fullcode = IIf(IsNull(Rs7("Fullcode").value), "", Rs7("Fullcode").value)
        Else
            Fullcode = ""
        End If
    Else
        sql = "select ID  from FixedAssets where Fullcode='" & Fullcode & "' "
        Rs7.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
        If Rs7.RecordCount > 0 Then
            ID = IIf(IsNull(Rs7("ID").value), 0, Rs7("ID").value)
        Else
            ID = 0
        End If
    End If
End Sub
Private Sub TxAssetscode_KeyPress(KeyAscii As Integer)
    Dim AsseID As Double
    GetAsseteCode_ID AsseID, TxAssetscode.Text, 1
    DcFixedAssets.BoundText = AsseID
End Sub
Private Sub TxtAqarid_Change()
    Dim TxtMod As String
    TxtMod = TxtModFlg.Text
    TxtModFlg.Text = ""
    TxtModFlg = TxtMod
End Sub
Public Function FindRec(ByVal RecId As Long)

    On Error GoTo ErrTrap
    
    RsSavRec.Find "Aqarid=" & RecId, , adSearchForward, 1

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
Sub CheCotIqar(ByRef bo As Boolean)

    Dim RsDetails As ADODB.Recordset
    Dim StrSQL As String
    Set RsDetails = New ADODB.Recordset
    StrSQL = "SELECT     *  from dbo.TblContract Where (Iqar =" & val(val(TxtAqarid.Text)) & ")"
    RsDetails.Open StrSQL, Cn, adOpenKeyset, adLockOptimistic, adCmdText
    If RsDetails.RecordCount > 0 Then
        bo = True
    End If
End Sub
Private Sub txtaqarname_Change()
    LblCaption.Caption = txtaqarname.Text
End Sub
Private Sub TxtCodeSales_Change()
    If Me.TxtModFlg = "N" Or Me.TxtModFlg = "E" Then
        DcbSales.BoundText = GeTEmpIDByEmpCode(TxtCodeSales.Text, True)
    End If
End Sub

Private Sub TxtContValueWithout_Change()
ClculteVAT
End Sub

Private Sub TXtFrom_KeyPress(KeyAscii As Integer)
    KeyAscii = KeyAscii_Num(KeyAscii, Me.txtFrom.Text, 1)
End Sub

Private Sub TxtModFlg_Change()

    If TxtModFlg.Text = "N" Then
    Ele(11).Enabled = True
    Cmd(20).Enabled = False
    Editbtn.Enabled = False
    CmdDelete.Enabled = False
        Me.btnNew.Enabled = False
        btnModify.Enabled = False
        btnDelete.Enabled = False
        Me.btnQuery.Enabled = False
        Grid.Enabled = False
        BtnUndo.Enabled = True
        Me.btnSave.Enabled = True
        BtnUpdate.Enabled = False
    ElseIf TxtModFlg.Text = "R" Then
    Ele(11).Enabled = False
        Cmd(20).Enabled = True
        Editbtn.Enabled = True
        CmdDelete.Enabled = True
        Grid.Enabled = True
        btnModify.Enabled = False
        btnDelete.Enabled = False
        If TxtAqarid.Text <> "" Then
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
    Ele(11).Enabled = True
        Cmd(20).Enabled = False
        Editbtn.Enabled = False
        CmdDelete.Enabled = False
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
    Dim rs As ADODB.Recordset
    Dim Rs1 As ADODB.Recordset
    Dim My_SQL As String
    ''''///sala
    ' RsDetails("LoungeCount").value = val(.TextMatrix(i, .ColIndex("LoungeCount")))
         
    Set rs = New ADODB.Recordset
    
    My_SQL = " SELECT    dbo.TblAqarDetai.Typed,TblAqarDetai.IsTax,dbo.TblAqarDetai.Status, dbo.TblAqarDetai.Id, dbo.TblAqarDetai.Aqarid, dbo.TblAqarDetai.length, dbo.TblAqarDetai.unitdesc, dbo.TblAqarDetai.unitno, "
    My_SQL = My_SQL & "                  dbo.TblAqarDetai.rentType, dbo.TblAqarDetai.meterPrice, dbo.TblAqarDetai.roomscount, dbo.TblAqarDetai.WCcount, dbo.TblAqarDetai.kithchencount,"
    My_SQL = My_SQL & "                   dbo.TblAqarDetai.RentValue, dbo.TblAqarDetai.haveFurniture, dbo.TblAqarDetai.LoungeCount, dbo.TblAqarDetai.namerentType, dbo.TblCustemers.CusName,"
    My_SQL = My_SQL & "                    dbo.TblCustemers.CusNamee, dbo.TblAqarDetai.unittype, dbo.TblAkarUnit.name, dbo.TblAkarUnit.namee, dbo.TblAqarDetai.customerid, dbo.TblAqarDetai.Floor,"
    My_SQL = My_SQL & "                     dbo.TblAqarDetai.ACCount, dbo.TblRentStatus.name AS statusname, dbo.TblRentStatus.namee AS statusnamee, dbo.TblAqarDetai.ACCountspleat,"
    My_SQL = My_SQL & "                     dbo.TblAqarDetai.UnitElectric , dbo.TblAqarDetai.Services, dbo.TblAqarDetai.Water ,dbo.TblAqarDetai.MiniRentValue"
    My_SQL = My_SQL & " FROM         dbo.TblAqarDetai LEFT OUTER JOIN"
    My_SQL = My_SQL & "                      dbo.TblRentStatus ON dbo.TblAqarDetai.Status = dbo.TblRentStatus.id LEFT OUTER JOIN"
    My_SQL = My_SQL & "                      dbo.TblAkarUnit ON dbo.TblAqarDetai.unittype = dbo.TblAkarUnit.id LEFT OUTER JOIN"
    My_SQL = My_SQL & "                      dbo.TblCustemers ON dbo.TblAqarDetai.customerid = dbo.TblCustemers.CusID"
    My_SQL = My_SQL & "  WHERE     (dbo.TblAqarDetai.Aqarid =" & val(Me.TxtAqarid.Text) & ")"
 
    If ChkOrder.value = vbChecked Then
        My_SQL = My_SQL & " order by dbo.TblAqarDetai.unitno "
    End If
    
    rs.Open My_SQL, Cn, adOpenKeyset, adLockReadOnly, adCmdText

    With Me.UnitsGrid
        .Rows = 1
        .Clear flexClearScrollable
        If rs.RecordCount > 0 Then
           .Rows = rs.RecordCount + 1
            rs.MoveFirst
            For i = 1 To .Rows - 1
                .TextMatrix(i, .ColIndex("Ser")) = i
                .TextMatrix(i, .ColIndex("Typed")) = IIf(IsNull(rs.Fields("Typed").value), 1, rs.Fields("Typed").value)
                .TextMatrix(i, .ColIndex("id")) = IIf(IsNull(rs.Fields("id").value), "", rs.Fields("id").value)
                .TextMatrix(i, .ColIndex("UnitElectric")) = IIf(IsNull(rs.Fields("UnitElectric").value), "", rs.Fields("UnitElectric").value)
                .TextMatrix(i, .ColIndex("MiniRentValue")) = IIf(IsNull(rs.Fields("MiniRentValue").value), "", rs.Fields("MiniRentValue").value)
                .TextMatrix(i, .ColIndex("unitno")) = IIf(IsNull(rs.Fields("unitno").value), "", rs.Fields("unitno").value)
                If SystemOptions.UserInterface = EnglishInterface Then
                    .TextMatrix(i, .ColIndex("customeridname")) = IIf(IsNull(rs.Fields("CusNamee").value), "", rs.Fields("CusNamee").value)
                    .TextMatrix(i, .ColIndex("nameunittype")) = IIf(IsNull(rs.Fields("namee").value), "", rs.Fields("namee").value)
                Else
                    .TextMatrix(i, .ColIndex("nameunittype")) = IIf(IsNull(rs.Fields("name").value), "", rs.Fields("name").value)
                    .TextMatrix(i, .ColIndex("customeridname")) = IIf(IsNull(rs.Fields("CusName").value), "", rs.Fields("CusName").value)
                End If
                 .TextMatrix(i, .ColIndex("IsTax")) = IIf(IsNull(rs("IsTax").value), "", rs("IsTax").value)
                .TextMatrix(i, .ColIndex("namerentType")) = IIf(IsNull(rs.Fields("namerentType").value), "", rs.Fields("namerentType").value)
                .TextMatrix(i, .ColIndex("rentType")) = IIf(IsNull(rs.Fields("rentType").value), "", rs.Fields("rentType").value)
                .TextMatrix(i, .ColIndex("unittype")) = IIf(IsNull(rs.Fields("unittype").value), "", rs.Fields("unittype").value)
                .TextMatrix(i, .ColIndex("length")) = IIf(IsNull(rs.Fields("length").value), "", rs.Fields("length").value)
                .TextMatrix(i, .ColIndex("meterPrice")) = IIf(IsNull(rs.Fields("meterPrice").value), "", rs.Fields("meterPrice").value)
                .TextMatrix(i, .ColIndex("roomscount")) = IIf(IsNull(rs.Fields("roomscount").value), "", rs.Fields("roomscount").value)
                .TextMatrix(i, .ColIndex("WCcount")) = IIf(IsNull(rs.Fields("WCcount").value), "", rs.Fields("WCcount").value)
                .TextMatrix(i, .ColIndex("kithchencount")) = IIf(IsNull(rs.Fields("kithchencount").value), "", rs.Fields("kithchencount").value)
                .TextMatrix(i, .ColIndex("unitdesc")) = IIf(IsNull(rs.Fields("unitdesc").value), "", rs.Fields("unitdesc").value)
                .TextMatrix(i, .ColIndex("RentValue")) = IIf(IsNull(rs.Fields("RentValue").value), "", rs.Fields("RentValue").value)
                .TextMatrix(i, .ColIndex("customerid")) = IIf(IsNull(rs.Fields("customerid").value), "", rs.Fields("customerid").value)
                .TextMatrix(i, .ColIndex("haveFurniture")) = IIf(IsNull(rs("haveFurniture").value), "", rs("haveFurniture").value)
                .TextMatrix(i, .ColIndex("Floor")) = IIf(IsNull(rs.Fields("Floor").value), "", rs.Fields("Floor").value)
                .TextMatrix(i, .ColIndex("LoungeCount")) = IIf(IsNull(rs.Fields("LoungeCount").value), "", rs.Fields("LoungeCount").value)
                .TextMatrix(i, .ColIndex("ACCount")) = IIf(IsNull(rs.Fields("ACCount").value), "", rs.Fields("ACCount").value)
                .TextMatrix(i, .ColIndex("ACCountspleat")) = IIf(IsNull(rs.Fields("ACCountspleat").value), "", rs.Fields("ACCountspleat").value)
                .TextMatrix(i, .ColIndex("StatusId")) = IIf(IsNull(rs.Fields("Status").value), 0, rs.Fields("Status").value)
                If SystemOptions.UserInterface = ArabicInterface Then
                    .TextMatrix(i, .ColIndex("Status")) = IIf(IsNull(rs.Fields("statusname").value), "‘«€—…", rs.Fields("statusname").value)
                Else
                    .TextMatrix(i, .ColIndex("Status")) = IIf(IsNull(rs.Fields("statusnamee").value), "empty", rs.Fields("statusnamee").value)
                End If
                rs.MoveNext
            Next
            rs.Close
            Set rs = Nothing
        End If
        .RowHeight(-1) = 300
    End With
    Set Rs1 = New ADODB.Recordset
    
    My_SQL = ""
    My_SQL = " SELECT     MaintEndDate, MaintStrDate, ElevatorNo, company, BuildNo, Elevatortype, MainCo, Aqarid, Id"
    My_SQL = My_SQL & " From dbo.TblAqarDetai2"
    My_SQL = My_SQL & " WHERE     (Aqarid =" & val(Me.TxtAqarid.Text) & ")"
    Rs1.Open My_SQL, Cn, adOpenKeyset, adLockReadOnly, adCmdText
    'rs1.Open My_SQL, Cn, adOpenStatic, adLockReadOnly, adCmdText
    With Me.VSFlexGrid2
        .Rows = 1
        .Clear flexClearScrollable
        If Rs1.RecordCount > 0 Then
           .Rows = Rs1.RecordCount + 1
            Rs1.MoveFirst
            For i = 1 To .Rows - 1
                .TextMatrix(i, .ColIndex("Ser")) = i
                .TextMatrix(i, .ColIndex("MainCo")) = IIf(IsNull(Rs1.Fields("MainCo").value), "", Rs1.Fields("MainCo").value)
                .TextMatrix(i, .ColIndex("Elevatortype")) = IIf(IsNull(Rs1.Fields("Elevatortype").value), "", Rs1.Fields("Elevatortype").value)
                .TextMatrix(i, .ColIndex("BuildNo")) = IIf(IsNull(Rs1.Fields("BuildNo").value), "", Rs1.Fields("BuildNo").value)
                 .TextMatrix(i, .ColIndex("company")) = IIf(IsNull(Rs1.Fields("company").value), "", Rs1.Fields("company").value)
                .TextMatrix(i, .ColIndex("ElevatorNo")) = IIf(IsNull(Rs1.Fields("ElevatorNo").value), "", Rs1.Fields("ElevatorNo").value)
                .TextMatrix(i, .ColIndex("MaintEndDate")) = IIf(IsNull(Rs1.Fields("MaintEndDate").value), "", Rs1.Fields("MaintEndDate").value)
                .TextMatrix(i, .ColIndex("MaintStrDate")) = IIf(IsNull(Rs1.Fields("MaintStrDate").value), "", Rs1.Fields("MaintStrDate").value)
                Rs1.MoveNext
            Next
            Rs1.Close
        End If
        .RowHeight(-1) = 300
    End With
    
    Set Rs1 = New ADODB.Recordset

    My_SQL = ""
    My_SQL = " SELECT  * from TblAqarDetai3"

    My_SQL = My_SQL & " WHERE     (Aqarid =" & val(Me.TxtAqarid.Text) & ")"
    Rs1.Open My_SQL, Cn, adOpenKeyset, adLockReadOnly, adCmdText
    'rs1.Open My_SQL, Cn, adOpenStatic, adLockReadOnly, adCmdText
    With Me.VSFlexGrid1
        .Rows = 1
        .Clear flexClearScrollable
        If Rs1.RecordCount > 0 Then
           .Rows = Rs1.RecordCount + 1
            Rs1.MoveFirst
            For i = 1 To .Rows - 1
                .TextMatrix(i, .ColIndex("Ser")) = i
                .TextMatrix(i, .ColIndex("Waset")) = IIf(IsNull(Rs1.Fields("Waset").value), "", Rs1.Fields("Waset").value)
                .TextMatrix(i, .ColIndex("Rate")) = IIf(IsNull(Rs1.Fields("Rate").value), "", Rs1.Fields("Rate").value)
                Rs1.MoveNext
            Next
            Rs1.Close
        End If
        .RowHeight(-1) = 300
    End With
    
  Set Rs1 = New ADODB.Recordset
  
    My_SQL = ""
    My_SQL = " SELECT  * from TblAqrOwin"
    My_SQL = My_SQL & " WHERE     (AqrID =" & val(Me.TxtAqarid.Text) & ")"
    Rs1.Open My_SQL, Cn, adOpenKeyset, adLockReadOnly, adCmdText
    'rs1.Open My_SQL, Cn, adOpenStatic, adLockReadOnly, adCmdText
    With Me.GridInstallments
        .Rows = 1
        .Clear flexClearScrollable
        If Rs1.RecordCount > 0 Then
           .Rows = Rs1.RecordCount + 1
            Rs1.MoveFirst
            For i = 1 To .Rows - 1
                .TextMatrix(i, .ColIndex("Ser")) = i
                .TextMatrix(i, .ColIndex("RecDateH")) = IIf(IsNull(Rs1.Fields("RecDateH").value), "", Rs1.Fields("RecDateH").value)
                .TextMatrix(i, .ColIndex("RecDate")) = IIf(IsNull(Rs1.Fields("RecDate").value), "", Rs1.Fields("RecDate").value)
                .TextMatrix(i, .ColIndex("value")) = IIf(IsNull(Rs1.Fields("value").value), "", Rs1.Fields("value").value)
                



                .TextMatrix(i, .ColIndex("valuewithout")) = IIf(IsNull(Rs1.Fields("valuewithout").value), 0, Rs1.Fields("valuewithout").value)
                
                .TextMatrix(i, .ColIndex("ValueAfterDiscount")) = IIf(IsNull(Rs1.Fields("ValueAfterDiscount").value), 0, Rs1.Fields("ValueAfterDiscount").value)
                .TextMatrix(i, .ColIndex("Discount")) = IIf(IsNull(Rs1.Fields("Discount").value), 0, Rs1.Fields("Discount").value)
                
                .TextMatrix(i, .ColIndex("VatPerc")) = IIf(IsNull(Rs1.Fields("VatPerc").value), 0, Rs1.Fields("VatPerc").value)
                .TextMatrix(i, .ColIndex("VatValue")) = IIf(IsNull(Rs1.Fields("VatValue").value), 0, Rs1.Fields("VatValue").value)
                
                If val(.TextMatrix(i, .ColIndex("valuewithout"))) = 0 Then
                 .TextMatrix(i, .ColIndex("valuewithout")) = .TextMatrix(i, .ColIndex("value"))
                End If
                
                
                
                .TextMatrix(i, .ColIndex("DMY")) = val(IIf(IsNull(Rs1.Fields("DMY").value), "", Rs1.Fields("DMY").value))
                .TextMatrix(i, .ColIndex("Cont")) = val(IIf(IsNull(Rs1.Fields("Cont").value), "", Rs1.Fields("Cont").value))
                .TextMatrix(i, .ColIndex("AllowDateH")) = IIf(IsNull(Rs1.Fields("AllowDateH").value), "", Rs1.Fields("AllowDateH").value)
                .TextMatrix(i, .ColIndex("AllowDate")) = IIf(IsNull(Rs1.Fields("AllowDate").value), "", Rs1.Fields("AllowDate").value)
                .TextMatrix(i, .ColIndex("PaymentNo")) = IIf(IsNull(Rs1.Fields("PaymentNo").value), "", Rs1.Fields("PaymentNo").value)
                Rs1.MoveNext
            Next
            Rs1.Close
        End If
        .RowHeight(-1) = 300
    End With

My_SQL = "  SELECT     oldunitNo,NewnitNo,  TransDate , dbo.TblUsers.UserName "
 My_SQL = My_SQL & "  FROM         dbo.UpdatesoldunitNo INNER JOIN"
My_SQL = My_SQL & "                       dbo.TblUsers ON dbo.UpdatesoldunitNo.UserId = dbo.TblUsers.UserID"
My_SQL = My_SQL & "  Where (dbo.UpdatesoldunitNo.Aqarid = " & val(Me.TxtAqarid.Text) & ")"
   loadgrid My_SQL, grdLoc, True, False
   
ErrTrap:
End Sub
Private Sub ShowTip()

    On Error GoTo ErrTrap
    
    Dim TTP As New clstooltip
    Dim Wrap As String
    Dim Msg As String
    
    Wrap = CHR(13) + CHR(10)

    With TTP
        .Create Me.hwnd, Me.Caption, 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 10000
        .DelayTime = 300
        Msg = "ÃœÌœ" & Wrap & "·› Õ ”Ã· ÃœÌœ " & Wrap & "≈÷€ÿ Â–« «·„› «Õ" & Wrap & "√Ê „› «Õ" & " F12 √Ê Enter"
            
        .AddControl btnNew, Msg, True
    End With

    With TTP
        .Create Me.hwnd, Me.Caption, 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 10000
        .DelayTime = 300
        Msg = " ⁄œÌ·" & Wrap & "· ⁄œÌ·  ”Ã· «·Õ«·Ï " & Wrap & "≈÷€ÿ Â–« «·„› «Õ" & Wrap & "√Ê „› «Õ" & " F11"
        .AddControl btnModify, Msg, True
    End With

    With TTP
        .Create Me.hwnd, Me.Caption, 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 10000
        .DelayTime = 300
        Msg = "Õ›Ÿ" & Wrap & "· ”ÃÌ· «·»Ì«‰«  œ«Œ· ﬁ«⁄œ… " & Wrap & "«·»Ì«‰«  ≈÷€ÿ Â–« «·„› «Õ" & Wrap & "√Ê „› «Õ" & " F10"
        .AddControl btnSave, Msg, True
    End With

    With TTP
        .Create Me.hwnd, Me.Caption, 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 10000
        .DelayTime = 300
        Msg = " —«Ã⁄" & Wrap & "·· —«Ã⁄ ⁄‰ «·⁄„·Ì… «·Õ«·Ì…" & Wrap & "≈÷€ÿ Â–« «·„› «Õ" & Wrap & "√Ê „› «Õ" & " F9"
        .AddControl BtnUndo, Msg, True
    End With

    With TTP
        .Create Me.hwnd, Me.Caption, 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 10000
        .DelayTime = 300
        Msg = "Õ–› «·”Ã·" & Wrap & "·Õ–› «·”Ã· «·Õ«·Ï" & Wrap & "≈÷€ÿ Â–« «·„› «Õ" & Wrap & "√Ê „› «Õ" & " F18"
        .AddControl btnDelete, Msg, True
    End With

    With TTP
        .Create Me.hwnd, Me.Caption, 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 10000
        .DelayTime = 300
        Msg = "Œ—ÊÃ" & Wrap & "·≈€·«ﬁ Â–Â «·‰«›–…" & Wrap & "≈÷€ÿ Â–« «·„› «Õ" & Wrap & "√Ê „› «Õ" & " Ctrl+x"
        .AddControl btnCancel, Msg, True
    End With

    With TTP
        .Create Me.hwnd, Me.Caption, 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 10000
        .DelayTime = 300
        Msg = "«·«Ê·" & Wrap & "··«‰ ﬁ«· «·Ï «·”Ã· «·«Ê·" & Wrap & "≈÷€ÿ Â–« «·„› «Õ" & Wrap & "√Ê „› «Õ" & " Home √Ê UpArrow"
        .AddControl btnFirst, Msg, True
    End With

    With TTP
        .Create Me.hwnd, Me.Caption, 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 10000
        .DelayTime = 300
        Msg = "«·”«»ﬁ" & Wrap & "··«‰ ﬁ«· «·Ï «·”Ã· «·”«»ﬁ" & Wrap & "≈÷€ÿ Â–« «·„› «Õ" & Wrap & "√Ê „› «Õ" & " PageUp √Ê LeftArrow"
        .AddControl btnPrevious, Msg, True
    End With

    With TTP
        .Create Me.hwnd, Me.Caption, 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 10000
        .DelayTime = 300
        Msg = "«· «·Ï" & Wrap & "··«‰ ﬁ«· «·Ï «·”Ã· «· «·Ï" & Wrap & "≈÷€ÿ Â–« «·„› «Õ" & Wrap & "√Ê „› «Õ" & " PageDown √Ê RightArrow"
        .AddControl btnNext, Msg, True
    End With

    With TTP
        .Create Me.hwnd, Me.Caption, 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 10000
        .DelayTime = 300
        Msg = "«·«ŒÌ—" & Wrap & "··«‰ ﬁ«· «·Ï «·”Ã· «·«ŒÌ—" & Wrap & "≈÷€ÿ Â–« «·„› «Õ" & Wrap & "√Ê „› «Õ" & " End √Ê DownArrow"
        .AddControl btnLast, Msg, True
    End With

ErrTrap:
End Sub

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

Private Sub txtsuckdateH_LostFocus()
    VBA.Calendar = vbCalGreg
    dpsuckdate.value = ToGregorianDate(txtsuckdateH.value)
End Sub


Private Sub txtFromPlanneddateH_LostFocus()
    VBA.Calendar = vbCalGreg
    txtFromPlanneddate.value = ToGregorianDate(txtFromPlanneddateH.value)
End Sub


Private Sub txtToPlanneddateH_LostFocus()
    VBA.Calendar = vbCalGreg
    txtToPlanneddate.value = ToGregorianDate(txtToPlanneddateH.value)
End Sub

Private Sub txtFromPlanneddate_Change()
    If Me.TxtModFlg.Text <> "R" Then
        txtFromPlanneddateH.value = ToHijriDate(txtFromPlanneddate.value)
    End If
End Sub
Private Sub txtToPlanneddate_Change()
    If Me.TxtModFlg.Text <> "R" Then
        txtToPlanneddateH.value = ToHijriDate(txtToPlanneddate.value)
    End If
End Sub
Private Sub txtto_KeyPress(KeyAscii As Integer)
    KeyAscii = KeyAscii_Num(KeyAscii, Me.txtTo.Text, 1)
End Sub
Private Sub UnitsGrid_AfterEdit(ByVal Row As Long, ByVal Col As Long)
Exit Sub
    Dim StrAccountCode As String
    Dim Msg As String
    Dim rs As New ADODB.Recordset
    Dim StrSQL As String
    Dim ClsAcc As New ClsAccounts
    Dim LngRow As Long

    With UnitsGrid
        Select Case .ColKey(Col)
            Case "customeridname"
                StrAccountCode = .ComboData
                LngRow = .FindRow(StrAccountCode, .FixedRows, .ColIndex("customerid"), False, True)
                .TextMatrix(Row, .ColIndex("customerid")) = StrAccountCode
            Case "nameunittype"
                StrAccountCode = .ComboData
                LngRow = .FindRow(StrAccountCode, .FixedRows, .ColIndex("unittype"), False, True)
                .TextMatrix(Row, .ColIndex("unittype")) = StrAccountCode
                If Me.TxtModFlg.Text = "E" Then
                    RegisterIqar .TextMatrix(Row, .ColIndex("unitno")), .TextMatrix(Row, .ColIndex("nameunittype"))
                End If
            Case "Status"
                StrAccountCode = .ComboData
                LngRow = .FindRow(StrAccountCode, .FixedRows, .ColIndex("Statusid"), False, True)
                .TextMatrix(Row, .ColIndex("Statusid")) = StrAccountCode
                If Me.TxtModFlg.Text = "E" Then
                    RegisterIqar .TextMatrix(Row, .ColIndex("unitno")), , , .TextMatrix(Row, .ColIndex("Status"))
                End If
            Case "MiniRentValue"
                If Me.TxtModFlg.Text = "E" Then
                    RegisterIqar .TextMatrix(Row, .ColIndex("unitno")), , .TextMatrix(Row, .ColIndex("MiniRentValue"))
                End If
            Case "namerentType"
                StrAccountCode = .ComboData
                LngRow = .FindRow(StrAccountCode, .FixedRows, .ColIndex("rentType"), False, True)
                .TextMatrix(Row, .ColIndex("rentType")) = StrAccountCode
        End Select
    End With
    If Row = UnitsGrid.Rows - 1 Then
            UnitsGrid.Rows = UnitsGrid.Rows + 1
    End If
    ReLineGrid
End Sub
Private Sub UnitsGrid_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
Exit Sub
    With UnitsGrid
        If .ColKey(Col) <> "Status" And .ColKey(Col) <> "customeridname" And .ColKey(Col) <> "nameunittype" And .ColKey(Col) <> "namerentType" Then
            UnitsGrid.ComboList = ""
        End If
        
        Select Case .ColKey(Col)
            Case "Floor"
                UnitsGrid.ComboList = ""
            Case "meterPrice"
                UnitsGrid.ComboList = ""
            Case "length"
                UnitsGrid.ComboList = ""
            Case "roomscount"
                UnitsGrid.ComboList = ""
            Case "WCcount"
                UnitsGrid.ComboList = ""
            Case "kithchencount"
                UnitsGrid.ComboList = ""
            Case "haveFurniture"
                UnitsGrid.ComboList = ""
            Case "isTax"
                UnitsGrid.ComboList = ""
            Case "unitdesc"
                UnitsGrid.ComboList = ""
            Case "customeridname"
                UnitsGrid.ComboList = ""
            Case "RentValue"
                Cancel = True
            Case "Status"
            'khaled
                'If SystemOptions.AllowChangeUnitIqar = False Then
                    'If (CheckUnitContract(val(UnitsGrid.TextMatrix(Row, UnitsGrid.ColIndex("id")))) = True Or CheckUnitContractMerg(val(UnitsGrid.TextMatrix(Row, UnitsGrid.ColIndex("id")))) = True) Then
                    'Cancel = True
                'Else
                '    UnitsGrid.ComboList = ""
                'End If
           Case "EditRow"
               ' If SystemOptions.AllowChangeUnitIqar = False Then
               '     Cancel = True
               ' Else
               '     UnitsGrid.ComboList = ""
               ' End If
        End Select
    End With
End Sub
Public Function RegisterIqar(Optional unitno As String, Optional unittype As String, Optional MiniRentValue As String = "", Optional Status As String = "")
    LogTextA = "—ﬁ„ «·ÊÕœ…" & unitno & CHR(13)
    If unittype <> "" Then
        LogTextA = LogTextA & "  ⁄œÌ·  ‰Ê⁄ «·ÊÕœ… «·Ï " & unittype & CHR(13)
    End If
    If MiniRentValue <> "" Then
        LogTextA = LogTextA & "  ⁄œÌ· «ﬁ· ﬁÌ„…  «ÃÌ—Ì…  «·Ï " & MiniRentValue & CHR(13)
    End If
    If Status <> "" Then
        LogTextA = LogTextA & "  ⁄œÌ·   Õ«·… «·ÊÕœ… «·Ï" & Status & CHR(13)
    End If
    AddToLogFile CInt(user_id), , Date, Time, LogTextA, LogTexte, "", Me.TxtModFlg.Text, " »Ì«‰«  «·⁄ﬁ«—", "", TxtAqarid.Text, txtaqarNo.Text
End Function
Function CuurentLogdata(Optional Currentmode As String)
    LogTextA = "    ‘«‘… " & ScreenNameArabic & CHR(13) & "«”„ «·⁄ﬁ«— " & CHR(13) & txtaqarname.Text & CHR(13) & " —ﬁ„ «·⁄ﬁ«—   " & txtaqarNo.Text & CHR(13) & "—ﬁ„ «·Õ—ﬂ…" & TxtAqarid.Text & CHR(13) & " «· «—ÌŒ " & Date & CHR(13) & "  «·„«·ﬂ   " & dcsupplier.Text
    LogTexte = "    Screen  " & ScreenNameEnglish & CHR(13) & "Real Estate Name " & CHR(13) & txtaqarname.Text & " Real Estate No. " & txtaqarNo.Text & CHR(13) & " Date " & Date & CHR(13) & " Owner" & dcsupplier.Text
    If Currentmode <> "D" Then
        AddToLogFile CInt(user_id), , Date, Time, LogTextA, LogTexte, Me.Name, Me.TxtModFlg, "", , val(TxtAqarid.Text), txtaqarNo.Text
    Else
        AddToLogFile CInt(user_id), , Date, Time, LogTextA, LogTexte, Me.Name, "D", "", , val(TxtAqarid.Text), txtaqarNo.Text
    End If
End Function
Private Sub UnitsGrid_CellButtonClick(ByVal Row As Long, ByVal Col As Long)
Exit Sub
    With Me.UnitsGrid
        Select Case .ColKey(Col)
            Case "EditRow"
                UpdateRow Row, val(.TextMatrix(Row, .ColIndex("id")))
        End Select
    End With
End Sub

Private Sub UnitsGrid_Click()
    On Error GoTo ErrTrap
    
    DcbTyped.ListIndex = val(UnitsGrid.TextMatrix(Me.UnitsGrid.Row, Me.UnitsGrid.ColIndex("Typed"))) - 1
    RecGID.Text = UnitsGrid.TextMatrix(Me.UnitsGrid.Row, Me.UnitsGrid.ColIndex("id"))
    DataBaseUnitNio = UnitsGrid.TextMatrix(Me.UnitsGrid.Row, Me.UnitsGrid.ColIndex("unitno"))
    UnitID.Text = UnitsGrid.TextMatrix(Me.UnitsGrid.Row, Me.UnitsGrid.ColIndex("unitno"))
    DCAkarUnit.BoundText = UnitsGrid.TextMatrix(Me.UnitsGrid.Row, Me.UnitsGrid.ColIndex("unittype"))
    TxtFloors.Text = UnitsGrid.TextMatrix(Me.UnitsGrid.Row, Me.UnitsGrid.ColIndex("Floor"))
    cBORENTTYPE.Text = (UnitsGrid.TextMatrix(Me.UnitsGrid.Row, Me.UnitsGrid.ColIndex("namerentType")))
    TxtLenght.Text = UnitsGrid.TextMatrix(Me.UnitsGrid.Row, Me.UnitsGrid.ColIndex("length"))
    txtMeterPrice.Text = UnitsGrid.TextMatrix(Me.UnitsGrid.Row, Me.UnitsGrid.ColIndex("meterPrice"))
    RentValue.Text = UnitsGrid.TextMatrix(Me.UnitsGrid.Row, Me.UnitsGrid.ColIndex("RentValue"))
    TxtMiniRentValue.Text = UnitsGrid.TextMatrix(Me.UnitsGrid.Row, Me.UnitsGrid.ColIndex("MiniRentValue"))
    TxtRooms.Text = UnitsGrid.TextMatrix(Me.UnitsGrid.Row, Me.UnitsGrid.ColIndex("roomscount"))
    TxtLoung.Text = UnitsGrid.TextMatrix(Me.UnitsGrid.Row, Me.UnitsGrid.ColIndex("LoungeCount"))
    TxtAccount.Text = UnitsGrid.TextMatrix(Me.UnitsGrid.Row, Me.UnitsGrid.ColIndex("ACCount"))
    TxtACCountÚSpleat.Text = UnitsGrid.TextMatrix(Me.UnitsGrid.Row, Me.UnitsGrid.ColIndex("ACCountspleat"))
    'UnitID.Text = UnitsGrid.TextMatrix(Me.UnitsGrid.Row, Me.UnitsGrid.ColIndex("haveFurniture"))
    TxtKitchn.Text = UnitsGrid.TextMatrix(Me.UnitsGrid.Row, Me.UnitsGrid.ColIndex("kithchencount"))
    BathNo.Text = UnitsGrid.TextMatrix(Me.UnitsGrid.Row, Me.UnitsGrid.ColIndex("WCcount"))
    UnitElc.Text = UnitsGrid.TextMatrix(Me.UnitsGrid.Row, Me.UnitsGrid.ColIndex("UnitElectric"))
    UnitStatus.BoundText = UnitsGrid.TextMatrix(Me.UnitsGrid.Row, Me.UnitsGrid.ColIndex("Statusid"))
    Disc.Text = UnitsGrid.TextMatrix(Me.UnitsGrid.Row, Me.UnitsGrid.ColIndex("unitdesc"))
    RenterDC.BoundText = UnitsGrid.TextMatrix(Me.UnitsGrid.Row, Me.UnitsGrid.ColIndex("customerid"))
    If UnitsGrid.ValueMatrix(Me.UnitsGrid.Row, Me.UnitsGrid.ColIndex("haveFurniture")) Then
        FerntChk = vbChecked
    Else
        FerntChk = vbUnchecked
    End If

    If UnitsGrid.ValueMatrix(Me.UnitsGrid.Row, Me.UnitsGrid.ColIndex("isTax")) Then
        chkIsTax = vbChecked
    Else
        chkIsTax = vbUnchecked
    End If
    
If Trim(Me.UnitStatus.Text) = "„ƒÃ—" Then
If SystemOptions.AllowChangeUnitIqar = False Then
UnitStatus.locked = True
Else
UnitStatus.locked = False
End If

Else
UnitStatus.locked = False
End If
ErrTrap:
End Sub

Private Sub UnitsGrid_StartEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
Exit Sub
    Dim rs As New ADODB.Recordset
    Dim StrSQL  As String
    Dim StrAccountType As String
    Dim StrComboList As String
    Dim Msg As String
    Dim StrAccountCode As String
    Dim StrAccountCode1 As String
    Dim ClsAcc As New ClsAccounts
    Dim LngRow As Long

    With UnitsGrid
        Select Case .ColKey(Col)
            Case "EditRow"
                .ColComboList(.ColIndex("EditRow")) = "..."
            Case "customeridname"
                StrSQL = "Select * from TblCustemers "
                rs.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText
                If SystemOptions.UserInterface = ArabicInterface Then
                    StrComboList = UnitsGrid.BuildComboList(rs, "CusName", "CusID")
                Else
                    StrComboList = UnitsGrid.BuildComboList(rs, "CusNamee", "CusID")
                End If
       
                If StrComboList <> "" Then
                    StrComboList = "|" & StrComboList
                End If
                 .ComboList = StrComboList
                  StrAccountCode = .ComboData
                LngRow = .FindRow(StrAccountCode, .FixedRows, .ColIndex("customerid"), False, True)
            Case "nameunittype"
                StrSQL = "select * from TblAkarUnit"
                rs.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

                If SystemOptions.UserInterface = ArabicInterface Then
                    StrComboList = UnitsGrid.BuildComboList(rs, "name", "id")
                Else
                    StrComboList = UnitsGrid.BuildComboList(rs, "namee", "id")
                End If
       
                If StrComboList <> "" Then
                    StrComboList = "|" & StrComboList
                End If
                 .ComboList = StrComboList
            Case "namerentType"
                StrSQL = "select * from TblRentType"
                rs.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

                If SystemOptions.UserInterface = ArabicInterface Then
                    StrComboList = UnitsGrid.BuildComboList(rs, "name", "id")
                Else
                    StrComboList = UnitsGrid.BuildComboList(rs, "namee", "id")
                End If
       
                If StrComboList <> "" Then
                    StrComboList = "|" & StrComboList
                End If
                 .ComboList = StrComboList
            Case "Status"
                StrSQL = "select * from TblRentStatus"
                rs.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

                If SystemOptions.UserInterface = ArabicInterface Then
                    StrComboList = UnitsGrid.BuildComboList(rs, "name", "id")
                Else
                    StrComboList = UnitsGrid.BuildComboList(rs, "namee", "id")
                End If
       
                If StrComboList <> "" Then
                    StrComboList = "|" & StrComboList
                End If
                 .ComboList = StrComboList
        End Select
    End With
    ReLineGrid
End Sub



Private Sub VSFlexGrid1_AfterEdit(ByVal Row As Long, ByVal Col As Long)
    If Row = VSFlexGrid1.Rows - 1 Then
        VSFlexGrid1.Rows = VSFlexGrid1.Rows + 1
    End If
    ReLineGrid
End Sub
Private Sub VSFlexGrid2_AfterEdit(ByVal Row As Long, ByVal Col As Long)
    If Row = VSFlexGrid2.Rows - 1 Then
            VSFlexGrid2.Rows = VSFlexGrid2.Rows + 1
    End If
    ReLineGrid
End Sub
'###########################################################################################################################################################################################
Private Sub UnitsGrid_EnterCell()

End Sub
Private Sub CmdDelete_Click()
    
    Dim Msg, StrSQL As String
    If (CheckUnitContract(val(RecGID.Text)) = True Or CheckUnitContractMerg(val(RecGID.Text)) = True) Then
        If SystemOptions.UserInterface = ArabicInterface Then
            MsgBox "·«Ì„ﬂ‰ Õ–› »Ì«‰«  Â–Â «·ÊÕœ… ·«— »«ÿÂ« »⁄ﬁœ "
        Else
            MsgBox "This Unit data cannot be deleted for being integrated with a contract record"
        End If
    Else
        If RecGID.Text <> "" Then
            If SystemOptions.UserInterface = ArabicInterface Then
                Msg = "”Ì „ Õ–› »Ì«‰«  «·”Ã· —ﬁ„ " & CHR(13)
                Msg = Msg + (RecGID.Text) & CHR(13)
                Msg = Msg + " Â·  —€» ›Ì Õ–› Â–Â «·»Ì«‰« ø"
            Else
                Msg = "Delete Recored File No. ?" & CHR(13)
                Msg = Msg + (RecGID.Text) & CHR(13)
                Msg = Msg + "  Are you sure you want to delete ?"
            End If
        
            If MsgBox(Msg, vbYesNo + vbQuestion + vbMsgBoxRight + vbMsgBoxRtlReading, App.title) = vbYes Then
                StrSQL = "delete From TblAqarDetai where  ID =" & val(RecGID.Text)
                Cn.Execute StrSQL, , adExecuteNoRecords
                If SystemOptions.UserInterface = ArabicInterface Then
                    MsgBox " „ Õ–› »‰Ã«Õ", vbInformation + vbOKOnly + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
                Else
                    MsgBox "Record deleted successfully", vbInformation + vbOKOnly + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
                End If
                FillGridWithData
            End If
        End If
    End If
End Sub

Public Function ReloadCombos()

    Dim Dcombos As ClsDataCombos
    Dim My_SQL As String
  
    My_SQL = "select UserID,UserName From tblUsers "
    fill_combo DCUser, My_SQL
 
    Set Dcombos = New ClsDataCombos
    Dcombos.GetAccountingCodes AccountVat
    Dcombos.GetAccountingCodes AccountVat2
    Dcombos.GetCustomersSuppliers 56, Me.dcCustomer
    Dcombos.GetCustomersSuppliers 57, Me.dcsupplier
    Dcombos.GetIqar DcbIqara
    Dcombos.getAkarUnit Me.DcbUnitType
    'Dcombos.GetIqarUnit 1, DcbUnitNo
    Dcombos.GetBranches dcBranch
    Dcombos.GetSalesRepData Me.DcboEmp
    Dcombos.GetEmployees Me.DcboEmpName
    Dcombos.GetUsers Me.DCboUserName
End Function
Sub ReloadUonit(Optional flg As Integer = 0)
Dim Dcombos As ClsDataCombos
Dim idd As Long
Dim idd1 As Long
Dim StrSQL As String
Set Dcombos = New ClsDataCombos
     StrSQL = " or id in(Select UntID from  TblIqrMerg where cont =" & val(TxtContNo.Text) & ")"
     StrSQL = StrSQL & " or id in (Select UnitNo from  TblContract    Where ContNo =" & val(TxtContNo.Text) & ")"
If val(DcbIqara.BoundText) > 0 Then
idd = val(DcbIqara.BoundText)
idd1 = val(DcbUnitType.BoundText)
If Me.TxtModFlg = "R" Or flg = 1 Then
Dcombos.GetIqarUnit idd, idd1, Me.DcbUnitNo, "R"
ElseIf Me.TxtModFlg = "N" Then
Dcombos.GetIqarUnit idd, idd1, Me.DcbUnitNo
ElseIf Me.TxtModFlg = "E" Then
Dcombos.GetIqarUnit idd, idd1, Me.DcbUnitNo, "E", StrSQL
End If
End If
End Sub

