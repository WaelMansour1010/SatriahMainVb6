VERSION 5.00
Object = "{C115893A-A3BF-43AF-B28D-69DB846077F3}#1.0#0"; "vsflex8u.ocx"
Object = "{0AFE7BE0-11B7-4A3E-978D-D4501E9A57FE}#1.0#0"; "c1sizer.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{FE5DCFAD-BC1D-11D2-94CF-004005455FAA}#1.4#0"; "ImpulseButton.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form FrmCitiesDistance 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   7830
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   11070
   Icon            =   "FrmCitiesDistance.frx":0000
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   RightToLeft     =   -1  'True
   ScaleHeight     =   7830
   ScaleWidth      =   11070
   ShowInTaskbar   =   0   'False
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
      Height          =   7770
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   11085
      _cx             =   19553
      _cy             =   13705
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
      Caption         =   "«·„”«›«  »Ì‰ «·„œ‰|»Ì«‰«  «·„Ê«‰∆|»Ì«‰«  |«‰Ê«⁄ «·‰ﬁ·|√‰Ê«⁄ «·—œÊœ"
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
      Begin C1SizerLibCtl.C1Elastic C1Elastic100 
         Height          =   7350
         Left            =   45
         TabIndex        =   1
         TabStop         =   0   'False
         Top             =   45
         Width           =   10995
         _cx             =   19394
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
            TabIndex        =   37
            Top             =   0
            Width           =   10995
            Begin VB.TextBox TxtVac_ID 
               Alignment       =   1  'Right Justify
               Appearance      =   0  'Flat
               BackColor       =   &H000000FF&
               Height          =   240
               Left            =   5670
               RightToLeft     =   -1  'True
               TabIndex        =   42
               Top             =   150
               Visible         =   0   'False
               Width           =   945
            End
            Begin VB.TextBox TxtModFlg 
               Alignment       =   1  'Right Justify
               Appearance      =   0  'Flat
               BackColor       =   &H0000FF00&
               Enabled         =   0   'False
               Height          =   285
               Left            =   4620
               RightToLeft     =   -1  'True
               TabIndex        =   41
               Text            =   "modflag"
               Top             =   120
               Visible         =   0   'False
               Width           =   465
            End
            Begin VB.Frame Frmo2 
               BackColor       =   &H00E0E0E0&
               BorderStyle     =   0  'None
               Enabled         =   0   'False
               Height          =   375
               Left            =   540
               RightToLeft     =   -1  'True
               TabIndex        =   38
               Top             =   690
               Visible         =   0   'False
               Width           =   3105
               Begin MSDataListLib.DataCombo DCUser 
                  CausesValidation=   0   'False
                  Height          =   315
                  Left            =   -255
                  TabIndex        =   39
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
                  TabIndex        =   40
                  Top             =   45
                  Width           =   855
               End
            End
            Begin MSComctlLib.ImageList GrdImageList 
               Left            =   7320
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
                     Picture         =   "FrmCitiesDistance.frx":57E2
                     Key             =   "CompanyName"
                  EndProperty
                  BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                     Picture         =   "FrmCitiesDistance.frx":5B7C
                     Key             =   "Ser"
                  EndProperty
                  BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                     Picture         =   "FrmCitiesDistance.frx":5F16
                     Key             =   "Vac_Name"
                  EndProperty
                  BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                     Picture         =   "FrmCitiesDistance.frx":62B0
                     Key             =   "ShareCount"
                  EndProperty
                  BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                     Picture         =   "FrmCitiesDistance.frx":664A
                     Key             =   "Dis_Count"
                  EndProperty
                  BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                     Picture         =   "FrmCitiesDistance.frx":69E4
                     Key             =   "Bouns"
                  EndProperty
                  BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                     Picture         =   "FrmCitiesDistance.frx":6D7E
                     Key             =   "SharesValue"
                  EndProperty
                  BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                     Picture         =   "FrmCitiesDistance.frx":7318
                     Key             =   "BuyValue"
                  EndProperty
               EndProperty
            End
            Begin ImpulseButton.ISButton btnLast 
               Height          =   315
               Left            =   570
               TabIndex        =   43
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
               ButtonImage     =   "FrmCitiesDistance.frx":76B2
               ColorButton     =   16777215
               AcclimateGrayTones=   -1  'True
               DrawFocusRectangle=   0   'False
               DisabledImageExtraction=   0
               DisabledImageStyle=   1
            End
            Begin ImpulseButton.ISButton btnNext 
               Height          =   315
               Left            =   1035
               TabIndex        =   44
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
               ButtonImage     =   "FrmCitiesDistance.frx":7A4C
               ColorButton     =   16777215
               DrawFocusRectangle=   0   'False
               DisabledImageStyle=   1
            End
            Begin ImpulseButton.ISButton btnPrevious 
               Height          =   315
               Left            =   1635
               TabIndex        =   45
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
               ButtonImage     =   "FrmCitiesDistance.frx":7DE6
               ColorButton     =   16777215
               DrawFocusRectangle=   0   'False
               DisabledImageStyle=   1
            End
            Begin ImpulseButton.ISButton btnFirst 
               Height          =   315
               Left            =   2100
               TabIndex        =   46
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
               ButtonImage     =   "FrmCitiesDistance.frx":8180
               ColorButton     =   16777215
               DrawFocusRectangle=   0   'False
               DisabledImageStyle=   1
            End
            Begin VB.Label Label1 
               Alignment       =   2  'Center
               BackStyle       =   0  'Transparent
               Caption         =   "«·„”«›«  »Ì‰ «·„œ‰"
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
               Left            =   8175
               RightToLeft     =   -1  'True
               TabIndex        =   47
               Top             =   90
               Width           =   2670
            End
         End
         Begin VB.Frame Frm2 
            BackColor       =   &H00E2E9E9&
            BorderStyle     =   0  'None
            Enabled         =   0   'False
            Height          =   1380
            Left            =   0
            RightToLeft     =   -1  'True
            TabIndex        =   16
            Top             =   3165
            Width           =   10920
            Begin VB.TextBox txtDriverPercentageUsed 
               Alignment       =   2  'Center
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
               Left            =   4230
               MaxLength       =   50
               RightToLeft     =   -1  'True
               TabIndex        =   203
               Top             =   1050
               Width           =   2010
            End
            Begin VB.TextBox txtTravelPriceUsed 
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
               Left            =   8160
               MaxLength       =   50
               RightToLeft     =   -1  'True
               TabIndex        =   202
               Tag             =   "⁄›Ê« Ì—ÃÏ «œŒ«·  ﬁÌ„… —œ «·—Õ·…"
               Top             =   1050
               Width           =   1410
            End
            Begin VB.TextBox txtDriverValueUsed 
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
               TabIndex        =   200
               Tag             =   "⁄›Ê« Ì—ÃÏ «œŒ«· ﬁÌ„… —œ «·”«∆ﬁ"
               Top             =   1050
               Width           =   2010
            End
            Begin VB.TextBox txtDistance 
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
               Left            =   8160
               MaxLength       =   50
               RightToLeft     =   -1  'True
               TabIndex        =   25
               Tag             =   "⁄›Ê« Ì—ÃÏ «œŒ«·  «·„”«›… ﬂ„"
               Top             =   390
               Width           =   1410
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
               Left            =   8160
               MaxLength       =   50
               RightToLeft     =   -1  'True
               TabIndex        =   24
               Top             =   30
               Width           =   1410
            End
            Begin VB.ComboBox CmbType 
               BackColor       =   &H80000018&
               Height          =   315
               ItemData        =   "FrmCitiesDistance.frx":851A
               Left            =   2280
               List            =   "FrmCitiesDistance.frx":852A
               RightToLeft     =   -1  'True
               Style           =   2  'Dropdown List
               TabIndex        =   23
               Top             =   1590
               Visible         =   0   'False
               Width           =   1005
            End
            Begin VB.TextBox txtKmPrice 
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
               Left            =   4230
               MaxLength       =   50
               RightToLeft     =   -1  'True
               TabIndex        =   21
               Tag             =   "⁄›Ê« Ì—ÃÏ «œŒ«·  ﬁÌ„… ﬂ„"
               Top             =   360
               Width           =   2010
            End
            Begin VB.TextBox txtTravelPrice 
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
               Left            =   8160
               MaxLength       =   50
               RightToLeft     =   -1  'True
               TabIndex        =   20
               Tag             =   "⁄›Ê« Ì—ÃÏ «œŒ«·  ﬁÌ„… —œ «·—Õ·…"
               Top             =   720
               Width           =   1410
            End
            Begin VB.TextBox txtDriverPercentage 
               Alignment       =   2  'Center
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
               Left            =   4230
               MaxLength       =   50
               RightToLeft     =   -1  'True
               TabIndex        =   19
               Top             =   720
               Width           =   2010
            End
            Begin VB.TextBox txtDriverValue 
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
               TabIndex        =   18
               Tag             =   "⁄›Ê« Ì—ÃÏ «œŒ«· ﬁÌ„… —œ «·”«∆ﬁ"
               Top             =   720
               Width           =   2010
            End
            Begin VB.TextBox txtDesil 
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
               TabIndex        =   17
               Tag             =   "⁄›Ê« Ì—ÃÏ «œŒ«·  ﬁÌ„… ﬂ„"
               Top             =   360
               Width           =   2010
            End
            Begin MSDataListLib.DataCombo DcboCountryID 
               Height          =   315
               Left            =   4230
               TabIndex        =   22
               Tag             =   "«Œ — «·„œÌ‰… „‰ „‰ ›÷·ﬂ"
               Top             =   30
               Width           =   3075
               _ExtentX        =   5424
               _ExtentY        =   556
               _Version        =   393216
               Text            =   ""
               RightToLeft     =   -1  'True
            End
            Begin MSDataListLib.DataCombo DcboCountryID1 
               Height          =   315
               Left            =   120
               TabIndex        =   26
               Tag             =   "«Œ — «·„œÌ‰…  «·Ï „‰ ›÷·ﬂ"
               Top             =   0
               Width           =   3075
               _ExtentX        =   5424
               _ExtentY        =   556
               _Version        =   393216
               Text            =   ""
               RightToLeft     =   -1  'True
            End
            Begin VB.Label Label1 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "‰”»… «·—œ „” ⁄„·…"
               Height          =   285
               Index           =   33
               Left            =   6480
               RightToLeft     =   -1  'True
               TabIndex        =   205
               Top             =   1050
               Width           =   1410
            End
            Begin VB.Label Label1 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "ﬁÌ„… «·—œ „” ⁄„·…"
               Height          =   285
               Index           =   32
               Left            =   9720
               RightToLeft     =   -1  'True
               TabIndex        =   204
               Top             =   1050
               Width           =   1200
            End
            Begin VB.Label Label1 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "ﬁÌ„… —œ «·”«∆ﬁ „” ⁄„·…"
               Height          =   345
               Index           =   31
               Left            =   2190
               RightToLeft     =   -1  'True
               TabIndex        =   201
               Top             =   1110
               Width           =   1830
            End
            Begin VB.Label Label1 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "„‰"
               Height          =   285
               Index           =   0
               Left            =   6720
               RightToLeft     =   -1  'True
               TabIndex        =   35
               Top             =   30
               Width           =   1050
            End
            Begin VB.Label Label1 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "«·ﬂÊœ"
               Height          =   195
               Index           =   3
               Left            =   9840
               RightToLeft     =   -1  'True
               TabIndex        =   34
               Top             =   30
               Width           =   1050
            End
            Begin VB.Label Label1 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "«·„”«›… ﬂ„"
               Height          =   285
               Index           =   1
               Left            =   9840
               RightToLeft     =   -1  'True
               TabIndex        =   33
               Top             =   390
               Width           =   1050
            End
            Begin VB.Label Label1 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "«·Ï"
               Height          =   285
               Index           =   4
               Left            =   3120
               RightToLeft     =   -1  'True
               TabIndex        =   32
               Top             =   30
               Width           =   690
            End
            Begin VB.Label Label1 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "«” Â·«ﬂ ﬂ„ œÌ“·"
               Height          =   285
               Index           =   5
               Left            =   6480
               RightToLeft     =   -1  'True
               TabIndex        =   31
               Top             =   390
               Width           =   1410
            End
            Begin VB.Label Label1 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "ﬁÌ„… —œ «·—Õ·…"
               Height          =   285
               Index           =   6
               Left            =   9840
               RightToLeft     =   -1  'True
               TabIndex        =   30
               Top             =   720
               Width           =   1050
            End
            Begin VB.Label Label1 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "‰”»… —œ «·”«∆ﬁ"
               Height          =   285
               Index           =   7
               Left            =   6480
               RightToLeft     =   -1  'True
               TabIndex        =   29
               Top             =   720
               Width           =   1410
            End
            Begin VB.Label Label1 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "ﬁÌ„… —œ «·”«∆ﬁ"
               Height          =   285
               Index           =   8
               Left            =   3000
               RightToLeft     =   -1  'True
               TabIndex        =   28
               Top             =   720
               Width           =   1050
            End
            Begin VB.Label Label1 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "ﬁÌ„… «·œÌ“·"
               Height          =   285
               Index           =   9
               Left            =   2880
               RightToLeft     =   -1  'True
               TabIndex        =   27
               Top             =   390
               Width           =   1050
            End
         End
         Begin C1SizerLibCtl.C1Elastic EltCont 
            Height          =   450
            Left            =   0
            TabIndex        =   2
            TabStop         =   0   'False
            Top             =   6780
            Width           =   10770
            _cx             =   18997
            _cy             =   794
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
               Left            =   9975
               TabIndex        =   3
               Top             =   75
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
               ButtonImage     =   "FrmCitiesDistance.frx":8543
               ColorButton     =   14871017
               DrawFocusRectangle=   0   'False
            End
            Begin ImpulseButton.ISButton btnSave 
               Height          =   330
               Left            =   7590
               TabIndex        =   4
               Top             =   75
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
               ButtonImage     =   "FrmCitiesDistance.frx":88DD
               ColorButton     =   14871017
               DrawFocusRectangle=   0   'False
            End
            Begin ImpulseButton.ISButton btnModify 
               Height          =   330
               Left            =   8835
               TabIndex        =   5
               Top             =   75
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
               ButtonImage     =   "FrmCitiesDistance.frx":8C77
               ColorButton     =   14871017
               DrawFocusRectangle=   0   'False
            End
            Begin ImpulseButton.ISButton BtnUndo 
               Height          =   330
               Left            =   6585
               TabIndex        =   6
               Top             =   75
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
               ButtonImage     =   "FrmCitiesDistance.frx":9011
               ColorButton     =   14871017
               DrawFocusRectangle=   0   'False
            End
            Begin ImpulseButton.ISButton btnDelete 
               Height          =   330
               Left            =   5460
               TabIndex        =   7
               Top             =   75
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
               ButtonImage     =   "FrmCitiesDistance.frx":93AB
               ColorButton     =   14871017
               DrawFocusRectangle=   0   'False
            End
            Begin ImpulseButton.ISButton btnQuery 
               Height          =   330
               Left            =   5880
               TabIndex        =   8
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
               ButtonImage     =   "FrmCitiesDistance.frx":9945
               ColorButton     =   14737632
               DrawFocusRectangle=   0   'False
            End
            Begin ImpulseButton.ISButton BtnUpdate 
               Height          =   330
               Left            =   6045
               TabIndex        =   9
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
               ButtonImage     =   "FrmCitiesDistance.frx":9CDF
               ColorButton     =   14871017
               DrawFocusRectangle=   0   'False
            End
            Begin ImpulseButton.ISButton BtnPrint 
               Height          =   285
               Left            =   4725
               TabIndex        =   10
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
               ButtonImage     =   "FrmCitiesDistance.frx":A079
               ColorButton     =   14871017
               DrawFocusRectangle=   0   'False
               DisabledImageStyle=   1
            End
            Begin ImpulseButton.ISButton btnCancel 
               Height          =   330
               Left            =   4200
               TabIndex        =   11
               Top             =   75
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
               ButtonImage     =   "FrmCitiesDistance.frx":A413
               ColorButton     =   14871017
               DrawFocusRectangle=   0   'False
               DisabledImageStyle=   1
            End
            Begin VB.Label LabCountRec 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Height          =   210
               Left            =   240
               RightToLeft     =   -1  'True
               TabIndex        =   15
               Top             =   75
               Width           =   540
            End
            Begin VB.Label LabCurrRec 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Height          =   210
               Left            =   1800
               RightToLeft     =   -1  'True
               TabIndex        =   14
               Top             =   75
               Width           =   675
            End
            Begin VB.Label Label2 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "⁄œœ «·”Ã·« :"
               Height          =   210
               Index           =   1
               Left            =   810
               RightToLeft     =   -1  'True
               TabIndex        =   13
               Top             =   75
               Width           =   975
            End
            Begin VB.Label Label2 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "«·”Ã· «·Õ«·Ì:"
               Height          =   210
               Index           =   0
               Left            =   2505
               RightToLeft     =   -1  'True
               TabIndex        =   12
               Top             =   75
               Width           =   975
            End
         End
         Begin VSFlex8UCtl.VSFlexGrid Grid 
            Height          =   2370
            Left            =   150
            TabIndex        =   36
            Top             =   720
            Width           =   10770
            _cx             =   18997
            _cy             =   4180
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
            Cols            =   14
            FixedRows       =   1
            FixedCols       =   1
            RowHeightMin    =   320
            RowHeightMax    =   0
            ColWidthMin     =   0
            ColWidthMax     =   0
            ExtendLastCol   =   -1  'True
            FormatString    =   $"FrmCitiesDistance.frx":A7AD
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
         Begin VSFlex8UCtl.VSFlexGrid VSFlexGrid1 
            Height          =   1800
            Left            =   150
            TabIndex        =   162
            Top             =   4635
            Width           =   10770
            _cx             =   18997
            _cy             =   3175
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
            Rows            =   2
            Cols            =   6
            FixedRows       =   1
            FixedCols       =   1
            RowHeightMin    =   320
            RowHeightMax    =   0
            ColWidthMin     =   0
            ColWidthMax     =   0
            ExtendLastCol   =   -1  'True
            FormatString    =   $"FrmCitiesDistance.frx":AA0F
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
            Height          =   240
            Left            =   9345
            TabIndex        =   163
            Top             =   6525
            Width           =   1575
            _ExtentX        =   2778
            _ExtentY        =   423
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
            ButtonImage     =   "FrmCitiesDistance.frx":AB07
            ColorButton     =   14871017
            DrawFocusRectangle=   0   'False
         End
      End
      Begin C1SizerLibCtl.C1Elastic C1Elastic1 
         Height          =   7350
         Left            =   11730
         TabIndex        =   48
         TabStop         =   0   'False
         Top             =   45
         Width           =   10995
         _cx             =   19394
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
            BackColor       =   &H00E2E9E9&
            BorderStyle     =   0  'None
            Enabled         =   0   'False
            Height          =   1740
            Left            =   0
            RightToLeft     =   -1  'True
            TabIndex        =   61
            Top             =   4440
            Width           =   10920
            Begin VB.TextBox NameE2 
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
               Left            =   2640
               MaxLength       =   50
               RightToLeft     =   -1  'True
               TabIndex        =   65
               Tag             =   "⁄›Ê« Ì—ÃÏ «œŒ«·  ﬁÌ„… —œ «·—Õ·…"
               Top             =   1065
               Width           =   5130
            End
            Begin VB.ComboBox Combo1 
               BackColor       =   &H80000018&
               Height          =   315
               ItemData        =   "FrmCitiesDistance.frx":B0A1
               Left            =   2280
               List            =   "FrmCitiesDistance.frx":B0B1
               RightToLeft     =   -1  'True
               Style           =   2  'Dropdown List
               TabIndex        =   64
               Top             =   1590
               Visible         =   0   'False
               Width           =   1005
            End
            Begin VB.TextBox ID2 
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
               Left            =   5040
               MaxLength       =   50
               RightToLeft     =   -1  'True
               TabIndex        =   63
               Top             =   180
               Width           =   2730
            End
            Begin VB.TextBox Name2 
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
               Left            =   2640
               MaxLength       =   50
               RightToLeft     =   -1  'True
               TabIndex        =   62
               Tag             =   "⁄›Ê« Ì—ÃÏ «œŒ«·  «·„”«›… ﬂ„"
               Top             =   615
               Width           =   5130
            End
            Begin VB.Label Label1 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "«·«”„ «‰Ã·Ì“Ì "
               Height          =   285
               Index           =   18
               Left            =   7920
               RightToLeft     =   -1  'True
               TabIndex        =   68
               Top             =   1080
               Width           =   1050
            End
            Begin VB.Label Label1 
               Alignment       =   2  'Center
               BackColor       =   &H00E2E9E9&
               Caption         =   "«·«”„ ⁄—»Ì"
               Height          =   285
               Index           =   15
               Left            =   7920
               RightToLeft     =   -1  'True
               TabIndex        =   67
               Top             =   630
               Width           =   1050
            End
            Begin VB.Label Label1 
               Alignment       =   2  'Center
               BackColor       =   &H00E2E9E9&
               Caption         =   "„”·”· "
               Height          =   195
               Index           =   14
               Left            =   7920
               RightToLeft     =   -1  'True
               TabIndex        =   66
               Top             =   240
               Width           =   1050
            End
         End
         Begin VB.Frame Frame1 
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   0  'None
            Height          =   720
            Left            =   0
            RightToLeft     =   -1  'True
            TabIndex        =   49
            Top             =   0
            Width           =   10995
            Begin VB.Frame Frame2 
               BackColor       =   &H00E0E0E0&
               BorderStyle     =   0  'None
               Enabled         =   0   'False
               Height          =   375
               Left            =   540
               RightToLeft     =   -1  'True
               TabIndex        =   52
               Top             =   690
               Visible         =   0   'False
               Width           =   3105
               Begin MSDataListLib.DataCombo DataCombo1 
                  CausesValidation=   0   'False
                  Height          =   315
                  Left            =   -255
                  TabIndex        =   53
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
                  Index           =   10
                  Left            =   2160
                  RightToLeft     =   -1  'True
                  TabIndex        =   54
                  Top             =   45
                  Width           =   855
               End
            End
            Begin VB.TextBox TxtModFlg2 
               Alignment       =   1  'Right Justify
               Appearance      =   0  'Flat
               BackColor       =   &H0000FF00&
               Enabled         =   0   'False
               Height          =   285
               Left            =   4620
               RightToLeft     =   -1  'True
               TabIndex        =   51
               Text            =   "modflag"
               Top             =   90
               Visible         =   0   'False
               Width           =   465
            End
            Begin VB.TextBox TxtVac_ID2 
               Alignment       =   1  'Right Justify
               Appearance      =   0  'Flat
               BackColor       =   &H000000FF&
               Height          =   240
               Left            =   5670
               RightToLeft     =   -1  'True
               TabIndex        =   50
               Top             =   150
               Visible         =   0   'False
               Width           =   945
            End
            Begin MSComctlLib.ImageList GrdImageList2 
               Left            =   7320
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
                     Picture         =   "FrmCitiesDistance.frx":B0CA
                     Key             =   "CompanyName"
                  EndProperty
                  BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                     Picture         =   "FrmCitiesDistance.frx":B464
                     Key             =   "Ser"
                  EndProperty
                  BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                     Picture         =   "FrmCitiesDistance.frx":B7FE
                     Key             =   "Vac_Name"
                  EndProperty
                  BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                     Picture         =   "FrmCitiesDistance.frx":BB98
                     Key             =   "ShareCount"
                  EndProperty
                  BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                     Picture         =   "FrmCitiesDistance.frx":BF32
                     Key             =   "Dis_Count"
                  EndProperty
                  BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                     Picture         =   "FrmCitiesDistance.frx":C2CC
                     Key             =   "Bouns"
                  EndProperty
                  BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                     Picture         =   "FrmCitiesDistance.frx":C666
                     Key             =   "SharesValue"
                  EndProperty
                  BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                     Picture         =   "FrmCitiesDistance.frx":CC00
                     Key             =   "BuyValue"
                  EndProperty
               EndProperty
            End
            Begin ImpulseButton.ISButton btnLast2 
               Height          =   315
               Left            =   570
               TabIndex        =   55
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
               ButtonImage     =   "FrmCitiesDistance.frx":CF9A
               ColorButton     =   16777215
               AcclimateGrayTones=   -1  'True
               DrawFocusRectangle=   0   'False
               DisabledImageExtraction=   0
               DisabledImageStyle=   1
            End
            Begin ImpulseButton.ISButton btnNext2 
               Height          =   315
               Left            =   1035
               TabIndex        =   56
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
               ButtonImage     =   "FrmCitiesDistance.frx":D334
               ColorButton     =   16777215
               DrawFocusRectangle=   0   'False
               DisabledImageStyle=   1
            End
            Begin ImpulseButton.ISButton btnPrevious2 
               Height          =   315
               Left            =   1635
               TabIndex        =   57
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
               ButtonImage     =   "FrmCitiesDistance.frx":D6CE
               ColorButton     =   16777215
               DrawFocusRectangle=   0   'False
               DisabledImageStyle=   1
            End
            Begin ImpulseButton.ISButton btnFirst2 
               Height          =   315
               Left            =   2100
               TabIndex        =   58
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
               ButtonImage     =   "FrmCitiesDistance.frx":DA68
               ColorButton     =   16777215
               DrawFocusRectangle=   0   'False
               DisabledImageStyle=   1
            End
            Begin VB.Label Label1 
               Alignment       =   2  'Center
               BackStyle       =   0  'Transparent
               Caption         =   "»Ì«‰«  «·„Ê«‰∆"
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
               Index           =   11
               Left            =   8175
               RightToLeft     =   -1  'True
               TabIndex        =   59
               Top             =   90
               Width           =   2670
            End
         End
         Begin VSFlex8UCtl.VSFlexGrid Grid2 
            Height          =   3570
            Left            =   150
            TabIndex        =   60
            Top             =   780
            Width           =   10770
            _cx             =   18997
            _cy             =   6297
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
            FormatString    =   $"FrmCitiesDistance.frx":DE02
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
         Begin C1SizerLibCtl.C1Elastic C1Elastic2 
            Height          =   1155
            Left            =   2565
            TabIndex        =   69
            TabStop         =   0   'False
            Top             =   6135
            Width           =   5865
            _cx             =   10345
            _cy             =   2037
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
            Begin ImpulseButton.ISButton btnNew2 
               Height          =   330
               Left            =   4575
               TabIndex        =   70
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
               ButtonImage     =   "FrmCitiesDistance.frx":DE9A
               ColorButton     =   14871017
               DrawFocusRectangle=   0   'False
            End
            Begin ImpulseButton.ISButton btnSave2 
               Height          =   330
               Left            =   3030
               TabIndex        =   71
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
               ButtonImage     =   "FrmCitiesDistance.frx":E234
               ColorButton     =   14871017
               DrawFocusRectangle=   0   'False
            End
            Begin ImpulseButton.ISButton btnModify2 
               Height          =   330
               Left            =   3795
               TabIndex        =   72
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
               ButtonImage     =   "FrmCitiesDistance.frx":E5CE
               ColorButton     =   14871017
               DrawFocusRectangle=   0   'False
            End
            Begin ImpulseButton.ISButton BtnUndo2 
               Height          =   330
               Left            =   2265
               TabIndex        =   73
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
               ButtonImage     =   "FrmCitiesDistance.frx":E968
               ColorButton     =   14871017
               DrawFocusRectangle=   0   'False
            End
            Begin ImpulseButton.ISButton btnDelete2 
               Height          =   330
               Left            =   1500
               TabIndex        =   74
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
               ButtonImage     =   "FrmCitiesDistance.frx":ED02
               ColorButton     =   14871017
               DrawFocusRectangle=   0   'False
            End
            Begin ImpulseButton.ISButton ISButton10 
               Height          =   330
               Left            =   5880
               TabIndex        =   75
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
               ButtonImage     =   "FrmCitiesDistance.frx":F29C
               ColorButton     =   14737632
               DrawFocusRectangle=   0   'False
            End
            Begin ImpulseButton.ISButton ISButton11 
               Height          =   330
               Left            =   6045
               TabIndex        =   76
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
               ButtonImage     =   "FrmCitiesDistance.frx":F636
               ColorButton     =   14871017
               DrawFocusRectangle=   0   'False
            End
            Begin ImpulseButton.ISButton ISButton12 
               Height          =   285
               Left            =   4725
               TabIndex        =   77
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
               ButtonImage     =   "FrmCitiesDistance.frx":F9D0
               ColorButton     =   14871017
               DrawFocusRectangle=   0   'False
               DisabledImageStyle=   1
            End
            Begin ImpulseButton.ISButton btnCancel2 
               Height          =   330
               Left            =   705
               TabIndex        =   78
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
               ButtonImage     =   "FrmCitiesDistance.frx":FD6A
               ColorButton     =   14871017
               DrawFocusRectangle=   0   'False
               DisabledImageStyle=   1
            End
            Begin VB.Label Label2 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "«·”Ã· «·Õ«·Ì:"
               Height          =   210
               Index           =   3
               Left            =   2505
               RightToLeft     =   -1  'True
               TabIndex        =   82
               Top             =   225
               Width           =   975
            End
            Begin VB.Label Label2 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "⁄œœ «·”Ã·« :"
               Height          =   210
               Index           =   2
               Left            =   810
               RightToLeft     =   -1  'True
               TabIndex        =   81
               Top             =   225
               Width           =   975
            End
            Begin VB.Label LabCurrRec2 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Height          =   210
               Left            =   1800
               RightToLeft     =   -1  'True
               TabIndex        =   80
               Top             =   240
               Width           =   675
            End
            Begin VB.Label LabCountRec2 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Height          =   210
               Left            =   240
               RightToLeft     =   -1  'True
               TabIndex        =   79
               Top             =   225
               Width           =   540
            End
         End
      End
      Begin C1SizerLibCtl.C1Elastic C1Elastic3 
         Height          =   7350
         Left            =   12030
         TabIndex        =   83
         TabStop         =   0   'False
         Top             =   45
         Width           =   10995
         _cx             =   19394
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
         Begin VB.Frame Frame5 
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   0  'None
            Height          =   720
            Left            =   0
            RightToLeft     =   -1  'True
            TabIndex        =   92
            Top             =   0
            Width           =   10995
            Begin VB.TextBox TxtVac_ID3 
               Alignment       =   1  'Right Justify
               Appearance      =   0  'Flat
               BackColor       =   &H000000FF&
               Height          =   240
               Left            =   5670
               RightToLeft     =   -1  'True
               TabIndex        =   97
               Top             =   150
               Visible         =   0   'False
               Width           =   945
            End
            Begin VB.TextBox TxtModFlg3 
               Alignment       =   1  'Right Justify
               Appearance      =   0  'Flat
               BackColor       =   &H0000FF00&
               Enabled         =   0   'False
               Height          =   285
               Left            =   4560
               RightToLeft     =   -1  'True
               TabIndex        =   96
               Text            =   "modflag"
               Top             =   90
               Visible         =   0   'False
               Width           =   465
            End
            Begin VB.Frame Frame6 
               BackColor       =   &H00E0E0E0&
               BorderStyle     =   0  'None
               Enabled         =   0   'False
               Height          =   375
               Left            =   540
               RightToLeft     =   -1  'True
               TabIndex        =   93
               Top             =   690
               Visible         =   0   'False
               Width           =   3105
               Begin MSDataListLib.DataCombo DataCombo2 
                  CausesValidation=   0   'False
                  Height          =   315
                  Left            =   -255
                  TabIndex        =   94
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
                  Index           =   19
                  Left            =   2160
                  RightToLeft     =   -1  'True
                  TabIndex        =   95
                  Top             =   45
                  Width           =   855
               End
            End
            Begin MSComctlLib.ImageList GrdImageList3 
               Left            =   7320
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
                     Picture         =   "FrmCitiesDistance.frx":10104
                     Key             =   "CompanyName"
                  EndProperty
                  BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                     Picture         =   "FrmCitiesDistance.frx":1049E
                     Key             =   "Ser"
                  EndProperty
                  BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                     Picture         =   "FrmCitiesDistance.frx":10838
                     Key             =   "Vac_Name"
                  EndProperty
                  BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                     Picture         =   "FrmCitiesDistance.frx":10BD2
                     Key             =   "ShareCount"
                  EndProperty
                  BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                     Picture         =   "FrmCitiesDistance.frx":10F6C
                     Key             =   "Dis_Count"
                  EndProperty
                  BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                     Picture         =   "FrmCitiesDistance.frx":11306
                     Key             =   "Bouns"
                  EndProperty
                  BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                     Picture         =   "FrmCitiesDistance.frx":116A0
                     Key             =   "SharesValue"
                  EndProperty
                  BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                     Picture         =   "FrmCitiesDistance.frx":11C3A
                     Key             =   "BuyValue"
                  EndProperty
               EndProperty
            End
            Begin ImpulseButton.ISButton btnLast3 
               Height          =   315
               Left            =   570
               TabIndex        =   98
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
               ButtonImage     =   "FrmCitiesDistance.frx":11FD4
               ColorButton     =   16777215
               AcclimateGrayTones=   -1  'True
               DrawFocusRectangle=   0   'False
               DisabledImageExtraction=   0
               DisabledImageStyle=   1
            End
            Begin ImpulseButton.ISButton btnNext3 
               Height          =   315
               Left            =   1035
               TabIndex        =   99
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
               ButtonImage     =   "FrmCitiesDistance.frx":1236E
               ColorButton     =   16777215
               DrawFocusRectangle=   0   'False
               DisabledImageStyle=   1
            End
            Begin ImpulseButton.ISButton btnPrevious3 
               Height          =   315
               Left            =   1635
               TabIndex        =   100
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
               ButtonImage     =   "FrmCitiesDistance.frx":12708
               ColorButton     =   16777215
               DrawFocusRectangle=   0   'False
               DisabledImageStyle=   1
            End
            Begin ImpulseButton.ISButton btnFirst3 
               Height          =   315
               Left            =   2100
               TabIndex        =   101
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
               ButtonImage     =   "FrmCitiesDistance.frx":12AA2
               ColorButton     =   16777215
               DrawFocusRectangle=   0   'False
               DisabledImageStyle=   1
            End
            Begin VB.Label Label1 
               Alignment       =   2  'Center
               BackStyle       =   0  'Transparent
               Caption         =   "»Ì«‰«  «·”›‰"
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
               Index           =   20
               Left            =   8175
               RightToLeft     =   -1  'True
               TabIndex        =   102
               Top             =   90
               Width           =   2670
            End
         End
         Begin VB.Frame Frame33 
            BackColor       =   &H00E2E9E9&
            BorderStyle     =   0  'None
            Enabled         =   0   'False
            Height          =   1860
            Left            =   -150
            RightToLeft     =   -1  'True
            TabIndex        =   84
            Top             =   4440
            Width           =   10920
            Begin VB.TextBox Name3 
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
               Left            =   2640
               MaxLength       =   50
               RightToLeft     =   -1  'True
               TabIndex        =   88
               Tag             =   "⁄›Ê« Ì—ÃÏ «œŒ«·  «·„”«›… ﬂ„"
               Top             =   495
               Width           =   5130
            End
            Begin VB.TextBox ID3 
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
               Left            =   4920
               MaxLength       =   50
               RightToLeft     =   -1  'True
               TabIndex        =   87
               Top             =   60
               Width           =   2850
            End
            Begin VB.ComboBox Combo2 
               BackColor       =   &H80000018&
               Height          =   315
               ItemData        =   "FrmCitiesDistance.frx":12E3C
               Left            =   2280
               List            =   "FrmCitiesDistance.frx":12E4C
               RightToLeft     =   -1  'True
               Style           =   2  'Dropdown List
               TabIndex        =   86
               Top             =   1710
               Visible         =   0   'False
               Width           =   1005
            End
            Begin VB.TextBox NameE3 
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
               Left            =   2640
               MaxLength       =   50
               RightToLeft     =   -1  'True
               TabIndex        =   85
               Tag             =   "⁄›Ê« Ì—ÃÏ «œŒ«·  ﬁÌ„… —œ «·—Õ·…"
               Top             =   945
               Width           =   5130
            End
            Begin MSDataListLib.DataCombo DcbHarbor 
               Height          =   315
               Left            =   2640
               TabIndex        =   118
               Top             =   1320
               Width           =   5130
               _ExtentX        =   9049
               _ExtentY        =   556
               _Version        =   393216
               Text            =   ""
               RightToLeft     =   -1  'True
            End
            Begin VB.Label lbl 
               Alignment       =   2  'Center
               BackColor       =   &H00E2E9E9&
               BackStyle       =   0  'Transparent
               Caption         =   "«·„Ì‰«¡"
               Height          =   195
               Index           =   55
               Left            =   7920
               RightToLeft     =   -1  'True
               TabIndex        =   119
               Top             =   1320
               Width           =   1050
            End
            Begin VB.Label Label1 
               Alignment       =   2  'Center
               BackColor       =   &H00E2E9E9&
               Caption         =   "„”·”· "
               Height          =   195
               Index           =   17
               Left            =   7920
               RightToLeft     =   -1  'True
               TabIndex        =   91
               Top             =   120
               Width           =   1050
            End
            Begin VB.Label Label1 
               Alignment       =   2  'Center
               BackColor       =   &H00E2E9E9&
               Caption         =   "«·«”„ ⁄—»Ì"
               Height          =   195
               Index           =   16
               Left            =   7920
               RightToLeft     =   -1  'True
               TabIndex        =   90
               Top             =   510
               Width           =   1050
            End
            Begin VB.Label Label1 
               Alignment       =   2  'Center
               BackColor       =   &H00E2E9E9&
               Caption         =   "«·«”„ «‰Ã·Ì“Ì "
               Height          =   195
               Index           =   12
               Left            =   7920
               RightToLeft     =   -1  'True
               TabIndex        =   89
               Top             =   960
               Width           =   1050
            End
         End
         Begin VSFlex8UCtl.VSFlexGrid Grid3 
            Height          =   3570
            Left            =   150
            TabIndex        =   103
            Top             =   780
            Width           =   10770
            _cx             =   18997
            _cy             =   6297
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
            FormatString    =   $"FrmCitiesDistance.frx":12E65
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
         Begin C1SizerLibCtl.C1Elastic C1Elastic4 
            Height          =   1020
            Left            =   2565
            TabIndex        =   104
            TabStop         =   0   'False
            Top             =   6270
            Width           =   5865
            _cx             =   10345
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
            Begin ImpulseButton.ISButton btnNew3 
               Height          =   330
               Left            =   4575
               TabIndex        =   105
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
               ButtonImage     =   "FrmCitiesDistance.frx":12EFA
               ColorButton     =   14871017
               DrawFocusRectangle=   0   'False
            End
            Begin ImpulseButton.ISButton btnSave3 
               Height          =   330
               Left            =   3030
               TabIndex        =   106
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
               ButtonImage     =   "FrmCitiesDistance.frx":13294
               ColorButton     =   14871017
               DrawFocusRectangle=   0   'False
            End
            Begin ImpulseButton.ISButton btnModify3 
               Height          =   330
               Left            =   3795
               TabIndex        =   107
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
               ButtonImage     =   "FrmCitiesDistance.frx":1362E
               ColorButton     =   14871017
               DrawFocusRectangle=   0   'False
            End
            Begin ImpulseButton.ISButton BtnUndo3 
               Height          =   330
               Left            =   2265
               TabIndex        =   108
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
               ButtonImage     =   "FrmCitiesDistance.frx":139C8
               ColorButton     =   14871017
               DrawFocusRectangle=   0   'False
            End
            Begin ImpulseButton.ISButton btnDelete3 
               Height          =   330
               Left            =   1440
               TabIndex        =   109
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
               ButtonImage     =   "FrmCitiesDistance.frx":13D62
               ColorButton     =   14871017
               DrawFocusRectangle=   0   'False
            End
            Begin ImpulseButton.ISButton ISButton18 
               Height          =   330
               Left            =   5880
               TabIndex        =   110
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
               ButtonImage     =   "FrmCitiesDistance.frx":142FC
               ColorButton     =   14737632
               DrawFocusRectangle=   0   'False
            End
            Begin ImpulseButton.ISButton ISButton19 
               Height          =   330
               Left            =   6045
               TabIndex        =   111
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
               ButtonImage     =   "FrmCitiesDistance.frx":14696
               ColorButton     =   14871017
               DrawFocusRectangle=   0   'False
            End
            Begin ImpulseButton.ISButton ISButton20 
               Height          =   285
               Left            =   4725
               TabIndex        =   112
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
               ButtonImage     =   "FrmCitiesDistance.frx":14A30
               ColorButton     =   14871017
               DrawFocusRectangle=   0   'False
               DisabledImageStyle=   1
            End
            Begin ImpulseButton.ISButton btnCancel3 
               Height          =   330
               Left            =   705
               TabIndex        =   113
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
               ButtonImage     =   "FrmCitiesDistance.frx":14DCA
               ColorButton     =   14871017
               DrawFocusRectangle=   0   'False
               DisabledImageStyle=   1
            End
            Begin VB.Label LabCountRec3 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Height          =   210
               Left            =   240
               RightToLeft     =   -1  'True
               TabIndex        =   117
               Top             =   225
               Width           =   540
            End
            Begin VB.Label LabCurrRec3 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Height          =   210
               Left            =   1800
               RightToLeft     =   -1  'True
               TabIndex        =   116
               Top             =   240
               Width           =   675
            End
            Begin VB.Label Label2 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "⁄œœ «·”Ã·« :"
               Height          =   210
               Index           =   5
               Left            =   810
               RightToLeft     =   -1  'True
               TabIndex        =   115
               Top             =   225
               Width           =   975
            End
            Begin VB.Label Label2 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "«·”Ã· «·Õ«·Ì:"
               Height          =   210
               Index           =   4
               Left            =   2505
               RightToLeft     =   -1  'True
               TabIndex        =   114
               Top             =   225
               Width           =   975
            End
         End
      End
      Begin C1SizerLibCtl.C1Elastic C1Elastic5 
         Height          =   7350
         Left            =   12330
         TabIndex        =   120
         TabStop         =   0   'False
         Top             =   45
         Width           =   10995
         _cx             =   19394
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
         Begin VB.Frame Frame8 
            BackColor       =   &H00E2E9E9&
            BorderStyle     =   0  'None
            Enabled         =   0   'False
            Height          =   1740
            Left            =   0
            RightToLeft     =   -1  'True
            TabIndex        =   132
            Top             =   4440
            Width           =   10920
            Begin VB.TextBox TxtModFlg4 
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
               Left            =   1320
               MaxLength       =   50
               RightToLeft     =   -1  'True
               TabIndex        =   161
               Tag             =   "⁄›Ê« Ì—ÃÏ «œŒ«·  «·„”«›… ﬂ„"
               Top             =   0
               Visible         =   0   'False
               Width           =   810
            End
            Begin VB.TextBox TxtAccount2 
               Alignment       =   1  'Right Justify
               Height          =   315
               Left            =   3465
               RightToLeft     =   -1  'True
               TabIndex        =   159
               Top             =   840
               Width           =   705
            End
            Begin VB.TextBox TxtAccount 
               Alignment       =   1  'Right Justify
               Height          =   315
               Left            =   8985
               RightToLeft     =   -1  'True
               TabIndex        =   156
               Top             =   840
               Width           =   705
            End
            Begin VB.TextBox TxtName3E 
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
               TabIndex        =   136
               Tag             =   "⁄›Ê« Ì—ÃÏ «œŒ«·  ﬁÌ„… —œ «·—Õ·…"
               Top             =   480
               Width           =   4050
            End
            Begin VB.ComboBox Combo3 
               BackColor       =   &H80000018&
               Height          =   315
               ItemData        =   "FrmCitiesDistance.frx":15164
               Left            =   2280
               List            =   "FrmCitiesDistance.frx":15174
               RightToLeft     =   -1  'True
               Style           =   2  'Dropdown List
               TabIndex        =   135
               Top             =   1590
               Visible         =   0   'False
               Width           =   1005
            End
            Begin VB.TextBox TxtTransID 
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
               Left            =   5520
               MaxLength       =   50
               RightToLeft     =   -1  'True
               TabIndex        =   134
               Top             =   60
               Width           =   4170
            End
            Begin VB.TextBox TxtName3 
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
               Left            =   5520
               MaxLength       =   50
               RightToLeft     =   -1  'True
               TabIndex        =   133
               Tag             =   "⁄›Ê« Ì—ÃÏ «œŒ«·  «·„”«›… ﬂ„"
               Top             =   495
               Width           =   4170
            End
            Begin MSDataListLib.DataCombo DcbAccount 
               Height          =   315
               Left            =   5520
               TabIndex        =   157
               Top             =   840
               Width           =   3375
               _ExtentX        =   5953
               _ExtentY        =   556
               _Version        =   393216
               Text            =   ""
               RightToLeft     =   -1  'True
            End
            Begin MSDataListLib.DataCombo DcbAccount2 
               Height          =   315
               Left            =   120
               TabIndex        =   160
               Top             =   840
               Width           =   3255
               _ExtentX        =   5741
               _ExtentY        =   556
               _Version        =   393216
               Text            =   ""
               RightToLeft     =   -1  'True
            End
            Begin VB.Label lbl 
               Alignment       =   2  'Center
               BackColor       =   &H00E2E9E9&
               BackStyle       =   0  'Transparent
               Caption         =   "Õ‹ «·„’—Ê›« "
               Height          =   405
               Index           =   1
               Left            =   4320
               RightToLeft     =   -1  'True
               TabIndex        =   158
               Top             =   840
               Width           =   1170
            End
            Begin VB.Label Label1 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "«·«”„ «‰Ã·Ì“Ì "
               Height          =   285
               Index           =   25
               Left            =   4200
               RightToLeft     =   -1  'True
               TabIndex        =   140
               Top             =   480
               Width           =   1170
            End
            Begin VB.Label Label1 
               Alignment       =   2  'Center
               BackColor       =   &H00E2E9E9&
               Caption         =   "«·«”„ ⁄—»Ì"
               Height          =   285
               Index           =   24
               Left            =   9840
               RightToLeft     =   -1  'True
               TabIndex        =   139
               Top             =   510
               Width           =   1050
            End
            Begin VB.Label Label1 
               Alignment       =   2  'Center
               BackColor       =   &H00E2E9E9&
               Caption         =   "„”·”· "
               Height          =   195
               Index           =   23
               Left            =   9720
               RightToLeft     =   -1  'True
               TabIndex        =   138
               Top             =   120
               Width           =   1050
            End
            Begin VB.Label lbl 
               Alignment       =   2  'Center
               BackColor       =   &H00E2E9E9&
               BackStyle       =   0  'Transparent
               Caption         =   "Õ‹ «·«Ì—«œ« "
               Height          =   405
               Index           =   0
               Left            =   9840
               RightToLeft     =   -1  'True
               TabIndex        =   137
               Top             =   840
               Width           =   1050
            End
         End
         Begin VB.Frame Frame4 
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   0  'None
            Height          =   720
            Left            =   0
            RightToLeft     =   -1  'True
            TabIndex        =   121
            Top             =   0
            Width           =   10995
            Begin VB.Frame Frame7 
               BackColor       =   &H00E0E0E0&
               BorderStyle     =   0  'None
               Enabled         =   0   'False
               Height          =   375
               Left            =   540
               RightToLeft     =   -1  'True
               TabIndex        =   124
               Top             =   690
               Visible         =   0   'False
               Width           =   3105
               Begin MSDataListLib.DataCombo DataCombo3 
                  CausesValidation=   0   'False
                  Height          =   315
                  Left            =   -255
                  TabIndex        =   125
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
                  Index           =   21
                  Left            =   2160
                  RightToLeft     =   -1  'True
                  TabIndex        =   126
                  Top             =   45
                  Width           =   855
               End
            End
            Begin VB.TextBox Text2 
               Alignment       =   1  'Right Justify
               Appearance      =   0  'Flat
               BackColor       =   &H0000FF00&
               Enabled         =   0   'False
               Height          =   285
               Left            =   4560
               RightToLeft     =   -1  'True
               TabIndex        =   123
               Text            =   "modflag"
               Top             =   90
               Visible         =   0   'False
               Width           =   465
            End
            Begin VB.TextBox Text1 
               Alignment       =   1  'Right Justify
               Appearance      =   0  'Flat
               BackColor       =   &H000000FF&
               Height          =   240
               Left            =   5670
               RightToLeft     =   -1  'True
               TabIndex        =   122
               Top             =   150
               Visible         =   0   'False
               Width           =   945
            End
            Begin MSComctlLib.ImageList ImageList1 
               Left            =   7320
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
                     Picture         =   "FrmCitiesDistance.frx":1518D
                     Key             =   "CompanyName"
                  EndProperty
                  BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                     Picture         =   "FrmCitiesDistance.frx":15527
                     Key             =   "Ser"
                  EndProperty
                  BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                     Picture         =   "FrmCitiesDistance.frx":158C1
                     Key             =   "Vac_Name"
                  EndProperty
                  BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                     Picture         =   "FrmCitiesDistance.frx":15C5B
                     Key             =   "ShareCount"
                  EndProperty
                  BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                     Picture         =   "FrmCitiesDistance.frx":15FF5
                     Key             =   "Dis_Count"
                  EndProperty
                  BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                     Picture         =   "FrmCitiesDistance.frx":1638F
                     Key             =   "Bouns"
                  EndProperty
                  BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                     Picture         =   "FrmCitiesDistance.frx":16729
                     Key             =   "SharesValue"
                  EndProperty
                  BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                     Picture         =   "FrmCitiesDistance.frx":16CC3
                     Key             =   "BuyValue"
                  EndProperty
               EndProperty
            End
            Begin ImpulseButton.ISButton btnLast4 
               Height          =   315
               Left            =   570
               TabIndex        =   127
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
               ButtonImage     =   "FrmCitiesDistance.frx":1705D
               ColorButton     =   16777215
               AcclimateGrayTones=   -1  'True
               DrawFocusRectangle=   0   'False
               DisabledImageExtraction=   0
               DisabledImageStyle=   1
            End
            Begin ImpulseButton.ISButton btnNext4 
               Height          =   315
               Left            =   1035
               TabIndex        =   128
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
               ButtonImage     =   "FrmCitiesDistance.frx":173F7
               ColorButton     =   16777215
               DrawFocusRectangle=   0   'False
               DisabledImageStyle=   1
            End
            Begin ImpulseButton.ISButton btnPrevious4 
               Height          =   315
               Left            =   1635
               TabIndex        =   129
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
               ButtonImage     =   "FrmCitiesDistance.frx":17791
               ColorButton     =   16777215
               DrawFocusRectangle=   0   'False
               DisabledImageStyle=   1
            End
            Begin ImpulseButton.ISButton btnFirst4 
               Height          =   315
               Left            =   2100
               TabIndex        =   130
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
               ButtonImage     =   "FrmCitiesDistance.frx":17B2B
               ColorButton     =   16777215
               DrawFocusRectangle=   0   'False
               DisabledImageStyle=   1
            End
            Begin VB.Label Label1 
               Alignment       =   2  'Center
               BackStyle       =   0  'Transparent
               Caption         =   "»Ì«‰«  «‰Ê«⁄ «·‰ﬁ·"
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
               Index           =   22
               Left            =   8175
               RightToLeft     =   -1  'True
               TabIndex        =   131
               Top             =   90
               Width           =   2670
            End
         End
         Begin VSFlex8UCtl.VSFlexGrid Grid4 
            Height          =   3570
            Left            =   150
            TabIndex        =   141
            Top             =   780
            Width           =   10770
            _cx             =   18997
            _cy             =   6297
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
            Cols            =   8
            FixedRows       =   1
            FixedCols       =   1
            RowHeightMin    =   320
            RowHeightMax    =   0
            ColWidthMin     =   0
            ColWidthMax     =   0
            ExtendLastCol   =   -1  'True
            FormatString    =   $"FrmCitiesDistance.frx":17EC5
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
         Begin C1SizerLibCtl.C1Elastic C1Elastic6 
            Height          =   1155
            Left            =   2565
            TabIndex        =   142
            TabStop         =   0   'False
            Top             =   6135
            Width           =   5865
            _cx             =   10345
            _cy             =   2037
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
               TabIndex        =   143
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
               ButtonImage     =   "FrmCitiesDistance.frx":18018
               ColorButton     =   14871017
               DrawFocusRectangle=   0   'False
            End
            Begin ImpulseButton.ISButton btnSave4 
               Height          =   330
               Left            =   3030
               TabIndex        =   144
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
               ButtonImage     =   "FrmCitiesDistance.frx":183B2
               ColorButton     =   14871017
               DrawFocusRectangle=   0   'False
            End
            Begin ImpulseButton.ISButton btnModify4 
               Height          =   330
               Left            =   3795
               TabIndex        =   145
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
               ButtonImage     =   "FrmCitiesDistance.frx":1874C
               ColorButton     =   14871017
               DrawFocusRectangle=   0   'False
            End
            Begin ImpulseButton.ISButton BtnUndo4 
               Height          =   330
               Left            =   2280
               TabIndex        =   146
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
               ButtonImage     =   "FrmCitiesDistance.frx":18AE6
               ColorButton     =   14871017
               DrawFocusRectangle=   0   'False
            End
            Begin ImpulseButton.ISButton btnDelete4 
               Height          =   330
               Left            =   1440
               TabIndex        =   147
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
               ButtonImage     =   "FrmCitiesDistance.frx":18E80
               ColorButton     =   14871017
               DrawFocusRectangle=   0   'False
            End
            Begin ImpulseButton.ISButton ISButton13 
               Height          =   330
               Left            =   5880
               TabIndex        =   148
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
               ButtonImage     =   "FrmCitiesDistance.frx":1941A
               ColorButton     =   14737632
               DrawFocusRectangle=   0   'False
            End
            Begin ImpulseButton.ISButton ISButton14 
               Height          =   330
               Left            =   6045
               TabIndex        =   149
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
               ButtonImage     =   "FrmCitiesDistance.frx":197B4
               ColorButton     =   14871017
               DrawFocusRectangle=   0   'False
            End
            Begin ImpulseButton.ISButton ISButton15 
               Height          =   285
               Left            =   4725
               TabIndex        =   150
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
               ButtonImage     =   "FrmCitiesDistance.frx":19B4E
               ColorButton     =   14871017
               DrawFocusRectangle=   0   'False
               DisabledImageStyle=   1
            End
            Begin ImpulseButton.ISButton btnCancel4 
               Height          =   330
               Left            =   705
               TabIndex        =   151
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
               ButtonImage     =   "FrmCitiesDistance.frx":19EE8
               ColorButton     =   14871017
               DrawFocusRectangle=   0   'False
               DisabledImageStyle=   1
            End
            Begin VB.Label Label2 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "«·”Ã· «·Õ«·Ì:"
               Height          =   210
               Index           =   7
               Left            =   2505
               RightToLeft     =   -1  'True
               TabIndex        =   155
               Top             =   225
               Width           =   975
            End
            Begin VB.Label Label2 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "⁄œœ «·”Ã·« :"
               Height          =   210
               Index           =   6
               Left            =   810
               RightToLeft     =   -1  'True
               TabIndex        =   154
               Top             =   225
               Width           =   975
            End
            Begin VB.Label LabCurrRec4 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Height          =   210
               Left            =   1800
               RightToLeft     =   -1  'True
               TabIndex        =   153
               Top             =   240
               Width           =   675
            End
            Begin VB.Label LabCountRec4 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Height          =   210
               Left            =   240
               RightToLeft     =   -1  'True
               TabIndex        =   152
               Top             =   225
               Width           =   540
            End
         End
      End
      Begin C1SizerLibCtl.C1Elastic C1Elastic7 
         Height          =   7350
         Left            =   12630
         TabIndex        =   164
         TabStop         =   0   'False
         Top             =   45
         Width           =   10995
         _cx             =   19394
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
         Begin VB.Frame Frame10 
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   0  'None
            Height          =   720
            Left            =   0
            RightToLeft     =   -1  'True
            TabIndex        =   174
            Top             =   0
            Width           =   10995
            Begin VB.TextBox TxtVac_ID5 
               Alignment       =   1  'Right Justify
               Appearance      =   0  'Flat
               BackColor       =   &H000000FF&
               Height          =   240
               Left            =   5670
               RightToLeft     =   -1  'True
               TabIndex        =   179
               Top             =   150
               Visible         =   0   'False
               Width           =   945
            End
            Begin VB.TextBox TxtModFlg5 
               Alignment       =   1  'Right Justify
               Appearance      =   0  'Flat
               BackColor       =   &H0000FF00&
               Enabled         =   0   'False
               Height          =   285
               Left            =   4560
               RightToLeft     =   -1  'True
               TabIndex        =   178
               Text            =   "modflag"
               Top             =   90
               Visible         =   0   'False
               Width           =   465
            End
            Begin VB.Frame Frame11 
               BackColor       =   &H00E0E0E0&
               BorderStyle     =   0  'None
               Enabled         =   0   'False
               Height          =   375
               Left            =   540
               RightToLeft     =   -1  'True
               TabIndex        =   175
               Top             =   690
               Visible         =   0   'False
               Width           =   3105
               Begin MSDataListLib.DataCombo DataCombo6 
                  CausesValidation=   0   'False
                  Height          =   315
                  Left            =   -255
                  TabIndex        =   176
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
                  Index           =   29
                  Left            =   2160
                  RightToLeft     =   -1  'True
                  TabIndex        =   177
                  Top             =   45
                  Width           =   855
               End
            End
            Begin MSComctlLib.ImageList GrdImageList5 
               Left            =   7320
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
                     Picture         =   "FrmCitiesDistance.frx":1A282
                     Key             =   "CompanyName"
                  EndProperty
                  BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                     Picture         =   "FrmCitiesDistance.frx":1A61C
                     Key             =   "Ser"
                  EndProperty
                  BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                     Picture         =   "FrmCitiesDistance.frx":1A9B6
                     Key             =   "Vac_Name"
                  EndProperty
                  BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                     Picture         =   "FrmCitiesDistance.frx":1AD50
                     Key             =   "ShareCount"
                  EndProperty
                  BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                     Picture         =   "FrmCitiesDistance.frx":1B0EA
                     Key             =   "Dis_Count"
                  EndProperty
                  BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                     Picture         =   "FrmCitiesDistance.frx":1B484
                     Key             =   "Bouns"
                  EndProperty
                  BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                     Picture         =   "FrmCitiesDistance.frx":1B81E
                     Key             =   "SharesValue"
                  EndProperty
                  BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                     Picture         =   "FrmCitiesDistance.frx":1BDB8
                     Key             =   "BuyValue"
                  EndProperty
               EndProperty
            End
            Begin ImpulseButton.ISButton btnLast5 
               Height          =   315
               Left            =   570
               TabIndex        =   180
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
               ButtonImage     =   "FrmCitiesDistance.frx":1C152
               ColorButton     =   16777215
               AcclimateGrayTones=   -1  'True
               DrawFocusRectangle=   0   'False
               DisabledImageExtraction=   0
               DisabledImageStyle=   1
            End
            Begin ImpulseButton.ISButton btnNext5 
               Height          =   315
               Left            =   1035
               TabIndex        =   181
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
               ButtonImage     =   "FrmCitiesDistance.frx":1C4EC
               ColorButton     =   16777215
               DrawFocusRectangle=   0   'False
               DisabledImageStyle=   1
            End
            Begin ImpulseButton.ISButton btnPrevious5 
               Height          =   315
               Left            =   1635
               TabIndex        =   182
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
               ButtonImage     =   "FrmCitiesDistance.frx":1C886
               ColorButton     =   16777215
               DrawFocusRectangle=   0   'False
               DisabledImageStyle=   1
            End
            Begin ImpulseButton.ISButton btnFirst5 
               Height          =   315
               Left            =   2100
               TabIndex        =   183
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
               ButtonImage     =   "FrmCitiesDistance.frx":1CC20
               ColorButton     =   16777215
               DrawFocusRectangle=   0   'False
               DisabledImageStyle=   1
            End
            Begin VB.Label Label1 
               Alignment       =   2  'Center
               BackStyle       =   0  'Transparent
               Caption         =   "»Ì«‰«  «‰Ê«⁄ «·—œÊœ"
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
               Index           =   30
               Left            =   8175
               RightToLeft     =   -1  'True
               TabIndex        =   184
               Top             =   90
               Width           =   2670
            End
         End
         Begin VB.Frame Frame9 
            BackColor       =   &H00E2E9E9&
            BorderStyle     =   0  'None
            Enabled         =   0   'False
            Height          =   1740
            Left            =   0
            RightToLeft     =   -1  'True
            TabIndex        =   165
            Top             =   4440
            Width           =   10920
            Begin VB.TextBox TxtName5 
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
               Left            =   5520
               MaxLength       =   50
               RightToLeft     =   -1  'True
               TabIndex        =   170
               Tag             =   "⁄›Ê« Ì—ÃÏ «œŒ«·  «·„”«›… ﬂ„"
               Top             =   495
               Width           =   4170
            End
            Begin VB.TextBox TxtSerialID 
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
               Left            =   5520
               MaxLength       =   50
               RightToLeft     =   -1  'True
               TabIndex        =   169
               Top             =   60
               Width           =   4170
            End
            Begin VB.ComboBox Combo4 
               BackColor       =   &H80000018&
               Height          =   315
               ItemData        =   "FrmCitiesDistance.frx":1CFBA
               Left            =   2280
               List            =   "FrmCitiesDistance.frx":1CFCA
               RightToLeft     =   -1  'True
               Style           =   2  'Dropdown List
               TabIndex        =   168
               Top             =   1590
               Visible         =   0   'False
               Width           =   1005
            End
            Begin VB.TextBox TxtName5E 
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
               TabIndex        =   167
               Tag             =   "⁄›Ê« Ì—ÃÏ «œŒ«·  ﬁÌ„… —œ «·—Õ·…"
               Top             =   480
               Width           =   4050
            End
            Begin VB.TextBox Text3 
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
               Left            =   1320
               MaxLength       =   50
               RightToLeft     =   -1  'True
               TabIndex        =   166
               Tag             =   "⁄›Ê« Ì—ÃÏ «œŒ«·  «·„”«›… ﬂ„"
               Top             =   0
               Visible         =   0   'False
               Width           =   810
            End
            Begin VB.Label Label1 
               Alignment       =   2  'Center
               BackColor       =   &H00E2E9E9&
               Caption         =   "„”·”· "
               Height          =   195
               Index           =   28
               Left            =   9720
               RightToLeft     =   -1  'True
               TabIndex        =   173
               Top             =   120
               Width           =   1050
            End
            Begin VB.Label Label1 
               Alignment       =   2  'Center
               BackColor       =   &H00E2E9E9&
               Caption         =   "«·«”„ ⁄—»Ì"
               Height          =   285
               Index           =   27
               Left            =   9840
               RightToLeft     =   -1  'True
               TabIndex        =   172
               Top             =   510
               Width           =   1050
            End
            Begin VB.Label Label1 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "«·«”„ «‰Ã·Ì“Ì "
               Height          =   285
               Index           =   26
               Left            =   4200
               RightToLeft     =   -1  'True
               TabIndex        =   171
               Top             =   480
               Width           =   1170
            End
         End
         Begin VSFlex8UCtl.VSFlexGrid Grid5 
            Height          =   3570
            Left            =   150
            TabIndex        =   185
            Top             =   780
            Width           =   10770
            _cx             =   18997
            _cy             =   6297
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
            Cols            =   8
            FixedRows       =   1
            FixedCols       =   1
            RowHeightMin    =   320
            RowHeightMax    =   0
            ColWidthMin     =   0
            ColWidthMax     =   0
            ExtendLastCol   =   -1  'True
            FormatString    =   $"FrmCitiesDistance.frx":1CFE3
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
         Begin C1SizerLibCtl.C1Elastic C1Elastic8 
            Height          =   1155
            Left            =   2565
            TabIndex        =   186
            TabStop         =   0   'False
            Top             =   6135
            Width           =   5865
            _cx             =   10345
            _cy             =   2037
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
               TabIndex        =   187
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
               ButtonImage     =   "FrmCitiesDistance.frx":1D138
               ColorButton     =   14871017
               DrawFocusRectangle=   0   'False
            End
            Begin ImpulseButton.ISButton btnSave5 
               Height          =   330
               Left            =   3030
               TabIndex        =   188
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
               ButtonImage     =   "FrmCitiesDistance.frx":1D4D2
               ColorButton     =   14871017
               DrawFocusRectangle=   0   'False
            End
            Begin ImpulseButton.ISButton btnModify5 
               Height          =   330
               Left            =   3795
               TabIndex        =   189
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
               ButtonImage     =   "FrmCitiesDistance.frx":1D86C
               ColorButton     =   14871017
               DrawFocusRectangle=   0   'False
            End
            Begin ImpulseButton.ISButton BtnUndo5 
               Height          =   330
               Left            =   2280
               TabIndex        =   190
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
               ButtonImage     =   "FrmCitiesDistance.frx":1DC06
               ColorButton     =   14871017
               DrawFocusRectangle=   0   'False
            End
            Begin ImpulseButton.ISButton btnDelete5 
               Height          =   330
               Left            =   1440
               TabIndex        =   191
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
               ButtonImage     =   "FrmCitiesDistance.frx":1DFA0
               ColorButton     =   14871017
               DrawFocusRectangle=   0   'False
            End
            Begin ImpulseButton.ISButton ISButton16 
               Height          =   330
               Left            =   5880
               TabIndex        =   192
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
               ButtonImage     =   "FrmCitiesDistance.frx":1E53A
               ColorButton     =   14737632
               DrawFocusRectangle=   0   'False
            End
            Begin ImpulseButton.ISButton ISButton17 
               Height          =   330
               Left            =   6045
               TabIndex        =   193
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
               ButtonImage     =   "FrmCitiesDistance.frx":1E8D4
               ColorButton     =   14871017
               DrawFocusRectangle=   0   'False
            End
            Begin ImpulseButton.ISButton ISButton21 
               Height          =   285
               Left            =   4725
               TabIndex        =   194
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
               ButtonImage     =   "FrmCitiesDistance.frx":1EC6E
               ColorButton     =   14871017
               DrawFocusRectangle=   0   'False
               DisabledImageStyle=   1
            End
            Begin ImpulseButton.ISButton btnCancel5 
               Height          =   330
               Left            =   705
               TabIndex        =   195
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
               ButtonImage     =   "FrmCitiesDistance.frx":1F008
               ColorButton     =   14871017
               DrawFocusRectangle=   0   'False
               DisabledImageStyle=   1
            End
            Begin VB.Label LabCountRec5 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Height          =   210
               Left            =   240
               RightToLeft     =   -1  'True
               TabIndex        =   199
               Top             =   225
               Width           =   540
            End
            Begin VB.Label LabCurrRec5 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Height          =   210
               Left            =   1800
               RightToLeft     =   -1  'True
               TabIndex        =   198
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
               TabIndex        =   197
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
               TabIndex        =   196
               Top             =   225
               Width           =   975
            End
         End
      End
   End
End
Attribute VB_Name = "FrmCitiesDistance"
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
Public Indx As Integer
'##################################################################################
Dim RsSavRec2 As ADODB.Recordset
Dim BKGrndPic2 As ClsBackGroundPic
Dim RecId2 As String
''##################################################################################
Dim RsSavRec3 As ADODB.Recordset
Dim RsSavRec5 As ADODB.Recordset
Dim BKGrndPic3 As ClsBackGroundPic
Dim rs2 As ADODB.Recordset
Dim RecId3 As String

Private Sub btnCancel4_Click()
 Unload Me
End Sub

Private Sub btnDelete4_Click()

    Dim MSGType As Integer
    Dim BlnRecExist As Boolean
    Dim StrMSG  As String
    Dim Msg As String
    On Error GoTo ErrTrap
    If TxtTransID.Text <> "" Then
      '  MSGType = MsgBox("Â·  —€» ›Ì Õ–› Â–« «·”Ã·", vbQuestion + vbMsgBoxRtlReading + vbYesNo + vbMsgBoxRight, App.title)
    If SystemOptions.UserInterface = ArabicInterface Then
       MSGType = MsgBox("Â·  —€» ›Ì Õ–› Â–« «·”Ã·", vbQuestion + vbMsgBoxRtlReading + vbYesNo + vbMsgBoxRight, App.title)
   Else
       MSGType = MsgBox("Confirm Delete", vbQuestion + vbMsgBoxRtlReading + vbYesNo + vbMsgBoxRight, App.title)
   End If

        If MSGType = vbYes Then
            rs2.find "ID=" & val(TxtTransID.Text), , adSearchForward, 1
            rs2.delete
           ' MsgBox " „  ⁄„·Ì… «·Õ–› »‰Ã«Õ", vbInformation + vbMsgBoxRtlReading + vbOKOnly + vbMsgBoxRight, App.title
      If SystemOptions.UserInterface = ArabicInterface Then
           MsgBox " „  ⁄„·Ì… «·Õ–› »‰Ã«Õ", vbInformation + vbMsgBoxRtlReading + vbOKOnly + vbMsgBoxRight, App.title
      Else
           MsgBox "Delete Successfully", vbInformation + vbMsgBoxRtlReading + vbOKOnly + vbMsgBoxRight, App.title
      End If
            '------------------------------ Move Next ---------------------------.
            FillGridWithData4
            btnNext4_Click
        End If
    End If

    Exit Sub
ErrTrap:
 
    Select Case Err.Number

        Case -2147217873, -2147467259
          '  StrMSG = "⁄›Ê« ·« ÌÃÊ“ Õ–› «·”Ã· ·«— »«ÿÂ »»Ì«‰«  √Œ—Ì"
           If SystemOptions.UserInterface = ArabicInterface Then
            StrMSG = "⁄›Ê« ·« ÌÃÊ“ Õ–› «·”Ã· ·«— »«ÿÂ »»Ì«‰«  √Œ—Ì"
            Else
            StrMSG = "Sorry ... This record can not be deleted because it is linked to other data"
            End If
            rs2.CancelUpdate
            MsgBox StrMSG, vbOKOnly + vbMsgBoxRight, App.title
            'clear the ConnectiOn Errors
            Cn.Errors.Clear
    End Select

End Sub

Private Sub btnFirst4_Click()
    On Error GoTo ErrTrap
    Dim Msg As String
    If Me.TxtModFlg4.Text = "N" Then
        FindRec4 val(TxtVac_ID3.Text)
        Me.TxtModFlg4.Text = "R"
    End If
    TxtModFlg4.Text = "R"
    If rs2.RecordCount = 0 Then
        clear_all Me
        Exit Sub
    End If

BegnieWork:
    rs2.MoveFirst
    FiLLTXT4

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
            MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading + vbOKOnly, App.title
            rs2.Requery
            Resume BegnieWork
    End Select
End Sub

Private Sub btnLast4_Click()
   On Error GoTo ErrTrap
    Dim Msg As String
    If Me.TxtModFlg4.Text = "N" Then
        FindRec4 val(TxtTransID.Text)
        Me.TxtModFlg4.Text = "R"
    End If

    TxtModFlg4.Text = "R"

    If rs2.RecordCount = 0 Then
        clear_all Me
        Exit Sub
    End If

BegnieWork:

    rs2.MoveLast
    FiLLTXT4
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
            MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading + vbOKOnly, App.title
            rs2.Requery
            Resume BegnieWork
    End Select
End Sub





Private Sub btnModify4_Click()
    Dim Msg As String
    On Error GoTo ErrTrap

    If TxtTransID.Text <> "" Then
        TxtModFlg4.Text = "E"
        Me.TxtName3.SetFocus
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
            MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading + vbOKOnly, App.title
            If rs2.EditMode <> adEditNone Then
                rs2.CancelUpdate
            End If
    End Select
End Sub



Private Sub btnNew4_Click()

    Dim My_SQL As String
    Dim rs As ADODB.Recordset
    On Error GoTo ErrTrap
    Set rs = New ADODB.Recordset
    '-----------------------------------
    TxtName3.Text = ""
    TxtName3E.Text = ""
    DcbAccount.BoundText = ""
    DcbAccount2.BoundText = ""
    TxtAccount.Text = ""
    TxtAccount2.Text = ""
    '-----------------------------------
    TxtModFlg4.Text = "N"

    My_SQL = "TblTypesTransport"
    rs.Open My_SQL, Cn, adOpenKeyset, adLockOptimistic, adCmdTable
    If rs.RecordCount > 0 Then
        TxtTransID.Text = rs.RecordCount + 1
    Else
        TxtTransID.Text = 1
    End If
    rs.Close
    TxtName3.SetFocus
ErrTrap:
End Sub

Private Sub btnNew5_Click()

    Dim My_SQL As String
    Dim rs As ADODB.Recordset
    On Error GoTo ErrTrap
    Set rs = New ADODB.Recordset
    '-----------------------------------
    TxtName5.Text = ""
    TxtName5E.Text = ""
    '-----------------------------------
    TxtModFlg5.Text = "N"

    My_SQL = "TblTypesTripStatus"
    rs.Open My_SQL, Cn, adOpenKeyset, adLockOptimistic, adCmdTable
    If rs.RecordCount > 0 Then
        TxtSerialID.Text = rs.RecordCount + 1
    Else
        TxtSerialID.Text = 1
    End If
    rs.Close
    TxtName5.SetFocus
ErrTrap:
End Sub

Private Sub btnNext4_Click()
    On Error GoTo ErrTrap
    Dim Msg As String
    If Me.TxtModFlg4.Text = "N" Then
        FindRec4 val(TxtTransID.Text)
        Me.TxtModFlg4.Text = "R"
    End If

    TxtModFlg4.Text = "R"

    If rs2.RecordCount = 0 Then
        clear_all Me
        Exit Sub
    End If

BegnieWork:

    If rs2.EOF Then
        rs2.MoveLast
    Else
        rs2.MoveNext
        If rs2.EOF Then
            rs2.MoveLast
        End If
    End If

    FiLLTXT4
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
            MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading + vbOKOnly, App.title
            rs2.Requery
            Resume BegnieWork
    End Select
End Sub

Private Sub btnPrevious4_Click()
    On Error GoTo ErrTrap
    Dim Msg As String

    If Me.TxtModFlg4.Text = "N" Then
        FindRec4 val(TxtTransID.Text)
        Me.TxtModFlg4.Text = "R"
    End If

    TxtModFlg4.Text = "R"

    If rs2.RecordCount = 0 Then
        clear_all Me
        Exit Sub
    End If

BegnieWork:

    rs2.MovePrevious

    If rs2.BOF Then
        rs2.MoveFirst
    End If

    FiLLTXT4
    Exit Sub
ErrTrap:

    Select Case Err.Number

        Case -2147217885
         '   Msg = "⁄›Ê« ·ﬁœ  „ Õ–› Â–« «·”Ã· «· «·Ï" & Chr(13)
         '   Msg = Msg & "„‰ ﬁ»· „” Œœ„ √Œ— ⁄·Ï «·‘»ﬂ… " & Chr(13)
         '   Msg = Msg & "”Ê› Ì „  ÕœÌÀ «·»Ì«‰« "
         If SystemOptions.UserInterface = ArabicInterface Then
            Msg = "⁄›Ê« ·ﬁœ  „ Õ–› Â–« «·”Ã· «· «·Ï" & CHR(13)
            Msg = Msg & "„‰ ﬁ»· „” Œœ„ √Œ— ⁄·Ï «·‘»ﬂ… " & CHR(13)
            Msg = Msg & "”Ê› Ì „  ÕœÌÀ «·»Ì«‰« "
           Else
           Msg = "Sorry..The following record has been deleted" & CHR(13)
           Msg = Msg & "By another user on the network " & CHR(13)
           Msg = Msg & "The data will be updated " & CHR(13)
       End If
            MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading + vbOKOnly, App.title
            rs2.Requery
            Resume BegnieWork
    End Select
End Sub

Private Sub btnSave4_Click()
    Dim Msg As String
    Dim StrVacCode As String
    Dim StrVacName As String
    Dim CtrlTxt As Control
    '---------------------- check if data Vaclete -----------------------
    If TxtName3.Text = "" Then
    If SystemOptions.UserInterface = ArabicInterface Then
      MsgBox "«·—Ã«¡ «œŒ«· «·«”„ ⁄—»Ì"
     Else
     MsgBox "Please Eneter Name"
     End If
     TxtName3.SetFocus
     Exit Sub
    End If
    
    If TxtName3E.Text = "" Then
    If SystemOptions.UserInterface = ArabicInterface Then
        MsgBox "«·—Ã«¡ «œŒ«· «·«”„ «‰Ã·Ì“Ì"
     Else
     MsgBox "Please Enter Name"
     End If
     TxtName3E.SetFocus
     Exit Sub
    End If
    If Me.DcbAccount.BoundText = "" Then
    If SystemOptions.UserInterface = ArabicInterface Then
        MsgBox "«·—Ã«¡ «Œ Ì«— Õ”«» «·«Ì—«œ« "
     Else
     MsgBox "Please Select Account"
     End If
     DcbAccount.SetFocus
     Exit Sub
    End If
    If Me.DcbAccount2.BoundText = "" Then
    If SystemOptions.UserInterface = ArabicInterface Then
        MsgBox "«·—Ã«¡ «Œ Ì«— Õ”«» «·„’—Ê›« "
     Else
     MsgBox "Please Select Account"
     End If
     DcbAccount2.SetFocus
     Exit Sub
    End If
    
    ' -------------------------------------- txtmodflg type -------------------
    Select Case Me.TxtModFlg4.Text
    '------------------------------ new record ----------------------------
        Case "N"
    '------------------------- save record -----------------------------
            AddNewRec4
            btnLast4_Click
        Case "E"
    '----------------------------- save edit -------------------------------
            FiLLRec4
    End Select

    Exit Sub
ErrTrap:
   ' MsgBox "Â‰«ﬂ Œÿ√ „« ›Ì ≈œŒ«· «·»Ì«‰« ", vbOKOnly + vbMsgBoxRight, App.title
If SystemOptions.UserInterface = ArabicInterface Then
    MsgBox "Â‰«ﬂ Œÿ√ „« ›Ì ≈œŒ«· «·»Ì«‰« ", vbOKOnly + vbMsgBoxRight, App.title
  Else
   MsgBox "Error...douring entering data", vbOKOnly + vbMsgBoxRight, App.title
End If
End Sub

Private Sub C1Tab1_Switch(OldTab As Integer, NewTab As Integer, Cancel As Integer)
Indx = NewTab
End Sub

Private Sub C1Tab1_Validate(Cancel As Boolean)
Indx = C1Tab1.CurrTab
End Sub

Private Sub Cmd_Click()
RemoveGridRow
End Sub

Private Sub DcbAccount_Change()
DcbAccount_Click (0)
End Sub

Private Sub DcbAccount_Click(Area As Integer)
TxtAccount.Text = getAccountSerial_Code("Account_Serial", "Account_Code", DcbAccount.BoundText)
End Sub

Private Sub DcbAccount_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF3 Then
        Account_search.show
        Account_search.case_id = 201

    End If
End Sub

Private Sub DcbAccount2_Change()
DcbAccount2_Click (0)
End Sub

Private Sub DcbAccount2_Click(Area As Integer)
TxtAccount2.Text = getAccountSerial_Code("Account_Serial", "Account_Code", DcbAccount2.BoundText)
End Sub

Private Sub DcbAccount2_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF3 Then
        Account_search.show
        Account_search.case_id = 202

    End If
End Sub

Private Sub Form_Load()
    On Error GoTo ErrTrap
    Dim i As Integer
    Dim My_SQL As String
    Dim Dcombos As ClsDataCombos
    
    'Indx = 4

    If Indx = 0 Then
        My_SQL = "TBLCitiesDistance"
        Set BKGrndPic = New ClsBackGroundPic
        Set RsSavRec = New ADODB.Recordset
        RsSavRec.CursorLocation = adUseClient
        RsSavRec.Open My_SQL, Cn, adOpenKeyset, adLockPessimistic, adCmdTableDirect
        Me.TxtModFlg.Text = "R"
    ElseIf Indx = 1 Then
        My_SQL = "TblHarborsData"
        Set BKGrndPic2 = New ClsBackGroundPic
        Set RsSavRec2 = New ADODB.Recordset
        RsSavRec2.CursorLocation = adUseClient
        RsSavRec2.Open My_SQL, Cn, adOpenKeyset, adLockPessimistic, adCmdTableDirect
        Me.TxtModFlg2.Text = "R"
    ElseIf Indx = 2 Then
        My_SQL = "TblShipsData"
        Set BKGrndPic3 = New ClsBackGroundPic
        Set RsSavRec3 = New ADODB.Recordset
        RsSavRec3.CursorLocation = adUseClient
        RsSavRec3.Open My_SQL, Cn, adOpenKeyset, adLockPessimistic, adCmdTableDirect
        Me.TxtModFlg3.Text = "R"
       ElseIf Indx = 3 Then
        My_SQL = "TblTypesTransport"
        Set rs2 = New ADODB.Recordset
        rs2.CursorLocation = adUseClient
        rs2.Open My_SQL, Cn, adOpenKeyset, adLockPessimistic, adCmdTableDirect
        Me.TxtModFlg4.Text = "R"
    ElseIf Indx = 4 Then
        My_SQL = "TblTypesTripStatus"
        Set RsSavRec5 = New ADODB.Recordset
        RsSavRec5.CursorLocation = adUseClient
        RsSavRec5.Open My_SQL, Cn, adOpenKeyset, adLockPessimistic, adCmdTableDirect
        Me.TxtModFlg5.Text = "R"
        
    End If
    
    Resize_Form Me
    
    'load tblUsers -----------------------------------------------
    My_SQL = "select UserID,UserName From tblUsers "
    fill_combo DCUser, My_SQL
    Set Dcombos = New ClsDataCombos
    Dcombos.getCountriesGovernments Me.DcboCountryID
    Dcombos.getCountriesGovernments Me.DcboCountryID1
    Dcombos.GetAccountingCodes Me.DcbAccount, True, False
    Dcombos.GetAccountingCodes Me.DcbAccount2, True, False
    Dcombos.GetHarbors Me.DcbHarbor
    Set cSearch = New clsDCboSearch
    Set cSearch.Client = Me.DcboCountryID
    ModFgLib.LinkFgColWithDataCombo Grid, Grid.ColIndex("CountryID"), Me.DcboCountryID
    ModFgLib.LinkFgColWithDataCombo Grid, Grid.ColIndex("CountryID1"), Me.DcboCountryID1
    
    If Indx = 0 Then
        FillGridWithData

        With Me.Grid
            .Cell(flexcpPicture, 0, .ColIndex("GovernmentName")) = Me.GrdImageList.ListImages("Vac_Name").ExtractIcon
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

        If OPEN_NEW_SCREEN = True Then
            'btnNew_Click
        End If
    ElseIf Indx = 1 Then
        FillGridWithData2
        btnFirst2_Click
    ElseIf Indx = 2 Then
        FillGridWithData3
        btnFirst3_Click
    ElseIf Indx = 3 Then
        FillGridWithData4
        btnFirst4_Click
    ElseIf Indx = 4 Then
        FillGridWithData5
        btnFirst5_Click
        
    End If
    
    'C1Tab1.TabVisible(0) = False
    'C1Tab1.TabVisible(1) = False
    'C1Tab1.TabVisible(2) = False

    'C1Tab1.TabVisible(Indx) = True
    C1Tab1.CurrTab = Indx

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
   Cmd.Caption = "Delete"
   
    Me.Caption = "Distances Between Cities "
    Label1(2).Caption = Me.Caption
  With VSFlexGrid1
    .TextMatrix(0, .ColIndex("Ser")) = "Ser"
    .TextMatrix(0, .ColIndex("Fullcode")) = "Code"
    .TextMatrix(0, .ColIndex("CusName")) = "Customer Name"
    .TextMatrix(0, .ColIndex("PriceComplete")) = "Price Complete"
    .TextMatrix(0, .ColIndex("PriceWithoutPart")) = "PriceWithout Part"
  End With
    With Me.Grid
        .TextMatrix(0, .ColIndex("Ser")) = "Ser"
        .TextMatrix(0, .ColIndex("CountryID")) = "From"
        .TextMatrix(0, .ColIndex("CountryID1")) = "To"
        .TextMatrix(0, .ColIndex("Distance")) = "Distance"
        .TextMatrix(0, .ColIndex("KmPrice")) = "Km Price "
        .TextMatrix(0, .ColIndex("Desil")) = "Desil"
        .TextMatrix(0, .ColIndex("TravelPrice")) = "Travel Cost"
        .TextMatrix(0, .ColIndex("DriverPercentage")) = "DriverP ercentage"
        .TextMatrix(0, .ColIndex("DriverValue")) = "Driver Cost"
        
        .TextMatrix(0, .ColIndex("TravelPriceUsed")) = "Travel Cost used"
        .TextMatrix(0, .ColIndex("DriverPercentageUsed")) = "DriverP ercentage used"
        .TextMatrix(0, .ColIndex("DriverValueUsed")) = "Driver Cost used"
        

    End With
    lbl(55).Caption = "Port"
    Label1(18).Caption = "Name English"
    Label1(12).Caption = "Name English"
    Label1(25).Caption = "Name English"
    Label1(26).Caption = "Name English"
    Label1(24).Caption = "Name Arabic"
    Label1(16).Caption = "Name Arabic"
    Label1(15).Caption = "Name Arabic"
    Label1(27).Caption = "Name Arabic"
    Label1(14).Caption = "ID"
    Label1(17).Caption = "ID"
    Label1(23).Caption = "ID"
    Label1(28).Caption = "ID"
    Label1(3).Caption = "ID"
    Label1(0).Caption = "From"
    Label1(4).Caption = "To"
    Label1(1).Caption = "Distance"
    Label1(5).Caption = "Desil Cost KM"
    Label1(9).Caption = "Desil"
    Label1(6).Caption = "Trip Cost"
    Label1(7).Caption = "Driver Percentage"
    Label1(8).Caption = "Driver Cost"
    Label2(0).Caption = "Curr. Rec."
    LabCurrRec5.Caption = "Curr. Rec."
    LabCountRec5.Caption = "Rec. Count."
    Label2(1).Caption = "Rec. Count."
    Label2(3).Caption = "Curr. Rec."
    Label2(2).Caption = "Rec. Count."
    Label2(4).Caption = "Curr. Rec."
    Label2(5).Caption = "Rec. Count."
    Label2(7).Caption = "Curr. Rec."
    Label2(6).Caption = "Rec. Count."
    Label1(11).Caption = "Data Of Port"
    Label1(20).Caption = "Data Of Ships"
    Label1(22).Caption = "Types Of Transport"
    btnNew.Caption = "New"
    btnModify.Caption = "Modify"
    btnSave.Caption = "Save"
    BtnUndo.Caption = "Undo"
    btnDelete.Caption = "Delete"
    btnCancel.Caption = "Exit"
    btnNew2.Caption = "New"
    btnModify2.Caption = "Modify"
    btnSave2.Caption = "Save"
    BtnUndo2.Caption = "Undo"
    btnDelete2.Caption = "Delete"
    btnCancel2.Caption = "Exit"
    btnNew3.Caption = "New"
    btnModify3.Caption = "Modify"
    btnSave3.Caption = "Save"
    BtnUndo3.Caption = "Undo"
    btnDelete3.Caption = "Delete"
    btnCancel3.Caption = "Exit"
    btnNew4.Caption = "New"
    btnModify4.Caption = "Modify"
    btnSave4.Caption = "Save"
    BtnUndo4.Caption = "Undo"
    btnDelete4.Caption = "Delete"
    btnCancel4.Caption = "Exit"
    

    btnModify5.Caption = "Modify"
    btnSave5.Caption = "Save"
    BtnUndo5.Caption = "Undo"
    btnDelete5.Caption = "Delete"
    btnCancel5.Caption = "Exit"
With Grid3
.TextMatrix(0, .ColIndex("Ser")) = "Serial"
.TextMatrix(0, .ColIndex("ID")) = "No#"
.TextMatrix(0, .ColIndex("Name")) = "Name Arabic "
.TextMatrix(0, .ColIndex("NameE")) = "Name English"
End With

With Grid5
.TextMatrix(0, .ColIndex("Ser")) = "Serial"
.TextMatrix(0, .ColIndex("ID")) = "No#"
.TextMatrix(0, .ColIndex("Name")) = "Name Arabic "
.TextMatrix(0, .ColIndex("NameE")) = "Name English"
End With
With Grid2
.TextMatrix(0, .ColIndex("Ser")) = "Serial"
.TextMatrix(0, .ColIndex("ID")) = "No#"
.TextMatrix(0, .ColIndex("Name")) = "Name Arabic "
.TextMatrix(0, .ColIndex("NameE")) = "Name English"
End With
With Grid4
.TextMatrix(0, .ColIndex("Ser")) = "Serial"
.TextMatrix(0, .ColIndex("ID")) = "No#"
.TextMatrix(0, .ColIndex("Name")) = "Name Arabic "
.TextMatrix(0, .ColIndex("NameE")) = "Name English"
.TextMatrix(0, .ColIndex("Account_Name")) = "Revenue Account"
.TextMatrix(0, .ColIndex("Account_Name2")) = "Expense Account"
End With
   lbl(0).Caption = "Revenue Account"
    lbl(1).Caption = "Expense Account"
End Sub
Private Sub BtnCancel_Click()
    Unload Me
End Sub

Private Sub btnDelete_Click()
    Dim MSGType As Integer
    Dim BlnRecExist As Boolean
    Dim StrMSG  As String
    Dim Msg As String
    On Error GoTo ErrTrap

    If DoPremis(Do_Delete, Me.Name, True) = False Then
        Exit Sub
    End If

    If TxtVac_ID.Text <> "" Then
        'MSGType = MsgBox("Â·  —€» ›Ì Õ–› Â–« «·”Ã·", vbQuestion + vbMsgBoxRtlReading + vbYesNo + vbMsgBoxRight, App.title)
  If SystemOptions.UserInterface = ArabicInterface Then
   MSGType = MsgBox("Â·  —€» ›Ì Õ–› Â–« «·”Ã·", vbQuestion + vbMsgBoxRtlReading + vbYesNo + vbMsgBoxRight, App.title)
   Else
   MSGType = MsgBox("Confirm Delete", vbQuestion + vbMsgBoxRtlReading + vbYesNo + vbMsgBoxRight, App.title)
   End If

        If MSGType = vbYes Then
            Cn.Execute "Delete from TblCitiesDistanceVendor where CityDisID=" & val(TxtVac_ID.Text) & ""
            RsSavRec.find "CitiesDistanceID=" & val(TxtVac_ID.Text), , adSearchForward, 1
            RsSavRec.delete
         '   MsgBox " „  ⁄„·Ì… «·Õ–› »‰Ã«Õ", vbInformation + vbMsgBoxRtlReading + vbOKOnly + vbMsgBoxRight, App.title
       If SystemOptions.UserInterface = ArabicInterface Then
          MsgBox " „  ⁄„·Ì… «·Õ–› »‰Ã«Õ", vbInformation + vbMsgBoxRtlReading + vbOKOnly + vbMsgBoxRight, App.title
        Else
          MsgBox "Delete Successfully", vbInformation + vbMsgBoxRtlReading + vbOKOnly + vbMsgBoxRight, App.title
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
          '  StrMSG = "⁄›Ê« ·« ÌÃÊ“ Õ–› «·”Ã· ·«— »«ÿÂ »»Ì«‰«  √Œ—Ì"
         If SystemOptions.UserInterface = ArabicInterface Then
            StrMSG = "⁄›Ê« ·« ÌÃÊ“ Õ–› «·”Ã· ·«— »«ÿÂ »»Ì«‰«  √Œ—Ì"
            Else
            StrMSG = "Sorry ... This record can not be deleted because it is linked to other data"
          End If
            RsSavRec.CancelUpdate
            MsgBox StrMSG, vbOKOnly + vbMsgBoxRight, App.title
            'clear the ConnectiOn Errors
            Cn.Errors.Clear
    End Select

End Sub

Private Sub BtnFirst_Click()
    On Error GoTo ErrTrap

    Dim Msg As String

    If Me.TxtModFlg.Text = "N" Then
        FindRec val(TxtVac_ID.Text)
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
        '    Msg = "⁄›Ê« ·ﬁœ  „ Õ–› Â–« «·”Ã· «· «·Ï" & Chr(13)
        '    Msg = Msg & "„‰ ﬁ»· „” Œœ„ √Œ— ⁄·Ï «·‘»ﬂ… " & Chr(13)
        '    Msg = Msg & "”Ê› Ì „  ÕœÌÀ «·»Ì«‰« "
           If SystemOptions.UserInterface = ArabicInterface Then
            Msg = "⁄›Ê« ·ﬁœ  „ Õ–› Â–« «·”Ã· «· «·Ï" & CHR(13)
            Msg = Msg & "„‰ ﬁ»· „” Œœ„ √Œ— ⁄·Ï «·‘»ﬂ… " & CHR(13)
            Msg = Msg & "”Ê› Ì „  ÕœÌÀ «·»Ì«‰« "
     Else
           Msg = "Sorry..The following record has been deleted" & CHR(13)
           Msg = Msg & "By another user on the network " & CHR(13)
           Msg = Msg & "The data will be updated " & CHR(13)
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
        FindRec val(TxtVac_ID.Text)
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

    If TxtVac_ID.Text <> "" Then
        TxtModFlg = "E"
        Frm2.Enabled = True
        Me.txtDistance.SetFocus
            VSFlexGrid1.Rows = VSFlexGrid1.Rows + 1
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
    '-----------------------------------
    Me.TxtVac_ID.Text = ""
    Me.txtDistance.Text = ""
    Me.DcboCountryID.BoundText = ""
    Me.DcboCountryID1.BoundText = ""
    txtDistance.Text = ""
    txtTravelPrice.Text = ""
    txtDriverPercentage.Text = ""
    txtDriverValue.Text = ""
    txtTravelPriceUsed.Text = ""
    txtDriverValueUsed.Text = ""
    txtDriverPercentageUsed.Text = ""
    txtDriverValueUsed = ""
    txtDesil.Text = ""
    txtKmPrice.Text = ""
    '-----------------------------------
    TxtModFlg.Text = "N"
    VSFlexGrid1.Clear flexClearScrollable, flexClearEverything
    VSFlexGrid1.Rows = 2
    My_SQL = "TBLCitiesDistance"
    rs.Open My_SQL, Cn, adOpenKeyset, adLockOptimistic, adCmdTable

    If rs.RecordCount > 0 Then
        TxtSerial.Text = rs.RecordCount + 1
    Else
        TxtSerial.Text = 1
    End If

    rs.Close
    CmbType.ListIndex = 0
    txtDistance.SetFocus
ErrTrap:
End Sub

Private Sub BtnNext_Click()
    On Error GoTo ErrTrap
    Dim Msg As String

    If Me.TxtModFlg.Text = "N" Then
        FindRec val(TxtVac_ID.Text)
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
         '   Msg = "⁄›Ê« ·ﬁœ  „ Õ–› Â–« «·”Ã· «· «·Ï" & Chr(13)
         '   Msg = Msg & "„‰ ﬁ»· „” Œœ„ √Œ— ⁄·Ï «·‘»ﬂ… " & Chr(13)
         '   Msg = Msg & "”Ê› Ì „  ÕœÌÀ «·»Ì«‰« "
      If SystemOptions.UserInterface = ArabicInterface Then
            Msg = "⁄›Ê« ·ﬁœ  „ Õ–› Â–« «·”Ã· «· «·Ï" & CHR(13)
            Msg = Msg & "„‰ ﬁ»· „” Œœ„ √Œ— ⁄·Ï «·‘»ﬂ… " & CHR(13)
            Msg = Msg & "”Ê› Ì „  ÕœÌÀ «·»Ì«‰« "
           Else
           Msg = "Sorry..The following record has been deleted" & CHR(13)
           Msg = Msg & "By another user on the network " & CHR(13)
           Msg = Msg & "The data will be updated " & CHR(13)
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
        FindRec val(TxtVac_ID.Text)
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
        '    Msg = "⁄›Ê« ·ﬁœ  „ Õ–› Â–« «·”Ã· «· «·Ï" & Chr(13)
        '    Msg = Msg & "„‰ ﬁ»· „” Œœ„ √Œ— ⁄·Ï «·‘»ﬂ… " & Chr(13)
        '    Msg = Msg & "”Ê› Ì „  ÕœÌÀ «·»Ì«‰« "
       If SystemOptions.UserInterface = ArabicInterface Then
            Msg = "⁄›Ê« ·ﬁœ  „ Õ–› Â–« «·”Ã· «· «·Ï" & CHR(13)
            Msg = Msg & "„‰ ﬁ»· „” Œœ„ √Œ— ⁄·Ï «·‘»ﬂ… " & CHR(13)
            Msg = Msg & "”Ê› Ì „  ÕœÌÀ «·»Ì«‰« "
        Else
           Msg = "Sorry..The following record has been deleted" & CHR(13)
           Msg = Msg & "By another user on the network " & CHR(13)
           Msg = Msg & "The data will be updated " & CHR(13)
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
    '---------------------- check if data Vaclete -----------------------

    For Each CtrlTxt In Me.Controls

        If TypeOf CtrlTxt Is TextBox Or TypeOf CtrlTxt Is ComboBox Then
            If CtrlTxt.Text = "" And CtrlTxt.Tag <> "" And CtrlTxt.Enabled = True Then
                MsgBox CtrlTxt.Tag, vbOKOnly + vbMsgBoxRight, App.title
                CtrlTxt.SetFocus
                Exit Sub
            End If
        End If

    Next

    '------------------------------ check if Empcode exist ----------------------

    StrVacName = IsRecExist("TBLCitiesDistance", "GovernmentName", Trim(txtDistance.Text), "GovernmentName", "Vac_ID<>'" & Trim(TxtVac_ID.Text) & "'")

    If StrVacName <> "" Then
       ' Msg = "·ﬁœ ”»ﬁ  ”ÃÌ· Â–« «·‰Ê⁄ „‰ ﬁ»·"
    If SystemOptions.UserInterface = ArabicInterface Then
        Msg = "·ﬁœ ”»ﬁ  ”ÃÌ· Â–« «·‰Ê⁄ „‰ ﬁ»·"
      Else
        Msg = "I have already registered this type before"
      End If
         
        MsgBox Msg, vbOKOnly + vbMsgBoxRight, App.title
        txtDistance.SetFocus
    
        Exit Sub

    End If

    ' -------------------------------------- txtmodflg type -------------------
    Select Case Me.TxtModFlg.Text

            '------------------------------ new record ----------------------------
        Case "N"
      
            '------------------------- save record -----------------------------
            AddNewRec
            BtnLast_Click

        Case "E"

            '----------------------------- save edit -------------------------------
            FiLLRec
    End Select

    Exit Sub
ErrTrap:
   ' MsgBox "Â‰«ﬂ Œÿ√ „« ›Ì ≈œŒ«· «·»Ì«‰« ", vbOKOnly + vbMsgBoxRight, App.title
  If SystemOptions.UserInterface = ArabicInterface Then
    MsgBox "Â‰«ﬂ Œÿ√ „« ›Ì ≈œŒ«· «·»Ì«‰« ", vbOKOnly + vbMsgBoxRight, App.title
  Else
   MsgBox "Error...douring entering data", vbOKOnly + vbMsgBoxRight, App.title
End If

End Sub
 
Private Sub BtnUndo_Click()
    FindRec val(TxtVac_ID.Text)
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
        Msg = "Number of records before update" & vbCrLf & FristCount & vbCrLf & "Number of records after  update" & vbCrLf & LastCount
        
        If LastCount > FristCount Then
            Msg = Msg + vbCrLf & "Number of new records" & vbCrLf & LastCount - FristCount
        Else
            Msg = Msg + vbCrLf & "Number of records deleted" & vbCrLf & FristCount - LastCount
        End If
    End If
End If
    MsgBox Msg, vbOKOnly + vbMsgBoxRight + vbInformation, App.title
ErrTrap:
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, _
                             UnloadMode As Integer)
    Dim IntResult As String
    Dim StrMSG As String
    On Error GoTo ErrTrap
    
    If Indx = 0 Then
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
    ElseIf Indx = 1 Then
        If Me.TxtModFlg2.Text <> "R" Then
            Select Case Me.TxtModFlg2.Text
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
                    btnSave2_Click
                Case vbCancel
                    Cancel = True
            End Select
        End If
    ElseIf Indx = 2 Then
        If Me.TxtModFlg3.Text <> "R" Then
            Select Case Me.TxtModFlg3.Text
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
                    btnSave3_Click
                Case vbCancel
                    Cancel = True
            End Select
        End If
        
    ElseIf Indx = 5 Then
        If Me.TxtModFlg5.Text <> "R" Then
            Select Case Me.TxtModFlg5.Text
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
                    btnSave5_Click
                Case vbCancel
                    Cancel = True
            End Select
        End If
        
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

Private Sub Form_Activate()
    Me.ZOrder 0
End Sub
Public Sub AddNewRec()
    On Error GoTo ErrTrap
    Dim StrRecID As String
    StrRecID = new_id("TBLCitiesDistance", "CitiesDistanceID", "")
    RsSavRec.AddNew
    RsSavRec.Fields("CitiesDistanceID").value = IIf(StrRecID <> "", StrRecID, Null)
    FiLLRec
ErrTrap:
End Sub
Public Sub FiLLRec()
Dim sql As String
Dim rs2 As ADODB.Recordset
Dim i As Integer
    On Error GoTo ErrTrap
 If Me.TxtModFlg.Text = "E" Then
 Cn.Execute "Delete from TblCitiesDistanceVendor where CityDisID=" & val(TxtVac_ID.Text) & ""
 End If
    RsSavRec.Fields("CityFromId").value = IIf(DcboCountryID.BoundText <> 0, val(DcboCountryID.BoundText), Null)
    RsSavRec.Fields("CityToId").value = IIf(DcboCountryID1.BoundText <> 0, val(DcboCountryID1.BoundText), Null)
    RsSavRec.Fields("Distance").value = val(txtDistance.Text)
    RsSavRec.Fields("KmPrice").value = val(txtKmPrice.Text)
    RsSavRec.Fields("TravelPrice").value = val(txtTravelPrice.Text)
    RsSavRec.Fields("DriverPercentage").value = val(txtDriverPercentage.Text)
    RsSavRec.Fields("DriverValue").value = val(txtDriverValue.Text)
    
    
    RsSavRec.Fields("DriverPercentageUsed").value = val(txtDriverPercentageUsed.Text)
    RsSavRec.Fields("TravelPriceUsed").value = val(txtTravelPriceUsed.Text)
    RsSavRec.Fields("DriverValueUsed").value = val(txtDriverValueUsed.Text)
    RsSavRec.Fields("Desil").value = val(txtDesil.Text)
    RsSavRec.update
    sql = "Select * from TblCitiesDistanceVendor  where 1=-1"
    Set rs2 = New ADODB.Recordset
    rs2.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
    With Me.VSFlexGrid1
    For i = 1 To .Rows - 1
    If val(.TextMatrix(i, .ColIndex("CusID"))) <> 0 Then
    rs2.AddNew
    rs2("CityDisID").value = val(TxtSerial.Text)
    rs2("CusID").value = val(.TextMatrix(i, .ColIndex("CusID")))
    rs2("PriceWithoutPart").value = val(.TextMatrix(i, .ColIndex("PriceWithoutPart")))
    rs2("PriceComplete").value = val(.TextMatrix(i, .ColIndex("PriceComplete")))
    rs2.update
    End If
    Next i
    End With
  '  MsgBox " „  ⁄„·Ì… «·Õ›Ÿ »‰Ã«Õ", vbInformation + vbOKOnly + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
   If SystemOptions.UserInterface = ArabicInterface Then
    MsgBox " „  ⁄„·Ì… «·Õ›Ÿ »‰Ã«Õ", vbInformation + vbOKOnly + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
   Else
    MsgBox "Saved Successfully", vbInformation + vbOKOnly + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
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
    TxtVac_ID.Text = IIf(IsNull(RsSavRec.Fields("CitiesDistanceID").value), "", RsSavRec.Fields("CitiesDistanceID").value)
    Me.DcboCountryID.BoundText = IIf(IsNull(RsSavRec.Fields("CityFromId").value), 0, RsSavRec.Fields("CityFromId").value)
    Me.DcboCountryID1.BoundText = IIf(IsNull(RsSavRec.Fields("CityToId").value), 0, RsSavRec.Fields("CityToId").value)
    txtDistance.Text = IIf(IsNull(RsSavRec.Fields("Distance").value), 0, RsSavRec.Fields("Distance").value)
    txtKmPrice.Text = IIf(IsNull(RsSavRec.Fields("KmPrice").value), 0, RsSavRec.Fields("KmPrice").value)
    txtTravelPrice.Text = IIf(IsNull(RsSavRec.Fields("TravelPrice").value), 0, RsSavRec.Fields("TravelPrice").value)
    txtDriverPercentage.Text = IIf(IsNull(RsSavRec.Fields("DriverPercentage").value), 0, RsSavRec.Fields("DriverPercentage").value)
    txtDriverValue.Text = IIf(IsNull(RsSavRec.Fields("DriverValue").value), 0, RsSavRec.Fields("DriverValue").value)
    
    txtTravelPriceUsed.Text = IIf(IsNull(RsSavRec.Fields("TravelPriceUsed").value), 0, RsSavRec.Fields("TravelPriceUsed").value)
    txtDriverPercentageUsed.Text = IIf(IsNull(RsSavRec.Fields("DriverPercentageUsed").value), 0, RsSavRec.Fields("DriverPercentageUsed").value)
   
    txtDriverValueUsed.Text = IIf(IsNull(RsSavRec.Fields("DriverValueUsed").value), 0, RsSavRec.Fields("DriverValueUsed").value)
    
    txtDesil.Text = IIf(IsNull(RsSavRec.Fields("Desil").value), 0, RsSavRec.Fields("Desil").value)

    LabCurrRec.Caption = RsSavRec.AbsolutePosition
    LabCountRec.Caption = RsSavRec.RecordCount
FillGridRet
    With Grid

        For i = 1 To .Rows - 1

            If Trim(TxtVac_ID.Text) = .TextMatrix(i, .ColIndex("CitiesDistanceID")) Then
                TxtSerial.Text = .TextMatrix(i, .ColIndex("Ser"))
                .Row = i
                Exit Sub
            End If

        Next

    End With

ErrTrap:

End Sub
Sub FillGridRet()
Dim sql As String
Dim rs2 As ADODB.Recordset
Set rs2 = New ADODB.Recordset
Dim i As Integer
   VSFlexGrid1.Clear flexClearScrollable, flexClearEverything
    VSFlexGrid1.Rows = 2
sql = " SELECT     dbo.TblCitiesDistanceVendor.ID, dbo.TblCitiesDistanceVendor.CityDisID, dbo.TblCitiesDistanceVendor.CusID, dbo.TblCustemers.CusName,"
sql = sql & "                       dbo.TblCustemers.CusNamee , dbo.TblCustemers.Fullcode, dbo.TblCitiesDistanceVendor.PriceWithoutPart, dbo.TblCitiesDistanceVendor.PriceComplete"
sql = sql & "  FROM         dbo.TblCitiesDistanceVendor LEFT OUTER JOIN"
sql = sql & "                       dbo.TblCustemers ON dbo.TblCitiesDistanceVendor.CusID = dbo.TblCustemers.CusID"
sql = sql & "  Where (dbo.TblCitiesDistanceVendor.CityDisID = " & val(TxtVac_ID.Text) & ")"
rs2.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
If rs2.RecordCount > 0 Then
With VSFlexGrid1
rs2.MoveFirst
.Rows = rs2.RecordCount + 1
For i = 1 To .Rows - 1
.TextMatrix(i, .ColIndex("PriceComplete")) = IIf(IsNull(rs2("PriceComplete").value), "", rs2("PriceComplete").value)
.TextMatrix(i, .ColIndex("PriceWithoutPart")) = IIf(IsNull(rs2("PriceWithoutPart").value), "", rs2("PriceWithoutPart").value)
.TextMatrix(i, .ColIndex("CusID")) = IIf(IsNull(rs2("CusID").value), "", rs2("CusID").value)
.TextMatrix(i, .ColIndex("Fullcode")) = IIf(IsNull(rs2("Fullcode").value), "", rs2("Fullcode").value)
If SystemOptions.UserInterface = ArabicInterface Then
.TextMatrix(i, .ColIndex("CusName")) = IIf(IsNull(rs2("CusName").value), "", rs2("CusName").value)
Else
.TextMatrix(i, .ColIndex("CusName")) = IIf(IsNull(rs2("CusNamee").value), "", rs2("CusNamee").value)
End If
rs2.MoveNext
Next i
End With
End If
End Sub
Public Sub EditRec(StrTable As String, _
                   RecId As String)
    'My_SQL = "select * From " & StrTable & " where "
    'RsSavRec.Open My_SQL, cn, adOpenStatic, adLockReadOnly, adCmdTableDirect
    FiLLRec

End Sub
Private Sub Grid_EnterCell()
    On Error GoTo ErrTrap
    FindRec val(Me.Grid.TextMatrix(Me.Grid.Row, Me.Grid.ColIndex("CitiesDistanceID")))
ErrTrap:
End Sub
Private Sub Grid4_EnterCell()
    On Error GoTo ErrTrap
    FindRec4 val(Me.Grid4.TextMatrix(Me.Grid4.Row, Me.Grid4.ColIndex("ID")))
ErrTrap:
End Sub
Private Sub TxtAccount_KeyPress(KeyAscii As Integer)
DcbAccount.BoundText = getAccountSerial_Code("Account_Code", "Account_Serial", TxtAccount.Text)
End Sub

Private Sub TxtAccount2_KeyPress(KeyAscii As Integer)
DcbAccount2.BoundText = getAccountSerial_Code("Account_Code", "Account_Serial", TxtAccount2.Text)
End Sub

Private Sub txtDistance_Change()

    If Me.TxtModFlg = "E" Or Me.TxtModFlg = "N" Then
        txtTravelPrice = val(txtDistance) * val(txtKmPrice)
        
        txtTravelPriceUsed = val(txtDistance) * val(txtKmPrice)
    End If

End Sub

Private Sub txtDriverPercentage_Change()

    If Me.TxtModFlg = "E" Or Me.TxtModFlg = "N" Then
        If val(txtDriverPercentage) <> 0 Then
            txtDriverValue = val(txtDriverPercentage) / 100 * val(txtTravelPrice)
        End If
    End If

End Sub

Private Sub txtDriverPercentageUsed_Change()

    If Me.TxtModFlg = "E" Or Me.TxtModFlg = "N" Then
        If val(txtDriverPercentageUsed) <> 0 Then
            txtDriverValueUsed = val(txtDriverPercentageUsed) / 100 * val(txtTravelPriceUsed)
        End If
    End If

End Sub

Private Sub txtKmPrice_Change()

    If Me.TxtModFlg = "E" Or Me.TxtModFlg = "N" Then
        txtDesil = val(txtDistance) * val(txtKmPrice)
    End If

End Sub


Private Sub TxtModFlg4_Change()
    If TxtModFlg4.Text = "N" Then
        Frame8.Enabled = True
        Me.btnNew4.Enabled = False
        btnModify4.Enabled = False
        btnDelete4.Enabled = False
        Grid4.Enabled = False
        BtnUndo4.Enabled = True
        Me.btnSave4.Enabled = True
    ElseIf TxtModFlg4.Text = "R" Then
        Grid4.Enabled = True
        btnModify4.Enabled = True
        btnDelete4.Enabled = True
        Me.btnNew4.Enabled = True
        BtnUndo4.Enabled = False
        Me.btnSave4.Enabled = False
        btnNext4.Enabled = True
        btnPrevious4.Enabled = True
        btnFirst4.Enabled = True
        btnLast4.Enabled = True
        Frame8.Enabled = False
    ElseIf TxtModFlg4.Text = "E" Then
        Frame8.Enabled = True
        Me.btnNew4.Enabled = False
        btnModify4.Enabled = False
        btnDelete4.Enabled = False
        BtnUndo4.Enabled = True
        Me.btnSave4.Enabled = True
        Grid4.Enabled = False
        btnNext4.Enabled = False
        btnPrevious4.Enabled = False
        btnFirst4.Enabled = False
        btnLast4.Enabled = False
    End If
End Sub

Private Sub txtTravelPrice_Change()
    txtDriverPercentage_Change
End Sub

Private Sub txtTravelPriceUsed_Change()
    txtDriverPercentageUsed_Change
End Sub
Private Sub TxtVac_ID_Change()
    Dim TxtMod As String
    TxtMod = TxtModFlg.Text
    TxtModFlg.Text = ""
    TxtModFlg = TxtMod
End Sub

Public Function FindRec(ByVal RecId As Long)
    On Error GoTo ErrTrap
    RsSavRec.find "CitiesDistanceID=" & RecId, , adSearchForward, 1

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
        '    btnNext.Enabled = False
        '    btnPrevious.Enabled = False
        '    btnFirst.Enabled = False
        '    btnLast.Enabled = False
    
    ElseIf TxtModFlg.Text = "R" Then
        Frm2.Enabled = False
        Grid.Enabled = True
        btnModify.Enabled = False
        btnDelete.Enabled = False

        If TxtVac_ID.Text <> "" Then
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
    Dim rs As ADODB.Recordset
    Dim My_SQL As String

    Set rs = New ADODB.Recordset
    My_SQL = "select * From TBLCitiesDistance order by CitiesDistanceID"
    rs.Open My_SQL, Cn, adOpenKeyset, adLockReadOnly, adCmdText

    With Me.Grid
        .Rows = 2
        .Clear flexClearScrollable

        If rs.RecordCount > 0 Then
            .Rows = rs.RecordCount + 1
            rs.MoveFirst

            For i = 1 To .Rows - 1
                .TextMatrix(i, .ColIndex("Ser")) = i
               
                .TextMatrix(i, .ColIndex("CitiesDistanceID")) = IIf(IsNull(rs.Fields("CitiesDistanceID").value), "", rs.Fields("CitiesDistanceID").value)
           
                .TextMatrix(i, .ColIndex("CountryID")) = IIf(IsNull(rs.Fields("CityFromId").value), "", rs.Fields("CityFromId").value)
            
                .TextMatrix(i, .ColIndex("CountryID1")) = IIf(IsNull(rs.Fields("CityToId").value), "", rs.Fields("CityToId").value)
            
                .TextMatrix(i, .ColIndex("Distance")) = IIf(IsNull(rs.Fields("Distance").value), "", rs.Fields("Distance").value)
            
                .TextMatrix(i, .ColIndex("KmPrice")) = IIf(IsNull(rs.Fields("KmPrice").value), "", rs.Fields("KmPrice").value)
                    
                .TextMatrix(i, .ColIndex("TravelPrice")) = IIf(IsNull(rs.Fields("TravelPrice").value), "", rs.Fields("TravelPrice").value)
            
                .TextMatrix(i, .ColIndex("DriverPercentage")) = IIf(IsNull(rs.Fields("DriverPercentage").value), "", rs.Fields("DriverPercentage").value)
                .TextMatrix(i, .ColIndex("DriverValue")) = IIf(IsNull(rs.Fields("DriverValue").value), "", rs.Fields("DriverValue").value)
                 
                .TextMatrix(i, .ColIndex("TravelPriceUsed")) = IIf(IsNull(rs.Fields("TravelPriceUsed").value), "", rs.Fields("TravelPriceUsed").value)
            
                .TextMatrix(i, .ColIndex("DriverPercentageUsed")) = IIf(IsNull(rs.Fields("DriverPercentageUsed").value), "", rs.Fields("DriverPercentageUsed").value)
                 
                 
                 .TextMatrix(i, .ColIndex("DriverValueUsed")) = IIf(IsNull(rs.Fields("DriverValueUsed").value), "", rs.Fields("DriverValueUsed").value)
                
                .TextMatrix(i, .ColIndex("Desil")) = IIf(IsNull(rs.Fields("Desil").value), "", rs.Fields("Desil").value)
                    
                rs.MoveNext
            Next

            rs.Close
        End If

        .RowHeight(-1) = 300
    End With

ErrTrap:
End Sub
'#####################################################################################################################################################################################################################################
Private Sub btnNew2_Click()

    Dim My_SQL As String
    Dim rs As ADODB.Recordset

    'If DoPremis(Do_New, Me.Name, True) = False Then
    '    Exit Sub
    'End If

    On Error GoTo ErrTrap
    Set rs = New ADODB.Recordset
    
    Frame3.Enabled = True
    '-----------------------------------
    Name2.Text = ""
    NameE2.Text = ""
    '-----------------------------------
    TxtModFlg2.Text = "N"

    My_SQL = "TblHarborsData"
    rs.Open My_SQL, Cn, adOpenKeyset, adLockOptimistic, adCmdTable

    If rs.RecordCount > 0 Then
        ID2.Text = rs.RecordCount + 1
    Else
        ID2.Text = 1
    End If

    rs.Close
    Name2.SetFocus
ErrTrap:
End Sub
Private Sub btnModify2_Click()
    Dim Msg As String

    'If DoPremis(Do_Edit, Me.Name, True) = False Then
    '    Exit Sub
    'End If

    On Error GoTo ErrTrap

    If TxtVac_ID2.Text <> "" Then
        TxtModFlg2.Text = "E"
        Frame3.Enabled = True
        Me.Name2.SetFocus
    End If

    Exit Sub
ErrTrap:

    Select Case Err.Number

        Case -2147467259
            'Could not update; currently locked.
           ' Msg = "⁄›Ê«" & Chr(13)
           ' Msg = Msg & " ·«Ì„ﬂ‰  ⁄œÌ· Â–« «·”Ã· ›Ï «·Êﬁ  «·Õ«·Ï" & Chr(13)
           ' Msg = Msg & "ÕÌÀ «‰Â ﬁÌœ «· ⁄œÌ· „‰ ﬁ»· „” Œœ„ «Œ— ⁄·Ï «·‘»ﬂ…"
       If SystemOptions.UserInterface = ArabicInterface Then
            Msg = "⁄›Ê«" & CHR(13)
            Msg = Msg & " ·«Ì„ﬂ‰  ⁄œÌ· Â–« «·”Ã· ›Ï «·Êﬁ  «·Õ«·Ï" & CHR(13)
            Msg = Msg & "ÕÌÀ «‰Â ﬁÌœ «· ⁄œÌ· „‰ ﬁ»· „” Œœ„ «Œ— ⁄·Ï «·‘»ﬂ…"
         Else
            Msg = "Sorry..." & CHR(13)
            Msg = Msg & " This record can not be edited at this time" & CHR(13)
            Msg = Msg & "Because it was modified by another user on the network"
        End If
            MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading + vbOKOnly, App.title
        
            If RsSavRec.EditMode <> adEditNone Then
                RsSavRec.CancelUpdate
                'RsSavRec.Requery
            End If

    End Select

End Sub
Private Sub btnSave2_Click()

    'On Error GoTo ErrTrap
    
    Dim Msg As String
    Dim StrVacCode As String
    Dim StrVacName As String
    Dim CtrlTxt As Control
    '---------------------- check if data Vaclete -----------------------
    If Name2.Text = "" Then
        MsgBox "«·—Ã«¡ «œŒ«· «·«”„ ⁄—»Ì"
    End If
    
    If NameE2.Text = "" Then
        MsgBox "«·—Ã«¡ «œŒ«· «·«”„ «‰Ã·Ì“Ì"
    End If
    '------------------------------ check if Empcode exist ----------------------
    'StrVacName = IsRecExist("TBLCitiesDistance", "GovernmentName", Trim(txtDistance.Text), "GovernmentName", "Vac_ID<>'" & Trim(TxtVac_ID.Text) & "'")
    'If StrVacName <> "" Then
    '   Msg = "·ﬁœ ”»ﬁ  ”ÃÌ· Â–« «·‰Ê⁄ „‰ ﬁ»·"
         
     '   MsgBox Msg, vbOKOnly + vbMsgBoxRight, App.title
    '    txtDistance.SetFocus
    
   '     Exit Sub
   ' End If

    ' -------------------------------------- txtmodflg type -------------------
    Select Case Me.TxtModFlg2.Text
    '------------------------------ new record ----------------------------
        Case "N"
    '------------------------- save record -----------------------------
            AddNewRec2
            btnLast2_Click
        Case "E"
    '----------------------------- save edit -------------------------------
            FiLLRec2
    End Select

    Exit Sub
ErrTrap:
  '  MsgBox "Â‰«ﬂ Œÿ√ „« ›Ì ≈œŒ«· «·»Ì«‰« ", vbOKOnly + vbMsgBoxRight, App.title
  If SystemOptions.UserInterface = ArabicInterface Then
    MsgBox "Â‰«ﬂ Œÿ√ „« ›Ì ≈œŒ«· «·»Ì«‰« ", vbOKOnly + vbMsgBoxRight, App.title
  Else
   MsgBox "Error...douring entering data", vbOKOnly + vbMsgBoxRight, App.title
End If


End Sub
Public Sub AddNewRec2()

    'On Error GoTo ErrTrap
    
    Dim StrRecID As String
    
    StrRecID = new_id("TblHarborsData", "ID", "")
    
    RsSavRec2.AddNew
    RsSavRec2.Fields("ID").value = IIf(StrRecID <> "", StrRecID, Null)
    FiLLRec2
ErrTrap:
End Sub
Public Sub FiLLRec2()

    On Error GoTo ErrTrap
 

    RsSavRec2.Fields("Name").value = Name2.Text
    RsSavRec2.Fields("NameE").value = NameE2.Text

    RsSavRec2.update
    
   ' MsgBox " „  ⁄„·Ì… «·Õ›Ÿ »‰Ã«Õ", vbInformation + vbOKOnly + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
    If SystemOptions.UserInterface = ArabicInterface Then
    MsgBox " „  ⁄„·Ì… «·Õ›Ÿ »‰Ã«Õ", vbInformation + vbOKOnly + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
   Else
    MsgBox "Saved Successfully", vbInformation + vbOKOnly + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
   End If
    FillGridWithData2
    
    TxtModFlg2.Text = "R"

    Exit Sub
ErrTrap:

    If RsSavRec2.EditMode <> adEditNone Then
        RsSavRec2.CancelUpdate
    End If

End Sub
Private Sub btnLast2_Click()

    On Error GoTo ErrTrap

    Dim Msg As String

    If Me.TxtModFlg2.Text = "N" Then
        FindRec2 val(TxtVac_ID2.Text)
        Me.TxtModFlg2.Text = "R"
    End If

    TxtModFlg2.Text = "R"

    If RsSavRec2.RecordCount = 0 Then
        clear_all Me
        Exit Sub
    End If

BegnieWork:

    RsSavRec2.MoveLast
    FiLLTXT2
    Exit Sub

ErrTrap:

    Select Case Err.Number

        Case -2147217885
       '     Msg = "⁄›Ê« ·ﬁœ  „ Õ–› Â–« «·”Ã· «· «·Ï" & Chr(13)
       '     Msg = Msg & "„‰ ﬁ»· „” Œœ„ √Œ— ⁄·Ï «·‘»ﬂ… " & Chr(13)
       '     Msg = Msg & "”Ê› Ì „  ÕœÌÀ «·»Ì«‰« "
        If SystemOptions.UserInterface = ArabicInterface Then
            Msg = "⁄›Ê« ·ﬁœ  „ Õ–› Â–« «·”Ã· «· «·Ï" & CHR(13)
            Msg = Msg & "„‰ ﬁ»· „” Œœ„ √Œ— ⁄·Ï «·‘»ﬂ… " & CHR(13)
            Msg = Msg & "”Ê› Ì „  ÕœÌÀ «·»Ì«‰« "
        Else
           Msg = "Sorry..The following record has been deleted" & CHR(13)
           Msg = Msg & "By another user on the network " & CHR(13)
           Msg = Msg & "The data will be updated " & CHR(13)
       End If
            MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading + vbOKOnly, App.title
            RsSavRec.Requery
            Resume BegnieWork
    End Select

End Sub
Public Sub FillGridWithData2()

    On Error GoTo ErrTrap

    Dim i As Integer
    Dim rs As ADODB.Recordset
    Dim My_SQL As String

    Set rs = New ADODB.Recordset
    My_SQL = "select * From TblHarborsData order by ID"
    rs.Open My_SQL, Cn, adOpenKeyset, adLockReadOnly, adCmdText

    With Me.Grid2
        .Rows = 2
        .Clear flexClearScrollable

        If rs.RecordCount > 0 Then
            .Rows = rs.RecordCount + 1
            rs.MoveFirst

            For i = 1 To .Rows - 1
                .TextMatrix(i, .ColIndex("Ser")) = i
                
                .TextMatrix(i, .ColIndex("ID")) = IIf(IsNull(rs.Fields("ID").value), "", rs.Fields("ID").value)
               
                .TextMatrix(i, .ColIndex("Name")) = IIf(IsNull(rs.Fields("Name").value), "", rs.Fields("Name").value)
           
                .TextMatrix(i, .ColIndex("NameE")) = IIf(IsNull(rs.Fields("NameE").value), "", rs.Fields("NameE").value)
                
                rs.MoveNext
            Next

            rs.Close
        End If

        .RowHeight(-1) = 300
    End With

ErrTrap:
End Sub
Public Sub FiLLTXT2()

    On Error GoTo ErrTrap
    
    Dim i As Integer
    
    Frame3.Enabled = False
    
    TxtVac_ID2.Text = IIf(IsNull(RsSavRec2.Fields("ID").value), "", RsSavRec2.Fields("ID").value)

    Me.Name2.Text = IIf(IsNull(RsSavRec2.Fields("Name").value), "", RsSavRec2.Fields("Name").value)
    Me.NameE2.Text = IIf(IsNull(RsSavRec2.Fields("NameE").value), "", RsSavRec2.Fields("NameE").value)

    LabCurrRec2.Caption = RsSavRec2.AbsolutePosition
    LabCountRec2.Caption = RsSavRec2.RecordCount

    With Grid2

        For i = 1 To .Rows - 1

            If Trim(TxtVac_ID2.Text) = .TextMatrix(i, .ColIndex("ID")) Then
                ID2.Text = .TextMatrix(i, .ColIndex("ID"))
                .Row = i
                Exit Sub
            End If

        Next

    End With

ErrTrap:

End Sub

Private Sub BtnUndo2_Click()
    FindRec2 val(TxtVac_ID2.Text)
    Me.TxtModFlg2.Text = "R"
End Sub
Private Sub btnDelete2_Click()

    Dim MSGType As Integer
    Dim BlnRecExist As Boolean
    Dim StrMSG  As String
    Dim Msg As String
    
    On Error GoTo ErrTrap

    'If DoPremis(Do_Delete, Me.Name, True) = False Then
    '    Exit Sub
    'End If

    If TxtVac_ID2.Text <> "" Then
       ' MSGType = MsgBox("Â·  —€» ›Ì Õ–› Â–« «·”Ã·", vbQuestion + vbMsgBoxRtlReading + vbYesNo + vbMsgBoxRight, App.title)
   If SystemOptions.UserInterface = ArabicInterface Then
       MSGType = MsgBox("Â·  —€» ›Ì Õ–› Â–« «·”Ã·", vbQuestion + vbMsgBoxRtlReading + vbYesNo + vbMsgBoxRight, App.title)
   Else
       MSGType = MsgBox("Confirm Delete", vbQuestion + vbMsgBoxRtlReading + vbYesNo + vbMsgBoxRight, App.title)
   End If

        If MSGType = vbYes Then
            RsSavRec2.find "ID=" & val(TxtVac_ID2.Text), , adSearchForward, 1
            RsSavRec2.delete
           ' MsgBox " „  ⁄„·Ì… «·Õ–› »‰Ã«Õ", vbInformation + vbMsgBoxRtlReading + vbOKOnly + vbMsgBoxRight, App.title
     If SystemOptions.UserInterface = ArabicInterface Then
        MsgBox " „  ⁄„·Ì… «·Õ–› »‰Ã«Õ", vbInformation + vbMsgBoxRtlReading + vbOKOnly + vbMsgBoxRight, App.title
       Else
        MsgBox "Delete Successfully", vbInformation + vbMsgBoxRtlReading + vbOKOnly + vbMsgBoxRight, App.title
    End If
            '------------------------------ Move Next ---------------------------.
            FillGridWithData2
            btnNext2_Click
        End If
    End If

    Exit Sub
ErrTrap:
 
    Select Case Err.Number

        Case -2147217873, -2147467259
          '  StrMSG = "⁄›Ê« ·« ÌÃÊ“ Õ–› «·”Ã· ·«— »«ÿÂ »»Ì«‰«  √Œ—Ì"
             If SystemOptions.UserInterface = ArabicInterface Then
            StrMSG = "⁄›Ê« ·« ÌÃÊ“ Õ–› «·”Ã· ·«— »«ÿÂ »»Ì«‰«  √Œ—Ì"
            Else
            StrMSG = "Sorry ... This record can not be deleted because it is linked to other data"
            End If
            RsSavRec2.CancelUpdate
            MsgBox StrMSG, vbOKOnly + vbMsgBoxRight, App.title
            'clear the ConnectiOn Errors
            Cn.Errors.Clear
    End Select

End Sub

Private Sub btnNext2_Click()
    On Error GoTo ErrTrap
    Dim Msg As String

    If Me.TxtModFlg2.Text = "N" Then
        FindRec2 val(TxtVac_ID2.Text)
        Me.TxtModFlg2.Text = "R"
    End If

    TxtModFlg2.Text = "R"

    If RsSavRec2.RecordCount = 0 Then
        clear_all Me
        Exit Sub
    End If

BegnieWork:

    If RsSavRec2.EOF Then
        RsSavRec2.MoveLast
    Else
        RsSavRec2.MoveNext
        If RsSavRec2.EOF Then
            RsSavRec2.MoveLast
        End If
    End If

    FiLLTXT2
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
            MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading + vbOKOnly, App.title
            RsSavRec2.Requery
            Resume BegnieWork
    End Select

End Sub
Private Sub btnCancel2_Click()
    Unload Me
End Sub
Private Sub btnPrevious2_Click()
    On Error GoTo ErrTrap
    Dim Msg As String

    If Me.TxtModFlg2.Text = "N" Then
        FindRec2 val(TxtVac_ID2.Text)
        Me.TxtModFlg2.Text = "R"
    End If

    TxtModFlg2.Text = "R"

    If RsSavRec2.RecordCount = 0 Then
        clear_all Me
        Exit Sub
    End If

BegnieWork:

    RsSavRec2.MovePrevious

    If RsSavRec2.BOF Then
        RsSavRec2.MoveFirst
    End If

    FiLLTXT2
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
            MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading + vbOKOnly, App.title
            RsSavRec2.Requery
            Resume BegnieWork
    End Select

End Sub
Private Sub btnFirst2_Click()
    On Error GoTo ErrTrap

    Dim Msg As String

    If Me.TxtModFlg2.Text = "N" Then
        FindRec2 val(TxtVac_ID2.Text)
        Me.TxtModFlg.Text = "R"
    End If

    TxtModFlg2.Text = "R"

    If RsSavRec2.RecordCount = 0 Then
        clear_all Me
        Exit Sub
    End If

BegnieWork:
    RsSavRec2.MoveFirst
    FiLLTXT2

    Exit Sub

ErrTrap:

    Select Case Err.Number

        Case -2147217885
        '    Msg = "⁄›Ê« ·ﬁœ  „ Õ–› Â–« «·”Ã· «· «·Ï" & Chr(13)
        '    Msg = Msg & "„‰ ﬁ»· „” Œœ„ √Œ— ⁄·Ï «·‘»ﬂ… " & Chr(13)
        '    Msg = Msg & "”Ê› Ì „  ÕœÌÀ «·»Ì«‰« "
       If SystemOptions.UserInterface = ArabicInterface Then
            Msg = "⁄›Ê« ·ﬁœ  „ Õ–› Â–« «·”Ã· «· «·Ï" & CHR(13)
            Msg = Msg & "„‰ ﬁ»· „” Œœ„ √Œ— ⁄·Ï «·‘»ﬂ… " & CHR(13)
            Msg = Msg & "”Ê› Ì „  ÕœÌÀ «·»Ì«‰« "
       Else
           Msg = "Sorry..The following record has been deleted" & CHR(13)
           Msg = Msg & "By another user on the network " & CHR(13)
           Msg = Msg & "The data will be updated " & CHR(13)
      End If
            MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading + vbOKOnly, App.title
            RsSavRec2.Requery
            Resume BegnieWork
    End Select

End Sub
Public Function FindRec2(ByVal RecId As Long)
    On Error GoTo ErrTrap
    RsSavRec2.find "ID=" & RecId, , adSearchForward, 1

    If Not (RsSavRec2.EOF) Then
        FiLLTXT2
    End If

    Exit Function
ErrTrap:

    If RsSavRec.EditMode <> adEditNone Then
        RsSavRec2.CancelUpdate
        BtnUndo2_Click
    End If
End Function

Private Sub TxtModFlg2_Change()
    If TxtModFlg2.Text = "N" Then
        Frame3.Enabled = True
        Me.btnNew2.Enabled = False
        btnModify2.Enabled = False
        btnDelete2.Enabled = False
        Grid2.Enabled = False
        BtnUndo2.Enabled = True
        Me.btnSave2.Enabled = True
    ElseIf TxtModFlg2.Text = "R" Then
        Frame3.Enabled = False
        Grid2.Enabled = True
        btnModify2.Enabled = True
        btnDelete2.Enabled = True
        Me.btnNew2.Enabled = True
        BtnUndo2.Enabled = False
        Me.btnSave2.Enabled = False
        btnNext2.Enabled = True
        btnPrevious2.Enabled = True
        btnFirst2.Enabled = True
        btnLast2.Enabled = True
    ElseIf TxtModFlg2.Text = "E" Then
        Frame3.Enabled = True
        Me.btnNew2.Enabled = False
        btnModify2.Enabled = False
        btnDelete2.Enabled = False
        BtnUndo2.Enabled = True
        Me.btnSave2.Enabled = True
        Grid2.Enabled = False
        btnNext2.Enabled = False
        btnPrevious2.Enabled = False
        btnFirst2.Enabled = False
        btnLast2.Enabled = False
    End If
End Sub
Private Sub Grid2_EnterCell()
    On Error GoTo ErrTrap
    FindRec2 val(Me.Grid2.TextMatrix(Me.Grid2.Row, Me.Grid2.ColIndex("ID")))
ErrTrap:
End Sub
'#####################################################################################################################################################################################################################################
'#####################################################################################################################################################################################################################################
Private Sub btnNew3_Click()

    Dim My_SQL As String
    Dim rs As ADODB.Recordset

    'If DoPremis(Do_New, Me.Name, True) = False Then
    '    Exit Sub
    'End If

    On Error GoTo ErrTrap
    Set rs = New ADODB.Recordset
    
    Frame3.Enabled = True
    '-----------------------------------
    Name3.Text = ""
    NameE3.Text = ""
    '-----------------------------------
    TxtModFlg3.Text = "N"

    My_SQL = "TblshipsData"
    rs.Open My_SQL, Cn, adOpenKeyset, adLockOptimistic, adCmdTable

    If rs.RecordCount > 0 Then
        ID3.Text = rs.RecordCount + 1
    Else
        ID3.Text = 1
    End If

    rs.Close
    Name3.SetFocus
ErrTrap:
End Sub
Private Sub btnModify3_Click()
    Dim Msg As String

    'If DoPremis(Do_Edit, Me.Name, True) = False Then
    '    Exit Sub
    'End If

    On Error GoTo ErrTrap

    If TxtVac_ID3.Text <> "" Then
        TxtModFlg3.Text = "E"
        Frame33.Enabled = True
        Me.Name3.SetFocus
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
            MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading + vbOKOnly, App.title
        
            If RsSavRec3.EditMode <> adEditNone Then
                RsSavRec3.CancelUpdate
            End If
    End Select
End Sub

Private Sub btnModify5_Click()
    Dim Msg As String

    'If DoPremis(Do_Edit, Me.Name, True) = False Then
    '    Exit Sub
    'End If

    On Error GoTo ErrTrap

    If TxtVac_ID5.Text <> "" Then
        TxtModFlg5.Text = "E"
        Frame9.Enabled = True
        'Me.Name5.SetFocus
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
            MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading + vbOKOnly, App.title
        
            If RsSavRec5.EditMode <> adEditNone Then
                RsSavRec5.CancelUpdate
            End If
    End Select
End Sub


Private Sub btnSave3_Click()

    'On Error GoTo ErrTrap
    
    Dim Msg As String
    Dim StrVacCode As String
    Dim StrVacName As String
    Dim CtrlTxt As Control
    '---------------------- check if data Vaclete -----------------------
    If Name3.Text = "" Then
        MsgBox "«·—Ã«¡ «œŒ«· «·«”„ ⁄—»Ì"
    End If
    
    If NameE3.Text = "" Then
        MsgBox "«·—Ã«¡ «œŒ«· «·«”„ «‰Ã·Ì“Ì"
    End If
    '------------------------------ check if Empcode exist ----------------------
    'StrVacName = IsRecExist("TBLCitiesDistance", "GovernmentName", Trim(txtDistance.Text), "GovernmentName", "Vac_ID<>'" & Trim(TxtVac_ID.Text) & "'")
    'If StrVacName <> "" Then
    '   Msg = "·ﬁœ ”»ﬁ  ”ÃÌ· Â–« «·‰Ê⁄ „‰ ﬁ»·"
         
     '   MsgBox Msg, vbOKOnly + vbMsgBoxRight, App.title
    '    txtDistance.SetFocus
    
   '     Exit Sub
   ' End If

    ' -------------------------------------- txtmodflg type -------------------
    Select Case Me.TxtModFlg3.Text
    '------------------------------ new record ----------------------------
        Case "N"
    '------------------------- save record -----------------------------
            AddNewRec3
            btnLast3_Click
        Case "E"
    '----------------------------- save edit -------------------------------
            FiLLRec3
    End Select

    Exit Sub
ErrTrap:
  '  MsgBox "Â‰«ﬂ Œÿ√ „« ›Ì ≈œŒ«· «·»Ì«‰« ", vbOKOnly + vbMsgBoxRight, App.title
If SystemOptions.UserInterface = ArabicInterface Then
    MsgBox "Â‰«ﬂ Œÿ√ „« ›Ì ≈œŒ«· «·»Ì«‰« ", vbOKOnly + vbMsgBoxRight, App.title
  Else
   MsgBox "Error...douring entering data", vbOKOnly + vbMsgBoxRight, App.title
End If

End Sub

Private Sub btnSave5_Click()

    'On Error GoTo ErrTrap
    
    Dim Msg As String
    Dim StrVacCode As String
    Dim StrVacName As String
    Dim CtrlTxt As Control
    '---------------------- check if data Vaclete -----------------------
    If TxtName5.Text = "" Then
        MsgBox "«·—Ã«¡ «œŒ«· «·«”„ ⁄—»Ì"
    End If
    
    If TxtName5E.Text = "" Then
        MsgBox "«·—Ã«¡ «œŒ«· «·«”„ «‰Ã·Ì“Ì"
    End If
    '------------------------------ check if Empcode exist ----------------------
    'StrVacName = IsRecExist("TBLCitiesDistance", "GovernmentName", Trim(txtDistance.Text), "GovernmentName", "Vac_ID<>'" & Trim(TxtVac_ID.Text) & "'")
    'If StrVacName <> "" Then
    '   Msg = "·ﬁœ ”»ﬁ  ”ÃÌ· Â–« «·‰Ê⁄ „‰ ﬁ»·"
         
     '   MsgBox Msg, vbOKOnly + vbMsgBoxRight, App.title
    '    txtDistance.SetFocus
    
   '     Exit Sub
   ' End If

    ' -------------------------------------- txtmodflg type -------------------
    Select Case Me.TxtModFlg5.Text
    '------------------------------ new record ----------------------------
        Case "N"
    '------------------------- save record -----------------------------
            AddNewRec5
            btnLast5_Click
        Case "E"
    '----------------------------- save edit -------------------------------
            FiLLRec5
    End Select

    Exit Sub
ErrTrap:
  '  MsgBox "Â‰«ﬂ Œÿ√ „« ›Ì ≈œŒ«· «·»Ì«‰« ", vbOKOnly + vbMsgBoxRight, App.title
If SystemOptions.UserInterface = ArabicInterface Then
    MsgBox "Â‰«ﬂ Œÿ√ „« ›Ì ≈œŒ«· «·»Ì«‰« ", vbOKOnly + vbMsgBoxRight, App.title
  Else
   MsgBox "Error...douring entering data", vbOKOnly + vbMsgBoxRight, App.title
End If

End Sub


Public Sub AddNewRec4()

    On Error GoTo ErrTrap
    Dim StrRecID As String
    StrRecID = new_id("TblTypesTransport", "ID", "")
    rs2.AddNew
    rs2.Fields("ID").value = IIf(StrRecID <> "", StrRecID, Null)
    FiLLRec4
ErrTrap:
End Sub

Public Sub AddNewRec5()

    On Error GoTo ErrTrap
    Dim StrRecID As String
    StrRecID = new_id("TblTypesTripStatus", "ID", "")
    RsSavRec5.AddNew
    RsSavRec5.Fields("ID").value = IIf(StrRecID <> "", StrRecID, Null)
    FiLLRec5
ErrTrap:
End Sub

Public Sub AddNewRec3()

    On Error GoTo ErrTrap
    
    Dim StrRecID As String
    
    StrRecID = new_id("TblShipsData", "ID", "")
    
    RsSavRec3.AddNew
    RsSavRec3.Fields("ID").value = IIf(StrRecID <> "", StrRecID, Null)
    FiLLRec3
ErrTrap:
End Sub
Public Sub FiLLRec3()
    On Error GoTo ErrTrap
    RsSavRec3.Fields("Name").value = Name3.Text
    RsSavRec3.Fields("NameE").value = NameE3.Text
    RsSavRec3.Fields("HarborID").value = val(DcbHarbor.BoundText)
    RsSavRec3.update
   ' MsgBox " „  ⁄„·Ì… «·Õ›Ÿ »‰Ã«Õ", vbInformation + vbOKOnly + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
   If SystemOptions.UserInterface = ArabicInterface Then
    MsgBox " „  ⁄„·Ì… «·Õ›Ÿ »‰Ã«Õ", vbInformation + vbOKOnly + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
   Else
    MsgBox "Saved Successfully", vbInformation + vbOKOnly + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
   End If
    FillGridWithData3
    TxtModFlg3.Text = "R"
    Exit Sub
ErrTrap:
    If RsSavRec3.EditMode <> adEditNone Then
        RsSavRec3.CancelUpdate
    End If
End Sub

Public Sub FiLLRec4()
    On Error GoTo ErrTrap
    rs2.Fields("Name").value = TxtName3.Text
    rs2.Fields("NameE").value = TxtName3E.Text
    rs2.Fields("AccountRevenue").value = Me.DcbAccount.BoundText
    rs2.Fields("AccountExpense").value = Me.DcbAccount2.BoundText
    rs2.update
    'MsgBox " „  ⁄„·Ì… «·Õ›Ÿ »‰Ã«Õ", vbInformation + vbOKOnly + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
    If SystemOptions.UserInterface = ArabicInterface Then
    MsgBox " „  ⁄„·Ì… «·Õ›Ÿ »‰Ã«Õ", vbInformation + vbOKOnly + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
   Else
    MsgBox "Saved Successfully", vbInformation + vbOKOnly + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
   End If
    FillGridWithData4
    TxtModFlg4.Text = "R"
    Exit Sub
ErrTrap:

    If rs2.EditMode <> adEditNone Then
        rs2.CancelUpdate
    End If

End Sub


Public Sub FiLLRec5()
    On Error GoTo ErrTrap
    RsSavRec5.Fields("Name").value = TxtName5.Text
    RsSavRec5.Fields("NameE").value = TxtName5E.Text
    
    RsSavRec5.update
    'MsgBox " „  ⁄„·Ì… «·Õ›Ÿ »‰Ã«Õ", vbInformation + vbOKOnly + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
    If SystemOptions.UserInterface = ArabicInterface Then
    MsgBox " „  ⁄„·Ì… «·Õ›Ÿ »‰Ã«Õ", vbInformation + vbOKOnly + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
   Else
    MsgBox "Saved Successfully", vbInformation + vbOKOnly + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
   End If
    FillGridWithData5
    TxtModFlg5.Text = "R"
    Exit Sub
ErrTrap:

    If RsSavRec5.EditMode <> adEditNone Then
        RsSavRec5.CancelUpdate
    End If

End Sub


Private Sub btnLast3_Click()

    On Error GoTo ErrTrap

    Dim Msg As String

    If Me.TxtModFlg3.Text = "N" Then
        FindRec3 val(TxtVac_ID3.Text)
        Me.TxtModFlg3.Text = "R"
    End If

    TxtModFlg3.Text = "R"

    If RsSavRec3.RecordCount = 0 Then
        clear_all Me
        Exit Sub
    End If

BegnieWork:

    RsSavRec3.MoveLast
    FiLLTXT3
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
            MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading + vbOKOnly, App.title
            RsSavRec3.Requery
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

    TxtModFlg5.Text = "R"

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
            MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading + vbOKOnly, App.title
            RsSavRec5.Requery
            Resume BegnieWork
    End Select

End Sub


Public Sub FillGridWithData3()

    On Error GoTo ErrTrap

    Dim i As Integer
    Dim rs As ADODB.Recordset
    Dim My_SQL As String

    Set rs = New ADODB.Recordset
    My_SQL = "select * From TblShipsData order by ID"
    rs.Open My_SQL, Cn, adOpenKeyset, adLockReadOnly, adCmdText

    With Me.Grid3
        .Rows = 2
        .Clear flexClearScrollable

        If rs.RecordCount > 0 Then
            .Rows = rs.RecordCount + 1
            rs.MoveFirst

            For i = 1 To .Rows - 1
                .TextMatrix(i, .ColIndex("ser")) = i
                
                .TextMatrix(i, .ColIndex("ID")) = IIf(IsNull(rs.Fields("ID").value), "", rs.Fields("ID").value)
               
                .TextMatrix(i, .ColIndex("Name")) = IIf(IsNull(rs.Fields("Name").value), "", rs.Fields("Name").value)
           
                .TextMatrix(i, .ColIndex("NameE")) = IIf(IsNull(rs.Fields("NameE").value), "", rs.Fields("NameE").value)
                
                rs.MoveNext
            Next

            rs.Close
        End If

        .RowHeight(-1) = 300
    End With

ErrTrap:
End Sub
Public Sub FillGridWithData4()

    On Error GoTo ErrTrap

    Dim i As Integer
    Dim rs As ADODB.Recordset
    Dim My_SQL As String

    Set rs = New ADODB.Recordset
    My_SQL = " SELECT     dbo.TblTypesTransport.ID, dbo.TblTypesTransport.Name, dbo.TblTypesTransport.NameE, dbo.TblTypesTransport.AccountRevenue, dbo.ACCOUNTS.Account_Name, "
    My_SQL = My_SQL & "                    dbo.ACCOUNTS.Account_Serial, dbo.ACCOUNTS.Account_NameEng, dbo.TblTypesTransport.AccountExpense, ACCOUNTS_1.Account_Name AS Account_Name2,"
    My_SQL = My_SQL & "                    ACCOUNTS_1.Account_Serial AS Account_Serial2, ACCOUNTS_1.Account_NameEng AS Account_NameEng2"
    My_SQL = My_SQL & "   FROM         dbo.TblTypesTransport LEFT OUTER JOIN"
    My_SQL = My_SQL & "                    dbo.ACCOUNTS ACCOUNTS_1 ON dbo.TblTypesTransport.AccountExpense = ACCOUNTS_1.Account_Code LEFT OUTER JOIN"
    My_SQL = My_SQL & "                    dbo.ACCOUNTS ON dbo.TblTypesTransport.AccountRevenue = dbo.ACCOUNTS.Account_Code"
    rs.Open My_SQL, Cn, adOpenKeyset, adLockReadOnly, adCmdText
    With Me.Grid4
        .Rows = 2
        .Clear flexClearScrollable

        If rs.RecordCount > 0 Then
            .Rows = rs.RecordCount + 1
            rs.MoveFirst
            For i = 1 To .Rows - 1
                .TextMatrix(i, .ColIndex("ser")) = i
                .TextMatrix(i, .ColIndex("ID")) = IIf(IsNull(rs.Fields("ID").value), "", rs.Fields("ID").value)
                .TextMatrix(i, .ColIndex("Name")) = IIf(IsNull(rs.Fields("Name").value), "", rs.Fields("Name").value)
                .TextMatrix(i, .ColIndex("NameE")) = IIf(IsNull(rs.Fields("NameE").value), "", rs.Fields("NameE").value)
                .TextMatrix(i, .ColIndex("AccountRevenue")) = IIf(IsNull(rs.Fields("AccountRevenue").value), "", rs.Fields("AccountRevenue").value)
                .TextMatrix(i, .ColIndex("AccountExpense")) = IIf(IsNull(rs.Fields("AccountExpense").value), "", rs.Fields("AccountExpense").value)
                If SystemOptions.UserInterface = ArabicInterface Then
                .TextMatrix(i, .ColIndex("Account_Name")) = IIf(IsNull(rs.Fields("Account_Name").value), "", rs.Fields("Account_Name").value)
                .TextMatrix(i, .ColIndex("Account_Name2")) = IIf(IsNull(rs.Fields("Account_Name2").value), "", rs.Fields("Account_Name2").value)
                Else
                .TextMatrix(i, .ColIndex("Account_Name")) = IIf(IsNull(rs.Fields("Account_NameEng").value), "", rs.Fields("Account_NameEng").value)
                .TextMatrix(i, .ColIndex("Account_Name2")) = IIf(IsNull(rs.Fields("Account_NameEng2").value), "", rs.Fields("Account_NameEng2").value)
                End If
                rs.MoveNext
            Next

            rs.Close
        End If

        .RowHeight(-1) = 300
    End With

ErrTrap:
End Sub

Public Sub FillGridWithData5()

    On Error GoTo ErrTrap

    Dim i As Integer
    Dim rs As ADODB.Recordset
    Dim My_SQL As String

    Set rs = New ADODB.Recordset
    My_SQL = "select * From TblTypesTripStatus order by ID"
    rs.Open My_SQL, Cn, adOpenKeyset, adLockReadOnly, adCmdText

    With Me.Grid5
        .Rows = 2
        .Clear flexClearScrollable

        If rs.RecordCount > 0 Then
            .Rows = RsSavRec5.RecordCount + 1
            rs.MoveFirst

            For i = 1 To .Rows - 1
                .TextMatrix(i, .ColIndex("ser")) = i
                
                .TextMatrix(i, .ColIndex("ID")) = IIf(IsNull(rs.Fields("ID").value), "", rs.Fields("ID").value)
               
                .TextMatrix(i, .ColIndex("Name")) = IIf(IsNull(rs.Fields("Name").value), "", rs.Fields("Name").value)
           
                .TextMatrix(i, .ColIndex("NameE")) = IIf(IsNull(rs.Fields("NameE").value), "", rs.Fields("NameE").value)
                
                rs.MoveNext
            Next

            rs.Close
        End If

        .RowHeight(-1) = 300
    End With

ErrTrap:
End Sub

Public Sub FiLLTXT3()

    On Error GoTo ErrTrap
    
    Dim i As Integer
    
    Frame33.Enabled = False
    
    TxtVac_ID3.Text = IIf(IsNull(RsSavRec3.Fields("ID").value), "", RsSavRec3.Fields("ID").value)

    Me.Name3.Text = IIf(IsNull(RsSavRec3.Fields("Name").value), "", RsSavRec3.Fields("Name").value)
    Me.NameE3.Text = IIf(IsNull(RsSavRec3.Fields("NameE").value), "", RsSavRec3.Fields("NameE").value)
DcbHarbor.BoundText = IIf(IsNull(RsSavRec3.Fields("HarborID").value), "", RsSavRec3.Fields("HarborID").value)

    LabCurrRec3.Caption = RsSavRec3.AbsolutePosition
    LabCountRec3.Caption = RsSavRec3.RecordCount

    With Grid3

        For i = 1 To .Rows - 1

            If Trim(TxtVac_ID3.Text) = .TextMatrix(i, .ColIndex("ID")) Then
                ID3.Text = .TextMatrix(i, .ColIndex("ID"))
                .Row = i
                Exit Sub
            End If
        Next
    End With

ErrTrap:

End Sub
Public Sub FiLLTXT4()
    On Error GoTo ErrTrap
    Dim i As Integer
    TxtTransID.Text = IIf(IsNull(rs2.Fields("ID").value), "", rs2.Fields("ID").value)
    Me.TxtName3.Text = IIf(IsNull(rs2.Fields("Name").value), "", rs2.Fields("Name").value)
    Me.TxtName3E.Text = IIf(IsNull(rs2.Fields("NameE").value), "", rs2.Fields("NameE").value)
   Me.DcbAccount.BoundText = IIf(IsNull(rs2.Fields("AccountRevenue").value), "", rs2.Fields("AccountRevenue").value)
   Me.DcbAccount2.BoundText = IIf(IsNull(rs2.Fields("AccountExpense").value), "", rs2.Fields("AccountExpense").value)
    LabCurrRec4.Caption = rs2.AbsolutePosition
    LabCountRec4.Caption = rs2.RecordCount
    With Grid4
        For i = 1 To .Rows - 1
            If Trim(TxtTransID.Text) = .TextMatrix(i, .ColIndex("ID")) Then
                TxtTransID.Text = .TextMatrix(i, .ColIndex("ID"))
                .Row = i
                Exit Sub
            End If
        Next
    End With

ErrTrap:

End Sub

Public Sub FiLLTXT5()

    On Error GoTo ErrTrap
    
    Dim i As Integer
    
    Frame9.Enabled = False
            
    TxtVac_ID5.Text = IIf(IsNull(RsSavRec5.Fields("ID").value), "", RsSavRec5.Fields("ID").value)

    TxtName5.Text = IIf(IsNull(RsSavRec5.Fields("Name").value), "", RsSavRec5.Fields("Name").value)
    TxtName5E.Text = IIf(IsNull(RsSavRec5.Fields("NameE").value), "", RsSavRec5.Fields("NameE").value)


    LabCurrRec5.Caption = RsSavRec5.AbsolutePosition
    LabCountRec5.Caption = RsSavRec5.RecordCount

    With Grid5

        For i = 1 To .Rows - 1

            If Trim(TxtVac_ID5.Text) = .TextMatrix(i, .ColIndex("ID")) Then
                TxtSerialID.Text = .TextMatrix(i, .ColIndex("ID"))
                .Row = i
                Exit Sub
            End If
        Next
    End With

ErrTrap:

End Sub

Private Sub BtnUndo4_Click()
    FindRec4 val(TxtTransID.Text)
    Me.TxtModFlg4.Text = "R"
End Sub


Private Sub BtnUndo3_Click()
    FindRec3 val(TxtVac_ID3.Text)
    Me.TxtModFlg3.Text = "R"
End Sub



Private Sub btnDelete3_Click()

    Dim MSGType As Integer
    Dim BlnRecExist As Boolean
    Dim StrMSG  As String
    Dim Msg As String
    
    On Error GoTo ErrTrap

    'If DoPremis(Do_Delete, Me.Name, True) = False Then
    '    Exit Sub
    'End If

    If TxtVac_ID3.Text <> "" Then
        'MSGType = MsgBox("Â·  —€» ›Ì Õ–› Â–« «·”Ã·", vbQuestion + vbMsgBoxRtlReading + vbYesNo + vbMsgBoxRight, App.title)
    If SystemOptions.UserInterface = ArabicInterface Then
       MSGType = MsgBox("Â·  —€» ›Ì Õ–› Â–« «·”Ã·", vbQuestion + vbMsgBoxRtlReading + vbYesNo + vbMsgBoxRight, App.title)
   Else
       MSGType = MsgBox("Confirm Delete", vbQuestion + vbMsgBoxRtlReading + vbYesNo + vbMsgBoxRight, App.title)
   End If

        If MSGType = vbYes Then
            RsSavRec3.find "ID=" & val(TxtVac_ID3.Text), , adSearchForward, 1
            RsSavRec3.delete
           ' MsgBox " „  ⁄„·Ì… «·Õ–› »‰Ã«Õ", vbInformation + vbMsgBoxRtlReading + vbOKOnly + vbMsgBoxRight, App.title
      If SystemOptions.UserInterface = ArabicInterface Then
        MsgBox " „  ⁄„·Ì… «·Õ–› »‰Ã«Õ", vbInformation + vbMsgBoxRtlReading + vbOKOnly + vbMsgBoxRight, App.title
        Else
        MsgBox "Delete Successfully", vbInformation + vbMsgBoxRtlReading + vbOKOnly + vbMsgBoxRight, App.title
      End If
            '------------------------------ Move Next ---------------------------.
            FillGridWithData3
            btnNext3_Click
        End If
    End If

    Exit Sub
ErrTrap:
 
    Select Case Err.Number

        Case -2147217873, -2147467259
           ' StrMSG = "⁄›Ê« ·« ÌÃÊ“ Õ–› «·”Ã· ·«— »«ÿÂ »»Ì«‰«  √Œ—Ì"
           If SystemOptions.UserInterface = ArabicInterface Then
            StrMSG = "⁄›Ê« ·« ÌÃÊ“ Õ–› «·”Ã· ·«— »«ÿÂ »»Ì«‰«  √Œ—Ì"
            Else
            StrMSG = "Sorry ... This record can not be deleted because it is linked to other data"
            End If
            RsSavRec3.CancelUpdate
            MsgBox StrMSG, vbOKOnly + vbMsgBoxRight, App.title
            'clear the ConnectiOn Errors
            Cn.Errors.Clear
    End Select

End Sub

Private Sub btnNext3_Click()
    On Error GoTo ErrTrap
    Dim Msg As String

    If Me.TxtModFlg3.Text = "N" Then
        FindRec3 val(TxtVac_ID3.Text)
        Me.TxtModFlg3.Text = "R"
    End If

    TxtModFlg3.Text = "R"

    If RsSavRec3.RecordCount = 0 Then
        clear_all Me
        Exit Sub
    End If

BegnieWork:

    If RsSavRec3.EOF Then
        RsSavRec3.MoveLast
    Else
        RsSavRec3.MoveNext
        If RsSavRec3.EOF Then
            RsSavRec3.MoveLast
        End If
    End If

    FiLLTXT3
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
            MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading + vbOKOnly, App.title
            RsSavRec3.Requery
            Resume BegnieWork
    End Select

End Sub
Private Sub btnCancel3_Click()
    Unload Me
End Sub
Private Sub btnPrevious3_Click()
    On Error GoTo ErrTrap
    Dim Msg As String

    If Me.TxtModFlg3.Text = "N" Then
        FindRec3 val(TxtVac_ID3.Text)
        Me.TxtModFlg3.Text = "R"
    End If

    TxtModFlg3.Text = "R"

    If RsSavRec3.RecordCount = 0 Then
        clear_all Me
        Exit Sub
    End If

BegnieWork:

    RsSavRec3.MovePrevious

    If RsSavRec3.BOF Then
        RsSavRec3.MoveFirst
    End If

    FiLLTXT3
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
            MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading + vbOKOnly, App.title
            RsSavRec3.Requery
            Resume BegnieWork
    End Select

End Sub
Private Sub btnFirst3_Click()
    On Error GoTo ErrTrap

    Dim Msg As String

    If Me.TxtModFlg3.Text = "N" Then
        FindRec3 val(TxtVac_ID3.Text)
        Me.TxtModFlg3.Text = "R"
    End If

    TxtModFlg3.Text = "R"

    If RsSavRec3.RecordCount = 0 Then
        clear_all Me
        Exit Sub
    End If

BegnieWork:
    RsSavRec3.MoveFirst
    FiLLTXT3

    Exit Sub

ErrTrap:

    Select Case Err.Number

        Case -2147217885
        
     '       Msg = "⁄›Ê« ·ﬁœ  „ Õ–› Â–« «·”Ã· «· «·Ï" & Chr(13)
     '       Msg = Msg & "„‰ ﬁ»· „” Œœ„ √Œ— ⁄·Ï «·‘»ﬂ… " & Chr(13)
     '       Msg = Msg & "”Ê› Ì „  ÕœÌÀ «·»Ì«‰« "
      If SystemOptions.UserInterface = ArabicInterface Then
            Msg = "⁄›Ê« ·ﬁœ  „ Õ–› Â–« «·”Ã· «· «·Ï" & CHR(13)
            Msg = Msg & "„‰ ﬁ»· „” Œœ„ √Œ— ⁄·Ï «·‘»ﬂ… " & CHR(13)
            Msg = Msg & "”Ê› Ì „  ÕœÌÀ «·»Ì«‰« "
     Else
           Msg = "Sorry..The following record has been deleted" & CHR(13)
           Msg = Msg & "By another user on the network " & CHR(13)
           Msg = Msg & "The data will be updated " & CHR(13)
     End If
            MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading + vbOKOnly, App.title
            RsSavRec3.Requery
            Resume BegnieWork
    End Select

End Sub
Public Function FindRec4(ByVal RecId As Long)
    On Error GoTo ErrTrap
    rs2.find "ID=" & RecId, , adSearchForward, 1

    If Not (rs2.EOF) Then
        FiLLTXT4
    End If

    Exit Function
ErrTrap:

    If rs2.EditMode <> adEditNone Then
        rs2.CancelUpdate
        BtnUndo4_Click
    End If
End Function
Public Function FindRec3(ByVal RecId As Long)
    On Error GoTo ErrTrap
    RsSavRec3.find "ID=" & RecId, , adSearchForward, 1

    If Not (RsSavRec3.EOF) Then
        FiLLTXT3
    End If

    Exit Function
ErrTrap:

    If RsSavRec3.EditMode <> adEditNone Then
        RsSavRec3.CancelUpdate
        BtnUndo3_Click
    End If
End Function

Private Sub TxtModFlg3_Change()
    If TxtModFlg3.Text = "N" Then
        Frame33.Enabled = True
        Me.btnNew3.Enabled = False
        btnModify3.Enabled = False
        btnDelete3.Enabled = False
        Grid3.Enabled = False
        BtnUndo3.Enabled = True
        Me.btnSave3.Enabled = True
        
    ElseIf TxtModFlg3.Text = "R" Then
        Frame33.Enabled = False
        Grid3.Enabled = True
        btnModify3.Enabled = True
        btnDelete3.Enabled = True
        Me.btnNew3.Enabled = True
        BtnUndo3.Enabled = False
        Me.btnSave3.Enabled = False
        btnNext3.Enabled = True
        btnPrevious3.Enabled = True
        btnFirst3.Enabled = True
        btnLast3.Enabled = True
        
    ElseIf TxtModFlg3.Text = "E" Then
        Frame33.Enabled = True
        Me.btnNew3.Enabled = False
        btnModify3.Enabled = False
        btnDelete3.Enabled = False
        BtnUndo3.Enabled = True
        Me.btnSave3.Enabled = True
        Grid3.Enabled = False
        btnNext3.Enabled = False
        btnPrevious3.Enabled = False
        btnFirst3.Enabled = False
        btnLast3.Enabled = False
    End If
End Sub
Private Sub Grid3_EnterCell()
    On Error GoTo ErrTrap
    FindRec3 val(Me.Grid3.TextMatrix(Me.Grid3.Row, Me.Grid3.ColIndex("ID")))
ErrTrap:
End Sub






Private Sub BtnUndo5_Click()
    FindRec5 val(TxtVac_ID5.Text)
    Me.TxtModFlg5.Text = "R"
End Sub
Private Sub btnDelete5_Click()

    Dim MSGType As Integer
    Dim BlnRecExist As Boolean
    Dim StrMSG  As String
    Dim Msg As String
    
    On Error GoTo ErrTrap

    'If DoPremis(Do_Delete, Me.Name, True) = False Then
    '    Exit Sub
    'End If

    If TxtVac_ID5.Text <> "" Then
        'MSGType = MsgBox("Â·  —€» ›Ì Õ–› Â–« «·”Ã·", vbQuestion + vbMsgBoxRtlReading + vbYesNo + vbMsgBoxRight, App.title)
    If SystemOptions.UserInterface = ArabicInterface Then
       MSGType = MsgBox("Â·  —€» ›Ì Õ–› Â–« «·”Ã·", vbQuestion + vbMsgBoxRtlReading + vbYesNo + vbMsgBoxRight, App.title)
   Else
       MSGType = MsgBox("Confirm Delete", vbQuestion + vbMsgBoxRtlReading + vbYesNo + vbMsgBoxRight, App.title)
   End If

        If MSGType = vbYes Then
            RsSavRec5.find "ID=" & val(TxtVac_ID5.Text), , adSearchForward, 1
            RsSavRec5.delete
           ' MsgBox " „  ⁄„·Ì… «·Õ–› »‰Ã«Õ", vbInformation + vbMsgBoxRtlReading + vbOKOnly + vbMsgBoxRight, App.title
      If SystemOptions.UserInterface = ArabicInterface Then
        MsgBox " „  ⁄„·Ì… «·Õ–› »‰Ã«Õ", vbInformation + vbMsgBoxRtlReading + vbOKOnly + vbMsgBoxRight, App.title
        Else
        MsgBox "Delete Successfully", vbInformation + vbMsgBoxRtlReading + vbOKOnly + vbMsgBoxRight, App.title
      End If
            '------------------------------ Move Next ---------------------------.
            FillGridWithData5
            btnNext5_Click
        End If
    End If

    Exit Sub
ErrTrap:
 
    Select Case Err.Number

        Case -2147217873, -2147467259
           ' StrMSG = "⁄›Ê« ·« ÌÃÊ“ Õ–› «·”Ã· ·«— »«ÿÂ »»Ì«‰«  √Œ—Ì"
           If SystemOptions.UserInterface = ArabicInterface Then
            StrMSG = "⁄›Ê« ·« ÌÃÊ“ Õ–› «·”Ã· ·«— »«ÿÂ »»Ì«‰«  √Œ—Ì"
            Else
            StrMSG = "Sorry ... This record can not be deleted because it is linked to other data"
            End If
            RsSavRec5.CancelUpdate
            MsgBox StrMSG, vbOKOnly + vbMsgBoxRight, App.title
            'clear the ConnectiOn Errors
            Cn.Errors.Clear
    End Select

End Sub

Private Sub btnNext5_Click()
    On Error GoTo ErrTrap
    Dim Msg As String

    If Me.TxtModFlg5.Text = "N" Then
        FindRec5 val(TxtVac_ID5.Text)
        Me.TxtModFlg5.Text = "R"
    End If

    TxtModFlg5.Text = "R"

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
            MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading + vbOKOnly, App.title
            RsSavRec5.Requery
            Resume BegnieWork
    End Select

End Sub
Private Sub btnCancel5_Click()
    Unload Me
End Sub
Private Sub btnPrevious5_Click()
    On Error GoTo ErrTrap
    Dim Msg As String

    If Me.TxtModFlg5.Text = "N" Then
        FindRec5 val(TxtVac_ID5.Text)
        Me.TxtModFlg5.Text = "R"
    End If

    TxtModFlg5.Text = "R"

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
            MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading + vbOKOnly, App.title
            RsSavRec5.Requery
            Resume BegnieWork
    End Select

End Sub
Private Sub btnFirst5_Click()
    On Error GoTo ErrTrap

    Dim Msg As String

    If Me.TxtModFlg5.Text = "N" Then
        FindRec5 val(TxtVac_ID5.Text)
        Me.TxtModFlg5.Text = "R"
    End If

    TxtModFlg5.Text = "R"

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
        
     '       Msg = "⁄›Ê« ·ﬁœ  „ Õ–› Â–« «·”Ã· «· «·Ï" & Chr(13)
     '       Msg = Msg & "„‰ ﬁ»· „” Œœ„ √Œ— ⁄·Ï «·‘»ﬂ… " & Chr(13)
     '       Msg = Msg & "”Ê› Ì „  ÕœÌÀ «·»Ì«‰« "
      If SystemOptions.UserInterface = ArabicInterface Then
            Msg = "⁄›Ê« ·ﬁœ  „ Õ–› Â–« «·”Ã· «· «·Ï" & CHR(13)
            Msg = Msg & "„‰ ﬁ»· „” Œœ„ √Œ— ⁄·Ï «·‘»ﬂ… " & CHR(13)
            Msg = Msg & "”Ê› Ì „  ÕœÌÀ «·»Ì«‰« "
     Else
           Msg = "Sorry..The following record has been deleted" & CHR(13)
           Msg = Msg & "By another user on the network " & CHR(13)
           Msg = Msg & "The data will be updated " & CHR(13)
     End If
            MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading + vbOKOnly, App.title
            RsSavRec5.Requery
            Resume BegnieWork
    End Select

End Sub

Public Function FindRec5(ByVal RecId As Long)
    On Error GoTo ErrTrap
    RsSavRec5.find "ID=" & RecId, , adSearchForward, 1

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
        Frame9.Enabled = True
        Me.btnNew5.Enabled = False
        btnModify5.Enabled = False
        btnDelete5.Enabled = False
        Grid5.Enabled = False
        BtnUndo5.Enabled = True
        Me.btnSave5.Enabled = True
        
    ElseIf TxtModFlg5.Text = "R" Then
        Frame9.Enabled = False
        Grid5.Enabled = True
        btnModify5.Enabled = True
        btnDelete5.Enabled = True
        Me.btnNew5.Enabled = True
        BtnUndo5.Enabled = False
        Me.btnSave5.Enabled = False
        btnNext5.Enabled = True
        btnPrevious5.Enabled = True
        btnFirst5.Enabled = True
        btnLast5.Enabled = True
        
    ElseIf TxtModFlg5.Text = "E" Then
        Frame9.Enabled = True
        Me.btnNew5.Enabled = False
        btnModify5.Enabled = False
        btnDelete5.Enabled = False
        BtnUndo5.Enabled = True
        Me.btnSave5.Enabled = True
        Grid5.Enabled = False
        btnNext5.Enabled = False
        btnPrevious5.Enabled = False
        btnFirst5.Enabled = False
        btnLast5.Enabled = False
    End If
End Sub
Private Sub Grid5_EnterCell()
    On Error GoTo ErrTrap
    FindRec5 val(Me.Grid5.TextMatrix(Me.Grid5.Row, Me.Grid5.ColIndex("ID")))
ErrTrap:
End Sub



'-------------------------------------------------------------
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

Private Sub RemoveGridRow()
If Me.TxtModFlg.Text <> "R" Then
    With Me.VSFlexGrid1
        If .Row <= 0 Then Exit Sub
        .RemoveItem .Row
    End With
End If
End Sub
Private Sub VSFlexGrid1_AfterEdit(ByVal Row As Long, ByVal Col As Long)
  Dim StrAccountCode As String
    Dim Msg As String
    Dim rs2 As New ADODB.Recordset
    Dim StrSQL As String
    Dim ClsAcc As New ClsAccounts
    Dim LngRow As Long
    With VSFlexGrid1
        Select Case .ColKey(Col)
           Case "CusName"
                 StrAccountCode = .ComboData
                 LngRow = .FindRow(StrAccountCode, .FixedRows, .ColIndex("CusID"), False, True)
                .TextMatrix(Row, .ColIndex("CusID")) = StrAccountCode
                If val(.TextMatrix(Row, .ColIndex("CusID"))) <> 0 Then
                StrSQL = " select * from TblCustemers where CusID=" & val(.TextMatrix(Row, .ColIndex("CusID"))) & ""
                Set rs2 = New ADODB.Recordset
                rs2.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText
                If rs2.RecordCount > 0 Then
                .TextMatrix(Row, .ColIndex("Fullcode")) = IIf(IsNull(rs2("Fullcode").value), "", rs2("Fullcode").value)
                Else
                .TextMatrix(Row, .ColIndex("Fullcode")) = ""
                End If
                Else
                .TextMatrix(Row, .ColIndex("Fullcode")) = ""
                End If
          Case "Fullcode"

                If .TextMatrix(Row, .ColIndex("Fullcode")) <> "" Then
                StrSQL = " select * from TblCustemers where Fullcode='" & (.TextMatrix(Row, .ColIndex("Fullcode"))) & "'"
                StrSQL = StrSQL & "    and     (Type = 2)"
                StrSQL = StrSQL & "  and    (  BranchId=0  or      BranchId in(" & Current_branchSql & "))"
                Set rs2 = New ADODB.Recordset
                rs2.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText
                If rs2.RecordCount > 0 Then
                .TextMatrix(Row, .ColIndex("CusID")) = IIf(IsNull(rs2("CusID").value), "", rs2("CusID").value)
                If SystemOptions.UserInterface = ArabicInterface Then
                .TextMatrix(Row, .ColIndex("CusName")) = IIf(IsNull(rs2("CusName").value), "", rs2("CusName").value)
                Else
                .TextMatrix(Row, .ColIndex("CusName")) = IIf(IsNull(rs2("CusNamee").value), "", rs2("CusNamee").value)
                End If
                Else
                .TextMatrix(Row, .ColIndex("CusID")) = 0
                .TextMatrix(Row, .ColIndex("CusName")) = ""
                End If
                Else
                .TextMatrix(Row, .ColIndex("CusID")) = 0
                .TextMatrix(Row, .ColIndex("CusName")) = ""
                End If
  
        End Select
              If Row = .Rows - 1 Then
            .Rows = .Rows + 1
        End If
    End With
End Sub

Private Sub VSFlexGrid1_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
With VSFlexGrid1
Select Case .ColKey(Col)
Case "PriceWithoutPart"
.ComboList = ""
Case "PriceComplete"
.ComboList = ""
Case "Fullcode"
.ComboList = ""
End Select
End With
End Sub

Private Sub VSFlexGrid1_KeyUp(KeyCode As Integer, Shift As Integer)
With VSFlexGrid1
Select Case .ColKey(.Col)
Case "Fullcode", "CusName"
           FrmCompanySearch.lblSearchtype.Caption = 11
           FrmCompanySearch.show vbModal
End Select
End With
           
End Sub



Private Sub VSFlexGrid1_StartEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
   Dim StrAccountCode As String
    Dim Msg As String
    Dim StrSQL As String
    Dim MyStrList As String
    Dim ClsAcc As New ClsAccounts
    Dim LngRow As Long
    Dim rs As ADODB.Recordset
    Set rs = New ADODB.Recordset
    With VSFlexGrid1
        Select Case .ColKey(Col)
         Case "CusName"
                StrSQL = "SELECT     CusID, CusName, CusNamee"
                StrSQL = StrSQL & "  From dbo.TblCustemers"
                StrSQL = StrSQL & "    WHERE     (Type = 2)"
                StrSQL = StrSQL & "  and    (  BranchId=0  or      BranchId in(" & Current_branchSql & "))"
                Set rs = New ADODB.Recordset
                rs.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText
                If Not (rs.BOF Or rs.EOF) Then
                If SystemOptions.UserInterface = ArabicInterface Then
                 MyStrList = .BuildComboList(rs, "CusName", "CusID")
                Else
                MyStrList = .BuildComboList(rs, "CusNamee", "CusID")
                End If
                .ColComboList(.ColIndex("CusName")) = "|" & MyStrList
                End If
                
        End Select

    End With
End Sub
