VERSION 5.00
Object = "{C115893A-A3BF-43AF-B28D-69DB846077F3}#1.0#0"; "vsflex8u.ocx"
Object = "{0AFE7BE0-11B7-4A3E-978D-D4501E9A57FE}#1.0#0"; "c1sizer.ocx"
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{FE5DCFAD-BC1D-11D2-94CF-004005455FAA}#1.4#0"; "ImpulseButton.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form FrmVizitScreen 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   " œ—Ì» «·⁄„·«¡"
   ClientHeight    =   9255
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   16515
   Icon            =   "FrmVizitScreen.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   RightToLeft     =   -1  'True
   ScaleHeight     =   9255
   ScaleWidth      =   16515
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
   Begin C1SizerLibCtl.C1Tab TabMain 
      Height          =   9165
      Left            =   60
      TabIndex        =   0
      Top             =   60
      Width           =   16260
      _cx             =   28681
      _cy             =   16166
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
      Caption         =   " œ—Ì» «·⁄„·«¡|«ÃÊ— «·Ìœ|«·„ﬂ« » «·„›Ê÷…| ⁄—Ì› «·⁄œ”« "
      Align           =   0
      CurrTab         =   3
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
         Height          =   8790
         Index           =   1
         Left            =   -17415
         TabIndex        =   1
         TabStop         =   0   'False
         Top             =   45
         Width           =   16170
         _cx             =   28522
         _cy             =   15505
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
            TabIndex        =   32
            Top             =   0
            Width           =   14550
            Begin VB.TextBox TxtVac_ID 
               Alignment       =   1  'Right Justify
               Appearance      =   0  'Flat
               BackColor       =   &H000000FF&
               Height          =   240
               Left            =   3030
               RightToLeft     =   -1  'True
               TabIndex        =   37
               Top             =   510
               Visible         =   0   'False
               Width           =   945
            End
            Begin VB.TextBox TxtModFlg 
               Alignment       =   1  'Right Justify
               Appearance      =   0  'Flat
               BackColor       =   &H0000FF00&
               Enabled         =   0   'False
               Height          =   285
               Left            =   2580
               RightToLeft     =   -1  'True
               TabIndex        =   36
               Text            =   "modflag"
               Top             =   90
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
               TabIndex        =   33
               Top             =   450
               Visible         =   0   'False
               Width           =   3105
               Begin MSDataListLib.DataCombo DCUser 
                  CausesValidation=   0   'False
                  Height          =   315
                  Left            =   -255
                  TabIndex        =   34
                  Tag             =   "„‰ ›÷·ﬂ √œŒ· —ﬁ„ «·ﬁ÷Ì…"
                  Top             =   15
                  Visible         =   0   'False
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
                  TabIndex        =   35
                  Top             =   45
                  Visible         =   0   'False
                  Width           =   855
               End
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
                     Picture         =   "FrmVizitScreen.frx":57E2
                     Key             =   "CompanyName"
                  EndProperty
                  BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                     Picture         =   "FrmVizitScreen.frx":5B7C
                     Key             =   "Ser"
                  EndProperty
                  BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                     Picture         =   "FrmVizitScreen.frx":5F16
                     Key             =   "Vac_Name"
                  EndProperty
                  BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                     Picture         =   "FrmVizitScreen.frx":62B0
                     Key             =   "ShareCount"
                  EndProperty
                  BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                     Picture         =   "FrmVizitScreen.frx":664A
                     Key             =   "Dis_Count"
                  EndProperty
                  BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                     Picture         =   "FrmVizitScreen.frx":69E4
                     Key             =   "Bouns"
                  EndProperty
                  BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                     Picture         =   "FrmVizitScreen.frx":6D7E
                     Key             =   "SharesValue"
                  EndProperty
                  BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                     Picture         =   "FrmVizitScreen.frx":7318
                     Key             =   "BuyValue"
                  EndProperty
               EndProperty
            End
            Begin ImpulseButton.ISButton btnLast 
               Height          =   315
               Left            =   90
               TabIndex        =   38
               Top             =   150
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
               ButtonImage     =   "FrmVizitScreen.frx":76B2
               ColorButton     =   14871017
               AcclimateGrayTones=   -1  'True
               DrawFocusRectangle=   0   'False
               DisabledImageExtraction=   0
               DisabledImageStyle=   1
            End
            Begin ImpulseButton.ISButton btnNext 
               Height          =   315
               Left            =   555
               TabIndex        =   39
               Top             =   150
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
               ButtonImage     =   "FrmVizitScreen.frx":7A4C
               ColorButton     =   14871017
               DrawFocusRectangle=   0   'False
               DisabledImageStyle=   1
            End
            Begin ImpulseButton.ISButton btnPrevious 
               Height          =   315
               Left            =   1155
               TabIndex        =   40
               Top             =   150
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
               ButtonImage     =   "FrmVizitScreen.frx":7DE6
               ColorButton     =   14871017
               DrawFocusRectangle=   0   'False
               DisabledImageStyle=   1
            End
            Begin ImpulseButton.ISButton btnFirst 
               Height          =   315
               Left            =   1620
               TabIndex        =   41
               Top             =   150
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
               ButtonImage     =   "FrmVizitScreen.frx":8180
               ColorButton     =   14871017
               DrawFocusRectangle=   0   'False
               DisabledImageStyle=   1
            End
            Begin VB.Label Label1 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   " œ—Ì» «·⁄„·«¡"
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
               Left            =   9855
               RightToLeft     =   -1  'True
               TabIndex        =   42
               Top             =   210
               Width           =   3390
            End
         End
         Begin VB.Frame Frm2 
            BackColor       =   &H00E2E9E9&
            BorderStyle     =   0  'None
            Height          =   3510
            Left            =   1425
            RightToLeft     =   -1  'True
            TabIndex        =   5
            Top             =   4890
            Width           =   14550
            Begin VB.ComboBox CmbType 
               BackColor       =   &H80000018&
               Height          =   315
               ItemData        =   "FrmVizitScreen.frx":851A
               Left            =   2280
               List            =   "FrmVizitScreen.frx":852A
               RightToLeft     =   -1  'True
               Style           =   2  'Dropdown List
               TabIndex        =   31
               Top             =   4110
               Visible         =   0   'False
               Width           =   1005
            End
            Begin VB.Frame Frame1 
               BackColor       =   &H00E2E9E9&
               Caption         =   "»Ì«‰«  «·⁄„Ì·"
               Height          =   735
               Index           =   0
               Left            =   6840
               RightToLeft     =   -1  'True
               TabIndex        =   26
               Top             =   0
               Width           =   6615
               Begin VB.TextBox TxtUserPass 
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
                  Left            =   120
                  MaxLength       =   50
                  PasswordChar    =   "*"
                  RightToLeft     =   -1  'True
                  TabIndex        =   27
                  Top             =   240
                  Width           =   1665
               End
               Begin MSDataListLib.DataCombo DcbUserID 
                  Height          =   315
                  Left            =   2760
                  TabIndex        =   28
                  Tag             =   "«Œ — «·œÊ·… „‰ ›÷·ﬂ"
                  Top             =   240
                  Width           =   3090
                  _ExtentX        =   5450
                  _ExtentY        =   556
                  _Version        =   393216
                  Enabled         =   0   'False
                  Text            =   ""
                  RightToLeft     =   -1  'True
               End
               Begin VB.Label Label1 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "«·«”„"
                  Height          =   285
                  Index           =   0
                  Left            =   5640
                  RightToLeft     =   -1  'True
                  TabIndex        =   30
                  Top             =   240
                  Width           =   810
               End
               Begin VB.Label Label1 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "»«”Ê—œ"
                  Height          =   285
                  Index           =   4
                  Left            =   1800
                  RightToLeft     =   -1  'True
                  TabIndex        =   29
                  Top             =   240
                  Width           =   810
               End
            End
            Begin VB.Frame Frame2 
               BackColor       =   &H00E2E9E9&
               Caption         =   "»Ì«‰«  «·„‰œÊ»"
               Height          =   735
               Index           =   1
               Left            =   0
               RightToLeft     =   -1  'True
               TabIndex        =   21
               Top             =   0
               Width           =   6615
               Begin VB.TextBox TxtEmpPass 
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
                  Left            =   120
                  MaxLength       =   50
                  PasswordChar    =   "*"
                  RightToLeft     =   -1  'True
                  TabIndex        =   22
                  Top             =   240
                  Width           =   1665
               End
               Begin MSDataListLib.DataCombo DcbEmpUsrID 
                  Height          =   315
                  Left            =   2760
                  TabIndex        =   23
                  Tag             =   "«Œ — «·œÊ·… „‰ ›÷·ﬂ"
                  Top             =   240
                  Width           =   3090
                  _ExtentX        =   5450
                  _ExtentY        =   556
                  _Version        =   393216
                  Enabled         =   0   'False
                  Text            =   ""
                  RightToLeft     =   -1  'True
               End
               Begin VB.Label Label1 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "»«”Ê—œ"
                  Height          =   285
                  Index           =   1
                  Left            =   1800
                  RightToLeft     =   -1  'True
                  TabIndex        =   25
                  Top             =   240
                  Width           =   810
               End
               Begin VB.Label Label1 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "«·«”„"
                  Height          =   285
                  Index           =   5
                  Left            =   5640
                  RightToLeft     =   -1  'True
                  TabIndex        =   24
                  Top             =   240
                  Width           =   810
               End
            End
            Begin VB.Frame Frame3 
               BackColor       =   &H00E2E9E9&
               Caption         =   "»Ì«‰«  «·“Ì«—…"
               Height          =   2655
               Left            =   2400
               RightToLeft     =   -1  'True
               TabIndex        =   10
               Top             =   840
               Width           =   10935
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
                  Left            =   7800
                  MaxLength       =   50
                  RightToLeft     =   -1  'True
                  TabIndex        =   13
                  Top             =   240
                  Width           =   1545
               End
               Begin VB.TextBox TxtCusRemark 
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
                  Height          =   795
                  Left            =   0
                  MultiLine       =   -1  'True
                  RightToLeft     =   -1  'True
                  ScrollBars      =   2  'Vertical
                  TabIndex        =   12
                  Top             =   600
                  Width           =   9330
               End
               Begin VB.TextBox TxtEmpRemark 
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
                  Height          =   915
                  Left            =   0
                  MaxLength       =   50
                  MultiLine       =   -1  'True
                  RightToLeft     =   -1  'True
                  ScrollBars      =   2  'Vertical
                  TabIndex        =   11
                  Top             =   1560
                  Width           =   9330
               End
               Begin MSComCtl2.DTPicker RecordDate 
                  Height          =   315
                  Left            =   5340
                  TabIndex        =   14
                  Top             =   240
                  Width           =   1545
                  _ExtentX        =   2725
                  _ExtentY        =   556
                  _Version        =   393216
                  Format          =   151126017
                  CurrentDate     =   38784
               End
               Begin MSDataListLib.DataCombo DcbScreen 
                  Height          =   315
                  Left            =   120
                  TabIndex        =   15
                  Tag             =   "«Œ — «·œÊ·… „‰ ›÷·ﬂ"
                  Top             =   240
                  Width           =   3810
                  _ExtentX        =   6720
                  _ExtentY        =   556
                  _Version        =   393216
                  Text            =   ""
                  RightToLeft     =   -1  'True
               End
               Begin VB.Label Label1 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "„"
                  Height          =   195
                  Index           =   3
                  Left            =   10065
                  RightToLeft     =   -1  'True
                  TabIndex        =   20
                  Top             =   240
                  Width           =   510
               End
               Begin VB.Label RecDate 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   " «—ÌŒ "
                  Height          =   285
                  Index           =   11
                  Left            =   7020
                  RightToLeft     =   -1  'True
                  TabIndex        =   19
                  Top             =   240
                  Width           =   690
               End
               Begin VB.Label Label1 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "„·«ÕŸ«  «·⁄„Ì·"
                  Height          =   285
                  Index           =   7
                  Left            =   9600
                  RightToLeft     =   -1  'True
                  TabIndex        =   18
                  Top             =   720
                  Width           =   1170
               End
               Begin VB.Label Label1 
                  Alignment       =   2  'Center
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "„·«ÕŸ«  «·„‰œÊ»"
                  Height          =   285
                  Index           =   6
                  Left            =   9720
                  RightToLeft     =   -1  'True
                  TabIndex        =   17
                  Top             =   1560
                  Width           =   1170
               End
               Begin VB.Label Label1 
                  Alignment       =   2  'Center
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "„« „"
                  Height          =   285
                  Index           =   8
                  Left            =   4065
                  RightToLeft     =   -1  'True
                  TabIndex        =   16
                  Top             =   240
                  Width           =   1170
               End
            End
            Begin VB.Frame Frame4 
               BackColor       =   &H00E2E9E9&
               Height          =   1335
               Left            =   240
               RightToLeft     =   -1  'True
               TabIndex        =   6
               Top             =   840
               Width           =   2055
               Begin ImpulseButton.ISButton CmdAdd 
                  Height          =   360
                  Left            =   1080
                  TabIndex        =   7
                  Top             =   120
                  Width           =   780
                  _ExtentX        =   1376
                  _ExtentY        =   635
                  ButtonStyle     =   1
                  ButtonPositionImage=   1
                  Caption         =   "≈÷«›…"
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
                  ButtonImage     =   "FrmVizitScreen.frx":8543
                  ColorButton     =   14871017
                  DrawFocusRectangle=   0   'False
               End
               Begin ImpulseButton.ISButton CmdDel 
                  Height          =   375
                  Left            =   1080
                  TabIndex        =   8
                  Top             =   915
                  Width           =   780
                  _ExtentX        =   1376
                  _ExtentY        =   661
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
                  ButtonImage     =   "FrmVizitScreen.frx":88DD
                  ColorButton     =   14871017
                  DrawFocusRectangle=   0   'False
               End
               Begin ImpulseButton.ISButton CmdMod 
                  Height          =   270
                  Left            =   1020
                  TabIndex        =   9
                  Top             =   570
                  Width           =   840
                  _ExtentX        =   1482
                  _ExtentY        =   476
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
                  ButtonImage     =   "FrmVizitScreen.frx":F13F
                  ColorButton     =   14871017
                  DrawFocusRectangle=   0   'False
               End
            End
         End
         Begin VSFlex8UCtl.VSFlexGrid FgItems 
            Height          =   8670
            Index           =   0
            Left            =   22500
            TabIndex        =   2
            Top             =   765
            Width           =   15975
            _cx             =   28178
            _cy             =   15293
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
            FormatString    =   $"FrmVizitScreen.frx":F4D9
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
         Begin C1SizerLibCtl.C1Elastic EltCont 
            Height          =   1260
            Left            =   0
            TabIndex        =   43
            TabStop         =   0   'False
            Top             =   7650
            Width           =   14505
            _cx             =   25585
            _cy             =   2223
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
               Left            =   11400
               TabIndex        =   44
               Top             =   675
               Visible         =   0   'False
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
               ButtonImage     =   "FrmVizitScreen.frx":F599
               ColorButton     =   14871017
               DrawFocusRectangle=   0   'False
            End
            Begin ImpulseButton.ISButton btnSave 
               Height          =   330
               Left            =   9030
               TabIndex        =   45
               Top             =   675
               Visible         =   0   'False
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
               ButtonImage     =   "FrmVizitScreen.frx":F933
               ColorButton     =   14871017
               DrawFocusRectangle=   0   'False
            End
            Begin ImpulseButton.ISButton btnModify 
               Height          =   330
               Left            =   10155
               TabIndex        =   46
               Top             =   675
               Visible         =   0   'False
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
               ButtonImage     =   "FrmVizitScreen.frx":FCCD
               ColorButton     =   14871017
               DrawFocusRectangle=   0   'False
            End
            Begin ImpulseButton.ISButton BtnUndo 
               Height          =   330
               Left            =   7665
               TabIndex        =   47
               Top             =   675
               Visible         =   0   'False
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
               ButtonImage     =   "FrmVizitScreen.frx":10067
               ColorButton     =   14871017
               DrawFocusRectangle=   0   'False
            End
            Begin ImpulseButton.ISButton btnDelete 
               Height          =   330
               Left            =   6420
               TabIndex        =   48
               Top             =   675
               Visible         =   0   'False
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
               ButtonImage     =   "FrmVizitScreen.frx":10401
               ColorButton     =   14871017
               DrawFocusRectangle=   0   'False
            End
            Begin ImpulseButton.ISButton btnQuery 
               Height          =   330
               Left            =   5760
               TabIndex        =   49
               TabStop         =   0   'False
               ToolTipText     =   "(Ctrl+F)  ··»ÕÀ ≈÷€ÿ Â–« «·„› «Õ √Ê ≈÷€ÿ "
               Top             =   210
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
               ButtonImage     =   "FrmVizitScreen.frx":1099B
               ColorButton     =   14737632
               DrawFocusRectangle=   0   'False
            End
            Begin ImpulseButton.ISButton BtnUpdate 
               Height          =   330
               Left            =   7485
               TabIndex        =   50
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
               ButtonImage     =   "FrmVizitScreen.frx":10D35
               ColorButton     =   14871017
               DrawFocusRectangle=   0   'False
            End
            Begin ImpulseButton.ISButton btnCancel 
               Height          =   330
               Left            =   1665
               TabIndex        =   51
               Top             =   675
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
               ButtonImage     =   "FrmVizitScreen.frx":110CF
               ColorButton     =   14871017
               DrawFocusRectangle=   0   'False
               DisabledImageStyle=   1
            End
            Begin ImpulseButton.ISButton ISButton2 
               Height          =   375
               Index           =   0
               Left            =   5040
               TabIndex        =   52
               TabStop         =   0   'False
               Top             =   720
               Width           =   1215
               _ExtentX        =   2143
               _ExtentY        =   661
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
               ButtonImage     =   "FrmVizitScreen.frx":11469
               ColorButton     =   14871017
               DrawFocusRectangle=   0   'False
               DisabledImageStyle=   1
            End
            Begin ImpulseButton.ISButton ISButton2 
               Height          =   375
               Index           =   1
               Left            =   3480
               TabIndex        =   53
               TabStop         =   0   'False
               Top             =   720
               Width           =   1215
               _ExtentX        =   2143
               _ExtentY        =   661
               ButtonStyle     =   1
               ButtonPositionImage=   1
               Caption         =   "ÿ»«⁄… «·ﬂ·"
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
               ButtonImage     =   "FrmVizitScreen.frx":17CCB
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
               Left            =   2505
               RightToLeft     =   -1  'True
               TabIndex        =   57
               Top             =   225
               Width           =   975
            End
            Begin VB.Label Label2 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "⁄œœ «·”Ã·« :"
               Height          =   210
               Index           =   1
               Left            =   810
               RightToLeft     =   -1  'True
               TabIndex        =   56
               Top             =   225
               Width           =   975
            End
            Begin VB.Label LabCurrRec 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Height          =   210
               Left            =   1800
               RightToLeft     =   -1  'True
               TabIndex        =   55
               Top             =   240
               Width           =   675
            End
            Begin VB.Label LabCountRec 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Height          =   210
               Left            =   240
               RightToLeft     =   -1  'True
               TabIndex        =   54
               Top             =   225
               Width           =   540
            End
         End
         Begin VSFlex8UCtl.VSFlexGrid Grid 
            Height          =   3405
            Left            =   150
            TabIndex        =   58
            Top             =   750
            Width           =   15810
            _cx             =   27887
            _cy             =   6006
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
            Cols            =   7
            FixedRows       =   1
            FixedCols       =   1
            RowHeightMin    =   320
            RowHeightMax    =   0
            ColWidthMin     =   0
            ColWidthMax     =   0
            ExtendLastCol   =   -1  'True
            FormatString    =   $"FrmVizitScreen.frx":1E52D
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
      Begin C1SizerLibCtl.C1Elastic ELe 
         Height          =   8790
         Index           =   0
         Left            =   -17115
         TabIndex        =   3
         TabStop         =   0   'False
         Top             =   45
         Width           =   16170
         _cx             =   28522
         _cy             =   15505
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
         Begin VB.TextBox TXTOrDer_no2 
            Alignment       =   1  'Right Justify
            Height          =   315
            Left            =   6315
            TabIndex        =   182
            Top             =   900
            Width           =   1395
         End
         Begin VB.TextBox TxtNoteSerial1 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFC0&
            Height          =   315
            Left            =   13335
            Locked          =   -1  'True
            RightToLeft     =   -1  'True
            TabIndex        =   129
            Top             =   990
            Width           =   1545
         End
         Begin VB.CommandButton cmdPrintNote 
            Caption         =   "ÿ»«⁄Â «·ﬁÌœ"
            Height          =   450
            Left            =   3405
            RightToLeft     =   -1  'True
            TabIndex        =   127
            Top             =   6960
            Width           =   2445
         End
         Begin VB.CommandButton cmdDelNote 
            Caption         =   "Õ–› «·ﬁÌœ "
            Height          =   450
            Left            =   10365
            RightToLeft     =   -1  'True
            TabIndex        =   126
            Top             =   6960
            Visible         =   0   'False
            Width           =   2910
         End
         Begin VB.CommandButton CmdCreateV2 
            Caption         =   "≈‰‘«¡ «·ﬁÌœ "
            Height          =   450
            Left            =   13335
            RightToLeft     =   -1  'True
            TabIndex        =   125
            Top             =   6960
            Visible         =   0   'False
            Width           =   2595
         End
         Begin VB.TextBox TxtNoteSerial 
            Alignment       =   1  'Right Justify
            Enabled         =   0   'False
            Height          =   450
            Left            =   5850
            Locked          =   -1  'True
            RightToLeft     =   -1  'True
            TabIndex        =   124
            Top             =   6960
            Width           =   3210
         End
         Begin VB.TextBox Text1 
            Alignment       =   1  'Right Justify
            Height          =   315
            Left            =   30
            RightToLeft     =   -1  'True
            TabIndex        =   123
            Top             =   7065
            Visible         =   0   'False
            Width           =   2190
         End
         Begin VB.TextBox TxtCashCustomerName 
            Alignment       =   1  'Right Justify
            Height          =   345
            Left            =   2160
            Locked          =   -1  'True
            TabIndex        =   121
            Top             =   1350
            Width           =   2880
         End
         Begin VB.TextBox TXTOrDer_no 
            Alignment       =   1  'Right Justify
            Height          =   315
            Index           =   1
            Left            =   5520
            TabIndex        =   120
            Top             =   540
            Visible         =   0   'False
            Width           =   1410
         End
         Begin VB.ComboBox DcbType 
            Height          =   315
            Left            =   10365
            RightToLeft     =   -1  'True
            Style           =   2  'Dropdown List
            TabIndex        =   108
            Top             =   8760
            Visible         =   0   'False
            Width           =   1320
         End
         Begin VB.ComboBox DCOPrType 
            Height          =   315
            Left            =   12510
            RightToLeft     =   -1  'True
            Style           =   2  'Dropdown List
            TabIndex        =   107
            Top             =   8790
            Visible         =   0   'False
            Width           =   1965
         End
         Begin VB.ComboBox DcbyearFactor 
            Enabled         =   0   'False
            Height          =   315
            Left            =   12495
            Locked          =   -1  'True
            RightToLeft     =   -1  'True
            TabIndex        =   106
            Top             =   1890
            Width           =   2385
         End
         Begin VB.TextBox TxtPlatNo 
            Alignment       =   1  'Right Justify
            Height          =   315
            Left            =   615
            Locked          =   -1  'True
            TabIndex        =   105
            Top             =   1890
            Width           =   2430
         End
         Begin VB.TextBox TxtManualNo2 
            Alignment       =   1  'Right Justify
            BackColor       =   &H0080FFFF&
            Height          =   285
            Index           =   2
            Left            =   8085
            Locked          =   -1  'True
            TabIndex        =   104
            Top             =   1920
            Width           =   2310
         End
         Begin VB.TextBox TxtManualNo2 
            Alignment       =   1  'Right Justify
            BackColor       =   &H0080FFFF&
            Height          =   285
            Index           =   1
            Left            =   4635
            Locked          =   -1  'True
            TabIndex        =   103
            Top             =   1920
            Width           =   1680
         End
         Begin VB.TextBox TXTOrDer_no 
            Alignment       =   1  'Right Justify
            Height          =   315
            Index           =   0
            Left            =   3135
            TabIndex        =   101
            Top             =   930
            Width           =   1425
         End
         Begin VB.ComboBox CBoBasedON 
            Height          =   315
            ItemData        =   "FrmVizitScreen.frx":1E637
            Left            =   7755
            List            =   "FrmVizitScreen.frx":1E639
            RightToLeft     =   -1  'True
            Style           =   2  'Dropdown List
            TabIndex        =   100
            Top             =   915
            Width           =   1560
         End
         Begin VB.TextBox TxtNoteID 
            Alignment       =   1  'Right Justify
            Height          =   315
            Left            =   -1155
            RightToLeft     =   -1  'True
            TabIndex        =   82
            Top             =   7140
            Visible         =   0   'False
            Width           =   2235
         End
         Begin VB.TextBox TxtSerial1 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFFC0&
            Enabled         =   0   'False
            Height          =   360
            Index           =   1
            Left            =   -30
            RightToLeft     =   -1  'True
            TabIndex        =   73
            Top             =   1290
            Visible         =   0   'False
            Width           =   2025
         End
         Begin VB.TextBox TxtRemarks 
            Alignment       =   2  'Center
            Height          =   555
            Left            =   90
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   72
            Top             =   2670
            Width           =   14220
         End
         Begin VB.Frame Fra_Header 
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   0  'None
            Height          =   570
            Index           =   0
            Left            =   0
            RightToLeft     =   -1  'True
            TabIndex        =   60
            Top             =   0
            Width           =   17985
            Begin VB.Frame Frame2 
               BackColor       =   &H00E0E0E0&
               BorderStyle     =   0  'None
               Enabled         =   0   'False
               Height          =   375
               Index           =   0
               Left            =   540
               RightToLeft     =   -1  'True
               TabIndex        =   63
               Top             =   450
               Visible         =   0   'False
               Width           =   3105
               Begin MSDataListLib.DataCombo DataCombo4 
                  CausesValidation=   0   'False
                  Height          =   315
                  Index           =   0
                  Left            =   -255
                  TabIndex        =   64
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
                  Index           =   9
                  Left            =   2160
                  RightToLeft     =   -1  'True
                  TabIndex        =   65
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
               Index           =   1
               Left            =   2580
               RightToLeft     =   -1  'True
               TabIndex        =   62
               Text            =   "modflag"
               Top             =   90
               Visible         =   0   'False
               Width           =   465
            End
            Begin VB.TextBox Text2 
               Alignment       =   1  'Right Justify
               Appearance      =   0  'Flat
               BackColor       =   &H000000FF&
               Height          =   240
               Index           =   0
               Left            =   3030
               RightToLeft     =   -1  'True
               TabIndex        =   61
               Top             =   510
               Visible         =   0   'False
               Width           =   945
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
                     Picture         =   "FrmVizitScreen.frx":1E63B
                     Key             =   "CompanyName"
                  EndProperty
                  BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                     Picture         =   "FrmVizitScreen.frx":1E9D5
                     Key             =   "Ser"
                  EndProperty
                  BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                     Picture         =   "FrmVizitScreen.frx":1ED6F
                     Key             =   "Vac_Name"
                  EndProperty
                  BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                     Picture         =   "FrmVizitScreen.frx":1F109
                     Key             =   "ShareCount"
                  EndProperty
                  BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                     Picture         =   "FrmVizitScreen.frx":1F4A3
                     Key             =   "Dis_Count"
                  EndProperty
                  BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                     Picture         =   "FrmVizitScreen.frx":1F83D
                     Key             =   "Bouns"
                  EndProperty
                  BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                     Picture         =   "FrmVizitScreen.frx":1FBD7
                     Key             =   "SharesValue"
                  EndProperty
                  BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                     Picture         =   "FrmVizitScreen.frx":20171
                     Key             =   "BuyValue"
                  EndProperty
               EndProperty
            End
            Begin ImpulseButton.ISButton btn_Last 
               Height          =   315
               Index           =   1
               Left            =   90
               TabIndex        =   66
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
               ButtonImage     =   "FrmVizitScreen.frx":2050B
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
               TabIndex        =   67
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
               ButtonImage     =   "FrmVizitScreen.frx":208A5
               ColorButton     =   14871017
               DrawFocusRectangle=   0   'False
               DisabledImageStyle=   1
            End
            Begin ImpulseButton.ISButton btn_Previous 
               Height          =   315
               Index           =   1
               Left            =   1155
               TabIndex        =   68
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
               ButtonImage     =   "FrmVizitScreen.frx":20C3F
               ColorButton     =   14871017
               DrawFocusRectangle=   0   'False
               DisabledImageStyle=   1
            End
            Begin ImpulseButton.ISButton btn_First 
               Height          =   315
               Index           =   1
               Left            =   1620
               TabIndex        =   69
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
               ButtonImage     =   "FrmVizitScreen.frx":20FD9
               ColorButton     =   14871017
               DrawFocusRectangle=   0   'False
               DisabledImageStyle=   1
            End
            Begin VB.Image ImgFavorites 
               Height          =   390
               Left            =   7560
               Picture         =   "FrmVizitScreen.frx":21373
               Stretch         =   -1  'True
               Top             =   0
               Width           =   525
            End
            Begin VB.Label Label1 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "«ÃÊ— «·Ìœ"
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
               Index           =   10
               Left            =   11340
               RightToLeft     =   -1  'True
               TabIndex        =   70
               Top             =   90
               Width           =   2640
            End
         End
         Begin VSFlex8UCtl.VSFlexGrid FgItems 
            Height          =   8670
            Index           =   1
            Left            =   22545
            TabIndex        =   4
            Top             =   765
            Width           =   15945
            _cx             =   28125
            _cy             =   15293
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
            FormatString    =   $"FrmVizitScreen.frx":24FDB
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
         Begin MSComCtl2.DTPicker XPDtbTrans 
            Height          =   300
            Left            =   10200
            TabIndex        =   74
            Top             =   960
            Width           =   1620
            _ExtentX        =   2858
            _ExtentY        =   529
            _Version        =   393216
            Format          =   44564481
            CurrentDate     =   38784
         End
         Begin MSDataListLib.DataCombo Dcbranch 
            Bindings        =   "FrmVizitScreen.frx":2509B
            Height          =   315
            Index           =   1
            Left            =   0
            TabIndex        =   75
            Top             =   840
            Width           =   2235
            _ExtentX        =   3942
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
         Begin MSDataListLib.DataCombo DcCustmer 
            Height          =   315
            Index           =   1
            Left            =   6765
            TabIndex        =   76
            Top             =   1380
            Width           =   3630
            _ExtentX        =   6403
            _ExtentY        =   556
            _Version        =   393216
            Locked          =   -1  'True
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
         Begin MSDataListLib.DataCombo DCboUserName 
            Height          =   315
            Index           =   1
            Left            =   11850
            TabIndex        =   83
            Top             =   7740
            Width           =   2880
            _ExtentX        =   5080
            _ExtentY        =   556
            _Version        =   393216
            Enabled         =   0   'False
            Text            =   ""
            RightToLeft     =   -1  'True
         End
         Begin ImpulseButton.ISButton btn_New 
            Height          =   285
            Index           =   1
            Left            =   12255
            TabIndex        =   84
            Top             =   8430
            Width           =   870
            _ExtentX        =   1535
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
            ButtonImage     =   "FrmVizitScreen.frx":250B0
            ColorButton     =   14871017
            DrawFocusRectangle=   0   'False
         End
         Begin ImpulseButton.ISButton btn_Save 
            Height          =   315
            Index           =   1
            Left            =   10395
            TabIndex        =   85
            Top             =   8400
            Width           =   780
            _ExtentX        =   1376
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
            ButtonImage     =   "FrmVizitScreen.frx":2544A
            ColorButton     =   14871017
            DrawFocusRectangle=   0   'False
         End
         Begin ImpulseButton.ISButton btn_Modify 
            Height          =   225
            Index           =   1
            Left            =   11205
            TabIndex        =   86
            Top             =   8430
            Width           =   840
            _ExtentX        =   1482
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
            ButtonImage     =   "FrmVizitScreen.frx":257E4
            ColorButton     =   14871017
            DrawFocusRectangle=   0   'False
         End
         Begin ImpulseButton.ISButton Btn_Undo 
            Height          =   225
            Index           =   1
            Left            =   9525
            TabIndex        =   87
            Top             =   8430
            Width           =   870
            _ExtentX        =   1535
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
            ButtonImage     =   "FrmVizitScreen.frx":25B7E
            ColorButton     =   14871017
            DrawFocusRectangle=   0   'False
         End
         Begin ImpulseButton.ISButton btn_Delete 
            Height          =   315
            Index           =   1
            Left            =   8760
            TabIndex        =   88
            Top             =   8400
            Width           =   765
            _ExtentX        =   1349
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
            ButtonImage     =   "FrmVizitScreen.frx":25F18
            ColorButton     =   14871017
            DrawFocusRectangle=   0   'False
         End
         Begin ImpulseButton.ISButton Btn_Update 
            Height          =   345
            Index           =   1
            Left            =   9735
            TabIndex        =   89
            TabStop         =   0   'False
            ToolTipText     =   " ÕœÌÀ ﬁ«⁄œ… «·»Ì«‰« "
            Top             =   8070
            Visible         =   0   'False
            Width           =   960
            _ExtentX        =   1693
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
            ButtonImage     =   "FrmVizitScreen.frx":264B2
            ColorButton     =   14871017
            DrawFocusRectangle=   0   'False
         End
         Begin ImpulseButton.ISButton btn_Cancel 
            Height          =   315
            Index           =   1
            Left            =   5085
            TabIndex        =   90
            Top             =   8370
            Width           =   885
            _ExtentX        =   1561
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
            ButtonImage     =   "FrmVizitScreen.frx":2684C
            ColorButton     =   14871017
            DrawFocusRectangle=   0   'False
            DisabledImageStyle=   1
         End
         Begin ImpulseButton.ISButton Btn_Print 
            Height          =   360
            Index           =   1
            Left            =   7410
            TabIndex        =   91
            TabStop         =   0   'False
            ToolTipText     =   "ÿ»«⁄… «·»Ì«‰«  "
            Top             =   8340
            Width           =   1155
            _ExtentX        =   2037
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
            ButtonImage     =   "FrmVizitScreen.frx":26BE6
            ColorButton     =   14871017
            DrawFocusRectangle=   0   'False
            DisabledImageStyle=   1
         End
         Begin ImpulseButton.ISButton btn_Query 
            Height          =   390
            Index           =   1
            Left            =   6075
            TabIndex        =   92
            TabStop         =   0   'False
            ToolTipText     =   "(Ctrl+F)  ··»ÕÀ ≈÷€ÿ Â–« «·„› «Õ √Ê ≈÷€ÿ "
            Top             =   8310
            Width           =   1140
            _ExtentX        =   2011
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
            ButtonImage     =   "FrmVizitScreen.frx":2D448
            ColorButton     =   14871017
            DrawFocusRectangle=   0   'False
         End
         Begin ImpulseButton.ISButton Cmd_DeleteRow 
            Height          =   300
            Index           =   1
            Left            =   2145
            TabIndex        =   93
            Top             =   7590
            Width           =   1650
            _ExtentX        =   2910
            _ExtentY        =   529
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
            ButtonImage     =   "FrmVizitScreen.frx":2D7E2
            ColorButton     =   14871017
            DrawFocusRectangle=   0   'False
         End
         Begin ImpulseButton.ISButton Cmd_DeleteAll 
            Height          =   300
            Index           =   1
            Left            =   255
            TabIndex        =   94
            Top             =   7575
            Width           =   1755
            _ExtentX        =   3096
            _ExtentY        =   529
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
            ButtonImage     =   "FrmVizitScreen.frx":2DD7C
            ColorButton     =   14871017
            DrawFocusRectangle=   0   'False
         End
         Begin MSDataListLib.DataCombo DCEquipments 
            Height          =   315
            Left            =   10815
            TabIndex        =   109
            Top             =   2220
            Visible         =   0   'False
            Width           =   4065
            _ExtentX        =   7170
            _ExtentY        =   556
            _Version        =   393216
            ListField       =   "6"
            Text            =   ""
            RightToLeft     =   -1  'True
         End
         Begin MSDataListLib.DataCombo DcbCarType 
            Bindings        =   "FrmVizitScreen.frx":2E316
            Height          =   315
            Left            =   12495
            TabIndex        =   110
            Top             =   1500
            Width           =   2385
            _ExtentX        =   4207
            _ExtentY        =   556
            _Version        =   393216
            Enabled         =   0   'False
            Locked          =   -1  'True
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
         Begin MSDataListLib.DataCombo DcbCarModel 
            Bindings        =   "FrmVizitScreen.frx":2E32B
            Height          =   315
            Left            =   585
            TabIndex        =   131
            Top             =   2250
            Width           =   2415
            _ExtentX        =   4260
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
         Begin C1SizerLibCtl.C1Tab TabMain2 
            Height          =   3765
            Left            =   135
            TabIndex        =   134
            Top             =   3210
            Width           =   16065
            _cx             =   28337
            _cy             =   6641
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
            Caption         =   "»Ì«‰« |»Ì«‰«  ›Ê« Ì— «·„»Ì⁄« "
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
            Begin C1SizerLibCtl.C1Elastic ELe 
               Height          =   3390
               Index           =   2
               Left            =   45
               TabIndex        =   135
               TabStop         =   0   'False
               Top             =   45
               Width           =   15975
               _cx             =   28178
               _cy             =   5980
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
               Begin VB.Frame Frame6 
                  Caption         =   "«·«Ã„«·Ï «·⁄«„"
                  Height          =   2145
                  Left            =   0
                  RightToLeft     =   -1  'True
                  TabIndex        =   171
                  Top             =   1170
                  Width           =   3450
                  Begin VB.TextBox txtGeneralTotal 
                     Alignment       =   1  'Right Justify
                     BackColor       =   &H80000010&
                     Height          =   315
                     Left            =   210
                     Locked          =   -1  'True
                     RightToLeft     =   -1  'True
                     TabIndex        =   176
                     Top             =   450
                     Width           =   1440
                  End
                  Begin VB.TextBox txtTotalDisc 
                     Alignment       =   1  'Right Justify
                     BackColor       =   &H80000010&
                     Height          =   285
                     Left            =   210
                     Locked          =   -1  'True
                     RightToLeft     =   -1  'True
                     TabIndex        =   175
                     Top             =   780
                     Width           =   1440
                  End
                  Begin VB.TextBox txtTotalBVat 
                     Alignment       =   1  'Right Justify
                     BackColor       =   &H80000010&
                     Height          =   285
                     Left            =   210
                     Locked          =   -1  'True
                     RightToLeft     =   -1  'True
                     TabIndex        =   174
                     Top             =   1170
                     Width           =   1440
                  End
                  Begin VB.TextBox txtTotalVat 
                     Alignment       =   1  'Right Justify
                     BackColor       =   &H80000010&
                     Height          =   285
                     Left            =   210
                     Locked          =   -1  'True
                     RightToLeft     =   -1  'True
                     TabIndex        =   173
                     Top             =   1530
                     Width           =   1440
                  End
                  Begin VB.TextBox txtTotalNet 
                     Alignment       =   1  'Right Justify
                     BackColor       =   &H80000010&
                     Height          =   285
                     Left            =   210
                     Locked          =   -1  'True
                     RightToLeft     =   -1  'True
                     TabIndex        =   172
                     Top             =   1830
                     Width           =   1440
                  End
                  Begin VB.Label lbl 
                     Alignment       =   1  'Right Justify
                     BackColor       =   &H00E2E9E9&
                     Caption         =   "«·Œ’„"
                     Height          =   225
                     Index           =   23
                     Left            =   1605
                     TabIndex        =   181
                     Top             =   780
                     Width           =   1650
                  End
                  Begin VB.Label lbl 
                     Alignment       =   1  'Right Justify
                     BackColor       =   &H00E2E9E9&
                     Caption         =   "«·«Ã„«·Ì ﬁ»· «·÷—Ì»…"
                     Height          =   225
                     Index           =   17
                     Left            =   1755
                     TabIndex        =   180
                     Top             =   1140
                     Width           =   1500
                  End
                  Begin VB.Label lbl 
                     Alignment       =   1  'Right Justify
                     BackColor       =   &H00E2E9E9&
                     Caption         =   "«·«Ã„«·Ì"
                     Height          =   225
                     Index           =   16
                     Left            =   1950
                     TabIndex        =   179
                     Top             =   480
                     Width           =   1305
                  End
                  Begin VB.Label lbl 
                     Alignment       =   1  'Right Justify
                     BackColor       =   &H00E2E9E9&
                     Caption         =   "«·ﬁÌ„… «·„÷«›…"
                     Height          =   285
                     Index           =   18
                     Left            =   2190
                     RightToLeft     =   -1  'True
                     TabIndex        =   178
                     Top             =   1470
                     Width           =   1065
                  End
                  Begin VB.Label lbl 
                     Alignment       =   1  'Right Justify
                     BackColor       =   &H00E2E9E9&
                     Caption         =   "«·’«›Ï "
                     Height          =   225
                     Index           =   19
                     Left            =   2130
                     TabIndex        =   177
                     Top             =   1890
                     Width           =   1125
                  End
               End
               Begin VB.Frame Frame7 
                  Caption         =   "«Ã„«·Ï «ÃÊ— «·Ìœ"
                  Height          =   1905
                  Left            =   3750
                  RightToLeft     =   -1  'True
                  TabIndex        =   154
                  Top             =   1470
                  Width           =   5805
                  Begin VB.TextBox txtTotal 
                     Alignment       =   1  'Right Justify
                     BackColor       =   &H80000010&
                     Height          =   285
                     Left            =   2370
                     Locked          =   -1  'True
                     RightToLeft     =   -1  'True
                     TabIndex        =   161
                     Top             =   585
                     Width           =   1380
                  End
                  Begin VB.TextBox txtVatYou 
                     Alignment       =   1  'Right Justify
                     BackColor       =   &H80000010&
                     Height          =   285
                     Left            =   2940
                     Locked          =   -1  'True
                     RightToLeft     =   -1  'True
                     TabIndex        =   160
                     Top             =   915
                     Width           =   780
                  End
                  Begin VB.TextBox txtDiscPercent 
                     Alignment       =   1  'Right Justify
                     BackColor       =   &H80000010&
                     Height          =   285
                     Left            =   690
                     RightToLeft     =   -1  'True
                     TabIndex        =   159
                     Top             =   630
                     Width           =   600
                  End
                  Begin VB.TextBox txtDiscValue 
                     Alignment       =   1  'Right Justify
                     BackColor       =   &H80000010&
                     Height          =   285
                     Left            =   90
                     RightToLeft     =   -1  'True
                     TabIndex        =   158
                     Top             =   240
                     Width           =   1200
                  End
                  Begin VB.TextBox txtTotal2 
                     Alignment       =   1  'Right Justify
                     BackColor       =   &H80000010&
                     Height          =   285
                     Left            =   2370
                     Locked          =   -1  'True
                     RightToLeft     =   -1  'True
                     TabIndex        =   157
                     Top             =   240
                     Width           =   1380
                  End
                  Begin VB.TextBox txtVat2 
                     Alignment       =   1  'Right Justify
                     BackColor       =   &H80000010&
                     Height          =   285
                     Left            =   2370
                     Locked          =   -1  'True
                     RightToLeft     =   -1  'True
                     TabIndex        =   156
                     Top             =   1230
                     Width           =   1380
                  End
                  Begin VB.TextBox txtNetInvoice2 
                     Alignment       =   1  'Right Justify
                     BackColor       =   &H80000010&
                     Height          =   285
                     Left            =   2370
                     Locked          =   -1  'True
                     RightToLeft     =   -1  'True
                     TabIndex        =   155
                     Top             =   1530
                     Width           =   1380
                  End
                  Begin VB.Label lbl 
                     Alignment       =   1  'Right Justify
                     BackColor       =   &H00E2E9E9&
                     Caption         =   "«·«Ã„«·Ì ﬁ»· «·÷—Ì»…"
                     Height          =   225
                     Index           =   9
                     Left            =   3840
                     TabIndex        =   170
                     Top             =   585
                     Width           =   1680
                  End
                  Begin VB.Label lbl 
                     Alignment       =   1  'Right Justify
                     BackColor       =   &H00E2E9E9&
                     Caption         =   "«·ﬁÌ„… «·„÷«›…"
                     Height          =   285
                     Index           =   65
                     Left            =   4320
                     RightToLeft     =   -1  'True
                     TabIndex        =   169
                     Top             =   1275
                     Width           =   1140
                  End
                  Begin VB.Label lbl 
                     Alignment       =   2  'Center
                     BackColor       =   &H00E2E9E9&
                     Caption         =   "%"
                     Height          =   225
                     Index           =   6
                     Left            =   2280
                     TabIndex        =   168
                     Top             =   930
                     Width           =   600
                  End
                  Begin VB.Label lbl 
                     Alignment       =   1  'Right Justify
                     BackColor       =   &H00E2E9E9&
                     Caption         =   "«·÷—Ì»…"
                     Height          =   225
                     Index           =   5
                     Left            =   4530
                     TabIndex        =   167
                     Top             =   915
                     Width           =   960
                  End
                  Begin VB.Label lbl 
                     Alignment       =   1  'Right Justify
                     BackColor       =   &H00E2E9E9&
                     BackStyle       =   0  'Transparent
                     Caption         =   "«·Œ’„ ‰”»…"
                     Height          =   225
                     Index           =   1
                     Left            =   1290
                     TabIndex        =   166
                     Top             =   660
                     Width           =   900
                  End
                  Begin VB.Label lbl 
                     Alignment       =   1  'Right Justify
                     BackColor       =   &H00E2E9E9&
                     BackStyle       =   0  'Transparent
                     Caption         =   "«·Œ’„ ﬁÌ„…"
                     Height          =   225
                     Index           =   0
                     Left            =   480
                     TabIndex        =   165
                     Top             =   240
                     Width           =   1680
                  End
                  Begin VB.Label lbl 
                     Alignment       =   1  'Right Justify
                     BackColor       =   &H00E2E9E9&
                     Caption         =   "«Ã„«·Ì "
                     Height          =   225
                     Index           =   43
                     Left            =   3840
                     TabIndex        =   164
                     Top             =   240
                     Width           =   1650
                  End
                  Begin VB.Label lbl 
                     Alignment       =   2  'Center
                     BackColor       =   &H00E2E9E9&
                     Caption         =   "%"
                     Height          =   225
                     Index           =   21
                     Left            =   90
                     TabIndex        =   163
                     Top             =   660
                     Width           =   600
                  End
                  Begin VB.Label lbl 
                     Alignment       =   1  'Right Justify
                     BackColor       =   &H00E2E9E9&
                     Caption         =   "«·’«›Ï "
                     Height          =   285
                     Index           =   22
                     Left            =   4410
                     RightToLeft     =   -1  'True
                     TabIndex        =   162
                     Top             =   1590
                     Width           =   1080
                  End
               End
               Begin VB.TextBox txtNet 
                  Alignment       =   1  'Right Justify
                  Height          =   285
                  Left            =   2235
                  Locked          =   -1  'True
                  RightToLeft     =   -1  'True
                  TabIndex        =   150
                  Top             =   3390
                  Width           =   3270
               End
               Begin VB.Frame Frame5 
                  Caption         =   "»Ì«‰«   ›Ê« Ì— «·„»Ì⁄« "
                  Height          =   1485
                  Left            =   3750
                  RightToLeft     =   -1  'True
                  TabIndex        =   139
                  Top             =   0
                  Width           =   5805
                  Begin VB.TextBox txtTotalInvoice 
                     Alignment       =   1  'Right Justify
                     BackColor       =   &H80000010&
                     Height          =   285
                     Left            =   2580
                     Locked          =   -1  'True
                     RightToLeft     =   -1  'True
                     TabIndex        =   144
                     Top             =   360
                     Width           =   1320
                  End
                  Begin VB.TextBox txtVat2Invoice 
                     Alignment       =   1  'Right Justify
                     BackColor       =   &H80000010&
                     Height          =   315
                     Left            =   2580
                     Locked          =   -1  'True
                     RightToLeft     =   -1  'True
                     TabIndex        =   143
                     Top             =   1020
                     Width           =   1335
                  End
                  Begin VB.TextBox txtDiscValueInvoice 
                     Alignment       =   1  'Right Justify
                     BackColor       =   &H80000010&
                     Height          =   285
                     Left            =   300
                     Locked          =   -1  'True
                     RightToLeft     =   -1  'True
                     TabIndex        =   142
                     Top             =   360
                     Width           =   1020
                  End
                  Begin VB.TextBox txtNetInvoice 
                     Alignment       =   1  'Right Justify
                     BackColor       =   &H80000010&
                     Height          =   285
                     Left            =   300
                     Locked          =   -1  'True
                     RightToLeft     =   -1  'True
                     TabIndex        =   141
                     Top             =   1020
                     Width           =   1020
                  End
                  Begin VB.TextBox txtTotalInvoiceBVat 
                     Alignment       =   1  'Right Justify
                     BackColor       =   &H80000010&
                     Height          =   285
                     Left            =   2580
                     Locked          =   -1  'True
                     RightToLeft     =   -1  'True
                     TabIndex        =   140
                     Top             =   660
                     Width           =   1320
                  End
                  Begin VB.Label lbl 
                     Alignment       =   2  'Center
                     BackColor       =   &H00E2E9E9&
                     Caption         =   "«Ã„«·Ì ﬁÿ⁄ «·€Ì«—"
                     Height          =   225
                     Index           =   10
                     Left            =   4020
                     TabIndex        =   149
                     Top             =   360
                     Width           =   1425
                  End
                  Begin VB.Label lbl 
                     Alignment       =   1  'Right Justify
                     BackColor       =   &H00E2E9E9&
                     Caption         =   "«·ﬁÌ„… «·„÷«›…"
                     Height          =   285
                     Index           =   11
                     Left            =   4170
                     RightToLeft     =   -1  'True
                     TabIndex        =   148
                     Top             =   1035
                     Width           =   1275
                  End
                  Begin VB.Label lbl 
                     Alignment       =   1  'Right Justify
                     BackColor       =   &H00E2E9E9&
                     Caption         =   "«·Œ’„ ﬁÌ„…"
                     Height          =   225
                     Index           =   12
                     Left            =   1470
                     TabIndex        =   147
                     Top             =   390
                     Width           =   885
                  End
                  Begin VB.Label lbl 
                     Alignment       =   1  'Right Justify
                     BackColor       =   &H00E2E9E9&
                     Caption         =   "«·’«›Ï "
                     Height          =   225
                     Index           =   13
                     Left            =   510
                     TabIndex        =   146
                     Top             =   1020
                     Width           =   1845
                  End
                  Begin VB.Label lbl 
                     Alignment       =   2  'Center
                     BackColor       =   &H00E2E9E9&
                     Caption         =   "«·«Ã„«·Ì ﬁ»· «·÷—Ì»…"
                     Height          =   225
                     Index           =   14
                     Left            =   4020
                     TabIndex        =   145
                     Top             =   660
                     Width           =   1425
                  End
               End
               Begin VSFlex8UCtl.VSFlexGrid FgItems 
                  Height          =   3285
                  Index           =   2
                  Left            =   23475
                  TabIndex        =   136
                  Top             =   645
                  Width           =   15900
                  _cx             =   28046
                  _cy             =   5794
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
                  FormatString    =   $"FrmVizitScreen.frx":2E340
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
               Begin VSFlex8UCtl.VSFlexGrid fg 
                  Height          =   3105
                  Left            =   9600
                  TabIndex        =   151
                  Top             =   90
                  Width           =   6180
                  _cx             =   10901
                  _cy             =   5477
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
                  Rows            =   12
                  Cols            =   7
                  FixedRows       =   1
                  FixedCols       =   1
                  RowHeightMin    =   320
                  RowHeightMax    =   0
                  ColWidthMin     =   0
                  ColWidthMax     =   0
                  ExtendLastCol   =   -1  'True
                  FormatString    =   $"FrmVizitScreen.frx":2E400
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
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "«·’«›Ï »⁄œ «·ﬁÌ„… «·„÷«›…"
                  Height          =   225
                  Index           =   47
                  Left            =   5505
                  TabIndex        =   152
                  Top             =   3420
                  Width           =   1920
               End
            End
            Begin C1SizerLibCtl.C1Elastic ELe 
               Height          =   3390
               Index           =   3
               Left            =   16710
               TabIndex        =   137
               TabStop         =   0   'False
               Top             =   45
               Width           =   15975
               _cx             =   28178
               _cy             =   5980
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
                  Height          =   3270
                  Index           =   3
                  Left            =   23340
                  TabIndex        =   138
                  Top             =   720
                  Width           =   15870
                  _cx             =   27993
                  _cy             =   5768
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
                  FormatString    =   $"FrmVizitScreen.frx":2E50F
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
               Begin VSFlex8UCtl.VSFlexGrid grdTrans 
                  Height          =   3300
                  Left            =   375
                  TabIndex        =   153
                  Top             =   -60
                  Width           =   15135
                  _cx             =   26696
                  _cy             =   5821
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
                  Rows            =   12
                  Cols            =   13
                  FixedRows       =   1
                  FixedCols       =   1
                  RowHeightMin    =   320
                  RowHeightMax    =   0
                  ColWidthMin     =   0
                  ColWidthMax     =   0
                  ExtendLastCol   =   -1  'True
                  FormatString    =   $"FrmVizitScreen.frx":2E5CF
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
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "—ﬁ„ «„— «·«’·«Õ"
            Height          =   255
            Index           =   20
            Left            =   4245
            TabIndex        =   133
            Top             =   960
            Width           =   1560
         End
         Begin VB.Label lblModel 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "«·ÿ—«“ "
            Height          =   255
            Left            =   3300
            TabIndex        =   132
            Top             =   2250
            Width           =   915
         End
         Begin VB.Label Label5 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "—ﬁ„ «·”‰œ"
            Height          =   255
            Index           =   1
            Left            =   15030
            TabIndex        =   130
            Top             =   960
            Width           =   900
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            BackStyle       =   0  'Transparent
            Caption         =   "—ﬁ„ «·ﬁÌœ"
            Height          =   405
            Index           =   14
            Left            =   9060
            RightToLeft     =   -1  'True
            TabIndex        =   128
            Top             =   7065
            Width           =   1125
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "«”„ «·⁄„Ì· «·‰ﬁœÏ"
            Height          =   315
            Index           =   33
            Left            =   5250
            TabIndex        =   122
            Top             =   1410
            Width           =   1410
         End
         Begin VB.Label lbl 
            Alignment       =   2  'Center
            BackColor       =   &H00E2E9E9&
            Caption         =   "%"
            Height          =   225
            Index           =   3
            Left            =   3600
            TabIndex        =   119
            Top             =   5190
            Width           =   645
         End
         Begin VB.Label LblYear 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "„ÊœÌ· «·„⁄œÂ/«·”Ì«—…"
            Height          =   255
            Left            =   14775
            TabIndex        =   118
            Top             =   1830
            Width           =   1200
         End
         Begin VB.Label LblPla 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "—ﬁ„ «··ÊÕ…"
            Height          =   255
            Left            =   3330
            TabIndex        =   117
            Top             =   1920
            Width           =   1155
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "«·‰Ê⁄"
            Height          =   285
            Index           =   123
            Left            =   11655
            RightToLeft     =   -1  'True
            TabIndex        =   116
            Top             =   9000
            Visible         =   0   'False
            Width           =   810
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "‰Ê⁄ «·⁄„·Ì…"
            Height          =   285
            Index           =   124
            Left            =   14835
            RightToLeft     =   -1  'True
            TabIndex        =   115
            Top             =   8745
            Visible         =   0   'False
            Width           =   765
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "«·„⁄œÂ/«·”Ì«—…"
            Height          =   240
            Index           =   125
            Left            =   14790
            RightToLeft     =   -1  'True
            TabIndex        =   114
            Top             =   2250
            Visible         =   0   'False
            Width           =   1185
         End
         Begin VB.Label lbltycar 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "‰Ê⁄ «·„⁄œÂ/«·”Ì«—…"
            Height          =   255
            Left            =   15120
            TabIndex        =   113
            Top             =   1470
            Width           =   900
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            BackStyle       =   0  'Transparent
            Caption         =   "—ﬁ„ «·‘«”ÌÂ"
            Height          =   195
            Index           =   119
            Left            =   10125
            TabIndex        =   112
            Top             =   1920
            Width           =   1410
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            BackStyle       =   0  'Transparent
            Caption         =   "⁄œ«œ «·ﬂÌ·Ê „ —"
            Height          =   195
            Index           =   118
            Left            =   6225
            TabIndex        =   111
            Top             =   1920
            Width           =   1395
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "»‰«¡ ⁄·Ï"
            Height          =   255
            Index           =   56
            Left            =   9240
            TabIndex        =   102
            Top             =   945
            Width           =   840
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Õ—— »Ê«”ÿ…  "
            Height          =   315
            Index           =   8
            Left            =   14880
            TabIndex        =   99
            Top             =   7680
            Width           =   990
         End
         Begin VB.Label LabCount_Rec 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Height          =   225
            Index           =   1
            Left            =   2610
            RightToLeft     =   -1  'True
            TabIndex        =   98
            Top             =   8070
            Width           =   585
         End
         Begin VB.Label LabCurr_Rec 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Height          =   225
            Index           =   1
            Left            =   4425
            RightToLeft     =   -1  'True
            TabIndex        =   97
            Top             =   8070
            Width           =   660
         End
         Begin VB.Label Label2 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "⁄œœ «·”Ã·« :"
            Height          =   225
            Index           =   2
            Left            =   3285
            RightToLeft     =   -1  'True
            TabIndex        =   96
            Top             =   8055
            Width           =   1140
         End
         Begin VB.Label Label2 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "«·”Ã· «·Õ«·Ì:"
            Height          =   225
            Index           =   3
            Left            =   5145
            RightToLeft     =   -1  'True
            TabIndex        =   95
            Top             =   8055
            Width           =   1080
         End
         Begin VB.Label lbl 
            Alignment       =   2  'Center
            BackColor       =   &H00E2E9E9&
            Caption         =   " «—ÌŒ «·”‰œ"
            Height          =   270
            Index           =   2
            Left            =   12030
            TabIndex        =   81
            Top             =   990
            Width           =   1245
         End
         Begin VB.Label lbl 
            Alignment       =   2  'Center
            BackColor       =   &H00E2E9E9&
            Caption         =   "—ﬁ„ «·⁄ﬁœ"
            Height          =   255
            Index           =   4
            Left            =   16200
            RightToLeft     =   -1  'True
            TabIndex        =   80
            Top             =   1275
            Width           =   960
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "«·›—⁄"
            Height          =   255
            Index           =   7
            Left            =   2010
            TabIndex        =   79
            Top             =   870
            Width           =   705
         End
         Begin VB.Label Label1 
            Alignment       =   2  'Center
            BackColor       =   &H00E2E9E9&
            BackStyle       =   0  'Transparent
            Caption         =   "„·«ÕŸ«    "
            Height          =   285
            Index           =   11
            Left            =   16020
            RightToLeft     =   -1  'True
            TabIndex        =   78
            Top             =   3060
            Width           =   1140
         End
         Begin VB.Label lbl 
            Alignment       =   2  'Center
            BackColor       =   &H00E2E9E9&
            Caption         =   "«·⁄„Ì·"
            Height          =   465
            Index           =   15
            Left            =   10095
            TabIndex        =   77
            Top             =   1410
            Width           =   1275
         End
         Begin VB.Label Label4 
            Alignment       =   1  'Right Justify
            Caption         =   "„·«ÕŸ« "
            Height          =   255
            Left            =   14520
            RightToLeft     =   -1  'True
            TabIndex        =   71
            Top             =   2850
            Width           =   780
         End
         Begin VB.Label Label3 
            Alignment       =   1  'Right Justify
            Caption         =   "Label3"
            Height          =   135
            Index           =   1
            Left            =   3615
            RightToLeft     =   -1  'True
            TabIndex        =   59
            Top             =   0
            Width           =   330
         End
      End
      Begin C1SizerLibCtl.C1Elastic ELe 
         Height          =   8790
         Index           =   4
         Left            =   -16815
         TabIndex        =   183
         TabStop         =   0   'False
         Top             =   45
         Width           =   16170
         _cx             =   28522
         _cy             =   15505
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
         Begin VB.Frame Frame1 
            BackColor       =   &H00E2E9E9&
            BorderStyle     =   0  'None
            Enabled         =   0   'False
            Height          =   1665
            Index           =   2
            Left            =   75
            RightToLeft     =   -1  'True
            TabIndex        =   192
            Top             =   4560
            Width           =   7365
            Begin VB.TextBox txtNamee 
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
               Index           =   2
               Left            =   1395
               MaxLength       =   50
               RightToLeft     =   -1  'True
               TabIndex        =   196
               Top             =   1020
               Width           =   2760
            End
            Begin VB.TextBox txtName 
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
               Index           =   2
               Left            =   1395
               MaxLength       =   50
               RightToLeft     =   -1  'True
               TabIndex        =   195
               Top             =   705
               Width           =   2760
            End
            Begin VB.TextBox TxtSerial1 
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
               Index           =   2
               Left            =   3090
               MaxLength       =   50
               RightToLeft     =   -1  'True
               TabIndex        =   194
               Top             =   330
               Width           =   1065
            End
            Begin VB.ComboBox Combo3 
               BackColor       =   &H80000018&
               Height          =   315
               ItemData        =   "FrmVizitScreen.frx":2E7F3
               Left            =   2280
               List            =   "FrmVizitScreen.frx":2E803
               RightToLeft     =   -1  'True
               Style           =   2  'Dropdown List
               TabIndex        =   193
               Top             =   3150
               Visible         =   0   'False
               Width           =   1005
            End
            Begin VB.Label Label3 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "«”„ «‰Ã·Ì“Ì"
               Height          =   285
               Index           =   0
               Left            =   4200
               RightToLeft     =   -1  'True
               TabIndex        =   199
               Top             =   1140
               Width           =   1500
            End
            Begin VB.Label Label5 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "«”„ ⁄—»Ì"
               Height          =   285
               Index           =   0
               Left            =   4350
               RightToLeft     =   -1  'True
               TabIndex        =   198
               Top             =   780
               Width           =   1350
            End
            Begin VB.Label Label6 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "«·ﬂÊœ "
               Height          =   195
               Index           =   0
               Left            =   4695
               RightToLeft     =   -1  'True
               TabIndex        =   197
               Top             =   450
               Width           =   990
            End
         End
         Begin VB.Frame Fra_Header 
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   0  'None
            Height          =   585
            Index           =   2
            Left            =   -135
            RightToLeft     =   -1  'True
            TabIndex        =   184
            Top             =   0
            Width           =   16170
            Begin VB.TextBox TxtModFlg2 
               Alignment       =   1  'Right Justify
               Appearance      =   0  'Flat
               BackColor       =   &H0000FF00&
               Enabled         =   0   'False
               Height          =   285
               Index           =   2
               Left            =   2580
               RightToLeft     =   -1  'True
               TabIndex        =   186
               Text            =   "modflag"
               Top             =   90
               Visible         =   0   'False
               Width           =   465
            End
            Begin VB.TextBox Text2 
               Alignment       =   1  'Right Justify
               Appearance      =   0  'Flat
               BackColor       =   &H000000FF&
               Height          =   240
               Index           =   2
               Left            =   3030
               RightToLeft     =   -1  'True
               TabIndex        =   185
               Top             =   510
               Visible         =   0   'False
               Width           =   945
            End
            Begin MSComctlLib.ImageList GrdImageList2 
               Index           =   2
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
                     Picture         =   "FrmVizitScreen.frx":2E81C
                     Key             =   "CompanyName"
                  EndProperty
                  BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                     Picture         =   "FrmVizitScreen.frx":2EBB6
                     Key             =   "Ser"
                  EndProperty
                  BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                     Picture         =   "FrmVizitScreen.frx":2EF50
                     Key             =   "Vac_Name"
                  EndProperty
                  BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                     Picture         =   "FrmVizitScreen.frx":2F2EA
                     Key             =   "ShareCount"
                  EndProperty
                  BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                     Picture         =   "FrmVizitScreen.frx":2F684
                     Key             =   "Dis_Count"
                  EndProperty
                  BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                     Picture         =   "FrmVizitScreen.frx":2FA1E
                     Key             =   "Bouns"
                  EndProperty
                  BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                     Picture         =   "FrmVizitScreen.frx":2FDB8
                     Key             =   "SharesValue"
                  EndProperty
                  BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                     Picture         =   "FrmVizitScreen.frx":30352
                     Key             =   "BuyValue"
                  EndProperty
               EndProperty
            End
            Begin ImpulseButton.ISButton btn_Last 
               Height          =   315
               Index           =   2
               Left            =   90
               TabIndex        =   187
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
               ButtonImage     =   "FrmVizitScreen.frx":306EC
               ColorButton     =   14871017
               AcclimateGrayTones=   -1  'True
               DrawFocusRectangle=   0   'False
               DisabledImageExtraction=   0
               DisabledImageStyle=   1
            End
            Begin ImpulseButton.ISButton btn_Next 
               Height          =   315
               Index           =   2
               Left            =   555
               TabIndex        =   188
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
               ButtonImage     =   "FrmVizitScreen.frx":30A86
               ColorButton     =   14871017
               DrawFocusRectangle=   0   'False
               DisabledImageStyle=   1
            End
            Begin ImpulseButton.ISButton btn_Previous 
               Height          =   315
               Index           =   2
               Left            =   1155
               TabIndex        =   189
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
               ButtonImage     =   "FrmVizitScreen.frx":30E20
               ColorButton     =   14871017
               DrawFocusRectangle=   0   'False
               DisabledImageStyle=   1
            End
            Begin ImpulseButton.ISButton btn_First 
               Height          =   315
               Index           =   2
               Left            =   1620
               TabIndex        =   190
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
               ButtonImage     =   "FrmVizitScreen.frx":311BA
               ColorButton     =   14871017
               DrawFocusRectangle=   0   'False
               DisabledImageStyle=   1
            End
            Begin VB.Label Label1 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "«‰Ê«⁄ «·—«Õ« "
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
               Index           =   12
               Left            =   11340
               RightToLeft     =   -1  'True
               TabIndex        =   191
               Top             =   90
               Width           =   2640
            End
         End
         Begin ImpulseButton.ISButton btn_New 
            Height          =   345
            Index           =   2
            Left            =   6255
            TabIndex        =   200
            Top             =   7680
            Width           =   975
            _ExtentX        =   1720
            _ExtentY        =   609
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
            ButtonImage     =   "FrmVizitScreen.frx":31554
            ColorButton     =   14871017
            DrawFocusRectangle=   0   'False
         End
         Begin ImpulseButton.ISButton btn_Save 
            Height          =   345
            Index           =   2
            Left            =   4575
            TabIndex        =   201
            Top             =   7650
            Width           =   780
            _ExtentX        =   1376
            _ExtentY        =   609
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
            ButtonImage     =   "FrmVizitScreen.frx":318EE
            ColorButton     =   14871017
            DrawFocusRectangle=   0   'False
         End
         Begin ImpulseButton.ISButton btn_Modify 
            Height          =   345
            Index           =   2
            Left            =   5385
            TabIndex        =   202
            Top             =   7680
            Width           =   840
            _ExtentX        =   1482
            _ExtentY        =   609
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
            ButtonImage     =   "FrmVizitScreen.frx":31C88
            ColorButton     =   14871017
            DrawFocusRectangle=   0   'False
         End
         Begin ImpulseButton.ISButton Btn_Undo 
            Height          =   345
            Index           =   2
            Left            =   3705
            TabIndex        =   203
            Top             =   7650
            Width           =   855
            _ExtentX        =   1508
            _ExtentY        =   609
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
            ButtonImage     =   "FrmVizitScreen.frx":32022
            ColorButton     =   14871017
            DrawFocusRectangle=   0   'False
         End
         Begin ImpulseButton.ISButton btn_Delete 
            Height          =   345
            Index           =   2
            Left            =   3045
            TabIndex        =   204
            Top             =   7680
            Width           =   645
            _ExtentX        =   1138
            _ExtentY        =   609
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
            ButtonImage     =   "FrmVizitScreen.frx":323BC
            ColorButton     =   14871017
            DrawFocusRectangle=   0   'False
         End
         Begin ImpulseButton.ISButton Btn_Update 
            Height          =   225
            Index           =   2
            Left            =   5625
            TabIndex        =   205
            TabStop         =   0   'False
            ToolTipText     =   " ÕœÌÀ ﬁ«⁄œ… «·»Ì«‰« "
            Top             =   6450
            Visible         =   0   'False
            Width           =   1605
            _ExtentX        =   2831
            _ExtentY        =   397
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
            ButtonImage     =   "FrmVizitScreen.frx":32956
            ColorButton     =   14871017
            DrawFocusRectangle=   0   'False
         End
         Begin ImpulseButton.ISButton btn_Cancel 
            Height          =   345
            Index           =   2
            Left            =   75
            TabIndex        =   206
            Top             =   7620
            Width           =   900
            _ExtentX        =   1588
            _ExtentY        =   609
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
            ButtonImage     =   "FrmVizitScreen.frx":32CF0
            ColorButton     =   14871017
            DrawFocusRectangle=   0   'False
            DisabledImageStyle=   1
         End
         Begin ImpulseButton.ISButton Btn_Print 
            Height          =   435
            Index           =   2
            Left            =   1995
            TabIndex        =   207
            TabStop         =   0   'False
            ToolTipText     =   "ÿ»«⁄… «·»Ì«‰«  "
            Top             =   7620
            Width           =   945
            _ExtentX        =   1667
            _ExtentY        =   767
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
            ButtonImage     =   "FrmVizitScreen.frx":3308A
            ColorButton     =   14871017
            DrawFocusRectangle=   0   'False
            DisabledImageStyle=   1
         End
         Begin ImpulseButton.ISButton btn_Query 
            Height          =   540
            Index           =   2
            Left            =   1275
            TabIndex        =   208
            TabStop         =   0   'False
            ToolTipText     =   "(Ctrl+F)  ··»ÕÀ ≈÷€ÿ Â–« «·„› «Õ √Ê ≈÷€ÿ "
            Top             =   7545
            Width           =   645
            _ExtentX        =   1138
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
            ButtonImage     =   "FrmVizitScreen.frx":398EC
            ColorButton     =   14871017
            DrawFocusRectangle=   0   'False
         End
         Begin VSFlex8Ctl.VSFlexGrid Grid2 
            Height          =   3480
            Left            =   0
            TabIndex        =   209
            Top             =   840
            Width           =   7575
            _cx             =   13361
            _cy             =   6138
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
            FormatString    =   $"FrmVizitScreen.frx":39C86
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
         Begin VB.Label Label2 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "«·”Ã· «·Õ«·Ì:"
            Height          =   225
            Index           =   7
            Left            =   5295
            RightToLeft     =   -1  'True
            TabIndex        =   217
            Top             =   6705
            Width           =   1830
         End
         Begin VB.Label Label2 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "⁄œœ «·”Ã·« :"
            Height          =   225
            Index           =   6
            Left            =   1890
            RightToLeft     =   -1  'True
            TabIndex        =   216
            Top             =   6705
            Width           =   1845
         End
         Begin VB.Label LabCurr_Rec 
            BackColor       =   &H00E2E9E9&
            Height          =   240
            Index           =   0
            Left            =   3735
            RightToLeft     =   -1  'True
            TabIndex        =   215
            Top             =   6720
            Width           =   1290
         End
         Begin VB.Label LabCount_Rec 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Height          =   240
            Index           =   0
            Left            =   615
            RightToLeft     =   -1  'True
            TabIndex        =   214
            Top             =   6720
            Width           =   990
         End
         Begin VB.Label Label2 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "«·”Ã· «·Õ«·Ì:"
            Height          =   225
            Index           =   4
            Left            =   5805
            RightToLeft     =   -1  'True
            TabIndex        =   213
            Top             =   7170
            Width           =   1260
         End
         Begin VB.Label Label2 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "⁄œœ «·”Ã·« :"
            Height          =   225
            Index           =   5
            Left            =   2460
            RightToLeft     =   -1  'True
            TabIndex        =   212
            Top             =   7170
            Width           =   1815
         End
         Begin VB.Label LabCurr_Rec 
            BackColor       =   &H00E2E9E9&
            Height          =   240
            Index           =   2
            Left            =   4275
            RightToLeft     =   -1  'True
            TabIndex        =   211
            Top             =   7185
            Width           =   1335
         End
         Begin VB.Label LabCount_Rec 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Height          =   240
            Index           =   2
            Left            =   1125
            RightToLeft     =   -1  'True
            TabIndex        =   210
            Top             =   7185
            Width           =   1050
         End
      End
      Begin C1SizerLibCtl.C1Elastic ELe 
         Height          =   8790
         Index           =   5
         Left            =   45
         TabIndex        =   218
         TabStop         =   0   'False
         Top             =   45
         Width           =   16170
         _cx             =   28522
         _cy             =   15505
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
            Caption         =   "«‰‘«¡ «·«’‰«›"
            Height          =   450
            Left            =   1200
            RightToLeft     =   -1  'True
            TabIndex        =   253
            Top             =   4620
            Visible         =   0   'False
            Width           =   2445
         End
         Begin VB.ComboBox cmbFlag 
            Height          =   315
            Index           =   2
            ItemData        =   "FrmVizitScreen.frx":39D15
            Left            =   1230
            List            =   "FrmVizitScreen.frx":39D17
            RightToLeft     =   -1  'True
            Style           =   2  'Dropdown List
            TabIndex        =   252
            Top             =   4260
            Visible         =   0   'False
            Width           =   1560
         End
         Begin VB.ComboBox cmbFlag 
            Height          =   315
            Index           =   1
            ItemData        =   "FrmVizitScreen.frx":39D19
            Left            =   2850
            List            =   "FrmVizitScreen.frx":39D1B
            RightToLeft     =   -1  'True
            Style           =   2  'Dropdown List
            TabIndex        =   251
            Top             =   4290
            Visible         =   0   'False
            Width           =   1560
         End
         Begin VB.ComboBox cmbFlag 
            Height          =   315
            Index           =   0
            ItemData        =   "FrmVizitScreen.frx":39D1D
            Left            =   2640
            List            =   "FrmVizitScreen.frx":39D1F
            RightToLeft     =   -1  'True
            TabIndex        =   250
            Text            =   "cmbFlag"
            Top             =   2850
            Width           =   1560
         End
         Begin VB.Frame Fra_Header 
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   0  'None
            Height          =   585
            Index           =   1
            Left            =   -135
            RightToLeft     =   -1  'True
            TabIndex        =   226
            Top             =   0
            Width           =   16170
            Begin VB.TextBox Text2 
               Alignment       =   1  'Right Justify
               Appearance      =   0  'Flat
               BackColor       =   &H000000FF&
               Height          =   240
               Index           =   1
               Left            =   3030
               RightToLeft     =   -1  'True
               TabIndex        =   228
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
               Index           =   3
               Left            =   2580
               RightToLeft     =   -1  'True
               TabIndex        =   227
               Text            =   "modflag"
               Top             =   90
               Visible         =   0   'False
               Width           =   465
            End
            Begin MSComctlLib.ImageList GrdImageList2 
               Index           =   3
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
                     Picture         =   "FrmVizitScreen.frx":39D21
                     Key             =   "CompanyName"
                  EndProperty
                  BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                     Picture         =   "FrmVizitScreen.frx":3A0BB
                     Key             =   "Ser"
                  EndProperty
                  BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                     Picture         =   "FrmVizitScreen.frx":3A455
                     Key             =   "Vac_Name"
                  EndProperty
                  BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                     Picture         =   "FrmVizitScreen.frx":3A7EF
                     Key             =   "ShareCount"
                  EndProperty
                  BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                     Picture         =   "FrmVizitScreen.frx":3AB89
                     Key             =   "Dis_Count"
                  EndProperty
                  BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                     Picture         =   "FrmVizitScreen.frx":3AF23
                     Key             =   "Bouns"
                  EndProperty
                  BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                     Picture         =   "FrmVizitScreen.frx":3B2BD
                     Key             =   "SharesValue"
                  EndProperty
                  BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                     Picture         =   "FrmVizitScreen.frx":3B857
                     Key             =   "BuyValue"
                  EndProperty
               EndProperty
            End
            Begin ImpulseButton.ISButton btn_Last 
               Height          =   315
               Index           =   3
               Left            =   90
               TabIndex        =   229
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
               ButtonImage     =   "FrmVizitScreen.frx":3BBF1
               ColorButton     =   14871017
               AcclimateGrayTones=   -1  'True
               DrawFocusRectangle=   0   'False
               DisabledImageExtraction=   0
               DisabledImageStyle=   1
            End
            Begin ImpulseButton.ISButton btn_Next 
               Height          =   315
               Index           =   3
               Left            =   555
               TabIndex        =   230
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
               ButtonImage     =   "FrmVizitScreen.frx":3BF8B
               ColorButton     =   14871017
               DrawFocusRectangle=   0   'False
               DisabledImageStyle=   1
            End
            Begin ImpulseButton.ISButton btn_Previous 
               Height          =   315
               Index           =   3
               Left            =   1155
               TabIndex        =   231
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
               ButtonImage     =   "FrmVizitScreen.frx":3C325
               ColorButton     =   14871017
               DrawFocusRectangle=   0   'False
               DisabledImageStyle=   1
            End
            Begin ImpulseButton.ISButton btn_First 
               Height          =   315
               Index           =   3
               Left            =   1620
               TabIndex        =   232
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
               ButtonImage     =   "FrmVizitScreen.frx":3C6BF
               ColorButton     =   14871017
               DrawFocusRectangle=   0   'False
               DisabledImageStyle=   1
            End
            Begin VB.Label Label1 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   " ⁄—Ì› «·⁄œ”« "
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
               Index           =   15
               Left            =   11340
               RightToLeft     =   -1  'True
               TabIndex        =   233
               Top             =   90
               Width           =   2640
            End
         End
         Begin VB.Frame Frame1 
            BackColor       =   &H00E2E9E9&
            BorderStyle     =   0  'None
            Height          =   3795
            Index           =   3
            Left            =   75
            RightToLeft     =   -1  'True
            TabIndex        =   219
            Top             =   720
            Width           =   6105
            Begin VB.TextBox txtPrice 
               Alignment       =   1  'Right Justify
               Height          =   315
               Left            =   90
               RightToLeft     =   -1  'True
               TabIndex        =   267
               Top             =   2100
               Width           =   1185
            End
            Begin VB.TextBox TxtSerial1 
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
               Index           =   3
               Left            =   3090
               MaxLength       =   50
               RightToLeft     =   -1  'True
               TabIndex        =   222
               Top             =   330
               Width           =   1065
            End
            Begin VB.TextBox txtName 
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
               Index           =   3
               Left            =   1395
               MaxLength       =   50
               RightToLeft     =   -1  'True
               TabIndex        =   221
               Top             =   705
               Width           =   2760
            End
            Begin VB.TextBox txtNamee 
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
               Index           =   3
               Left            =   1395
               MaxLength       =   50
               RightToLeft     =   -1  'True
               TabIndex        =   220
               Top             =   1020
               Width           =   2760
            End
            Begin MSDataListLib.DataCombo DCBoMain 
               Height          =   360
               Index           =   2
               Left            =   1470
               TabIndex        =   256
               Top             =   2520
               Width           =   1395
               _ExtentX        =   2461
               _ExtentY        =   635
               _Version        =   393216
               BackColor       =   16761024
               Text            =   ""
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "MS Sans Serif"
                  Size            =   9.75
                  Charset         =   178
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
            End
            Begin MSDataListLib.DataCombo DCBoMain 
               Height          =   360
               Index           =   5
               Left            =   4380
               TabIndex        =   258
               Top             =   2520
               Width           =   1485
               _ExtentX        =   2619
               _ExtentY        =   635
               _Version        =   393216
               BackColor       =   16761024
               Text            =   ""
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "MS Sans Serif"
                  Size            =   9.75
                  Charset         =   178
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
            End
            Begin MSDataListLib.DataCombo DCBoMain 
               Height          =   360
               Index           =   3
               Left            =   1470
               TabIndex        =   260
               Top             =   3060
               Width           =   1365
               _ExtentX        =   2408
               _ExtentY        =   635
               _Version        =   393216
               BackColor       =   16761024
               Text            =   ""
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "MS Sans Serif"
                  Size            =   9.75
                  Charset         =   178
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
            End
            Begin MSDataListLib.DataCombo DCBoMain 
               Height          =   360
               Index           =   6
               Left            =   4380
               TabIndex        =   261
               Top             =   3000
               Width           =   1485
               _ExtentX        =   2619
               _ExtentY        =   635
               _Version        =   393216
               BackColor       =   16761024
               Text            =   ""
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "MS Sans Serif"
                  Size            =   9.75
                  Charset         =   178
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
            End
            Begin MSDataListLib.DataCombo cmbGroupId 
               Height          =   315
               Left            =   90
               TabIndex        =   265
               Top             =   1350
               Width           =   4065
               _ExtentX        =   7170
               _ExtentY        =   556
               _Version        =   393216
               ListField       =   "6"
               Text            =   ""
               RightToLeft     =   -1  'True
            End
            Begin MSDataListLib.DataCombo cmbUnitID 
               Height          =   315
               Left            =   90
               TabIndex        =   266
               Top             =   1710
               Width           =   4065
               _ExtentX        =   7170
               _ExtentY        =   556
               _Version        =   393216
               ListField       =   "6"
               Text            =   ""
               RightToLeft     =   -1  'True
            End
            Begin VB.Label lbl 
               Alignment       =   2  'Center
               BackColor       =   &H00E2E9E9&
               BackStyle       =   0  'Transparent
               Caption         =   "From CYL"
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
               Height          =   345
               Index           =   26
               Left            =   90
               RightToLeft     =   -1  'True
               TabIndex        =   263
               Top             =   3000
               Width           =   1380
            End
            Begin VB.Label lbl 
               Alignment       =   2  'Center
               BackColor       =   &H00E2E9E9&
               BackStyle       =   0  'Transparent
               Caption         =   "To CYL"
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
               Height          =   345
               Index           =   25
               Left            =   2700
               RightToLeft     =   -1  'True
               TabIndex        =   262
               Top             =   3000
               Width           =   1380
            End
            Begin VB.Label lbl 
               Alignment       =   2  'Center
               BackColor       =   &H00E2E9E9&
               BackStyle       =   0  'Transparent
               Caption         =   "To Sph"
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
               Height          =   345
               Index           =   24
               Left            =   2700
               RightToLeft     =   -1  'True
               TabIndex        =   259
               Top             =   2520
               Width           =   1380
            End
            Begin VB.Label lbl 
               Alignment       =   2  'Center
               BackColor       =   &H00E2E9E9&
               BackStyle       =   0  'Transparent
               Caption         =   "From Sph"
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
               Height          =   345
               Index           =   152
               Left            =   30
               RightToLeft     =   -1  'True
               TabIndex        =   257
               Top             =   2490
               Width           =   1380
            End
            Begin VB.Label Label3 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "«·«‘«—…"
               Height          =   285
               Index           =   6
               Left            =   4020
               RightToLeft     =   -1  'True
               TabIndex        =   255
               Top             =   2190
               Width           =   990
            End
            Begin VB.Label Label3 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "„Ã„Ê⁄… «·’‰›"
               Height          =   285
               Index           =   4
               Left            =   4230
               RightToLeft     =   -1  'True
               TabIndex        =   249
               Top             =   1470
               Width           =   1500
            End
            Begin VB.Label Label3 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "«·ÊÕœ…"
               Height          =   285
               Index           =   3
               Left            =   4200
               RightToLeft     =   -1  'True
               TabIndex        =   248
               Top             =   1860
               Width           =   1500
            End
            Begin VB.Label Label6 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "«·ﬂÊœ "
               Height          =   195
               Index           =   1
               Left            =   4695
               RightToLeft     =   -1  'True
               TabIndex        =   225
               Top             =   450
               Width           =   990
            End
            Begin VB.Label Label5 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "«”„ ⁄—»Ì"
               Height          =   285
               Index           =   2
               Left            =   4350
               RightToLeft     =   -1  'True
               TabIndex        =   224
               Top             =   780
               Width           =   1350
            End
            Begin VB.Label Label3 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "«”„ «‰Ã·Ì“Ì"
               Height          =   285
               Index           =   2
               Left            =   4200
               RightToLeft     =   -1  'True
               TabIndex        =   223
               Top             =   1140
               Width           =   1500
            End
         End
         Begin ImpulseButton.ISButton btn_New 
            Height          =   345
            Index           =   3
            Left            =   6255
            TabIndex        =   234
            Top             =   8280
            Width           =   975
            _ExtentX        =   1720
            _ExtentY        =   609
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
            ButtonImage     =   "FrmVizitScreen.frx":3CA59
            ColorButton     =   14871017
            DrawFocusRectangle=   0   'False
         End
         Begin ImpulseButton.ISButton btn_Save 
            Height          =   345
            Index           =   3
            Left            =   4575
            TabIndex        =   235
            Top             =   8250
            Width           =   780
            _ExtentX        =   1376
            _ExtentY        =   609
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
            ButtonImage     =   "FrmVizitScreen.frx":3CDF3
            ColorButton     =   14871017
            DrawFocusRectangle=   0   'False
         End
         Begin ImpulseButton.ISButton btn_Modify 
            Height          =   345
            Index           =   3
            Left            =   5385
            TabIndex        =   236
            Top             =   8280
            Width           =   840
            _ExtentX        =   1482
            _ExtentY        =   609
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
            ButtonImage     =   "FrmVizitScreen.frx":3D18D
            ColorButton     =   14871017
            DrawFocusRectangle=   0   'False
         End
         Begin ImpulseButton.ISButton Btn_Undo 
            Height          =   345
            Index           =   3
            Left            =   3705
            TabIndex        =   237
            Top             =   8250
            Width           =   855
            _ExtentX        =   1508
            _ExtentY        =   609
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
            ButtonImage     =   "FrmVizitScreen.frx":3D527
            ColorButton     =   14871017
            DrawFocusRectangle=   0   'False
         End
         Begin ImpulseButton.ISButton btn_Delete 
            Height          =   345
            Index           =   3
            Left            =   3045
            TabIndex        =   238
            Top             =   8280
            Width           =   645
            _ExtentX        =   1138
            _ExtentY        =   609
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
            ButtonImage     =   "FrmVizitScreen.frx":3D8C1
            ColorButton     =   14871017
            DrawFocusRectangle=   0   'False
         End
         Begin ImpulseButton.ISButton Btn_Update 
            Height          =   225
            Index           =   3
            Left            =   5625
            TabIndex        =   239
            TabStop         =   0   'False
            ToolTipText     =   " ÕœÌÀ ﬁ«⁄œ… «·»Ì«‰« "
            Top             =   7050
            Visible         =   0   'False
            Width           =   1605
            _ExtentX        =   2831
            _ExtentY        =   397
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
            ButtonImage     =   "FrmVizitScreen.frx":3DE5B
            ColorButton     =   14871017
            DrawFocusRectangle=   0   'False
         End
         Begin ImpulseButton.ISButton btn_Cancel 
            Height          =   345
            Index           =   3
            Left            =   75
            TabIndex        =   240
            Top             =   8220
            Width           =   900
            _ExtentX        =   1588
            _ExtentY        =   609
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
            ButtonImage     =   "FrmVizitScreen.frx":3E1F5
            ColorButton     =   14871017
            DrawFocusRectangle=   0   'False
            DisabledImageStyle=   1
         End
         Begin ImpulseButton.ISButton Btn_Print 
            Height          =   435
            Index           =   3
            Left            =   1995
            TabIndex        =   241
            TabStop         =   0   'False
            ToolTipText     =   "ÿ»«⁄… «·»Ì«‰«  "
            Top             =   8220
            Width           =   945
            _ExtentX        =   1667
            _ExtentY        =   767
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
            ButtonImage     =   "FrmVizitScreen.frx":3E58F
            ColorButton     =   14871017
            DrawFocusRectangle=   0   'False
            DisabledImageStyle=   1
         End
         Begin ImpulseButton.ISButton btn_Query 
            Height          =   540
            Index           =   3
            Left            =   1275
            TabIndex        =   242
            TabStop         =   0   'False
            ToolTipText     =   "(Ctrl+F)  ··»ÕÀ ≈÷€ÿ Â–« «·„› «Õ √Ê ≈÷€ÿ "
            Top             =   8145
            Width           =   645
            _ExtentX        =   1138
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
            ButtonImage     =   "FrmVizitScreen.frx":44DF1
            ColorButton     =   14871017
            DrawFocusRectangle=   0   'False
         End
         Begin VSFlex8Ctl.VSFlexGrid Grid3 
            Height          =   2160
            Left            =   6480
            TabIndex        =   243
            Top             =   780
            Width           =   9585
            _cx             =   16907
            _cy             =   3810
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
            FormatString    =   $"FrmVizitScreen.frx":4518B
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
         Begin VSFlex8UCtl.VSFlexGrid grdSphCYL 
            Height          =   1695
            Left            =   6480
            TabIndex        =   254
            Top             =   2970
            Width           =   9630
            _cx             =   16986
            _cy             =   2990
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
            Cols            =   3
            FixedRows       =   1
            FixedCols       =   1
            RowHeightMin    =   320
            RowHeightMax    =   0
            ColWidthMin     =   0
            ColWidthMax     =   0
            ExtendLastCol   =   -1  'True
            FormatString    =   $"FrmVizitScreen.frx":4521A
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
            RightToLeft     =   0   'False
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
         Begin VSFlex8Ctl.VSFlexGrid GrdItems 
            Height          =   2160
            Left            =   6420
            TabIndex        =   264
            Top             =   4680
            Width           =   9585
            _cx             =   16907
            _cy             =   3810
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
            Cols            =   11
            FixedRows       =   1
            FixedCols       =   1
            RowHeightMin    =   320
            RowHeightMax    =   0
            ColWidthMin     =   0
            ColWidthMax     =   0
            ExtendLastCol   =   -1  'True
            FormatString    =   $"FrmVizitScreen.frx":4528C
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
         Begin VB.Label LabCurr_Rec 
            BackColor       =   &H00E2E9E9&
            Height          =   240
            Index           =   3
            Left            =   4275
            RightToLeft     =   -1  'True
            TabIndex        =   247
            Top             =   7785
            Width           =   1335
         End
         Begin VB.Label Label2 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "⁄œœ «·”Ã·« :"
            Height          =   225
            Index           =   11
            Left            =   2670
            RightToLeft     =   -1  'True
            TabIndex        =   246
            Top             =   7770
            Width           =   1455
         End
         Begin VB.Label Label2 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "«·”Ã· «·Õ«·Ì:"
            Height          =   225
            Index           =   10
            Left            =   5805
            RightToLeft     =   -1  'True
            TabIndex        =   245
            Top             =   7770
            Width           =   1260
         End
         Begin VB.Label LabCount_Rec 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Height          =   240
            Index           =   3
            Left            =   1185
            RightToLeft     =   -1  'True
            TabIndex        =   244
            Top             =   7680
            Width           =   990
         End
      End
   End
End
Attribute VB_Name = "FrmVizitScreen"
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
Dim MintDone As Integer
Public mIndex As Integer
Dim mIndex2 As Integer
Dim Dcombos As ClsDataCombos
Dim mGridClicked As Boolean
Dim mDiscEnter As Boolean
Dim rsDummy As New ADODB.Recordset
Dim s As String

Private Sub cmbFlag_Change(Index As Integer)
    
        Dim StrSQL As String

        Dim s As String
    
        
        s = "Select sph as sph , spht  as SPHName From SPHTable  Where 1 = 1 "
        
        If cmbFlag(0).ListIndex = 0 Then
            s = s & " and  sph< 0"
        ElseIf cmbFlag(0).ListIndex = 1 Then
            s = s & " and  sph> 0"
        ElseIf cmbFlag(0).ListIndex = 2 Then
            s = s & " and  sph> 0"
        End If
        s = s & " order by id"
        
        fill_combo DCBoMain(2), s
        fill_combo DCBoMain(5), s
        
        
        s = " Select   CLY as CLY ,CLYT  as CLYName  From CLYTable   "
        s = s & " Where 1 = 1 "
        If cmbFlag(0).ListIndex = 0 Then
            s = s & " and  CLY < 0"
        ElseIf cmbFlag(0).ListIndex = 1 Then
            s = s & " and  CLY < 0"
        ElseIf cmbFlag(0).ListIndex = 2 Then
            s = s & " and  CLY > 0"
        End If
        s = s & " ORDER BY CLY"
        
        fill_combo DCBoMain(3), s
        fill_combo DCBoMain(6), s
        
End Sub

Private Sub cmbFlag_Click(Index As Integer)
cmbFlag_Change Index
End Sub

Private Sub Command1_Click()

Dim mNewCode As String
s = "Select * from SPHTable Where SPH > = " & val(DCBoMain(2).BoundText) & " and SPH <= " & val(DCBoMain(5).BoundText)
s = "Select * from SPHTable Where SPH > = " & val(DCBoMain(2).BoundText) & " and SPH <= " & val(DCBoMain(5).BoundText)
GenreateItems
Dim tRs As New ADODB.Recordset
Dim tRs2 As New ADODB.Recordset

    Dim mMaxId As Long
    Dim rsDummy As New ADODB.Recordset
    s = "SELECT Max(ItemID) MaxID  FROM tblItems AS te "
    rsDummy.Open s, Cn, adOpenKeyset, adLockOptimistic, adCmdText
    If Not rsDummy.EOF Then
        mMaxId = val(rsDummy!MaxID & "")
    End If
    
     s = "SELECT * FROM tblItems WHERE ItemID = -1 "
    tRs.Open s, Cn, adOpenKeyset, adLockOptimistic, adCmdText
    
        
        '**********************
    Dim i As Long
    For i = 1 To GrdItems.Rows - 1
    
        tRs.AddNew
        II = II + 1
        mMaxId = mMaxId + 1
    
        tRs!ItemID = mMaxId
        tRs!HaveSerial = 0
        tRs!HaveGuarantee = 0
        tRs!DealerPrice = 0
        tRs!GuaranteeValue = 0
        tRs!GuaranteeType = 0
        tRs!IsArchive = 0
        tRs!ItemType = 0
        tRs!AssbliedItem = 0
        tRs!RelatedItem = 0
        tRs!ItemCase = 1
        tRs!AssbliedItem = 0
        tRs!ItemName = Trim(GrdItems.TextMatrix(i, GrdItems.ColIndex("ItemName")))
        tRs!GroupID = val(GrdItems.TextMatrix(i, GrdItems.ColIndex("GroupID")))
        
        mNewCode = GetNewCode(val(GrdItems.TextMatrix(i, GrdItems.ColIndex("GroupID"))), "tblItems")
        tRs!Fullcode = mNewCode
        tRs!itemcode = mNewCode
        tRs!barCodeNO = mNewCode
        tRs!code = mNewCode
            
        tRs!SphereID = Trim(GrdItems.TextMatrix(i, GrdItems.ColIndex("SphereID")))
        tRs!CylinderID = Trim(GrdItems.TextMatrix(i, GrdItems.ColIndex("CylinderID")))
        tRs.update
        
        
        s = "SELECT * FROM TblItemsUnits WHERE ItemID = -1 "
        Set tRs2 = New ADODB.Recordset
        tRs2.Open s, Cn, adOpenKeyset, adLockOptimistic, adCmdText
        tRs2.AddNew
        tRs2!ItemID = mMaxId
        tRs2!UnitID = val(GrdItems.TextMatrix(i, GrdItems.ColIndex("UnitID")))
        tRs2!DefaultUnit = 1
        tRs2!UnitFactor = 1
        tRs2!unitsalesprice = val(txtPrice)
        
        tRs2!ForUnit = 0
        tRs2!MethodCalc = 0
        tRs2.update
        
        
    Next
            
 '   MsgBox " „ «‰‘«¡ «·«’‰«›"
    
End Sub


Private Sub GenreateItems()
Dim s As String
            'cmbFlag(CC).AddItem "--"
            'cmbFlag(CC).AddItem "-+"
            'cmbFlag(CC).AddItem "++"

s = " Select   * From CLYTable   "
s = s & " Where CLY > = " & val(DCBoMain(3).BoundText) & " and CLY <= " & val(DCBoMain(6).BoundText)
If cmbFlag(0).ListIndex = 0 Then
    s = s & " and  CLY < 0"
ElseIf cmbFlag(0).ListIndex = 1 Then
    s = s & " and  CLY < 0"
ElseIf cmbFlag(0).ListIndex = 2 Then
    s = s & " and  CLY > 0"
End If
s = s & " ORDER BY CLY"
Dim rsDummy As New ADODB.Recordset
Set rsDummy = New ADODB.Recordset
rsDummy.Open s, Cn, adOpenStatic, adLockReadOnly
grdSphCYL.Cols = 3
grdSphCYL.Rows = 1
Dim i As Long
Dim j As Long
j = 4
Do While Not rsDummy.EOF
    
   grdSphCYL.Cols = j
   grdSphCYL.TextMatrix(0, j - 1) = rsDummy!CLYT & ""
    
    j = j + 1
    
    
    rsDummy.MoveNext

Loop



s = " Select   * From "
s = s & " SPHTable Where SPH > = " & val(DCBoMain(2).BoundText) & " and SPH <= " & val(DCBoMain(5).BoundText)
If cmbFlag(0).ListIndex = 0 Then
    s = s & " and  SPH < 0"
ElseIf cmbFlag(0).ListIndex = 1 Then
    s = s & " and  SPH > 0"
ElseIf cmbFlag(0).ListIndex = 2 Then
    s = s & " and  SPH > 0"
End If
s = s & " ORDER BY SPH"
Set rsDummy = New ADODB.Recordset
rsDummy.Open s, Cn, adOpenStatic, adLockReadOnly

grdSphCYL.Rows = 1


j = 1
i = grdSphCYL.Rows
Do While Not rsDummy.EOF
    grdSphCYL.Rows = i + 1
    grdSphCYL.TextMatrix(i, 2) = rsDummy!SPHT & ""
    For j = 1 To grdSphCYL.Cols - 1
        If j > 2 Then
            grdSphCYL.TextMatrix(i, j) = txtName(mIndex) & " " & grdSphCYL.TextMatrix(0, j) & " " & rsDummy!SPHT & ""
        End If
      '  j = j + 1
    Next j
    
    i = i + 1
    
    
    rsDummy.MoveNext

Loop

GrdItems.Rows = 1
Dim n As Long
n = 1
For i = 1 To grdSphCYL.Rows - 1
    For j = 0 To grdSphCYL.Cols - 1
        If j >= 3 Then
            GrdItems.Rows = GrdItems.Rows + 1
            GrdItems.TextMatrix(n, 0) = n
            GrdItems.TextMatrix(n, GrdItems.ColIndex("GroupID")) = cmbGroupId.BoundText
            GrdItems.TextMatrix(n, GrdItems.ColIndex("UnitID")) = cmbUnitID.BoundText
            GrdItems.TextMatrix(n, GrdItems.ColIndex("GroupName")) = cmbGroupId.Text
            GrdItems.TextMatrix(n, GrdItems.ColIndex("UnitName")) = cmbUnitID.Text
            
            GrdItems.TextMatrix(n, GrdItems.ColIndex("ItemName")) = grdSphCYL.TextMatrix(i, j)
            GrdItems.TextMatrix(n, GrdItems.ColIndex("SPH")) = grdSphCYL.TextMatrix(i, 2)
            GrdItems.TextMatrix(n, GrdItems.ColIndex("SphereID")) = GetSphCylID(0, grdSphCYL.TextMatrix(i, 2))
            
            
            GrdItems.TextMatrix(n, GrdItems.ColIndex("CYL")) = grdSphCYL.TextMatrix(0, j)
            GrdItems.TextMatrix(n, GrdItems.ColIndex("CylinderID")) = GetSphCylID(1, grdSphCYL.TextMatrix(0, j))
            n = n + 1
        End If
    Next
    
Next

End Sub

Private Function GetSphCylID(ByVal mType As Integer, ByVal MSTR As String) As Long

Dim s As String
If mType = 0 Then
    s = "Select Id from SPHTable Where SPH =  " & val(MSTR)
ElseIf mType = 1 Then
    s = "Select Id from CLYTable Where CLY  =  " & val(MSTR)
End If
Dim rsDummy As ADODB.Recordset
Set rsDummy = New ADODB.Recordset
rsDummy.Open s, Cn, adOpenKeyset, adLockReadOnly
If Not rsDummy.EOF Then
    GetSphCylID = val(rsDummy!ID & "")
End If
End Function



Private Sub Grid2_EnterCell()
    On Error GoTo ErrTrap
    FindRec val(Me.Grid2.TextMatrix(Me.Grid2.Row, Me.Grid2.ColIndex("id")))
ErrTrap:
End Sub


Private Sub Grid3_EnterCell()
    On Error GoTo ErrTrap
    FindRec val(Me.Grid3.TextMatrix(Me.Grid3.Row, Me.Grid3.ColIndex("id")))
ErrTrap:
End Sub

Private Sub ImgFavorites_Click()
AddTofaforites Me.Name, Me.Caption, Me.Caption



End Sub
Private Sub CmdCreateV2_Click()
Dim s As String
'CHECKaCCOUNTS
Dim StrAccountCodeCridet As String
                StrAccountCodeCridet = get_account_code_branch(77, my_branch)
        
                If StrAccountCodeCridet = "NO branch" Then
                If SystemOptions.UserInterface = ArabicInterface Then
                    MsgBox "·„ Ì „ «‰‘«¡ «·›—⁄", vbCritical
                Else
                    MsgBox "No Branch Created", vbCritical
                End If

                Exit Sub
            Else

                If StrAccountCodeCridet = "NO account" Then
                    If SystemOptions.UserInterface = ArabicInterface Then
                        MsgBox "·„ Ì „  ÕœÌœ Õ”«»   «»—«œ«  «·’Ì«‰…", vbCritical
                    Else
                        MsgBox "Please Select Account VAT ", vbCritical
                    End If

                    Exit Sub
         
                End If
            End If


'END CHECK
If val(TxtNoteSerial.Text) = 0 Then
If createVoucher2 Then
       'FindRec val(TXTLCNO.Text)
       
            s = "Update TblHandWages Set NoteID = " & val(TxtNoteID) & ",NoteSerial = '" & Trim(TxtNoteSerial) & "' Where Id = " & val(TxtSerial1(mIndex))
            
                    
            Cn.Execute s
            
            FindRec val(TxtSerial1(mIndex).Text)
        If SystemOptions.UserInterface = ArabicInterface Then
           ' MsgBox " „ «‰‘«¡ «·ﬁÌœ"
            If val(TxtNoteID) <> 0 Then
                CmdCreateV2.Enabled = False
                cmdPrintNote.Enabled = True
                cmdDelNote.Enabled = True
                btn_Save(mIndex).Enabled = False
            Else
                CmdCreateV2.Enabled = True
                cmdPrintNote.Enabled = False
                cmdDelNote.Enabled = False
            End If
        Else
        
          '  MsgBox "Done"
        End If
    Else
        CmdCreateV2.Enabled = True
        cmdPrintNote.Enabled = False
        cmdDelNote.Enabled = False
    End If
    
End If
'Me.TxtModFlg2(mIndex).Text = "R"
End Sub
Function createVoucher2() As Boolean

'ee
Dim NoteID As Long
Dim NoteDate As Date
Dim NoteSerial As String
Dim Notevalue As Double
Dim notytype As Integer
Dim des As String
des = "    Õ”«» «·" & TxtNoteSerial.Text


Dim tablename As String
Dim Filedname As String
Dim NoteSerial1 As Long
Dim BranchID As Integer
Dim sql As String
Dim mRate  As Double
tablename = "TblHandWages"

Filedname = "ID"
NoteSerial1 = TxtNoteSerial1

BranchID = val(Dcbranch(mIndex).BoundText)
mRate = 1

'



notytype = 1100
Notevalue = val(txtNet)

'mAccNO = val(DboParentAccount.BoundText)
NoteDate = (XPDtbTrans.value)
 
If Notevalue > 0 Then
    CreateNotes NoteID, NoteDate, BranchID, notytype, Notevalue, NoteSerial, (NoteSerial1), tablename, Filedname, val(TxtSerial1(mIndex)), des                   ', recordDateH.value
                                              TxtNoteID.Text = NoteID
                                                     TxtNoteSerial.Text = NoteSerial

    If Not CREATE_VOUCHER_GE2(val(TxtNoteID.Text), BranchID, val(DCboUserName(mIndex).BoundText), NoteDate) Then createVoucher2 = False Else createVoucher2 = True
    RsSavRec.Resync adAffectCurrent

    updateNotesValueAndNobytext val(TxtNoteSerial.Text), Format(txtNet.Text, "###.00")
'
'
'    StrSQL = "update  " & tablename & "   set NoteID=" & NoteID & ",NoteSerial='" & NoteSerial & "'"
'
'    StrSQL = StrSQL & " Where " & Filedname & " = " & NoteSerial1 & ""
'    Cn.Execute StrSQL
     
     
 
End If
End Function


Public Function CREATE_VOUCHER_GE2(general_noteid As Long, BranchID As Integer, UserID As Long _
, NoteDate As Date) As Boolean
Dim StrSQL As String
     StrSQL = "Delete From DOUBLE_ENTREY_VOUCHERS Where Notes_ID=" & general_noteid
    Cn.Execute StrSQL, , adExecuteNoRecords
    
    Dim i As Integer
    Dim sql As String
    Dim StoreID6 As Integer
    Dim rs2 As ADODB.Recordset
    Set rs2 = New ADODB.Recordset
    Dim Notevalue As Double
    Dim LngDevID As Long
    Dim Msg As String
    Dim StrAccountCodeDebt As String
    Dim StrAccountCodeCridet As String
    Dim X As Integer
   
    Dim rs As New ADODB.Recordset
    Dim notes_serial As String
    Dim notes_id As String
    Msg = "    Õ”«» " & TxtSerial1(mIndex).Text
    notes_id = general_noteid
    my_branch = val(Dcbranch(mIndex))
    LngDevID = new_id("DOUBLE_ENTREY_VOUCHERS", "Double_Entry_Vouchers_ID", "")
    Dim line_no As Integer
    line_no = 1
    
    Dim s As String
    Dim mRate As Double
    mRate = 1
    ' „‰ Õ”«» «·⁄„Ì·
    StrAccountCodeDebt = ModAccounts.GetMyAccountCode("TblCustemers", "CusID", val(Me.DcCustmer(mIndex).BoundText))
    

   
    Notevalue = val(txtNet.Text)
    If Notevalue > 0 Then
        
       ' StrAccountCodeDebt = Trim(DboParentAccount.BoundText)
        If ModAccounts.AddNewDev(LngDevID, line_no, StrAccountCodeDebt, Notevalue, 0, Msg & "    Õ”«»  «·⁄„Ì·  ", val(notes_id), , , , XPDtbTrans.value, val(DCboUserName(mIndex).BoundText), , , , , , CLng(mRate), , , setfoxy_Line, , , , , , , , , _
        val(branch_id), , , , , , , , , , , , , , , , , , , , , , , , DcCustmer(mIndex).BoundText) = False Then
            GoTo ErrTrap
        End If
       ' «·Ï Õ”«» «·ﬁÌ„… «·„÷«›…
        GetValueAddedAccount XPDtbTrans.value, , StrAccountCodeCridet, 1, 10
        
        line_no = line_no + 1

        If ModAccounts.AddNewDev(LngDevID, line_no, StrAccountCodeCridet, val(txtVat2), 1, Msg & "    Õ”«»  «·ﬁÌ„… «·„÷«›… ", val(notes_id), , , , XPDtbTrans.value, val(DCboUserName(mIndex).BoundText), , , , , , CLng(mRate), , , setfoxy_Line, , , , , , , , , val(branch_id)) = False Then
            GoTo ErrTrap
        End If
        line_no = line_no + 1
    End If

    
    ' «·«ÿ—«›
    
     ' «·Ï Õ”«» «Ì—«œ«  «·Õ«ÊÌ« 
         
    Notevalue = val(txtTotal.Text)
    If Notevalue > 0 Then
    
                StrAccountCodeCridet = get_account_code_branch(77, my_branch)
        
                If StrAccountCodeCridet = "NO branch" Then
                If SystemOptions.UserInterface = ArabicInterface Then
                    MsgBox "·„ Ì „ «‰‘«¡ «·›—⁄", vbCritical
                Else
                    MsgBox "No Branch Created", vbCritical
                End If

                GoTo ErrTrap
            Else

                If StrAccountCodeCridet = "NO account" Then
                    If SystemOptions.UserInterface = ArabicInterface Then
                        MsgBox "·„ Ì „  ÕœÌœ Õ”«»   «»—«œ«  «·’Ì«‰…", vbCritical
                    Else
                        MsgBox "Please Select Account VAT ", vbCritical
                    End If

                    GoTo ErrTrap
         
                End If
            End If

        
        
 
        
        If ModAccounts.AddNewDev(LngDevID, line_no, StrAccountCodeCridet, Notevalue, 1, Msg & "    Õ”«» «Ì—«œ«  «·’Ì«‰…  ", val(notes_id), , , , XPDtbTrans.value, val(DCboUserName(mIndex).BoundText), , , , , , CLng(mRate), , , setfoxy_Line, , , , , , , , , _
        val(branch_id)) = False Then
            GoTo ErrTrap
        End If

        line_no = line_no + 1
    End If
    

    updateNotesValueAndNobytext (val(notes_id))
    CREATE_VOUCHER_GE2 = True
    Exit Function
ErrTrap:
CREATE_VOUCHER_GE2 = False
TxtNoteID = ""
TxtNoteSerial = ""
CmdCreateV2.Enabled = True
  End Function

Private Sub cmdDelNote_Click()

Dim X As Integer
Dim Msg As String
Dim StrSQL As String
    
        X = vbYes

      If X = vbYes Then
        StrSQL = "Delete From DOUBLE_ENTREY_VOUCHERS Where Notes_ID=" & val(Me.TxtNoteID.Text)
        Cn.Execute StrSQL, , adExecuteNoRecords
        StrSQL = "Delete From Notes Where NoteID=" & val(Me.TxtNoteID.Text)
        Cn.Execute StrSQL, , adExecuteNoRecords
        Cn.Execute " Update ContainerContracts set NoteID=null ,NoteSerial=null where ID=" & val(TxtSerial1(mIndex).Text)
       
        
     '   RsSavRec.Requery
        TxtNoteID = ""
        TxtNoteSerial = ""
        Dim s As String
        s = "Update TblHandWages Set NoteID = " & val(TxtNoteID) & ",NoteSerial = '" & Trim(TxtNoteSerial) & "' Where Id = " & val(TxtSerial1(mIndex))
                    
            Cn.Execute s
        End If
'
'         FindRec val(TxtSerial1(mIndex).Text)
'         TxtModFlg2(mIndex).Text = ""
'         TxtNoteSerial = ""
'          If SystemOptions.UserInterface = ArabicInterface Then
'            Msg = " „  Õ–› «·ﬁÌœ   "
'
'
'            If val(TxtNoteID) <> 0 Then
'                CmdCreateV2.Enabled = False
'                cmdPrintNote.Enabled = True
'                cmdDelNote.Enabled = True
'                btn_Save(mIndex).Enabled = False
'                btn_Modify(mIndex).Enabled = False
'             Else
'                CmdCreateV2.Enabled = True
'                cmdPrintNote.Enabled = False
'                cmdDelNote.Enabled = False
'            End If
'        Else
'            Msg = " This voucher deleted  "
'        End If
'     '   MsgBox Msg
'       End If

  


End Sub

Private Sub cmdPrintNote_Click()

ShowGL_cc Me.TxtNoteSerial.Text, , 1100

End Sub
Private Sub CBoBasedON_Change()
    If CBoBasedON.ListIndex = 0 Then
        Frame5.Visible = True
    '    lbl(20).Caption = "—ﬁ„ Õ—ﬂ… ﬁÿ⁄ «·€Ì«— "
    lbl(20).Visible = True
    TXTOrDer_no(0).Visible = True
    Else
    lbl(20).Visible = False
    TXTOrDer_no(0).Visible = False
        lbl(20).Caption = "—ﬁ„ «„— «·«’·«Õ"
        Frame5.Visible = False
    End If
    Frame5.Visible = True
    If Me.TxtModFlg2(mIndex).Text = "N" Or Me.TxtModFlg2(mIndex).Text = "E" Then
        
        If TXTOrDer_no(0).Text <> "" Then
            TXTOrDer_no(0).Text = ""
            TXTOrDer_no(1).Text = ""
        End If
        
        LoadCar
    End If
End Sub

Private Sub CBoBasedON_Click()
CBoBasedON_Change
End Sub

Private Sub grdTrans_CellButtonClick(ByVal Row As Long, ByVal Col As Long)
 With Me.grdTrans

        Select Case .ColKey(Col)

            Case "Show"
                frmsalebill.show
                frmsalebill.XPBtnMove_Click (2)
                frmsalebill.Retrive val(grdTrans.TextMatrix(Row, grdTrans.ColIndex("Transaction_ID")))
                
            End Select
    End With
End Sub

Private Sub lbl_MouseMove(Index As Integer, _
                          Button As Integer, _
                          Shift As Integer, _
                          X As Single, _
                          Y As Single)

    If val(lbl(Index).Caption) <> 0 Then
        lbl(Index).ToolTipText = WriteNo(lbl(Index).Caption, 0, True)
    End If

End Sub


Private Sub LoadCar()
    Dim StrSQL As String
    
        txtTotalInvoice = ""
                          txtTotalInvoiceBVat = ""
                          txtDiscValueInvoice = ""
                          txtVat2Invoice = ""
                          txtNetInvoice = ""
         'If CBoBasedON.ListIndex = 0 Then
                             
                     'StrSQL = "SELECT     Sum(Transaction_Details.ShowPrice * Transaction_Details.ShowQty) Total"
                     
                     StrSQL = " Select Sum(T.VatYou) VatYou,Sum(T.Vat) Vat,sum(t.ItemDiscount) ItemDiscount,  Sum(T.Trans_Discount) Trans_Discount,Sum(T.netvalue) netvalue,Sum(T.Total) Total from ("
                     StrSQL = StrSQL & " SELECT     t.VatYou, t.Vat,  t.netvalue,"
                    StrSQL = StrSQL & "                        Trans_Discount ="
                    StrSQL = StrSQL & "                   CASE  Trans_DiscountType WHEN 1 then"
                    StrSQL = StrSQL & "                     t.Trans_Discount"
                    StrSQL = StrSQL & "                 WHEN 2 THEN"
                    
                    StrSQL = StrSQL & "                     ("
                    StrSQL = StrSQL & "                       SELECT SUM("
                    StrSQL = StrSQL & "                                  Transaction_Details.ShowPrice * Transaction_Details.ShowQty"
                    StrSQL = StrSQL & "                              )"
                    StrSQL = StrSQL & "                       From Transaction_Details"
                    StrSQL = StrSQL & "                       Where dbo.Transaction_Details.Transaction_ID = t.Transaction_ID"
                    StrSQL = StrSQL & " ) * t.Trans_Discount /100 end,"
                    
                     
                    StrSQL = StrSQL & " Total =  ( SELECT SUM("
                    StrSQL = StrSQL & "                        Transaction_Details.ShowPrice * Transaction_Details.ShowQty)"
                        
                    StrSQL = StrSQL & "    From Transaction_Details"
                    StrSQL = StrSQL & " Where dbo.Transaction_Details.Transaction_ID = t.Transaction_ID),"
                    
                    
                    StrSQL = StrSQL & "    ItemDiscount = ("
                     StrSQL = StrSQL & "      SELECT SUM("
                    StrSQL = StrSQL & "                                     ("
                    StrSQL = StrSQL & "                                         Case Transaction_Details.ItemDiscountType"
                    StrSQL = StrSQL & "                                              WHEN 2 THEN ItemDiscount"
                    StrSQL = StrSQL & "                                              WHEN 3 THEN Transaction_Details.ShowPrice * Transaction_Details.ShowQty *"
                    StrSQL = StrSQL & "                                                   ItemDiscount / 100"
                    StrSQL = StrSQL & "                                         End"
                    StrSQL = StrSQL & "                                     )"
                    StrSQL = StrSQL & "                                 ) "
                    '+ SUM(Transaction_Details.TotalDiscountPerLine)"
                    StrSQL = StrSQL & "                          From Transaction_Details"
                    StrSQL = StrSQL & "                          Where dbo.Transaction_Details.Transaction_ID = t.Transaction_ID"
                    StrSQL = StrSQL & "                      )"
                     
                     StrSQL = StrSQL & " FROM         "
                     
                     StrSQL = StrSQL & "                      dbo.Transactions t "
                     StrSQL = StrSQL & " Where (t.Transaction_Type = 21) And (t.order_no = '" & val(TXTOrDer_no(0).Text) & "')"
                    StrSQL = StrSQL & " ) T"
                      Dim rsDummy As New ADODB.Recordset
                      Dim mTotal As Double
                      rsDummy.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText
                      If Not rsDummy.EOF Then
                          txtTotalInvoice = Round(val(rsDummy!Total & ""), 2)
                          txtTotalInvoiceBVat = Round(val(rsDummy!netvalue & ""), 2)
                          txtDiscValueInvoice = Round(val(rsDummy!ItemDiscount & ""), 2) + Round(val(rsDummy!Trans_Discount & ""), 2)
                          txtVat2Invoice = Round(val(rsDummy!Vat & ""), 2)
                          txtNetInvoice = Round(val(rsDummy!netvalue & "") + val(rsDummy!Vat & ""), 2)
                      End If
                   ' End If

Set Dcombos = New ClsDataCombos
Dim sql As String
  
        If val(Me.DcbType.ListIndex) = 1 Then
          If SystemOptions.UserInterface = ArabicInterface Then
              sql = " SELECT     dbo.tblordermaintenancetypes.PartID, dbo.FixedAssets.Name"
          Else
              sql = " SELECT     dbo.tblordermaintenancetypes.PartID, dbo.FixedAssets.NameE"
          End If
        sql = sql & " FROM         dbo.tblordermaintenancetypes LEFT OUTER JOIN"
        sql = sql & "                    dbo.FixedAssets ON dbo.tblordermaintenancetypes.PartID = dbo.FixedAssets.id"
        sql = sql & "  Where (dbo.tblordermaintenancetypes.OrderID = " & val(TXTOrDer_no(1).Text) & ") And (dbo.tblordermaintenancetypes.TypeTrans = 2)"
        Dcombos.ClearMyDataCombo DCEquipments
        fill_combo DCEquipments, sql
        DoEvents
        DoEvents
        GetOrderMaintdet
    Else
        Dcombos.GetEquipments DCEquipments
        GetOrderMaint
    End If
    
    
                         
        StrSQL = " Select NoteSerial1,Transaction_Date,Transaction_ID,"
        StrSQL = StrSQL & " round(sum(netvalue) + sum(Vat),2)  as NetInvoice,t.remark,"
        StrSQL = StrSQL & " Sum(T.VatYou) VatYou,round(Sum(T.Vat),2) Vat,round(Sum(T.Trans_Discount),2) + round(Sum(T.ItemDiscount ),2)Trans_Discount,  round(Sum(T.netvalue),2) Total2,round(Sum(T.Total),2) Total from ("
        StrSQL = StrSQL & " SELECT     remark,Transaction_ID,NoteSerial1,Transaction_Date, t.VatYou, t.Vat,  t.netvalue,"
        
        StrSQL = StrSQL & "                        Trans_Discount ="
        StrSQL = StrSQL & "                   CASE  Trans_DiscountType WHEN 1 then"
        StrSQL = StrSQL & "                     t.Trans_Discount"
        StrSQL = StrSQL & "                 WHEN 2 THEN"
        
        StrSQL = StrSQL & "                     ("
        StrSQL = StrSQL & "                       SELECT SUM("
        StrSQL = StrSQL & "                                  Transaction_Details.ShowPrice * Transaction_Details.ShowQty"
        StrSQL = StrSQL & "                              )"
        StrSQL = StrSQL & "                       From Transaction_Details"
        StrSQL = StrSQL & "                       Where dbo.Transaction_Details.Transaction_ID = t.Transaction_ID"
        StrSQL = StrSQL & " ) * t.Trans_Discount /100 end,"
        
        
        
        StrSQL = StrSQL & " Total =  ( SELECT SUM("
        StrSQL = StrSQL & "                        Transaction_Details.ShowPrice * Transaction_Details.ShowQty)"
        
        StrSQL = StrSQL & "    From Transaction_Details"
        StrSQL = StrSQL & " Where dbo.Transaction_Details.Transaction_ID = t.Transaction_ID),"
        StrSQL = StrSQL & " ItemDiscount =  (SELECT SUM("
        StrSQL = StrSQL & "                        ItemDiscount)"
        
        StrSQL = StrSQL & "    From Transaction_Details"
        StrSQL = StrSQL & " Where dbo.Transaction_Details.Transaction_ID = t.Transaction_ID)"
        
        
        StrSQL = StrSQL & " FROM         "
        
        StrSQL = StrSQL & "                      dbo.Transactions t "
        StrSQL = StrSQL & " Where (t.Transaction_Type = 21) And (t.order_no = '" & val(TXTOrDer_no(0).Text) & "')"
        StrSQL = StrSQL & " ) T Group By NoteSerial1,Transaction_Date,Transaction_ID,remark"

        loadgrid StrSQL, grdTrans, True, False
        
    txtGeneralTotal = val(txtTotalInvoice) + val(txtTotal2)
    txtTotalDisc = val(txtDiscValueInvoice) + val(txtDiscValue)
    txtTotalVat = val(txtVat2Invoice) + val(txtVat2)
    txtTotalBVat = val(txtTotal) + val(txtTotalInvoiceBVat)
    txtTotalNet = val(txtNetInvoice) + val(txtNet)
    
 
End Sub



Sub GetOrderMaint()
If 1 = 1 Then
Dim sql As String
Dim rs2 As ADODB.Recordset
Set rs2 = New ADODB.Recordset
sql = " SELECT     EquepID"
sql = sql & " From dbo.TblOrderMaint"
sql = sql & "  where ID =" & val(TXTOrDer_no(1).Text) & ""
rs2.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
If rs2.RecordCount > 0 Then
DCEquipments.BoundText = IIf(IsNull(rs2("EquepID").value), "", rs2("EquepID").value)
Else
DCEquipments.BoundText = 0
End If
End If
End Sub


Sub GetOrderMaintdet()
 End Sub





Private Sub DcbType_Change()
LoadCar
End Sub

Private Sub DcbType_Click()
DcbType_Change
End Sub

Private Sub DCEquipments_KeyUp(KeyCode As Integer, Shift As Integer)
  'www
   If KeyCode = vbKeyF3 Then
         Load FrmCasrShearches
        FrmCasrShearches.SendForm = "FrmOut"
        FrmCasrShearches.show vbModal
    End If
End Sub


Private Sub Cmd_DeleteAll_Click(Index As Integer)
If mIndex = 1 Then
    If Me.TxtModFlg2(mIndex).Text <> "R" Then
    
            fg.Rows = 1
            fg.Rows = 2

    
        
    End If
ElseIf mIndex = 2 Then
End If

End Sub

Private Sub Cmd_DeleteRow_Click(Index As Integer)
If Me.TxtModFlg.Text <> "R" Then
If Index = 1 Then
    

    RemoveGridRow
End If


End If
End Sub
Private Sub RemoveGridRow()

    With Me.fg
'MsgBox .Row
        If .Row <= 0 Then
                .Rows = 2
        Exit Sub
        Else
        .RemoveItem .Row
        End If
    End With
End Sub

Private Sub btn_Cancel_Click(Index As Integer)

   Unload Me
End Sub

Private Sub btn_Delete_Click(Index As Integer)
   Dim MSGType As Integer
   Dim StrSQL As String
    Dim BlnRecExist As Boolean
    Dim StrMSG  As String
    Dim Msg As String
    On Error GoTo ErrTrap
    
If Index = 3 Then

'If TxtNoteSerial <> "" Then
'MsgBox "·« Ì„ﬂ‰ «·Õ–› «Ê «· ⁄œÌ· «·« »⁄œ Õ–› «·ﬁÌœ"
'Exit Sub
End If





If Index = 1 Then
    If TxtNoteSerial <> "" Then
        cmdDelNote_Click
    End If
End If
    'Index = TabMain.CurrTab
    'If DoPremis(Do_Delete, Me.name, True) = False Then
    '    Exit Sub
    'End If
    If TxtSerial1(mIndex).Text <> "" Then
        '    If CheckDelCountry(Val(Me.TxtVac_ID.text)) = False Then
        '        Msg = "·«Ì„ﬂ‰ Õ–› Â–« «·”Ã·...!!!"
        '        MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
        '        Exit Sub
        '    End If
        If SystemOptions.UserInterface = ArabicInterface Then
        MSGType = MsgBox("Â·  —€» ›Ì Õ–› Â–« «·”Ã·", vbQuestion + vbMsgBoxRtlReading + vbYesNo + vbMsgBoxRight, App.title)
        Else
        MSGType = MsgBox("Confirm Delete", vbQuestion + vbMsgBoxRtlReading + vbYesNo + vbMsgBoxRight, App.title)
        End If

        If MSGType = vbYes Then
            RsSavRec.Find "id=" & val(TxtSerial1(mIndex).Text), , adSearchForward, 1
           ' CuurentLogdata ("D")
            RsSavRec.delete
            Dim s As String
            If mIndex = 1 Then
                s = " Delete From TblHandWages2 Where MasterID = " & val(TxtSerial1(mIndex).Text)
                Cn.Execute s
            End If
            
            
            If SystemOptions.UserInterface = ArabicInterface Then
            MsgBox " „  ⁄„·Ì… «·Õ–› »‰Ã«Õ", vbInformation + vbMsgBoxRtlReading + vbOKOnly + vbMsgBoxRight, App.title
            Else
            MsgBox "Delete Successfully", vbInformation + vbMsgBoxRtlReading + vbOKOnly + vbMsgBoxRight, App.title
            End If
            '------------------------------ Move Next ---------------------------.
            
            btn_Next_Click mIndex
        End If
    End If

    Exit Sub
ErrTrap:
 
    Select Case Err.Number

        Case -2147217873, -2147467259
        If SystemOptions.UserInterface = ArabicInterface Then
            StrMSG = "⁄›Ê« ·« ÌÃÊ“ Õ–› «·”Ã· ·«— »«ÿÂ »»Ì«‰«  √Œ—Ì"
            Else
            StrMSG = "Sorry you can not delete the record of its connection with other data"
            End If
            RsSavRec.CancelUpdate
            MsgBox StrMSG, vbOKOnly + vbMsgBoxRight, App.title
            'clear the ConnectiOn Errors
            Cn.Errors.Clear
    End Select

End Sub

Private Sub btn_Modify_Click(Index As Integer)
    Dim Msg As String

    If DoPremis(Do_Edit, Me.Name, True) = False Then
        Exit Sub
    End If

    On Error GoTo ErrTrap

If Index = 1 Then

'    If TxtNoteSerial <> "" Then
'        MsgBox "·« Ì„ﬂ‰ «·Õ–› «Ê «· ⁄œÌ· «·« »⁄œ Õ–› «·ﬁÌœ"
'        Exit Sub
'    End If


  
 

      Set rsDummy = New ADODB.Recordset
      s = "select * from TblCardAuthorizationReform where WorkOrder = " & val(TXTOrDer_no(0).Text) & " "
      rsDummy.Open s, Cn, adOpenStatic, adLockOptimistic, adCmdText
      If Not rsDummy.EOF Then
          If val(rsDummy!IsEndAll & "") <> 0 Then
               If SystemOptions.UserInterface = ArabicInterface Then
                  MsgBox "·« Ì„ﬂ‰ «· ⁄œÌ· ⁄·Ï Â–Â «·›« Ê—… ·ÊÃÊœ «„— «’·«Õ  „ «‰Â«∆Â"
              Else
                  MsgBox "This invoice cannot be modified due to a repair order that has been terminated"
              End If
              Exit Sub
          End If
      End If
 

Frame1(2).Enabled = True
    If TxtSerial1(mIndex).Text <> "" Then
   '     TxtModFlg2(mIndex) = "E"
    
        Frm2.Enabled = True
        
        DcCustmer(mIndex).SetFocus
    End If
    
End If
    If TxtSerial1(mIndex).Text <> "" Then
        TxtModFlg2(mIndex) = "E"
        Frame1(2).Enabled = True
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
            Msg = "Sorry" & CHR(13)
            Msg = Msg & " You can not edit this record now" & CHR(13)
            Msg = Msg & "Where it was being edited by another user on the network"
           End If
            MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading + vbOKOnly, App.title
        
            If RsSavRec.EditMode <> adEditNone Then
                RsSavRec.CancelUpdate
                'RsSavRec.Requery
            End If

    End Select

End Sub

Public Sub btn_New_Click(Index As Integer)
    Dim My_SQL As String
    Dim rs As ADODB.Recordset

    If DoPremis(Do_New, Me.Name, True) = False Then
        Exit Sub
    End If

    On Error GoTo ErrTrap
    Set rs = New ADODB.Recordset
    Frm2.Enabled = True
    Frame1(2).Enabled = True
    TXTOrDer_no(0) = ""
    TXTOrDer_no(1) = ""
    clear_all Me
    TxtModFlg2(mIndex).Text = "N"
    If mIndex = 1 Then
        My_SQL = "ContainerContracts"
        DCboUserName(mIndex).BoundText = user_id
        Dcbranch(1).BoundText = branch_id
            
        fg.Rows = 1
        fg.Rows = 2
      
   ElseIf mIndex = 2 Then
        My_SQL = "TblOffice"
   ElseIf mIndex = 3 Then
        My_SQL = "TblLensesTypes"
        
     
   
   
    End If
'    rs.Open My_SQL, Cn, adOpenKeyset, adLockOptimistic, adCmdTable
'
''    If rs.RecordCount > 0 Then
''        TxtSerial1(mIndex).Text = rs.RecordCount + 1
''    Else
''        TxtSerial1(mIndex).Text = 1
''    End If
'
'    rs.Close
    'CmbType.ListIndex = 0
    'TxtVacName.SetFocus
ErrTrap:

End Sub

Private Sub Btn_Print_Click(Index As Integer)
  If mIndex = 1 Or mIndex = 2 Or mIndex = 4 Then
    
    print_report
   ' PrintRercord
ElseIf mIndex = 3 Then
    
End If
End Sub



Private Sub PrintRercord()
  Dim MySQL As String
    Dim RsData As New ADODB.Recordset
    Dim xApp As New CRAXDRT.Application
    Dim xReport As CRAXDRT.Report
    Dim CViewer As ClsReportViewer
    Dim StrReportTitle As String
    Dim StrFileName As String
    Dim Msg As String
 
'--------------------------------------------------------------------------------------------
   
   

    
    
 
        
MySQL = " SELECT TT.ID,"
MySQL = MySQL & "         TT.NoteSerial1,TblHandWages2.Price,TblHandWages2.Name,"
MySQL = MySQL & "         TblCardAuthorizationReform.CusID,TT.Total,TT.Total2,TT.VatYou,TT.Vat2,TT.Net,TT.OrDer_no2,"
MySQL = MySQL & "         TT.OrDer_no,TT.DiscValue,TT.DiscPercent,"
MySQL = MySQL & "         TT.RecordDate,"
MySQL = MySQL & "         TT.BranchID,"
MySQL = MySQL & "'" & CBoBasedON.Text & "' as CBoBasedON,  "
MySQL = MySQL & "'" & DcbCarType.Text & "' as DcbCarType,  "
MySQL = MySQL & "'" & DcbyearFactor.Text & "' as DcbyearFactor,  "
MySQL = MySQL & "'" & DCEquipments.Text & "' as DCEquipments,  "
MySQL = MySQL & "         TblCardAuthorizationReform.ClientName,TblCardAuthorizationReform.Shaseh,TblCardAuthorizationReform.CarMeter,PlateNo,"
MySQL = MySQL & "         b.branch_name"
MySQL = MySQL & "  FROM   TblHandWages TT"
MySQL = MySQL & "         LEFT OUTER JOIN TblBranchesData AS b"
MySQL = MySQL & "              ON  TT.BranchID = b.branch_id"
MySQL = MySQL & "         LEFT OUTER JOIN TblCardAuthorizationReform"
MySQL = MySQL & "              ON  TblCardAuthorizationReform.WorkOrder = tt.OrDer_no2"
MySQL = MySQL & "              LEFT OUTER JOIN TblHandWages2 ON TblHandWages2.MasterID = TT.ID"
MySQL = MySQL & "  Where 1 = 1"
MySQL = MySQL & "         AND (NOT (TT.ID IS NULL))"
        
MySQL = MySQL & "  And (TT.ID =" & val(TxtSerial1(mIndex).Text) & ")"
   
            If SystemOptions.UserInterface = ArabicInterface Then
                StrFileName = App.path & "\REPORTS\REPORTS NEW\" & "TblHandWages.rpt"
            Else
                StrFileName = App.path & "\REPORTS\REPORTS NEW\" & "TblHandWages.rpt"
    
            
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
        If SystemOptions.UserInterface = ArabicInterface Then
        Msg = "·« ÊÃœ »Ì«‰«  ··⁄—÷"
        Else
        Msg = "No Data"
        End If
        MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
        RsData.Close
        Set RsData = Nothing
        Screen.MousePointer = vbDefault
        Exit Sub
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
      '  xReport.ParameterFields(2).AddCurrentValue RPTComment_Eng
        'xReport.ParameterFields(4).AddCurrentValue get_branch_name(val(my_branch))
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

End Sub

Private Sub btn_Query_Click(Index As Integer)

If mIndex = 1 Then
   FrmProjectSearch.Indx = 7
    FrmProjectSearch.Indx2 = 8
    FrmProjectSearch.C1Tab1.CurrTab = 7
    FrmProjectSearch.C1Tab1.TabVisible(6) = False
    FrmProjectSearch.C1Tab1.TabCaption(7) = Me.Caption
    FrmProjectSearch.Caption = Me.Caption
 
    FrmProjectSearch.show vbModal
End If
End Sub

Private Sub btn_Save_Click(Index As Integer)
   On Error GoTo ErrTrap
    Dim Msg As String
    Dim StrVacCode As String
    Dim StrVacName As String
    Dim CtrlTxt As Control
    Dim i As Long
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

If mIndex < 2 Then
    If SystemOptions.UserInterface = ArabicInterface Then
        If TxtCashCustomerName.Text = "" Then
            MsgBox "Ì—ÃÏ «œŒ«· «·⁄„Ì·"
            DcCustmer(mIndex).SetFocus
            Exit Sub
        End If
    Else
        If DcCustmer(mIndex).Text = "" Then
            MsgBox "Please Enter Name"
            DcCustmer(mIndex).SetFocus
            Exit Sub
        End If
    End If
End If
If mIndex = 1 Then
    If val(txtTotal2) = 0 Then
        If SystemOptions.UserInterface = ArabicInterface Then
           MsgBox " Õœœ «·„»·€ «Ê·«"
        Else
            MsgBox " Please Enter amount "
        End If
        txtTotal2.SetFocus
         Exit Sub
    End If
    
    Dim StrSQL  As String
    
         If SystemOptions.MaintOrderCantRepeatBillBuy Then
            Dim rs2 As New ADODB.Recordset
            

            StrSQL = "SELECT NoteSerial1,OrDer_no2   FROM TblHandWages where  IsNull(OrDer_no,0)  = '" & val(TXTOrDer_no(0).Text) & "' and Id <> " & val(TxtSerial1(mIndex).Text)
            rs2.Open StrSQL, Cn, adOpenForwardOnly, adLockReadOnly
            If Not rs2.EOF Then
                MsgBox "Â–« «·«„— ·« Ì„ﬂ‰ «œ—«ÃÂ ›ﬁœ «œ—Ã „‰ ﬁ»· ›Ï «·›« Ê—… —ﬁ„" & rs2!NoteSerial1 & ""
                TXTOrDer_no2 = ""
                TXTOrDer_no(0) = ""
                TXTOrDer_no(1).Text = ""
                
               ' Cmd(2).Enabled = True
              
                Exit Sub
            End If
        End If
   
    

End If


   
    '------------------------------ check if Empcode exist ----------------------

   

    ' -------------------------------------- txtmodflg type -------------------
    Select Case TxtModFlg2(mIndex).Text

            '------------------------------ new record ----------------------------
        Case "N"
      
            '------------------------- save record -----------------------------
            
            If mIndex = 1 Then
                FiLLRec1
            ElseIf mIndex = 2 Then
                AddNewRec
               FiLLRec2
            ElseIf mIndex = 3 Then
              '  AddNewRec
               FiLLRec3
               
            End If
            If mIndex = 0 Then
                BtnLast_Click
            Else
               
                btn_Last_Click CInt(mIndex)
            End If

        Case "E"

            '----------------------------- save edit -------------------------------
            If mIndex = 0 Then
                FiLLRec
            ElseIf mIndex = 1 Then
                FiLLRec1
          ElseIf mIndex = 2 Then
                FiLLRec2
          ElseIf mIndex = 3 Then
                FiLLRec3
          
            End If
    End Select

    Exit Sub
ErrTrap:
If SystemOptions.UserInterface = ArabicInterface Then
    MsgBox "Â‰«ﬂ Œÿ√ „« ›Ì ≈œŒ«· «·»Ì«‰« ", vbOKOnly + vbMsgBoxRight, App.title
 Else
  MsgBox "Sorry...error douring insert data", vbOKOnly + vbMsgBoxRight, App.title
End If
 
End Sub

Private Sub Btn_Undo_Click(Index As Integer)
    Undo
End Sub
Private Sub Undo()
    On Error GoTo ErrTrap

    Select Case TxtModFlg2(mIndex).Text

        Case "N"
            clear_all Me
            TxtModFlg2(mIndex).Text = "R"
           
            btn_First_Click (mIndex)
        Case "E"
            RsSavRec.Find "ID='" & val(TxtSerial1(mIndex).Text) & "'", , adSearchForward, adBookmarkFirst

            If RsSavRec.EOF Or RsSavRec.BOF Then
                TxtModFlg2(mIndex).Text = "R"
                Exit Sub
            End If

            If mIndex = 1 Then
                FiLLTXT1
            ElseIf mIndex = 2 Then
                FiLLTXT2
            ElseIf mIndex = 3 Then
                FiLLTXT3
            End If
            TxtModFlg2(mIndex).Text = "R"
    End Select

    Exit Sub
ErrTrap:
End Sub
  Private Sub GetCardAuthorizationData()
   DcbyearFactor.Text = ""
            TxtPlatNo = ""
            DcbCarType.BoundText = ""
            
            TxtManualNo2(2).Text = """"
             TxtManualNo2(1).Text = ""
  If val(TXTOrDer_no(1)) <> 0 Then
  
        Dim rs2 As New ADODB.Recordset
        Dim orderStatus As Integer
        Dim StrSQL As String
        MintDone = 0
    
        Set rs2 = New ADODB.Recordset
        StrSQL = "select * from TblCardAuthorizationReform where WorkOrder = " & val(TXTOrDer_no(0).Text) & " "
        rs2.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText
        If rs2.RecordCount > 0 Then
            orderStatus = IIf(IsNull(rs2("OrderStatus").value), 0, rs2("OrderStatus").value)
            TxtCashCustomerName.Text = IIf(IsNull(rs2("ClientName").value), "", rs2("ClientName").value)
            'DCOPrType =
                  
                  
                  
            
            DcbyearFactor.Text = val(rs2!YearFact & "")
            TxtPlatNo = Trim(rs2!PlateNo & "")
            DcbCarType.BoundText = val(rs2!CarTypeID & "")
            DcbCarModel.BoundText = IIf(IsNull(rs2("CarModelID").value), "", rs2("CarModelID").value)
            TxtManualNo2(2).Text = Trim(rs2!Shaseh & "")
             TxtManualNo2(1).Text = Trim(rs2!CarMeter & "")
               DcCustmer(mIndex).BoundText = val(rs2!CusID & "")
                If val(rs2!CusID & "") = 0 Then
                    StrSQL = "SELECT tc.CusID FROM TblCustemers AS tc WHERE tc.CusName LIKE N'%" & Trim(TxtCashCustomerName.Text) & "%'"
                    Dim rsDummy As New ADODB.Recordset
                    rsDummy.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText
                    If Not rsDummy.EOF Then
                        DcCustmer(mIndex).BoundText = val(rsDummy!CusID & "")
                    End If
                End If
                  
        
                  
                  
        End If
        If orderStatus = 2 Or orderStatus = 4 Or orderStatus = 5 Then
            MintDone = 1
        End If
                If CBoBasedON.ListIndex = 0 Then
                             
                     'StrSQL = "SELECT     Sum(Transaction_Details.ShowPrice * Transaction_Details.ShowQty) Total"
                     StrSQL = "SELECT     t.VatYou, t.Vat, t.Trans_Discount, t.netvalue,"
                     
                    StrSQL = StrSQL & " Total =   SELECT SUM("
                    StrSQL = StrSQL & "                        Transaction_Details.ShowPrice * Transaction_Details.ShowQty)"
                        
                    StrSQL = StrSQL & "    From Transaction_Details"
                    StrSQL = StrSQL & " Where dbo.Transaction_Details.Transaction_ID = t.Transaction_ID),"
                    StrSQL = StrSQL & " ItemDiscount =   SELECT SUM("
                    StrSQL = StrSQL & "                        ItemDiscount)"
                        
                    StrSQL = StrSQL & "    From Transaction_Details"
                    StrSQL = StrSQL & " Where dbo.Transaction_Details.Transaction_ID = t.Transaction_ID)"
                        
                     
                     StrSQL = StrSQL & " FROM         "
                     StrSQL = StrSQL & "                      dbo.Transaction_Details RIGHT OUTER JOIN"
                     StrSQL = StrSQL & "                      dbo.Transactions t ON dbo.Transaction_Details.Transaction_ID = t.Transaction_ID"
                     StrSQL = StrSQL & " Where (t.Transaction_Type = 21) And (t.NoteSerial1 = '" & val(TXTOrDer_no(0).Text) & "')"
                          
                      Set rsDummy = New ADODB.Recordset
                      Dim mTotal As Double
                   '   rsDummy.Open strSql, Cn, adOpenStatic, adLockReadOnly, adCmdText
                   '   If Not rsDummy.EOF Then
                   '       txtTotalInvoice = val(rsDummy!Total & "")
                   '       txtDiscValueInvoice = val(rsDummy!ItemDiscount & "")
                   '       txtVat2Invoice = val(rsDummy!Vat & "")
                   '       txtNetInvoice = val(rsDummy!netvalue & "")
                   '   End If
                    End If
                  
        
    End If
End Sub
Public Sub FiLLTXT1(Optional Lngid As Long = 0)

    On Error GoTo ErrTrap
    
       Dim RsDetails As New ADODB.Recordset
    Dim StrSQL As String
    Dim RsNotes As New ADODB.Recordset
    Dim RsTest  As ADODB.Recordset
    Dim RsReplace As ADODB.Recordset
    Dim LngPartID As Long
    Dim RsPartDetails As ADODB.Recordset
    Dim i As Long

     If Lngid <> 0 Then
        RsSavRec.Find "ID=" & Lngid, , adSearchForward, adBookmarkFirst

        If RsSavRec.BOF Or RsSavRec.EOF Then
            Exit Sub
        End If
    End If

    On Error GoTo ErrTrap

    'Frm2.Enabled = False
    TxtSerial1(mIndex).Text = IIf(IsNull(RsSavRec.Fields("id").value), "", RsSavRec.Fields("id").value)
    XPDtbTrans.value = IIf(IsNull(RsSavRec("RecordDate").value), Date, RsSavRec("RecordDate").value)
    Me.TxtNoteSerial1.Text = IIf(IsNull(RsSavRec("NoteSerial1").value), "", RsSavRec("NoteSerial1").value)
  txtDiscValue = IIf(IsNull(RsSavRec("DiscValue").value), "", RsSavRec("DiscValue").value)
    txtDiscPercent = IIf(IsNull(RsSavRec("DiscPercent").value), "", RsSavRec("DiscPercent").value)
    
    
    txtTotal2 = IIf(IsNull(RsSavRec("Total2").value), "", RsSavRec("Total2").value)
    txtVat2 = IIf(IsNull(RsSavRec("Vat2").value), "", RsSavRec("Vat2").value)
    txtVatYou = IIf(IsNull(RsSavRec("VatYou").value), "", RsSavRec("VatYou").value)
    txtNet = IIf(IsNull(RsSavRec("Net").value), "", RsSavRec("Net").value)
    
    
   CBoBasedON.ListIndex = IIf(IsNull(RsSavRec("CBoBasedON").value), -1, RsSavRec("CBoBasedON").value)
    TXTOrDer_no(0) = IIf(IsNull(RsSavRec("OrDer_no").value), "", RsSavRec("OrDer_no").value)
    TXTOrDer_no(1) = IIf(IsNull(RsSavRec("OrDer_no2").value), "", RsSavRec("OrDer_no2").value)
    
    TXTOrDer_no2 = IIf(IsNull(RsSavRec("RowsEstimatedID").value), "", RsSavRec("RowsEstimatedID").value)
    

   
    Dcbranch(mIndex).BoundText = IIf(IsNull(RsSavRec("BranchID").value), "", RsSavRec("BranchID").value)
    TxtRemarks = IIf(IsNull(RsSavRec("Remarks").value), "", RsSavRec("Remarks").value)
    Me.DCboUserName(1).BoundText = IIf(IsNull(RsSavRec("UserID").value), "", RsSavRec("UserID").value)

    
     
    
 
     
    
    
    GetCardAuthorizationData
    
'      TxtNoteID = RsSavRec!NoteID & ""
'    TxtNoteSerial = RsSavRec!NoteSerial & ""
    LoadCar
     TxtNoteID = RsSavRec!NoteID & ""
    TxtNoteSerial = RsSavRec!NoteSerial & ""
    
     If val(TxtNoteID) <> 0 Then
        CmdCreateV2.Enabled = False
        cmdPrintNote.Enabled = True
        cmdDelNote.Enabled = True

     Else
        CmdCreateV2.Enabled = True
        cmdPrintNote.Enabled = False
        cmdDelNote.Enabled = False

    End If
    LabCurr_Rec(mIndex).Caption = RsSavRec.AbsolutePosition
    LabCount_Rec(mIndex).Caption = RsSavRec.RecordCount

    Dim s As String
    
    





            
    
    
    s = " SELECT TblHandWages2.*,TblEmpDepartments.DepartmentName,TblEmpDepartments.DepartmentNamee "
    
    s = s & " from TblHandWages2 Left Outer Join TblEmpDepartments On TblHandWages2.DeparmentID =TblEmpDepartments.DeparmentID  "
    s = s & " Where MasterID = " & val(TxtSerial1(mIndex))
    
    loadgrid s, fg, True, True
CalcTotal2
ErrTrap:

End Sub

 Public Sub FiLLTXT3(Optional Lngid As Long = 0)

    On Error GoTo ErrTrap
    
       Dim RsDetails As New ADODB.Recordset
    Dim StrSQL As String
    Dim RsNotes As New ADODB.Recordset
    Dim RsTest  As ADODB.Recordset
    Dim RsReplace As ADODB.Recordset
    Dim LngPartID As Long
    Dim RsPartDetails As ADODB.Recordset
    Dim i As Long

     If Lngid <> 0 Then
        RsSavRec.Find "ID=" & Lngid, , adSearchForward, adBookmarkFirst

        If RsSavRec.BOF Or RsSavRec.EOF Then
            Exit Sub
        End If
    End If

    On Error GoTo ErrTrap

    'Frm2.Enabled = False
    TxtSerial1(mIndex).Text = IIf(IsNull(RsSavRec.Fields("id").value), "", RsSavRec.Fields("id").value)
    'XPDtbTrans.value = IIf(IsNull(RsSavRec("RecordDate").value), Date, RsSavRec("RecordDate").value)
    'Me.TxtNoteSerial1.Text = IIf(IsNull(RsSavRec("NoteSerial1").value), "", RsSavRec("NoteSerial1").value)
   
   cmbFlag(0).ListIndex = IIf(IsNull(RsSavRec("Flag").value), -1, RsSavRec("Flag").value)

   
    cmbGroupId.BoundText = IIf(IsNull(RsSavRec("GroupId").value), "", RsSavRec("GroupId").value)
    cmbUnitID.BoundText = IIf(IsNull(RsSavRec("UnitId").value), "", RsSavRec("UnitId").value)
    
   
    DCBoMain(2).BoundText = IIf(IsNull(RsSavRec("FromSPH").value), "", RsSavRec("FromSPH").value)
    DCBoMain(5).BoundText = IIf(IsNull(RsSavRec("TOSPH").value), "", RsSavRec("TOSPH").value)
    DCBoMain(3).BoundText = IIf(IsNull(RsSavRec("FROMCYL").value), "", RsSavRec("FROMCYL").value)
    DCBoMain(6).BoundText = IIf(IsNull(RsSavRec("TOCYL").value), "", RsSavRec("TOCYL").value)
        
    
    
     
   txtPrice = IIf(IsNull(RsSavRec("Price").value), "", RsSavRec("Price").value)
    txtName(mIndex).Text = IIf(IsNull(RsSavRec.Fields("name").value), "", RsSavRec.Fields("name").value)
    txtNamee(mIndex).Text = IIf(IsNull(RsSavRec.Fields("nameE").value), "", RsSavRec.Fields("nameE").value)
    
    LabCurr_Rec(mIndex).Caption = RsSavRec.AbsolutePosition
    LabCount_Rec(mIndex).Caption = RsSavRec.RecordCount


     
    
    
    LabCurr_Rec(mIndex).Caption = RsSavRec.AbsolutePosition
    LabCount_Rec(mIndex).Caption = RsSavRec.RecordCount

    Dim s As String
    
    
    
    

    With Grid3

        For i = 1 To .Rows - 1

            If Trim(TxtSerial1(mIndex).Text) = .TextMatrix(i, .ColIndex("id")) Then
                TxtSerial1(mIndex).Text = .TextMatrix(i, .ColIndex("Ser"))
                .Row = i
                Exit Sub
            End If

        Next

    End With
 
GenreateItems


            
'
'
'    s = " SELECT TblHandWages2.*,TblEmpDepartments.DepartmentName,TblEmpDepartments.DepartmentNamee "
'
'    s = s & " from TblHandWages2 Left Outer Join TblEmpDepartments On TblHandWages2.DeparmentID =TblEmpDepartments.DeparmentID  "
'    s = s & " Where MasterID = " & val(TxtSerial1(mIndex))
'
'    loadgrid s, fg, True, True

ErrTrap:

End Sub

 

Private Sub Fg_AfterEdit(ByVal Row As Long, ByVal Col As Long)
    



    Dim StrAccountCode As String
    Dim Msg As String
    Dim rs As New ADODB.Recordset
    Dim StrSQL As String
    Dim ClsAcc As New ClsAccounts
    Dim LngRow As Long
Dim StrComboList As String
    'Dim Rs2 As ADODB.Recordset
  
On Error GoTo ErrTrap
    With fg
     Select Case .ColKey(Col)
    
           Case "DepartmentName"
                StrAccountCode = .ComboData
            
                LngRow = .FindRow(StrAccountCode, .FixedRows, .ColIndex("DepartmentName"), False, True)
                .TextMatrix(Row, .ColIndex("DeparmentID")) = StrAccountCode
        
 Case "Name"
       If .TextMatrix(Row, .ColIndex("DeparmentID")) = "" Then
        If Row - 1 > 1 Then
            .TextMatrix(Row, .ColIndex("DeparmentID")) = .TextMatrix(Row - 1, .ColIndex("DeparmentID"))
            .TextMatrix(Row, .ColIndex("DepartmentName")) = .TextMatrix(Row - 1, .ColIndex("DepartmentName"))
        End If
       End If
       
     Case "FixedAssetsName2"
      
    Case "FixedAssetsName3"
     
 Case "EmpName"
    
    Case "FromDate", "ToDate"
   
       
   Case "NoDays"
     'Fg.TextMatrix(Row, Fg.ColIndex("FromDate")) = StrDate
  
    End Select
    
    CalcTotal2
    
    End With
ErrTrap:
End Sub


Private Sub FG_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
With fg

   Select Case .ColKey(Col)
        Case "Name", "Rem", "Price", "Discount", "Total", "Vat", "VatValue"
            .ComboList = ""
        Case "NoteNo"
            .ComboList = ""
        Case "DayMeter"
            .ComboList = ""
        Case "CustName", "Total"
            Cancel = True
        End Select
        
    End With
End Sub

Private Sub CalcTotal2()
    
   With fg
    .IsSubtotal(.Rows - 1) = True
    Dim SngTotal As Single
    If .Rows > 1 Then
        txtTotal2 = .Aggregate(flexSTSum, .FixedRows, .ColIndex("Price"), .Rows - 1, .ColIndex("Price"))
    End If
    

    
    End With
    CalculteValueAdded2
    txtNetInvoice2 = Round(val(txtTotal) + val(txtVat2), 2)
    txtGeneralTotal = val(txtTotalInvoice) + val(txtTotal2)
    txtTotalDisc = val(txtDiscValueInvoice) + val(txtDiscValue)
    txtTotalVat = val(txtVat2Invoice) + val(txtVat2)
    txtTotalBVat = val(txtTotal) + val(txtTotalInvoiceBVat)
    txtTotalNet = val(txtNetInvoice) + val(txtNet)
End Sub


Public Sub CalculteValueAdded2(Optional posDelete As Boolean = False)

txtTotal = val(txtTotal2) - val(txtDiscValue)

txtNet = val(txtTotal) + val(txtVat2)
If SystemOptions.PriceWithVAT = True Then Exit Sub
'If (TxtModFlg2(mIndex).Text = "R" Or TxtModFlg2(mIndex).Text = "") Then Exit Sub
 Dim Percentg As Double
'If val(txtVatYou) = 0 Then Percentg = 5: txtVatYou = Percentg Else Percentg = val(txtVatYou)

Dim cCompanyInfo As New ClsCompanyInfo
If mdifrmmain.taxes.Visible = True Then
'If TransType = 9 And ReturnSales = True Then

    Dim AccountVATCreit As String
    
    If SystemOptions.AllItemInVAT = True Then
        Percentg = val(cCompanyInfo.VATItems)
    Else
      PercentgValueAddedAccount_Transec XPDtbTrans.value, 10, 1, AccountVATCreit, Percentg

    End If
    txtVatYou = Percentg
    If Percentg = -1 Then
        Percentg = 0
    Else

    End If
   
     txtVat2 = val(txtTotal) * Percentg / 100
     
     
     txtNet = val(txtTotal) + val(txtVat2)
    

End If

End Sub


 

Private Sub Fg_KeyDown(KeyCode As Integer, Shift As Integer)
'GridKeyDown Fg, KeyCode, Shift, False, False, Fg.Row
 'GridKeyDown Fg, KeyCode, Shift
 mGridClicked = True
        Dim mOldRow As Long
        mOldRow = fg.Row
     GridKeyDown fg, KeyCode, Shift, False, False, fg.Row
     If mOldRow <> fg.Row And fg.Row <> 1 Then
        fg.TextMatrix(fg.Row, fg.ColIndex("DeparmentID")) = fg.TextMatrix(fg.Row - 1, fg.ColIndex("DeparmentID"))
        fg.TextMatrix(fg.Row, fg.ColIndex("DepartmentName")) = fg.TextMatrix(fg.Row - 1, fg.ColIndex("DepartmentName"))
     End If
   '  mGridClicked = False
End Sub

Private Sub fg_KeyDownEdit(ByVal Row As Long, ByVal Col As Long, KeyCode As Integer, ByVal Shift As Integer)
mGridClicked = True
Fg_KeyDown KeyCode, 0
'mGridClicked = False
End Sub

Private Sub fg_StartEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
  Dim rs As New ADODB.Recordset
    Dim StrSQL  As String
    Dim StrAccountType As String
    Dim StrComboList As String
    Dim Msg As String
    Dim StrAccountCode As String
    Dim StrAccountCode1 As String

    Dim ClsAcc As New ClsAccounts
    Dim LngRow As Long
   
    With fg

        Select Case .ColKey(Col)
              
        
             
            Case "DepartmentName"
                .TextMatrix(Row, .ColIndex("DepartmentName")) = ""
                StrSQL = "SELECT DeparmentID,DepartmentName  FROM TblEmpDepartments "
                rs.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

                If SystemOptions.UserInterface = ArabicInterface Then
                    StrComboList = fg.BuildComboList(rs, "DepartmentName", "DeparmentID")
                Else
                    StrComboList = fg.BuildComboList(rs, "DepartmentNamee", "DeparmentID")
                End If
       
                If StrComboList <> "" Then
                    StrComboList = "|" & StrComboList
                End If
                 .ComboList = StrComboList
          Case "GroupName"
   
                 
          Case "GroupName2"
  
          Case "FixedAssetsName"
 
          Case "FixedAssetsName2"
                 
                 
          End Select
        End With
End Sub




Private Sub Dcbranch_Change(Index As Integer)
    If Me.TxtModFlg2(mIndex) <> "R" Then
        TxtNoteSerial1.Text = ""
        TxtNoteSerial.Text = ""
   End If
End Sub

Private Sub Dcbranch_Click(Index As Integer, Area As Integer)
    If Me.TxtModFlg2(mIndex) <> "R" Then
    TxtNoteSerial1.Text = ""
   TxtNoteSerial.Text = ""
   End If
   
End Sub

 
 

Private Sub Contract_period_no_KeyPress(KeyAscii As Integer)
    
End Sub
Private Sub Contract_period_no_Change()
 

End Sub

Private Sub cmdopenFrmCust_Click()
    'Dim frm As FrmCustemers
    OpenScreen CustomersScreen, DcCustmer(mIndex).BoundText
End Sub

Private Sub DcCustmer_Change(Index As Integer)
DcCustmer_Click Index, 0
End Sub

Private Sub DcCustmer_Click(Index As Integer, Area As Integer)
  If val(DcCustmer(Index).BoundText) = 0 Then Exit Sub
  

    Dim EmpCode  As String
    GetTblCustemersCode , , DcCustmer(Index).BoundText, EmpCode
    'Me.txtTotalValue.Text = EmpCode
    
If Me.TxtModFlg2(mIndex).Text <> "R" Then
If val(DcCustmer(Index).BoundText) <> 0 Then



End If
End If
End Sub

Private Sub DcCustmer_KeyUp(Index As Integer, KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyF3 Then
        Dim Frm As New FrmCustemerSearch
        Frm.SearchType = 604
        Frm.RetrunType = 0
        Frm.mIndex = Index
        Frm.show vbModal
    End If
End Sub

Private Sub StrDate_Change()

End Sub

Private Sub txtDiscPercent_Change()
 If Me.TxtModFlg2(mIndex) = "R" Then Exit Sub
If mDiscEnter Then Exit Sub
If val(txtTotal2) <> 0 Then

    txtDiscValue = val(txtDiscPercent) * val(txtTotal2) / 100

End If

CalcTotal2
mDiscEnter = False
End Sub

Private Sub txtDiscValue_Change()
 If Me.TxtModFlg2(mIndex) = "R" Then Exit Sub
mDiscEnter = True
If val(txtTotal2) <> 0 Then
    txtDiscPercent = Round(val(txtDiscValue) / val(txtTotal2) * 100, 2)
End If

CalcTotal2
mDiscEnter = False
End Sub

Private Sub TXTOrDer_no_Validate(Index As Integer, Cancel As Boolean)
    If Me.TxtModFlg2(mIndex) = "R" Then Exit Sub
        Dim s As String
           Dim rs2 As New ADODB.Recordset
           ' If CBoBasedON.ListIndex = 0 And val(TXTOrDer_no2.Text) <> 0 Then Exit Sub
              '  Dim s As String
            If CBoBasedON.ListIndex = 1 Then TXTOrDer_no(0) = TXTOrDer_no2.Text
            If CBoBasedON.ListIndex = -1 Then Exit Sub
            'Else
                If Index <> 1 Then
                TXTOrDer_no(1).Text = TXTOrDer_no2.Text
                
                End If
            'End If
            
            Dim StrSQL As String
            Dim orderStatus As Integer
     
            MintDone = 0
            Set rs2 = New ADODB.Recordset
            If CBoBasedON.ListIndex = 1 Then
                StrSQL = "select * from TblCardAuthorizationReform where WorkOrder = " & val(TXTOrDer_no(1).Text) & " "
            Else
                StrSQL = "select * from TblRowsEstimated where ID =" & val(TXTOrDer_no2.Text) & " "
                
                rs2.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText
                If Not rs2.EOF Then
                    If val(rs2!CBoBasedON & "") = 1 And val(rs2!order_no & "") <> 0 Then
                        StrSQL = "select * from TblCardAuthorizationReform where WorkOrder = " & val(rs2!order_no & "")
                        
                        TXTOrDer_no(0) = rs2!order_no & ""
                       ' txttotal2 = Rs2!HandWagesAmount & ""
                        fg.TextMatrix(1, fg.ColIndex("Name")) = "„‰ Õ—ﬂ… «·ﬁÿ⁄ «·„ﬁœ—… "
                        fg.TextMatrix(1, fg.ColIndex("Price")) = val(rs2!TotalAfterDiscount & "") + val(rs2!Vat2 & "")
                        CalcTotal2
                      '  TXTOrDer_no(1) = dd
                    Else
                        
                    Exit Sub
                        
                    End If
                Else
                    TXTOrDer_no(0) = ""
                End If
            End If
            Set rs2 = New ADODB.Recordset
            rs2.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText
            If rs2.RecordCount > 0 Then
                If (rs2!IsEndAll & "") = 1 Then
                    If SystemOptions.UserInterface = ArabicInterface Then
                        MsgBox "·« Ì„ﬂ‰ «·⁄„· ⁄·Ï «„— „€·ﬁ"
                    Else
                        MsgBox "Cannot work on a closed command"
                    End If
                    GoTo Exits
                End If
                
                
                
                orderStatus = IIf(IsNull(rs2("OrderStatus").value), 0, rs2("OrderStatus").value)
                TxtCashCustomerName.Text = IIf(IsNull(rs2("ClientName").value), "", rs2("ClientName").value)
                'DCOPrType =
                
                DcCustmer(mIndex).BoundText = val(rs2!CusID & "")
                If val(rs2!CusID & "") = 0 Then
                    StrSQL = "SELECT tc.CusID FROM TblCustemers AS tc WHERE tc.CusName LIKE N'%" & Trim(TxtCashCustomerName.Text) & "%'"
                    Dim rsDummy As New ADODB.Recordset
                    rsDummy.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText
                    If Not rsDummy.EOF Then
                        DcCustmer(mIndex).BoundText = val(rsDummy!CusID & "")
                    End If
                End If
                
                 TXTOrDer_no(1) = val(rs2!ID & "")
                DcbyearFactor.Text = val(rs2!YearFact & "")
                TxtPlatNo = Trim(rs2!PlateNo & "")
                DcbCarType.BoundText = val(rs2!CarTypeID & "")
                
                TxtManualNo2(2).Text = Trim(rs2!Shaseh & "")
                 TxtManualNo2(1).Text = Trim(rs2!CarMeter & "")
                
                DcbCarModel.BoundText = IIf(IsNull(rs2("CarModelID").value), "", rs2("CarModelID").value)
                 
                If orderStatus = 2 Or orderStatus = 4 Or orderStatus = 5 Then
                    MintDone = 1
                End If
                If Me.TxtModFlg2(mIndex) = "N" Or Me.TxtModFlg2(mIndex) = "E" Then
                    
                    Dim RsData3 As New ADODB.Recordset
                    
                    
                    s = "Select TblCardAuthorizationReformItems.qty, tblitems.itemid,TblCardAuthorizationReformItems.Price ,TblCardAuthorizationReformItems.TotalWithVat ,tblItems.ItemCode,tblItems.ItemName from TblCardAuthorizationReformItems Left Outer Join tblItems On tblItems.ItemID =TblCardAuthorizationReformItems.ItemID Left Outer join TblCardAuthorizationReform On TblCardAuthorizationReform.Id = TblCardAuthorizationReformItems.id"
                    
                    s = s & "  Where (dbo.TblCardAuthorizationReform.WorkOrder = " & val(TXTOrDer_no(0).Text) & ") "
                           
'                     RsData3.Open s, Cn, adOpenStatic, adLockReadOnly, adCmdText
'                     Fg.Rows = 1
'                     Do While Not RsData3.EOF
'                        Fg.Rows = Fg.Rows + 1
'                        Fg.TextMatrix(Fg.Rows - 1, Fg.ColIndex("Code")) = RsData3!ItemID & ""
'                        Fg.TextMatrix(Fg.Rows - 1, Fg.ColIndex("Name")) = RsData3!ItemID & ""
'                        Fg.TextMatrix(Fg.Rows - 1, Fg.ColIndex("Price")) = RsData3!Price & ""
'                        Fg.TextMatrix(Fg.Rows - 1, Fg.ColIndex("Count")) = RsData3!Qty & ""
'
'
'
'
'                        RsData3.MoveNext
'                    Loop
                    LoadCar
           
       
                     Exit Sub
                End If
                LoadCar
            Else
Exits:
                TxtCashCustomerName.Text = ""
                MintDone = -1
                TXTOrDer_no(0) = ""
                TXTOrDer_no(1) = ""
                TXTOrDer_no2 = ""
                DcbCarType.Text = ""
                DcbyearFactor.Text = ""
                DCEquipments.Text = ""
                TxtManualNo2(2) = ""
                TxtManualNo2(1) = ""
                TxtPlatNo = ""
                DcbCarModel.Text = ""
                DcCustmer(1).Text = ""
            End If
            
            CalcTotal2
'End If
End Sub

Private Sub TXTOrDer_no2_Validate(Cancel As Boolean)
TXTOrDer_no_Validate 0, False
End Sub

Private Sub XPDtbTrans_Change()
    If Me.TxtModFlg2(mIndex) <> "R" Then
        TxtNoteSerial1.Text = ""
        TxtNoteSerial.Text = ""
   End If
       
    CalcTotal2
End Sub


Private Sub btn_First_Click(Index As Integer)
  On Error GoTo ErrTrap

    Dim Msg As String

  
    If Me.TxtModFlg2(mIndex).Text = "N" Then
        FindRec val(TxtSerial1(mIndex).Text)
        TxtModFlg2(mIndex).Text = "R"
    End If

    TxtModFlg2(mIndex) = "R"

    If RsSavRec.RecordCount = 0 Then
        clear_all Me
        Exit Sub
    End If

BegnieWork:
    RsSavRec.MoveFirst
    If mIndex = 1 Then
        FiLLTXT1
   ElseIf mIndex = 2 Then
        FiLLTXT2
   ElseIf mIndex = 3 Then
        FiLLTXT3
        
    End If

    Exit Sub

ErrTrap:

    Select Case Err.Number

        Case -2147217885
        If SystemOptions.UserInterface = ArabicInterface Then
            Msg = "⁄›Ê« ·ﬁœ  „ Õ–› Â–« «·”Ã· " & CHR(13)
            Msg = Msg & "„‰ ﬁ»· „” Œœ„ √Œ— ⁄·Ï «·‘»ﬂ… " & CHR(13)
            Msg = Msg & "”Ê› Ì „  ÕœÌÀ «·»Ì«‰« "
          Else
            Msg = "Sorry I have been deleted the record" & CHR(13)
            Msg = Msg & "By another user on the network " & CHR(13)
            Msg = Msg & "Data will be updated"
        End If
            MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading + vbOKOnly, App.title
            RsSavRec.Requery
            Resume BegnieWork
    End Select
End Sub
Private Sub btn_Next_Click(Index As Integer)
On Error GoTo ErrTrap
    Dim Msg As String

    If Me.TxtModFlg2(mIndex).Text = "N" Then
        FindRec val(TxtSerial1(mIndex).Text)
        TxtModFlg2(mIndex).Text = "R"
    End If

    TxtModFlg2(mIndex) = "R"

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

    If mIndex = 1 Then
        FiLLTXT1
  ElseIf mIndex = 2 Then
        FiLLTXT2
ElseIf mIndex = 3 Then
        FiLLTXT3
      '  FillGridWithData2
    End If
    Exit Sub
ErrTrap:

    Select Case Err.Number

        Case -2147217885
        If SystemOptions.UserInterface = ArabicInterface Then
            Msg = "⁄›Ê« ·ﬁœ  „ Õ–› Â–« «·”Ã· " & CHR(13)
            Msg = Msg & "„‰ ﬁ»· „” Œœ„ √Œ— ⁄·Ï «·‘»ﬂ… " & CHR(13)
            Msg = Msg & "”Ê› Ì „  ÕœÌÀ «·»Ì«‰« "
        Else
            Msg = "Sorry I have been deleted the  record" & CHR(13)
            Msg = Msg & "By another user on the network " & CHR(13)
            Msg = Msg & "Data will be updated"
        End If
            MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading + vbOKOnly, App.title
            RsSavRec.Requery
            Resume BegnieWork
    End Select
End Sub

Private Sub btn_Previous_Click(Index As Integer)
  On Error GoTo ErrTrap
    Dim Msg As String

    If TxtModFlg2(mIndex).Text = "N" Then
        FindRec val(TxtSerial1(mIndex).Text)
        TxtModFlg2(mIndex).Text = "R"
    End If

    Me.TxtModFlg2(mIndex) = "R"

    If RsSavRec.RecordCount = 0 Then
        clear_all Me
        Exit Sub
    End If

BegnieWork:

    RsSavRec.MovePrevious

    If RsSavRec.BOF Then
        RsSavRec.MoveFirst
    End If

    If mIndex = 1 Then
        FiLLTXT1
     ElseIf mIndex = 2 Then
        FiLLTXT2
    ElseIf mIndex = 3 Then
        FiLLTXT3
      '  FillGridWithData2
    End If
    Exit Sub
ErrTrap:

    Select Case Err.Number

        Case -2147217885
        If SystemOptions.UserInterface = ArabicInterface Then
            Msg = "⁄›Ê« ·ﬁœ  „ Õ–› Â–« «·”Ã· " & CHR(13)
            Msg = Msg & "„‰ ﬁ»· „” Œœ„ √Œ— ⁄·Ï «·‘»ﬂ… " & CHR(13)
            Msg = Msg & "”Ê› Ì „  ÕœÌÀ «·»Ì«‰« "
             Else
            Msg = "Sorry I have been deleted the  record" & CHR(13)
            Msg = Msg & "By another user on the network " & CHR(13)
            Msg = Msg & "Data will be updated"
        End If
            MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading + vbOKOnly, App.title
            RsSavRec.Requery
            Resume BegnieWork
    End Select
End Sub
Private Sub btn_Last_Click(Index As Integer)
  On Error GoTo ErrTrap

    Dim Msg As String

    If Me.TxtModFlg.Text = "N" Then
        FindRec val(TxtSerial1(mIndex).Text)
        Me.TxtModFlg2(mIndex).Text = "R"
    End If

    Me.TxtModFlg2(mIndex) = "R"

    If RsSavRec.RecordCount = 0 Then
        clear_all Me
        Exit Sub
    End If

BegnieWork:

    RsSavRec.MoveLast
    If mIndex = 1 Then
        FiLLTXT1
      ElseIf mIndex = 2 Then
        FiLLTXT2
        FillGridWithData2
      ElseIf mIndex = 3 Then
        FiLLTXT3
    End If
    
    Exit Sub

ErrTrap:

    Select Case Err.Number

        Case -2147217885
       If SystemOptions.UserInterface = ArabicInterface Then
            Msg = "⁄›Ê« ·ﬁœ  „ Õ–› Â–« «·”Ã· " & CHR(13)
            Msg = Msg & "„‰ ﬁ»· „” Œœ„ √Œ— ⁄·Ï «·‘»ﬂ… " & CHR(13)
            Msg = Msg & "”Ê› Ì „  ÕœÌÀ «·»Ì«‰« "
        Else
            Msg = "Sorry I have been deleted the record" & CHR(13)
            Msg = Msg & "By another user on the network " & CHR(13)
            Msg = Msg & "Data will be updated"
        End If
            MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading + vbOKOnly, App.title
            RsSavRec.Requery
            Resume BegnieWork
    End Select
End Sub

Private Sub TxtModFlg2_Change(Index As Integer)
 On Error GoTo ErrTrap

    Select Case Me.TxtModFlg2(mIndex).Text

        Case "R"
            '        Me.Caption = " ’—ÌÕ Œ—ÊÃ „ƒﬁ "
            Me.btn_Save(Index).Enabled = False
            Me.Btn_Undo(Index).Enabled = False
            Me.btn_New(Index).Enabled = True
            Me.btn_Modify(Index).Enabled = True
            Me.btn_Delete(Index).Enabled = True
            Me.btn_Query(Index).Enabled = True
            btn_Previous(Index).Enabled = True
            btn_First(Index).Enabled = True
            btn_Last(Index).Enabled = True
            btn_Next(Index).Enabled = True
       
            XPDtbTrans.Enabled = False
            

'            If rs.RecordCount < 1 Then
'                btn_Previous(Index).Enabled = False
'                btn_First(Index).Enabled = False
'                btn_Last(Index).Enabled = False
'                btn_Next(Index).Enabled = False
'                Me.btn_Modify(Index).Enabled = False
'                Me.btn_Delete(Index).Enabled = False
'            End If

        Case "N"
            '        Me.Caption = " ’—ÌÕ Œ—ÊÃ „ƒﬁ ( ÃœÌœ )"
            Me.btn_Save(Index).Enabled = True
            Me.Btn_Undo(Index).Enabled = True
            Me.btn_New(Index).Enabled = False
            Me.btn_Modify(Index).Enabled = False
            Me.btn_Delete(Index).Enabled = False
            Me.btn_Query(Index).Enabled = False
            '      btn_Previous(Index).Enabled = False
            '      btn_First(Index).Enabled = False
            '      btn_Last(Index).Enabled = False
            '      btn_Next(Index).Enabled = False
           
            XPDtbTrans.Enabled = True
            XPDtbTrans.value = Date

        Case "E"
            '        Me.Caption = " ’—ÌÕ Œ—ÊÃ „ƒﬁ (  ⁄œÌ· )"
            Me.btn_Save(Index).Enabled = True
            Me.Btn_Undo(Index).Enabled = True
            Me.btn_New(Index).Enabled = False
            Me.btn_Modify(Index).Enabled = False
            Me.btn_Delete(Index).Enabled = False
            Me.btn_Query(Index).Enabled = False
            
            btn_Previous(Index).Enabled = False
            btn_First(Index).Enabled = False
            btn_Last(Index).Enabled = False
            btn_Next(Index).Enabled = False
      

            XPDtbTrans.Enabled = True
    End Select

    Exit Sub
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



    Set XPic = Me.btn_First(1).ButtonImage
    Set Me.btn_First(1).ButtonImage = Me.btn_Last(1).ButtonImage
    Set Me.btn_Last(1).ButtonImage = XPic
    Set XPic = Me.btn_Previous(1).ButtonImage
    Set Me.btn_Previous(1).ButtonImage = Me.btn_Next(1).ButtonImage
    Set Me.btn_Next(1).ButtonImage = XPic
    

    
lbl(20).Caption = "Job order"

    Me.Caption = "Old Contract Data"
    Label1(2).Caption = Me.Caption

    With Me.Grid
        .TextMatrix(0, .ColIndex("Ser")) = "Ser"
'        .TextMatrix(0, .ColIndex("CusID")) = "Customer Name"
'        .TextMatrix(0, .ColIndex("ContractNo")) = "Contract No."
'        .TextMatrix(0, .ColIndex("ContractDate")) = "Contract Date"
'.TextMatrix(0, .ColIndex("ContractValue")) = "Contract Value"
'.TextMatrix(0, .ColIndex("EndGuranteeDate")) = "End Gurantee Date"
'.TextMatrix(0, .ColIndex("NetValue")) = "Maintenance Value"
'.TextMatrix(0, .ColIndex("Remarks")) = "Remarks"

    End With
    Label1(0).Caption = "Cont.No"
    Label1(4).Caption = "Contr.Date"
    Label1(1).Caption = "Customer"
    Label1(6).Caption = "End Gurantee"
    Label1(9).Caption = "Type"
     Label1(5).Caption = "Maint. Value"
    Label1(8).Caption = "Contract Value"
    Label1(7).Caption = "Remarks"
'XPPnlTime.Caption = "Print End Contruct"
    Label1(3).Caption = "ID"
    'lbl(1).Caption = "From"
 'lbl(0).Caption = "To"
    Label2(0).Caption = "Curr. Rec."
    Label2(1).Caption = "Rec. Count."
    
    Label2(3).Caption = "Curr. Rec."
    Label2(2).Caption = "Rec. Count."
'BtnPrint.Caption = "Print"
    btnNew.Caption = "New"
    btnModify.Caption = "Modify"
    btnSave.Caption = "Save"
    BtnUndo.Caption = "Undo"
    btnDelete.Caption = "Delete"
    btnCancel.Caption = "Exit"
    Frame5.Caption = "Invoice data"
    cmdPrintNote.Caption = "Print Note"
    'Label5.Caption = "No"
    lbltycar.Caption = "Type of car"
    LblYear.Caption = "car model"
    LblPla.Caption = "Plate Number"
    lblModel.Caption = "Style"
    Label4.Caption = "Remark"
    lbl(2).Caption = "Date"
    lbl(56).Caption = "Based on"
    lbl(7).Caption = "Branch"
    lbl(15).Caption = "Customer"
    lbl(33).Caption = "Customer cash"
    lbl(119).Caption = "Chassis"
    lbl(118).Caption = "kilometer"
    lbl(43).Caption = "Total hand wages"
    lbl(10).Caption = "Total spare parts"
    lbl(14).Caption = "Total before tax"
    lbl(11).Caption = "Vat"
    lbl(12).Caption = "Discount"
    lbl(13).Caption = "Net"
    lbl(0).Caption = "Discount"
    lbl(1).Caption = "Disc Percent"
    lbl(9).Caption = "Total"
    lbl(65).Caption = "Vat Value"
    lbl(5).Caption = "Vat"
    lbl(47).Caption = "Net"
    Label1(14).Caption = "Entry No"

    Label1(10).Caption = "Hand wages"
    'Label5.Caption = "No"
    'Label5.Caption = "No"
     With Me.fg
        .TextMatrix(0, .ColIndex("Ser")) = "Ser"
        .TextMatrix(0, .ColIndex("Name")) = "Name"
        .TextMatrix(0, .ColIndex("Price")) = "Price"
    End With
        
        
     With Me.grdTrans
        .TextMatrix(0, .ColIndex("Ser")) = "Ser"
        .TextMatrix(0, .ColIndex("Show")) = "Show"
        .TextMatrix(0, .ColIndex("NoteSerial1")) = "Note Serial"
        .TextMatrix(0, .ColIndex("Transaction_Date")) = "Date"
        .TextMatrix(0, .ColIndex("Trans_Discount")) = "Discount"
        .TextMatrix(0, .ColIndex("Total")) = "Total"
        .TextMatrix(0, .ColIndex("Total2")) = "Total B Vat"
        .TextMatrix(0, .ColIndex("Vat")) = "Vate"
        .TextMatrix(0, .ColIndex("NetInvoice")) = "Net"
        .TextMatrix(0, .ColIndex("remark")) = "Remark"
        
        
    End With
        
        
    btn_New(1).Caption = "New"
    btn_Modify(1).Caption = "Modify"
    btn_Save(1).Caption = "Save"
    Btn_Undo(1).Caption = "Undo"
    btn_Delete(1).Caption = "Delete"
    Btn_Print(1).Caption = "Print"
    btn_Query(1).Caption = "Search"
    Btn_Update(1).Caption = "Refresh"
    btn_Cancel(1).Caption = "Exit"
    Cmd_DeleteRow(1).Caption = "Delete Row"
    Cmd_DeleteAll(1).Caption = "Delete all"
    lbl(8).Caption = "User Name"
    If mIndex = 1 Then
        Me.Caption = "Hand wages"
        TabMain.TabCaption(1) = "Hand wages"
        TabMain2.TabCaption(0) = "Data"
        TabMain2.TabCaption(1) = "Invoices"
    End If
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

'    If DoPremis(Do_Delete, Me.name, True) = False Then
'        Exit Sub
'    End If
If val(DcbEmpUsrID.BoundText) = 0 Then
If SystemOptions.UserInterface = ArabicInterface Then
MsgBox "Ì—ÃÏ «Œ Ì«— «·„‰œÊ»"
Else
MsgBox "Please Select Employee"
End If
Exit Sub
End If
    If TxtVac_ID.Text <> "" Then
    If SystemOptions.UserInterface = ArabicInterface Then
        MSGType = MsgBox("Â·  —€» ›Ì Õ–› Â–« «·”Ã·", vbQuestion + vbMsgBoxRtlReading + vbYesNo + vbMsgBoxRight, App.title)
        Else
        MSGType = MsgBox("ÂConfirm Delete", vbQuestion + vbMsgBoxRtlReading + vbYesNo + vbMsgBoxRight, App.title)
     End If

        If MSGType = vbYes Then
        Cn.Execute "Update TblUserScreen set FlgWork=null where id=" & val(Me.DcbScreen.BoundText) & ""
            RsSavRec.Find "id=" & val(TxtVac_ID.Text), , adSearchForward, 1
            RsSavRec.delete
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
        If SystemOptions.UserInterface = ArabicInterface Then
            StrMSG = "⁄›Ê« ·« ÌÃÊ“ Õ–› «·”Ã· ·«— »«ÿÂ »»Ì«‰«  √Œ—Ì"
            Else
            StrMSG = "Sorry... Can not Delete.  is related to with other data"
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
        If SystemOptions.UserInterface = ArabicInterface Then
            Msg = "⁄›Ê« ·ﬁœ  „ Õ–› Â–« «·”Ã· «· «·Ï" & CHR(13)
            Msg = Msg & "„‰ ﬁ»· „” Œœ„ √Œ— ⁄·Ï «·‘»ﬂ… " & CHR(13)
            Msg = Msg & "”Ê› Ì „  ÕœÌÀ «·»Ì«‰« "
            Else
            Msg = "Sorry I've been to delete this record" & CHR(13)
            Msg = Msg & "From Another user on network " & CHR(13)
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
        If SystemOptions.UserInterface = ArabicInterface Then
            Msg = "⁄›Ê« ·ﬁœ  „ Õ–› Â–« «·”Ã· «· «·Ï" & CHR(13)
            Msg = Msg & "„‰ ﬁ»· „” Œœ„ √Œ— ⁄·Ï «·‘»ﬂ… " & CHR(13)
            Msg = Msg & "”Ê› Ì „  ÕœÌÀ «·»Ì«‰« "
              Else
            Msg = "Sorry I've been to delete this record" & CHR(13)
            Msg = Msg & "From Another user on network " & CHR(13)
            Msg = Msg & "Data will be updated"
            End If
            MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading + vbOKOnly, App.title
            RsSavRec.Requery
            Resume BegnieWork
    End Select

End Sub

Private Sub btnModify_Click()
    Dim Msg As String

'    If DoPremis(Do_Edit, Me.name, True) = False Then
'        Exit Sub
'    End If
If val(DcbEmpUsrID.BoundText) = 0 Then
If SystemOptions.UserInterface = ArabicInterface Then
MsgBox "Ì—ÃÏ «Œ Ì«— «·„‰œÊ»"
Else
MsgBox "Please Select Employee"
End If
Exit Sub
End If
Frame3.Enabled = False

    On Error GoTo ErrTrap

    If TxtVac_ID.Text <> "" Then
        TxtModFlg = "E"
       ' Frm2.Enabled = True
      
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
            Msg = "Sorry" & CHR(13)
            Msg = Msg & " ·Currently can not be edited" & CHR(13)
            Msg = Msg & "Where it was being edited by another user on the network"
           
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

'    If DoPremis(Do_New, Me.name, True) = False Then
'        Exit Sub
'    End If

    On Error GoTo ErrTrap
    Set rs = New ADODB.Recordset
   ' Frm2.Enabled = True
   
    '-----------------------------------
    Me.TxtVac_ID.Text = ""
 
    Frame3.Enabled = False
    '-----------------------------------
    TxtModFlg.Text = "N"
clear_all Me
FillGridWithData
 My_SQL = "select ID,Name From TblUserScreen WHERE     (FlgWork IS NULL)"
    fill_combo Me.DcbScreen, My_SQL
    My_SQL = "TblVisitScreen"
    rs.Open My_SQL, Cn, adOpenKeyset, adLockOptimistic, adCmdTable

    If rs.RecordCount > 0 Then
        TxtSerial.Text = rs.RecordCount + 1
    Else
        TxtSerial.Text = 1
    End If

    rs.Close
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
        If SystemOptions.UserInterface = ArabicInterface Then
            Msg = "⁄›Ê« ·ﬁœ  „ Õ–› Â–« «·”Ã· «· «·Ï" & CHR(13)
            Msg = Msg & "„‰ ﬁ»· „” Œœ„ √Œ— ⁄·Ï «·‘»ﬂ… " & CHR(13)
            Msg = Msg & "”Ê› Ì „  ÕœÌÀ «·»Ì«‰« "
              Else
            Msg = "Sorry I've been to delete this record" & CHR(13)
            Msg = Msg & "From Another user on network " & CHR(13)
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
        If SystemOptions.UserInterface = ArabicInterface Then
            Msg = "⁄›Ê« ·ﬁœ  „ Õ–› Â–« «·”Ã· «· «·Ï" & CHR(13)
            Msg = Msg & "„‰ ﬁ»· „” Œœ„ √Œ— ⁄·Ï «·‘»ﬂ… " & CHR(13)
            Msg = Msg & "”Ê› Ì „  ÕœÌÀ «·»Ì«‰« "
              Else
            Msg = "Sorry I've been to delete this record" & CHR(13)
            Msg = Msg & "From Another user on network " & CHR(13)
            Msg = Msg & "Data will be updated"
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
   If val(DcbScreen.BoundText) = 0 Then
   If SystemOptions.UserInterface = ArabicInterface Then
   MsgBox "Ì—ÃÏ «Œ Ì«— „« „"
   Else
   MsgBox "Please Select Screen"
   End If
   DcbScreen.SetFocus
   Exit Sub
   End If
If val(DcbEmpUsrID.BoundText) = 0 Then
If SystemOptions.UserInterface = ArabicInterface Then
MsgBox "Ì—ÃÏ «œŒ«· ”„ «·„‰œÊ»"
Else
MsgBox "Please Enter Employee Name"
End If
Exit Sub
End If
If TxtEmpRemark.Text = "" Then
If SystemOptions.UserInterface = ArabicInterface Then
MsgBox "Ì—ÃÏ «œŒ«· „·«ÕŸ«  «·„‰œÊ»"
Else
MsgBox "Please Enter Remarks"
End If
Exit Sub
End If
    '------------------------------ check if Empcode exist ----------------------

    'StrVacName = IsRecExist("TblVisit", "ID", Trim(TxtContractNo.Text), "ID", "Vac_ID<>'" & Trim(TxtVac_ID.Text) & "'")

    'If StrVacName <> "" Then
    'If SystemOptions.UserInterface = ArabicInterface Then
    '    Msg = "·ﬁœ ”»ﬁ  ”ÃÌ· Â–« «·‰Ê⁄ „‰ ﬁ»·"
    '    Else
    '    Msg = "This type already exists"
    '    End If
    '    MsgBox Msg, vbOKOnly + vbMsgBoxRight, App.title
    '    TxtContractNo.SetFocus
    '
    '    Exit Sub

    'End If

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
If SystemOptions.UserInterface = ArabicInterface Then
    MsgBox "Â‰«ﬂ Œÿ√ „« ›Ì ≈œŒ«· «·»Ì«‰« ", vbOKOnly + vbMsgBoxRight, App.title
   Else
   MsgBox "Sorry...error in douring enter data", vbOKOnly + vbMsgBoxRight, App.title
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
        Msg = "No new data found"
    Else
        Msg = "No Rec.Before Update " & vbCrLf & FristCount & vbCrLf & "No Rec.After Update" & vbCrLf & LastCount
        
        If LastCount > FristCount Then
            Msg = Msg + vbCrLf & "No New Record" & vbCrLf & LastCount - FristCount
        Else
            Msg = Msg + vbCrLf & "No Deleted Record"" & vbCrLf & FristCount - LastCount"
        End If
    End If
End If
    MsgBox Msg, vbOKOnly + vbMsgBoxRight + vbInformation, App.title
ErrTrap:
End Sub


Private Sub cmdAdd_Click()
   If val(DcbScreen.BoundText) = 0 Then
   If SystemOptions.UserInterface = ArabicInterface Then
   MsgBox "Ì—ÃÏ «Œ Ì«— „« „"
   Else
   MsgBox "Please Select Screen"
   End If
   DcbScreen.SetFocus
   Exit Sub
   End If
If val(DcbEmpUsrID.BoundText) = 0 Then
If SystemOptions.UserInterface = ArabicInterface Then
MsgBox "Ì—ÃÏ «œŒ«· ”„ «·„‰œÊ»"
Else
MsgBox "Please Enter Employee Name"
End If
Exit Sub
End If
If TxtEmpRemark.Text = "" Then
If SystemOptions.UserInterface = ArabicInterface Then
MsgBox "Ì—ÃÏ «œŒ«· „·«ÕŸ«  «·„‰œÊ»"
Else
MsgBox "Please Enter Remarks"
End If
Exit Sub
End If
AddNewRec
End Sub

Private Sub CmdDel_Click()
btnDelete_Click
End Sub

Private Sub CmdMod_Click()
   If val(DcbScreen.BoundText) = 0 Then
   If SystemOptions.UserInterface = ArabicInterface Then
   MsgBox "Ì—ÃÏ «Œ Ì«— „« „"
   Else
   MsgBox "Please Select Screen"
   End If
   DcbScreen.SetFocus
   Exit Sub
   End If
If val(DcbEmpUsrID.BoundText) = 0 Then
If SystemOptions.UserInterface = ArabicInterface Then
MsgBox "Ì—ÃÏ «œŒ«· ”„ «·„‰œÊ»"
Else
MsgBox "Please Enter Employee Name"
End If
Exit Sub
End If
If TxtEmpRemark.Text = "" Then
If SystemOptions.UserInterface = ArabicInterface Then
MsgBox "Ì—ÃÏ «œŒ«· „·«ÕŸ«  «·„‰œÊ»"
Else
MsgBox "Please Enter Remarks"
End If
Exit Sub
End If
FiLLRec
End Sub

Private Sub DcbEmpUsrID_Change()
If val(Me.DcbEmpUsrID.BoundText) <> 0 Then
Frame4.Enabled = True
Else
Frame4.Enabled = False
End If
End Sub

Private Sub Form_Load()
'    On Error GoTo ErrTrap
    Dim i As Integer
    Dim My_SQL As String
    Dim Dcombos As ClsDataCombos
Frame3.Enabled = False
Frame2(0).Enabled = False
Frame4.Enabled = False
    RecordDate = Date

    TabMain.TabVisible(0) = False
     TabMain.TabVisible(1) = False
     TabMain.TabVisible(2) = False
     TabMain.TabVisible(3) = False
     If mIndex = 0 Then
        TabMain.TabVisible(0) = True
        TabMain.CurrTab = 0
    ElseIf mIndex = 1 Then
        TabMain.TabVisible(1) = True
        TabMain.CurrTab = 1
        Me.Caption = "√ÃÊ— «·Ìœ"
      
    ElseIf mIndex = 2 Then
        Me.Width = Grid2.Width + 400
        TabMain.TabVisible(mIndex) = True
        TabMain.CurrTab = mIndex
       ' Me.Width = Grid.Width + 400
    ScreenNameArabic = "«‰Ê«⁄ „ﬂ« » «· ›ÊÌ÷"
     
    ElseIf mIndex = 3 Then
        'Me.Width = GRID2.Width + 400
        TabMain.TabVisible(mIndex) = True
        TabMain.CurrTab = mIndex
       ' Me.Width = Grid.Width + 400
    ScreenNameArabic = " ⁄—Ì› «·⁄œ”« "
     
     Dim CC As Integer
        For CC = 0 To cmbFlag.count - 1
            cmbFlag(CC).Clear
            cmbFlag(CC).AddItem "--"
            cmbFlag(CC).AddItem "-+"
            cmbFlag(CC).AddItem "++"
            
        Next
    End If

    With Me.CBoBasedON
        .Clear
        '.AddItem "»·«"
        .AddItem "ﬁÿ⁄ «·€Ì«— «· ﬁœÌ—Ì…"
        .AddItem "«„— «’·«Õ-Ê—‘ "

    End With

   
    Set Dcombos = New ClsDataCombos
    Dcombos.GetUsers Me.DCboUserName(1)
    
    Dcombos.GetCustomersSuppliers 1, DcCustmer(mIndex), , , 1
    Dcombos.GetTblCarsDataGroup Me.DcbCarType
    Dcombos.GetBranches Me.Dcbranch(1)
   ' Dcombos.GetTblCarModels Me.DcbCarModel
    
    If SystemOptions.UserInterface = EnglishInterface Then
        My_SQL = "SELECT id,ISNULL(ModelE,Model) ModelName from TblCarModels"
    Else
        My_SQL = "SELECT id, Model from TblCarModels"
    End If
    fill_combo DcbCarModel, My_SQL
      'Dim ii As Integer
     
      For II = 1900 To 2100
        Me.DcbyearFactor.AddItem (II)
      Next II
      
Me.Dcbranch(1).BoundText = branch_id
    Resize_Form Me
    
    
    SetDtpickerDate Me.XPDtbTrans
    
    If mIndex = 1 Then
        My_SQL = "TblHandWages"
       ' Set BKGrndPic = New ClsBackGroundPic
        Set RsSavRec = New ADODB.Recordset
        RsSavRec.CursorLocation = adUseClient
        RsSavRec.Open My_SQL, Cn, adOpenKeyset, adLockPessimistic, adCmdTableDirect
        TxtModFlg2(mIndex).Text = "R"
        DCboUserName(mIndex).BoundText = user_id
       

        

        btn_First_Click (mIndex)
    ElseIf mIndex = 0 Then

        My_SQL = "select ID,Name From TblUserScreen "
        fill_combo Me.DcbScreen, My_SQL
        Set Dcombos = New ClsDataCombos
        Set cSearch = New clsDCboSearch
        Dcombos.GetUsers DcbUserID
        Dcombos.GetUserComp DcbEmpUsrID
        
        My_SQL = "TblVisitScreen"
        Set BKGrndPic = New ClsBackGroundPic
        Set RsSavRec = New ADODB.Recordset
        RsSavRec.CursorLocation = adUseClient
        RsSavRec.Open My_SQL, Cn, adOpenKeyset, adLockPessimistic, adCmdTableDirect
        Me.TxtModFlg.Text = "R"
        Resize_Form Me
        
        'load tblUsers -----------------------------------------------
    
    
        FillGridWithData
    
        With Me.Grid
    '        .Cell(flexcpPicture, 0, .ColIndex("ContractNo")) = Me.GrdImageList.ListImages("Vac_Name").ExtractIcon
    '        .Cell(flexcpPicture, 0, .ColIndex("Ser")) = Me.GrdImageList.ListImages("Ser").ExtractIcon
    
            For i = 0 To .Cols - 1
                .Cell(flexcpPictureAlignment, 0, i) = flexPicAlignRightCenter
            Next
       
            .ExtendLastCol = True
            .WallPaper = BKGrndPic.Picture
            .RowHeight(-1) = 300
        End With
    
        BtnFirst_Click
    ElseIf mIndex = 2 Then
       My_SQL = "TblOffice"
        Set BKGrndPic = New ClsBackGroundPic
        Set RsSavRec = New ADODB.Recordset
        RsSavRec.CursorLocation = adUseClient
        RsSavRec.Open My_SQL, Cn, adOpenKeyset, adLockPessimistic, adCmdTableDirect
        FillGridWithData2
        
        If SystemOptions.UserInterface = EnglishInterface Then
           
        End If
        Me.Caption = "«‰Ê«⁄ «·„ﬂ« » «·„›Ê÷…"
       btn_First_Click (mIndex)
    
    
    ElseIf mIndex = 3 Then
       My_SQL = "TblLensesTypes"
        Set BKGrndPic = New ClsBackGroundPic
        Set RsSavRec = New ADODB.Recordset
        RsSavRec.CursorLocation = adUseClient
        RsSavRec.Open My_SQL, Cn, adOpenKeyset, adLockPessimistic, adCmdTableDirect
        FillGridWithData3
        
        If SystemOptions.UserInterface = EnglishInterface Then
           
        End If
        Me.Caption = "«‰Ê«⁄ «·⁄œ”« "
       btn_First_Click (mIndex)
    
    

        
        Set Dcombos = New ClsDataCombos
        Dcombos.GetItemSGroups cmbGroupId
        Dcombos.GetItemsUnits cmbUnitID
        
        Dim StrSQL As String

     
        If SystemOptions.UserInterface = ArabicInterface Then
        
            StrSQL = "Select sph as sph , spht  as SPHName From SPHTable   order by id"
        Else
            StrSQL = "Select sph as sph , spht  as SPHName From SPHTable   order by id"
        End If
        
        fill_combo DCBoMain(2), StrSQL
        fill_combo DCBoMain(5), StrSQL
        
        
        
        If SystemOptions.UserInterface = ArabicInterface Then
        
            StrSQL = "Select   CLY as CLY ,CLYT  as CLYName From CLYTable   order by id"
        Else
            StrSQL = "Select   CLY as CLY ,CLYT  as CLYName From CLYTable   order by id"
        End If
        
        fill_combo DCBoMain(3), StrSQL
        fill_combo DCBoMain(6), StrSQL
        
        
    End If
    ShowTip

    If SystemOptions.UserInterface = EnglishInterface Then
        SetInterface Me
        ChangeLang
    End If

   ' If OPEN_NEW_SCREEN = True Then
   '     btnNew_Click
   ' End If
Me.DcbUserID.BoundText = user_id
ErrTrap:
End Sub
Private Function GetTblCarModels(MyCombo As DataCombo, _
                                Optional BolLoadAdmins As Boolean = True, Optional Index As Integer)
    Dim StrSQL As String

    If SystemOptions.UserInterface = ArabicInterface Then
   StrSQL = "SELECT id, Model from TblCarModels where CarID=" & Index & ""
    Else
        StrSQL = "SELECT id,Model From TblCarModels where CarID=" & Index & ""
    End If
 
   
End Function

Private Sub Form_QueryUnload(Cancel As Integer, _
                             UnloadMode As Integer)
   Dim IntResult As String
    Dim StrMSG As String
    On Error GoTo ErrTrap
    Dim mSelectModFlg As String
    If mIndex = 0 Then
        mSelectModFlg = Me.TxtModFlg.Text
    Else
        mSelectModFlg = Me.TxtModFlg2(mIndex).Text
    End If
    
    
    If mSelectModFlg <> "R" Then

        Select Case mSelectModFlg

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
                If mIndex = 0 Then
                    btnSave_Click
                Else
                    btn_Save_Click CInt(mIndex)
                End If

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
  '  Set FrmVacancy = Nothing

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
    If mIndex = 0 Then
        StrRecID = new_id("TblVisitScreen", "id", "")
            RsSavRec.AddNew
            RsSavRec.Fields("id").value = IIf(StrRecID <> "", StrRecID, Null)
            FiLLRec
 
    ElseIf mIndex = 1 Then
        StrRecID = new_id("TblHandWages", "id", "")
 
       ElseIf mIndex = 2 Then
        StrRecID = new_id("TblOffice", "id", "")
        RsSavRec.AddNew
        
       ElseIf mIndex = 3 Then
        StrRecID = new_id("TblLensesTypes", "id", "")
        RsSavRec.AddNew
        
                

    End If
    
    
'    RsSavRec.AddNew
'    RsSavRec.Fields("id").value = IIf(StrRecID <> "", StrRecID, Null)
   
    If mIndex = 1 Then
        FiLLRec1
    ElseIf mIndex = 2 Then
        FiLLRec2
    ElseIf mIndex = 3 Then
        FiLLRec3

    End If
    
ErrTrap:

End Sub

Public Sub FiLLRec()
    On Error GoTo ErrTrap
     RsSavRec.Fields("UserID").value = IIf(DcbUserID.BoundText <> 0, val(DcbUserID.BoundText), Null)
     RsSavRec.Fields("EmpUsrID").value = IIf(DcbEmpUsrID.BoundText <> 0, val(DcbEmpUsrID.BoundText), Null)
     RsSavRec.Fields("ScreenID").value = IIf(DcbScreen.BoundText <> 0, val(DcbScreen.BoundText), Null)
     RsSavRec.Fields("UserPass").value = TxtUserPass.Text
     RsSavRec.Fields("EmpPass").value = TxtEmpPass.Text
     RsSavRec.Fields("RecordDate").value = RecordDate.value
     RsSavRec.Fields("CusRemark").value = TxtCusRemark.Text
     RsSavRec.Fields("EmpRemark").value = TxtEmpRemark.Text
     RsSavRec.update
     Cn.Execute "Update TblUserScreen set FlgWork=1 where id=" & val(Me.DcbScreen.BoundText) & ""
    If SystemOptions.UserInterface = ArabicInterface Then
    MsgBox " „  ⁄„·Ì… «·Õ›Ÿ »‰Ã«Õ", vbInformation + vbOKOnly + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
    Else
    MsgBox "Save Successfully", vbInformation + vbOKOnly + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
    End If
    FiLLTXT
    FillGridWithData
    
    TxtModFlg = "R"

    Exit Sub
ErrTrap:

    If RsSavRec.EditMode <> adEditNone Then
        RsSavRec.CancelUpdate
    End If

End Sub



Public Sub FiLLRec1()
    On Error GoTo ErrTrap
    
   
        If TxtNoteSerial1.Text = "" Then
                If Voucher_coding(val(Dcbranch(mIndex).BoundText), XPDtbTrans.value, 81, 1100, , , , , , , "TblHandWages") = "error" Then
                    If SystemOptions.UserInterface = ArabicInterface Then
                        MsgBox " ·« Ì„ﬂ‰ «÷«›… ”‰œ  ÃœÌœ ·«‰ﬂ  ⁄œÌ  «·Õœ «·–Ì ﬁ„  » ÕœÌœ… „‰ «·”‰œ«   ": Exit Sub
                    Else
                        MsgBox " Cant't Create Expenses Voucher to this Process no You exceed the maximum number ": Exit Sub
                    End If

                Else
         
                    If Voucher_coding(val(Dcbranch(mIndex).BoundText), XPDtbTrans.value, 81, 1100, , , , , , , "TblHandWages") = "" Then
                        If SystemOptions.UserInterface = ArabicInterface Then
                            
                            TxtNoteSerial1.locked = False
                            MsgBox " ·«»œ „‰ ﬂ «»… —ﬁ„ «·”‰œ ÌœÊÌ« ﬂ„« Õœœ   ": Exit Sub
                        Else
                            MsgBox "  Enter Voucher No Manually or Define Coding ": Exit Sub
                        End If

                    Else
                        TxtNoteSerial1.Text = Voucher_coding(val(Dcbranch(mIndex).BoundText), XPDtbTrans.value, 81, 1100, , , , , , , "TblHandWages")
                    End If
                End If
            End If
    
    
    If TxtModFlg2(mIndex).Text = "N" Then
           ' TxtSerial1(mIndex).Text = CStr(new_id("TblAdditionsAssest", "ID", "", True))

       
        RsSavRec.AddNew
        TxtSerial1(mIndex).Text = new_id("TblHandWages", "id", "")
        RsSavRec.Fields("ID").value = val(TxtSerial1(mIndex).Text)
    End If
    RsSavRec("NoteSerial1").value = Trim$(Me.TxtNoteSerial1.Text)
    RsSavRec.Fields("BranchID").value = IIf(Dcbranch(mIndex).Text <> "", Trim(Dcbranch(mIndex).BoundText), Null)
    RsSavRec("RecordDate").value = XPDtbTrans.value
    
    RsSavRec("CBoBasedON").value = CBoBasedON.ListIndex
   DCboUserName(mIndex).BoundText = IIf(DCboUserName(mIndex).Text <> "", Trim(DCboUserName(mIndex).BoundText), user_id)
   'RsSavRec("RecType").value = cmbRecType.ListIndex
    'RsSavRec("ContractNo").value = txtContractNo.Text
    'RsSavRec("RecName").value = txtRecName.Text
    'RsSavRec("RecordTime").value = XPDtbTransTime.Value
    

    RsSavRec.Fields("OrDer_no").value = val(TXTOrDer_no(0).Text)
    RsSavRec.Fields("OrDer_no2").value = val(TXTOrDer_no(1).Text)
    RsSavRec.Fields("RowsEstimatedID").value = val(TXTOrDer_no2.Text)
 
    
    RsSavRec.Fields("DiscValue").value = val(txtDiscValue.Text)
    RsSavRec.Fields("Total2").value = val(txtTotal2.Text)
    RsSavRec.Fields("VatYou").value = val(txtVatYou.Text)
    RsSavRec.Fields("DiscPercent").value = val(txtDiscPercent.Text)
    
    RsSavRec.Fields("Total").value = val(txtTotal.Text)
    RsSavRec.Fields("Vat2").value = val(txtVat2.Text)
    RsSavRec.Fields("Net").value = val(txtNet.Text)
    
    RsSavRec("Remarks").value = TxtRemarks.Text
    
    
    '*********************
     
    
    
      
   

    RsSavRec.update
    cmdDelNote_Click
    Dim s As String
                
    If mIndex = 1 Then
        s = " Delete From TblHandWages2 Where MasterID = " & val(TxtSerial1(mIndex).Text)
    
        
        
    End If
    Cn.Execute s
    
    s = "Select * from TblHandWages2 Where Id = -1"
    'saveGrid s, fg, "Name", "ID", "MasterID", val(TxtSerial1(mIndex).Text)
    saveGrid s, fg, "Name", "", "MasterID", val(TxtSerial1(mIndex).Text)
    
    CmdCreateV2_Click
'
If SystemOptions.UserInterface = ArabicInterface Then
    MsgBox " „  ⁄„·Ì… «·Õ›Ÿ »‰Ã«Õ", vbInformation + vbOKOnly + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
 Else
   MsgBox "Save Successfully", vbInformation + vbOKOnly + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
End If
    'CuurentLogdata
    
    TxtModFlg2(mIndex) = "R"

    Exit Sub
ErrTrap:

    If RsSavRec.EditMode <> adEditNone Then
        RsSavRec.CancelUpdate
    End If

End Sub


Public Sub FiLLTXT(Optional TypeTr As Integer = 0)

    On Error GoTo ErrTrap
    Dim i As Integer
    Dim Shifttime As Date
   ' Frm2.Enabled = False
   Dim My_SQL As String
  My_SQL = "select ID,Name From TblUserScreen "
  If TypeTr = 0 Then
  My_SQL = My_SQL & " WHERE     (FlgWork IS NULL)"
  End If
    fill_combo Me.DcbScreen, My_SQL
  '  Frame3.Enabled = False
'Frame2.Enabled = False
'Me.DcbEmpUsrID.BoundText = 0
'TxtEmpPass.Text = ""
'TxtUserPass.Text = ""
    TxtVac_ID.Text = IIf(IsNull(RsSavRec.Fields("id").value), "", RsSavRec.Fields("id").value)
    TxtSerial.Text = IIf(IsNull(RsSavRec.Fields("id").value), "", RsSavRec.Fields("id").value)
    Me.DcbUserID.BoundText = IIf(IsNull(RsSavRec.Fields("UserID").value), "", RsSavRec.Fields("UserID").value)
   ' Me.DcbEmpUsrID.BoundText = IIf(IsNull(RsSavRec.Fields("EmpUsrID").value), "", RsSavRec.Fields("EmpUsrID").value)
  '  TxtEmpPass.Text = IIf(IsNull(RsSavRec.Fields("EmpPass").value), "", RsSavRec.Fields("EmpPass").value)
  '  TxtUserPass.Text = IIf(IsNull(RsSavRec.Fields("UserPass").value), "", RsSavRec.Fields("UserPass").value)
    RecordDate.value = IIf(IsNull(RsSavRec.Fields("RecordDate").value), Date, RsSavRec.Fields("RecordDate").value)
    TxtCusRemark.Text = IIf(IsNull(RsSavRec.Fields("CusRemark").value), "", RsSavRec.Fields("CusRemark").value)
    TxtEmpRemark.Text = IIf(IsNull(RsSavRec.Fields("EmpRemark").value), "", RsSavRec.Fields("EmpRemark").value)
    Me.DcbScreen.BoundText = IIf(IsNull(RsSavRec.Fields("ScreenID").value), "", RsSavRec.Fields("ScreenID").value)


    LabCurrRec.Caption = RsSavRec.AbsolutePosition
    LabCountRec.Caption = RsSavRec.RecordCount

    With Grid

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

Public Sub EditRec(StrTable As String, _
                   RecId As String)
    FiLLRec

End Sub

Private Sub Grid_Click()
    On Error GoTo ErrTrap
    FindRec val(Me.Grid.TextMatrix(Me.Grid.Row, Me.Grid.ColIndex("id")))
    FiLLTXT 1
ErrTrap:
End Sub

Private Sub Grid_EnterCell()
    On Error GoTo ErrTrap
    FindRec val(Me.Grid.TextMatrix(Me.Grid.Row, Me.Grid.ColIndex("id")))
ErrTrap:
End Sub

Function CheckPassworUserComp() As Double
Dim StrSQL As String
Dim Rs3 As ADODB.Recordset
Set Rs3 = New ADODB.Recordset
       StrSQL = "Select * From TblUserComp Where  Password='" & Trim(Me.TxtEmpPass.Text) & "'"
Rs3.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText
If Rs3.RecordCount > 0 Then
CheckPassworUserComp = IIf(IsNull(Rs3("id").value), 0, Rs3("id").value)
Else
CheckPassworUserComp = 0
End If
End Function


Function CheckPassworUser() As Boolean
Dim StrSQL As String
Dim Rs3 As ADODB.Recordset
Set Rs3 = New ADODB.Recordset
       StrSQL = "Select * From TblUsers Where UserID=" & Me.DcbUserID.BoundText & " AND PassWord='" & Trim(Me.TxtUserPass.Text) & "'"
Rs3.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText
If Rs3.RecordCount > 0 Then
CheckPassworUser = True
Else
CheckPassworUser = False
End If
End Function


Function print_report(Optional NoteSerial As String, Optional Ind As Integer = 0)
    

    
    Dim MySQL As String
    Dim RsData As New ADODB.Recordset
    Dim xApp As New CRAXDRT.Application
    Dim xReport As CRAXDRT.Report
    Dim CViewer As ClsReportViewer
    Dim StrReportTitle As String
    Dim StrFileName As String
    Dim Msg As String
    Dim s As String
    
 Dim RsData2  As New ADODB.Recordset
        Dim RsData3  As New ADODB.Recordset
        Dim RsDetails1 As New ADODB.Recordset
        Dim StrSQL As String
 If mIndex = 1 Then
     MySQL = " SELECT    distinct  '" & DcbCarModel.Text & "' as CarModel,TblHandWages.Remarks, TblHandWages.NoteSerial1,TblHandWages.Total2,TblHandWages.OrDer_no,TblHandWages.Total,TblHandWages.VatYou,TblHandWages.DiscValue,TblHandWages.Net,TblHandWages.Net,"
     MySQL = MySQL & "                     TblHandWages2.Name ,TblHandWages.Remarks,"
     MySQL = MySQL & "                     dbo.TblCardAuthorizationReformDetails.Type , dbo.TblCardAuthorizationReformDetails.Mainte,  dbo.TblMaintenanceWork.NameE, "
     MySQL = MySQL & "                     dbo.TblCardAuthorizationReformDetails.EmpID, TblEmployee_2.Emp_Name AS fiter, TblEmployee_2.Emp_Namee AS fitere,"
     MySQL = MySQL & "                     dbo.TblCardAuthorizationReformDetails.empsuper, TblEmployee_1.Emp_Name AS NameSuper, TblEmployee_1.Emp_Namee AS NamesuperE,"
     MySQL = MySQL & "                     dbo.TblCardAuthorizationReformDetails.Deptid, dbo.TblEmpDepartments.DepartmentName, dbo.TblEmpDepartments.DepartmentNamee,"
     MySQL = MySQL & "                     dbo.TblCardAuthorizationReformDetails.Dpeterial, dbo.TblCardAuthorizationReformDetails.DeptBr, dbo.TblCardAuthorizationReformDetails.DeptColor,"
     MySQL = MySQL & "                     dbo.TblCardAuthorizationReformDetails.PriceFitter, dbo.TblCardAuthorizationReformDetails.payed, dbo.TblCardAuthorizationReformDetails.allocation,"
     MySQL = MySQL & "                     dbo.TblCardAuthorizationReformDetails.TimOut, dbo.TblCardAuthorizationReformDetails.TimeEnter, dbo.TblCardAuthorizationReformDetails.DateExit,"
     MySQL = MySQL & "                     dbo.TblCardAuthorizationReformDetails.DateEnter, dbo.TblCardAuthorizationReformDetails.finish, dbo.TblCardAuthorizationReformDetails.nohours,"
     MySQL = MySQL & "                     dbo.TblCardAuthorizationReformDetails.bill, dbo.TblCardAuthorizationReformDetails.comp, dbo.TblCardAuthorizationReformDetails.[count],"
     MySQL = MySQL & "                     dbo.TblCardAuthorizationReformDetails.[Value], dbo.TblCardAuthorizationReform.RecordDate, dbo.TblCardAuthorizationReform.ClientName,"
     MySQL = MySQL & "                     dbo.TblCardAuthorizationReform.Telephone, dbo.TblCardAuthorizationReform.Posted, dbo.TblCardAuthorizationReform.CarTypeID,"
     MySQL = MySQL & "                     dbo.TBLCarTypes.name AS CarName, dbo.TBLCarTypes.namee AS CarNameE, dbo.TblCardAuthorizationReform.CarModelID, dbo.TblCarModels.Model,"
     MySQL = MySQL & "                     dbo.TblCarModels.ModelE, dbo.TblCardAuthorizationReform.PlateNo, dbo.TblCardAuthorizationReform.BranchID, dbo.TblBranchesData.branch_name,"
     MySQL = MySQL & "                     dbo.TblBranchesData.branch_namee, dbo.TblCardAuthorizationReform.ColorID, dbo.TblColor.name AS Color, dbo.TblColor.namee AS ColorE,"
     MySQL = MySQL & "                     dbo.TblCardAuthorizationReform.YearFact, dbo.TblCardAuthorizationReform.OrderStatus, dbo.TblCardAuthorizationReform.Accept,"
     MySQL = MySQL & "                     dbo.TblCardAuthorizationReform.EndDate, dbo.TblCardAuthorizationReform.subcar1, dbo.TblCardAuthorizationReform.subcar2,"
     MySQL = MySQL & "                     dbo.TblCardAuthorizationReform.subcar3, dbo.TblCardAuthorizationReform.subcar4, dbo.TblCardAuthorizationReform.subcar5,"
     MySQL = MySQL & "                     dbo.TblCardAuthorizationReform.subcar6, dbo.TblCardAuthorizationReform.subcar7, dbo.TblCardAuthorizationReform.subcar8,"
     MySQL = MySQL & "                     dbo.TblCardAuthorizationReform.subcar9, dbo.TblCardAuthorizationReform.subcar10, dbo.TblCardAuthorizationReform.Month_Day,"
     MySQL = MySQL & "                     dbo.TblCardAuthorizationReform.Granty, dbo.TblCardAuthorizationReform.DateStartG, dbo.TblCardAuthorizationReform.DateEndG,"
     MySQL = MySQL & "                     dbo.TblCardAuthorizationReform.CarMeter, dbo.TblCardAuthorizationReform.LongGranty, dbo.TblCardAuthorizationReform.PayFirst,"
     MySQL = MySQL & "                     dbo.TblCardAuthorizationReform.AmountAccept, dbo.TblCardAuthorizationReform.Complaint, dbo.TblCardAuthorizationReform.Noteinitial,"
     MySQL = MySQL & "                     dbo.TblCardAuthorizationReform.Shaseh, dbo.TblCardAuthorizationReform.NotAccept, dbo.TblCardAuthorizationReform.EmpID2,"
     MySQL = MySQL & "                     dbo.TblCardAuthorizationReform.EmpID1, dbo.TblCardAuthorizationReform.EmpID AS EmPPID, dbo.TblCardAuthorizationReform.typerequest,"
     MySQL = MySQL & "                     dbo.TblCardAuthorizationReform.FitterID, dbo.TblUsers.UserName, dbo.TblCardAuthorizationReform.ClientCode, dbo.TblCardAuthorizationReform.mobile,"
     MySQL = MySQL & "                     dbo.TblCardAuthorizationReform.Cash, dbo.TblCardAuthorizationReform.Accoun, dbo.TblCardAuthorizationReform.credit, dbo.TblCardAuthorizationReform.box,"
     MySQL = MySQL & "                     dbo.TblCardAuthorizationReform.fax, dbo.TblCardAuthorizationReform.email, dbo.TblCardAuthorizationReform.address, dbo.TblCardAuthorizationReform.boxzip,"
     MySQL = MySQL & "                     dbo.TblCardAuthorizationReform.codereg, dbo.TblCardAuthorizationReform.codedoor, dbo.TblCardAuthorizationReform.typereg,"
     MySQL = MySQL & "                     dbo.TblCardAuthorizationReform.DateEnter AS DateEnterR, dbo.TblCardAuthorizationReform.persons, dbo.TblCardAuthorizationReform.Companies,"
     MySQL = MySQL & "                     dbo.TblCardAuthorizationReform.driver, dbo.TblCardAuthorizationReform.DateAcutExite, dbo.TblCardAuthorizationReform.DateExptExit,"
     MySQL = MySQL & "                     dbo.TblCardAuthorizationReform.TimeAcutExite, dbo.TblCardAuthorizationReform.TimeExptExit, dbo.TblCardAuthorizationReform.DateExit AS DateExitR,"
     MySQL = MySQL & "                     dbo.TblCardAuthorizationReform.subcar11, dbo.TblCardAuthorizationReform.subcar12, dbo.TblCardAuthorizationReform.subcar13,"
     MySQL = MySQL & "                     dbo.TblCardAuthorizationReform.subcar14, dbo.TblCardAuthorizationReform.ResonUnderWait, dbo.TblCardAuthorizationReform.Remarkcar,"
     MySQL = MySQL & "                     dbo.TblCardAuthorizationReform.Payed AS PayedR, dbo.TblCardAuthorizationReform.finish AS finishR, dbo.TblCardAuthorizationReform.PrivateCop,"
     MySQL = MySQL & "                     dbo.TblCardAuthorizationReform.ReComentClient, dbo.TblCardAuthorizationReform.wait, dbo.TblCardAuthorizationReform.notAcepted,"
     MySQL = MySQL & "                     dbo.TblCardAuthorizationReform.NoteSerial, dbo.TblCardAuthorizationReform.CodeComputer, dbo.TblCardAuthorizationReform.ID,"
     MySQL = MySQL & "                     dbo.TblCardAuthorizationReform.TypeCustomer, dbo.TblCardAuthorizationReform.OverKM, dbo.TblCustemers.CusName, TblCustemers.VATNO, dbo.TblCustemers.CusNamee,"
     MySQL = MySQL & "                     dbo.TblCardAuthorizationReform.SendSMS, dbo.TblCardAuthorizationReform.TypeOrder, dbo.TblCardAuthorizationReform.WorkOrder,"
     MySQL = MySQL & "                     dbo.TblCardAuthorizationReform.ShowPriceOrder, dbo.TblCardAuthorizationReform.AuthoOrder, dbo.TblCardAuthorizationReform.LastWorOrder,"
     MySQL = MySQL & "                     dbo.TblCustemers.Fullcode, dbo.TblCustemers.CustGID, dbo.TblCustemers.ExpireDateH, dbo.TblCardAuthorizationReform.RecordeTime,"
     MySQL = MySQL & "                     dbo.TblCardAuthorizationReform.CarMetarOut,TblHandWages2.Id TblHandWages2ID"
    
     MySQL = MySQL & "                     FROM            TBLCarTypes RIGHT OUTER JOIN"
      MySQL = MySQL & "                                             TblColor RIGHT OUTER JOIN"
     MySQL = MySQL & "                                              TblCardAuthorizationReform LEFT OUTER JOIN"
     MySQL = MySQL & "                                              TblCustemers ON TblCardAuthorizationReform.CusID = TblCustemers.CusID LEFT OUTER JOIN"
     MySQL = MySQL & "                                              TblUsers ON TblCardAuthorizationReform.FitterID = TblUsers.UserID ON TblColor.Id = TblCardAuthorizationReform.ColorID LEFT OUTER JOIN"
     MySQL = MySQL & "                                              TblBranchesData ON TblCardAuthorizationReform.BranchID = TblBranchesData.branch_id LEFT OUTER JOIN"
     MySQL = MySQL & "                                              TblCarModels ON TblCardAuthorizationReform.CarModelID = TblCarModels.Id LEFT OUTER JOIN"
     MySQL = MySQL & "                                              TblEmpDepartments RIGHT OUTER JOIN"
     MySQL = MySQL & "                                              TblCardAuthorizationReformDetails LEFT OUTER JOIN"
     MySQL = MySQL & "                                              TblMaintenanceWork ON TblCardAuthorizationReformDetails.Mainte = TblMaintenanceWork.Id ON"
     MySQL = MySQL & "                                              TblEmpDepartments.DeparmentID = TblCardAuthorizationReformDetails.Deptid LEFT OUTER JOIN"
      MySQL = MySQL & "                                             TblEmployee AS TblEmployee_1 ON TblCardAuthorizationReformDetails.empsuper = TblEmployee_1.Emp_ID LEFT OUTER JOIN"
      MySQL = MySQL & "                                             TblEmployee AS TblEmployee_2 ON TblCardAuthorizationReformDetails.EmpID = TblEmployee_2.Emp_ID ON"
      MySQL = MySQL & "                                             TblCardAuthorizationReform.ID = TblCardAuthorizationReformDetails.ID ON TBLCarTypes.id = TblCardAuthorizationReform.CarTypeID"
      MySQL = MySQL & "                                           LEFT OUTER JOIN TblHandWages"
        MySQL = MySQL & "                                           ON TblHandWages.OrDer_no =   TblCardAuthorizationReform.WorkOrder "
    
        MySQL = MySQL & "                                           LEFT OUTER JOIN TblHandWages2 "
        MySQL = MySQL & "                                           ON TblHandWages.Id = TblHandWages2.MasterID "
        MySQL = MySQL & "  Where (TblHandWages.Id  =  " & val(TxtSerial1(1).Text) & ") "
     'and (dbo.TblCardAuthorizationReformDetails.type=0)"

     ' RsDetails1.Open MySQL, Cn, adOpenStatic, adLockReadOnly, adCmdText
    
    
    
     
                
                    
        StrSQL = "SELECT     Transaction_Details.ShowPrice,dbo.Transactions.Transaction_ID, dbo.Transactions.Transaction_Date, dbo.Transactions.Transaction_Type, dbo.Transactions.NoteSerial1, "
        StrSQL = StrSQL & "                      dbo.Transactions.Transaction_Serial, dbo.Transactions.Transaction_HijriDate, dbo.Transactions.TransactionComment, dbo.Transactions.OpOrderID,"
        StrSQL = StrSQL & "                      dbo.Transactions.OldOpOrderID, dbo.Transaction_Details.UnitId,dbo.Transaction_Details.OperPrice, dbo.Transaction_Details.ID, dbo.Transaction_Details.ShowQty, dbo.Transaction_Details.Item_ID,"
        StrSQL = StrSQL & "                      dbo.TblItems.itemname , dbo.TblItems.ItemNamee, dbo.TblItems.fullcode , dbo.Transaction_Details.showPrice"
        StrSQL = StrSQL & " ,ShowPrice2 = (SELECT Top 1 TblItemsUnits.UnitSalesPrice"
        StrSQL = StrSQL & "                 From TblItemsUnits"
        StrSQL = StrSQL & "                 Where ItemID = Transaction_Details.Item_ID"
        StrSQL = StrSQL & "                        AND UnitID           = Transaction_Details.UnitId  )"
        StrSQL = StrSQL & " FROM         dbo.TblItems RIGHT OUTER JOIN"
        StrSQL = StrSQL & "                      dbo.Transaction_Details ON dbo.TblItems.ItemID = dbo.Transaction_Details.Item_ID RIGHT OUTER JOIN"
        StrSQL = StrSQL & "                      dbo.Transactions ON dbo.Transaction_Details.Transaction_ID = dbo.Transactions.Transaction_ID"
        StrSQL = StrSQL & " Where (dbo.Transactions.Transaction_Type = 21) And  (Transactions.order_no = '" & val(TXTOrDer_no(0).Text) & "')"
            
            
            
            
            RsData2.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText
            
             

    
     
     
    
     
     
    
       If SystemOptions.UserInterface = ArabicInterface Then
                StrFileName = App.path & "\REPORTS\REPORTS NEW\" & "TblHandWages2.rpt"
            Else
                StrFileName = App.path & "\REPORTS\REPORTS NEW\" & "TblHandWages2.rpt"
        End If
        
        If Dir(StrFileName) = "" Then
            'GetMsgs 139, vbExclamation  RepCardAutintcationShow
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
           ' xReport.ParameterFields(15).AddCurrentValue Me.DcboFitter.text
            ' xReport.ParameterFields(2).AddCurrentValue RPTComment_Arabic
            StrReportTitle = "" '& StrAccountName
            'If Me.DTPickerAccFrom.value <> Empty Or Me.DTPickerAccFrom.value <> Null Then
            '    StrReportTitle = StrReportTitle + " »œ«Ì… „‰ " & Format(Me.DTPickerAccFrom.value, "yyyy/M/d") & ""
            'End If
            'If Me.DTPickerAccTo.value <> Empty Or Me.DTPickerAccTo.value <> Null Then
            '    StrReportTitle = StrReportTitle + " ≈·Ï " & Format(Me.DTPickerAccTo.value, "yyyy/M/d") & " "
            'End If
        Else
     
            xReport.ParameterFields(1).AddCurrentValue cCompanyInfo.ArabCompanyName ' RPTCompany_Name_Eng
        
          '  xReport.ParameterFields(15).AddCurrentValue Me.DcboFitter.text
            xReport.ParameterFields(4).AddCurrentValue get_branch_name(val(my_branch))
            StrReportTitle = ""
            'If Me.DTPickerAccFrom.value <> Empty Or Me.DTPickerAccFrom.value <> Null Then
            '    StrReportTitle = StrReportTitle + " From Date " & (Me.DTPickerAccFrom.value) & ""
            'End If
            'If Me.DTPickerAccTo.value <> Empty Or Me.DTPickerAccTo.value <> Null Then
            '    StrReportTitle = StrReportTitle + " To Date :  " & (Me.DTPickerAccTo.value) & ""
            'End If
        End If
    
        xReport.ParameterFields(3).AddCurrentValue user_name
        
            xReport.OpenSubreport("Out").Database.SetDataSource RsData2
           ' xReport.OpenSubreport("RepCarBillMaintene").Database.SetDataSource RsData3
           ' xReport.OpenSubreport("RepCar").Database.SetDataSource RsData3
                Dim i As Integer
                 xReport.EnableParameterPrompting = False
             For i = 1 To xReport.ParameterFields.count
                 Select Case xReport.ParameterFields.Item(i).ParameterFieldName
                 
                Case "TotalNet"
                    xReport.ParameterFields.Item(i).AddCurrentValue "" & WriteNo(Format(val(val(txtNet)), "0.00"), 0, True, ".") & ""
                Case "TotalNet2"
                    'xReport.ParameterFields.Item(i).AddCurrentValue "" & (val(val(LbToTalExtra.Caption) + (val(lbl(23))) + (val(Me.lbTotalMente.Caption)) * 1.05)) & ""
                    
                 Case "txtTotalInvoice"
                     xReport.ParameterFields.Item(i).AddCurrentValue "" & val(txtTotalInvoice)
                Case "txtDiscValueInvoice"
                     xReport.ParameterFields.Item(i).AddCurrentValue "" & val(txtDiscValueInvoice)
                Case "txtTotalInvoiceBVat"
                     xReport.ParameterFields.Item(i).AddCurrentValue "" & val(txtTotalInvoiceBVat)
                Case "txtVat2Invoice"
                     xReport.ParameterFields.Item(i).AddCurrentValue "" & val(txtVat2Invoice)
                Case "txtNetInvoice"
                     xReport.ParameterFields.Item(i).AddCurrentValue "" & val(txtNetInvoice)
                
                Case "txtTotal2"
                     xReport.ParameterFields.Item(i).AddCurrentValue "" & val(txtTotal2)
                Case "txtDiscValue"
                     xReport.ParameterFields.Item(i).AddCurrentValue "" & val(txtDiscValue)
                Case "txtTotal"
                     xReport.ParameterFields.Item(i).AddCurrentValue "" & val(txtTotal)
                Case "txtVat2"
                     xReport.ParameterFields.Item(i).AddCurrentValue "" & val(txtVat2)
                Case "txtNet"
                     xReport.ParameterFields.Item(i).AddCurrentValue "" & val(txtNet)
                
                Case "txtNetInvoice"
                     xReport.ParameterFields.Item(i).AddCurrentValue "" & val(txtNetInvoice)
                Case "txtGeneralTotal"
                     xReport.ParameterFields.Item(i).AddCurrentValue "" & val(txtGeneralTotal)
                Case "txtTotalDisc"
                     xReport.ParameterFields.Item(i).AddCurrentValue "" & val(txtTotalDisc)
                Case "txtTotalBVat"
                     xReport.ParameterFields.Item(i).AddCurrentValue "" & val(txtTotalBVat)
                Case "txtTotalVat"
                     xReport.ParameterFields.Item(i).AddCurrentValue "" & val(txtTotalVat)
                Case "txtTotalNet"
                     xReport.ParameterFields.Item(i).AddCurrentValue "" & val(txtTotalNet)
                     
                     
                 Case "TotalVat"
                     xReport.ParameterFields.Item(i).AddCurrentValue "" & val(txtVat2)
                     
                 Case "DisckPercent"
                     xReport.ParameterFields.Item(i).AddCurrentValue "" & txtDiscPercent & ""
                 Case "TotalPriceBeDisk"
                     xReport.ParameterFields.Item(i).AddCurrentValue "" & (txtTotalInvoice) & ""
                 Case "TotalAfterDisc"
                     xReport.ParameterFields.Item(i).AddCurrentValue "0"
                     
                Case "TotalHand"
                     xReport.ParameterFields.Item(i).AddCurrentValue CStr(val(Me.txtTotal2))
                Case "VATRegNo"
                    If SystemOptions.VATNoAccordActivity = False Then
                        xReport.ParameterFields(i).AddCurrentValue cCompanyInfo.VATRegNo
                    Else
                        xReport.ParameterFields(i).AddCurrentValue GetRegVATNo(val(Dcbranch(mIndex).BoundText))
                    End If
                 End Select
             Next i
            
        
          '  xReport.ParameterFields(4).AddCurrentValue WriteNo(Format(val(TxtAdvanceValue.text), "0.00"), 0, True, ".")
           ' xReport.ParameterFields(6).AddCurrentValue val(lbl(23).Caption)
           '  xReport.ParameterFields(7).AddCurrentValue DBIssueDate.value
    '    xReport.ParameterFields(8).AddCurrentValue IIf(IsNumeric(fg.TextMatrix(Me.fg.FixedRows, fg.ColIndex("PartValue"))), val(fg.TextMatrix(Me.fg.FixedRows, fg.ColIndex("PartValue"))), 0)
    ' xReport.ParameterFields(9).AddCurrentValue val(lbl(22).Caption)
     ' xReport.ParameterFields(10).AddCurrentValue val(TxtDiscount.text)
      ' xReport.ParameterFields(11).AddCurrentValue txtDiscountDES.text
      Dim Total As String
       Dim dif As String
      Dim totl As Double
      
       xReport.ParameterFields(12).AddCurrentValue "0"
          xReport.ParameterFields(13).AddCurrentValue "0"
            xReport.ParameterFields(14).AddCurrentValue (txtTotal)
            xReport.ParameterFields(15).AddCurrentValue "0"
           
    '    xReport.ParameterFields(5).AddCurrentValue ToHijriDate(RsData("notedate").value)
        xReport.reporttitle = StrReportTitle
        xReport.EnableParameterPrompting = False
        xReport.ApplicationName = App.title
        xReport.ReportAuthor = App.title
        Set CViewer = New ClsReportViewer
        CViewer.FireReport xReport, WindowTarget, "", , , , StrFileName
    
        RsData.Close
        Set RsData = Nothing
        Screen.MousePointer = vbDefault


Else



         

    Dim EmpReport As ClsEmployeeReport
    

    Dim rs As ADODB.Recordset
    
    Set cCompanyInfo = New ClsCompanyInfo
 Dim sql As String

 sql = "SELECT     dbo.TblVisitScreen.ScreenID, dbo.TblUserScreen.Name AS ScreenName, dbo.TblUserScreen.Mdiol, dbo.TblUserScreen.FlgWork, dbo.TblVisitScreen.EmpRemark, "
 sql = sql & "                      dbo.TblVisitScreen.CusRemark, dbo.TblVisitScreen.RecordDate, dbo.TblVisitScreen.EmpPass, dbo.TblVisitScreen.UserPass, dbo.TblVisitScreen.ID,"
 sql = sql & "                     dbo.TblVisitScreen.UserID , dbo.TblUsers.UserName, dbo.TblVisitScreen.EmpUsrID, dbo.TblUserComp.Name, dbo.TblUserComp.NameE"
 sql = sql & " FROM         dbo.TblUserComp RIGHT OUTER JOIN"
 sql = sql & "                     dbo.TblVisitScreen ON dbo.TblUserComp.ID = dbo.TblVisitScreen.EmpUsrID LEFT OUTER JOIN"
 sql = sql & "                     dbo.TblUsers ON dbo.TblVisitScreen.UserID = dbo.TblUsers.UserID LEFT OUTER JOIN"
 sql = sql & "                     dbo.TblUserScreen ON dbo.TblVisitScreen.ScreenID = dbo.TblUserScreen.ID"
If Ind = 0 Then
sql = sql & " where dbo.TblVisitScreen.ID=" & val(TxtVac_ID.Text) & " "
End If
    Set rs = New ADODB.Recordset
    rs.Open sql, Cn, adOpenStatic, adLockPessimistic, adCmdText

    If SystemOptions.UserInterface = ArabicInterface Then
        Set xReport = xApp.OpenReport(App.path & "\reports\REPORTS NEW\RepVisitScreen.rpt")
    Else

        Set xReport = xApp.OpenReport(App.path & "\reports\REPORTS NEW\RepVisitScreen.rpt")
    End If


    xReport.Database.SetDataSource rs
     Dim cAccountReport As New ClsReportViewer
    Dim FrmReport As New FrmReportViewer
    FrmReport.CRViewer.ReportSource = xReport
    FrmReport.txtPath = (App.path & "\reports\REPORTS NEW\RepVisitScreen.rpt")
  '  xReport.reporttitle = "  «·⁄ﬁÊœ «·”«»ﬁ…"

       xReport.ParameterFields(1).AddCurrentValue cCompanyInfo.ArabCompanyName
       xReport.ParameterFields(2).AddCurrentValue user_name
    FrmReport.CRViewer.ViewReport
 cAccountReport.CreateLogo xReport

    FrmReport.show
    Screen.MousePointer = vbDefault
End If

End Function

Private Sub ISButton2_Click(Index As Integer)
print_report , Index
End Sub

Private Sub TxtEmpPass_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyReturn Then
Me.DcbEmpUsrID.BoundText = CheckPassworUserComp()

End If
End Sub

Private Sub TxtVac_ID_Change()
    Dim TxtMod As String
    TxtMod = TxtModFlg.Text
    TxtModFlg.Text = ""
    TxtModFlg = TxtMod
End Sub

Public Function FindRec(ByVal RecId As Long)
    
    'RsSavRec.Filter = adFilterNone
    
    
    
        On Error GoTo ErrTrap
    RsSavRec.Find "id=" & RecId, , adSearchForward, 1
    If mIndex2 = 0 Then mIndex2 = mIndex
    If Not (RsSavRec.EOF) Then
        If mIndex = 0 Then
            FiLLTXT
        ElseIf mIndex = 1 Then
            FiLLTXT1
        ElseIf mIndex = 2 Then
            FiLLTXT2
            'FillGridWithData2
        ElseIf mIndex = 3 Then
            FiLLTXT3
            GenreateItems
        End If
    End If

    Exit Function
ErrTrap:

    If RsSavRec.EditMode <> adEditNone Then
        RsSavRec.CancelUpdate
        If mIndex = 0 Then
            BtnUndo_Click
        Else
            Btn_Undo_Click (mIndex)
       
        End If
        
        
    End If
End Function

Private Sub TxtModFlg_Change()
btnDelete.Enabled = False
    If TxtModFlg.Text = "N" Then
       ' Frm2.Enabled = True
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
     '   Frm2.Enabled = False
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
       ' Frm2.Enabled = True
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
    My_SQL = "SELECT     dbo.TblVisitScreen.ScreenID, dbo.TblUserScreen.Name AS ScreenName, dbo.TblUserScreen.Mdiol, dbo.TblUserScreen.FlgWork, dbo.TblVisitScreen.EmpRemark, "
    My_SQL = My_SQL & "                  dbo.TblVisitScreen.CusRemark, dbo.TblVisitScreen.RecordDate, dbo.TblVisitScreen.EmpPass, dbo.TblVisitScreen.UserPass, dbo.TblVisitScreen.ID,"
    My_SQL = My_SQL & "                  dbo.TblVisitScreen.UserID , dbo.TblUsers.UserName, dbo.TblVisitScreen.EmpUsrID, dbo.TblUserComp.Name, dbo.TblUserComp.NameE"
    My_SQL = My_SQL & "  FROM         dbo.TblUserComp RIGHT OUTER JOIN"
    My_SQL = My_SQL & "                  dbo.TblVisitScreen ON dbo.TblUserComp.ID = dbo.TblVisitScreen.EmpUsrID LEFT OUTER JOIN"
    My_SQL = My_SQL & "                  dbo.TblUsers ON dbo.TblVisitScreen.UserID = dbo.TblUsers.UserID LEFT OUTER JOIN"
    My_SQL = My_SQL & "                  dbo.TblUserScreen ON dbo.TblVisitScreen.ScreenID = dbo.TblUserScreen.ID"
    My_SQL = My_SQL & " order by  dbo.TblVisitScreen.ID                  "
    rs.Open My_SQL, Cn, adOpenKeyset, adLockReadOnly, adCmdText

    With Me.Grid
        .Rows = 2
        .Clear flexClearScrollable

        If rs.RecordCount > 0 Then
            .Rows = rs.RecordCount + 1
            rs.MoveFirst

            For i = 1 To .Rows - 1
                .TextMatrix(i, .ColIndex("Ser")) = i
                .TextMatrix(i, .ColIndex("id")) = IIf(IsNull(rs.Fields("id").value), "", rs.Fields("id").value)
                .TextMatrix(i, .ColIndex("RecordDate")) = IIf(IsNull(rs.Fields("RecordDate").value), "", rs.Fields("RecordDate").value)
                '.TextMatrix(i, .ColIndex("FromTime")) = IIf(IsNull(rs.Fields("FromTime").value), "", rs.Fields("FromTime").value)
               '.TextMatrix(i, .ColIndex("ToTime")) = IIf(IsNull(rs.Fields("ToTime").value), "", rs.Fields("ToTime").value)
                .TextMatrix(i, .ColIndex("ScreenName")) = IIf(IsNull(rs.Fields("ScreenName").value), "", rs.Fields("ScreenName").value)
                .TextMatrix(i, .ColIndex("CusRemark")) = IIf(IsNull(rs.Fields("CusRemark").value), "", rs.Fields("CusRemark").value)
                .TextMatrix(i, .ColIndex("EmpRemark")) = IIf(IsNull(rs.Fields("EmpRemark").value), "", rs.Fields("EmpRemark").value)
               If SystemOptions.UserInterface = ArabicInterface Then
               .TextMatrix(i, .ColIndex("Name")) = IIf(IsNull(rs.Fields("Name").value), "", rs.Fields("Name").value)
               Else
               .TextMatrix(i, .ColIndex("Name")) = IIf(IsNull(rs.Fields("NameE").value), "", rs.Fields("NameE").value)
               End If
                rs.MoveNext
            Next

            rs.Close
        End If

        .RowHeight(-1) = 300
    End With

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





Public Sub FiLLRec2()
    On Error GoTo ErrTrap

    RsSavRec.Fields("name").value = IIf(txtName(mIndex).Text <> "", Trim(txtName(mIndex).Text), Null)
    RsSavRec.Fields("namee").value = IIf(txtNamee(mIndex).Text <> "", Trim(txtNamee(mIndex).Text), Null)
    

    RsSavRec.update
    MsgBox " „  ⁄„·Ì… «·Õ›Ÿ »‰Ã«Õ", vbInformation + vbOKOnly + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
    FillGridWithData2
    TxtModFlg2(mIndex) = "R"

    Exit Sub
ErrTrap:

    If RsSavRec.EditMode <> adEditNone Then
        RsSavRec.CancelUpdate
    End If

End Sub


Public Sub FiLLRec3()
    On Error GoTo ErrTrap

  
    If TxtModFlg2(mIndex).Text = "N" Then
           ' TxtSerial1(mIndex).Text = CStr(new_id("TblAdditionsAssest", "ID", "", True))

       
        RsSavRec.AddNew
        TxtSerial1(mIndex).Text = new_id("TblLensesTypes", "id", "")
        RsSavRec.Fields("ID").value = val(TxtSerial1(mIndex).Text)
    End If
    
    RsSavRec.Fields("name").value = IIf(txtName(mIndex).Text <> "", Trim(txtName(mIndex).Text), Null)
    RsSavRec.Fields("namee").value = IIf(txtNamee(mIndex).Text <> "", Trim(txtNamee(mIndex).Text), Null)
    

    RsSavRec.Fields("GroupId").value = IIf(cmbGroupId.Text <> "", Trim(cmbGroupId.BoundText), Null)
    RsSavRec.Fields("UnitID").value = IIf(cmbUnitID.Text <> "", Trim(cmbUnitID.BoundText), Null)
    
    RsSavRec.Fields("FromSPH").value = IIf(DCBoMain(2).Text <> "", Trim(DCBoMain(2).BoundText), Null)
    RsSavRec.Fields("TOSPH").value = IIf(DCBoMain(5).Text <> "", Trim(DCBoMain(5).BoundText), Null)
    RsSavRec.Fields("FROMCYL").value = IIf(DCBoMain(3).Text <> "", Trim(DCBoMain(3).BoundText), Null)
    RsSavRec.Fields("TOCYL").value = IIf(DCBoMain(6).Text <> "", Trim(DCBoMain(6).BoundText), Null)
    RsSavRec.Fields("Price").value = val(txtPrice)
    
    
    RsSavRec("Flag").value = cmbFlag(0).ListIndex
   
   'RsSavRec("RecType").value = cmbRecType.ListIndex


    RsSavRec.update
    
    Command1_Click
    MsgBox " „  ⁄„·Ì… «·Õ›Ÿ »‰Ã«Õ", vbInformation + vbOKOnly + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
    FillGridWithData3
    FiLLTXT3
    
    TxtModFlg2(mIndex) = "R"

    Exit Sub
ErrTrap:

    If RsSavRec.EditMode <> adEditNone Then
        RsSavRec.CancelUpdate
    End If

End Sub


Public Sub FillGridWithData2()

    On Error GoTo ErrTrap

    Dim i As Integer
    Dim rs As ADODB.Recordset
    Dim My_SQL As String

    Set rs = New ADODB.Recordset
    My_SQL = "select * From TblOffice order by id"
    rs.Open My_SQL, Cn, adOpenKeyset, adLockReadOnly, adCmdText

    With Me.Grid2
        .Rows = 2
        .Clear flexClearScrollable

        If rs.RecordCount > 0 Then
            .Rows = rs.RecordCount + 1
            rs.MoveFirst

            For i = 1 To .Rows - 1
                .TextMatrix(i, .ColIndex("Ser")) = i
             
                .TextMatrix(i, .ColIndex("name")) = IIf(IsNull(rs.Fields("name").value), "", rs.Fields("name").value)
                .TextMatrix(i, .ColIndex("namee")) = IIf(IsNull(rs.Fields("namee").value), "", rs.Fields("namee").value)
                
               
                .TextMatrix(i, .ColIndex("id")) = IIf(IsNull(rs.Fields("id").value), "", rs.Fields("id").value)
                
            
                '    .Cell(flexcpBackColor, .Rows - 1, 1, .Rows - 1, .Cols - 1) = vbYellow
                rs.MoveNext
            Next

            rs.Close
        End If

        .RowHeight(-1) = 300
    End With

ErrTrap:
End Sub



Public Sub FillGridWithData3()

    On Error GoTo ErrTrap

    Dim i As Integer
    Dim rs As ADODB.Recordset
    Dim My_SQL As String

    Set rs = New ADODB.Recordset
    My_SQL = "select * From TblLensesTypes order by id"
    rs.Open My_SQL, Cn, adOpenKeyset, adLockReadOnly, adCmdText

    With Me.Grid3
        .Rows = 2
        .Clear flexClearScrollable

        If rs.RecordCount > 0 Then
            .Rows = rs.RecordCount + 1
            rs.MoveFirst

            For i = 1 To .Rows - 1
                .TextMatrix(i, .ColIndex("Ser")) = i
             
                .TextMatrix(i, .ColIndex("name")) = IIf(IsNull(rs.Fields("name").value), "", rs.Fields("name").value)
                .TextMatrix(i, .ColIndex("namee")) = IIf(IsNull(rs.Fields("namee").value), "", rs.Fields("namee").value)
                
               
                .TextMatrix(i, .ColIndex("id")) = IIf(IsNull(rs.Fields("id").value), "", rs.Fields("id").value)
                
            
                '    .Cell(flexcpBackColor, .Rows - 1, 1, .Rows - 1, .Cols - 1) = vbYellow
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
    Frame1(mIndex).Enabled = False
    TxtSerial1(mIndex).Text = IIf(IsNull(RsSavRec.Fields("id").value), "", RsSavRec.Fields("id").value)
    txtName(mIndex).Text = IIf(IsNull(RsSavRec.Fields("name").value), "", RsSavRec.Fields("name").value)
    txtNamee(mIndex).Text = IIf(IsNull(RsSavRec.Fields("nameE").value), "", RsSavRec.Fields("nameE").value)
    
    LabCurr_Rec(mIndex).Caption = RsSavRec.AbsolutePosition
    LabCount_Rec(mIndex).Caption = RsSavRec.RecordCount

    With Grid2

        For i = 1 To .Rows - 1

            If Trim(TxtSerial1(mIndex).Text) = .TextMatrix(i, .ColIndex("id")) Then
                TxtSerial1(mIndex).Text = .TextMatrix(i, .ColIndex("Ser"))
                .Row = i
                Exit Sub
            End If

        Next

    End With

ErrTrap:

End Sub





Private Function GetNewCode(LngParentGroupID As Long, Optional ByVal mTableName As String = "", Optional ByVal mTableGroupName As String = "Groups", Optional ByVal mFieldGroup As String = "GroupID") As String
    Dim rs As ADODB.Recordset
    Dim StrSQL As String
    Dim StrParentCode  As String
    Dim StrNewGroupCode As String
    Dim StrLastGroupCode As String
    Dim IntTemp As String
    If mTableName = "" Then
        mTableName = "Groups"
    End If
    On Error GoTo ErrTrap
    StrSQL = "Select Max(Code) Code From " & mTableName & "  Where " & mFieldGroup & " =" & LngParentGroupID & ""
    Set rs = New ADODB.Recordset
    rs.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

    If Not (rs.EOF Or rs.BOF) Then
        StrParentCode = IIf(IsNull(rs("Code").value), "", rs("Code").value)
    Else
        StrParentCode = "000"
    End If

     Set rs = New ADODB.Recordset
    StrSQL = "Select * From " & mTableGroupName & "   Where GroupID=" & LngParentGroupID & " Order By GroupID"
    rs.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText
    Dim mTmpGroup2  As String
    If Not rs.BOF Then
        StrNewGroupCode = rs!code & ""
        mTmpGroup2 = Replace(StrParentCode, StrNewGroupCode, "")
    End If
    If Trim(mTmpGroup2) = "" Then mTmpGroup2 = "000"
    rs.Close
    Dim mTmp As Long
    mTmp = val(mTmpGroup2) + 1
    If Len(CStr(mTmp)) = 1 Then
        StrParentCode = "00" & mTmp
    ElseIf Len(CStr(mTmp)) = 2 Then
        StrParentCode = "0" & mTmp
    ElseIf Len(CStr(mTmp)) = 3 Then
        StrParentCode = "" & mTmp
    End If
    Set rs = New ADODB.Recordset
    StrSQL = "Select * From " & mTableGroupName & "   Where GroupID=" & LngParentGroupID & " Order By GroupID"
    rs.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

    If rs.BOF Or rs.EOF Then
        StrNewGroupCode = StrParentCode & "1"
    Else
        rs.MoveLast
        StrLastGroupCode = IIf(IsNull(rs("Code").value), "", rs("Code").value)
        IntTemp = val(mId(StrLastGroupCode, Len(StrParentCode) + 1))
        StrNewGroupCode = StrLastGroupCode & StrParentCode
    End If

    rs.Close
    Set rs = Nothing
    GetNewCode = StrNewGroupCode
    Exit Function
ErrTrap:
End Function


