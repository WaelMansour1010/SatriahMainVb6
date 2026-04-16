VERSION 5.00
Object = "{C115893A-A3BF-43AF-B28D-69DB846077F3}#1.0#0"; "vsflex8u.ocx"
Object = "{0AFE7BE0-11B7-4A3E-978D-D4501E9A57FE}#1.0#0"; "c1sizer.ocx"
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{FE5DCFAD-BC1D-11D2-94CF-004005455FAA}#1.4#0"; "ImpulseButton.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#12.0#0"; "SuiteCtrls.ocx"
Begin VB.Form FrmVizitScreen 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   9255
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   17610
   Icon            =   "FrmVizitScreen.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   RightToLeft     =   -1  'True
   ScaleHeight     =   9255
   ScaleWidth      =   17610
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
      Top             =   120
      Width           =   18780
      _cx             =   33126
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
      Caption         =   " œ—Ì» «·⁄„·«¡|«ÃÊ— «·Ìœ|«·„ﬂ« » «·„›Ê÷…| ⁄—Ì› «·⁄œ”« | ‰»Â«  «·ÿ·»«  «·œ«Œ·Ì…|«·„⁄—÷| ‰»ÌÂ«  «·„⁄„·|a|«⁄„«— «·œÌÊ‰"
      Align           =   0
      CurrTab         =   8
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
         Left            =   -21435
         TabIndex        =   1
         TabStop         =   0   'False
         Top             =   45
         Width           =   18690
         _cx             =   32967
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
            Width           =   16875
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
            Left            =   1680
            RightToLeft     =   -1  'True
            TabIndex        =   5
            Top             =   4890
            Width           =   16740
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
                  Format          =   143589377
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
            Left            =   25935
            TabIndex        =   2
            Top             =   765
            Width           =   18555
            _cx             =   32729
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
            Width           =   16725
            _cx             =   29501
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
            Left            =   135
            TabIndex        =   58
            Top             =   750
            Width           =   18285
            _cx             =   32253
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
         Left            =   -21135
         TabIndex        =   3
         TabStop         =   0   'False
         Top             =   45
         Width           =   18690
         _cx             =   32967
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
            Left            =   7245
            TabIndex        =   182
            Top             =   900
            Width           =   1680
         End
         Begin VB.TextBox TxtNoteSerial1 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFC0&
            Height          =   315
            Left            =   15480
            Locked          =   -1  'True
            RightToLeft     =   -1  'True
            TabIndex        =   129
            Top             =   990
            Width           =   1680
         End
         Begin VB.CommandButton cmdPrintNote 
            Caption         =   "ÿ»«⁄Â «·ﬁÌœ"
            Height          =   450
            Left            =   3900
            RightToLeft     =   -1  'True
            TabIndex        =   127
            Top             =   7020
            Width           =   2790
         End
         Begin VB.CommandButton cmdDelNote 
            Caption         =   "Õ–› «·ﬁÌœ "
            Height          =   450
            Left            =   12000
            RightToLeft     =   -1  'True
            TabIndex        =   126
            Top             =   6960
            Visible         =   0   'False
            Width           =   3345
         End
         Begin VB.CommandButton CmdCreateV2 
            Caption         =   "≈‰‘«¡ «·ﬁÌœ "
            Height          =   450
            Left            =   15480
            RightToLeft     =   -1  'True
            TabIndex        =   125
            Top             =   6960
            Visible         =   0   'False
            Width           =   2940
         End
         Begin VB.TextBox TxtNoteSerial 
            Alignment       =   1  'Right Justify
            Enabled         =   0   'False
            Height          =   450
            Left            =   6690
            Locked          =   -1  'True
            RightToLeft     =   -1  'True
            TabIndex        =   124
            Top             =   6930
            Width           =   3765
         End
         Begin VB.TextBox Text1 
            Alignment       =   1  'Right Justify
            Height          =   315
            Left            =   660
            RightToLeft     =   -1  'True
            TabIndex        =   123
            Top             =   7140
            Visible         =   0   'False
            Width           =   2520
         End
         Begin VB.TextBox TxtCashCustomerName 
            Alignment       =   1  'Right Justify
            Height          =   345
            Left            =   2520
            Locked          =   -1  'True
            TabIndex        =   121
            Top             =   1350
            Width           =   3345
         End
         Begin VB.TextBox TXTOrDer_no 
            Alignment       =   1  'Right Justify
            Height          =   315
            Index           =   1
            Left            =   6420
            TabIndex        =   120
            Top             =   540
            Visible         =   0   'False
            Width           =   1545
         End
         Begin VB.ComboBox DcbType 
            Height          =   315
            Left            =   12000
            RightToLeft     =   -1  'True
            Style           =   2  'Dropdown List
            TabIndex        =   108
            Top             =   8760
            Visible         =   0   'False
            Width           =   1530
         End
         Begin VB.ComboBox DCOPrType 
            Height          =   315
            Left            =   14505
            RightToLeft     =   -1  'True
            Style           =   2  'Dropdown List
            TabIndex        =   107
            Top             =   8790
            Visible         =   0   'False
            Width           =   2220
         End
         Begin VB.ComboBox DcbyearFactor 
            Enabled         =   0   'False
            Height          =   315
            Left            =   14505
            Locked          =   -1  'True
            RightToLeft     =   -1  'True
            TabIndex        =   106
            Top             =   1890
            Width           =   2655
         End
         Begin VB.TextBox TxtPlatNo 
            Alignment       =   1  'Right Justify
            Height          =   315
            Left            =   690
            Locked          =   -1  'True
            TabIndex        =   105
            Top             =   1890
            Width           =   2790
         End
         Begin VB.TextBox TxtManualNo2 
            Alignment       =   1  'Right Justify
            BackColor       =   &H0080FFFF&
            Height          =   285
            Index           =   2
            Left            =   9345
            Locked          =   -1  'True
            TabIndex        =   104
            Top             =   1920
            Width           =   2655
         End
         Begin VB.TextBox TxtManualNo2 
            Alignment       =   1  'Right Justify
            BackColor       =   &H0080FFFF&
            Height          =   285
            Index           =   1
            Left            =   5310
            Locked          =   -1  'True
            TabIndex        =   103
            Top             =   1920
            Width           =   1935
         End
         Begin VB.TextBox TXTOrDer_no 
            Alignment       =   1  'Right Justify
            Height          =   315
            Index           =   0
            Left            =   3630
            TabIndex        =   101
            Top             =   930
            Width           =   1680
         End
         Begin VB.ComboBox CBoBasedON 
            Height          =   315
            ItemData        =   "FrmVizitScreen.frx":1E637
            Left            =   8925
            List            =   "FrmVizitScreen.frx":1E639
            RightToLeft     =   -1  'True
            Style           =   2  'Dropdown List
            TabIndex        =   100
            Top             =   915
            Width           =   1800
         End
         Begin VB.TextBox TxtNoteID 
            Alignment       =   1  'Right Justify
            Height          =   315
            Left            =   4890
            RightToLeft     =   -1  'True
            TabIndex        =   82
            Top             =   7410
            Visible         =   0   'False
            Width           =   2640
         End
         Begin VB.TextBox TxtSerial1 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFFC0&
            Enabled         =   0   'False
            Height          =   360
            Index           =   1
            Left            =   0
            RightToLeft     =   -1  'True
            TabIndex        =   73
            Top             =   1290
            Visible         =   0   'False
            Width           =   2370
         End
         Begin VB.TextBox TxtRemarks 
            Alignment       =   2  'Center
            Height          =   555
            Left            =   135
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   72
            Top             =   2670
            Width           =   16455
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
            Width           =   20790
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
            Left            =   26070
            TabIndex        =   4
            Top             =   765
            Width           =   18420
            _cx             =   32491
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
            Left            =   11865
            TabIndex        =   74
            Top             =   990
            Width           =   1800
            _ExtentX        =   3175
            _ExtentY        =   529
            _Version        =   393216
            Format          =   139329537
            CurrentDate     =   38784
         End
         Begin MSDataListLib.DataCombo dcBranch 
            Bindings        =   "FrmVizitScreen.frx":2509B
            Height          =   315
            Index           =   1
            Left            =   0
            TabIndex        =   75
            Top             =   840
            Width           =   2655
            _ExtentX        =   4683
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
            Left            =   7815
            TabIndex        =   76
            Top             =   1380
            Width           =   4185
            _ExtentX        =   7382
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
            Left            =   13665
            TabIndex        =   83
            Top             =   7740
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
            Left            =   14235
            TabIndex        =   84
            Top             =   8430
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
            ButtonImage     =   "FrmVizitScreen.frx":250B0
            ColorButton     =   14871017
            DrawFocusRectangle=   0   'False
         End
         Begin ImpulseButton.ISButton btn_Save 
            Height          =   315
            Index           =   1
            Left            =   12000
            TabIndex        =   85
            Top             =   8400
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
            ButtonImage     =   "FrmVizitScreen.frx":2544A
            ColorButton     =   14871017
            DrawFocusRectangle=   0   'False
         End
         Begin ImpulseButton.ISButton btn_Modify 
            Height          =   225
            Index           =   1
            Left            =   12975
            TabIndex        =   86
            Top             =   8430
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
            ButtonImage     =   "FrmVizitScreen.frx":257E4
            ColorButton     =   14871017
            DrawFocusRectangle=   0   'False
         End
         Begin ImpulseButton.ISButton Btn_Undo 
            Height          =   225
            Index           =   1
            Left            =   11025
            TabIndex        =   87
            Top             =   8430
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
            ButtonImage     =   "FrmVizitScreen.frx":25B7E
            ColorButton     =   14871017
            DrawFocusRectangle=   0   'False
         End
         Begin ImpulseButton.ISButton btn_Delete 
            Height          =   315
            Index           =   1
            Left            =   10170
            TabIndex        =   88
            Top             =   8400
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
            ButtonImage     =   "FrmVizitScreen.frx":25F18
            ColorButton     =   14871017
            DrawFocusRectangle=   0   'False
         End
         Begin ImpulseButton.ISButton Btn_Update 
            Height          =   345
            Index           =   1
            Left            =   11310
            TabIndex        =   89
            TabStop         =   0   'False
            ToolTipText     =   " ÕœÌÀ ﬁ«⁄œ… «·»Ì«‰« "
            Top             =   8070
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
            ButtonImage     =   "FrmVizitScreen.frx":264B2
            ColorButton     =   14871017
            DrawFocusRectangle=   0   'False
         End
         Begin ImpulseButton.ISButton btn_Cancel 
            Height          =   315
            Index           =   1
            Left            =   5865
            TabIndex        =   90
            Top             =   8370
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
            ButtonImage     =   "FrmVizitScreen.frx":2684C
            ColorButton     =   14871017
            DrawFocusRectangle=   0   'False
            DisabledImageStyle=   1
         End
         Begin ImpulseButton.ISButton Btn_Print 
            Height          =   360
            Index           =   1
            Left            =   8520
            TabIndex        =   91
            TabStop         =   0   'False
            ToolTipText     =   "ÿ»«⁄… «·»Ì«‰«  "
            Top             =   8340
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
            ButtonImage     =   "FrmVizitScreen.frx":26BE6
            ColorButton     =   14871017
            DrawFocusRectangle=   0   'False
            DisabledImageStyle=   1
         End
         Begin ImpulseButton.ISButton btn_Query 
            Height          =   390
            Index           =   1
            Left            =   6975
            TabIndex        =   92
            TabStop         =   0   'False
            ToolTipText     =   "(Ctrl+F)  ··»ÕÀ ≈÷€ÿ Â–« «·„› «Õ √Ê ≈÷€ÿ "
            Top             =   8310
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
            ButtonImage     =   "FrmVizitScreen.frx":2D448
            ColorButton     =   14871017
            DrawFocusRectangle=   0   'False
         End
         Begin ImpulseButton.ISButton Cmd_DeleteRow 
            Height          =   300
            Index           =   1
            Left            =   2520
            TabIndex        =   93
            Top             =   7590
            Width           =   1800
            _ExtentX        =   3175
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
            Left            =   270
            TabIndex        =   94
            Top             =   7575
            Width           =   2100
            _ExtentX        =   3704
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
            Left            =   12555
            TabIndex        =   109
            Top             =   2220
            Visible         =   0   'False
            Width           =   4605
            _ExtentX        =   8123
            _ExtentY        =   556
            _Version        =   393216
            ListField       =   "6"
            Text            =   ""
            RightToLeft     =   -1  'True
         End
         Begin MSDataListLib.DataCombo DcbCarType 
            Bindings        =   "FrmVizitScreen.frx":2E316
            Height          =   315
            Left            =   14505
            TabIndex        =   110
            Top             =   1500
            Width           =   2655
            _ExtentX        =   4683
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
            Left            =   690
            TabIndex        =   131
            Top             =   2250
            Width           =   2790
            _ExtentX        =   4921
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
            Width           =   18555
            _cx             =   32729
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
               Width           =   15960
               _cx             =   28152
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
                  Width           =   3390
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
                  Left            =   3765
                  RightToLeft     =   -1  'True
                  TabIndex        =   154
                  Top             =   1470
                  Width           =   5790
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
                  Left            =   2265
                  Locked          =   -1  'True
                  RightToLeft     =   -1  'True
                  TabIndex        =   150
                  Top             =   3390
                  Width           =   3270
               End
               Begin VB.Frame Frame5 
                  Caption         =   "»Ì«‰«   ›Ê« Ì— «·„»Ì⁄« "
                  Height          =   1485
                  Left            =   3765
                  RightToLeft     =   -1  'True
                  TabIndex        =   139
                  Top             =   0
                  Width           =   5790
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
                  Left            =   23505
                  TabIndex        =   136
                  Top             =   645
                  Width           =   15825
                  _cx             =   27914
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
                  Left            =   9555
                  TabIndex        =   151
                  Top             =   90
                  Width           =   6150
                  _cx             =   10848
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
                  Left            =   5535
                  TabIndex        =   152
                  Top             =   3420
                  Width           =   1875
               End
            End
            Begin C1SizerLibCtl.C1Elastic ELe 
               Height          =   3390
               Index           =   3
               Left            =   16695
               TabIndex        =   137
               TabStop         =   0   'False
               Top             =   45
               Width           =   15960
               _cx             =   28152
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
                  Left            =   23370
                  TabIndex        =   138
                  Top             =   720
                  Width           =   15840
                  _cx             =   27940
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
                  Width           =   15075
                  _cx             =   26591
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
            Left            =   4890
            TabIndex        =   133
            Top             =   960
            Width           =   1800
         End
         Begin VB.Label lblModel 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "«·ÿ—«“ "
            Height          =   255
            Left            =   3765
            TabIndex        =   132
            Top             =   2250
            Width           =   1125
         End
         Begin VB.Label Label5 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "—ﬁ„ «·”‰œ"
            Height          =   255
            Index           =   1
            Left            =   17445
            TabIndex        =   130
            Top             =   960
            Width           =   975
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            BackStyle       =   0  'Transparent
            Caption         =   "—ﬁ„ «·ﬁÌœ"
            Height          =   405
            Index           =   14
            Left            =   10455
            RightToLeft     =   -1  'True
            TabIndex        =   128
            Top             =   7050
            Width           =   1260
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "«”„ «·⁄„Ì· «·‰ﬁœÏ"
            Height          =   315
            Index           =   33
            Left            =   6135
            TabIndex        =   122
            Top             =   1410
            Width           =   1530
         End
         Begin VB.Label lbl 
            Alignment       =   2  'Center
            BackColor       =   &H00E2E9E9&
            Caption         =   "%"
            Height          =   225
            Index           =   3
            Left            =   4185
            TabIndex        =   119
            Top             =   5190
            Width           =   705
         End
         Begin VB.Label LblYear 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "„ÊœÌ· «·„⁄œÂ/«·”Ì«—…"
            Height          =   255
            Left            =   17010
            TabIndex        =   118
            Top             =   1830
            Width           =   1410
         End
         Begin VB.Label LblPla 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "—ﬁ„ «··ÊÕ…"
            Height          =   255
            Left            =   3900
            TabIndex        =   117
            Top             =   1920
            Width           =   1260
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "«·‰Ê⁄"
            Height          =   285
            Index           =   123
            Left            =   13530
            RightToLeft     =   -1  'True
            TabIndex        =   116
            Top             =   9000
            Visible         =   0   'False
            Width           =   840
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "‰Ê⁄ «·⁄„·Ì…"
            Height          =   285
            Index           =   124
            Left            =   17160
            RightToLeft     =   -1  'True
            TabIndex        =   115
            Top             =   8745
            Visible         =   0   'False
            Width           =   840
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "«·„⁄œÂ/«·”Ì«—…"
            Height          =   240
            Index           =   125
            Left            =   17160
            RightToLeft     =   -1  'True
            TabIndex        =   114
            Top             =   2250
            Visible         =   0   'False
            Width           =   1260
         End
         Begin VB.Label lbltycar 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "‰Ê⁄ «·„⁄œÂ/«·”Ì«—…"
            Height          =   255
            Left            =   17445
            TabIndex        =   113
            Top             =   1470
            Width           =   1110
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            BackStyle       =   0  'Transparent
            Caption         =   "—ﬁ„ «·‘«”ÌÂ"
            Height          =   195
            Index           =   119
            Left            =   11715
            TabIndex        =   112
            Top             =   1920
            Width           =   1665
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            BackStyle       =   0  'Transparent
            Caption         =   "⁄œ«œ «·ﬂÌ·Ê „ —"
            Height          =   195
            Index           =   118
            Left            =   7245
            TabIndex        =   111
            Top             =   1920
            Width           =   1545
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "»‰«¡ ⁄·Ï"
            Height          =   255
            Index           =   56
            Left            =   10725
            TabIndex        =   102
            Top             =   945
            Width           =   990
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Õ—— »Ê«”ÿ…  "
            Height          =   315
            Index           =   8
            Left            =   17160
            TabIndex        =   99
            Top             =   7680
            Width           =   1260
         End
         Begin VB.Label LabCount_Rec 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Height          =   225
            Index           =   1
            Left            =   3075
            RightToLeft     =   -1  'True
            TabIndex        =   98
            Top             =   8070
            Width           =   555
         End
         Begin VB.Label LabCurr_Rec 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Height          =   225
            Index           =   1
            Left            =   5160
            RightToLeft     =   -1  'True
            TabIndex        =   97
            Top             =   8070
            Width           =   705
         End
         Begin VB.Label Label2 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "⁄œœ «·”Ã·« :"
            Height          =   225
            Index           =   2
            Left            =   3765
            RightToLeft     =   -1  'True
            TabIndex        =   96
            Top             =   8055
            Width           =   1395
         End
         Begin VB.Label Label2 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "«·”Ã· «·Õ«·Ì:"
            Height          =   225
            Index           =   3
            Left            =   6000
            RightToLeft     =   -1  'True
            TabIndex        =   95
            Top             =   8055
            Width           =   1245
         End
         Begin VB.Label lbl 
            Alignment       =   2  'Center
            BackColor       =   &H00E2E9E9&
            Caption         =   " «—ÌŒ «·”‰œ"
            Height          =   270
            Index           =   2
            Left            =   13935
            TabIndex        =   81
            Top             =   990
            Width           =   1410
         End
         Begin VB.Label lbl 
            Alignment       =   2  'Center
            BackColor       =   &H00E2E9E9&
            Caption         =   "—ﬁ„ «·⁄ﬁœ"
            Height          =   255
            Index           =   4
            Left            =   18690
            RightToLeft     =   -1  'True
            TabIndex        =   80
            Top             =   1275
            Width           =   1110
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "«·›—⁄"
            Height          =   255
            Index           =   7
            Left            =   2370
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
            Left            =   18555
            RightToLeft     =   -1  'True
            TabIndex        =   78
            Top             =   3060
            Width           =   1245
         End
         Begin VB.Label lbl 
            Alignment       =   2  'Center
            BackColor       =   &H00E2E9E9&
            Caption         =   "«·⁄„Ì·"
            Height          =   465
            Index           =   15
            Left            =   11715
            TabIndex        =   77
            Top             =   1410
            Width           =   1395
         End
         Begin VB.Label Label4 
            Alignment       =   1  'Right Justify
            Caption         =   "„·«ÕŸ« "
            Height          =   255
            Index           =   0
            Left            =   16725
            RightToLeft     =   -1  'True
            TabIndex        =   71
            Top             =   2850
            Width           =   990
         End
         Begin VB.Label Label3 
            Alignment       =   1  'Right Justify
            Caption         =   "Label3"
            Height          =   135
            Index           =   1
            Left            =   4185
            RightToLeft     =   -1  'True
            TabIndex        =   59
            Top             =   0
            Width           =   405
         End
      End
      Begin C1SizerLibCtl.C1Elastic ELe 
         Height          =   8790
         Index           =   4
         Left            =   -20835
         TabIndex        =   183
         TabStop         =   0   'False
         Top             =   45
         Width           =   18690
         _cx             =   32967
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
            Left            =   135
            RightToLeft     =   -1  'True
            TabIndex        =   192
            Top             =   4560
            Width           =   8520
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
            Width           =   18690
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
            Left            =   7245
            TabIndex        =   200
            Top             =   7680
            Width           =   1125
            _ExtentX        =   1984
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
            Left            =   5310
            TabIndex        =   201
            Top             =   7650
            Width           =   825
            _ExtentX        =   1455
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
            Left            =   6270
            TabIndex        =   202
            Top             =   7680
            Width           =   975
            _ExtentX        =   1720
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
            Left            =   4320
            TabIndex        =   203
            Top             =   7650
            Width           =   990
            _ExtentX        =   1746
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
            Left            =   3480
            TabIndex        =   204
            Top             =   7680
            Width           =   840
            _ExtentX        =   1482
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
            Left            =   6555
            TabIndex        =   205
            TabStop         =   0   'False
            ToolTipText     =   " ÕœÌÀ ﬁ«⁄œ… «·»Ì«‰« "
            Top             =   6450
            Visible         =   0   'False
            Width           =   1815
            _ExtentX        =   3201
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
            Left            =   135
            TabIndex        =   206
            Top             =   7620
            Width           =   975
            _ExtentX        =   1720
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
            Left            =   2370
            TabIndex        =   207
            TabStop         =   0   'False
            ToolTipText     =   "ÿ»«⁄… «·»Ì«‰«  "
            Top             =   7620
            Width           =   975
            _ExtentX        =   1720
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
            Left            =   1530
            TabIndex        =   208
            TabStop         =   0   'False
            ToolTipText     =   "(Ctrl+F)  ··»ÕÀ ≈÷€ÿ Â–« «·„› «Õ √Ê ≈÷€ÿ "
            Top             =   7545
            Width           =   705
            _ExtentX        =   1244
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
            Width           =   8790
            _cx             =   15505
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
            Left            =   6135
            RightToLeft     =   -1  'True
            TabIndex        =   217
            Top             =   6705
            Width           =   2100
         End
         Begin VB.Label Label2 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "⁄œœ «·”Ã·« :"
            Height          =   225
            Index           =   6
            Left            =   2235
            RightToLeft     =   -1  'True
            TabIndex        =   216
            Top             =   6705
            Width           =   2085
         End
         Begin VB.Label LabCurr_Rec 
            BackColor       =   &H00E2E9E9&
            Height          =   240
            Index           =   0
            Left            =   4320
            RightToLeft     =   -1  'True
            TabIndex        =   215
            Top             =   6720
            Width           =   1545
         End
         Begin VB.Label LabCount_Rec 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Height          =   240
            Index           =   0
            Left            =   690
            RightToLeft     =   -1  'True
            TabIndex        =   214
            Top             =   6720
            Width           =   1125
         End
         Begin VB.Label Label2 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "«·”Ã· «·Õ«·Ì:"
            Height          =   225
            Index           =   4
            Left            =   6690
            RightToLeft     =   -1  'True
            TabIndex        =   213
            Top             =   7170
            Width           =   1545
         End
         Begin VB.Label Label2 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "⁄œœ «·”Ã·« :"
            Height          =   225
            Index           =   5
            Left            =   2790
            RightToLeft     =   -1  'True
            TabIndex        =   212
            Top             =   7170
            Width           =   2100
         End
         Begin VB.Label LabCurr_Rec 
            BackColor       =   &H00E2E9E9&
            Height          =   240
            Index           =   2
            Left            =   4890
            RightToLeft     =   -1  'True
            TabIndex        =   211
            Top             =   7185
            Width           =   1530
         End
         Begin VB.Label LabCount_Rec 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Height          =   240
            Index           =   2
            Left            =   1245
            RightToLeft     =   -1  'True
            TabIndex        =   210
            Top             =   7185
            Width           =   1275
         End
      End
      Begin C1SizerLibCtl.C1Elastic ELe 
         Height          =   8790
         Index           =   5
         Left            =   -20535
         TabIndex        =   218
         TabStop         =   0   'False
         Top             =   45
         Width           =   18690
         _cx             =   32967
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
         Begin VB.CommandButton Command3 
            Caption         =   " ÕœÌÀ «·«”⁄«— ›ﬁÿ"
            Height          =   450
            Left            =   5715
            RightToLeft     =   -1  'True
            TabIndex        =   369
            Top             =   5580
            Width           =   1395
         End
         Begin VB.CommandButton Command1 
            Caption         =   "«‰‘«¡ «·«’‰«›"
            Height          =   450
            Left            =   1380
            RightToLeft     =   -1  'True
            TabIndex        =   252
            Top             =   5400
            Visible         =   0   'False
            Width           =   2805
         End
         Begin VB.ComboBox cmbFlag 
            Height          =   315
            Index           =   2
            ItemData        =   "FrmVizitScreen.frx":39D15
            Left            =   1380
            List            =   "FrmVizitScreen.frx":39D17
            RightToLeft     =   -1  'True
            Style           =   2  'Dropdown List
            TabIndex        =   251
            Top             =   5040
            Visible         =   0   'False
            Width           =   1830
         End
         Begin VB.ComboBox cmbFlag 
            Height          =   315
            Index           =   1
            ItemData        =   "FrmVizitScreen.frx":39D19
            Left            =   3345
            List            =   "FrmVizitScreen.frx":39D1B
            RightToLeft     =   -1  'True
            Style           =   2  'Dropdown List
            TabIndex        =   250
            Top             =   5070
            Visible         =   0   'False
            Width           =   1815
         End
         Begin VB.ComboBox cmbFlag 
            Height          =   315
            Index           =   0
            ItemData        =   "FrmVizitScreen.frx":39D1D
            Left            =   3075
            List            =   "FrmVizitScreen.frx":39D1F
            RightToLeft     =   -1  'True
            TabIndex        =   249
            Text            =   "cmbFlag"
            Top             =   2850
            Width           =   1815
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
            Width           =   18690
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
            Height          =   4155
            Index           =   3
            Left            =   135
            RightToLeft     =   -1  'True
            TabIndex        =   219
            Top             =   720
            Width           =   6975
            Begin VB.TextBox txtPrice 
               Alignment       =   1  'Right Justify
               Height          =   315
               Left            =   90
               RightToLeft     =   -1  'True
               TabIndex        =   264
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
               TabIndex        =   254
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
               TabIndex        =   256
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
               TabIndex        =   258
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
               TabIndex        =   259
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
               TabIndex        =   262
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
               TabIndex        =   263
               Top             =   1710
               Width           =   4065
               _ExtentX        =   7170
               _ExtentY        =   556
               _Version        =   393216
               ListField       =   "6"
               Text            =   ""
               RightToLeft     =   -1  'True
            End
            Begin MSDataListLib.DataCombo cmbEyeDet 
               Height          =   315
               Index           =   13
               Left            =   1080
               TabIndex        =   357
               Top             =   3480
               Width           =   3075
               _ExtentX        =   5424
               _ExtentY        =   556
               _Version        =   393216
               Style           =   2
               Text            =   ""
               RightToLeft     =   -1  'True
            End
            Begin MSDataListLib.DataCombo cmbEyeDet 
               Height          =   315
               Index           =   14
               Left            =   1080
               TabIndex        =   359
               Top             =   3840
               Width           =   3075
               _ExtentX        =   5424
               _ExtentY        =   556
               _Version        =   393216
               Style           =   2
               Text            =   ""
               RightToLeft     =   -1  'True
            End
            Begin VB.Label Label3 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "«·”⁄—"
               Height          =   285
               Index           =   8
               Left            =   1440
               RightToLeft     =   -1  'True
               TabIndex        =   364
               Top             =   2160
               Width           =   990
            End
            Begin VB.Label lbl«”„«·ÊÕœ… 
               Alignment       =   2  'Center
               AutoSize        =   -1  'True
               BackColor       =   &H00E2E9E9&
               Caption         =   "DIAM"
               Height          =   195
               Index           =   30
               Left            =   4410
               RightToLeft     =   -1  'True
               TabIndex        =   360
               Top             =   3945
               Width           =   435
            End
            Begin VB.Label lbl«”„«·ÊÕœ… 
               Alignment       =   2  'Center
               AutoSize        =   -1  'True
               BackColor       =   &H00E2E9E9&
               Caption         =   "Index"
               Height          =   195
               Index           =   31
               Left            =   4380
               RightToLeft     =   -1  'True
               TabIndex        =   358
               Top             =   3525
               Width           =   405
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
               TabIndex        =   261
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
               TabIndex        =   260
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
               TabIndex        =   257
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
               TabIndex        =   255
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
               TabIndex        =   253
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
               TabIndex        =   248
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
               TabIndex        =   247
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
            Left            =   7245
            TabIndex        =   234
            Top             =   8280
            Width           =   1125
            _ExtentX        =   1984
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
            Left            =   5310
            TabIndex        =   235
            Top             =   8250
            Width           =   825
            _ExtentX        =   1455
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
            Left            =   6270
            TabIndex        =   236
            Top             =   8280
            Width           =   975
            _ExtentX        =   1720
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
            Left            =   4320
            TabIndex        =   237
            Top             =   8250
            Width           =   990
            _ExtentX        =   1746
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
            Left            =   3480
            TabIndex        =   238
            Top             =   8280
            Width           =   840
            _ExtentX        =   1482
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
            Left            =   2655
            TabIndex        =   239
            TabStop         =   0   'False
            ToolTipText     =   " ÕœÌÀ ﬁ«⁄œ… «·»Ì«‰« "
            Top             =   7170
            Visible         =   0   'False
            Width           =   1935
            _ExtentX        =   3413
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
            Left            =   135
            TabIndex        =   240
            Top             =   8220
            Width           =   975
            _ExtentX        =   1720
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
            Left            =   2370
            TabIndex        =   241
            TabStop         =   0   'False
            ToolTipText     =   "ÿ»«⁄… «·»Ì«‰«  "
            Top             =   8220
            Width           =   975
            _ExtentX        =   1720
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
            Left            =   1530
            TabIndex        =   242
            TabStop         =   0   'False
            ToolTipText     =   "(Ctrl+F)  ··»ÕÀ ≈÷€ÿ Â–« «·„› «Õ √Ê ≈÷€ÿ "
            Top             =   8145
            Width           =   705
            _ExtentX        =   1244
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
         Begin C1SizerLibCtl.C1Tab C1Tab1 
            Height          =   7290
            Left            =   7110
            TabIndex        =   282
            Top             =   720
            Width           =   11445
            _cx             =   20188
            _cy             =   12859
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
            Caption         =   " ›«’Ì· 1|«·⁄œ”« "
            Align           =   0
            CurrTab         =   1
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
               Height          =   6915
               Index           =   16
               Left            =   -12000
               TabIndex        =   283
               TabStop         =   0   'False
               Top             =   45
               Width           =   11355
               _cx             =   20029
               _cy             =   12197
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
               Begin VB.TextBox TxtBrandType 
                  Alignment       =   2  'Center
                  Height          =   345
                  Left            =   8775
                  TabIndex        =   292
                  Top             =   1950
                  Visible         =   0   'False
                  Width           =   1545
               End
               Begin VB.TextBox TxtModel 
                  Alignment       =   2  'Center
                  Height          =   315
                  Left            =   2460
                  TabIndex        =   291
                  Top             =   1170
                  Width           =   645
               End
               Begin VB.TextBox TxtColorCode 
                  Alignment       =   2  'Center
                  Height          =   300
                  Left            =   1155
                  TabIndex        =   290
                  Top             =   1155
                  Width           =   780
               End
               Begin VB.TextBox TxtSize 
                  Alignment       =   2  'Center
                  Height          =   300
                  Left            =   0
                  TabIndex        =   289
                  Top             =   1155
                  Width           =   780
               End
               Begin VB.CommandButton cmdLoadFile 
                  Caption         =   " Õ„Ì· «·„·›..."
                  Height          =   270
                  Left            =   10590
                  TabIndex        =   288
                  Top             =   4410
                  Visible         =   0   'False
                  Width           =   765
               End
               Begin VB.TextBox txtFile 
                  Height          =   345
                  Left            =   9555
                  Locked          =   -1  'True
                  TabIndex        =   287
                  Top             =   2565
                  Visible         =   0   'False
                  Width           =   900
               End
               Begin VB.CommandButton cmdSelectFile 
                  Caption         =   " ÕœÌœ «·„·›..."
                  Height          =   240
                  Left            =   8265
                  RightToLeft     =   -1  'True
                  TabIndex        =   286
                  Top             =   1320
                  Visible         =   0   'False
                  Width           =   1155
               End
               Begin VB.CommandButton Command8 
                  Caption         =   " ÕœÌÀ «·«’‰«›"
                  Height          =   180
                  Left            =   8130
                  TabIndex        =   285
                  Top             =   1845
                  Visible         =   0   'False
                  Width           =   1425
               End
               Begin VB.ComboBox cboMasterType 
                  Appearance      =   0  'Flat
                  Height          =   315
                  Left            =   0
                  RightToLeft     =   -1  'True
                  Style           =   2  'Dropdown List
                  TabIndex        =   284
                  Top             =   240
                  Width           =   3105
               End
               Begin VSFlex8UCtl.VSFlexGrid FgItems 
                  Height          =   6810
                  Index           =   7
                  Left            =   13425
                  TabIndex        =   293
                  Top             =   540
                  Width           =   9675
                  _cx             =   17066
                  _cy             =   12012
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
               Begin MSDataListLib.DataCombo cmbEyeDet 
                  Height          =   315
                  Index           =   0
                  Left            =   0
                  TabIndex        =   294
                  Top             =   690
                  Width           =   3105
                  _ExtentX        =   5477
                  _ExtentY        =   556
                  _Version        =   393216
                  Style           =   2
                  Text            =   ""
                  RightToLeft     =   -1  'True
               End
               Begin MSDataListLib.DataCombo cmbEyeDet 
                  Height          =   315
                  Index           =   1
                  Left            =   0
                  TabIndex        =   295
                  Top             =   2625
                  Width           =   3105
                  _ExtentX        =   5477
                  _ExtentY        =   556
                  _Version        =   393216
                  Style           =   2
                  Text            =   ""
                  RightToLeft     =   -1  'True
               End
               Begin MSDataListLib.DataCombo cmbEyeDet 
                  Height          =   315
                  Index           =   2
                  Left            =   0
                  TabIndex        =   296
                  Top             =   3030
                  Width           =   3105
                  _ExtentX        =   5477
                  _ExtentY        =   556
                  _Version        =   393216
                  Style           =   2
                  Text            =   ""
                  RightToLeft     =   -1  'True
               End
               Begin MSDataListLib.DataCombo cmbEyeDet 
                  Height          =   315
                  Index           =   4
                  Left            =   0
                  TabIndex        =   297
                  Top             =   1500
                  Width           =   3105
                  _ExtentX        =   5477
                  _ExtentY        =   556
                  _Version        =   393216
                  Style           =   2
                  Text            =   ""
                  RightToLeft     =   -1  'True
               End
               Begin MSDataListLib.DataCombo cmbEyeDet 
                  Height          =   315
                  Index           =   5
                  Left            =   4770
                  TabIndex        =   298
                  Top             =   135
                  Visible         =   0   'False
                  Width           =   2970
                  _ExtentX        =   5239
                  _ExtentY        =   556
                  _Version        =   393216
                  Style           =   2
                  Text            =   ""
                  RightToLeft     =   -1  'True
               End
               Begin MSDataListLib.DataCombo cmbEyeDet 
                  Height          =   315
                  Index           =   6
                  Left            =   0
                  TabIndex        =   299
                  Top             =   3450
                  Width           =   3105
                  _ExtentX        =   5477
                  _ExtentY        =   556
                  _Version        =   393216
                  Style           =   2
                  Text            =   ""
                  RightToLeft     =   -1  'True
               End
               Begin MSDataListLib.DataCombo cmbEyeDet 
                  Height          =   315
                  Index           =   8
                  Left            =   0
                  TabIndex        =   300
                  Top             =   2250
                  Width           =   3105
                  _ExtentX        =   5477
                  _ExtentY        =   556
                  _Version        =   393216
                  Style           =   2
                  Text            =   ""
                  RightToLeft     =   -1  'True
               End
               Begin MSDataListLib.DataCombo cmbEyeDet 
                  Height          =   315
                  Index           =   7
                  Left            =   8910
                  TabIndex        =   301
                  Top             =   5730
                  Visible         =   0   'False
                  Width           =   3090
                  _ExtentX        =   5450
                  _ExtentY        =   556
                  _Version        =   393216
                  Style           =   2
                  Text            =   ""
                  RightToLeft     =   -1  'True
               End
               Begin MSDataListLib.DataCombo cmbEyeDet 
                  Height          =   315
                  Index           =   9
                  Left            =   0
                  TabIndex        =   302
                  Top             =   4830
                  Width           =   3105
                  _ExtentX        =   5477
                  _ExtentY        =   556
                  _Version        =   393216
                  Style           =   2
                  Text            =   ""
                  RightToLeft     =   -1  'True
               End
               Begin VSFlex8Ctl.VSFlexGrid tmpGrd 
                  Height          =   345
                  Left            =   8910
                  TabIndex        =   303
                  Top             =   3420
                  Visible         =   0   'False
                  Width           =   1035
                  _cx             =   1826
                  _cy             =   609
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
                  BackColor       =   8421631
                  ForeColor       =   -2147483640
                  BackColorFixed  =   -2147483633
                  ForeColorFixed  =   -2147483630
                  BackColorSel    =   -2147483635
                  ForeColorSel    =   -2147483634
                  BackColorBkg    =   -2147483636
                  BackColorAlternate=   8421631
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
                  Cols            =   40
                  FixedRows       =   1
                  FixedCols       =   1
                  RowHeightMin    =   0
                  RowHeightMax    =   0
                  ColWidthMin     =   0
                  ColWidthMax     =   0
                  ExtendLastCol   =   0   'False
                  FormatString    =   ""
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
               Begin MSDataListLib.DataCombo cmbEyeDet 
                  Height          =   315
                  Index           =   24
                  Left            =   5295
                  TabIndex        =   304
                  Top             =   5715
                  Visible         =   0   'False
                  Width           =   3090
                  _ExtentX        =   5450
                  _ExtentY        =   556
                  _Version        =   393216
                  Style           =   2
                  Text            =   ""
                  RightToLeft     =   -1  'True
               End
               Begin MSDataListLib.DataCombo cmbEyeDet 
                  Height          =   315
                  Index           =   25
                  Left            =   5295
                  TabIndex        =   305
                  Top             =   6135
                  Visible         =   0   'False
                  Width           =   3090
                  _ExtentX        =   5450
                  _ExtentY        =   556
                  _Version        =   393216
                  Style           =   2
                  Text            =   ""
                  RightToLeft     =   -1  'True
               End
               Begin MSDataListLib.DataCombo cmbEyeDet 
                  Height          =   315
                  Index           =   19
                  Left            =   9945
                  TabIndex        =   306
                  Top             =   2085
                  Visible         =   0   'False
                  Width           =   3090
                  _ExtentX        =   5450
                  _ExtentY        =   556
                  _Version        =   393216
                  Style           =   2
                  Text            =   ""
                  RightToLeft     =   -1  'True
               End
               Begin MSDataListLib.DataCombo cmbEyeDet 
                  Height          =   315
                  Index           =   10
                  Left            =   0
                  TabIndex        =   307
                  Top             =   1920
                  Width           =   3105
                  _ExtentX        =   5477
                  _ExtentY        =   556
                  _Version        =   393216
                  Style           =   2
                  Text            =   ""
                  RightToLeft     =   -1  'True
               End
               Begin MSDataListLib.DataCombo cmbEyeDet 
                  Height          =   315
                  Index           =   17
                  Left            =   0
                  TabIndex        =   308
                  Top             =   6255
                  Width           =   3105
                  _ExtentX        =   5477
                  _ExtentY        =   556
                  _Version        =   393216
                  Style           =   2
                  Text            =   ""
                  RightToLeft     =   -1  'True
               End
               Begin MSDataListLib.DataCombo cmbEyeDet 
                  Height          =   315
                  Index           =   20
                  Left            =   0
                  TabIndex        =   309
                  Top             =   5295
                  Width           =   3105
                  _ExtentX        =   5477
                  _ExtentY        =   556
                  _Version        =   393216
                  Style           =   2
                  Text            =   ""
                  RightToLeft     =   -1  'True
               End
               Begin MSDataListLib.DataCombo cmbEyeDet 
                  Height          =   315
                  Index           =   21
                  Left            =   0
                  TabIndex        =   310
                  Top             =   5820
                  Width           =   3105
                  _ExtentX        =   5477
                  _ExtentY        =   556
                  _Version        =   393216
                  Style           =   2
                  Text            =   ""
                  RightToLeft     =   -1  'True
               End
               Begin MSDataListLib.DataCombo cmbEyeDet 
                  Height          =   315
                  Index           =   3
                  Left            =   5295
                  TabIndex        =   311
                  Top             =   6450
                  Visible         =   0   'False
                  Width           =   3090
                  _ExtentX        =   5450
                  _ExtentY        =   556
                  _Version        =   393216
                  Style           =   2
                  Text            =   ""
                  RightToLeft     =   -1  'True
               End
               Begin MSDataListLib.DataCombo cmbEyeDet 
                  Height          =   315
                  Index           =   12
                  Left            =   0
                  TabIndex        =   342
                  Top             =   4440
                  Width           =   3105
                  _ExtentX        =   5477
                  _ExtentY        =   556
                  _Version        =   393216
                  Style           =   2
                  Text            =   ""
                  RightToLeft     =   -1  'True
               End
               Begin MSDataListLib.DataCombo cmbEyeDet 
                  Height          =   315
                  Index           =   15
                  Left            =   0
                  TabIndex        =   343
                  Top             =   3975
                  Width           =   3105
                  _ExtentX        =   5477
                  _ExtentY        =   556
                  _Version        =   393216
                  Style           =   2
                  Text            =   ""
                  RightToLeft     =   -1  'True
               End
               Begin MSDataListLib.DataCombo cmbEyeDet 
                  Height          =   315
                  Index           =   16
                  Left            =   5160
                  TabIndex        =   344
                  Top             =   4920
                  Visible         =   0   'False
                  Width           =   3105
                  _ExtentX        =   5477
                  _ExtentY        =   556
                  _Version        =   393216
                  Style           =   2
                  Text            =   ""
                  RightToLeft     =   -1  'True
               End
               Begin MSDataListLib.DataCombo cmbEyeDet 
                  Height          =   315
                  Index           =   11
                  Left            =   6330
                  TabIndex        =   348
                  Top             =   2655
                  Visible         =   0   'False
                  Width           =   1290
                  _ExtentX        =   2275
                  _ExtentY        =   556
                  _Version        =   393216
                  Style           =   2
                  Text            =   ""
                  RightToLeft     =   -1  'True
               End
               Begin MSDataListLib.DataCombo cmbEyeDet 
                  Height          =   315
                  Index           =   22
                  Left            =   6450
                  TabIndex        =   349
                  Top             =   1335
                  Visible         =   0   'False
                  Width           =   1170
                  _ExtentX        =   2064
                  _ExtentY        =   556
                  _Version        =   393216
                  Style           =   2
                  Text            =   ""
                  RightToLeft     =   -1  'True
               End
               Begin MSDataListLib.DataCombo cmbEyeDet 
                  Height          =   315
                  Index           =   23
                  Left            =   6450
                  TabIndex        =   350
                  Top             =   2160
                  Visible         =   0   'False
                  Width           =   1170
                  _ExtentX        =   2064
                  _ExtentY        =   556
                  _Version        =   393216
                  Style           =   2
                  Text            =   ""
                  RightToLeft     =   -1  'True
               End
               Begin VB.Label lbl«”„«·ÊÕœ… 
                  Alignment       =   2  'Center
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "Sphere"
                  Height          =   540
                  Index           =   36
                  Left            =   7620
                  TabIndex        =   353
                  Top             =   1320
                  Visible         =   0   'False
                  Width           =   645
               End
               Begin VB.Label lbl«”„«·ÊÕœ… 
                  Alignment       =   2  'Center
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "Cylinder"
                  Height          =   555
                  Index           =   35
                  Left            =   7620
                  TabIndex        =   352
                  Top             =   2220
                  Visible         =   0   'False
                  Width           =   645
               End
               Begin VB.Label lbl«”„«·ÊÕœ… 
                  Alignment       =   2  'Center
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "Division"
                  Height          =   810
                  Index           =   33
                  Left            =   7620
                  TabIndex        =   351
                  Top             =   2775
                  Visible         =   0   'False
                  Width           =   645
               End
               Begin VB.Label lbl«”„«·ÊÕœ… 
                  Alignment       =   2  'Center
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "Coating"
                  Height          =   225
                  Index           =   32
                  Left            =   3870
                  TabIndex        =   347
                  Top             =   4470
                  Width           =   645
               End
               Begin VB.Label lbl«”„«·ÊÕœ… 
                  Alignment       =   2  'Center
                  AutoSize        =   -1  'True
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "Light Adaptation"
                  Height          =   195
                  Index           =   29
                  Left            =   3225
                  RightToLeft     =   -1  'True
                  TabIndex        =   346
                  Top             =   3975
                  Width           =   1170
               End
               Begin VB.Label lbl«”„«·ÊÕœ… 
                  Alignment       =   2  'Center
                  AutoSize        =   -1  'True
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "Breaking"
                  Height          =   195
                  Index           =   28
                  Left            =   9030
                  RightToLeft     =   -1  'True
                  TabIndex        =   345
                  Top             =   4800
                  Visible         =   0   'False
                  Width           =   645
               End
               Begin VB.Label lbl 
                  Alignment       =   2  'Center
                  BackStyle       =   0  'Transparent
                  Caption         =   "Age"
                  Height          =   435
                  Index           =   93
                  Left            =   8385
                  TabIndex        =   331
                  Top             =   6165
                  Visible         =   0   'False
                  Width           =   1425
               End
               Begin VB.Label lbl 
                  Alignment       =   2  'Center
                  BackStyle       =   0  'Transparent
                  Caption         =   "Gender"
                  Height          =   405
                  Index           =   92
                  Left            =   8385
                  TabIndex        =   330
                  Top             =   5835
                  Visible         =   0   'False
                  Width           =   1425
               End
               Begin VB.Label lbl«”„«·ÊÕœ… 
                  Alignment       =   2  'Center
                  AutoSize        =   -1  'True
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "Material "
                  Height          =   195
                  Index           =   15
                  Left            =   3870
                  RightToLeft     =   -1  'True
                  TabIndex        =   329
                  Top             =   3525
                  Width           =   525
               End
               Begin VB.Label lbl«”„«·ÊÕœ… 
                  Alignment       =   2  'Center
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "Shapes"
                  Height          =   315
                  Index           =   14
                  Left            =   7875
                  TabIndex        =   328
                  Top             =   165
                  Visible         =   0   'False
                  Width           =   1425
               End
               Begin VB.Label lbl«”„«·ÊÕœ… 
                  Alignment       =   2  'Center
                  AutoSize        =   -1  'True
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "Brand Type"
                  Height          =   195
                  Index           =   13
                  Left            =   3615
                  RightToLeft     =   -1  'True
                  TabIndex        =   327
                  Top             =   1575
                  Width           =   780
               End
               Begin VB.Label lbl«”„«·ÊÕœ… 
                  Alignment       =   2  'Center
                  AutoSize        =   -1  'True
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "Design"
                  Height          =   195
                  Index           =   11
                  Left            =   3870
                  RightToLeft     =   -1  'True
                  TabIndex        =   326
                  Top             =   3075
                  Width           =   525
               End
               Begin VB.Label lbl«”„«·ÊÕœ… 
                  Alignment       =   2  'Center
                  AutoSize        =   -1  'True
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "Type"
                  Height          =   195
                  Index           =   10
                  Left            =   4005
                  RightToLeft     =   -1  'True
                  TabIndex        =   325
                  Top             =   2685
                  Width           =   390
               End
               Begin VB.Label lbl«”„«·ÊÕœ… 
                  Alignment       =   2  'Center
                  AutoSize        =   -1  'True
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "Brand"
                  Height          =   195
                  Index           =   9
                  Left            =   4005
                  RightToLeft     =   -1  'True
                  TabIndex        =   324
                  Top             =   795
                  Width           =   390
               End
               Begin VB.Label lbl«”„«·ÊÕœ… 
                  Alignment       =   2  'Center
                  AutoSize        =   -1  'True
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "Category"
                  Height          =   195
                  Index           =   16
                  Left            =   3750
                  RightToLeft     =   -1  'True
                  TabIndex        =   323
                  Top             =   2295
                  Width           =   645
               End
               Begin VB.Label lbl«”„«·ÊÕœ… 
                  Alignment       =   2  'Center
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "Source"
                  Height          =   330
                  Index           =   17
                  Left            =   8010
                  TabIndex        =   322
                  Top             =   630
                  Visible         =   0   'False
                  Width           =   1665
               End
               Begin VB.Label lbl«”„«·ÊÕœ… 
                  Alignment       =   2  'Center
                  AutoSize        =   -1  'True
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "Colors"
                  Height          =   195
                  Index           =   18
                  Left            =   4005
                  RightToLeft     =   -1  'True
                  TabIndex        =   321
                  Top             =   4905
                  Width           =   390
               End
               Begin VB.Label lbl«”„«·ÊÕœ… 
                  Alignment       =   2  'Center
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "Brand Type"
                  Height          =   435
                  Index           =   19
                  Left            =   9300
                  TabIndex        =   320
                  Top             =   1455
                  Visible         =   0   'False
                  Width           =   900
               End
               Begin VB.Label lbl«”„«·ÊÕœ… 
                  Alignment       =   2  'Center
                  AutoSize        =   -1  'True
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "Model"
                  Height          =   195
                  Index           =   20
                  Left            =   4005
                  RightToLeft     =   -1  'True
                  TabIndex        =   319
                  Top             =   1170
                  Width           =   390
               End
               Begin VB.Label lbl«”„«·ÊÕœ… 
                  Alignment       =   2  'Center
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "Color Code"
                  Height          =   180
                  Index           =   21
                  Left            =   1935
                  TabIndex        =   318
                  Top             =   1185
                  Width           =   525
               End
               Begin VB.Label lbl«”„«·ÊÕœ… 
                  Alignment       =   2  'Center
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "Size"
                  Height          =   240
                  Index           =   22
                  Left            =   645
                  TabIndex        =   317
                  Top             =   1155
                  Width           =   900
               End
               Begin VB.Label lbl«”„«·ÊÕœ… 
                  Alignment       =   2  'Center
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "Group"
                  Height          =   255
                  Index           =   24
                  Left            =   8520
                  TabIndex        =   316
                  Top             =   2205
                  Visible         =   0   'False
                  Width           =   1800
               End
               Begin VB.Label lbl«”„«·ÊÕœ… 
                  Alignment       =   2  'Center
                  AutoSize        =   -1  'True
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "Origin"
                  Height          =   195
                  Index           =   34
                  Left            =   4005
                  RightToLeft     =   -1  'True
                  TabIndex        =   315
                  Top             =   1980
                  Width           =   390
               End
               Begin VB.Label lbl«”„«·ÊÕœ… 
                  Alignment       =   2  'Center
                  AutoSize        =   -1  'True
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "Service"
                  Height          =   195
                  Index           =   27
                  Left            =   3870
                  RightToLeft     =   -1  'True
                  TabIndex        =   314
                  Top             =   6270
                  Width           =   525
               End
               Begin VB.Label lbl«”„«·ÊÕœ… 
                  Alignment       =   2  'Center
                  AutoSize        =   -1  'True
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "Usage"
                  Height          =   195
                  Index           =   23
                  Left            =   3870
                  RightToLeft     =   -1  'True
                  TabIndex        =   313
                  Top             =   5280
                  Width           =   525
               End
               Begin VB.Label lbl«”„«·ÊÕœ… 
                  Alignment       =   2  'Center
                  AutoSize        =   -1  'True
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "Packing"
                  Height          =   195
                  Index           =   25
                  Left            =   3870
                  RightToLeft     =   -1  'True
                  TabIndex        =   312
                  Top             =   5895
                  Width           =   525
               End
            End
            Begin C1SizerLibCtl.C1Elastic ELe 
               Height          =   6915
               Index           =   27
               Left            =   45
               TabIndex        =   332
               TabStop         =   0   'False
               Top             =   45
               Width           =   11355
               _cx             =   20029
               _cy             =   12197
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
                  Height          =   6825
                  Index           =   8
                  Left            =   15240
                  TabIndex        =   333
                  Top             =   585
                  Width           =   11070
                  _cx             =   19526
                  _cy             =   12039
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
                  FormatString    =   $"FrmVizitScreen.frx":4524B
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
                  Height          =   6945
                  Index           =   28
                  Left            =   0
                  TabIndex        =   334
                  TabStop         =   0   'False
                  Top             =   0
                  Width           =   12255
                  _cx             =   21616
                  _cy             =   12250
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
                  Begin MSDataListLib.DataCombo cmbEyeDet 
                     Height          =   315
                     Index           =   18
                     Left            =   0
                     TabIndex        =   335
                     Top             =   8355
                     Width           =   1500
                     _ExtentX        =   2646
                     _ExtentY        =   556
                     _Version        =   393216
                     Style           =   2
                     Text            =   ""
                     RightToLeft     =   -1  'True
                  End
                  Begin VSFlex8Ctl.VSFlexGrid Grid3 
                     Height          =   2280
                     Left            =   0
                     TabIndex        =   354
                     Top             =   0
                     Width           =   10920
                     _cx             =   19262
                     _cy             =   4022
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
                     FormatString    =   $"FrmVizitScreen.frx":4530B
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
                     Height          =   1890
                     Left            =   0
                     TabIndex        =   355
                     Top             =   2310
                     Width           =   11055
                     _cx             =   19500
                     _cy             =   3334
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
                     FormatString    =   $"FrmVizitScreen.frx":4539A
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
                     Height          =   2565
                     Left            =   0
                     TabIndex        =   356
                     Top             =   4680
                     Width           =   10920
                     _cx             =   19262
                     _cy             =   4524
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
                     FormatString    =   $"FrmVizitScreen.frx":4540C
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
                  Begin VB.Label lbl«”„«·ÊÕœ… 
                     Alignment       =   2  'Center
                     BackColor       =   &H00E2E9E9&
                     Caption         =   "Base Curve"
                     Height          =   510
                     Index           =   26
                     Left            =   1500
                     TabIndex        =   336
                     Top             =   8505
                     Width           =   735
                  End
               End
            End
            Begin C1SizerLibCtl.C1Elastic ELe 
               Height          =   6915
               Index           =   9
               Left            =   12090
               TabIndex        =   337
               TabStop         =   0   'False
               Top             =   45
               Width           =   11355
               _cx             =   20029
               _cy             =   12197
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
               Begin MSDataListLib.DataCombo cmbEyeDet55 
                  Height          =   315
                  Index           =   26
                  Left            =   0
                  TabIndex        =   338
                  Top             =   6960
                  Width           =   1545
                  _ExtentX        =   2725
                  _ExtentY        =   556
                  _Version        =   393216
                  Style           =   2
                  Text            =   ""
                  RightToLeft     =   -1  'True
               End
               Begin MSDataListLib.DataCombo cmbEyeDet5241654 
                  Height          =   315
                  Index           =   27
                  Left            =   0
                  TabIndex        =   339
                  Top             =   8805
                  Width           =   1545
                  _ExtentX        =   2725
                  _ExtentY        =   556
                  _Version        =   393216
                  Style           =   2
                  Text            =   ""
                  RightToLeft     =   -1  'True
               End
               Begin VB.Label lbl«”„«·ÊÕœ… 
                  Alignment       =   2  'Center
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "Base Curve"
                  Height          =   510
                  Index           =   1
                  Left            =   1425
                  TabIndex        =   341
                  Top             =   8955
                  Width           =   765
               End
               Begin VB.Label lbl«”„«·ÊÕœ… 
                  Alignment       =   2  'Center
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "Breaking"
                  Height          =   405
                  Index           =   0
                  Left            =   1545
                  TabIndex        =   340
                  Top             =   7140
                  Width           =   645
               End
            End
         End
         Begin VB.Label LabCurr_Rec 
            BackColor       =   &H00E2E9E9&
            Height          =   240
            Index           =   3
            Left            =   3630
            RightToLeft     =   -1  'True
            TabIndex        =   246
            Top             =   7785
            Width           =   1530
         End
         Begin VB.Label Label2 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "⁄œœ «·”Ã·« :"
            Height          =   225
            Index           =   11
            Left            =   1680
            RightToLeft     =   -1  'True
            TabIndex        =   245
            Top             =   7770
            Width           =   1665
         End
         Begin VB.Label Label2 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "«·”Ã· «·Õ«·Ì:"
            Height          =   225
            Index           =   10
            Left            =   5310
            RightToLeft     =   -1  'True
            TabIndex        =   244
            Top             =   7770
            Width           =   1515
         End
         Begin VB.Label LabCount_Rec 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Height          =   240
            Index           =   3
            Left            =   0
            RightToLeft     =   -1  'True
            TabIndex        =   243
            Top             =   7680
            Width           =   1110
         End
      End
      Begin C1SizerLibCtl.C1Elastic ELe 
         Height          =   8790
         Index           =   6
         Left            =   -20235
         TabIndex        =   265
         TabStop         =   0   'False
         Top             =   45
         Width           =   18690
         _cx             =   32967
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
         Begin VB.CommandButton Command2 
            Caption         =   "⁄—÷"
            Height          =   375
            Left            =   4755
            RightToLeft     =   -1  'True
            TabIndex        =   368
            Top             =   840
            Width           =   1245
         End
         Begin VB.Frame Frame9 
            Height          =   615
            Left            =   6270
            RightToLeft     =   -1  'True
            TabIndex        =   365
            Top             =   720
            Width           =   4605
            Begin VB.OptionButton Option1 
               Alignment       =   1  'Right Justify
               Caption         =   "«·ÿ·»«  «·œ«Œ·Ì…  ÕÊÌ·"
               Height          =   375
               Index           =   1
               Left            =   120
               RightToLeft     =   -1  'True
               TabIndex        =   367
               Top             =   120
               Width           =   1815
            End
            Begin VB.OptionButton Option1 
               Alignment       =   1  'Right Justify
               Caption         =   "«·ÿ·»«  «·œ«Œ·Ì… ‘—«¡"
               Height          =   375
               Index           =   0
               Left            =   2040
               RightToLeft     =   -1  'True
               TabIndex        =   366
               Top             =   120
               Value           =   -1  'True
               Width           =   1815
            End
         End
         Begin VB.Frame Frame8 
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   0  'None
            Height          =   660
            Index           =   0
            Left            =   0
            RightToLeft     =   -1  'True
            TabIndex        =   266
            Top             =   0
            Width           =   18690
            Begin VB.Label Label1 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   " ‰»ÌÂ«  «·ÿ·»«  «·œ«Œ·Ì… ( ÕÊÌ· - ‘—«¡) "
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
               Left            =   8235
               RightToLeft     =   -1  'True
               TabIndex        =   267
               Top             =   210
               Width           =   5010
            End
         End
         Begin VSFlex8UCtl.VSFlexGrid FgItems 
            Height          =   8670
            Index           =   4
            Left            =   25935
            TabIndex        =   268
            Top             =   765
            Width           =   18555
            _cx             =   32729
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
            FormatString    =   $"FrmVizitScreen.frx":455CA
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
         Begin VSFlex8UCtl.VSFlexGrid grdTransfer 
            Height          =   7275
            Left            =   270
            TabIndex        =   269
            Top             =   1680
            Width           =   18420
            _cx             =   32491
            _cy             =   12832
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
            HighLight       =   0
            AllowSelection  =   -1  'True
            AllowBigSelection=   0   'False
            AllowUserResizing=   1
            SelectionMode   =   0
            GridLines       =   1
            GridLinesFixed  =   2
            GridLineWidth   =   1
            Rows            =   50
            Cols            =   21
            FixedRows       =   1
            FixedCols       =   1
            RowHeightMin    =   320
            RowHeightMax    =   0
            ColWidthMin     =   0
            ColWidthMax     =   0
            ExtendLastCol   =   -1  'True
            FormatString    =   $"FrmVizitScreen.frx":4568A
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
         Begin MSDataListLib.DataCombo cmbStoreID 
            Height          =   315
            Left            =   11025
            TabIndex        =   280
            Top             =   1020
            Width           =   4740
            _ExtentX        =   8361
            _ExtentY        =   556
            _Version        =   393216
            ListField       =   "6"
            Text            =   ""
            RightToLeft     =   -1  'True
         End
         Begin VB.Label Label3 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "«·„Œ“‰ «·„ÿ·Ê» „‰Â"
            Height          =   285
            Index           =   5
            Left            =   15900
            RightToLeft     =   -1  'True
            TabIndex        =   281
            Top             =   1050
            Width           =   1680
         End
      End
      Begin C1SizerLibCtl.C1Elastic ELe 
         Height          =   8790
         Index           =   7
         Left            =   -19935
         TabIndex        =   270
         TabStop         =   0   'False
         Top             =   45
         Width           =   18690
         _cx             =   32967
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
         Begin VB.Frame Frame8 
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   0  'None
            Height          =   660
            Index           =   1
            Left            =   0
            RightToLeft     =   -1  'True
            TabIndex        =   271
            Top             =   0
            Width           =   18420
            Begin VB.Label Label1 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "«·„⁄—÷"
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
               Index           =   16
               Left            =   8235
               RightToLeft     =   -1  'True
               TabIndex        =   272
               Top             =   210
               Width           =   5010
            End
         End
         Begin VSFlex8UCtl.VSFlexGrid FgItems 
            Height          =   8670
            Index           =   5
            Left            =   25935
            TabIndex        =   273
            Top             =   765
            Width           =   18555
            _cx             =   32729
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
            FormatString    =   $"FrmVizitScreen.frx":459D8
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
         Begin VSFlex8UCtl.VSFlexGrid VSFlexGrid2 
            Height          =   3105
            Left            =   135
            TabIndex        =   274
            Top             =   750
            Width           =   18285
            _cx             =   32253
            _cy             =   5477
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
            FormatString    =   $"FrmVizitScreen.frx":45A98
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
         Begin VSFlex8UCtl.VSFlexGrid VSFlexGrid3 
            Height          =   4035
            Left            =   270
            TabIndex        =   275
            Top             =   4470
            Width           =   18420
            _cx             =   32491
            _cy             =   7117
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
            FormatString    =   $"FrmVizitScreen.frx":45BCD
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
         Index           =   8
         Left            =   -19635
         TabIndex        =   276
         TabStop         =   0   'False
         Top             =   45
         Width           =   18690
         _cx             =   32967
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
         Begin VB.Frame Frame8 
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   0  'None
            Height          =   690
            Index           =   2
            Left            =   0
            RightToLeft     =   -1  'True
            TabIndex        =   277
            Top             =   0
            Width           =   18690
            Begin VB.Label Label1 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   " ‰»ÌÂ«  «·„⁄„·"
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
               Index           =   18
               Left            =   8235
               RightToLeft     =   -1  'True
               TabIndex        =   278
               Top             =   210
               Width           =   5010
            End
         End
         Begin VSFlex8UCtl.VSFlexGrid FgItems 
            Height          =   8670
            Index           =   6
            Left            =   25935
            TabIndex        =   279
            Top             =   765
            Width           =   18555
            _cx             =   32729
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
            FormatString    =   $"FrmVizitScreen.frx":45DA7
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
         Begin VSFlex8UCtl.VSFlexGrid grdTransfer2 
            Height          =   7425
            Left            =   135
            TabIndex        =   361
            Top             =   1380
            Width           =   18420
            _cx             =   32491
            _cy             =   13097
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
            HighLight       =   0
            AllowSelection  =   -1  'True
            AllowBigSelection=   0   'False
            AllowUserResizing=   1
            SelectionMode   =   0
            GridLines       =   1
            GridLinesFixed  =   2
            GridLineWidth   =   1
            Rows            =   50
            Cols            =   29
            FixedRows       =   1
            FixedCols       =   1
            RowHeightMin    =   320
            RowHeightMax    =   0
            ColWidthMin     =   0
            ColWidthMax     =   0
            ExtendLastCol   =   -1  'True
            FormatString    =   $"FrmVizitScreen.frx":45E67
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
         Begin MSDataListLib.DataCombo cmbStoreID2 
            Height          =   315
            Left            =   11715
            TabIndex        =   362
            Top             =   930
            Width           =   4740
            _ExtentX        =   8361
            _ExtentY        =   556
            _Version        =   393216
            ListField       =   "6"
            Text            =   ""
            RightToLeft     =   -1  'True
         End
         Begin VB.Label Label3 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "„Œ“‰ «·„⁄„· "
            Height          =   285
            Index           =   7
            Left            =   16590
            RightToLeft     =   -1  'True
            TabIndex        =   363
            Top             =   960
            Width           =   1680
         End
      End
      Begin C1SizerLibCtl.C1Elastic ELe 
         Height          =   8790
         Index           =   10
         Left            =   -19335
         TabIndex        =   370
         TabStop         =   0   'False
         Top             =   45
         Width           =   18690
         _cx             =   32967
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
         Begin VB.Frame Fra_Header 
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   0  'None
            Height          =   585
            Index           =   3
            Left            =   -135
            RightToLeft     =   -1  'True
            TabIndex        =   375
            Top             =   0
            Width           =   18690
            Begin VB.TextBox Text2 
               Alignment       =   1  'Right Justify
               Appearance      =   0  'Flat
               BackColor       =   &H000000FF&
               Height          =   240
               Index           =   3
               Left            =   3030
               RightToLeft     =   -1  'True
               TabIndex        =   377
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
               Index           =   7
               Left            =   2580
               RightToLeft     =   -1  'True
               TabIndex        =   376
               Text            =   "modflag"
               Top             =   90
               Visible         =   0   'False
               Width           =   465
            End
            Begin MSComctlLib.ImageList GrdImageList2 
               Index           =   7
               Left            =   3120
               Top             =   30
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
                     Picture         =   "FrmVizitScreen.frx":462F3
                     Key             =   "CompanyName"
                  EndProperty
                  BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                     Picture         =   "FrmVizitScreen.frx":4668D
                     Key             =   "Ser"
                  EndProperty
                  BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                     Picture         =   "FrmVizitScreen.frx":46A27
                     Key             =   "Vac_Name"
                  EndProperty
                  BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                     Picture         =   "FrmVizitScreen.frx":46DC1
                     Key             =   "ShareCount"
                  EndProperty
                  BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                     Picture         =   "FrmVizitScreen.frx":4715B
                     Key             =   "Dis_Count"
                  EndProperty
                  BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                     Picture         =   "FrmVizitScreen.frx":474F5
                     Key             =   "Bouns"
                  EndProperty
                  BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                     Picture         =   "FrmVizitScreen.frx":4788F
                     Key             =   "SharesValue"
                  EndProperty
                  BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                     Picture         =   "FrmVizitScreen.frx":47E29
                     Key             =   "BuyValue"
                  EndProperty
               EndProperty
            End
            Begin ImpulseButton.ISButton btn_Last 
               Height          =   315
               Index           =   7
               Left            =   90
               TabIndex        =   378
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
               ButtonImage     =   "FrmVizitScreen.frx":481C3
               ColorButton     =   14871017
               AcclimateGrayTones=   -1  'True
               DrawFocusRectangle=   0   'False
               DisabledImageExtraction=   0
               DisabledImageStyle=   1
            End
            Begin ImpulseButton.ISButton btn_Next 
               Height          =   315
               Index           =   7
               Left            =   555
               TabIndex        =   379
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
               ButtonImage     =   "FrmVizitScreen.frx":4855D
               ColorButton     =   14871017
               DrawFocusRectangle=   0   'False
               DisabledImageStyle=   1
            End
            Begin ImpulseButton.ISButton btn_Previous 
               Height          =   315
               Index           =   7
               Left            =   1155
               TabIndex        =   380
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
               ButtonImage     =   "FrmVizitScreen.frx":488F7
               ColorButton     =   14871017
               DrawFocusRectangle=   0   'False
               DisabledImageStyle=   1
            End
            Begin ImpulseButton.ISButton btn_First 
               Height          =   315
               Index           =   7
               Left            =   1620
               TabIndex        =   381
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
               ButtonImage     =   "FrmVizitScreen.frx":48C91
               ColorButton     =   14871017
               DrawFocusRectangle=   0   'False
               DisabledImageStyle=   1
            End
            Begin MSComDlg.CommonDialog CD1 
               Left            =   0
               Top             =   0
               _ExtentX        =   847
               _ExtentY        =   847
               _Version        =   393216
            End
            Begin VB.Label Label1 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "›Ê« Ì— «·„»Ì⁄« "
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
               Index           =   19
               Left            =   11340
               RightToLeft     =   -1  'True
               TabIndex        =   382
               Top             =   90
               Width           =   2640
            End
         End
         Begin VB.Frame Frame1 
            BackColor       =   &H00E2E9E9&
            BorderStyle     =   0  'None
            Height          =   1335
            Index           =   1
            Left            =   270
            RightToLeft     =   -1  'True
            TabIndex        =   371
            Top             =   780
            Width           =   18150
            Begin VB.CheckBox chkIsDateLine 
               Alignment       =   1  'Right Justify
               Caption         =   "«÷«›…  «—ÌŒ «·”ÿ— ··ﬁÌœ"
               Height          =   315
               Left            =   3420
               RightToLeft     =   -1  'True
               TabIndex        =   421
               Top             =   870
               Width           =   2055
            End
            Begin VB.CheckBox chkIsAddOnly 
               Alignment       =   1  'Right Justify
               Caption         =   "«” Ì—«œ «·«÷«›… ›ﬁÿ"
               Height          =   315
               Left            =   10140
               RightToLeft     =   -1  'True
               TabIndex        =   418
               Top             =   1080
               Width           =   2055
            End
            Begin VB.CheckBox chkIsDiscountOnly 
               Alignment       =   1  'Right Justify
               Caption         =   "«” Ì—«œ «·Œ’„ ›ﬁÿ"
               Height          =   315
               Left            =   12510
               RightToLeft     =   -1  'True
               TabIndex        =   417
               Top             =   1080
               Width           =   2055
            End
            Begin VB.CommandButton cmdDelNote7 
               Caption         =   "Õ–› «·ﬁÌœ "
               Height          =   450
               Left            =   2265
               RightToLeft     =   -1  'True
               TabIndex        =   414
               Top             =   390
               Visible         =   0   'False
               Width           =   2115
            End
            Begin VB.CommandButton cmdPrintNote7 
               Caption         =   "ÿ»«⁄Â «·ﬁÌœ"
               Height          =   450
               Left            =   120
               RightToLeft     =   -1  'True
               TabIndex        =   413
               Top             =   390
               Width           =   2115
            End
            Begin VB.CheckBox chkIsVat 
               Alignment       =   1  'Right Justify
               Caption         =   "«·„»·€ ‘«„· «·÷—Ì»…"
               Height          =   315
               Left            =   9060
               RightToLeft     =   -1  'True
               TabIndex        =   406
               Top             =   780
               Width           =   2055
            End
            Begin VB.TextBox txtNoteID7 
               Alignment       =   1  'Right Justify
               Height          =   315
               Left            =   390
               RightToLeft     =   -1  'True
               TabIndex        =   405
               Top             =   -30
               Visible         =   0   'False
               Width           =   2280
            End
            Begin VB.TextBox TxtNoteSerial7 
               Alignment       =   1  'Right Justify
               Enabled         =   0   'False
               Height          =   450
               Left            =   5130
               Locked          =   -1  'True
               RightToLeft     =   -1  'True
               TabIndex        =   403
               Top             =   450
               Width           =   3255
            End
            Begin VB.TextBox TxtNoteSerial17 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00FFFFC0&
               Height          =   315
               Left            =   13080
               Locked          =   -1  'True
               RightToLeft     =   -1  'True
               TabIndex        =   402
               Top             =   30
               Width           =   1455
            End
            Begin VB.CommandButton CmdCreateV7 
               Caption         =   "«‰‘«¡ «·ﬁÌœ"
               Height          =   285
               Left            =   5610
               TabIndex        =   401
               Top             =   900
               Visible         =   0   'False
               Width           =   1485
            End
            Begin VB.TextBox Text3 
               Height          =   285
               Left            =   7170
               Locked          =   -1  'True
               TabIndex        =   400
               Top             =   60
               Visible         =   0   'False
               Width           =   975
            End
            Begin VB.CommandButton Command5 
               Caption         =   " ÕœÌœ «·„·›..."
               Height          =   255
               Left            =   12690
               RightToLeft     =   -1  'True
               TabIndex        =   399
               Top             =   780
               Width           =   1305
            End
            Begin VB.CommandButton Command4 
               Caption         =   " Õ„Ì· «·„·›..."
               Height          =   285
               Left            =   11190
               TabIndex        =   398
               Top             =   750
               Width           =   1485
            End
            Begin VB.ComboBox Combo1 
               BackColor       =   &H80000018&
               Height          =   315
               ItemData        =   "FrmVizitScreen.frx":4902B
               Left            =   2280
               List            =   "FrmVizitScreen.frx":4903B
               RightToLeft     =   -1  'True
               Style           =   2  'Dropdown List
               TabIndex        =   373
               Top             =   3150
               Visible         =   0   'False
               Width           =   1005
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
               Index           =   7
               Left            =   360
               MaxLength       =   50
               RightToLeft     =   -1  'True
               TabIndex        =   372
               Top             =   -60
               Visible         =   0   'False
               Width           =   1065
            End
            Begin MSDataListLib.DataCombo dcBranch 
               Bindings        =   "FrmVizitScreen.frx":49054
               Height          =   315
               Index           =   7
               Left            =   2880
               TabIndex        =   407
               Top             =   0
               Width           =   2895
               _ExtentX        =   5106
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
            Begin MSComCtl2.DTPicker XPDtbTrans7 
               Height          =   300
               Left            =   10170
               TabIndex        =   410
               Top             =   0
               Width           =   1560
               _ExtentX        =   2752
               _ExtentY        =   529
               _Version        =   393216
               Format          =   143196161
               CurrentDate     =   38784
            End
            Begin MSDataListLib.DataCombo DcboBankName 
               Height          =   315
               Left            =   10200
               TabIndex        =   411
               Top             =   420
               Width           =   3795
               _ExtentX        =   6694
               _ExtentY        =   556
               _Version        =   393216
               Text            =   ""
               RightToLeft     =   -1  'True
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "«”„ «·»‰ﬂ"
               Height          =   285
               Index           =   29
               Left            =   13800
               RightToLeft     =   -1  'True
               TabIndex        =   412
               Top             =   420
               Width           =   1215
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "«·›—⁄"
               Height          =   255
               Index           =   28
               Left            =   5790
               TabIndex        =   409
               Top             =   0
               Width           =   600
            End
            Begin VB.Label lbl 
               Alignment       =   2  'Center
               BackColor       =   &H00E2E9E9&
               Caption         =   " «—ÌŒ «·”‰œ"
               Height          =   270
               Index           =   27
               Left            =   11820
               TabIndex        =   408
               Top             =   30
               Width           =   1215
            End
            Begin VB.Label Label1 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               BackStyle       =   0  'Transparent
               Caption         =   "—ﬁ„ «·ﬁÌœ"
               Height          =   405
               Index           =   20
               Left            =   8385
               RightToLeft     =   -1  'True
               TabIndex        =   404
               Top             =   570
               Width           =   1095
            End
            Begin VB.Label Label6 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "«·ﬂÊœ "
               Height          =   195
               Index           =   2
               Left            =   13980
               RightToLeft     =   -1  'True
               TabIndex        =   374
               Top             =   30
               Width           =   990
            End
         End
         Begin ImpulseButton.ISButton btn_New 
            Height          =   345
            Index           =   7
            Left            =   7245
            TabIndex        =   383
            Top             =   7740
            Width           =   1125
            _ExtentX        =   1984
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
            ButtonImage     =   "FrmVizitScreen.frx":49069
            ColorButton     =   14871017
            DrawFocusRectangle=   0   'False
         End
         Begin ImpulseButton.ISButton btn_Save 
            Height          =   345
            Index           =   7
            Left            =   5310
            TabIndex        =   384
            Top             =   7710
            Width           =   825
            _ExtentX        =   1455
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
            ButtonImage     =   "FrmVizitScreen.frx":49403
            ColorButton     =   14871017
            DrawFocusRectangle=   0   'False
         End
         Begin ImpulseButton.ISButton btn_Modify 
            Height          =   345
            Index           =   7
            Left            =   6270
            TabIndex        =   385
            Top             =   7770
            Width           =   975
            _ExtentX        =   1720
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
            ButtonImage     =   "FrmVizitScreen.frx":4979D
            ColorButton     =   14871017
            DrawFocusRectangle=   0   'False
         End
         Begin ImpulseButton.ISButton Btn_Undo 
            Height          =   345
            Index           =   7
            Left            =   4320
            TabIndex        =   386
            Top             =   7680
            Width           =   990
            _ExtentX        =   1746
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
            ButtonImage     =   "FrmVizitScreen.frx":49B37
            ColorButton     =   14871017
            DrawFocusRectangle=   0   'False
         End
         Begin ImpulseButton.ISButton btn_Delete 
            Height          =   345
            Index           =   7
            Left            =   3480
            TabIndex        =   387
            Top             =   7710
            Width           =   840
            _ExtentX        =   1482
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
            ButtonImage     =   "FrmVizitScreen.frx":49ED1
            ColorButton     =   14871017
            DrawFocusRectangle=   0   'False
         End
         Begin ImpulseButton.ISButton Btn_Update 
            Height          =   225
            Index           =   7
            Left            =   6555
            TabIndex        =   388
            TabStop         =   0   'False
            ToolTipText     =   " ÕœÌÀ ﬁ«⁄œ… «·»Ì«‰« "
            Top             =   6450
            Visible         =   0   'False
            Width           =   1815
            _ExtentX        =   3201
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
            ButtonImage     =   "FrmVizitScreen.frx":4A46B
            ColorButton     =   14871017
            DrawFocusRectangle=   0   'False
         End
         Begin ImpulseButton.ISButton btn_Cancel 
            Height          =   345
            Index           =   7
            Left            =   135
            TabIndex        =   389
            Top             =   7650
            Width           =   975
            _ExtentX        =   1720
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
            ButtonImage     =   "FrmVizitScreen.frx":4A805
            ColorButton     =   14871017
            DrawFocusRectangle=   0   'False
            DisabledImageStyle=   1
         End
         Begin ImpulseButton.ISButton Btn_Print 
            Height          =   435
            Index           =   7
            Left            =   2370
            TabIndex        =   390
            TabStop         =   0   'False
            ToolTipText     =   "ÿ»«⁄… «·»Ì«‰«  "
            Top             =   7650
            Width           =   975
            _ExtentX        =   1720
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
            ButtonImage     =   "FrmVizitScreen.frx":4AB9F
            ColorButton     =   14871017
            DrawFocusRectangle=   0   'False
            DisabledImageStyle=   1
         End
         Begin ImpulseButton.ISButton btn_Query 
            Height          =   540
            Index           =   7
            Left            =   1530
            TabIndex        =   391
            TabStop         =   0   'False
            ToolTipText     =   "(Ctrl+F)  ··»ÕÀ ≈÷€ÿ Â–« «·„› «Õ √Ê ≈÷€ÿ "
            Top             =   7575
            Width           =   705
            _ExtentX        =   1244
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
            ButtonImage     =   "FrmVizitScreen.frx":51401
            ColorButton     =   14871017
            DrawFocusRectangle=   0   'False
         End
         Begin VSFlex8Ctl.VSFlexGrid GrdExcel 
            Height          =   3990
            Left            =   420
            TabIndex        =   392
            Top             =   2370
            Width           =   18540
            _cx             =   32702
            _cy             =   7038
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
            Cols            =   12
            FixedRows       =   1
            FixedCols       =   1
            RowHeightMin    =   320
            RowHeightMax    =   0
            ColWidthMin     =   0
            ColWidthMax     =   0
            ExtendLastCol   =   -1  'True
            FormatString    =   $"FrmVizitScreen.frx":5179B
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
         Begin MSDataListLib.DataCombo DCboUserName 
            Height          =   315
            Index           =   7
            Left            =   13620
            TabIndex        =   415
            Top             =   6720
            Width           =   3360
            _ExtentX        =   5927
            _ExtentY        =   556
            _Version        =   393216
            Enabled         =   0   'False
            Text            =   ""
            RightToLeft     =   -1  'True
         End
         Begin ImpulseButton.ISButton Cmd_DeleteRow 
            Height          =   300
            Index           =   7
            Left            =   2685
            TabIndex        =   419
            Top             =   6465
            Width           =   1800
            _ExtentX        =   3175
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
            ButtonImage     =   "FrmVizitScreen.frx":51982
            ColorButton     =   14871017
            DrawFocusRectangle=   0   'False
         End
         Begin ImpulseButton.ISButton Cmd_DeleteAll 
            Height          =   300
            Index           =   7
            Left            =   450
            TabIndex        =   420
            Top             =   6450
            Width           =   2100
            _ExtentX        =   3704
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
            ButtonImage     =   "FrmVizitScreen.frx":51F1C
            ColorButton     =   14871017
            DrawFocusRectangle=   0   'False
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Õ—— »Ê«”ÿ…  "
            Height          =   315
            Index           =   30
            Left            =   17130
            TabIndex        =   416
            Top             =   6660
            Width           =   1245
         End
         Begin VB.Label LabCount_Rec 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Height          =   240
            Index           =   7
            Left            =   1245
            RightToLeft     =   -1  'True
            TabIndex        =   397
            Top             =   7185
            Width           =   1275
         End
         Begin VB.Label LabCurr_Rec 
            BackColor       =   &H00E2E9E9&
            Height          =   240
            Index           =   7
            Left            =   4890
            RightToLeft     =   -1  'True
            TabIndex        =   396
            Top             =   7185
            Width           =   1530
         End
         Begin VB.Label Label2 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "⁄œœ «·”Ã·« :"
            Height          =   225
            Index           =   13
            Left            =   2850
            RightToLeft     =   -1  'True
            TabIndex        =   395
            Top             =   7200
            Width           =   2085
         End
         Begin VB.Label Label2 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "«·”Ã· «·Õ«·Ì:"
            Height          =   225
            Index           =   12
            Left            =   6690
            RightToLeft     =   -1  'True
            TabIndex        =   394
            Top             =   7170
            Width           =   1545
         End
         Begin VB.Label LabCurr_Rec 
            BackColor       =   &H00E2E9E9&
            Height          =   240
            Index           =   4
            Left            =   4320
            RightToLeft     =   -1  'True
            TabIndex        =   393
            Top             =   6720
            Width           =   1545
         End
      End
      Begin C1SizerLibCtl.C1Elastic ELe 
         Height          =   8790
         Index           =   11
         Left            =   45
         TabIndex        =   422
         TabStop         =   0   'False
         Top             =   45
         Width           =   18690
         _cx             =   32967
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
         Begin VB.TextBox txtTotalStill 
            Alignment       =   1  'Right Justify
            Height          =   465
            Left            =   7470
            RightToLeft     =   -1  'True
            TabIndex        =   467
            Top             =   8130
            Visible         =   0   'False
            Width           =   2805
         End
         Begin VB.CommandButton CmdSelectCus 
            Caption         =   " ÕœÌœ>>"
            Height          =   330
            Left            =   5130
            RightToLeft     =   -1  'True
            TabIndex        =   466
            Top             =   4020
            Width           =   4320
         End
         Begin VB.CommandButton CmdSelectEmp 
            Caption         =   " ÕœÌœ>>"
            Height          =   330
            Left            =   5160
            RightToLeft     =   -1  'True
            TabIndex        =   465
            Top             =   4830
            Width           =   4320
         End
         Begin VB.TextBox Text4 
            Alignment       =   1  'Right Justify
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   178
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   465
            Left            =   3315
            RightToLeft     =   -1  'True
            TabIndex        =   435
            Text            =   "Text1"
            Top             =   7320
            Visible         =   0   'False
            Width           =   390
         End
         Begin VB.OptionButton Option1 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "⁄„·«¡"
            ForeColor       =   &H00800000&
            Height          =   255
            Index           =   2
            Left            =   8220
            RightToLeft     =   -1  'True
            TabIndex        =   434
            Top             =   1260
            Value           =   -1  'True
            Width           =   1485
         End
         Begin VB.OptionButton Option2 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "„Ê—œÌ‰"
            ForeColor       =   &H00C00000&
            Height          =   255
            Left            =   6615
            RightToLeft     =   -1  'True
            TabIndex        =   433
            Top             =   1230
            Width           =   1455
         End
         Begin VB.Frame Frame12 
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   178
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   660
            Left            =   30
            RightToLeft     =   -1  'True
            TabIndex        =   430
            Top             =   -150
            Width           =   17475
            Begin VB.Label Label1 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "   ﬁ«—Ì— «⁄„«— «·œÌÊ‰"
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
               Left            =   7800
               RightToLeft     =   -1  'True
               TabIndex        =   431
               Top             =   300
               Width           =   3390
            End
         End
         Begin VB.TextBox CurrenrEmployeeIDs 
            Alignment       =   1  'Right Justify
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   178
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   4440
            RightToLeft     =   -1  'True
            TabIndex        =   429
            Top             =   7080
            Visible         =   0   'False
            Width           =   1575
         End
         Begin VB.TextBox StrCusID 
            Alignment       =   1  'Right Justify
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   178
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   4440
            RightToLeft     =   -1  'True
            TabIndex        =   428
            Top             =   7200
            Visible         =   0   'False
            Width           =   1575
         End
         Begin VB.CommandButton Command6 
            BackColor       =   &H00E2E9E9&
            Caption         =   "„”Õ"
            Height          =   555
            Left            =   5640
            RightToLeft     =   -1  'True
            TabIndex        =   427
            Top             =   6000
            Width           =   1560
         End
         Begin VB.Frame Frame10 
            BackColor       =   &H00E2E9E9&
            Caption         =   "ÿ»ﬁ« · «—ÌŒ"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   9.75
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   615
            Left            =   5280
            RightToLeft     =   -1  'True
            TabIndex        =   424
            Top             =   480
            Width           =   5265
            Begin VB.OptionButton Rd 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "«’œ«— «·›« Ê—…"
               ForeColor       =   &H00800000&
               Height          =   255
               Index           =   0
               Left            =   2700
               RightToLeft     =   -1  'True
               TabIndex        =   426
               Top             =   240
               Width           =   1605
            End
            Begin VB.OptionButton Rd 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "«·«” Õﬁ«ﬁ"
               ForeColor       =   &H00800000&
               Height          =   255
               Index           =   1
               Left            =   720
               RightToLeft     =   -1  'True
               TabIndex        =   425
               Top             =   240
               Value           =   -1  'True
               Width           =   1245
            End
         End
         Begin VB.TextBox TxtSearchCode 
            Alignment       =   2  'Center
            Height          =   315
            Left            =   10860
            TabIndex        =   423
            Top             =   3570
            Width           =   1710
         End
         Begin XtremeSuiteControls.CheckBox CheckEmp 
            Height          =   375
            Left            =   13290
            TabIndex        =   432
            Top             =   4200
            Width           =   3075
            _Version        =   786432
            _ExtentX        =   5424
            _ExtentY        =   661
            _StockProps     =   79
            Caption         =   "„‰œÊ» „Õœœ"
            BackColor       =   14871017
            UseVisualStyle  =   -1  'True
            TextAlignment   =   1
            RightToLeft     =   -1  'True
         End
         Begin MSComCtl2.DTPicker DTP_Date 
            Height          =   345
            Left            =   105
            TabIndex        =   437
            TabStop         =   0   'False
            Top             =   750
            Width           =   1680
            _ExtentX        =   2963
            _ExtentY        =   609
            _Version        =   393216
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   178
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            CalendarBackColor=   12648447
            CalendarTitleBackColor=   10383715
            CustomFormat    =   "yyyy/M/d"
            Format          =   138870787
            CurrentDate     =   37140
         End
         Begin ImpulseButton.ISButton CmdPrint 
            Height          =   330
            Left            =   4080
            TabIndex        =   438
            Top             =   7080
            Visible         =   0   'False
            Width           =   1680
            _ExtentX        =   2963
            _ExtentY        =   582
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
            ButtonImage     =   "FrmVizitScreen.frx":524B6
            ColorButton     =   14871017
            DrawFocusRectangle=   0   'False
         End
         Begin MSDataListLib.DataCombo dcBranch 
            Height          =   315
            Index           =   8
            Left            =   5100
            TabIndex        =   439
            Top             =   2610
            Width           =   7545
            _ExtentX        =   13309
            _ExtentY        =   556
            _Version        =   393216
            BackColor       =   16777215
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
         Begin MSDataListLib.DataCombo DcbEmployee 
            Height          =   315
            Left            =   5220
            TabIndex        =   440
            Top             =   4410
            Width           =   5295
            _ExtentX        =   9340
            _ExtentY        =   556
            _Version        =   393216
            BackColor       =   16777215
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
         Begin MSDataListLib.DataCombo DBCboClientName 
            Height          =   315
            Left            =   5130
            TabIndex        =   441
            Top             =   3570
            Width           =   5655
            _ExtentX        =   9975
            _ExtentY        =   556
            _Version        =   393216
            BackColor       =   16777215
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
         Begin ImpulseButton.ISButton BtnPrint 
            Height          =   555
            Index           =   0
            Left            =   3960
            TabIndex        =   442
            Top             =   6000
            Width           =   1560
            _ExtentX        =   2752
            _ExtentY        =   979
            ButtonPositionImage=   1
            Caption         =   "ÿ»«⁄…  Õ·Ì·Ì"
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
            ButtonImage     =   "FrmVizitScreen.frx":52850
            ColorButton     =   14871017
            DrawFocusRectangle=   0   'False
         End
         Begin ImpulseButton.ISButton BtnPrint 
            Height          =   555
            Index           =   1
            Left            =   2160
            TabIndex        =   443
            Top             =   6000
            Width           =   1560
            _ExtentX        =   2752
            _ExtentY        =   979
            ButtonPositionImage=   1
            Caption         =   "ÿ»«⁄… «Ã„«·Ì"
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
            ButtonImage     =   "FrmVizitScreen.frx":52BEA
            ColorButton     =   14871017
            DrawFocusRectangle=   0   'False
         End
         Begin XtremeSuiteControls.CheckBox ChekCustomer 
            Height          =   375
            Left            =   13290
            TabIndex        =   444
            Top             =   3120
            Width           =   3075
            _Version        =   786432
            _ExtentX        =   5424
            _ExtentY        =   661
            _StockProps     =   79
            Caption         =   "⁄„Ì·/„Ê—œ"
            BackColor       =   14871017
            UseVisualStyle  =   -1  'True
            TextAlignment   =   1
            RightToLeft     =   -1  'True
         End
         Begin XtremeSuiteControls.CheckBox CheckAllCustomer 
            Height          =   375
            Left            =   12210
            TabIndex        =   445
            Top             =   3480
            Width           =   4155
            _Version        =   786432
            _ExtentX        =   7329
            _ExtentY        =   661
            _StockProps     =   79
            Caption         =   "«Œ Ì«— «ﬂÀ— „‰ ⁄„Ì· /„Ê—œ"
            BackColor       =   14871017
            UseVisualStyle  =   -1  'True
            TextAlignment   =   1
            RightToLeft     =   -1  'True
         End
         Begin XtremeSuiteControls.CheckBox CheckAllEMp 
            Height          =   375
            Left            =   11970
            TabIndex        =   446
            Top             =   4680
            Visible         =   0   'False
            Width           =   4395
            _Version        =   786432
            _ExtentX        =   7752
            _ExtentY        =   661
            _StockProps     =   79
            Caption         =   "«Œ Ì«— «ﬂÀ— „‰ „‰œÊ»"
            BackColor       =   14871017
            UseVisualStyle  =   -1  'True
            TextAlignment   =   1
            RightToLeft     =   -1  'True
         End
         Begin MSComCtl2.DTPicker FromDate 
            Height          =   345
            Left            =   2925
            TabIndex        =   447
            TabStop         =   0   'False
            Top             =   7500
            Visible         =   0   'False
            Width           =   1680
            _ExtentX        =   2963
            _ExtentY        =   609
            _Version        =   393216
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   178
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            CalendarBackColor=   12648447
            CalendarTitleBackColor=   10383715
            CheckBox        =   -1  'True
            CustomFormat    =   "yyyy/M/d"
            Format          =   138870787
            CurrentDate     =   37140
         End
         Begin MSComCtl2.DTPicker ToDate 
            Height          =   345
            Left            =   30
            TabIndex        =   448
            TabStop         =   0   'False
            Top             =   7500
            Visible         =   0   'False
            Width           =   1680
            _ExtentX        =   2963
            _ExtentY        =   609
            _Version        =   393216
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   178
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            CalendarBackColor=   12648447
            CalendarTitleBackColor=   10383715
            CheckBox        =   -1  'True
            CustomFormat    =   "yyyy/M/d"
            Format          =   138870787
            CurrentDate     =   37140
         End
         Begin ImpulseButton.ISButton BtnPrint22 
            Height          =   555
            Left            =   90
            TabIndex        =   449
            Top             =   7860
            Visible         =   0   'False
            Width           =   1560
            _ExtentX        =   2752
            _ExtentY        =   979
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
            ButtonImage     =   "FrmVizitScreen.frx":52F84
            ColorButton     =   14871017
            DrawFocusRectangle=   0   'False
         End
         Begin MSComCtl2.DTPicker FromDate1 
            Height          =   345
            Left            =   2925
            TabIndex        =   450
            TabStop         =   0   'False
            Top             =   7860
            Visible         =   0   'False
            Width           =   1680
            _ExtentX        =   2963
            _ExtentY        =   609
            _Version        =   393216
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   178
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            CalendarBackColor=   12648447
            CalendarTitleBackColor=   10383715
            CheckBox        =   -1  'True
            CustomFormat    =   "yyyy/M/d"
            Format          =   138870787
            CurrentDate     =   37140
         End
         Begin MSComCtl2.DTPicker ToDate1 
            Height          =   345
            Left            =   30
            TabIndex        =   451
            TabStop         =   0   'False
            Top             =   7860
            Visible         =   0   'False
            Width           =   1680
            _ExtentX        =   2963
            _ExtentY        =   609
            _Version        =   393216
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   178
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            CalendarBackColor=   12648447
            CalendarTitleBackColor=   10383715
            CheckBox        =   -1  'True
            CustomFormat    =   "yyyy/M/d"
            Format          =   138870787
            CurrentDate     =   37140
         End
         Begin ImpulseButton.ISButton ISButton6 
            Height          =   315
            Left            =   90
            TabIndex        =   452
            Top             =   8460
            Visible         =   0   'False
            Width           =   1560
            _ExtentX        =   2752
            _ExtentY        =   556
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
            ButtonImage     =   "FrmVizitScreen.frx":5331E
            ColorButton     =   14871017
            DrawFocusRectangle=   0   'False
         End
         Begin MSComCtl2.DTPicker DTPickerAccFrom 
            Height          =   345
            Left            =   3120
            TabIndex        =   453
            TabStop         =   0   'False
            Top             =   7860
            Visible         =   0   'False
            Width           =   1680
            _ExtentX        =   2963
            _ExtentY        =   609
            _Version        =   393216
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   178
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            CalendarBackColor=   12648447
            CalendarTitleBackColor=   10383715
            CheckBox        =   -1  'True
            CustomFormat    =   "yyyy/M/d"
            Format          =   138870787
            CurrentDate     =   37140
         End
         Begin MSComCtl2.DTPicker DTPickerAccTo 
            Height          =   345
            Left            =   1470
            TabIndex        =   454
            TabStop         =   0   'False
            Top             =   8010
            Visible         =   0   'False
            Width           =   1680
            _ExtentX        =   2963
            _ExtentY        =   609
            _Version        =   393216
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   178
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            CalendarBackColor=   12648447
            CalendarTitleBackColor=   10383715
            CheckBox        =   -1  'True
            CustomFormat    =   "yyyy/M/d"
            Format          =   138870787
            CurrentDate     =   37140
         End
         Begin VSFlex8Ctl.VSFlexGrid grdAging 
            Height          =   3630
            Left            =   12150
            TabIndex        =   436
            Top             =   7680
            Visible         =   0   'False
            Width           =   2670
            _cx             =   4710
            _cy             =   6403
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
            Cols            =   23
            FixedRows       =   1
            FixedCols       =   1
            RowHeightMin    =   320
            RowHeightMax    =   0
            ColWidthMin     =   0
            ColWidthMax     =   0
            ExtendLastCol   =   -1  'True
            FormatString    =   $"FrmVizitScreen.frx":536B8
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
         Begin VSFlex8Ctl.VSFlexGrid grdAging2 
            Height          =   1470
            Left            =   7410
            TabIndex        =   464
            Top             =   7020
            Visible         =   0   'False
            Width           =   10050
            _cx             =   17727
            _cy             =   2593
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
            Cols            =   20
            FixedRows       =   1
            FixedCols       =   1
            RowHeightMin    =   320
            RowHeightMax    =   0
            ColWidthMin     =   0
            ColWidthMax     =   0
            ExtendLastCol   =   -1  'True
            FormatString    =   $"FrmVizitScreen.frx":53A38
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
         Begin ImpulseButton.ISButton btn_Cancel 
            Height          =   345
            Index           =   0
            Left            =   420
            TabIndex        =   469
            Top             =   6030
            Width           =   975
            _ExtentX        =   1720
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
            ButtonImage     =   "FrmVizitScreen.frx":53D40
            ColorButton     =   14871017
            DrawFocusRectangle=   0   'False
            DisabledImageStyle=   1
         End
         Begin MSDataListLib.DataCombo dcClass 
            Height          =   315
            Left            =   5100
            TabIndex        =   470
            Top             =   2130
            Width           =   7545
            _ExtentX        =   13309
            _ExtentY        =   556
            _Version        =   393216
            BackColor       =   16777215
            Text            =   ""
            RightToLeft     =   -1  'True
         End
         Begin MSDataListLib.DataCombo DcCustomerType 
            Height          =   315
            Left            =   5100
            TabIndex        =   472
            Tag             =   "⁄›Ê« Ì—ÃÏ «œŒ«· √”„ «·ÕÌ"
            Top             =   1620
            Width           =   7545
            _ExtentX        =   13309
            _ExtentY        =   556
            _Version        =   393216
            Text            =   ""
            RightToLeft     =   -1  'True
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "‰Ê⁄ «·⁄„Ì·"
            Height          =   285
            Index           =   21
            Left            =   12615
            RightToLeft     =   -1  'True
            TabIndex        =   473
            Top             =   1650
            Width           =   3750
         End
         Begin VB.Label Label4 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "«· ’‰Ì›"
            ForeColor       =   &H00000000&
            Height          =   375
            Index           =   1
            Left            =   13815
            TabIndex        =   471
            Top             =   2070
            Width           =   2550
         End
         Begin VB.Label Label6 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "«Ã„«·Ì «·„ »ﬁÌ"
            Height          =   375
            Index           =   4
            Left            =   10110
            RightToLeft     =   -1  'True
            TabIndex        =   468
            Top             =   7860
            Visible         =   0   'False
            Width           =   1215
         End
         Begin VB.Label Label5 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   " «—ÌŒ «·ﬁÌ«”"
            Height          =   375
            Index           =   3
            Left            =   1560
            RightToLeft     =   -1  'True
            TabIndex        =   463
            Top             =   720
            Width           =   1350
         End
         Begin VB.Label Label6 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "›—⁄ „⁄Ì‰"
            Height          =   375
            Index           =   3
            Left            =   13320
            RightToLeft     =   -1  'True
            TabIndex        =   462
            Top             =   2670
            Width           =   3075
         End
         Begin VB.Label Label7 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   " «—ÌŒ «·«” Õﬁ«ﬁ „‰"
            Height          =   375
            Left            =   5805
            RightToLeft     =   -1  'True
            TabIndex        =   461
            Top             =   7860
            Visible         =   0   'False
            Width           =   1350
         End
         Begin VB.Label Label8 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "«·Ï  «—ÌŒ"
            Height          =   375
            Left            =   1410
            RightToLeft     =   -1  'True
            TabIndex        =   460
            Top             =   7500
            Visible         =   0   'False
            Width           =   1350
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00C0FFFF&
            Caption         =   "Â–Â «·‘«‘…  ﬁÊ„ »«ŸÂ«— »Ì«‰«  «⁄„«— «·œÌÊ‰ ÿ»ﬁ« · «—ÌŒ «·ﬁÌ«”"
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
            Height          =   1380
            Index           =   31
            Left            =   10080
            RightToLeft     =   -1  'True
            TabIndex        =   459
            Top             =   6480
            Visible         =   0   'False
            Width           =   6975
         End
         Begin VB.Shape Shape1 
            BorderWidth     =   2
            Height          =   1455
            Left            =   9780
            Top             =   6000
            Visible         =   0   'False
            Width           =   6975
         End
         Begin VB.Label Label24 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "«·Ï  «—ÌŒ"
            Height          =   375
            Left            =   1410
            RightToLeft     =   -1  'True
            TabIndex        =   458
            Top             =   7860
            Visible         =   0   'False
            Width           =   1350
         End
         Begin VB.Label Label25 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   " «—ÌŒ «’œ«— «·›« Ê—… „‰"
            Height          =   375
            Left            =   5475
            RightToLeft     =   -1  'True
            TabIndex        =   457
            Top             =   7500
            Visible         =   0   'False
            Width           =   1710
         End
         Begin VB.Label Label46 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "„‰"
            Height          =   375
            Left            =   5940
            RightToLeft     =   -1  'True
            TabIndex        =   456
            Top             =   7950
            Visible         =   0   'False
            Width           =   390
         End
         Begin VB.Label Label47 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "«·Ì"
            Height          =   375
            Left            =   4830
            RightToLeft     =   -1  'True
            TabIndex        =   455
            Top             =   8220
            Visible         =   0   'False
            Width           =   390
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
Dim cSearchDCombo As clsDCboSearch
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

Public mToStoreID As Long
Dim mItemIDS As String
Dim mDateTrans As Date

Private Sub BtnPrint_Click(Index As Integer)
    
   
   print_report66 Index
    'If Me.Option1.value = True Then
     
    'Else
    '    print_report2 Index
    'End If
End Sub

Private Sub PrintAging(ByVal Ind As Long)
    
    Dim MySQL As String
    Dim RsData As New ADODB.Recordset
    Dim xApp As New CRAXDRT.Application
    Dim xReport As CRAXDRT.Report
    Dim CViewer As ClsReportViewer
    Dim StrReportTitle As String
    Dim StrFileName As String
    Dim Msg As String
        
        
    
    MySQL = "Select * from TblAging  Order By AGEID"
   
 
    If Ind = 0 Then
        MySQL = "SELECT * FROM TblAging WHERE ISNULL(StillAmount,0) <> 0"
        If SystemOptions.UserInterface = ArabicInterface Then
            StrFileName = App.path & "\REPORTS\REPORTS NEW\" & "Aging1.rpt"
        Else
            StrFileName = App.path & "\REPORTS\REPORTS NEW\" & "Aging1E.rpt"
        End If
    Else
        MySQL = " SELECT TblAging.* FROM TblAging INNER JOIN Ageng_type ON Ageng_type.id = TblAging.AGEID"
        MySQL = MySQL & " ORDER BY   Ageng_type.id,TblAging.Account_Code"
        If SystemOptions.UserInterface = ArabicInterface Then
            StrFileName = App.path & "\REPORTS\REPORTS NEW\" & "Aging2.rpt"
        Else
            StrFileName = App.path & "\REPORTS\REPORTS NEW\" & "Aging2E.rpt"
        End If
    End If
    
    If Dir(StrFileName) = "" Then
        'GetMsgs 139, vbExclamation
        Screen.MousePointer = vbDefault
        Exit Sub
    End If

    Set RsData = New ADODB.Recordset
    RsData.Open MySQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

    If RsData.BOF Or RsData.EOF Then
        If SystemOptions.UserInterface = ArabicInterface Then
            Msg = "·« ÊÃœ »Ì«‰«  ··⁄—÷"
        Else
            Msg = "No data"
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

End Sub
Function print_report66(Optional Ind As Integer = 0)
    Dim MySQL As String
    Dim RsData As New ADODB.Recordset
    Dim xApp As New CRAXDRT.Application
    Dim xReport As CRAXDRT.Report
    Dim CViewer As ClsReportViewer
    Dim StrReportTitle As String
    Dim StrFileName As String
    Dim Msg As String
       Dim mSql1 As String
    Dim mSql2 As String

    Dim x As Integer
    MySQL = "Select * from TblAging  "
     MySQL = MySQL & " WHERE 1 = 1  "
    If Not IsNull(DTP_Date.value) Then
        MySQL = MySQL & " and TblAging.DueDate <=" & SQLDate(DTP_Date.value, True) & ""
    End If
       If ChekCustomer.value = vbChecked Then
        If val(DBCboClientName.BoundText) <> 0 Then
            MySQL = MySQL & " and TblAging.CusID =" & val(DBCboClientName.BoundText) & ""
        End If
    End If
    'StrCusID = ""
    If CheckAllCustomer.value = vbChecked Then
         If StrCusID.Text <> "" Then
            MySQL = MySQL & " and TblAging.CusID in (" & (StrCusID.Text) & ")"
        End If
    End If
     MySQL = MySQL & " ORDER BY DueDate  "
    RsData.Open MySQL, Cn, adOpenKeyset, adLockReadOnly
    If Not RsData.EOF Then
          If SystemOptions.UserInterface = ArabicInterface Then
                        x = MsgBox("ÌÊÃœ ⁄„— œÌ‰  „ ⁄„·Â „”»ﬁ« » «—ÌŒ " & DTP_Date.value & "" & " Â·  Êœ ⁄—÷Â ‰⁄„/·«", vbInformation + vbYesNo)
                    Else
                        x = MsgBox("No Contract For This Employee Create Contarct y / n", vbInformation + vbYesNo)
                    End If
     
                    If x = vbYes Then
                        loadgrid MySQL, grdAging, True, False
                        PrintAging Ind
                        Exit Function
                    End If
    End If

    RsData.Close
 
    

    
    grdAging2.Rows = 1
    grdAging.Rows = 1
    
Dim mWhereCus As String


'
'
'   MySQL = ""
'  MySQL = MySQL & " update"
' MySQL = MySQL & " Accounts"
' MySQL = MySQL & " SET BalanceAging ="
'
'MySQL = MySQL & " (SELECT SUM(XB.TransNet)"
'
'MySQL = MySQL & " FROM   ("
'MySQL = MySQL & "            SELECT dev.Account_Code,"
'MySQL = MySQL & "                   dev.Credit_Or_Debit,"
'MySQL = MySQL & "                   dev.branch_id,"
'MySQL = MySQL & "                   dev.Notes_ID,"
'MySQL = MySQL & "                   NotesTypeName,"
'MySQL = MySQL & "                   ISNULL(Dev.DueDate, Notes.NoteDate) DueDate,"
'MySQL = MySQL & "                   Notes.NoteDate,"
'MySQL = MySQL & "                   Notes.NoteType,"
'MySQL = MySQL & "                   Notes.NoteSerial,"
'MySQL = MySQL & "                   dev.[Value]             AS Note_Value,"
'MySQL = MySQL & "                   a.Account_Name          AS CusName,"
'MySQL = MySQL & "                   ISNULL(dev.[Value], 0)  AS TransNet "
'
'MySQL = MySQL & " ,  dbo.GetDeptAgeID(DATEDIFF(day,ISNULL(dev.DueDate, Notes.NoteDate),"
'MySQL = MySQL & SQLDate(DTP_Date.value, True) & ")) AS AgeID,   Datediff(day,ISNULL(dev.DueDate, Notes.NoteDate), " & SQLDate(DTP_Date.value, True) & " ) AS DiffDate"
'
'
'MySQL = MySQL & "            FROM   DOUBLE_ENTREY_VOUCHERS  AS dev"
'MySQL = MySQL & "                   INNER JOIN Notes"
'MySQL = MySQL & "                        ON  Notes.NoteId = dev.Notes_Id"
'MySQL = MySQL & "                   LEFT OUTER JOIN TblNotesTypes"
'MySQL = MySQL & "                        ON  Notes.NoteType = TblNotesTypes.NotesType"
'MySQL = MySQL & "                   LEFT OUTER JOIN ACCOUNTS AS a"
'MySQL = MySQL & "                        ON  a.Account_Code = dev.Account_Code"
'MySQL = MySQL & "            Where (dev.Posted Is Null)"
'MySQL = MySQL & "                   AND ISNULL(dev.[Value], 0) <> 0"
'MySQL = MySQL & "                   AND dev.Credit_Or_Debit = 1"
'MySQL = MySQL & "            Union all"
'MySQL = MySQL & "            SELECT dev.Account_Code,"
'MySQL = MySQL & "                   dev.Credit_Or_Debit,"
'MySQL = MySQL & "                   dev.branch_id,"
'MySQL = MySQL & "                   dev.Notes_ID,"
'MySQL = MySQL & "                   NotesTypeName = 'ﬁÌœ «›  «ÕÌ',"
'MySQL = MySQL & "                   ISNULL(Dev.DueDate, Notes1.NoteDate) DueDate,"
'MySQL = MySQL & "                   Notes1.NoteDate,"
'MySQL = MySQL & "                   Notes1.NoteType,"
'MySQL = MySQL & "                   Notes1.NoteSerial,"
'MySQL = MySQL & "                   dev.[Value]             AS Note_Value,"
'MySQL = MySQL & "                   a.Account_Name          AS CusName,"
'MySQL = MySQL & "                   ISNULL(dev.[Value], 0)  AS TransNet "
'
'MySQL = MySQL & " ,  dbo.GetDeptAgeID(DATEDIFF(day,ISNULL(dev.DueDate, Notes1.NoteDate),"
'MySQL = MySQL & SQLDate(DTP_Date.value, True) & ")) AS AgeID,   Datediff(day,ISNULL(dev.DueDate, Notes1.NoteDate), " & SQLDate(DTP_Date.value, True) & " ) AS DiffDate"
'
'
'MySQL = MySQL & "             FROM   DOUBLE_ENTREY_VOUCHERS1 AS dev"
'MySQL = MySQL & "                    INNER JOIN Notes1"
'MySQL = MySQL & "                        ON  Notes1.NoteId = dev.Notes_Id"
'MySQL = MySQL & "                  LEFT OUTER JOIN TblNotesTypes"
'MySQL = MySQL & "                       ON  Notes1.NoteType = TblNotesTypes.NotesType"
'MySQL = MySQL & "                  LEFT OUTER JOIN ACCOUNTS AS a"
'MySQL = MySQL & "                       ON  a.Account_Code = dev.Account_Code"
'MySQL = MySQL & "           Where (dev.Posted Is Null)"
'MySQL = MySQL & "                  AND ISNULL(dev.[Value], 0) <> 0"
'MySQL = MySQL & "                  AND dev.Credit_Or_Debit = 1"
'MySQL = MySQL & "       ) XB"
'MySQL = MySQL & "       LEFT OUTER JOIN dbo.Ageng_type"
'MySQL = MySQL & "            ON  XB.AgeID = dbo.Ageng_type.id"
'MySQL = MySQL & " Where 1 = 1"
''
'If Not IsNull(DTP_Date.value) Then
'    MySQL = MySQL & " and XB.DueDate <=" & SQLDate(DTP_Date.value, True) & ""
'End If
'    If Option1(2).value = True Then
'        mWhereCus = " and Account_Code  In (Select Account_Code from TblCustemers Where  Type = " & mCusType & ")"
'
'    ElseIf Option2.value = True Then
'        mWhereCus = " and Account_Code  In (Select Account_Code from TblCustemers Where  Type = " & mCusType & ")"
'
'    End If
'
'
'     If ChekCustomer.value = vbChecked Then
'        If val(DBCboClientName.BoundText) <> 0 Then
'            MySQL = MySQL & " and Account_Code  In (Select Account_Code from TblCustemers Where  TblCustemers.CusId = " & val(DBCboClientName.BoundText) & " )"
'            mWhereCus = " and Account_Code  In (Select Account_Code from TblCustemers Where  TblCustemers.CusId = " & val(DBCboClientName.BoundText) & " )"
'        End If
'    Else
'         If val(DBCboClientName.BoundText) <> 0 Then
'            mWhereCus = " and Account_Code  In (Select Account_Code from TblCustemers Where  TblCustemers.CusId = " & val(DBCboClientName.BoundText) & " )"
'
'        End If
'    End If
'    MySQL = MySQL & mWhereCus
'
'
'
'
'   ' End If
'    'Account_Code  In (Select Account_Code from TblCustemers Where  TblCustemers.CusId = " & val(DBCboClientName.BoundText) & ")"
'
'        If CheckEmp.value = vbChecked Then
'        If val(DcbEmployee.BoundText) <> 0 Then
'            MySQL = MySQL & " and Account_Code  "
'            MySQL = MySQL & " In (Select Account_Code from TblCustemers Where  TblCustemers.EmpID = " & val(DcbEmployee.BoundText) & ")"
'
'
'        End If
'    End If
'
'
'    MySQL = MySQL & " AND Credit_Or_Debit = 1"
'
'
'
'MySQL = MySQL & "       AND ISNULL(TransNet, 0) <> 0  AND ACCOUNTS.Account_Code = xb.Account_Code"
'MySQL = MySQL & " GROUP BY XB.Account_Code,CusName"
'
'MySQL = MySQL & " ) WHERE 1 = 1 " & mWhereCus
''MySQL = MySQL & " ,AgeID"
'
'
'
'   mSql2 = MySQL
'
'Cn.Execute mSql2

'-------------------------------
   
  Dim mCusType  As Integer
    If Option1(2).value = True Then
        mCusType = 1
    Else
        mCusType = 2
    End If
   
MySQL = ""
MySQL = MySQL & " SELECT "

If mCusType = 2 Then
MySQL = MySQL & " SUM (TransNet) AS TransNet,XB.Account_Code"

Else
MySQL = MySQL & "        XB.Account_Code,"
MySQL = MySQL & " Xb.NotesTypeName ,"
MySQL = MySQL & "        XB.DueDate,"
MySQL = MySQL & "        DiffDate,"
MySQL = MySQL & "        XB.NoteDate,"
MySQL = MySQL & "        XB.NoteType,"
MySQL = MySQL & "        XB.Note_Value TransNet,"
MySQL = MySQL & "        XB.CusName,"
',BalanceAging,"
       '--XB.CusID,
MySQL = MySQL & "        xb.AgeID,"
MySQL = MySQL & "        XB.NoteSerial,"
MySQL = MySQL & "        dbo.Ageng_type.Name,"
MySQL = MySQL & "        dbo.Ageng_type.[From],"
MySQL = MySQL & "        dbo.Ageng_type.[To],"
MySQL = MySQL & "        dbo.Ageng_type.Color,"
MySQL = MySQL & "        dbo.Ageng_type.NameE,"

'       --    BranchId,


If SystemOptions.UserInterface = ArabicInterface Then
    
    MySQL = MySQL & "        ISNULL(NotesTypeName, 'ﬁÌœ «›  «ÕÏ') AS TransactionTypeName"
Else
    
    MySQL = MySQL & "        ISNULL(NotesTypeName, 'Opening entry') AS TransactionTypeName"
End If
End If

MySQL = MySQL & " FROM   ("
MySQL = MySQL & "            SELECT dev.Account_Code,"
MySQL = MySQL & "                   dev.Credit_Or_Debit,"
MySQL = MySQL & "                   dev.branch_id,"
MySQL = MySQL & "                   dev.Notes_ID,"
If SystemOptions.UserInterface = ArabicInterface Then
    MySQL = MySQL & "                   NotesTypeName,"
Else
    MySQL = MySQL & "                   NotesTypeNamee as NotesTypeName,"
End If

MySQL = MySQL & "                   ISNULL(Dev.DueDate, Notes.NoteDate) DueDate,"
MySQL = MySQL & "                   Notes.NoteDate,"
MySQL = MySQL & "                   Notes.NoteType,"
MySQL = MySQL & "                   Notes.NoteSerial,"
MySQL = MySQL & "                   dev.[Value]             AS Note_Value,"

If SystemOptions.UserInterface = ArabicInterface Then
    MySQL = MySQL & "                   a.Account_Name          AS CusName,"
Else
    MySQL = MySQL & "                   a.Account_NameEng          AS CusName,"
End If


'a.BalanceAging,"
MySQL = MySQL & "                   ISNULL(dev.[Value], 0)  AS TransNet "

MySQL = MySQL & " ,  dbo.GetDeptAgeID(DATEDIFF(day,ISNULL(dev.DueDate, Notes.NoteDate),"
MySQL = MySQL & SQLDate(DTP_Date.value, True) & ")) AS AgeID,   Datediff(day,ISNULL(dev.DueDate, Notes.NoteDate), " & SQLDate(DTP_Date.value, True) & " ) AS DiffDate"


MySQL = MySQL & "            FROM   DOUBLE_ENTREY_VOUCHERS  AS dev"
MySQL = MySQL & "                   INNER JOIN Notes"
MySQL = MySQL & "                        ON  Notes.NoteId = dev.Notes_Id"
MySQL = MySQL & "                   LEFT OUTER JOIN TblNotesTypes"
MySQL = MySQL & "                        ON  Notes.NoteType = TblNotesTypes.NotesType"
MySQL = MySQL & "                   LEFT OUTER JOIN ACCOUNTS AS a"
MySQL = MySQL & "                        ON  a.Account_Code = dev.Account_Code"
MySQL = MySQL & "            Where (dev.Posted Is Null)"
MySQL = MySQL & "                   AND ISNULL(dev.[Value], 0) <> 0"
MySQL = MySQL & "                   AND dev.Credit_Or_Debit = 0"
MySQL = MySQL & "            Union all"
MySQL = MySQL & "            SELECT dev.Account_Code,"
MySQL = MySQL & "                   dev.Credit_Or_Debit,"
MySQL = MySQL & "                   dev.branch_id,"
MySQL = MySQL & "                   dev.Notes_ID,"

If SystemOptions.UserInterface = ArabicInterface Then
    MySQL = MySQL & "                   NotesTypeName = 'ﬁÌœ «›  «ÕÌ',"
Else
    MySQL = MySQL & "                   NotesTypeName = 'Opening entry',"
End If
MySQL = MySQL & "                   ISNULL(Dev.DueDate, Notes1.NoteDate) DueDate,"
MySQL = MySQL & "                   Notes1.NoteDate,"
MySQL = MySQL & "                   Notes1.NoteType,"
MySQL = MySQL & "                   Notes1.NoteSerial,"
MySQL = MySQL & "                   dev.[Value]             AS Note_Value,"


If SystemOptions.UserInterface = ArabicInterface Then
    MySQL = MySQL & "                   a.Account_Name          AS CusName,"
Else
    MySQL = MySQL & "                   a.Account_NameEng          AS CusName,"
End If


'a.BalanceAging,"
MySQL = MySQL & "                   ISNULL(dev.[Value], 0)  AS TransNet "

MySQL = MySQL & " ,  dbo.GetDeptAgeID(DATEDIFF(day,ISNULL(dev.DueDate, Notes1.NoteDate),"
MySQL = MySQL & SQLDate(DTP_Date.value, True) & ")) AS AgeID,   Datediff(day,ISNULL(dev.DueDate, Notes1.NoteDate), " & SQLDate(DTP_Date.value, True) & " ) AS DiffDate"

    
MySQL = MySQL & "             FROM   DOUBLE_ENTREY_VOUCHERS1 AS dev"
MySQL = MySQL & "                    INNER JOIN Notes1"
MySQL = MySQL & "                        ON  Notes1.NoteId = dev.Notes_Id"
MySQL = MySQL & "                  LEFT OUTER JOIN TblNotesTypes"
MySQL = MySQL & "                       ON  Notes1.NoteType = TblNotesTypes.NotesType"
MySQL = MySQL & "                  LEFT OUTER JOIN ACCOUNTS AS a"
MySQL = MySQL & "                       ON  a.Account_Code = dev.Account_Code"
MySQL = MySQL & "           Where (dev.Posted Is Null)"
MySQL = MySQL & "                  AND ISNULL(dev.[Value], 0) <> 0"
MySQL = MySQL & "                  AND dev.Credit_Or_Debit = 0"
MySQL = MySQL & "       ) XB"
MySQL = MySQL & "       Right OUTER JOIN dbo.Ageng_type"
MySQL = MySQL & "            ON  XB.AgeID = dbo.Ageng_type.id"
MySQL = MySQL & " Where 1 = 1"
'
If Not IsNull(DTP_Date.value) Then
    MySQL = MySQL & " and XB.DueDate <=" & SQLDate(DTP_Date.value, True) & ""
End If

    If Option1(2).value = True Then
        mCusType = 1
        MySQL = MySQL & " and Account_Code  In (Select Account_Code from TblCustemers Where  Type = " & mCusType & ")"
        
    ElseIf Option2.value = True Then
        mCusType = 2
        MySQL = MySQL & " and Account_Code  In (Select Account_Code from TblCustemers Where  Type = " & mCusType & ")"
        
    End If
    If ChekCustomer.value = vbChecked Then
        If val(DBCboClientName.BoundText) <> 0 Then
            MySQL = MySQL & " and Account_Code  In (Select Account_Code from TblCustemers Where  TblCustemers.CusId = " & val(DBCboClientName.BoundText) & " and Type = " & mCusType & ")"
        Else
            
        End If

    Else
         If val(DBCboClientName.BoundText) <> 0 Then
            MySQL = MySQL & " and Account_Code  In (Select Account_Code from TblCustemers Where  TblCustemers.CusId = " & val(DBCboClientName.BoundText) & " )"
        
        End If
        
    End If
    
        If dcClass.Text <> "" And val(dcClass.BoundText) <> 0 Then
            
             MySQL = MySQL & " and Account_Code  In (Select Account_Code from TblCustemers Where  TblCustemers.ClassCustomersId = " & val(dcClass.BoundText) & " and Type = " & mCusType & ")"
        End If
        
            If (DcCustomerType.Text) <> "" And val(DcCustomerType.BoundText) <> 0 Then
            
             MySQL = MySQL & " and Account_Code  In (Select Account_Code from TblCustemers Where  TblCustemers.CustomerTypeID = " & val(DcCustomerType.BoundText) & " and Type = " & mCusType & ")"
        End If
        
    If CheckAllCustomer.value = vbChecked Then
         If StrCusID.Text <> "" Then
           ' MySQL = MySQL & " and TblCustemers.CusID in (" & (StrCusID.Text) & ")"
            MySQL = MySQL & " and Account_Code  In ( Select  TblCustemers.Account_Code from TblCustemers Where TblCustemers.CusID in (" & (StrCusID.Text) & ") )"
        End If
    Else
        If StrCusID.Text <> "" Then
           ' MySQL = MySQL & " and TblCustemers.CusID in (" & (StrCusID.Text) & ")"
            MySQL = MySQL & " and Account_Code  In ( Select  TblCustemers.Account_Code from TblCustemers Where TblCustemers.CusID in (" & (StrCusID.Text) & ") )"
        End If
    End If
     
  
  
   ' End If
    'Account_Code  In (Select Account_Code from TblCustemers Where  TblCustemers.CusId = " & val(DBCboClientName.BoundText) & ")"
   
        If CheckEmp.value = vbChecked Then
        If val(DcbEmployee.BoundText) <> 0 Then
            MySQL = MySQL & " and Account_Code  "
            MySQL = MySQL & " In (Select Account_Code from TblCustemers Where  TblCustemers.EmpID = " & val(DcbEmployee.BoundText) & ")"
            
            
        End If
    End If
           
If mCusType = 1 Then
MySQL = MySQL & " Order By"
MySQL = MySQL & "       Account_Code,"
MySQL = MySQL & "       XB.NoteSerial,"
MySQL = MySQL & "       Xb.DueDate"
'
Else
MySQL = MySQL & " GROUP BY Account_Code   "
End If
   
    
    
    mSql1 = MySQL
    
  
  
   
   MySQL = ""
   
If mCusType = 2 Then
    MySQL = MySQL & " SELECT SUM(XB.TransNet) AS TransNet ,SUM(XB.TransNet) AS TransNet22,Account_Code,CusName"


Else
      MySQL = MySQL & " Select  SUM (TransNet) AS TransNet,Account_Code,CusName"

End If
   
   


MySQL = ""
MySQL = MySQL & " SELECT "

If mCusType = 1 Then
MySQL = MySQL & "   SUM (TransNet) AS TransNet,Account_Code,CusName"

Else
MySQL = MySQL & "        XB.Account_Code,"
MySQL = MySQL & " Xb.NotesTypeName ,"
MySQL = MySQL & "        XB.DueDate,"
MySQL = MySQL & "        DiffDate,"
MySQL = MySQL & "        XB.NoteDate,"
MySQL = MySQL & "        XB.NoteType,"
MySQL = MySQL & "        XB.Note_Value TransNet,"
MySQL = MySQL & "        XB.CusName,"
',BalanceAging,"
       '--XB.CusID,
MySQL = MySQL & "        xb.AgeID,"
MySQL = MySQL & "        XB.NoteSerial,"
MySQL = MySQL & "        dbo.Ageng_type.Name,"
MySQL = MySQL & "        dbo.Ageng_type.[From],"
MySQL = MySQL & "        dbo.Ageng_type.[To],"
MySQL = MySQL & "        dbo.Ageng_type.Color,"
MySQL = MySQL & "        dbo.Ageng_type.NameE,"

'       --    BranchId,

If SystemOptions.UserInterface = ArabicInterface Then
    
    MySQL = MySQL & "        ISNULL(NotesTypeName, 'ﬁÌœ «›  «ÕÏ') AS TransactionTypeName"
Else
    
    MySQL = MySQL & "        ISNULL(NotesTypeName, 'Opening entry') AS TransactionTypeName"
End If

End If

',AgeID"




MySQL = MySQL & " FROM   ("
MySQL = MySQL & "            SELECT dev.Account_Code,"
MySQL = MySQL & "                   dev.Credit_Or_Debit,"
MySQL = MySQL & "                   dev.branch_id,"
MySQL = MySQL & "                   dev.Notes_ID,"
If SystemOptions.UserInterface = ArabicInterface Then
    MySQL = MySQL & "                   NotesTypeName,"
Else
    MySQL = MySQL & "                   NotesTypeNamee as NotesTypeName,"
End If
MySQL = MySQL & "                   ISNULL(Dev.DueDate, Notes.NoteDate) DueDate,"
MySQL = MySQL & "                   Notes.NoteDate,"
MySQL = MySQL & "                   Notes.NoteType,"
MySQL = MySQL & "                   Notes.NoteSerial,"
MySQL = MySQL & "                   dev.[Value]             AS Note_Value,"
If SystemOptions.UserInterface = ArabicInterface Then
    MySQL = MySQL & "                   a.Account_Name          AS CusName,"
Else
    MySQL = MySQL & "                   a.Account_NameEng          AS CusName,"
End If


MySQL = MySQL & "                   ISNULL(dev.[Value], 0)  AS TransNet "

MySQL = MySQL & " ,  dbo.GetDeptAgeID(DATEDIFF(day,ISNULL(dev.DueDate, Notes.NoteDate),"
MySQL = MySQL & SQLDate(DTP_Date.value, True) & ")) AS AgeID,   Datediff(day,ISNULL(dev.DueDate, Notes.NoteDate), " & SQLDate(DTP_Date.value, True) & " ) AS DiffDate"


MySQL = MySQL & "            FROM   DOUBLE_ENTREY_VOUCHERS  AS dev"
MySQL = MySQL & "                   INNER JOIN Notes"
MySQL = MySQL & "                        ON  Notes.NoteId = dev.Notes_Id"
MySQL = MySQL & "                   LEFT OUTER JOIN TblNotesTypes"
MySQL = MySQL & "                        ON  Notes.NoteType = TblNotesTypes.NotesType"
MySQL = MySQL & "                   LEFT OUTER JOIN ACCOUNTS AS a"
MySQL = MySQL & "                        ON  a.Account_Code = dev.Account_Code"
MySQL = MySQL & "            Where (dev.Posted Is Null)"
MySQL = MySQL & "                   AND ISNULL(dev.[Value], 0) <> 0"
MySQL = MySQL & "                   AND dev.Credit_Or_Debit = 1"
MySQL = MySQL & "            Union all"
MySQL = MySQL & "            SELECT dev.Account_Code,"
MySQL = MySQL & "                   dev.Credit_Or_Debit,"
MySQL = MySQL & "                   dev.branch_id,"
MySQL = MySQL & "                   dev.Notes_ID,"
If SystemOptions.UserInterface = ArabicInterface Then
    MySQL = MySQL & "                   NotesTypeName = 'ﬁÌœ «›  «ÕÌ',"
Else
    MySQL = MySQL & "                   NotesTypeName = 'Opening entry',"
End If
MySQL = MySQL & "                   ISNULL(Dev.DueDate, Notes1.NoteDate) DueDate,"
MySQL = MySQL & "                   Notes1.NoteDate,"
MySQL = MySQL & "                   Notes1.NoteType,"
MySQL = MySQL & "                   Notes1.NoteSerial,"
MySQL = MySQL & "                   dev.[Value]             AS Note_Value,"



If SystemOptions.UserInterface = ArabicInterface Then
    MySQL = MySQL & "                   a.Account_Name          AS CusName,"
Else
    MySQL = MySQL & "                   a.Account_NameEng          AS CusName,"
End If


MySQL = MySQL & "                   ISNULL(dev.[Value], 0)  AS TransNet "

MySQL = MySQL & " ,  dbo.GetDeptAgeID(DATEDIFF(day,ISNULL(dev.DueDate, Notes1.NoteDate),"
MySQL = MySQL & SQLDate(DTP_Date.value, True) & ")) AS AgeID,   Datediff(day,ISNULL(dev.DueDate, Notes1.NoteDate), " & SQLDate(DTP_Date.value, True) & " ) AS DiffDate"

    
MySQL = MySQL & "             FROM   DOUBLE_ENTREY_VOUCHERS1 AS dev"
MySQL = MySQL & "                    INNER JOIN Notes1"
MySQL = MySQL & "                        ON  Notes1.NoteId = dev.Notes_Id"
MySQL = MySQL & "                  LEFT OUTER JOIN TblNotesTypes"
MySQL = MySQL & "                       ON  Notes1.NoteType = TblNotesTypes.NotesType"
MySQL = MySQL & "                  LEFT OUTER JOIN ACCOUNTS AS a"
MySQL = MySQL & "                       ON  a.Account_Code = dev.Account_Code"
MySQL = MySQL & "           Where (dev.Posted Is Null)"
MySQL = MySQL & "                  AND ISNULL(dev.[Value], 0) <> 0"
MySQL = MySQL & "                  AND dev.Credit_Or_Debit = 1"
MySQL = MySQL & "       ) XB"
MySQL = MySQL & "       Right OUTER JOIN dbo.Ageng_type"
MySQL = MySQL & "            ON  XB.AgeID = dbo.Ageng_type.id"
MySQL = MySQL & " Where 1 = 1"
'
If Not IsNull(DTP_Date.value) Then
    MySQL = MySQL & " and XB.DueDate <=" & SQLDate(DTP_Date.value, True) & ""
End If
    If Option1(2).value = True Then
        MySQL = MySQL & " and Account_Code  In (Select Account_Code from TblCustemers Where  Type = " & mCusType & ")"
    ElseIf Option2.value = True Then
        MySQL = MySQL & " and Account_Code  In (Select Account_Code from TblCustemers Where  Type = " & mCusType & ")"
    End If
    

     If ChekCustomer.value = vbChecked Then
        If val(DBCboClientName.BoundText) <> 0 Then
            MySQL = MySQL & " and Account_Code  In (Select Account_Code from TblCustemers Where  TblCustemers.CusId = " & val(DBCboClientName.BoundText) & " )"
        
        End If
    Else
         If val(DBCboClientName.BoundText) <> 0 Then
            MySQL = MySQL & " and Account_Code  In (Select Account_Code from TblCustemers Where  TblCustemers.CusId = " & val(DBCboClientName.BoundText) & " )"
        
        End If
    End If
    
            If Trim(dcClass.Text) <> "" And val(dcClass.BoundText) <> 0 Then
            
             MySQL = MySQL & " and Account_Code  In (Select Account_Code from TblCustemers Where  TblCustemers.ClassCustomersId = " & val(dcClass.BoundText) & " and Type = " & mCusType & ")"
        End If
  
            
            If Trim(DcCustomerType.Text) <> "" And val(DcCustomerType.BoundText) <> 0 Then
            
             MySQL = MySQL & " and Account_Code  In (Select Account_Code from TblCustemers Where  TblCustemers.CustomerTypeID = " & val(DcCustomerType.BoundText) & " and Type = " & mCusType & ")"
        End If

     
   ' End If
    'Account_Code  In (Select Account_Code from TblCustemers Where  TblCustemers.CusId = " & val(DBCboClientName.BoundText) & ")"
   
        If CheckEmp.value = vbChecked Then
        If val(DcbEmployee.BoundText) <> 0 Then
            MySQL = MySQL & " and Account_Code  "
            MySQL = MySQL & " In (Select Account_Code from TblCustemers Where  TblCustemers.EmpID = " & val(DcbEmployee.BoundText) & ")"
            
            
        End If
    End If
           

    MySQL = MySQL & " AND Credit_Or_Debit = 1"
    


MySQL = MySQL & "       AND ISNULL(TransNet, 0) <> 0"


If mCusType = 2 Then
MySQL = MySQL & " Order By"
MySQL = MySQL & "       Account_Code,"
MySQL = MySQL & "       XB.NoteSerial,"
MySQL = MySQL & "       Xb.DueDate"
'
Else
MySQL = MySQL & " GROUP BY Account_Code  ,CusName"
End If
   


'MySQL = MySQL & " ,AgeID"

   
   
   mSql2 = MySQL

    If Option1(2).value = True Then
        loadgrid mSql1, grdAging, True, False
        loadgrid mSql2, grdAging2, False, False
    Else
        loadgrid mSql2, grdAging, True, False
        loadgrid mSql1, grdAging2, False, False
    End If
    
    
'
'
'    MySQL = MySQL & "                                          AND (DATEDIFF(DAY, '31-Aug-2020', RptLedger_Sub2.RecordDate) < 0)"
'    MySQL = MySQL & "                            ) XB"
'
'    MySQL = MySQL & "                               LEFT OUTER JOIN dbo.Ageng_type"
'    MySQL = MySQL & "                                    ON  XB.ID = dbo.Ageng_type.id"
'
'    MySQL = MySQL & "                        WHERE  XB.DueDate >= '01-Aug-2020'"
'    MySQL = MySQL & "                               AND XB.DueDate <= '31-Aug-2020'"
'
'    MySQL = MySQL & "                        Order By "
'    MySQL = MySQL & "                               Xb.ID , DueDate "
'
'
   

   Dim i As Long
   Dim mValue As Double
   Dim mCusId As Long
    Dim j As Long
    Dim mValue2 As Double
   Dim mCusId2 As Long
    Dim mPayedValue As Double

Dim mAccount_Code As String
Dim mAccount_Code2 As String
Dim Balance As String

'If grdAging.Rows > 1 Then
'    mAccount_Code = Trim(grdAging.TextMatrix(1, grdAging.ColIndex("Account_Code")))
'    WriteCustomerBalPublic mAccount_Code, Balance, , 0, , , , , FromDate1.value, 1
'    grdAging.TextMatrix(1, grdAging.ColIndex("Balance")) = Balance
'End If

'If grdAging2.Rows > 1 Then
'    Balance = ""
'    mAccount_Code = Trim(grdAging2.TextMatrix(1, grdAging2.ColIndex("Account_Code")))
'    WriteCustomerBalPublic mAccount_Code, Balance, , 1, , , , , FromDate1.value, 1
'    grdAging2.TextMatrix(1, grdAging2.ColIndex("Balance")) = Balance
'End If
txtTotalStill = ""
Dim mJ As Long
mJ = 1
   For i = 1 To grdAging.Rows - 1
   
     
'     If I = 1 Then
'        mAccount_Code = Trim(grdAging.TextMatrix(I, grdAging.ColIndex("Account_Code")))
'        WriteCustomerBalPublic mAccount_Code, Balance, , 0, , , , , FromDate1.value, 1
'        grdAging.TextMatrix(I, grdAging.ColIndex("Balance")) = Balance
'     End If
      mValue = val(grdAging.TextMatrix(i, grdAging.ColIndex("TransNet")))
      mAccount_Code = Trim(grdAging.TextMatrix(i, grdAging.ColIndex("Account_Code")))
      
        If val(grdAging.TextMatrix(i, grdAging.ColIndex("PayedValue"))) <> mValue Then
        
        
        mJ = grdAging2.FindRow(mAccount_Code, grdAging2.FixedRows, grdAging2.ColIndex("Account_Code"), False, True)
        'mJ = grdAging2.FindRow("dsfdsf", grdAging2.FixedRows, grdAging2.ColIndex("Account_Code"), False, True)
       ' For j = mJ To grdAging2.Rows - 1
'
       j = mJ
            If mJ <> -1 Then
                mValue2 = val(grdAging2.TextMatrix(j, grdAging2.ColIndex("TransNet")))
                mAccount_Code2 = Trim(grdAging2.TextMatrix(j, grdAging2.ColIndex("Account_Code")))
                mPayedValue = val(grdAging.TextMatrix(i, grdAging.ColIndex("PayedValue")))
    
    
               If mValue2 <> 0 And mAccount_Code2 = mAccount_Code And mValue <> mPayedValue Then
    
                    If mValue - mPayedValue = mValue2 Then
                        grdAging.TextMatrix(i, grdAging.ColIndex("PayedValue")) = mValue2
                        grdAging2.TextMatrix(j, grdAging2.ColIndex("TransNet")) = 0
                    ElseIf mValue - mPayedValue > mValue2 Then
                        grdAging.TextMatrix(i, grdAging.ColIndex("PayedValue")) = val(grdAging.TextMatrix(i, grdAging.ColIndex("PayedValue"))) + mValue2
                        grdAging2.TextMatrix(j, grdAging2.ColIndex("TransNet")) = 0
                    ElseIf mValue - mPayedValue < mValue2 Then
                        grdAging.TextMatrix(i, grdAging.ColIndex("PayedValue")) = mPayedValue + mValue - mPayedValue
                        grdAging2.TextMatrix(j, grdAging2.ColIndex("TransNet")) = mValue2 - (mValue - mPayedValue)
                        grdAging.TextMatrix(i, grdAging.ColIndex("TransNetGrid2")) = mValue2 - (mValue - mPayedValue)
                        'mValue - grdAging.TextMatrix(i, grdAging.ColIndex("PayedValue")) + mValue2
                    End If
               End If
               grdAging2.TextMatrix(j, grdAging2.ColIndex("StillAmount")) = val(grdAging2.TextMatrix(j, grdAging2.ColIndex("TransNet"))) - val(grdAging2.TextMatrix(j, grdAging2.ColIndex("PayedValue")))
            End If
'
'
'            If mAccount_Code2 <> mAccount_Code Then
'                GoTo ExitFor
'            End If
'        Next
        
      End If
      'mJ = j + 1
ExitFor:
      grdAging.TextMatrix(i, grdAging.ColIndex("StillAmount")) = val(grdAging.TextMatrix(i, grdAging.ColIndex("TransNet"))) - val(grdAging.TextMatrix(i, grdAging.ColIndex("PayedValue")))
      If val(grdAging.TextMatrix(i, grdAging.ColIndex("StillAmount"))) = 0 Then
        grdAging.TextMatrix(i, grdAging.ColIndex("StillAmount")) = ""
        grdAging.RowHidden(i) = True
      End If
      txtTotalStill = val(txtTotalStill) + val(grdAging.TextMatrix(i, grdAging.ColIndex("StillAmount")))
   Next
    s = "Delete TblAging "
    Cn.Execute s
    
    
    

    
    s = "Select * from TblAging  "
    
    
    
    saveGrid s, grdAging, "StillAmount", "Id", "Credit_Or_Debit", 0



    Dim rsDummyT As New ADODB.Recordset
    Dim rsDummyT2 As New ADODB.Recordset
    Dim rs As New ADODB.Recordset
    s = " Select Account_Code,AGEID,CusName from TblAging"
    s = s & " GROUP BY Account_Code,AGEID,CusName"
    Set rsDummyT = New ADODB.Recordset
    rsDummyT.Open s, Cn, adOpenStatic, adLockReadOnly
    
    Do While Not rsDummyT.EOF
        s = "Select * from Ageng_type where Id Not In (Select  AGEID from TblAging Where  Account_Code = N'" & Trim(rsDummyT!Account_code & "") & "' )"
        Set rsDummyT2 = New ADODB.Recordset
        rsDummyT2.Open s, Cn, adOpenStatic, adLockReadOnly
        Do While Not rsDummyT2.EOF
            s = "Select * from TblAging "
            rs.Open s, Cn, adOpenKeyset, adLockOptimistic
            rs.AddNew
            rs!AGEID = rsDummyT2!ID
            rs!Account_code = rsDummyT!Account_code & ""
            rs!CusName = rsDummyT!CusName & ""
            rs!To = rsDummyT2!To & ""
            rs!From = rsDummyT2!From & ""
            If SystemOptions.UserInterface = ArabicInterface Then
                rs!Name = rsDummyT2!Name & ""
            Else
                rs!Name = rsDummyT2!Name & ""
            End If
            rs.update
            rs.Close
            rsDummyT2.MoveNext
        Loop
        's = "Select Account_Code,AGEID from TblAging Where  Account_Code = " & Trim(rsDummyT!Account_code & "")
        
        rsDummyT.MoveNext
    Loop
    
 
    s = "Select * from TblAging "
'    saveGrid s, grdAging2, "CusId", "Id", "Credit_Or_Debit", 1


    Set RsData = New ADODB.Recordset
    RsData.Open s, Cn, adOpenStatic, adLockReadOnly, adCmdText

    If RsData.BOF Or RsData.EOF Then
        If SystemOptions.UserInterface = ArabicInterface Then
            Msg = "·« ÊÃœ »Ì«‰«  ··⁄—÷"
        Else
            Msg = "No data"
        End If
        'MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
        RsData.Close
        Set RsData = Nothing
        Screen.MousePointer = vbDefault
        Exit Function
    End If
StrCusID = ""
    RsData.Close
    Set RsData = Nothing
    PrintAging Ind
    Screen.MousePointer = vbDefault
End Function

Private Sub chkIsVat_Click()
grdExcel.ColHidden(grdExcel.ColIndex("VatValue")) = False
grdExcel.ColHidden(grdExcel.ColIndex("AmountNet")) = False
Dim Percentg As Double
Dim Notevalue As Double, mVat As Double, mValue As Double
If chkIsVat.value = vbUnchecked Then
    grdExcel.TextMatrix(0, grdExcel.ColIndex("AmountNet")) = "«·„»·€ «·«Ã„«·Ì"
Else
    grdExcel.TextMatrix(0, grdExcel.ColIndex("AmountNet")) = "«·„»·€ »œÊ‰ «·÷—Ì»…"
End If

PercentgValueAddedAccount_Transec XPDtbTrans7.value, 21, 1, , Percentg

Dim i As Long
For i = 1 To grdExcel.Rows - 1
    Notevalue = Abs(val(grdExcel.TextMatrix(i, grdExcel.ColIndex("Amount"))))
    If Notevalue <> 0 Then
        If chkIsVat.value = vbChecked Then
            If Percentg = 5 Then
                mValue = Notevalue / 1.05
            ElseIf Percentg = 15 Then
                mValue = Notevalue / 1.15
            End If
            mVat = Notevalue - mValue
            grdExcel.TextMatrix(i, grdExcel.ColIndex("VatValue")) = Round(mVat, 3)
 
        
             grdExcel.TextMatrix(i, grdExcel.ColIndex("AmountNet")) = Round(Notevalue - mVat, 3)
        Else
            If Percentg = 5 Then
                mValue = Notevalue / 1.05
            ElseIf Percentg = 15 Then
                mValue = Notevalue * 1.15
            End If
            mVat = mValue - Notevalue
            grdExcel.TextMatrix(i, grdExcel.ColIndex("VatValue")) = Round(mVat, 3)
        
            grdExcel.TextMatrix(i, grdExcel.ColIndex("AmountNet")) = Round(mValue, 3)
            
            
        End If
    End If
Next

End Sub

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

Private Sub cmbGroupId_Click(Area As Integer)
  cmbEyeDet(8).BoundText = cmbGroupId.BoundText
End Sub

Private Sub cmbStoreID2_Click(Area As Integer)
GetDataStoreQty2
End Sub

Private Sub CmdCreateV7_Click()

'If (TxtNoteSerial7.Text) = "" Then
cmdDelNote7_Click
    If createVoucher7 Then
       'FindRec val(TXTLCNO.Text)
       
            s = "Update TblCaptinTrans Set NoteID = " & val(txtNoteID7) & ",NoteSerial = '" & Trim(TxtNoteSerial) & "' Where Id = " & val(TxtSerial1(mIndex))
            
                    
            Cn.Execute s
            
            FindRec val(TxtSerial1(mIndex).Text)
            If SystemOptions.UserInterface = ArabicInterface Then
               ' MsgBox " „ «‰‘«¡ «·ﬁÌœ"
                If val(txtNoteID7) <> 0 Then
                    CmdCreateV7.Enabled = False
                    cmdPrintNote7.Enabled = True
                    cmdDelNote7.Enabled = True
                    btn_Save(mIndex).Enabled = False
                Else
                    CmdCreateV7.Enabled = True
                    cmdPrintNote7.Enabled = False
                    cmdDelNote7.Enabled = False
                End If
            Else
            
              '  MsgBox "Done"
            End If
    Else
        CmdCreateV7.Enabled = True
        cmdPrintNote7.Enabled = False
        cmdDelNote7.Enabled = False
    End If
    
'End If

End Sub

Private Sub cmdDelNote7_Click()

Dim x As Integer
Dim Msg As String
Dim StrSQL As String
    
        x = vbYes

      If x = vbYes Then
        StrSQL = "Delete From DOUBLE_ENTREY_VOUCHERS Where Notes_ID=" & val(Me.txtNoteID7.Text)
        Cn.Execute StrSQL, , adExecuteNoRecords
        StrSQL = "Delete From Notes Where NoteID=" & val(Me.txtNoteID7.Text)
        Cn.Execute StrSQL, , adExecuteNoRecords
       ' Cn.Execute " Update TblCaptinTrans set NoteID=null ,NoteSerial=null where ID=" & val(TxtSerial1(mIndex).Text)
       
        
     '   RsSavRec.Requery
        txtNoteID7 = ""
        TxtNoteSerial7 = ""
        Dim s As String
        s = "Update TblCaptinTrans Set NoteID = " & val(txtNoteID7) & ",NoteSerial = '" & Trim(TxtNoteSerial7) & "' Where Id = " & val(TxtSerial1(mIndex))
                    
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

Private Sub cmdPrintNote7_Click()
ShowGL_cc Me.TxtNoteSerial7.Text, , 23001
End Sub

Private Sub CmdSelectCus_Click()
    Dim Indxx As Long
    If Me.Option2.value = True Then
        Indxx = 3
       
        FrmSelectVendor.Indxx = Indxx
        Load FrmSelectVendor
        FrmSelectVendor.Indxx = Indxx
        FrmSelectVendor.show
        FrmSelectVendor.Indxx = Indxx
    Else
        Indxx = 4
         FrmSelectVendor.mEmpId = 0
        If DcbEmployee.Text <> "" And val(DcbEmployee.BoundText) <> 0 Then
            
            FrmSelectVendor.mEmpId = val(DcbEmployee.BoundText)
        End If
        FrmSelectVendor.Indxx = Indxx
        Load FrmSelectVendor
        FrmSelectVendor.Indxx = Indxx
        FrmSelectVendor.show
        FrmSelectVendor.Indxx = Indxx
    End If

End Sub

Private Sub CmdSelectEmp_Click()
    Load FrmSelectEmployee
    FrmSelectEmployee.lblFlag.Caption = 2
    FrmSelectEmployee.show
End Sub

Private Sub Command1_Click()
    CreateItems
End Sub


Sub CreateItems(Optional ByVal IsRefreshOnly As Boolean = False)


If Not IsRefreshOnly Then
    Dim mNewCode As String
    
    
    If val(DCBoMain(2).BoundText) > val(DCBoMain(5).BoundText) Then
     Dim xx As String
        
        xx = (val(DCBoMain(5).BoundText))
        
        DCBoMain(5).BoundText = DCBoMain(2).BoundText
        DCBoMain(2).BoundText = xx
    End If
    
    
    
    If val(DCBoMain(3).BoundText) > val(DCBoMain(6).BoundText) Then
     Dim xxx As String
        
        xxx = (val(DCBoMain(6).BoundText))
        
        DCBoMain(6).BoundText = DCBoMain(3).BoundText
        DCBoMain(3).BoundText = xxx
    End If
    
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
        
        Dim rsDummyCheck As New ADODB.Recordset
        s = "Select * from tblItems Where LensesTypesID =" & val(TxtSerial1(mIndex)) & " and ItemID In  (Select Item_ID FROM Transaction_Details  ) "
        rsDummyCheck.Open s, Cn, adOpenForwardOnly, adLockReadOnly
        If Not rsDummyCheck.EOF Then
            MsgBox "·«Ì„ﬂ‰ «· ⁄œÌ· «Ê «·Õ–› ·ÊÃÊœ ⁄œ”«   „ ⁄·ÌÂ« ›Ê« Ì—"
            Exit Sub
        End If
    End If
        
    s = "Delete tblItems Where LensesTypesID =" & val(TxtSerial1(mIndex))
    Cn.Execute s
    
    
     s = "SELECT * FROM tblItems WHERE ItemID = -1 "
    tRs.Open s, Cn, adOpenKeyset, adLockOptimistic, adCmdText
    Dim mm As Long
        
        '**********************
    Dim i As Long
    For i = 1 To GrdItems.Rows - 1
    
        tRs.AddNew
        II = II + 1
        mMaxId = mMaxId + 1
    
        tRs!ItemID = mMaxId
        tRs!LensesTypesID = val(TxtSerial1(mIndex))
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

        'Trim (GrdItems.TextMatrix(i, GrdItems.ColIndex("ItemName")))
        
        tRs!GroupID = val(GrdItems.TextMatrix(i, GrdItems.ColIndex("GroupID")))
        If i = 1 Then
            mNewCode = GetNewCode2(val(GrdItems.TextMatrix(i, GrdItems.ColIndex("GroupID"))), "tblItems")
        Else
            mNewCode = "0" & val(mNewCode) + 1
        End If
        tRs!Fullcode = mNewCode
        tRs!itemcode = mNewCode
        tRs!barCodeNO = mNewCode
        tRs!code = mNewCode
            
        tRs!SphereID = Trim(GrdItems.TextMatrix(i, GrdItems.ColIndex("SphereID")))
        tRs!CylinderID = Trim(GrdItems.TextMatrix(i, GrdItems.ColIndex("CylinderID")))
        
        tRs("MasterType").value = cboMasterType.ListIndex
        
        
        cmbEyeDet(22).BoundText = val(tRs!SphereID & "")
        cmbEyeDet(23).BoundText = val(tRs!CylinderID & "")
        tRs!ItemName = IIf(FnGenrateName = "", Trim(GrdItems.TextMatrix(i, GrdItems.ColIndex("ItemName"))), FnGenrateName)
        tRs!ItemNamee = IIf(FnGenrateName = "", Trim(GrdItems.TextMatrix(i, GrdItems.ColIndex("ItemName"))), FnGenrateName)
        
        mm = 0
        For mm = 0 To cmbEyeDet.count - 1
            If mm <> 7 And mm <> 23 And mm <> 22 And mm <> 8 Then
                tRs(GetFieldName(mm)).value = val(Me.cmbEyeDet(mm).BoundText)
            End If
        Next
        
        
        tRs.update
        
        
        s = "SELECT * FROM TblItemsUnits WHERE ItemID = -1 "
        Set tRs2 = New ADODB.Recordset
        tRs2.Open s, Cn, adOpenKeyset, adLockOptimistic, adCmdText
        tRs2.AddNew
        tRs2!ItemID = mMaxId
        tRs2!UnitID = val(GrdItems.TextMatrix(i, GrdItems.ColIndex("UnitID")))
        tRs2!DefaultUnit = 1
        tRs2!UnitFactor = 1
        tRs2!UnitSalesPrice = IIf(val(GrdItems.TextMatrix(i, GrdItems.ColIndex("Price"))) = 0, val(TxtPrice), val(GrdItems.TextMatrix(i, GrdItems.ColIndex("UnitID"))))
        
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
            grdSphCYL.TextMatrix(i, j) = TxtName(mIndex) & " " & grdSphCYL.TextMatrix(0, j) & " " & rsDummy!SPHT & ""
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



Private Sub Command2_Click()
GetDataStoreQty
End Sub

Private Sub Command3_Click()
Dim tRs2 As New ADODB.Recordset

Dim i As Long

    
    Dim mm As Long
        
        '**********************
    
    For i = 1 To GrdItems.Rows - 1
    
        
        s = "SELECT * FROM TblItemsUnits WHERE ItemID =  " & val(GrdItems.TextMatrix(i, GrdItems.ColIndex("ItemID")))
        s = s & " And UnitID =  " & val(GrdItems.TextMatrix(i, GrdItems.ColIndex("UnitID")))
        s = s & " and ItemID In (Select ItemID from tblItems Where  LensesTypesID =  " & val(TxtSerial1(mIndex)) & ")"
        s = s & " "
        Set tRs2 = New ADODB.Recordset
        tRs2.Open s, Cn, adOpenKeyset, adLockOptimistic, adCmdText
        tRs2!UnitSalesPrice = IIf(val(GrdItems.TextMatrix(i, GrdItems.ColIndex("Price"))) = 0, val(TxtPrice), val(GrdItems.TextMatrix(i, GrdItems.ColIndex("Price"))))
        
        tRs2.update
        
        
    Next
    MsgBox " „  ÕœÌÀ «·«”⁄«—"

End Sub

Private Sub Command4_Click()
'ExportToExcel Me, Grd, "TT", , "grdItems"
tmpGrd.Rows = 1

Dim i As Long

    grdExcel.ColHidden(grdExcel.ColIndex("VatValue")) = True
    grdExcel.ColHidden(grdExcel.ColIndex("AmountNet")) = True
    
    grdExcel.Rows = 1
    FromExcel grdExcel, tmpGrd, Me, , , txtFile.Text, "TblEmployee"
    chkIsVat_Click
'For i = 0 To GrdExcel.Cols - 1
'    If GrdExcel.ColEditMask(i) <> "" Then
'        GrdExcel.ColHidden(i) = False
'    End If
'    'Grd.ColComboList(i) = ""
'Next
End Sub

Private Sub Command5_Click()
CD1.ShowOpen
txtFile.Text = CD1.filename
End Sub

Private Sub grdTransfer_CellButtonClick(ByVal Row As Long, ByVal Col As Long)
Select Case grdTransfer.ColKey(Col)
   
    Case "CreateOrder"
        
        
            If Trim(Me.grdTransfer.TextMatrix(Me.grdTransfer.Row, Me.grdTransfer.ColIndex("RequestTypeName"))) = " ÕÊÌ· „Œ“‰Ï" Then
                CreateIssueVoucher Row, grdTransfer
            Else
                CreatePurchOrder Row, grdTransfer
            End If
        
   
End Select
End Sub

Private Sub grdTransfer2_CellButtonClick(ByVal Row As Long, ByVal Col As Long)
Select Case grdTransfer2.ColKey(Col)
   
    Case "CreateOrder"
        
                If val(Me.grdTransfer2.TextMatrix(Me.grdTransfer2.Row, Me.grdTransfer2.ColIndex("StatusID"))) = 2 Then
                    If val(grdTransfer2.TextMatrix(Row, grdTransfer2.ColIndex("RequestTypeNo"))) = 2 Then
                    CreateIssueVoucher Row, grdTransfer2
                        
                    ElseIf val(grdTransfer2.TextMatrix(Row, grdTransfer2.ColIndex("RequestTypeNo"))) = 1 Then
                        CreatePurchOrder Row, grdTransfer2
                    End If
                End If
           
    Case "IsFinish"
        
            Dim s As String
            s = "Update Transaction_Details Set IsFinish = 1 Where Id =  " & val(Me.grdTransfer2.TextMatrix(Me.grdTransfer2.Row, Me.grdTransfer2.ColIndex("mmID")))
            Cn.Execute s
            MsgBox " „ «·«‰Â«¡"
       GetDataStoreQty2
   
End Select
End Sub


Private Sub Grid2_EnterCell()
    On Error GoTo ErrTrap
    FindRec val(Me.GRID2.TextMatrix(Me.GRID2.Row, Me.GRID2.ColIndex("id")))
ErrTrap:
End Sub


Private Sub Grid3_EnterCell()
    On Error GoTo ErrTrap
    FindRec val(Me.Grid3.TextMatrix(Me.Grid3.Row, Me.Grid3.ColIndex("id")))
    Dim s As String
    s = ""


s = " SELECT ti.ItemID,"
s = s & "       ti.ItemName,"
s = s & "       g.GroupName,Ti.ItemCode,"
s = s & "       TblUnites.UnitName,Ti.GroupID,TblItemsUnits.UnitID,Ti.SphereID,Ti.CylinderID,"
s = s & "       SPHTable.SPH , CLYTable.CLY as CYL"
s = s & "       ,ti.CylinderID,"
s = s & "       TblItemsUnits.UnitSalesPrice  AS Price"
s = s & " FROM   TblItems                      AS ti"
s = s & "       INNER JOIN Groups             AS g"
s = s & "            ON  g.GroupID = ti.GroupID"
s = s & "       INNER JOIN TblItemsUnits"
s = s & "            ON  TblItemsUnits.ItemID = ti.ItemID"
s = s & "       INNER JOIN TblUnites"
s = s & "            ON  TblUnites.UnitID = TblItemsUnits.UnitID"
s = s & "            LEFT OUTER JOIN CLYTable"
s = s & "            ON CLYTable.ID = ti.CylinderID"
s = s & "            LEFT OUTER JOIN SPHTable"
s = s & "            ON SPHTable.ID = ti.SphereID"
s = s & " Where LensesTypesID =  " & val(TxtSerial1(mIndex))

loadgrid s, GrdItems, True

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
       
            s = "Update TblHandWages Set NoteID = " & val(TXTNoteID) & ",NoteSerial = '" & Trim(TxtNoteSerial) & "' Where Id = " & val(TxtSerial1(mIndex))
            
                    
            Cn.Execute s
            
            FindRec val(TxtSerial1(mIndex).Text)
        If SystemOptions.UserInterface = ArabicInterface Then
           ' MsgBox " „ «‰‘«¡ «·ﬁÌœ"
            If val(TXTNoteID) <> 0 Then
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

BranchID = val(dcBranch(mIndex).BoundText)
mRate = 1

'



notytype = 1100
Notevalue = val(txtNet)

'mAccNO = val(DboParentAccount.BoundText)
NoteDate = (XPDtbTrans.value)
 
If Notevalue > 0 Then
    CreateNotes NoteID, NoteDate, BranchID, notytype, Notevalue, NoteSerial, (NoteSerial1), tablename, Filedname, val(TxtSerial1(mIndex)), des                   ', recordDateH.value
                                              TXTNoteID.Text = NoteID
                                                     TxtNoteSerial.Text = NoteSerial

    If Not CREATE_VOUCHER_GE2(val(TXTNoteID.Text), BranchID, val(DCboUserName(mIndex).BoundText), NoteDate) Then createVoucher2 = False Else createVoucher2 = True
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



Function createVoucher7() As Boolean

'ee
Dim NoteID As Long
Dim NoteDate As Date
Dim NoteSerial As String
Dim Notevalue As Double
Dim notytype As Integer
Dim des As String
des = "    Õ”«» «·" & TxtNoteSerial7.Text


Dim tablename As String
Dim Filedname As String
Dim NoteSerial1 As Long
Dim BranchID As Integer
Dim sql As String
Dim mRate  As Double
tablename = "TblCaptinTrans"

Filedname = "ID"
NoteSerial1 = val(TxtNoteSerial17)

BranchID = val(dcBranch(mIndex).BoundText)
mRate = 1

'

Dim i As Long
Notevalue = 0
For i = 1 To grdExcel.Rows - 1
    Notevalue = Notevalue + Abs(val(grdExcel.TextMatrix(i, grdExcel.ColIndex("Amount"))))
Next

notytype = 23001


'mAccNO = val(DboParentAccount.BoundText)
NoteDate = (XPDtbTrans7.value)
 
If Notevalue > 0 Then
    CreateNotes NoteID, NoteDate, BranchID, notytype, Notevalue, NoteSerial, (NoteSerial1), tablename, Filedname, val(TxtSerial1(mIndex)), des                   ', recordDateH.value
                                              txtNoteID7.Text = NoteID
                                                     TxtNoteSerial7.Text = NoteSerial

    If Not CREATE_VOUCHER_GECaptin(val(txtNoteID7.Text), BranchID, val(DCboUserName(mIndex).BoundText), NoteDate) Then createVoucher7 = False Else createVoucher7 = True
    RsSavRec.Resync adAffectCurrent

    updateNotesValueAndNobytext val(TxtNoteSerial7.Text), Format(Notevalue, "###.00")
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
    Dim x As Integer
   
    Dim rs As New ADODB.Recordset
    Dim notes_serial As String
    Dim notes_id As String
    Msg = "    Õ”«» " & TxtSerial1(mIndex).Text
    notes_id = general_noteid
    my_branch = val(dcBranch(mIndex))
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

        If ModAccounts.AddNewDev(LngDevID, line_no, StrAccountCodeCridet, val(TxtVAt2), 1, Msg & "    Õ”«»  «·ﬁÌ„… «·„÷«›… ", val(notes_id), , , , XPDtbTrans.value, val(DCboUserName(mIndex).BoundText), , , , , , CLng(mRate), , , setfoxy_Line, , , , , , , , , val(branch_id)) = False Then
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
TXTNoteID = ""
TxtNoteSerial = ""
CmdCreateV2.Enabled = True
  End Function

Private Sub cmdDelNote_Click()

Dim x As Integer
Dim Msg As String
Dim StrSQL As String
    
        x = vbYes

      If x = vbYes Then
        StrSQL = "Delete From DOUBLE_ENTREY_VOUCHERS Where Notes_ID=" & val(Me.TXTNoteID.Text)
        Cn.Execute StrSQL, , adExecuteNoRecords
        StrSQL = "Delete From Notes Where NoteID=" & val(Me.TXTNoteID.Text)
        Cn.Execute StrSQL, , adExecuteNoRecords
        Cn.Execute " Update ContainerContracts set NoteID=null ,NoteSerial=null where ID=" & val(TxtSerial1(mIndex).Text)
       
        
     '   RsSavRec.Requery
        TXTNoteID = ""
        TxtNoteSerial = ""
        Dim s As String
        s = "Update TblHandWages Set NoteID = " & val(TXTNoteID) & ",NoteSerial = '" & Trim(TxtNoteSerial) & "' Where Id = " & val(TxtSerial1(mIndex))
                    
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




Public Function CREATE_VOUCHER_GECaptin(general_noteid As Long, BranchID As Integer, UserID As Long _
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
    Dim mCustName As String
    Dim Msg As String
    Dim StrAccountCodeDebt As String
    Dim StrAccountCodeCridet As String
    Dim x As Integer
   Dim AccountVATCreit As String
    Dim rs As New ADODB.Recordset
    Dim notes_serial As String
    Dim notes_id As String
    Msg = "    Õ”«» " & TxtSerial1(mIndex).Text
    notes_id = general_noteid
    my_branch = val(dcBranch(mIndex))
    LngDevID = new_id("DOUBLE_ENTREY_VOUCHERS", "Double_Entry_Vouchers_ID", "")
    Dim line_no As Integer
    line_no = 1
    Dim Percentg  As Double
    Dim s As String
    Dim mRate As Double
    Dim DateEntry As Date
    mRate = 1
    ' „‰ Õ”«» «·⁄„Ì·
    
    'XPDtbTrans
    
    StrAccountCodeCridet = get_account_code_branch(2, my_branch)
        
    If StrAccountCodeCridet = "NO branch" Then
        If SystemOptions.UserInterface = ArabicInterface Then
            MsgBox "·„ Ì „ «‰‘«¡ «·›—⁄", vbCritical
        Else
            MsgBox "Branch Not Created", vbCritical
        End If

        GoTo ErrTrap
    ElseIf StrAccountCodeCridet = "NO account" Then

        If SystemOptions.UserInterface = ArabicInterface Then
            MsgBox "·„ Ì „  ÕœÌœ Õ”«»  «·„»Ì⁄«   ›Ì «·›—⁄ ·Â–… «·⁄„·Ì…", vbCritical
        Else
            MsgBox "Sales Account Not Defined in this Branch", vbCritical
        End If

        GoTo ErrTrap
         
    End If
    Dim mVat As Double
    Dim mDisc As String
    For i = 1 To grdExcel.Rows - 1
            mCustName = Trim(grdExcel.TextMatrix(i, grdExcel.ColIndex("CompanyName")))
            Notevalue = val(grdExcel.TextMatrix(i, grdExcel.ColIndex("Amount")))
            If IsDate(grdExcel.TextMatrix(i, grdExcel.ColIndex("DateEntry"))) Then
                DateEntry = CDate(grdExcel.TextMatrix(i, grdExcel.ColIndex("DateEntry")))
            Else
                DateEntry = XPDtbTrans7.value
            End If
            If Notevalue < 0 Then
               StrAccountCodeDebt = Trim(grdExcel.TextMatrix(i, grdExcel.ColIndex("Account_Code")))
            Else
                'StrAccountCodeDebt = ModAccounts.GetMyAccountCode("TblCustemers", "CusID", val(Me.DcCustmer(mIndex).BoundText))
                StrAccountCodeDebt = get_bank_Account(val(Me.DcboBankName.BoundText), "Account_Code")
            End If
            
            mDisc = Trim(grdExcel.TextMatrix(i, grdExcel.ColIndex("CompanyName"))) & " " & Trim(grdExcel.TextMatrix(i, grdExcel.ColIndex("OperationName"))) & " " & Trim(grdExcel.TextMatrix(i, grdExcel.ColIndex("DateEntry")))
        
            Notevalue = Abs(Notevalue)
            PercentgValueAddedAccount_Transec XPDtbTrans7.value, 21, 1, AccountVATCreit, Percentg

          
            If Notevalue <> 0 Then
                mVat = val(grdExcel.TextMatrix(i, grdExcel.ColIndex("VatValue")))
'                mVat = Notevalue * Percentg / 100
'
                If chkIsVat.value = vbUnchecked Then
                    Notevalue = Notevalue + mVat
                End If
               ' StrAccountCodeDebt = Trim(DboParentAccount.BoundText)
               
                If ModAccounts.AddNewDev(LngDevID, line_no, StrAccountCodeDebt, Notevalue, 0, mDisc & "    Õ”«»  «·⁄„Ì·  ", val(notes_id), , , , IIf(chkIsDateLine.value = vbChecked, DateEntry, XPDtbTrans7.value), val(DCboUserName(mIndex).BoundText), , , , , , CLng(mRate), , , setfoxy_Line, , , , , , , , , _
                val(my_branch)) = False Then
                    GoTo ErrTrap
                End If
               ' «·Ï Õ”«» «·ﬁÌ„… «·„÷«›…
                
                
                line_no = line_no + 1
                
                
            
                
                
                
                If ModAccounts.AddNewDev(LngDevID, line_no, AccountVATCreit, val(mVat), 1, Msg & "    Õ”«»  «·ﬁÌ„… «·„÷«›… ", val(notes_id), , , , IIf(chkIsDateLine.value = vbChecked, DateEntry, XPDtbTrans7.value), val(DCboUserName(mIndex).BoundText), , , , , , CLng(mRate), , , setfoxy_Line, , , , , , , , , val(my_branch)) = False Then
                    GoTo ErrTrap
                End If
                line_no = line_no + 1
            
        
            
            ' «·«ÿ—«›
            
             ' «·Ï Õ”«» «Ì—«œ«  «·Õ«ÊÌ« 
                 
                If chkIsVat.value = vbUnchecked Then
                    Notevalue = Abs(val(grdExcel.TextMatrix(i, grdExcel.ColIndex("Amount"))))
                Else
                   Notevalue = val(grdExcel.TextMatrix(i, grdExcel.ColIndex("AmountNet")))
                End If
                
                
                 StrAccountCodeCridet = get_account_code_branch(2, my_branch)
                If ModAccounts.AddNewDev(LngDevID, line_no, StrAccountCodeCridet, Notevalue, 1, mDisc & "    Õ”«» «·„»Ì⁄«   ", val(notes_id), , , , IIf(chkIsDateLine.value = vbChecked, DateEntry, XPDtbTrans7.value), val(DCboUserName(mIndex).BoundText), , , , , , CLng(mRate), , , setfoxy_Line, , , , , , , , , _
                val(branch_id)) = False Then
                    GoTo ErrTrap
                End If
        
                line_no = line_no + 1
        End If
      Next

    updateNotesValueAndNobytext (val(notes_id))
    CREATE_VOUCHER_GECaptin = True
    Exit Function
ErrTrap:
CREATE_VOUCHER_GECaptin = False
txtNoteID7 = ""
TxtNoteSerial7 = ""
CmdCreateV7.Enabled = True
  End Function


Private Sub cmdPrintNote_Click()

ShowGL_cc Me.TxtNoteSerial.Text, , 1100

End Sub
Private Sub CBoBasedON_Change()
    If mIndex = 8 Then Exit Sub
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
                          x As Single, _
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
                          txtTotalInvoice = Round(val(rsDummy!total & ""), 2)
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
    txtTotalVat = val(txtVat2Invoice) + val(TxtVAt2)
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
    
            FG.Rows = 1
            FG.Rows = 2

    
        
    End If
ElseIf mIndex = 7 Then
    If Me.TxtModFlg2(mIndex).Text <> "R" Then
    
            grdExcel.Rows = 1
            grdExcel.Rows = 2

    
        
    End If

End If

End Sub

Private Sub Cmd_DeleteRow_Click(Index As Integer)
If Me.TxtModFlg.Text <> "R" Then

    

    RemoveGridRow




End If
End Sub
Private Sub RemoveGridRow()
    If mIndex = 1 Then
        With Me.FG
    'MsgBox .Row
            If .Row <= 0 Then
                    .Rows = 2
            Exit Sub
            Else
            .RemoveItem .Row
            End If
        End With
    ElseIf mIndex = 7 Then
        With Me.grdExcel
    'MsgBox .Row
            If .Row <= 0 Then
                    .Rows = 2
            Exit Sub
            Else
            .RemoveItem .Row
            End If
        End With
    End If
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



    Dim rsDummyCheck As New ADODB.Recordset
    s = "Select * from tblItems Where LensesTypesID =" & val(TxtSerial1(mIndex)) & " and ItemID In  (Select Item_ID FROM Transaction_Details  ) "
    rsDummyCheck.Open s, Cn, adOpenForwardOnly, adLockReadOnly
    If Not rsDummyCheck.EOF Then
        MsgBox "·«Ì„ﬂ‰ «· ⁄œÌ· «Ê «·Õ–› ·ÊÃÊœ ⁄œ”«   „ ⁄·ÌÂ« ›Ê« Ì—"
        Exit Sub
    End If
    
   
'If TxtNoteSerial <> "" Then
'MsgBox "·« Ì„ﬂ‰ «·Õ–› «Ê «· ⁄œÌ· «·« »⁄œ Õ–› «·ﬁÌœ"
'Exit Sub
End If





If mIndex = 1 Then
    If TxtNoteSerial <> "" Then
        cmdDelNote_Click
    End If
ElseIf mIndex = 7 Then
    cmdDelNote7_Click
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
ElseIf mIndex = 3 Then

    Dim rsDummyCheck As New ADODB.Recordset
    s = "Select * from tblItems Where LensesTypesID =" & val(TxtSerial1(mIndex)) & " and ItemID In  (Select Item_ID FROM Transaction_Details  ) "
    rsDummyCheck.Open s, Cn, adOpenForwardOnly, adLockReadOnly
    If Not rsDummyCheck.EOF Then
        MsgBox "·«Ì„ﬂ‰ «· ⁄œÌ· «Ê «·Õ–› ·ÊÃÊœ ⁄œ”«   „ ⁄·ÌÂ« ›Ê« Ì—"
        Exit Sub
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
        dcBranch(1).BoundText = branch_id
            
        FG.Rows = 1
        FG.Rows = 2
      
   ElseIf mIndex = 2 Then
        My_SQL = "TblOffice"
   ElseIf mIndex = 3 Then
        My_SQL = "TblLensesTypes"
        
  ElseIf mIndex = 7 Then
        My_SQL = "TblCaptinTrans"
        DCboUserName(mIndex).BoundText = user_id
        dcBranch(mIndex).BoundText = branch_id
        grdExcel.Rows = 1
        chkIsDiscountOnly.value = vbChecked
        chkIsAddOnly.value = vbChecked
    
   
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
'   On Error GoTo ErrTrap
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
   
    
ElseIf mIndex = 7 Then
    
        If DcboBankName.Text = "" Then
            MsgBox "Ì—ÃÏ «œŒ«· «·»‰ﬂ"
            DcboBankName.SetFocus
            Exit Sub
        End If
 
        If dcBranch(mIndex).Text = "" Then
            MsgBox "Ì—ÃÏ «œŒ«· «·›—⁄"
            dcBranch(mIndex).SetFocus
            Exit Sub
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
            
                 Dim rsDummyCheck As New ADODB.Recordset
                 s = "Select * from tblItems Where LensesTypesID =" & val(TxtSerial1(mIndex)) & " and ItemID In  (Select Item_ID FROM Transaction_Details  ) "
                 rsDummyCheck.Open s, Cn, adOpenForwardOnly, adLockReadOnly
                 If Not rsDummyCheck.EOF Then
                     MsgBox "·«Ì„ﬂ‰ «· ⁄œÌ· «Ê «·Õ–› ·ÊÃÊœ ⁄œ”«   „ ⁄·ÌÂ« ›Ê« Ì—"
                     Exit Sub
                 End If
                 
                
              '  AddNewRec
               FiLLRec3
       
            ElseIf mIndex = 7 Then
            
                 
        If TxtNoteSerial17.Text = "" Then
                If Voucher_coding(val(dcBranch(mIndex).BoundText), XPDtbTrans7.value, 7, 23001, , , , , , , "TblCaptinTrans") = "error" Then
                    If SystemOptions.UserInterface = ArabicInterface Then
                        MsgBox " ·« Ì„ﬂ‰ «÷«›… ”‰œ  ÃœÌœ ·«‰ﬂ  ⁄œÌ  «·Õœ «·–Ì ﬁ„  » ÕœÌœ… „‰ «·”‰œ«   ": Exit Sub
                    Else
                        MsgBox " Cant't Create Expenses Voucher to this Process no You exceed the maximum number ": Exit Sub
                    End If

                Else
         
                    If Voucher_coding(val(dcBranch(mIndex).BoundText), XPDtbTrans7.value, 7, 23001, , , , , , , "TblCaptinTrans") = "" Then
                        If SystemOptions.UserInterface = ArabicInterface Then
                            
                            TxtNoteSerial17.locked = False
                            MsgBox " ·«»œ „‰ ﬂ «»… —ﬁ„ «·”‰œ ÌœÊÌ« ﬂ„« Õœœ   ": Exit Sub
                        Else
                            MsgBox "  Enter Voucher No Manually or Define Coding ": Exit Sub
                        End If

                    Else
                        TxtNoteSerial17.Text = Voucher_coding(val(dcBranch(mIndex).BoundText), XPDtbTrans7.value, 7, 23001, , , , , , , "TblCaptinTrans")
                    End If
                End If
            End If
                
              '  AddNewRec
               FiLLRec7
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
          ElseIf mIndex = 7 Then
                FiLLRec7
          
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
            ElseIf mIndex = 7 Then
                FiLLTXT7
        
                
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
    TxtVAt2 = IIf(IsNull(RsSavRec("Vat2").value), "", RsSavRec("Vat2").value)
    txtVatYou = IIf(IsNull(RsSavRec("VatYou").value), "", RsSavRec("VatYou").value)
    txtNet = IIf(IsNull(RsSavRec("Net").value), "", RsSavRec("Net").value)
    
    
   CBoBasedON.ListIndex = IIf(IsNull(RsSavRec("CBoBasedON").value), -1, RsSavRec("CBoBasedON").value)
    TXTOrDer_no(0) = IIf(IsNull(RsSavRec("OrDer_no").value), "", RsSavRec("OrDer_no").value)
    TXTOrDer_no(1) = IIf(IsNull(RsSavRec("OrDer_no2").value), "", RsSavRec("OrDer_no2").value)
    
    TXTOrDer_no2 = IIf(IsNull(RsSavRec("RowsEstimatedID").value), "", RsSavRec("RowsEstimatedID").value)
    

   
    dcBranch(mIndex).BoundText = IIf(IsNull(RsSavRec("BranchID").value), "", RsSavRec("BranchID").value)
    TxtRemarks = IIf(IsNull(RsSavRec("Remarks").value), "", RsSavRec("Remarks").value)
    Me.DCboUserName(1).BoundText = IIf(IsNull(RsSavRec("UserID").value), "", RsSavRec("UserID").value)

    
     
    
 
     
    
    
    GetCardAuthorizationData
    
'      TxtNoteID = RsSavRec!NoteID & ""
'    TxtNoteSerial = RsSavRec!NoteSerial & ""
    LoadCar
     TXTNoteID = RsSavRec!NoteID & ""
    TxtNoteSerial = RsSavRec!NoteSerial & ""
    
     If val(TXTNoteID) <> 0 Then
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
    
    loadgrid s, FG, True, True
CalcTotal2
ErrTrap:

End Sub
Public Sub FiLLTXT7(Optional Lngid As Long = 0)

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
    XPDtbTrans7.value = IIf(IsNull(RsSavRec("RecordDate").value), Date, RsSavRec("RecordDate").value)
    Me.TxtNoteSerial17.Text = IIf(IsNull(RsSavRec("NoteSerial1").value), "", RsSavRec("NoteSerial1").value)
    

   
    dcBranch(mIndex).BoundText = IIf(IsNull(RsSavRec("BranchID").value), "", RsSavRec("BranchID").value)
    DcboBankName.BoundText = IIf(IsNull(RsSavRec("BankID").value), "", RsSavRec("BankID").value)
     
   ' TxtRemarks = IIf(IsNull(RsSavRec("Remarks").value), "", RsSavRec("Remarks").value)
    Me.DCboUserName(7).BoundText = IIf(IsNull(RsSavRec("UserID").value), "", RsSavRec("UserID").value)
  If RsSavRec.Fields("IsVat").value = True Then
        chkIsVat.value = vbChecked
    Else
        chkIsVat.value = vbUnchecked
     End If
    
     
    
 
     
    
    
     txtNoteID7 = RsSavRec!NoteID & ""
    TxtNoteSerial7 = RsSavRec!NoteSerial & ""
    
     If val(txtNoteID7) <> 0 Then
        CmdCreateV7.Enabled = False
        cmdPrintNote7.Enabled = True
        cmdDelNote7.Enabled = True

     Else
        CmdCreateV7.Enabled = True
        cmdPrintNote7.Enabled = False
        cmdDelNote7.Enabled = False

    End If
    LabCurr_Rec(mIndex).Caption = RsSavRec.AbsolutePosition
    LabCount_Rec(mIndex).Caption = RsSavRec.RecordCount

    Dim s As String
    
    





            
    
    
    s = " SELECT TblCaptinTrans2.* "
    s = s & " ,Account_Code = (SELECT Account_Code FROM TblBoxesData AS tbd WHERE tbd.empid = TblCaptinTrans2.Emp_ID)"
    s = s & " from TblCaptinTrans2 "
    s = s & " Where MasterID = " & val(TxtSerial1(mIndex))
    
    loadgrid s, grdExcel, True, True
    chkIsVat_Click
ErrTrap:

End Sub


 Public Sub FiLLTXT3(Optional Lngid As Long = 0, Optional ByVal misGrid As Boolean = False)

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
        
    
    
    If Not (IsNull(RsSavRec("MasterType").value)) Then
          
            Me.cboMasterType.ListIndex = IIf(IsNull(RsSavRec("MasterType").value), 1, (RsSavRec("MasterType").value))
         
    Else
        Me.cboMasterType.ListIndex = 1
    End If
    
      
    Dim mm As Long
    For mm = 0 To cmbEyeDet.count - 1
        If mm <> 7 And mm <> 23 And mm <> 22 And mm <> 8 Then
            Me.cmbEyeDet(mm).BoundText = IIf(IsNull(RsSavRec(GetFieldName(mm)).value), "", RsSavRec(GetFieldName(mm)).value)
        End If
    Next
    'cmbSex.ListIndex = IIf(IsNull(rs("SexID").value), -1, rs("SexID").value)
    'cmbAge.ListIndex = IIf(IsNull(rs("AGEID").value), -1, rs("AGEID").value)
    
    
    cmbEyeDet(8).BoundText = cmbGroupId.BoundText
    

    
    
     
   TxtPrice = IIf(IsNull(RsSavRec("Price").value), "", RsSavRec("Price").value)
    TxtName(mIndex).Text = IIf(IsNull(RsSavRec.Fields("name").value), "", RsSavRec.Fields("name").value)
    TxtNameE(mIndex).Text = IIf(IsNull(RsSavRec.Fields("nameE").value), "", RsSavRec.Fields("nameE").value)
    
    LabCurr_Rec(mIndex).Caption = RsSavRec.AbsolutePosition
    LabCount_Rec(mIndex).Caption = RsSavRec.RecordCount


     
    
    
    LabCurr_Rec(mIndex).Caption = RsSavRec.AbsolutePosition
    LabCount_Rec(mIndex).Caption = RsSavRec.RecordCount

    Dim s As String
    
    
    
    

'GenreateItems
s = ""


s = " SELECT ti.ItemID,"
s = s & "       ti.ItemName,"
s = s & "       g.GroupName,Ti.ItemCode,"
s = s & "       TblUnites.UnitName,Ti.GroupID,TblItemsUnits.UnitID,Ti.SphereID,Ti.CylinderID,"
s = s & "       SPHTable.SPH , CLYTable.CLY as CYL"
s = s & "       ,ti.CylinderID,"
s = s & "       TblItemsUnits.UnitSalesPrice  AS Price"
s = s & " FROM   TblItems                      AS ti"
s = s & "       INNER JOIN Groups             AS g"
s = s & "            ON  g.GroupID = ti.GroupID"
s = s & "       INNER JOIN TblItemsUnits"
s = s & "            ON  TblItemsUnits.ItemID = ti.ItemID"
s = s & "       INNER JOIN TblUnites"
s = s & "            ON  TblUnites.UnitID = TblItemsUnits.UnitID"
s = s & "            LEFT OUTER JOIN CLYTable"
s = s & "            ON CLYTable.ID = ti.CylinderID"
s = s & "            LEFT OUTER JOIN SPHTable"
s = s & "            ON SPHTable.ID = ti.SphereID"
s = s & " Where LensesTypesID =  " & val(TxtSerial1(mIndex))

loadgrid s, GrdItems, True
            


    If misGrid Then

    With Grid3

        For i = 1 To .Rows - 1

            If Trim(TxtSerial1(mIndex).Text) = .TextMatrix(i, .ColIndex("id")) Then
                TxtSerial1(mIndex).Text = .TextMatrix(i, .ColIndex("Ser"))
                .Row = i
                Exit Sub
            End If

        Next

    End With
 End If
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

 

Private Sub FG_AfterEdit(ByVal Row As Long, ByVal Col As Long)
    



    Dim StrAccountCode As String
    Dim Msg As String
    Dim rs As New ADODB.Recordset
    Dim StrSQL As String
    Dim ClsAcc As New ClsAccounts
    Dim LngRow As Long
Dim StrComboList As String
    'Dim Rs2 As ADODB.Recordset
  
On Error GoTo ErrTrap
    With FG
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
With FG

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
    
   With FG
    .IsSubtotal(.Rows - 1) = True
    Dim SngTotal As Single
    If .Rows > 1 Then
        txtTotal2 = .Aggregate(flexSTSum, .FixedRows, .ColIndex("Price"), .Rows - 1, .ColIndex("Price"))
    End If
    

    
    End With
    CalculteValueAdded2
    txtNetInvoice2 = Round(val(txtTotal) + val(TxtVAt2), 2)
    txtGeneralTotal = val(txtTotalInvoice) + val(txtTotal2)
    txtTotalDisc = val(txtDiscValueInvoice) + val(txtDiscValue)
    txtTotalVat = val(txtVat2Invoice) + val(TxtVAt2)
    txtTotalBVat = val(txtTotal) + val(txtTotalInvoiceBVat)
    txtTotalNet = val(txtNetInvoice) + val(txtNet)
End Sub


Public Sub CalculteValueAdded2(Optional posDelete As Boolean = False)

txtTotal = val(txtTotal2) - val(txtDiscValue)

txtNet = val(txtTotal) + val(TxtVAt2)
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
   
     TxtVAt2 = val(txtTotal) * Percentg / 100
     
     
     txtNet = val(txtTotal) + val(TxtVAt2)
    

End If

End Sub


 

Private Sub Fg_KeyDown(KeyCode As Integer, Shift As Integer)
'GridKeyDown Fg, KeyCode, Shift, False, False, Fg.Row
 'GridKeyDown Fg, KeyCode, Shift
 mGridClicked = True
        Dim mOldRow As Long
        mOldRow = FG.Row
     GridKeyDown FG, KeyCode, Shift, False, False, FG.Row
     If mOldRow <> FG.Row And FG.Row <> 1 Then
        FG.TextMatrix(FG.Row, FG.ColIndex("DeparmentID")) = FG.TextMatrix(FG.Row - 1, FG.ColIndex("DeparmentID"))
        FG.TextMatrix(FG.Row, FG.ColIndex("DepartmentName")) = FG.TextMatrix(FG.Row - 1, FG.ColIndex("DepartmentName"))
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
   
    With FG

        Select Case .ColKey(Col)
              
        
             
            Case "DepartmentName"
                .TextMatrix(Row, .ColIndex("DepartmentName")) = ""
                StrSQL = "SELECT DeparmentID,DepartmentName  FROM TblEmpDepartments "
                rs.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

                If SystemOptions.UserInterface = ArabicInterface Then
                    StrComboList = FG.BuildComboList(rs, "DepartmentName", "DeparmentID")
                Else
                    StrComboList = FG.BuildComboList(rs, "DepartmentNamee", "DeparmentID")
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



Public Sub CreateIssueVoucher(ByVal Row As Long, ByVal mGrid As vsFlexGrid)

    On Error GoTo errortrap
    'DeleteTransactiomsVoucher Val(Text1.text)


Dim mmID  As Long

    Dim i As Long
    Dim LngCurItemID As Double
    Dim LngUnitID As Long
    Dim UnitFactor As Double
    

    Dim RsTrans As ADODB.Recordset
    Dim rsTrans2 As ADODB.Recordset
    Dim rsTransDet As ADODB.Recordset
    Dim s As String
    Dim mOrderId As Long
    Dim mmTransaction_ID As Long
    Dim sql As String
    Dim MYWAER As String
    Dim StrSQL As String
    Dim RsNotes As ADODB.Recordset
    Dim MYinvnum As String
    Dim note_id As Long

    Dim RSTransDetails As ADODB.Recordset
    Dim RsTemp As New ADODB.Recordset
    Dim RowNum As Integer
    Dim StrSqlDel As String
    Dim SearchResault As Integer
    'Dim Note_ID As Long
    Dim RsDetalis  As ADODB.Recordset
    Dim BeginTrans As Boolean
    Dim LnItemID As Long
 
    Dim StrCurrentItemName As String
    Dim DblNotesTotal As Double

    Dim IntLineNO As Integer
    Dim StrAccountCode As String
    '  Dim RowNum As Integer
    Dim Frm As Form
    Dim Msg As String
    Dim MYTEXT As String
    '›Ì Õ«·… «·«‰ «Ã «·‰„ÿÌ
    Dim TxtNoteSerialV As String
    Dim mDate As Date
    Dim mStoreId As Integer
    Dim mUserID As Long
    Dim xyeas As Boolean
    Dim Transaction_ID As Long
    
    Dim general_noteid As Long
    Dim RsNotesGeneral As ADODB.Recordset
     Dim rs As New ADODB.Recordset
    
    Dim mNoteSerial1 As String
    Dim TxtNoteSerial1V As String
    Dim mTotalCost As Double
    Dim CurrentVoucherNo As String
    
    
        
    
    
    
    
    Dim RsTest As New ADODB.Recordset
    Dim RsRepeat As ADODB.Recordset
    
    
    
    
    Dim NoteID As Long
    
    Dim LngItemID As Long
    Dim Posted As Integer
    Dim Transaction_Type As Integer
    Dim Transaction_Type2 As Integer
    Dim mStoreId2 As Long
    Dim mBranchID As Long
    Dim mEmp_ID As Long
    Dim mNoteId As Long
   On Error GoTo ErrTrap
   Screen.MousePointer = vbArrowHourglass
         Transaction_Type = 10
         Transaction_Type2 = 11
         Posted = 0

    

ll:
Dim TransBegine As Boolean

Cn.BeginTrans
TransBegine = True
 
 '   mOrderId = val(mGrid.TextMatrix(i, mGrid.ColIndex("Transaction_ID")))
    s = " select * from Transactions WHERE Transaction_ID =  " & mOrderId
    Set RsTrans = New ADODB.Recordset
    RsTrans.Open s, Cn, adOpenKeyset, adLockOptimistic
    rs.Open s, Cn, adOpenKeyset, adLockOptimistic
        
        
      
        MYTEXT = CStr(new_id("Transactions", "Transaction_Serial", "", True, "Transaction_Type=10 or Transaction_Type=992"))
        my_branch = val(mGrid.TextMatrix(Row, mGrid.ColIndex("BranchID")))
        Dim mNoteSerial As String
            
        
        mDate = Date
        TxtNoteSerialV = Notes_coding(val(my_branch), mDate)
        mStoreId = val(mGrid.TextMatrix(Row, mGrid.ColIndex("StoreIDAvi")))
        mBranchID = val(mGrid.TextMatrix(Row, mGrid.ColIndex("BranchID")))
        mStoreId2 = val(mGrid.TextMatrix(Row, mGrid.ColIndex("StoreId")))
        mEmp_ID = val(mGrid.TextMatrix(Row, mGrid.ColIndex("Emp_ID")))
        mmID = val(mGrid.TextMatrix(Row, mGrid.ColIndex("mmID")))
        mUserID = user_id
        mNoteSerial1 = Voucher_coding(val(my_branch), Date, 12, 190, , 10, , mStoreId)
        Transaction_ID = CStr(new_id("Transactions", "Transaction_ID", "", True))
        TxtNoteSerial1V = Voucher_coding(val(my_branch), Date, 12, 190, , 10, , mStoreId)
        mNoteSerial = Notes_coding(val(branch_id), Date)
        CurrentVoucherNo = ""
        
        
       
        
                   rs.AddNew
            rs("Transaction_ID").value = CStr(new_id("Transactions", "Transaction_ID", "", True))
             mmTransaction_ID = val(rs("Transaction_ID").value)
           ' rs.update
          
            'Me.oldtxtNoteSerial1.Text = Trim$(Me.TxtNoteSerial1.Text)


  
  
        Screen.MousePointer = vbArrowHourglass
    '     rs("NoteSerial").value = IIf(Trim(Me.TxtNoteSerial.text) = "", Null, Trim(Me.TxtNoteSerial.text))
        rs("NoteSerial1").value = mNoteSerial1
        rs("OldNoteSerial1").value = mNoteSerial1
    
'        rs("NoteId").value = val(TxtNoteID.text)
        rs("Transaction_Serial").value = mNoteSerial1
        rs("Transaction_Date").value = Date
        rs("Transaction_Type").value = Transaction_Type
        rs("OrderID").value = 0
        rs("UserID").value = user_id
        rs("StoreID").value = mStoreId
        rs("BranchId").value = mBranchID
        Dim FromstoreAr As String
         Dim FromstoreEn As String
         
       Dim TostoreAr As String
         Dim TostoreEn As String
             
            getStorenames val(mStoreId), FromstoreAr, FromstoreEn
             getStorenames val(mStoreId2), TostoreAr, TostoreEn
         
      
              rs("CusID").value = rs("CusID").value = val(mGrid.TextMatrix(Row, mGrid.ColIndex("CusID")))
        rs("Emp_ID").value = mEmp_ID

'    If Trim$(Me.TxtCashCustomerName.Text) <> "" Then
'        rs("CashCustomerName").value = Trim$(Me.TxtCashCustomerName.Text)
'    Else
        rs("CashCustomerName").value = Null
'    End If
' rs("TransactionComment").value = IIf(Trim$(TxtBillComment.Text) = "", Null, Trim$(TxtBillComment.Text))
'  rs("InspectionReport").value = IIf(Trim$(TxtInspectionReport.Text) = "", Null, Trim$(TxtInspectionReport.Text))
  
  rs("BillBasedOn").value = 0
 rs("order_no").value = 0
   
   
  's("Phone").value = IIf(Trim$(TxtPhone.Text) = "", Null, Trim$(TxtPhone.Text))
  
  
   
        rs.update


        RowNum = Row

            If mGrid.TextMatrix(RowNum, mGrid.ColIndex("ItemID")) <> "" Then

                'Check Repeat Serial
              
       Set RSTransDetails = New ADODB.Recordset
     '   RSTransDetails.Open "[Transaction_Details]", Cn, adOpenStatic, adLockOptimistic, adCmdTable
 StrSQL = "SELECT     dbo.Transaction_Details.* from dbo.Transaction_Details Where (Transaction_ID = -1)"
   RSTransDetails.Open StrSQL, Cn, adOpenKeyset, adLockOptimistic, adCmdText
                RSTransDetails.AddNew
                RSTransDetails("FromstoreAr").value = FromstoreAr
                RSTransDetails("TostoreAr").value = TostoreAr
                RSTransDetails("FromstoreEn").value = FromstoreEn
                RSTransDetails("TostoreEn").value = TostoreEn
                'STransDetails("IsExpirDate").value = IIf(mGrid.TextMatrix(RowNum, mGrid.ColIndex("IsExpirDate")) = "", Null, val(mGrid.TextMatrix(RowNum, mGrid.ColIndex("IsExpirDate"))))
                RSTransDetails("OriginalQty").value = IIf((mGrid.TextMatrix(RowNum, mGrid.ColIndex("Quantity")) = ""), Null, val(mGrid.TextMatrix(RowNum, mGrid.ColIndex("Quantity"))))
                'STransDetails("order_no").value = IIf(mGrid.TextMatrix(RowNum, mGrid.ColIndex("order_no")) = "", Null, mGrid.TextMatrix(RowNum, mGrid.ColIndex("order_no")))
                RSTransDetails("OrderArrivalDate").value = Date '.TextMatrix(RowNum, mGrid.ColIndex("OrderArrivalDate")) = "", Null, mGrid.TextMatrix(RowNum, mGrid.ColIndex("OrderArrivalDate")))
                'STransDetails("FoxyNo").value = IIf((mGrid.TextMatrix(RowNum, mGrid.ColIndex("FoxyNo")) = ""), Null, mGrid.TextMatrix(RowNum, mGrid.ColIndex("FoxyNo")))
                RSTransDetails("Transaction_ID").value = val(Transaction_ID)
                RSTransDetails("Item_ID").value = IIf((mGrid.TextMatrix(RowNum, mGrid.ColIndex("ItemID")) = ""), Null, val(mGrid.TextMatrix(RowNum, mGrid.ColIndex("ItemID"))))
                'RSTransDetails("Remarks").value = IIf((mGrid.TextMatrix(RowNum, mGrid.ColIndex("Remarks")) = ""), Null, Trim$(mGrid.TextMatrix(RowNum, mGrid.ColIndex("Remarks"))))
                If Not mGrid.TextMatrix(RowNum, mGrid.ColIndex("ItemName")) = "" Then
                    StrSQL = "select * From TblItems where ItemID=" & mGrid.TextMatrix(RowNum, mGrid.ColIndex("ItemID"))
                    RsTemp.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText

                    If Not (RsTemp.EOF Or RsTemp.BOF) Then
                        If RsTemp("HaveSerial").value = True Then
                            'RSTransDetails("ItemSerial").value = IIf((mGrid.TextMatrix(RowNum, mGrid.ColIndex("Serial")) = ""), Null, mGrid.TextMatrix(RowNum, mGrid.ColIndex("Serial")))
                        End If
                    End If

                    RsTemp.Close
                End If
                RSTransDetails("ItemBalance").value = IIf((mGrid.TextMatrix(RowNum, mGrid.ColIndex("ItemBalance2")) = ""), Null, val(mGrid.TextMatrix(RowNum, mGrid.ColIndex("ItemBalance2"))))
                RSTransDetails("ShowQty").value = IIf((mGrid.TextMatrix(RowNum, mGrid.ColIndex("Quantity")) = ""), Null, val(mGrid.TextMatrix(RowNum, mGrid.ColIndex("Quantity"))))
                RSTransDetails("showPrice").value = IIf((mGrid.TextMatrix(RowNum, mGrid.ColIndex("Price")) = ""), Null, val(mGrid.TextMatrix(RowNum, mGrid.ColIndex("Price"))))
            
           '     RSTransDetails("ItemDiscountType").value = IIf((mGrid.TextMatrix(RowNum, mGrid.ColIndex("DiscountType")) = ""), Null, val(mGrid.TextMatrix(RowNum, mGrid.ColIndex("DiscountType"))))
           '     RSTransDetails("ItemCase").value = IIf((mGrid.TextMatrix(RowNum, mGrid.ColIndex("ItemCase")) = ""), Null, val(mGrid.TextMatrix(RowNum, mGrid.ColIndex("ItemCase"))))
           '     RSTransDetails("ItemDiscount").value = IIf((mGrid.TextMatrix(RowNum, mGrid.ColIndex("DiscountVal")) = ""), Null, val(mGrid.TextMatrix(RowNum, mGrid.ColIndex("DiscountVal"))))
           '     RSTransDetails("guaranteeTime").value = IIf((mGrid.TextMatrix(RowNum, mGrid.ColIndex("guaranteeTime")) = ""), Null, val(mGrid.TextMatrix(RowNum, mGrid.ColIndex("guaranteeTime"))))
            
                RSTransDetails("ColorID").value = 1 'IIf((mGrid.TextMatrix(RowNum, mGrid.ColIndex("ColorID")) = ""), 1, val(mGrid.TextMatrix(RowNum, mGrid.ColIndex("ColorID"))))
                RSTransDetails("ItemSize").value = 1 'IIf((mGrid.TextMatrix(RowNum, mGrid.ColIndex("ItemSize")) = ""), 1, Trim$(mGrid.TextMatrix(RowNum, mGrid.ColIndex("ItemSize"))))
                RSTransDetails("ClassId").value = 1 'IIf((mGrid.TextMatrix(RowNum, mGrid.ColIndex("ClassId")) = ""), 1, val(mGrid.TextMatrix(RowNum, mGrid.ColIndex("ClassId"))))

                RSTransDetails("UnitID").value = val(mGrid.TextMatrix(RowNum, mGrid.ColIndex("UnitID")))

                RSTransDetails("BranchId").value = mBranchID
                RSTransDetails("OrderArrivalDate").value = Date ' IIf(Not IsDate(mGrid.TextMatrix(RowNum, mGrid.ColIndex("OrderArrivalDate"))), Null, mGrid.TextMatrix(RowNum, mGrid.ColIndex("OrderArrivalDate")))
            'RSTransDetails("ItemsDetailsNewidea").value = IIf(mGrid.TextMatrix(RowNum, mGrid.ColIndex("ItemsDetailsNewidea")) = "", Null, mGrid.TextMatrix(RowNum, mGrid.ColIndex("ItemsDetailsNewidea")))

                Dim RsUnitData As ADODB.Recordset
                
                
                Dim DblQty As Double
        
                LngCurItemID = val(mGrid.TextMatrix(RowNum, mGrid.ColIndex("ItemID")))
                LngUnitID = val(mGrid.TextMatrix(RowNum, mGrid.ColIndex("UnitID")))
                DblQty = val(mGrid.TextMatrix(RowNum, mGrid.ColIndex("Quantity")))

                StrSQL = "Select * From TblItemsUnits Where ItemID=" & LngCurItemID
                StrSQL = StrSQL + " AND UnitID=" & LngUnitID
                Set RsUnitData = New ADODB.Recordset
                RsUnitData.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

                If Not (RsUnitData.BOF Or RsUnitData.EOF) Then
                    RSTransDetails("QtyBySmalltUnit").value = RsUnitData("UnitFactor").value
                
                    RSTransDetails("Quantity").value = RSTransDetails("QtyBySmalltUnit").value * RSTransDetails("showqty").value
                    RSTransDetails("Price").value = val(IIf((mGrid.TextMatrix(RowNum, mGrid.ColIndex("Price")) = ""), 0, val(mGrid.TextMatrix(RowNum, mGrid.ColIndex("Price"))))) / RSTransDetails("QtyBySmalltUnit").value
            
                End If
                'RSTransDetails("Height").value = val(mGrid.TextMatrix(RowNum, mGrid.ColIndex("Height")))
                'RSTransDetails("Width").value = val(mGrid.TextMatrix(RowNum, mGrid.ColIndex("Width")))
            
                RSTransDetails("price").value = Round(mGrid.TextMatrix(RowNum, mGrid.ColIndex("Price")) / RSTransDetails("Quantity").value, 2)
                'RSTransDetails("ProductionDate").value = IIf((mGrid.TextMatrix(RowNum, mGrid.ColIndex("ProductionDate")) = ""), Null, Format((mGrid.TextMatrix(RowNum, mGrid.ColIndex("ProductionDate"))), "DD/mm/YYYY"))
                'RSTransDetails("ExpiryDate").value = IIf((mGrid.TextMatrix(RowNum, mGrid.ColIndex("ExpiryDate")) = ""), Null, Format((mGrid.TextMatrix(RowNum, mGrid.ColIndex("ExpiryDate"))), "DD/mm/YYYY"))
                'RSTransDetails("LotNO").value = IIf(mGrid.TextMatrix(RowNum, mGrid.ColIndex("LotNO")) = "", Null, mGrid.TextMatrix(RowNum, mGrid.ColIndex("LotNO")))
             Dim OldQty As Double
             Dim OldCost As Double
              Dim NewQty As Double
               Dim NewCost As Double
               
getItemCostData Date, RSTransDetails("Item_ID").value, val(mStoreId), val(Transaction_ID), OldQty, OldCost, NewQty, NewCost
       RSTransDetails("OldQty").value = NewQty
       RSTransDetails("OldCost").value = NewCost
       
      RSTransDetails("NewQty").value = RSTransDetails("OldQty").value - RSTransDetails("Quantity").value
       RSTransDetails("NewCost").value = RSTransDetails("OldCost").value ' ((RSTransDetails("OldQty").value * RSTransDetails("OldCost").value) + (RSTransDetails("Quantity").value * RSTransDetails("Price").value)) / (RSTransDetails("Quantity").value + RSTransDetails("OldQty").value)
       
       
                RSTransDetails.update
            End If


'SaveItemsData val(Transaction_ID), -1
        '≈÷«›… «·»÷«∆⁄ ≈·Ï «·„Œ“‰ «·ÃœÌœ
        rs.AddNew
        rs("NoteSerial").value = IIf(Trim(mNoteSerial) = "", Null, Trim(mNoteSerial))
        rs("NoteSerial1").value = IIf(Trim(mNoteSerial1) = "", Null, Trim(mNoteSerial1))
        rs("OldNoteSerial1").value = Trim$(mNoteSerial1) '
        rs("NoteId").value = mNoteId
        rs("Transaction_ID").value = CStr(new_id("Transactions", "Transaction_ID", "", True))
        rs("Transaction_Date").value = Date
        rs("Transaction_Type").value = Transaction_Type2
        rs("UserID").value = user_id
         rs("StoreID").value = mStoreId2
        rs("ReturnID").value = val(Transaction_ID)
        rs("BranchId").value = mBranchID
     
        rs("CusID").value = mGrid.TextMatrix(RowNum, mGrid.ColIndex("CusID"))
             
        rs.update


        
        If mGrid.TextMatrix(RowNum, mGrid.ColIndex("ItemID")) <> "" Then
            RSTransDetails.AddNew
            'RSTransDetails("order_no").value = 'IIf(mGrid.TextMatrix(RowNum, mGrid.ColIndex("order_no")) = "", Null, mGrid.TextMatrix(RowNum, mGrid.ColIndex("order_no")))
            RSTransDetails("OrderArrivalDate").value = Date 'IIf(mGrid.TextMatrix(RowNum, mGrid.ColIndex("OrderArrivalDate")) = "", Null, mGrid.TextMatrix(RowNum, mGrid.ColIndex("OrderArrivalDate")))
            'RSTransDetails("FoxyNo").value = IIf((mGrid.TextMatrix(RowNum, mGrid.ColIndex("FoxyNo")) = ""), Null, mGrid.TextMatrix(RowNum, mGrid.ColIndex("FoxyNo")))
            RSTransDetails("Transaction_ID").value = rs("Transaction_ID").value
            RSTransDetails("Item_ID").value = IIf((mGrid.TextMatrix(RowNum, mGrid.ColIndex("ItemID")) = ""), Null, val(mGrid.TextMatrix(RowNum, mGrid.ColIndex("ItemID"))))
             
            RSTransDetails("UnitID").value = IIf((mGrid.TextMatrix(RowNum, mGrid.ColIndex("UnitID")) = ""), Null, val(mGrid.TextMatrix(RowNum, mGrid.ColIndex("UnitID"))))

            If Not mGrid.TextMatrix(RowNum, mGrid.ColIndex("ItemID")) = "" Then
                StrSQL = "select * From TblItems where ItemID=" & mGrid.TextMatrix(RowNum, mGrid.ColIndex("ItemID"))
                RsTemp.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText

                If Not (RsTemp.EOF Or RsTemp.BOF) Then
                    If RsTemp("HaveSerial").value = True Then
                        'RSTransDetails("ItemSerial").value = IIf((mGrid.TextMatrix(RowNum, mGrid.ColIndex("Serial")) = ""), Null, mGrid.TextMatrix(RowNum, mGrid.ColIndex("Serial")))
                    End If
                End If

                RsTemp.Close
            End If

            RSTransDetails("BranchId").value = mBranchID
            
            
            RSTransDetails("OrderArrivalDate").value = Null 'IIf(Not IsDate(mGrid.TextMatrix(RowNum, mGrid.ColIndex("OrderArrivalDate"))), Null, mGrid.TextMatrix(RowNum, mGrid.ColIndex("OrderArrivalDate")))
 
            'RSTransDetails("ItemDiscountType").value = IIf((mGrid.TextMatrix(RowNum, mGrid.ColIndex("DiscountType")) = ""), Null, val(mGrid.TextMatrix(RowNum, mGrid.ColIndex("DiscountType"))))
            'RSTransDetails("ItemCase").value = IIf((mGrid.TextMatrix(RowNum, mGrid.ColIndex("ItemCase")) = ""), Null, val(mGrid.TextMatrix(RowNum, mGrid.ColIndex("ItemCase"))))
            'RSTransDetails("ItemDiscount").value = IIf((mGrid.TextMatrix(RowNum, mGrid.ColIndex("DiscountVal")) = ""), Null, val(mGrid.TextMatrix(RowNum, mGrid.ColIndex("DiscountVal"))))
            'RSTransDetails("guaranteeTime").value = IIf((mGrid.TextMatrix(RowNum, mGrid.ColIndex("guaranteeTime")) = ""), Null, val(mGrid.TextMatrix(RowNum, mGrid.ColIndex("guaranteeTime"))))
            
            RSTransDetails("ColorID").value = 1 'IIf((mGrid.TextMatrix(RowNum, mGrid.ColIndex("ColorID")) = ""), 1, val(mGrid.TextMatrix(RowNum, mGrid.ColIndex("ColorID"))))
            RSTransDetails("ItemSize").value = 1 'IIf((mGrid.TextMatrix(RowNum, mGrid.ColIndex("ItemSize")) = ""), 1, Trim$(mGrid.TextMatrix(RowNum, mGrid.ColIndex("ItemSize"))))
            RSTransDetails("ClassId").value = 1 ' IIf((mGrid.TextMatrix(RowNum, mGrid.ColIndex("ClassId")) = ""), 1, val(mGrid.TextMatrix(RowNum, mGrid.ColIndex("ClassId"))))
           
            RSTransDetails("ShowPrice").value = IIf((mGrid.TextMatrix(RowNum, mGrid.ColIndex("Price")) = ""), Null, val(mGrid.TextMatrix(RowNum, mGrid.ColIndex("Price"))))
            RSTransDetails("ShowQty").value = IIf((mGrid.TextMatrix(RowNum, mGrid.ColIndex("Quantity")) = ""), Null, val(mGrid.TextMatrix(RowNum, mGrid.ColIndex("Quantity"))))

            '---------------------------------------------------
            Dim RsUnitData1 As ADODB.Recordset
        
            LngCurItemID = val(mGrid.TextMatrix(RowNum, mGrid.ColIndex("ItemID")))
            LngUnitID = val(mGrid.TextMatrix(RowNum, mGrid.ColIndex("UnitID")))
            DblQty = val(mGrid.TextMatrix(RowNum, mGrid.ColIndex("Quantity")))

            StrSQL = "Select * From TblItemsUnits Where ItemID=" & LngCurItemID
            StrSQL = StrSQL + " AND UnitID=" & LngUnitID
            Set RsUnitData1 = New ADODB.Recordset
            RsUnitData1.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

            If Not (RsUnitData1.BOF And RsUnitData1.EOF) Then
                RSTransDetails("QtyBySmalltUnit").value = RsUnitData1("UnitFactor").value
                RSTransDetails("Quantity").value = RSTransDetails("QtyBySmalltUnit").value * RSTransDetails("showqty").value
                RSTransDetails("Price").value = val(IIf((mGrid.TextMatrix(RowNum, mGrid.ColIndex("Price")) = ""), 0, val(mGrid.TextMatrix(RowNum, mGrid.ColIndex("Price"))))) / RSTransDetails("QtyBySmalltUnit").value
            
            End If

            'RSTransDetails("price").Value = Round(mGrid.TextMatrix(RowNum, mGrid.ColIndex("Valu")) / RSTransDetails("Quantity").Value, 2)
            'RSTransDetails("ProductionDate").value = IIf((mGrid.TextMatrix(RowNum, mGrid.ColIndex("ProductionDate")) = ""), Null, Format((mGrid.TextMatrix(RowNum, mGrid.ColIndex("ProductionDate"))), "DD/mm/YYYY"))
            'RSTransDetails("ExpiryDate").value = IIf((mGrid.TextMatrix(RowNum, mGrid.ColIndex("ExpiryDate")) = ""), Null, Format((mGrid.TextMatrix(RowNum, mGrid.ColIndex("ExpiryDate"))), "DD/mm/YYYY"))
            'RSTransDetails("LotNO").value = IIf(mGrid.TextMatrix(RowNum, mGrid.ColIndex("LotNO")) = "", Null, mGrid.TextMatrix(RowNum, mGrid.ColIndex("LotNO")))
                                   RSTransDetails("FromstoreAr").value = FromstoreAr
            RSTransDetails("TostoreAr").value = TostoreAr
           RSTransDetails("FromstoreEn").value = FromstoreEn
            RSTransDetails("TostoreEn").value = TostoreEn
            getItemCostData Date, RSTransDetails("Item_ID").value, val(mStoreId), val(Transaction_ID), OldQty, OldCost, NewQty, NewCost
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
        
'SaveItemsData rs("Transaction_ID").value, 1

        Cn.CommitTrans
        BeginTrans = False
    
 CreateNotes NoteID, Date, branch_id, 190, 0, mNoteSerial, mNoteSerial1, "Transactions", "Transaction_ID", val(Transaction_ID), mNoteSerial1, ToHijriDate(Date)
           mNoteId = NoteID
     '      If TxtNoteSerial.text = "" Then
     '      TxtNoteSerial.text = NoteSerial
     '      End If
           
           general_noteid = NoteID
    
        Dim LngDevID As Long
        Dim LngDevNO  As Integer
        Dim StrTempAccountCode As String
        Dim StrTempDes As String

        LngDevID = new_id("DOUBLE_ENTREY_VOUCHERS", "Double_Entry_Vouchers_ID", "")
        '----------------
        Dim Account_Code_dynamic As String
        'SngTemp = NewGrid.GetItemsCostTotal * RSTransDetails("quantity").value / Cnt
        Dim SngTemp  As Double
        SngTemp = val((mGrid.TextMatrix(RowNum, mGrid.ColIndex("Price"))))

        If SngTemp > 0 Then
            '1 work with branch
            '2 work with inventory
            '3 work with groups

            If detect_inventory_work_type = 1 Then
                ' 1«·„Œ“Ê‰ ›Ì «·›—⁄
                Account_Code_dynamic = get_account_code_branch(0, my_branch)
        
                If Account_Code_dynamic = "NO branch" Then
                    MsgBox "·„ Ì „ «‰‘«¡ «·›—⁄", vbCritical
                    
                    GoTo ErrTrap
                Else

                    If Account_Code_dynamic = "NO account" Then
                        MsgBox "·„ Ì „  ÕœÌœ Õ”«»   ﬂ·›… «·„Œ“Ê‰ ›Ì «·›—⁄ ·Â–… «·⁄„·Ì…", vbCritical
                        
                        GoTo ErrTrap
         
                    End If
                End If

                StrTempAccountCode = Account_Code_dynamic '«·„Œ“Ê‰ 0 ›Ì «·›—⁄
    
                If SystemOptions.UserInterface = ArabicInterface Then
                    StrTempDes = "√–‰  ÕÊÌ· »÷«∆⁄ »Ì‰ «·„Œ«“‰  —ﬁ„ " & mNoteSerial1
                Else
                    StrTempDes = "  Moving Items Vchr  No. " & mNoteSerial1
                End If
        
                LngDevNO = 0

                If ModAccounts.AddNewDev(LngDevID, LngDevNO, StrTempAccountCode, SngTemp, 1, StrTempDes, general_noteid, , , , mDate, user_id, val(Transaction_ID), , , , , , , , , , , , , , , , , val(mBranchID), , , , , , , , , , , , , , , , , , , , , , , , , Posted) = False Then
                    GoTo ErrTrap
                End If
     
                LngDevNO = 1

                If ModAccounts.AddNewDev(LngDevID, LngDevNO, StrTempAccountCode, SngTemp, 0, StrTempDes, general_noteid, , , , Date, user_id, val(Transaction_ID), , , , , , , , , , , , , , , , , val(branch_id), , , , , , , , , , , , , , , , , , , , , , , , , Posted) = False Then
                    GoTo ErrTrap
                End If
    
            ElseIf detect_inventory_work_type = 2 Then
                Account_Code_dynamic = get_store_Account(mStoreId, "Account_Code")

                If Account_Code_dynamic = "" Then
                    MsgBox "·„ Ì „  ÕœÌœ Õ”«»  ··„Œ“Ê‰ «·”·⁄Ì ·Â–« «·„Œ“‰ ›Ì Â–« «·›—⁄      " & mStoreId, vbCritical
                    GoTo ErrTrap
                End If
    
                StrTempAccountCode = Account_Code_dynamic  '„Õ“Ê‰ «·”·⁄Ì ··„Œ“‰

                ' StrTempAccountCode = "a1a2a5" '„Õ“Ê‰ «·»÷«⁄…
                If SystemOptions.UserInterface = ArabicInterface Then
                    StrTempDes = "√–‰  ÕÊÌ· »Ì‰ «·„Œ«“‰   —ﬁ„ " & mNoteSerial1
                Else
                    StrTempDes = " Moving Items Vchr  No. " & mNoteSerial1
                End If
    
                LngDevNO = 1

                If ModAccounts.AddNewDev(LngDevID, LngDevNO, StrTempAccountCode, SngTemp, 1, StrTempDes, general_noteid, , , , Date, user_id, val(Transaction_ID), , , , , , , , , , , , , , , , , val(mBranchID), , , , , , , , , , , , , , , , , , , , , , , , , Posted) = False Then
                    GoTo ErrTrap
                End If

                '«·„Œ“Ê‰ «·”·⁄Ì ⁄·Ï „” ÊÏ «·„Œ“‰
    
                Account_Code_dynamic = get_store_Account(CInt(mStoreId2), "Account_Code")

                If Account_Code_dynamic = "" Then
                    MsgBox "·„ Ì „  ÕœÌœ Õ”«»  ··„Œ“Ê‰ «·”·⁄Ì ·Â–« «·„Œ“‰ ›Ì Â–« «·›—⁄    " & mStoreId2, vbCritical
                    GoTo ErrTrap
                End If
    
                StrTempAccountCode = Account_Code_dynamic  '„Õ“Ê‰ «·”·⁄Ì ··„Œ“‰

                ' StrTempAccountCode = "a1a2a5" '„Õ“Ê‰ «·»÷«⁄…
                If SystemOptions.UserInterface = ArabicInterface Then
                    StrTempDes = "√–‰  ÕÊÌ· »Ì‰ «·„Œ«“‰   —ﬁ„ " & mNoteSerial1
                Else
                    StrTempDes = " Moving Items Vchr  No. " & mNoteSerial1
                End If
    
                LngDevNO = 0

                If ModAccounts.AddNewDev(LngDevID, LngDevNO, StrTempAccountCode, SngTemp, 0, StrTempDes, general_noteid, , , , Date, user_id, val(Transaction_ID), , , , , , , , , , , , , , , , , val(mBranchID), , , , , , , , , , , , , , , , , , , , , , , , , Posted) = False Then
                    GoTo ErrTrap
                End If
                Dim BranchId1  As Integer
                Dim BranchID2  As Integer
                Dim DeptSide1 As String
                Dim CreditSide1 As String
                Dim noteid1 As Double
                
                BranchID2 = GetInventoryBranch(CInt(mStoreId2))
                BranchId1 = GetInventoryBranch(CInt(mStoreId))
             
    LngDevNO = 1
If BranchId1 <> BranchID2 Then

 DeptSide1 = getBranchCurrentAccount(BranchId1)
CreditSide1 = getBranchCurrentAccount(BranchID2)
LngDevNO = LngDevNO + 1
If CreditSide1 <> "" Then
     If ModAccounts.AddNewDev(LngDevID, LngDevNO, CreditSide1, SngTemp, 0, StrTempDes, general_noteid, , , , Date, user_id, val(Transaction_ID), , , , , , , , , , , , , , , , , BranchId1, , , , , , , , , , , , , , , , , , , , , , , , , Posted) = False Then
                    GoTo ErrTrap
                End If
End If
                LngDevNO = LngDevNO + 1
                If DeptSide1 <> "" Then
     If ModAccounts.AddNewDev(LngDevID, LngDevNO, DeptSide1, SngTemp, 1, StrTempDes, general_noteid, , , , Date, user_id, val(Transaction_ID), , , , , , , , , , , , , , , , , BranchID2, , , , , , , , , , , , , , , , , , , , , , , , , Posted) = False Then
                    GoTo ErrTrap
                End If
                End If
                noteid1 = val(general_noteid)
                updateNotesValueAndNobytext noteid1, CDbl(SngTemp)
                
                
End If
            ElseIf detect_inventory_work_type = 3 Then
                Dim groupAccount As String
             
                Dim line_value As Single
                

                With mGrid

                        i = Row

                        If mGrid.TextMatrix(Row, mGrid.ColIndex("ItemID")) <> "" Then
    
                            ' groupAccount = get_item_group_account(mGrid.TextMatrix(i, mGrid.ColIndex("Code")), DCboStoreName.BoundText, 2)
                            groupAccount = get_item_group_account_inventory(mGrid.TextMatrix(i, mGrid.ColIndex("ItemID")), mStoreId, 0)

                            If groupAccount = "Error" Then
                                If SystemOptions.UserInterface = ArabicInterface Then
                                    MsgBox "«·’‰› ›Ì «·”ÿ— —ﬁ„" & i & " €Ì— „Õœœ —ﬁ„ Õ”«»  «·„Œ“Ê‰ «·”·⁄Ì ··„Œ“‰ «·„Õœœ   ·„Ã„Ê⁄ …" & mStoreId
                                Else
                                    MsgBox "Item in line no " & i & "Group Name Account Not Defined" & mStoreId
                                End If

                                GoTo ErrTrap
                            End If

                            line_value = mGrid.TextMatrix(i, mGrid.ColIndex("Price")) * mGrid.TextMatrix(i, mGrid.ColIndex("Quantity"))
    
                            If SystemOptions.UserInterface = ArabicInterface Then
                                StrTempDes = "√–‰  ÕÊÌ· Ì÷«∆⁄ »Ì‰ «·„Œ«“‰  —ﬁ„ " & mNoteSerial1
                            Else
                                StrTempDes = "moving items   No. " & mNoteSerial1
                            End If

                            LngDevNO = LngDevNO + 1

                            If ModAccounts.AddNewDev(LngDevID, LngDevNO, groupAccount, line_value, 1, StrTempDes, general_noteid, , , , Date, user_id, val(Transaction_ID), , , , , , , , , , , , , , , , , val(BranchId1), , , , , , , , , , , , , , , , , , , , , , , , , Posted) = False Then
                                GoTo ErrTrap
                            End If
    
                        End If



                End With
 
                With mGrid

                        i = Row

                        If mGrid.TextMatrix(i, mGrid.ColIndex("ItemID")) <> "" Then
    
                            ' groupAccount = get_item_group_account(mGrid.TextMatrix(i, mGrid.ColIndex("Code")), DCboStoreName.BoundText, 2)
                            groupAccount = get_item_group_account_inventory(mGrid.TextMatrix(i, mGrid.ColIndex("ItemID")), CInt(mStoreId2), 0)

                            If groupAccount = "Error" Then
                                If SystemOptions.UserInterface = ArabicInterface Then
                                    MsgBox "«·’‰› ›Ì «·”ÿ— —ﬁ„" & i & " €Ì— „Õœœ —ﬁ„ Õ”«»  «·„Œ“Ê‰ «·”⁄·⁄Ì ··„Œ“‰ «·„Õœœ   ·„Ã„Ê⁄ …" & mStoreId2
                                Else
                                    MsgBox "Item in line no " & i & "Group Name Account Not Defined" & mStoreId2
                                End If

                                GoTo ErrTrap
                            End If

                            line_value = mGrid.TextMatrix(i, mGrid.ColIndex("Price")) * mGrid.TextMatrix(i, mGrid.ColIndex("Quantity"))
    
                            If SystemOptions.UserInterface = ArabicInterface Then
                                StrTempDes = "√–‰  ÕÊÌ·   —ﬁ„ " & mNoteSerial1
                            Else
                                StrTempDes = " Moving Items No. " & mNoteSerial1
                            End If

                            LngDevNO = LngDevNO + 1

                            If ModAccounts.AddNewDev(LngDevID, LngDevNO, groupAccount, line_value, 0, StrTempDes, general_noteid, , , , Date, user_id, val(Transaction_ID), , , , , , , , , , , , , , , , , val(BranchID2), , , , , , , , , , , , , , , , , , , , , , , , , Posted) = False Then
                                GoTo ErrTrap
                            End If
    
                        End If



                End With

            End If

            '----------------
            'LngDevID = LngDevID + 1
            'LngDevNO = 0
        End If

 
 
TransBegine = False

    

'    If detect_inventory_work_type = 3 Then
'
'        With mGrid
'
'            For i = 1 To mGrid.Rows - 1
'
'                If mGrid.TextMatrix(i, mGrid.ColIndex("Code")) <> "" Then
'
'                    ' groupAccount = get_item_group_account(mGrid.TextMatrix(i, mGrid.ColIndex("Code")), DCboStoreName.BoundText, 2)
'                    groupAccount = get_item_group_account_inventory(mGrid.TextMatrix(i, mGrid.ColIndex("Code")), val(DCboStoreName.BoundText), 0)
'
'                    If groupAccount = "Error" Then
'                        If SystemOptions.UserInterface = ArabicInterface Then
'                            MsgBox "«·’‰› ›Ì «·”ÿ— —ﬁ„" & i & " €Ì— „Õœœ —ﬁ„ Õ”«»  «·„Œ“Ê‰ «·”⁄·⁄Ì ··„Œ“‰ «·„Õœœ   ·„Ã„Ê⁄ …"
'                        Else
'                            MsgBox "Item in line no " & i & "Group Name Account Not Defined"
'                        End If
'
'                        Exit Sub
'                    End If
'                End If
'
'            Next i
'
'        End With
'
'    End If


 
    

      '  If Trim(CurrentVoucherNo) <> "" And DateChanged <> True Then
      '      TxtNoteSerialV = CurrentVoucherNo '—ﬁ„ «·ﬁÌœ
      '      TxtNoteSerial1V = Trim(CurrentVoucherSerialNo)
      '  End If

        
  
'        Text1.text = Transaction_ID
'
'        If SystemOptions.TypicalProduction = True Then
'            Exit Sub
'        End If
'
'        'Create big notes
' Dim NoteID As Long
'  Dim NoteDate As Date
'    Dim NoteSerial As String
'    Dim Notevalue As Double
'    Dim des As String
'If CurrentVoucherNo <> "" Then
'NoteSerial = CurrentVoucherNo
'End If
'
'
''*****************************************************************
'    Dim TOTAL_COST As Double
'    With mGrid
'
'        For i = 1 To mGrid.Rows - 1
'
'            If mGrid.TextMatrix(i, mGrid.ColIndex("Code")) <> "" And val(mGrid.TextMatrix(i, mGrid.ColIndex("ItemType"))) <> 1 Then
'                LngCurItemID = val(mGrid.TextMatrix(i, mGrid.ColIndex("Code")))
'                LngUnitID = val(mGrid.Cell(flexcpData, i, mGrid.ColIndex("UnitID")))
'
'                GetUnitNoOfItems LngCurItemID, LngUnitID, UnitFactor
'
'                '           TOTAL_COST = TOTAL_COST + (mGrid.TextMatrix(i, mGrid.ColIndex("Count")) * ModItemCostPrice.GetCostItemPrice(mGrid.TextMatrix(i, mGrid.ColIndex("Code")), 0, , , SystemOptions.SysMainStockCostMethod, , , XPDtbBill.value, , LngUnitID))
'                        'CostPrice
'                TOTAL_COST = TOTAL_COST + val(mGrid.TextMatrix(i, mGrid.ColIndex("ItemCostPrice"))) * mGrid.TextMatrix(i, mGrid.ColIndex("Count"))
'            End If
'
'        Next i
'
'    End With
'    '*****************************************************************
'
' CreateNotes NoteID, (XPDtbBill.value), val(Dcbranch.BoundText), 180, TOTAL_COST, NoteSerial, TxtNoteSerial1V, "Transactions", "Transaction_ID", Transaction_ID, TxtNoteSerial1V, ToHijriDate(XPDtbBill.value)
'          ' TxtNoteID.text = NoteID
'           general_noteid = NoteID
'
'        CREATE_VOUCHER_GE Transaction_ID, TxtNoteSerialV, TxtNoteSerial1V, general_noteid, val(Me.Dcbranch.BoundText)
'
'    End If
'
    '
' End If

s = "Update Transaction_Details Set RequestTypeNo = null, TransferMoveID = " & mmTransaction_ID & "  Where Transaction_ID = " & val(mGrid.TextMatrix(RowNum, mGrid.ColIndex("Transaction_ID"))) & " and [Id] = " & mmID
Cn.Execute s

GetDataStoreQty
MsgBox " „ «‰‘«¡ ”‰œ«  «· ÕÊÌ·« "
Exit Sub
ErrTrap:
errortrap:
    If TransBegine = True Then
        TransBegine = False
        Cn.RollbackTrans
    End If

End Sub



Public Sub CreatePurchOrder(ByVal Row As Long, ByVal mGrid As vsFlexGrid)

    On Error GoTo errortrap
    'DeleteTransactiomsVoucher Val(Text1.text)




    Dim i As Long
    Dim LngCurItemID As Double
    Dim LngUnitID As Long
    Dim UnitFactor As Double
    

    Dim RsTrans As ADODB.Recordset
    Dim rsTrans2 As ADODB.Recordset
    Dim rsTransDet As ADODB.Recordset
    Dim s As String
    Dim mOrderId As Long
    
    Dim sql As String
    Dim MYWAER As String
    Dim StrSQL As String
    Dim RsNotes As ADODB.Recordset
    Dim MYinvnum As String
    Dim note_id As Long

    Dim RSTransDetails As ADODB.Recordset
    Dim RsTemp As New ADODB.Recordset
    Dim RowNum As Integer
    Dim StrSqlDel As String
    Dim SearchResault As Integer
    'Dim Note_ID As Long
    Dim RsDetalis  As ADODB.Recordset
    Dim BeginTrans As Boolean
    Dim LnItemID As Long
 
    Dim StrCurrentItemName As String
    Dim DblNotesTotal As Double

    Dim IntLineNO As Integer
    Dim StrAccountCode As String
    '  Dim RowNum As Integer
    Dim Frm As Form
    Dim Msg As String
    Dim MYTEXT As String
    '›Ì Õ«·… «·«‰ «Ã «·‰„ÿÌ
    Dim TxtNoteSerialV As String
    Dim mDate As Date
    Dim mStoreId As Integer
    Dim mUserID As Long
    Dim xyeas As Boolean
    Dim Transaction_ID As Long
    
    Dim general_noteid As Long
    Dim RsNotesGeneral As ADODB.Recordset
     Dim rs As New ADODB.Recordset
    
    Dim mNoteSerial1 As String
    Dim TxtNoteSerial1V As String
    Dim mTotalCost As Double
    Dim CurrentVoucherNo As String
    
    
        
    
    
    
    
    Dim RsTest As New ADODB.Recordset
    Dim RsRepeat As ADODB.Recordset
    
    
    
    
    Dim NoteID As Long
    
    Dim LngItemID As Long
    Dim Posted As Integer
    Dim Transaction_Type As Integer
    Dim Transaction_Type2 As Integer
    Dim mStoreId2 As Long
    Dim mBranchID As Long
    Dim mEmp_ID As Long
    Dim mNoteId As Long
    Dim mmID  As Long
   On Error GoTo ErrTrap
   Screen.MousePointer = vbArrowHourglass
         Transaction_Type = 38
         'Transaction_Type2 = 11
         Posted = 0

    

ll:
Dim TransBegine As Boolean
Dim mmTransaction_ID  As Long
Dim Sanad_No As Integer
Transaction_Type = 38
Sanad_No = 38

Dim mPrice As Double

Cn.BeginTrans
TransBegine = True
 
 '   mOrderId = val(mGrid.TextMatrix(i, mGrid.ColIndex("Transaction_ID")))
    s = " select * from Transactions WHERE Transaction_ID =  " & mOrderId
    Set RsTrans = New ADODB.Recordset
    RsTrans.Open s, Cn, adOpenKeyset, adLockOptimistic
    rs.Open s, Cn, adOpenKeyset, adLockOptimistic
        
        
         
      
        MYTEXT = CStr(new_id("Transactions", "Transaction_Serial", "", True, "Transaction_Type=38"))
        my_branch = val(mGrid.TextMatrix(Row, mGrid.ColIndex("BranchID")))
        Dim mNoteSerial As String
            
        
        mDate = Date
        TxtNoteSerialV = Notes_coding(val(my_branch), mDate)
        mStoreId = val(mGrid.TextMatrix(Row, mGrid.ColIndex("StoreIDAvi")))
        mBranchID = val(mGrid.TextMatrix(Row, mGrid.ColIndex("BranchID")))
        mStoreId2 = val(mGrid.TextMatrix(Row, mGrid.ColIndex("StoreId")))
        mEmp_ID = val(mGrid.TextMatrix(Row, mGrid.ColIndex("Emp_ID")))
        mmID = val(mGrid.TextMatrix(Row, mGrid.ColIndex("mmID")))
        mUserID = user_id
        mNoteSerial1 = Voucher_coding(val(my_branch), Date, Sanad_No, 170, , Transaction_Type, , mStoreId)
        
        Transaction_ID = CStr(new_id("Transactions", "Transaction_ID", "", True))
        TxtNoteSerial1V = Voucher_coding(val(my_branch), Date, Sanad_No, 170, , Transaction_Type, , mStoreId)
        'mNoteSerial = Notes_coding(val(branch_id), Date)
        CurrentVoucherNo = ""
        
        
       
        
                   rs.AddNew
            rs("Transaction_ID").value = CStr(new_id("Transactions", "Transaction_ID", "", True))
             mmTransaction_ID = val(rs("Transaction_ID").value & "")
           ' rs.update
          
            'Me.oldtxtNoteSerial1.Text = Trim$(Me.TxtNoteSerial1.Text)


  
  
        Screen.MousePointer = vbArrowHourglass
    '     rs("NoteSerial").value = IIf(Trim(Me.TxtNoteSerial.text) = "", Null, Trim(Me.TxtNoteSerial.text))
        rs("NoteSerial1").value = mNoteSerial1
        rs("OldNoteSerial1").value = mNoteSerial1
        rs("Shipping_Pos") = 0
         rs("OldOpOrderID").value = Null
         rs("shipped").value = 0
         rs("InternalFlag").value = 0
         
        rs("purchaseType").value = 0


rs("BillBasedOn").value = 0

rs("OPrType").value = 0
rs("OrderType").value = 3
rs("ContactTime").value = FormatDateTime(Date, vbShortTime)
    'rs("CusID").value = IIf(DBCboClientName.BoundText = "", Null, val(DBCboClientName.BoundText))
    
        
   
mStoreId = mStoreId2
        rs("Currency_id").value = 1
   
 
         
         
'        rs("NoteId").value = val(TxtNoteID.text)
        rs("Transaction_Serial").value = mNoteSerial1
        rs("Transaction_Date").value = Date
        rs("Transaction_Type").value = Transaction_Type
        rs("OrderID").value = 0
        rs("UserID").value = user_id
        
        rs("StoreID").value = mStoreId
'        rs("StoreID1").value = mStoreId2
        rs("BranchId").value = mBranchID
        rs("TaxFound").value = 0
        rs("total").value = 0
        Dim FromstoreAr As String
         Dim FromstoreEn As String
         
       Dim TostoreAr As String
         Dim TostoreEn As String
             
            getStorenames val(mStoreId), FromstoreAr, FromstoreEn
             getStorenames val(mStoreId2), TostoreAr, TostoreEn
         
      
              rs("CusID").value = rs("CusID").value = val(mGrid.TextMatrix(Row, mGrid.ColIndex("CusID")))
        rs("Emp_ID").value = mEmp_ID

'    If Trim$(Me.TxtCashCustomerName.Text) <> "" Then
'        rs("CashCustomerName").value = Trim$(Me.TxtCashCustomerName.Text)
'    Else
        rs("CashCustomerName").value = Null
'    End If
' rs("TransactionComment").value = IIf(Trim$(TxtBillComment.Text) = "", Null, Trim$(TxtBillComment.Text))
'  rs("InspectionReport").value = IIf(Trim$(TxtInspectionReport.Text) = "", Null, Trim$(TxtInspectionReport.Text))
  
  rs("BillBasedOn").value = 0
 rs("order_no").value = 0
   
   
  's("Phone").value = IIf(Trim$(TxtPhone.Text) = "", Null, Trim$(TxtPhone.Text))
  
  
   
        rs.update


        RowNum = Row

            If mGrid.TextMatrix(RowNum, mGrid.ColIndex("ItemID")) <> "" Then

                'Check Repeat Serial
              
       Set RSTransDetails = New ADODB.Recordset
     '   RSTransDetails.Open "[Transaction_Details]", Cn, adOpenStatic, adLockOptimistic, adCmdTable
 StrSQL = "SELECT     dbo.Transaction_Details.* from dbo.Transaction_Details Where (Transaction_ID = -1)"
   RSTransDetails.Open StrSQL, Cn, adOpenKeyset, adLockOptimistic, adCmdText
   
                s = "SELECT tiu.UnitPurPrice FROM TblItemsUnits AS tiu Where ItemId = " & val((mGrid.TextMatrix(RowNum, mGrid.ColIndex("ItemID")))) & " And UnitID = " & val(mGrid.TextMatrix(RowNum, mGrid.ColIndex("UnitID")))
                Dim rsDummy As New ADODB.Recordset
                rsDummy.Open s, Cn, adOpenStatic, adLockReadOnly
                If Not rsDummy.EOF Then
                    mPrice = val(rsDummy!UnitPurPrice & "")
                    mGrid.TextMatrix(RowNum, mGrid.ColIndex("Price")) = mPrice
                End If
                
                RSTransDetails.AddNew
                RSTransDetails("FromstoreAr").value = FromstoreAr
                RSTransDetails("TostoreAr").value = TostoreAr
                RSTransDetails("FromstoreEn").value = FromstoreEn
                RSTransDetails("TostoreEn").value = TostoreEn
                'STransDetails("IsExpirDate").value = IIf(mGrid.TextMatrix(RowNum, mGrid.ColIndex("IsExpirDate")) = "", Null, val(mGrid.TextMatrix(RowNum, mGrid.ColIndex("IsExpirDate"))))
                RSTransDetails("OriginalQty").value = IIf((mGrid.TextMatrix(RowNum, mGrid.ColIndex("Quantity")) = ""), Null, val(mGrid.TextMatrix(RowNum, mGrid.ColIndex("Quantity"))))
                'STransDetails("order_no").value = IIf(mGrid.TextMatrix(RowNum, mGrid.ColIndex("order_no")) = "", Null, mGrid.TextMatrix(RowNum, mGrid.ColIndex("order_no")))
                RSTransDetails("OrderArrivalDate").value = Date '.TextMatrix(RowNum, mGrid.ColIndex("OrderArrivalDate")) = "", Null, mGrid.TextMatrix(RowNum, mGrid.ColIndex("OrderArrivalDate")))
                'STransDetails("FoxyNo").value = IIf((mGrid.TextMatrix(RowNum, mGrid.ColIndex("FoxyNo")) = ""), Null, mGrid.TextMatrix(RowNum, mGrid.ColIndex("FoxyNo")))
                RSTransDetails("Transaction_ID").value = val(Transaction_ID)
                RSTransDetails("Item_ID").value = IIf((mGrid.TextMatrix(RowNum, mGrid.ColIndex("ItemID")) = ""), Null, val(mGrid.TextMatrix(RowNum, mGrid.ColIndex("ItemID"))))
                'RSTransDetails("Remarks").value = IIf((mGrid.TextMatrix(RowNum, mGrid.ColIndex("Remarks")) = ""), Null, Trim$(mGrid.TextMatrix(RowNum, mGrid.ColIndex("Remarks"))))
                If Not mGrid.TextMatrix(RowNum, mGrid.ColIndex("ItemName")) = "" Then
                    StrSQL = "select * From TblItems where ItemID=" & mGrid.TextMatrix(RowNum, mGrid.ColIndex("ItemID"))
                    RsTemp.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText

                    If Not (RsTemp.EOF Or RsTemp.BOF) Then
                        If RsTemp("HaveSerial").value = True Then
                            'RSTransDetails("ItemSerial").value = IIf((mGrid.TextMatrix(RowNum, mGrid.ColIndex("Serial")) = ""), Null, mGrid.TextMatrix(RowNum, mGrid.ColIndex("Serial")))
                        End If
                    End If

                    RsTemp.Close
                End If
                RSTransDetails("ItemBalance").value = IIf((mGrid.TextMatrix(RowNum, mGrid.ColIndex("ItemBalance2")) = ""), Null, val(mGrid.TextMatrix(RowNum, mGrid.ColIndex("ItemBalance2"))))
                RSTransDetails("ShowQty").value = IIf((mGrid.TextMatrix(RowNum, mGrid.ColIndex("Quantity")) = ""), Null, val(mGrid.TextMatrix(RowNum, mGrid.ColIndex("Quantity"))))
                RSTransDetails("showPrice").value = mPrice 'IIf((mGrid.TextMatrix(RowNum, mGrid.ColIndex("Price")) = ""), Null, val(mGrid.TextMatrix(RowNum, mGrid.ColIndex("Price"))))
                RSTransDetails("Quantity").value = IIf((mGrid.TextMatrix(RowNum, mGrid.ColIndex("Quantity")) = ""), Null, val(mGrid.TextMatrix(RowNum, mGrid.ColIndex("Quantity"))))
            
           '     RSTransDetails("ItemDiscountType").value = IIf((mGrid.TextMatrix(RowNum, mGrid.ColIndex("DiscountType")) = ""), Null, val(mGrid.TextMatrix(RowNum, mGrid.ColIndex("DiscountType"))))
           '     RSTransDetails("ItemCase").value = IIf((mGrid.TextMatrix(RowNum, mGrid.ColIndex("ItemCase")) = ""), Null, val(mGrid.TextMatrix(RowNum, mGrid.ColIndex("ItemCase"))))
           '     RSTransDetails("ItemDiscount").value = IIf((mGrid.TextMatrix(RowNum, mGrid.ColIndex("DiscountVal")) = ""), Null, val(mGrid.TextMatrix(RowNum, mGrid.ColIndex("DiscountVal"))))
           '     RSTransDetails("guaranteeTime").value = IIf((mGrid.TextMatrix(RowNum, mGrid.ColIndex("guaranteeTime")) = ""), Null, val(mGrid.TextMatrix(RowNum, mGrid.ColIndex("guaranteeTime"))))
            
                RSTransDetails("ColorID").value = 1 'IIf((mGrid.TextMatrix(RowNum, mGrid.ColIndex("ColorID")) = ""), 1, val(mGrid.TextMatrix(RowNum, mGrid.ColIndex("ColorID"))))
                RSTransDetails("ItemSize").value = 1 'IIf((mGrid.TextMatrix(RowNum, mGrid.ColIndex("ItemSize")) = ""), 1, Trim$(mGrid.TextMatrix(RowNum, mGrid.ColIndex("ItemSize"))))
                RSTransDetails("ClassId").value = 1 'IIf((mGrid.TextMatrix(RowNum, mGrid.ColIndex("ClassId")) = ""), 1, val(mGrid.TextMatrix(RowNum, mGrid.ColIndex("ClassId"))))

                RSTransDetails("UnitID").value = val(mGrid.TextMatrix(RowNum, mGrid.ColIndex("UnitID")))

                RSTransDetails("BranchId").value = mBranchID
                RSTransDetails("OrderArrivalDate").value = Date ' IIf(Not IsDate(mGrid.TextMatrix(RowNum, mGrid.ColIndex("OrderArrivalDate"))), Null, mGrid.TextMatrix(RowNum, mGrid.ColIndex("OrderArrivalDate")))
            'RSTransDetails("ItemsDetailsNewidea").value = IIf(mGrid.TextMatrix(RowNum, mGrid.ColIndex("ItemsDetailsNewidea")) = "", Null, mGrid.TextMatrix(RowNum, mGrid.ColIndex("ItemsDetailsNewidea")))

                Dim RsUnitData As ADODB.Recordset
                
                
                Dim DblQty As Double
        
                LngCurItemID = val(mGrid.TextMatrix(RowNum, mGrid.ColIndex("ItemID")))
                LngUnitID = val(mGrid.TextMatrix(RowNum, mGrid.ColIndex("UnitID")))
                DblQty = val(mGrid.TextMatrix(RowNum, mGrid.ColIndex("Quantity")))

                StrSQL = "Select * From TblItemsUnits Where ItemID=" & LngCurItemID
                StrSQL = StrSQL + " AND UnitID=" & LngUnitID
                Set RsUnitData = New ADODB.Recordset
                RsUnitData.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

                If Not (RsUnitData.BOF Or RsUnitData.EOF) Then
                    RSTransDetails("QtyBySmalltUnit").value = RsUnitData("UnitFactor").value
                
                    RSTransDetails("Quantity").value = RSTransDetails("QtyBySmalltUnit").value * RSTransDetails("showqty").value
                    RSTransDetails("Price").value = val(IIf((mGrid.TextMatrix(RowNum, mGrid.ColIndex("Price")) = ""), 0, val(mGrid.TextMatrix(RowNum, mGrid.ColIndex("Price"))))) / RSTransDetails("QtyBySmalltUnit").value
            
                End If
                'RSTransDetails("Height").value = val(mGrid.TextMatrix(RowNum, mGrid.ColIndex("Height")))
                'RSTransDetails("Width").value = val(mGrid.TextMatrix(RowNum, mGrid.ColIndex("Width")))
            
                RSTransDetails("price").value = Round(mGrid.TextMatrix(RowNum, mGrid.ColIndex("Price")) / RSTransDetails("Quantity").value, 2)
                'RSTransDetails("ProductionDate").value = IIf((mGrid.TextMatrix(RowNum, mGrid.ColIndex("ProductionDate")) = ""), Null, Format((mGrid.TextMatrix(RowNum, mGrid.ColIndex("ProductionDate"))), "DD/mm/YYYY"))
                'RSTransDetails("ExpiryDate").value = IIf((mGrid.TextMatrix(RowNum, mGrid.ColIndex("ExpiryDate")) = ""), Null, Format((mGrid.TextMatrix(RowNum, mGrid.ColIndex("ExpiryDate"))), "DD/mm/YYYY"))
                'RSTransDetails("LotNO").value = IIf(mGrid.TextMatrix(RowNum, mGrid.ColIndex("LotNO")) = "", Null, mGrid.TextMatrix(RowNum, mGrid.ColIndex("LotNO")))
             Dim OldQty As Double
             Dim OldCost As Double
              Dim NewQty As Double
               Dim NewCost As Double
               
getItemCostData Date, RSTransDetails("Item_ID").value, val(mStoreId), val(Transaction_ID), OldQty, OldCost, NewQty, NewCost
       RSTransDetails("OldQty").value = NewQty
       RSTransDetails("OldCost").value = NewCost
       
      RSTransDetails("NewQty").value = RSTransDetails("OldQty").value - RSTransDetails("Quantity").value
       RSTransDetails("NewCost").value = RSTransDetails("OldCost").value ' ((RSTransDetails("OldQty").value * RSTransDetails("OldCost").value) + (RSTransDetails("Quantity").value * RSTransDetails("Price").value)) / (RSTransDetails("Quantity").value + RSTransDetails("OldQty").value)
       
       
                RSTransDetails.update
            End If


'SaveItemsData val(Transaction_ID), -1
        '≈÷«›… «·»÷«∆⁄ ≈·Ï «·„Œ“‰ «·ÃœÌœ
        
'SaveItemsData rs("Transaction_ID").value, 1

        Cn.CommitTrans
        BeginTrans = False
    

s = "Update Transaction_Details Set RequestTypeNo = null,PurchaseRequestID = " & mmTransaction_ID & "  Where Transaction_ID = " & val(mGrid.TextMatrix(RowNum, mGrid.ColIndex("Transaction_ID"))) & " and [Id] = " & mmID
Cn.Execute s

GetDataStoreQty
MsgBox " „ «‰‘«¡ ÿ·»«  «·‘—«¡"
Exit Sub
ErrTrap:
errortrap:
    If TransBegine = True Then
        TransBegine = False
        Cn.RollbackTrans
    End If

End Sub






Private Sub Dcbranch_Change(Index As Integer)
    If Index <> 8 Then
    
    If Me.TxtModFlg2(mIndex) <> "R" Then
        TxtNoteSerial1.Text = ""
        TxtNoteSerial.Text = ""
   End If
   End If
End Sub

Private Sub Dcbranch_Click(Index As Integer, Area As Integer)
    If Index <> 8 Then
    If Me.TxtModFlg2(mIndex) <> "R" Then
    TxtNoteSerial1.Text = ""
   TxtNoteSerial.Text = ""
   End If
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

Private Sub Option1_Click(Index As Integer)
    If Me.Option1(2).value = True Then
        Reload 1
    End If

End Sub

Private Sub Option2_Click()
    If Me.Option2.value = True Then
        Reload 2
    End If

End Sub

Private Sub txtDiscPercent_Change()
If mIndex = 8 Then Exit Sub
 If Me.TxtModFlg2(mIndex) = "R" Then Exit Sub
If mDiscEnter Then Exit Sub
If val(txtTotal2) <> 0 Then

    txtDiscValue = val(txtDiscPercent) * val(txtTotal2) / 100

End If

CalcTotal2
mDiscEnter = False
End Sub

Private Sub txtDiscValue_Change()

If mIndex = 8 Then Exit Sub
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
                        FG.TextMatrix(1, FG.ColIndex("Name")) = "„‰ Õ—ﬂ… «·ﬁÿ⁄ «·„ﬁœ—… "
                        FG.TextMatrix(1, FG.ColIndex("Price")) = val(rs2!TotalAfterDiscount & "") + val(rs2!Vat2 & "")
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

Private Sub TxtSearchCode_KeyPress(KeyAscii As Integer)
 Dim CUSTID As Integer

    If KeyAscii = vbKeyReturn Then
        GetCustomersDetail CUSTID, , TxtSearchCode.Text, 1
        DBCboClientName.BoundText = CUSTID
    End If
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
   ElseIf mIndex = 7 Then
        FiLLTXT7
        
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
   ElseIf mIndex = 7 Then
        FiLLTXT7
        
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
   ElseIf mIndex = 7 Then
        FiLLTXT7
        
        
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
   ElseIf mIndex = 7 Then
        FiLLTXT7
        
        
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
    'Label4.Caption = "Remark"
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
    
    Option1(2).Caption = "Customers"

    Label1(10).Caption = "Hand wages"
    'Label5.Caption = "No"
    'Label5.Caption = "No"
     With Me.FG
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
        
    With grdAging
        .TextMatrix(0, .ColIndex("Ser")) = "Ser"
        .TextMatrix(0, .ColIndex("CusName")) = "Cus/Sup Name"
        .TextMatrix(0, .ColIndex("NoteSerial")) = "Note Serial"
        .TextMatrix(0, .ColIndex("DueDate")) = "Date"
        .TextMatrix(0, .ColIndex("TransactionTypeName")) = "Transaction Type"
        .TextMatrix(0, .ColIndex("TransNet")) = "Value"
        .TextMatrix(0, .ColIndex("PayedValue")) = "Payed Value"
        .TextMatrix(0, .ColIndex("StillAmount")) = "Still Amount"
        .TextMatrix(0, .ColIndex("DiffDate")) = "Aging"
        .TextMatrix(0, .ColIndex("Name")) = "Status"
        .TextMatrix(0, .ColIndex("From")) = "From"
        .TextMatrix(0, .ColIndex("TO")) = "TO"
    
    End With
        Frame12.Caption = "Age of debt"
        Frame10.Caption = "Date"
        
        Rd(0).RightToLeft = False
Rd(0).Caption = "Bill Date"
Rd(1).RightToLeft = False
Rd(1).Caption = "Due Date"
Frame7.Caption = "Date By"
Label8.Caption = "To Date"
Label24.Caption = "To Date"
Label25.Caption = "Date Invoice"
Label7.Caption = "Due Date"
BtnPrint22.Caption = "Print"



 'Fra(0).Caption = "Payment Method"
'OptPayType(0).RightToLeft = False
'OptPayType(1).RightToLeft = False
'OptPayType(2).RightToLeft = False
'OptPayType(0).Caption = "Cash"
'OptPayType(1).Caption = "Credit"
'OptPayType(2).Caption = "All"
    C1Tab1.Caption = "AGEING REPORT|Reports VAT"
    Label1(22).Caption = "AGEING REPORT"
    Label1(3).Caption = "VAT Reports "
    'Label2.Caption = "To Date"
    'Label3.Caption = "From Date"
    'Label4.Caption = "Transaction Type"
'    Label9.Caption = "Branch"
'    Label10.Caption = "Store"
'    Label11.Caption = "Item"
    
    Command3.Caption = "Clear"
    Command6.Caption = "Clear"
    btn_Cancel(0).Caption = "Exit"
Label5(3).Caption = "Date"
    lbl(1).Caption = "This screen displays the value added data according to the terms"
    Label1(21).Caption = "Customer Type"
    Label4(1).Caption = "Customer Type"
    Option2.Caption = "Vendor"
    lbl(25).Caption = "Category"
    Label6(3).Caption = "Branch"
    lbl(25).Caption = "Category"
   'Me.Caption = "PAYABLE AGEING REPORT (BY INVOICE)"
    Label1(2).Caption = "PAYABLE AGEING REPORT (BY INVOICE)"
    
    BtnPrint(0).Caption = "Analytical Printing"
    BtnPrint(1).Caption = "Total Printing"
    Command2.Caption = "Clear"
   ' Label7.Caption = "From Date"
   ' Label8.Caption = "To Date"
    
    ChekCustomer.Caption = "Cust/Supp"
    CheckAllCustomer.Caption = "Choose More Cust/Supp"
    CheckEmp.Caption = "Employee"
    CheckAllEMp.RightToLeft = False
    ChekCustomer.RightToLeft = False
    CheckAllCustomer.RightToLeft = False
    CheckEmp.RightToLeft = False
    CheckAllEMp.Caption = "Choose More Employee"
    CmdSelectCus.Caption = "Select >>"
    CmdSelectEmp.Caption = "Select >>"
    Label1(1).Caption = "Al SATTARYAH GROUP"
       
        
        
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

        If TxtNoteSerial17.Text = "" Then
                If Voucher_coding(val(dcBranch(mIndex).BoundText), XPDtbTrans7.value, 7, 23001, , , , , , , "TblCaptinTrans") = "error" Then
                    If SystemOptions.UserInterface = ArabicInterface Then
                        MsgBox " ·« Ì„ﬂ‰ «÷«›… ”‰œ  ÃœÌœ ·«‰ﬂ  ⁄œÌ  «·Õœ «·–Ì ﬁ„  » ÕœÌœ… „‰ «·”‰œ«   ": Exit Sub
                    Else
                        MsgBox " Cant't Create Expenses Voucher to this Process no You exceed the maximum number ": Exit Sub
                    End If

                Else
         
                    If Voucher_coding(val(dcBranch(mIndex).BoundText), XPDtbTrans7.value, 7, 23001, , , , , , , "TblCaptinTrans") = "" Then
                        If SystemOptions.UserInterface = ArabicInterface Then
                            
                            TxtNoteSerial17.locked = False
                            MsgBox " ·«»œ „‰ ﬂ «»… —ﬁ„ «·”‰œ ÌœÊÌ« ﬂ„« Õœœ   ": Exit Sub
                        Else
                            MsgBox "  Enter Voucher No Manually or Define Coding ": Exit Sub
                        End If

                    Else
                        TxtNoteSerial17.Text = Voucher_coding(val(dcBranch(mIndex).BoundText), XPDtbTrans7.value, 7, 23001, , , , , , , "TblCaptinTrans")
                    End If
                End If
            End If
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
    TabMain.TabVisible(4) = False
     TabMain.TabVisible(5) = False
     TabMain.TabVisible(6) = False
     TabMain.TabVisible(7) = False
     TabMain.TabVisible(8) = False
    
          
    If mIndex = 0 Then
        TabMain.TabVisible(0) = True
        TabMain.CurrTab = 0
    ElseIf mIndex = 1 Then
        Me.dcBranch(1).BoundText = branch_id
        TabMain.TabVisible(1) = True
        TabMain.CurrTab = 1
        Me.Caption = "√ÃÊ— «·Ìœ"
      
    ElseIf mIndex = 2 Then
        Me.Width = GRID2.Width + 400
        TabMain.TabVisible(mIndex) = True
        TabMain.CurrTab = mIndex
       ' Me.Width = Grid.Width + 400
    ScreenNameArabic = "«‰Ê«⁄ „ﬂ« » «· ›ÊÌ÷"
     
    ElseIf mIndex = 8 Then
        'Me.Width = Grid2.Width + 400
        TabMain.TabVisible(mIndex) = True
        TabMain.CurrTab = mIndex
        
       ' Me.Width = Grid.Width + 400
        ScreenNameArabic = "√⁄„«— «·œÌÊ‰"
        
           DTP_Date.value = Date
            Me.Caption = ScreenNameArabic
            todate.value = Date
            FromDate.value = Date
            ToDate1.value = Date
            FromDate1.value = Date
            todate.value = Date

            
    If SystemOptions.UserInterface = ArabicInterface Then
        My_SQL = "  select  ID,Name  from ClassCustomers  "
    Else
        My_SQL = "  select  ID,Namee  from ClassCustomers  "
    End If

    fill_combo dcClass, My_SQL


       
            
            
            FromDate.value = ""
            todate.value = ""
            FromDate1.value = ""
            ToDate1.value = ""
            
            DBCboClientName.Enabled = False
            CmdSelectCus.Enabled = False
            DcbEmployee.Enabled = False
            CmdSelectEmp.Enabled = False
        
        
            ScreenNameArabic = "  ﬁ—Ì— «⁄„«— «·œÌÊ‰ ⁄·Ï «·⁄„·«¡ Ê «·„Ê—œÌ‰  "
            ScreenNameEnglish = "  Agenig Report"
            RegisterLogInOut Me.Name, ScreenNameArabic, ScreenNameEnglish, "1"
            
            
            Set Dcombos = New ClsDataCombos
            ' Dcombos.GetEmployees Me.DCmboEmp, True
            Set cSearchDCombo = New clsDCboSearch
            ' Set cSearchDCombo.Client = DCmboEmp
            
            Dcombos.GetSalesRepData Me.DcbEmployee
            Dcombos.GetBranches Me.dcBranch(mIndex)
    

            Dcombos.GetCustomerType Me.DcCustomerType



            
            
  
    
    
    
    
    


  
    
    
    
    
   
    
    Resize_Form Me
         
     
    ElseIf mIndex = 7 Then
        Me.Width = grdExcel.Width + 400
        TabMain.TabVisible(mIndex) = True
        TabMain.CurrTab = mIndex
        
       ' Me.Width = Grid.Width + 400
    ScreenNameArabic = "«·ﬂ»« ‰"
        Set Dcombos = New ClsDataCombos
     Dcombos.GetBanks Me.DcboBankName
  
    Dcombos.GetUsers Me.DCboUserName(7)
    
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
    
    Dcombos.GetCustomersSuppliers 1, DcCustmer(1), , , 1
    Dcombos.GetTblCarsDataGroup Me.DcbCarType
    Dcombos.GetBranches Me.dcBranch(1)
    Dcombos.GetBranches Me.dcBranch(7)
    
    
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
      
'Me.Dcbranch(1).BoundText = branch_id
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
        
    ElseIf mIndex = 7 Then
        My_SQL = "TblCaptinTrans"
       ' Set BKGrndPic = New ClsBackGroundPic
        Set RsSavRec = New ADODB.Recordset
        RsSavRec.CursorLocation = adUseClient
        RsSavRec.Open My_SQL, Cn, adOpenKeyset, adLockPessimistic, adCmdTableDirect
        TxtModFlg2(mIndex).Text = "R"
       ' DCboUserName(mIndex).BoundText = user_id
       

                DCboUserName(mIndex).BoundText = user_id

        btn_First_Click (mIndex)
        Me.Caption = "›Ê« Ì— «·„»Ì⁄« "
        
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
    
    
    With Me.cboMasterType
        .Clear

     .AddItem "NA"
            .AddItem "Frames"
            .AddItem "Optical Lens"
            .AddItem "Contact Lens"

            .AddItem "Lens Care Product"

            .AddItem "Accessories"

            
       

    End With


    
            
         Dcombos.GetItemSGroups cmbEyeDet(8), False
           'Dcombos.GetItemSGroups cmbEyeDet(9), False
          Dcombos.GetItemsColors cmbEyeDet(9)
        Dim str As String
        Dim mm As Long
        For mm = 0 To cmbEyeDet.count - 1
              If mm <> 8 And mm <> 9 Then
                  
                  If mm = 22 Then
                  
                      str = " SELECT     ID, SPHT as Name"
                  ElseIf mm = 23 Then
                       str = " SELECT     ID, CLYT as Name"
                  Else
                      If SystemOptions.UserInterface = ArabicInterface Then
                          str = " SELECT     ID, Namee"
                      Else
                          str = " SELECT     ID, Namee"
                      End If
                  End If
                  str = str & "                   From " & GetTableName(mm)
                  If mm <> 7 Then
                  fill_combo cmbEyeDet(mm), str
                  End If
              End If
        Next
        


        
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
    ElseIf mIndex = 4 Then
        Dcombos.GetStores cmbStoreID
        TabMain.TabVisible(4) = True
        
        Me.Caption = " ‰»ÌÂ«  «·ÿ·»«  «·œ«Œ·Ì… ( ÕÊÌ· - ‘—«¡) "
        
        
        Dim dstore As Integer
        Dim dBox As Integer
        Dim usertype As Integer
        Dim EmpID As Integer
        Dim userbranchid As Integer
        Dim CUSTID As Integer
        Dim dStore2 As Integer
        'GetBranchData branch_id, dstore, dBox

        GetUserData user_id, usertype, userbranchid, dstore, dBox, , EmpID, , CUSTID, dStore2
        'intDef
            
           
           cmbStoreID.BoundText = dstore
          
'          Me.DcBranch.BoundText = userbranchid
'          Me.DCboStoreName.BoundText = dstore
'          Me.DcboBox.BoundText = dBox
'          Me.DcboEmp.BoundText = EmpID
'          Me.DCboStore2Name.BoundText = dStore2

        
        GetDataStoreQty
    'Store_ID
    ElseIf mIndex = 6 Then
        Dcombos.GetStores cmbStoreID2
        TabMain.TabVisible(6) = True
        
        Me.Caption = " ‰»ÌÂ«  «·„⁄„·"
        
        
        
        
        'GetBranchData branch_id, dstore, dBox

       ' GetUserData user_id, usertype, userbranchid, dstore, dBox, , EmpID, , CUSTID, dStore2
        
            Set rsDummy = New ADODB.Recordset
        s = "Select * from tblStore WHERE BranchId = " & branch_id & " and IsNull(IsLab,0) = 1"
        Set rsDummy = New ADODB.Recordset
        rsDummy.Open s, Cn, adOpenStatic, adLockReadOnly
        If Not rsDummy.EOF Then
            cmbStoreID2.BoundText = val(rsDummy!StoreId & "")
            
        End If
    


        'intDef
            
           
       
          
'          Me.DcBranch.BoundText = userbranchid
'          Me.DCboStoreName.BoundText = dstore
'          Me.DcboBox.BoundText = dBox
'          Me.DcboEmp.BoundText = EmpID
'          Me.DCboStore2Name.BoundText = dStore2

        
        GetDataStoreQty2
    'Store_ID
    
    End If
    ShowTip
TabMain.CurrTab = mIndex
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

Public Function GetDataStoreQty()
Dim s As String
  s = " SELECT td.Quantity,"
  s = s & "        td.Item_ID ItemID,td.ID  as mmID,"
    s = s & "        td.UnitId,t.Emp_Id,td.ShowPrice as price,"
    s = s & "        td.StoreIDAvi,t.NoteSerial1,"
    s = s & "        td.RequestTypeNo,"
    s = s & "        td.ItemBalance2,"
    s = s & "        RequestTypeName = CASE RequestTypeNo WHEN 1 THEN 'ÿ·» ‘—«¡' WHEN 2 THEN ' ÕÊÌ· „Œ“‰Ï' END,"
    s = s & "           IsNull(td.StoreIDAvi2,t.StoreID) StoreID,"
    s = s & "        t.BranchId,"
    s = s & "        TblBranchesData.branch_name AS BranchName,t.Transaction_ID,t.CusID,"
    s = s & "        TblItems.ItemName,TblUnites.UnitName,"
    s = s & "        TblStore.StoreName,"
    s = s & "               TblStore2.StoreName      AS StoreAviName"
    s = s & "        FROM   Transaction_Details      AS td"
    s = s & "               INNER JOIN Transactions  AS t"
    s = s & "                    ON  t.Transaction_ID = td.Transaction_ID"
    s = s & "               LEFT OUTER JOIN TblStore"
    s = s & "                    ON  TblStore.StoreID = td.StoreIDAvi2"
    s = s & "               LEFT OUTER JOIN TblStore TblStore2"
    s = s & "                    ON  TblStore2.StoreID = td.StoreIDAvi"
    s = s & "               LEFT OUTER JOIN TblItems"
    s = s & "                    ON  TblItems.ItemID = td.Item_ID"
    s = s & "               LEFT OUTER JOIN TblUnites"
    s = s & "                    ON  TblUnites.UnitId = td.UnitId"
    s = s & "               LEFT OUTER JOIN TblBranchesData"
    s = s & "                    ON  TblBranchesData.branch_id = t.BranchId"
    s = s & "        Where "
    If Option1(0).value Then
    
        s = s & "         td.StoreIDAvi2 = " & val(cmbStoreID.BoundText)
        s = s & " and IsNull(td.RequestTypeNo, 0) =1"
    Else
         s = s & "         td.StoreIDAvi = " & val(cmbStoreID.BoundText)
         s = s & " and IsNull(td.RequestTypeNo, 0) =2"
    End If
    
    loadgrid s, grdTransfer, True, False



End Function



Public Function GetDataStoreQty2()
Dim s As String
  s = " SELECT td.Quantity,"
  s = s & " Status =(CASE"
s = s & " when ISNULL(PurchaseRequestID, 0) <> 0 THEN"
s = s & "     ' „ ⁄„· ÿ·» ‘—«¡'"
s = s & "  WHEN ISNULL(TransferMoveID, 0) <> 0 THEN"
s = s & "     ' „ «” ·«„ «·ﬁÿ⁄…'"
s = s & "  WHEN ISNULL(RequestTypeNo, 0) = 2 THEN"
s = s & "     '„ÿ·Ê»… ·· ÕÊÌ·'"
s = s & " WHEN ISNULL(RequestTypeNo, 0) = 1 THEN"
s = s & "     '„ÿ·Ê»… ··‘—«¡'"
s = s & " END),"
           
  s = s & " StatusID =(CASE"
s = s & " when ISNULL(PurchaseRequestID, 0) <> 0 THEN"
s = s & "     '1'"
s = s & "  WHEN ISNULL(TransferMoveID, 0) <> 0 THEN"
s = s & "     '2'"
s = s & "  WHEN ISNULL(RequestTypeNo, 0) = 2 THEN"
s = s & "     '3'"
s = s & " WHEN ISNULL(RequestTypeNo, 0) = 1 THEN"
s = s & "     '4'"
s = s & " END),"
                      
           
  s = s & "        td.Item_ID ItemID,td.ID  as mmID,TblCustemers.CusName as CustomerName ,t.Transaction_Date,"
    s = s & "        td.UnitId,t.Emp_Id,td.ShowPrice as price,"
    s = s & "        td.StoreIDAvi,t.NoteSerial1,"
    s = s & "        td.RequestTypeNo,"
    s = s & "        td.ItemBalance2,"
    s = s & "        RequestTypeName = CASE RequestTypeNo WHEN 1 THEN 'ÿ·» ‘—«¡' WHEN 2 THEN ' ÕÊÌ· „Œ“‰Ï' END,"
    s = s & "        td.StoreIDAvi2 as StoreID,"
    s = s & "        t.BranchId,"
    s = s & "        TblBranchesData.branch_name AS BranchName,t.Transaction_ID,t.CusID,"
    s = s & "        TblItems.ItemName,TblUnites.UnitName,"
    s = s & "        TblStore.StoreName,"
    s = s & "               TblStore2.StoreName      AS StoreAviName"
    s = s & "        FROM   Transaction_Details      AS td"
    s = s & "               INNER JOIN Transactions  AS t"
    s = s & "                    ON  t.Transaction_ID = td.Transaction_ID"
    s = s & "               LEFT OUTER JOIN TblCustemers"
    s = s & "                    ON  TblCustemers.CusID= t.CusID"
    s = s & "               LEFT OUTER JOIN TblStore"
    s = s & "                    ON  TblStore.StoreID = td.StoreIDAvi2"
    s = s & "               LEFT OUTER JOIN TblStore TblStore2"
    s = s & "                    ON  TblStore2.StoreID = td.StoreIDAvi"
    s = s & "               LEFT OUTER JOIN TblItems"
    s = s & "                    ON  TblItems.ItemID = td.Item_ID"
    s = s & "               LEFT OUTER JOIN TblUnites"
    s = s & "                    ON  TblUnites.UnitId = td.UnitId"
    s = s & "               LEFT OUTER JOIN TblBranchesData"
    s = s & "                    ON  TblBranchesData.branch_id = t.BranchId"
    s = s & "        Where  IsNull(IsFinish,0) = 0 and  (IsNull(td.RequestTypeNo, 0) <> 0  Or IsNull(PurchaseRequestID,0) <> 0 Or IsNull(TransferMoveID,0) <> 0)"
    s = s & "        and td.StoreIDAvi2 = " & val(cmbStoreID2.BoundText)
    
    
    loadgrid s, grdTransfer2, True, False
'    for i = 1 to
CheckQtyFromStore grdTransfer2

End Function


Sub GetItemBalanceInStore(Optional LngRow As Long, Optional ColorID As Integer, Optional itemsize As Integer, Optional ClassId As Integer, Optional StrItemSerial As String, Optional LngItemID As Long, Optional TransactionDate As Variant, Optional ByVal mStoreId As Long = 1, Optional ByRef mGrid As vsFlexGrid = Nothing)
   
    Dim RsTest As ADODB.Recordset
    Dim DblOutQty As Double
    Dim LngColorID As Long
    Dim LngClassId As Long
    Dim mItemBalance As Double
    Dim StrItemSize As String
        LngColorID = ColorID
        StrItemSize = itemsize
         LngClassId = ClassId
With mGrid
   
            Set RsTest = GetItemQuantityStock(LngItemID, val(mStoreId), TransactionDate, , , , StrItemSerial, True, LngColorID, StrItemSize, LngClassId)
              If RsTest.EOF Or RsTest.BOF Then
                    .TextMatrix(LngRow, .ColIndex("ItemBalance")) = 0
              Else
              ' (Round(RsTest("totalqty").value, 2)
              
              mItemBalance = IIf(IsNull(RsTest("totalqty").value), 0, RsTest("totalqty").value)
              
                    .TextMatrix(LngRow, .ColIndex("ItemBalance")) = Round(mItemBalance, 2)
              End If
              
         
 End With

End Sub


Private Sub CheckQtyFromStore(mGrid As vsFlexGrid)
Dim StrSQL  As String
Dim mDateTrans As Date
Dim Begin  As Boolean
Dim mItemId As Long
Dim mStoreId As Long
Dim mQuantity As Double
mDateTrans = Date
Dim mItemBalance As Double
Dim i As Long
For i = 1 To mGrid.Rows - 1
    mItemId = val(mGrid.TextMatrix(i, mGrid.ColIndex("ItemID")))
    mStoreId = val(mGrid.TextMatrix(i, mGrid.ColIndex("StoreIDAvi")))
    mItemBalance = val(mGrid.TextMatrix(i, mGrid.ColIndex("ItemBalance")))
    mQuantity = val(mGrid.TextMatrix(i, mGrid.ColIndex("Quantity")))
    GetItemBalanceInStore i, 1, 1, 1, , mItemId, mDateTrans, mStoreId, grdTransfer2
    
    If mItemBalance <= mQuantity Then
          StrSQL = ""
          
        StrSQL = StrSQL & " SELECT "
        StrSQL = StrSQL & "        T.StoreID,"
        StrSQL = StrSQL & "        ts.StoreName,"
        StrSQL = StrSQL & "        ts.storenamee,SUM(T.transqty) as ItemBalance2"
        StrSQL = StrSQL & " FROM   ("
        StrSQL = StrSQL & "            SELECT SUM("
        StrSQL = StrSQL & "                       ("
        StrSQL = StrSQL & "                           dbo.Transaction_Details.Quantity / ISNULL("
        StrSQL = StrSQL & "                               dbo.GetItemUnitFactor(dbo.Transaction_Details.Item_ID, 1),"
        StrSQL = StrSQL & " 1"
        StrSQL = StrSQL & "                           ) * StockEffect"
        StrSQL = StrSQL & "                       )"
        StrSQL = StrSQL & "                   ) AS transqty,"
        StrSQL = StrSQL & "                   StockEffect,"
        StrSQL = StrSQL & "                   transactions.StoreId"
        StrSQL = StrSQL & "            From dbo.transactions"
        StrSQL = StrSQL & "                   INNER JOIN dbo.TransactionTypes"
        StrSQL = StrSQL & "                        ON  dbo.Transactions.Transaction_Type = dbo.TransactionTypes.Transaction_Type"
        StrSQL = StrSQL & "                   INNER JOIN dbo.Transaction_Details"
        StrSQL = StrSQL & "                        ON  dbo.Transactions.Transaction_ID = dbo.Transaction_Details.Transaction_ID"
        StrSQL = StrSQL & "            WHERE  dbo.Transactions.Transaction_Date <= " & SQLDate(mDateTrans, True) & ""
        StrSQL = StrSQL & "                       Group By"
        StrSQL = StrSQL & "                              dbo.Transactions.Transaction_Type,"
        StrSQL = StrSQL & "                              dbo.TransactionTypes.StockEffect,"
        StrSQL = StrSQL & "                              dbo.Transaction_Details.Item_ID,"
        StrSQL = StrSQL & "                              transactions.StoreId"
        StrSQL = StrSQL & "                       Having (dbo.TransactionTypes.StockEffect <> 0)"
        StrSQL = StrSQL & "                       AND (dbo.Transaction_Details.Item_ID = " & mItemId & " )"
        '               --TblItems.ItemID)
        'AND dbo.Transactions.StoreID = 1
        StrSQL = StrSQL & "                   ) T"
        StrSQL = StrSQL & "                   INNER JOIN TblStore AS ts"
        StrSQL = StrSQL & "                        ON  ts.StoreID = T.StoreID"
        StrSQL = StrSQL & "            Group By"
        StrSQL = StrSQL & "                   t.StoreID,"
        StrSQL = StrSQL & "                   ts.StoreName,"
        StrSQL = StrSQL & "                   ts.storenamee"
        StrSQL = StrSQL & "            Having (SUM(t.transqty) >=  " & val(mQuantity) & " )"
        StrSQL = StrSQL & "            ORDER BY  SUM(T.transqty) DEsc "
        Dim rsDummy As New ADODB.Recordset
        Set rsDummy = New ADODB.Recordset
        rsDummy.Open StrSQL, Cn, adOpenStatic, adLockReadOnly
    If rsDummy.EOF Then
        mGrid.TextMatrix(i, mGrid.ColIndex("RequestTypeName")) = "ÿ·» ‘—«¡ œ«Œ·Ï"
        mGrid.TextMatrix(i, mGrid.ColIndex("RequestTypeNo")) = "1"
        mGrid.TextMatrix(i, mGrid.ColIndex("ItemBalance2")) = ""
        
    Else
        mGrid.TextMatrix(i, mGrid.ColIndex("RequestTypeName")) = "ÿ·»  ÕÊÌ· „Œ“‰Ì"
        mGrid.TextMatrix(i, mGrid.ColIndex("RequestTypeNo")) = "2"
        mGrid.TextMatrix(i, mGrid.ColIndex("StoreIDAvi")) = rsDummy!StoreId
      '  mGrid.TextMatrix(i, mGrid.ColIndex("StoreIDAviName")) = rsDummy!StoreName
        mGrid.TextMatrix(i, mGrid.ColIndex("ItemBalance2")) = rsDummy!ItemBalance2
    



    End If
    End If
    
Next
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
        
        
       ElseIf mIndex = 7 Then
        StrRecID = new_id("TblCaptinTrans", "id", "")
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
    ElseIf mIndex = 7 Then
        FiLLRec7

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
                If Voucher_coding(val(dcBranch(mIndex).BoundText), XPDtbTrans.value, 81, 1100, , , , , , , "TblHandWages") = "error" Then
                    If SystemOptions.UserInterface = ArabicInterface Then
                        MsgBox " ·« Ì„ﬂ‰ «÷«›… ”‰œ  ÃœÌœ ·«‰ﬂ  ⁄œÌ  «·Õœ «·–Ì ﬁ„  » ÕœÌœ… „‰ «·”‰œ«   ": Exit Sub
                    Else
                        MsgBox " Cant't Create Expenses Voucher to this Process no You exceed the maximum number ": Exit Sub
                    End If

                Else
         
                    If Voucher_coding(val(dcBranch(mIndex).BoundText), XPDtbTrans.value, 81, 1100, , , , , , , "TblHandWages") = "" Then
                        If SystemOptions.UserInterface = ArabicInterface Then
                            
                            TxtNoteSerial1.locked = False
                            MsgBox " ·«»œ „‰ ﬂ «»… —ﬁ„ «·”‰œ ÌœÊÌ« ﬂ„« Õœœ   ": Exit Sub
                        Else
                            MsgBox "  Enter Voucher No Manually or Define Coding ": Exit Sub
                        End If

                    Else
                        TxtNoteSerial1.Text = Voucher_coding(val(dcBranch(mIndex).BoundText), XPDtbTrans.value, 81, 1100, , , , , , , "TblHandWages")
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
    RsSavRec.Fields("BranchID").value = IIf(dcBranch(mIndex).Text <> "", Trim(dcBranch(mIndex).BoundText), Null)
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
    RsSavRec.Fields("Vat2").value = val(TxtVAt2.Text)
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
    saveGrid s, FG, "Name", "", "MasterID", val(TxtSerial1(mIndex).Text)
    
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





Public Sub FiLLRec7()
   ' On Error GoTo ErrTrap
    
   
    
    
    If TxtModFlg2(mIndex).Text = "N" Then
           ' TxtSerial1(mIndex).Text = CStr(new_id("TblAdditionsAssest", "ID", "", True))

       
        RsSavRec.AddNew
        TxtSerial1(mIndex).Text = new_id("TblCaptinTrans", "id", "")
   '     RsSavRec.Fields("ID").value = val(TxtSerial1(mIndex).Text)
    End If
    RsSavRec("NoteSerial1").value = Trim$(Me.TxtNoteSerial17.Text)
    RsSavRec.Fields("BranchID").value = IIf(dcBranch(mIndex).Text <> "", Trim(dcBranch(mIndex).BoundText), Null)
    RsSavRec.Fields("BankID").value = IIf(DcboBankName.Text <> "", Trim(DcboBankName.BoundText), Null)
    RsSavRec("RecordDate").value = XPDtbTrans7.value
    
       If chkIsVat.value = vbChecked Then
            RsSavRec.Fields("IsVat").value = 1
        Else
            RsSavRec.Fields("IsVat").value = 0
        End If
               
   RsSavRec.Fields("UserID").value = IIf(DCboUserName(mIndex).Text <> "", Trim(DCboUserName(mIndex).BoundText), user_id)
   'RsSavRec("RecType").value = cmbRecType.ListIndex
    'RsSavRec("ContractNo").value = txtContractNo.Text
    'RsSavRec("RecName").value = txtRecName.Text
    'RsSavRec("RecordTime").value = XPDtbTransTime.Value
    

    
   ' RsSavRec("Remarks").value = TxtRemarks.Text
    
    
    '*********************
     
    
    
      
   

    RsSavRec.update
    'cmdDelNote7
    Dim s As String
                
    
        s = " Delete  TblCaptinTrans2 Where MasterID = " & val(TxtSerial1(mIndex).Text)
    
        
        
   
    Cn.Execute s
    
    s = "Select  MasterID,Emp_ID,CompanyName,OperationName,EmpName,typename,Account_Name,DateEntry,Amount"
    s = s & " from TblCaptinTrans2 Where Id = -1"
    'saveGrid s, fg, "Name", "ID", "MasterID", val(TxtSerial1(mIndex).Text)
    saveGrid s, grdExcel, "CompanyName", "", "MasterID", val(TxtSerial1(mIndex).Text)
    
    CmdCreateV7_Click
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
                     xReport.ParameterFields.Item(i).AddCurrentValue "" & val(TxtVAt2)
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
                     xReport.ParameterFields.Item(i).AddCurrentValue "" & val(TxtVAt2)
                     
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
                        xReport.ParameterFields(i).AddCurrentValue GetRegVATNo(val(dcBranch(mIndex).BoundText))
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
      Dim total As String
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
        ElseIf mIndex = 7 Then
            FiLLTXT7
        
            
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

    RsSavRec.Fields("name").value = IIf(TxtName(mIndex).Text <> "", Trim(TxtName(mIndex).Text), Null)
    RsSavRec.Fields("namee").value = IIf(TxtNameE(mIndex).Text <> "", Trim(TxtNameE(mIndex).Text), Null)
    

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
    
If val(DCBoMain(2).BoundText) > val(DCBoMain(5).BoundText) Then
 Dim xx As String
    
    xx = (val(DCBoMain(5).BoundText))
    
    DCBoMain(5).BoundText = DCBoMain(2).BoundText
    DCBoMain(2).BoundText = xx
End If



If val(DCBoMain(3).BoundText) > val(DCBoMain(6).BoundText) Then
 Dim xxx As String
    
    xxx = (val(DCBoMain(6).BoundText))
    
    DCBoMain(6).BoundText = DCBoMain(3).BoundText
    DCBoMain(3).BoundText = xxx
End If

    RsSavRec.Fields("name").value = IIf(TxtName(mIndex).Text <> "", Trim(TxtName(mIndex).Text), Null)
    RsSavRec.Fields("namee").value = IIf(TxtNameE(mIndex).Text <> "", Trim(TxtNameE(mIndex).Text), Null)
    

    RsSavRec.Fields("GroupId").value = IIf(cmbGroupId.Text <> "", Trim(cmbGroupId.BoundText), Null)
    RsSavRec.Fields("UnitID").value = IIf(cmbUnitID.Text <> "", Trim(cmbUnitID.BoundText), Null)
    
    RsSavRec.Fields("FromSPH").value = IIf(DCBoMain(2).Text <> "", Trim(DCBoMain(2).BoundText), Null)
    RsSavRec.Fields("TOSPH").value = IIf(DCBoMain(5).Text <> "", Trim(DCBoMain(5).BoundText), Null)
    RsSavRec.Fields("FROMCYL").value = IIf(DCBoMain(3).Text <> "", Trim(DCBoMain(3).BoundText), Null)
    RsSavRec.Fields("TOCYL").value = IIf(DCBoMain(6).Text <> "", Trim(DCBoMain(6).BoundText), Null)
    RsSavRec.Fields("Price").value = val(TxtPrice)
    
    RsSavRec("MasterType").value = cboMasterType.ListIndex
    Dim mm As Long
    
    For mm = 0 To cmbEyeDet.count - 1
        If mm <> 7 And mm <> 23 And mm <> 22 And mm <> 8 Then
            RsSavRec(GetFieldName(mm)).value = val(Me.cmbEyeDet(mm).BoundText)
        End If
    Next
    
    RsSavRec("Flag").value = cmbFlag(0).ListIndex
   
   'RsSavRec("RecType").value = cmbRecType.ListIndex


    RsSavRec.update
    
    Command1_Click
    MsgBox " „  ⁄„·Ì… «·Õ›Ÿ »‰Ã«Õ", vbInformation + vbOKOnly + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
    
    Dim mmIDD2 As Long
    mmIDD2 = val(TxtSerial1(mIndex).Text)
    FillGridWithData3
    FindRec mmIDD2
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

    With Me.GRID2
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
    TxtName(mIndex).Text = IIf(IsNull(RsSavRec.Fields("name").value), "", RsSavRec.Fields("name").value)
    TxtNameE(mIndex).Text = IIf(IsNull(RsSavRec.Fields("nameE").value), "", RsSavRec.Fields("nameE").value)
    
    LabCurr_Rec(mIndex).Caption = RsSavRec.AbsolutePosition
    LabCount_Rec(mIndex).Caption = RsSavRec.RecordCount

    With GRID2

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




Private Function GetNewCode2(LngParentGroupID As Long, Optional ByVal mTableName As String = "", Optional ByVal mTableGroupName As String = "Groups", Optional ByVal mFieldGroup As String = "GroupID") As String
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
    Dim mm As Double
    mm = val(mTmpGroup2)
    mTmp = val(mm) + 1
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
        IntTemp = val(mId(StrLastGroupCode, Len(StrParentCode)))
        If IntTemp = 0 Then
            IntTemp = val(mId(StrLastGroupCode, Len(StrParentCode)))
        End If
        IntTemp = val(mId(StrLastGroupCode, Len(StrParentCode) - 1))
        StrNewGroupCode = StrLastGroupCode & StrParentCode & IntTemp
    End If

    rs.Close
    Set rs = Nothing
    GetNewCode2 = StrNewGroupCode
    Exit Function
ErrTrap:
End Function







Private Sub cmbEyeDet_Change(Index As Integer)
On Error Resume Next
 If SystemOptions.IsAutoNameItems = False Then Exit Sub
 
If Me.TxtModFlg2(mIndex).Text = "N" Or Me.TxtModFlg2(mIndex).Text = "E" Then
      If Index = 8 Then
    cmbGroupId.BoundText = cmbEyeDet(8).BoundText
      End If

        cmbEyeDet(7).Text = ""
        DoEvents
        
 
        
         If 1 = 1 Then
    If cboMasterType.ListIndex = 1 Then 'frames
    Dim mNameAutoGen As String
    Dim mNameAutoGenEnG As String
    
  mNameAutoGen = cmbEyeDet(0).Text      'brand
'mNameAutoGen = mNameAutoGen & "," & cmbEyeDet(4).Text     'brand Type
'    mNameAutoGen = mNameAutoGen & "," & cmbEyeDet(3).Text     'collection
     
    
             mNameAutoGen = mNameAutoGen & "," & TxtModel
             mNameAutoGen = mNameAutoGen & "," & TxtColorCode
             mNameAutoGen = mNameAutoGen & "," & TxtSize
             
          mNameAutoGenEnG = GetArabicName(val(cmbEyeDet(8).BoundText), 8)  'gategory
          mNameAutoGenEnG = mNameAutoGenEnG & "," & GetArabicName(val(cmbEyeDet(1).BoundText), 1) ''  Type
          mNameAutoGenEnG = mNameAutoGenEnG & "," & GetArabicName(val(cmbEyeDet(2).BoundText), 2) '   '  Design
          mNameAutoGenEnG = mNameAutoGenEnG & "," & GetArabicName(val(cmbEyeDet(5).BoundText), 2) '   '  Shape
          mNameAutoGenEnG = mNameAutoGenEnG & "," & GetArabicName(val(cmbEyeDet(6).BoundText), 6) '  '  Material
          mNameAutoGenEnG = mNameAutoGenEnG & "," & GetArabicName(val(cmbEyeDet(24).BoundText), 24) '  '  Gender
          mNameAutoGenEnG = mNameAutoGenEnG & "," & GetArabicName(val(cmbEyeDet(25).BoundText), 25) '    '  Age
          mNameAutoGenEnG = mNameAutoGenEnG & "," & GetArabicName(val(cmbEyeDet(3).BoundText), 19) '   '  Group/collection
          mNameAutoGenEnG = mNameAutoGenEnG & "," & GetArabicName(val(cmbEyeDet(9).BoundText), 9) '     '  Color
          
     ElseIf cboMasterType.ListIndex = 2 Then 'Optical Lens
            mNameAutoGen = cmbEyeDet(0).Text      'brand
   
             mNameAutoGen = mNameAutoGen & "," & TxtModel
             mNameAutoGen = mNameAutoGen & "," & cmbEyeDet(13).Text 'index
             mNameAutoGen = mNameAutoGen & "," & cmbEyeDet(6).Text 'imaterial
             mNameAutoGen = mNameAutoGen & "," & cmbEyeDet(2).Text 'Design
             mNameAutoGen = mNameAutoGen & "," & cmbEyeDet(1).Text 'Type
             
          mNameAutoGenEnG = mNameAutoGenEnG & "," & GetArabicName(val(cmbEyeDet(0).BoundText), 0) ''  brand      'brand
          mNameAutoGenEnG = mNameAutoGenEnG & "," & TxtModel 'Model
         mNameAutoGenEnG = mNameAutoGenEnG & "," & GetArabicName(val(cmbEyeDet(13).BoundText), 13) ''  index
         mNameAutoGenEnG = mNameAutoGenEnG & "," & GetArabicName(val(cmbEyeDet(6).BoundText), 6) '  '  Material
         mNameAutoGenEnG = mNameAutoGenEnG & "," & GetArabicName(val(cmbEyeDet(2).BoundText), 2) '   '  Design
         mNameAutoGenEnG = mNameAutoGenEnG & "," & GetArabicName(val(cmbEyeDet(1).BoundText), 1) ''  Type
          
           
     ElseIf cboMasterType.ListIndex = 3 Then 'Contact Lens
     
     ElseIf cboMasterType.ListIndex = 4 Then 'Lens Care Product
     
     ElseIf cboMasterType.ListIndex = 5 Then 'Accessories
     
     
     End If
            
            'XPTxtName = mNameAutoGen
            'XPTxtNamee = mNameAutoGenEnG
            
        End If
        
End If
End Sub

Private Function FnGenrateName() As String

On Error Resume Next
 If SystemOptions.IsAutoNameItems = False Then Exit Function
 

     
            cmbEyeDet(8).BoundText = cmbGroupId.BoundText
    

        cmbEyeDet(7).Text = ""
        DoEvents
        
 
        
         If 1 = 1 Then
    If cboMasterType.ListIndex = 1 Then 'frames
    Dim mNameAutoGen As String
    Dim mNameAutoGenEnG As String
    
  mNameAutoGen = cmbEyeDet(0).Text      'brand
'mNameAutoGen = mNameAutoGen & "," & cmbEyeDet(4).Text     'brand Type
'    mNameAutoGen = mNameAutoGen & "," & cmbEyeDet(3).Text     'collection
     
    
             mNameAutoGen = mNameAutoGen & "," & TxtModel
             mNameAutoGen = mNameAutoGen & "," & TxtColorCode
             mNameAutoGen = mNameAutoGen & "," & TxtSize
             
          mNameAutoGenEnG = GetArabicName(val(cmbEyeDet(8).BoundText), 8)  'gategory
          mNameAutoGenEnG = mNameAutoGenEnG & "," & GetArabicName(val(cmbEyeDet(1).BoundText), 1) ''  Type
          mNameAutoGenEnG = mNameAutoGenEnG & "," & GetArabicName(val(cmbEyeDet(2).BoundText), 2) '   '  Design
          mNameAutoGenEnG = mNameAutoGenEnG & "," & GetArabicName(val(cmbEyeDet(5).BoundText), 2) '   '  Shape
          mNameAutoGenEnG = mNameAutoGenEnG & "," & GetArabicName(val(cmbEyeDet(6).BoundText), 6) '  '  Material
          mNameAutoGenEnG = mNameAutoGenEnG & "," & GetArabicName(val(cmbEyeDet(24).BoundText), 24) '  '  Gender
          mNameAutoGenEnG = mNameAutoGenEnG & "," & GetArabicName(val(cmbEyeDet(25).BoundText), 25) '    '  Age
          mNameAutoGenEnG = mNameAutoGenEnG & "," & GetArabicName(val(cmbEyeDet(3).BoundText), 19) '   '  Group/collection
          mNameAutoGenEnG = mNameAutoGenEnG & "," & GetArabicName(val(cmbEyeDet(9).BoundText), 9) '     '  Color
          
     ElseIf cboMasterType.ListIndex = 2 Then 'Optical Lens
            mNameAutoGen = cmbEyeDet(0).Text      'brand
   
             mNameAutoGen = mNameAutoGen & "," & TxtModel
             mNameAutoGen = mNameAutoGen & "," & cmbEyeDet(13).Text 'index
             mNameAutoGen = mNameAutoGen & "," & cmbEyeDet(6).Text 'imaterial
             mNameAutoGen = mNameAutoGen & "," & cmbEyeDet(2).Text 'Design
             mNameAutoGen = mNameAutoGen & "," & cmbEyeDet(1).Text 'Type
             
          mNameAutoGenEnG = mNameAutoGenEnG & "," & GetArabicName(val(cmbEyeDet(0).BoundText), 0) ''  brand      'brand
          mNameAutoGenEnG = mNameAutoGenEnG & "," & TxtModel 'Model
         mNameAutoGenEnG = mNameAutoGenEnG & "," & GetArabicName(val(cmbEyeDet(13).BoundText), 13) ''  index
         mNameAutoGenEnG = mNameAutoGenEnG & "," & GetArabicName(val(cmbEyeDet(6).BoundText), 6) '  '  Material
         mNameAutoGenEnG = mNameAutoGenEnG & "," & GetArabicName(val(cmbEyeDet(2).BoundText), 2) '   '  Design
         mNameAutoGenEnG = mNameAutoGenEnG & "," & GetArabicName(val(cmbEyeDet(1).BoundText), 1) ''  Type
          
           
     ElseIf cboMasterType.ListIndex = 3 Then 'Contact Lens
     
     ElseIf cboMasterType.ListIndex = 4 Then 'Lens Care Product
     
     ElseIf cboMasterType.ListIndex = 5 Then 'Accessories
     
     
     End If
            
            'XPTxtName = mNameAutoGen
            'XPTxtNamee = mNameAutoGenEnG
            
        End If
        

FnGenrateName = mNameAutoGenEnG & "," & (cmbEyeDet(22).Text) & "," & cmbEyeDet(23).Text
End Function

Function GetArabicName(ID As Integer, mm As Integer) As String
            
            
          Dim i As Integer
Dim BrnchIDes As String
BrnchIDes = "-1"
Dim rs2 As ADODB.Recordset
Set rs2 = New ADODB.Recordset
Dim sql As String
  sql = " SELECT     Name"
 sql = sql & "                   From " & GetTableName(mm)
 sql = sql & " where id=" & ID
If mm = 8 Then
sql = "select groupname as Name from groups where groupid=" & ID
End If

If mm = 9 Then
sql = "select colorname as Name from TblItemsColors where colorid=" & ID
End If
 
rs2.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
If rs2.RecordCount > 0 Then
 
GetArabicName = IIf(IsNull(rs2("Name").value), -1, rs2("Name").value)
 
 Else
 GetArabicName = ""
End If
 

End Function
Private Sub cmbEyeDet_Click(Index As Integer, Area As Integer)
 On Error Resume Next
    Dim OverHead As Double
    OverHead = 0
End Sub

Private Sub cmbSex_Change()
'If Me.TxtModFlg.Text = "N" Or Me.TxtModFlg.Text = "E" Then
'    If SystemOptions.IsAutoNameItems Then
'           Dim i As Long
'               mNameAutoGen = ""
'
'           For i = 0 To cmbEyeDet.count - 1
'                 If i <> 8 Then
'                   If i = 1 Then
'                     '  If cmbEyeDet(i).Text <> "" And val(cmbEyeDet(i).BoundText) <> 0 Then
'
'                       If cmbEyeDet(i).Text <> "" And val(cmbEyeDet(i).BoundText) <> 0 Then
'                           mNameAutoGen = mNameAutoGen & IIf(mNameAutoGen <> "", "-", "") & cmbEyeDet(i).Text & IIf(cmbEyeDet(i).Text = "" And val(cmbEyeDet(i).BoundText) <> 0, "", "-") & cmbSex.Text & IIf(cmbSex.Text = "", "", "-") & cmbAge.Text
'                       Else
'
'                           mNameAutoGen = mNameAutoGen & IIf(cmbSex.Text = "", "", "-") & cmbSex.Text & IIf(cmbSex.Text = "", "", "-") & cmbAge.Text
'                       End If
'                   Else
'                       If cmbEyeDet(i).Text <> "" And val(cmbEyeDet(i).BoundText) <> 0 Then
'                           mNameAutoGen = mNameAutoGen & IIf(mNameAutoGen <> "", "-", "") & cmbEyeDet(i).Text
'                       End If
'                   End If
'               End If
'           Next
'       End If
'        If mNameAutoGen <> "" Then
'           mNameAutoGen = cmbEyeDet(8).Text & "-" & mNameAutoGen
'           XPTxtName = mNameAutoGen
'           XPTxtNamee = mNameAutoGen
'       End If
'    End If
End Sub

Private Sub cmbSex_Click()
If Me.TxtModFlg2(mIndex).Text = "N" Or Me.TxtModFlg2(mIndex).Text = "E" Then
    cmbEyeDet_Change 0
End If
End Sub




Private Function GetTableName(ByVal mIndex As String) As String

Select Case mIndex
Case 0
    GetTableName = "tblBrands"
Case 1
    GetTableName = "tblTypeItems"
Case 2
    GetTableName = "tblDesign"
Case 3
    GetTableName = "tblCollections"
Case 4
    GetTableName = "tblShapes"
Case 5
    GetTableName = "tblShapesNew"
Case 6
    GetTableName = "tblMaterial"
Case 10
    GetTableName = "tblOrigin"
Case 11
    GetTableName = "tblDivision"
Case 12
    GetTableName = "tblCoating"
Case 13
    GetTableName = "tblIndexs"
Case 14
    GetTableName = "tblDIAM"
Case 15
    GetTableName = "tblLightAdaptation"
Case 16
    GetTableName = "tblBreaking"
Case 17
    GetTableName = "tblService"
Case 18
    GetTableName = "tblBaseCurve"
    
Case 19
    GetTableName = "tblGroupEye"
Case 20
    GetTableName = "tblUsage"
Case 21
    GetTableName = "tblPacking"
Case 22
    GetTableName = "SPHTable"
Case 23
    GetTableName = "CLYTable"
    

    Case 24
    GetTableName = "TblSex"
    Case 25
    GetTableName = "TblAge"
    
    
    
End Select





End Function


Private Function GetFieldName(ByVal mIndex As String) As String
Select Case mIndex
Case 0
    GetFieldName = "BrandsID"
Case 1
    GetFieldName = "TypeItemsID"
Case 2
    GetFieldName = "DesignID"
Case 3
    GetFieldName = "CollectionsID"
Case 4
    GetFieldName = "ShapesID"
Case 5
    GetFieldName = "ShapesNewID"
Case 6
    GetFieldName = "MaterialID"
Case 7
    GetFieldName = "GroupID"
Case 8
    GetFieldName = "NationalityID"
Case 9

    GetFieldName = "ColorID11"
    
Case 10
    GetFieldName = "OriginID"
Case 11
    GetFieldName = "DivisionID"
Case 12
    GetFieldName = "CoatingID"
Case 13
    GetFieldName = "IndexsID"
Case 14
    GetFieldName = "DIAMID"
Case 15
    GetFieldName = "LightAdaptationID"
Case 16
    GetFieldName = "BreakingID"
Case 17
    GetFieldName = "ServiceID"
Case 18
    GetFieldName = "BaseCurveID"
    
Case 19
    GetFieldName = "GroupEyeID"
Case 20
    GetFieldName = "UsageID"
Case 21
    GetFieldName = "PackingID"
Case 22
    GetFieldName = "SphereID"
        
Case 23
    GetFieldName = "CylinderID"
            
Case 24
    GetFieldName = "SexID"
                     
Case 25
    GetFieldName = "AGEID"
                              
  
End Select

End Function



Public Function GetGridFileName(ByVal G As Object, Optional MainFormName As String = "") As String
    Dim GlobalGridName As String
    Dim IndexS As String
    Dim MainContainerName As String

    On Error Resume Next
    IndexS = G.Index

    MainContainerName = GetMainForm(G.Container)
    GlobalGridName = MainContainerName & "\" & G.Name & IndexS & MainFormName
    GlobalGridName = "Import"
    GetGridFileName = App.path & GlobalGridName & ".xls"

End Function
Public Function GetMainForm(ByVal Obj) As String
    Dim n As String
    On Error Resume Next
    n = Obj.Container.Name

    If n = "" Then
        GetMainForm = Obj.Name
    Else
        GetMainForm = GetMainForm(Obj.Container)
    End If
End Function




Public Sub FromExcel(ByRef mGrid As Object, _
                     ByRef mtmpGrd As Object, _
                     Frm As Form, _
                     Optional MainFormName As String = "", _
                     Optional ProgressBar As Object = Nothing, Optional ByVal XlsFileName As String = "", Optional ByVal MainTableName As String = "")


    ' If Not i Then Exit Sub
       Dim cProgress As ClsProgress
       Dim i As Long, jj As Long, j As Long, H As Long
    '    Dim mtmpGrd As VSFlexGrid
    If XlsFileName = "" Then
        XlsFileName = GetGridFileName(mGrid, MainFormName)
    End If
    If FileExists(XlsFileName) Then

        mtmpGrd.FixedCols = 0
        mtmpGrd.FixedRows = 0

        mtmpGrd.loadgrid XlsFileName, flexFileExcel

        mtmpGrd.backcolor = &HFFFFFF
        mtmpGrd.BackColorAlternate = &HE9E9E9
        mtmpGrd.BackColorBkg = &H8000000C
        mtmpGrd.BackColorFixed = &H8000000F
        mtmpGrd.BackColorFrozen = &HC0FFFF
        mtmpGrd.BackColorSel = &H8000000D
        mtmpGrd.ForeColor = &H80000008
        mtmpGrd.ForeColorFixed = &HFF0000
        mtmpGrd.ForeColorSel = &H8000000E
        mtmpGrd.GridColor = &H8000000F
        mtmpGrd.GridColorFixed = &H80000010
        mtmpGrd.FixedCols = 1
        mtmpGrd.FixedRows = 1
        '·«‰ Loaded ÌŒ ›Ì
        mtmpGrd.Cols = mGrid.Cols + 1
        mtmpGrd.ColKey(mtmpGrd.Cols - 1) = "Loaded"
        mtmpGrd.ColHidden(mtmpGrd.Cols - 1) = True
        mtmpGrd.AutoSize 0, mtmpGrd.Cols - 1
    End If
    mGrid.Rows = 1
    
    For i = 1 To mtmpGrd.Rows - 1
        If i <= mtmpGrd.Rows - 1 Then
            If chkIsDiscountOnly.value = vbUnchecked Then
                If mtmpGrd.TextMatrix(i, 4) = "Œ’„" Then
                    mtmpGrd.RemoveItem i
                    i = i - 1
    
                End If
            End If
            If chkIsAddOnly.value = vbUnchecked Then
                If mtmpGrd.TextMatrix(i, 4) = "√÷«›…" Or mtmpGrd.TextMatrix(i, 4) = "«÷«›…" Or mtmpGrd.TextMatrix(i, 4) = "«÷«›Â" Then
                    mtmpGrd.RemoveItem i
                    i = i - 1
                End If
            End If
        End If
    Next
    
    mGrid.Rows = mtmpGrd.Rows

    '********************************
    If Not ProgressBar Is Nothing Then
        ProgressBar.Min = 1
        ProgressBar.Max = IIf(mGrid.Rows > 2, mGrid.Rows - 1, 2)    ' mGrid.Rows - 1
        ProgressBar.Visible = True
        '********************************
    End If
        Set cProgress = New ClsProgress
       cProgress.ProgressType = Waiting
    

    



    
    Dim Hide As Integer
    For i = 1 To mtmpGrd.Rows - 1
        '********************************
        If Not ProgressBar Is Nothing Then
            ProgressBar.value = i
            DoEvents
            ProgressBar.Refresh
        End If
        cProgress.StartProgress
       DoEvents
        '********************************
        jj = 0
        For j = 1 To mGrid.Cols - 1
            If j = 18 Then
                j = 18
            End If
            If Not mGrid.ColHidden(j) Then
                jj = jj + 1
                       If mGrid.ColKey(j) = "Account_Code" Then
                            GoTo NextCol
                            
                     
                    j = j
                End If
                Debug.Print i & " " & mGrid.TextMatrix(i, j)
                If InStr(1, mGrid.ColComboList(j), "#") Then
                    Hide = 0
                    For H = j - 1 To 1 Step -1
                        Hide = Hide + IIf(mGrid.ColHidden(H), 1, 0)
                    Next
                    mGrid.TextMatrix(i, j) = mtmpGrd.TextMatrix(i, j - Hide)
                    'Replace(Trim(mtmpGrd.TextMatrix(i, jj)), "'", "")
                Else
                    mGrid.TextMatrix(i, j) = Replace(Trim(mtmpGrd.TextMatrix(i, jj)), "'", "")
                End If
                If Trim(mGrid.ColEditMask(j)) = "Date" Then
                    GetFieldID mGrid.ColEditMask(j), i, j, mGrid
                End If
                'pValue = Split(G.ColComboList(j), ";")
            Else
                j = j
                If j = 34 Then
                j = j
                End If
                If Trim(mGrid.ColEditMask(j)) <> "" Then
                    GetFieldID mGrid.ColEditMask(j), i, j, mGrid, MainTableName
                End If
                If Trim(mGrid.ColComboList(j)) <> "" Then
                    GetIDCombo Trim(mGrid.ColComboList(j)), i, j, mGrid
                End If
            End If
            If Trim(Replace(Trim(mtmpGrd.TextMatrix(i, 1)), "'", "")) = "" Then
                mGrid.Rows = i + 1:  Exit Sub
            End If
NextCol:
        Next
        ' DisplayOrderTotals
NextRow:
    Next
    '********************************
    If Not ProgressBar Is Nothing Then
        ProgressBar.Visible = False
    End If
           DoEvents
    cProgress.FinishProgress
    cProgress.StopProgess
    Set cProgress = Nothing
    MsgBox " „ «·«œ—«Ã"
    '********************************
    
End Sub

Private Sub GetIDCombo(ByVal mTableColID As String, ByVal mRow As Long, ByVal mCol As Long, ByVal mGrid As Object)
Dim mTxt As String
mTxt = Trim(mGrid.TextMatrix(mRow, mCol - 1))
Select Case mTableColID
Case "sexID"
    If mTxt = "Male" Or mTxt = "–ﬂ—" Then
        mTxt = 1
    Else
        mTxt = 2
    End If
Case "MaritalStatusID"
'    DcbMatrial.AddItem "√⁄“»"
'      DcbMatrial.AddItem "„ “ÊÃ"
    If mTxt = "√⁄“»" Or mTxt = "Single" Then
        mTxt = 0
    ElseIf mTxt = "„ “ÊÃ" Or UCase(mTxt) = "MARRIED" Then
        mTxt = 1
    ElseIf mTxt = "„ÿ·ﬁ/„ÿ·›…" Or UCase(mTxt) = "DIVORCED" Then
        mTxt = 2
    ElseIf mTxt = "«—„·/√—„·…" Or UCase(mTxt) = "WIDOWED" Then
        mTxt = 3
        
    End If
    
Case "Status_id"
'    DcbMatrial.AddItem "√⁄“»"
'      DcbMatrial.AddItem "„ “ÊÃ"
    If mTxt = "Ã«—Ì «·«Â·«ﬂ" Or mTxt = "Ã«—Ï «·«Â·«ﬂ" Then
        mTxt = 0
    ElseIf mTxt = "„ Êﬁ›" Or UCase(mTxt) = "Stoped" Then
        mTxt = 1
    ElseIf mTxt = " „ «· Œ·’ »«·»Ì⁄" Or UCase(mTxt) = " „ «· Œ·’ »«·»Ì⁄" Then
        mTxt = 2
    ElseIf mTxt = " „ «·«Â·«ﬂ »«· Œ—Ìœ" Or UCase(mTxt) = " „ «·«Â·«ﬂ »«· Œ—Ìœ" Then
        mTxt = 3
        
    End If
    
 Case "Depreciation_Type_id"
'    DcbMatrial.AddItem "√⁄“»"
'      DcbMatrial.AddItem "„ “ÊÃ"
    If mTxt = "«·ﬁ”ÿ «·À«» " Or mTxt = "«·ﬁ”ÿ «·À«» " Then
        mTxt = 0
    ElseIf mTxt = "«·ﬁ”ÿ  «·„ ‰«ﬁ’" Or UCase(mTxt) = "«·ﬁ”ÿ  «·„ ‰«ﬁ’" Then
        mTxt = 1

    End If
       
Case "Emp_Name1.Emp_Name2.Emp_Name3.Emp_Name4"
    mTxt = mGrid.TextMatrix(mRow, mCol - 4) + " " + mGrid.TextMatrix(mRow, mCol - 3) + " " + mGrid.TextMatrix(mRow, mCol - 2) + " " + mGrid.TextMatrix(mRow, mCol - 1)
Case ""
End Select
mGrid.TextMatrix(mRow, mCol) = mTxt
End Sub

Public Function ToHijriDate(ByVal GregorianDate As String) As String
    Dim HijriDate As String, DateFormat As String
    ' DateFormat = "long date"
    
    DateFormat = "dd-mm-yyyy"
    HijriDate = ConvertDate(GregorianDate, vbCalGreg, vbCalHijri, DateFormat)
    ToHijriDate = HijriDate
    
End Function
Private Function ConvertDate(ByRef StringIn As String, _
                             ByRef OldCalender As Integer, _
                             ByVal NewCalender As Integer, _
                             ByRef NewFormat As String) As String
                             If StringIn = "" Then Exit Function
On Error Resume Next
    Dim SavedCal As Integer
    Dim d As Date, s As String
    SavedCal = Calendar
    Calendar = OldCalender
    d = CDate(StringIn)
    Calendar = NewCalender
    s = CStr(d)
    ConvertDate = Format(s, NewFormat)
    Calendar = SavedCal
End Function

Public Function ToGregorianDate(ByVal HijriDate As String) As Date
    Dim GregorianDate As String, DateFormat As String
  If HijriDate = "" Then Exit Function
    DateFormat = "dd/mm/yyyy"
    
    GregorianDate = ConvertDate(HijriDate, vbCalHijri, vbCalGreg, DateFormat)
    If DateDiff("D", "01/01/1900", GregorianDate) < 0 Then
    GregorianDate = Date
    End If
    ToGregorianDate = GregorianDate
End Function

Public Function CheckDateIsHij(ByVal mDate As String) As Integer
    If Not IsDate(mDate) Then CheckDateIsHij = 3: Exit Function
    
    If Trim(mDate) = "" Then CheckDateIsHij = 3: Exit Function
    
    If year(mDate) < 1800 Then
        CheckDateIsHij = 1
    Else
        CheckDateIsHij = 2
    End If
End Function


Private Sub GetFieldID(ByVal mTableColName As String, ByVal mRow As Long, ByVal mCol As Long, ByVal mGrid As Object, Optional ByVal MainTableName As String = "")
    Dim mTableName As String
    Dim mFieldIDName As String
    Dim mFieldName As String
    Dim xx As Variant
    Dim mValue As String
    Dim rsDummy As New ADODB.Recordset
    Dim rsDummy2 As New ADODB.Recordset
    If mCol = 67 Then
        mCol = 67
    End If
    If mGrid.ColKey(mCol) = "NationlID" Then
        mCol = mCol
    End If
    Dim mValue2 As String
    If mGrid.ColKey(mCol) = "DeanID" Then
        mCol = mCol
    End If
    If mGrid.ColKey(mCol) = "DOBH" Then
        mCol = mCol
    End If
    If mTableColName = "Date" Then
        If CheckDateIsHij(Trim(mGrid.TextMatrix(mRow, mCol - 1))) = 1 Then
            'If Trim(mGrid.TextMatrix(mRow, mCol - 1)) <> "" Then
                mGrid.TextMatrix(mRow, mCol) = Trim(mGrid.TextMatrix(mRow, mCol - 1))
                mGrid.TextMatrix(mRow, mCol - 1) = ToGregorianDate(mGrid.TextMatrix(mRow, mCol))
            'Else
            'End If
        ElseIf CheckDateIsHij(Trim(mGrid.TextMatrix(mRow, mCol - 1))) = 2 Then
            If Trim(mGrid.TextMatrix(mRow, mCol - 1)) = "" Then
                mGrid.TextMatrix(mRow, mCol - 1) = ToGregorianDate(Trim(mGrid.TextMatrix(mRow, mCol)))
            Else
                mGrid.TextMatrix(mRow, mCol) = ToHijriDate(Trim(mGrid.TextMatrix(mRow, mCol - 1)))
            End If
        ElseIf CheckDateIsHij(Trim(mGrid.TextMatrix(mRow, mCol - 1))) = 3 Then
            If mGrid.TextMatrix(mRow, mCol) <> "" Then
                mGrid.TextMatrix(mRow, mCol - 1) = ToGregorianDate(Trim(mGrid.TextMatrix(mRow, mCol)))
            End If
            'mGrid.TextMatrix(mRow, mCol - 1) = ToGregorianDate(mGrid.TextMatrix(mRow, mCol))
        Else
        
        End If
        Exit Sub
    End If
    xx = Split(mTableColName, ",")
    mTableName = xx(0)
    mFieldIDName = xx(1)
    mFieldName = xx(2)
    
 If mRow = 50 Then
 mRow = mRow
 End If
    mValue = Trim(mGrid.TextMatrix(mRow, mCol - 1))
Dim strValue As String
strValue = ""
Dim mValue3 As String

mValue3 = mValue
If (Right(mValue, 1)) = "Â" Then
    strValue = "…"
ElseIf (Right(mValue, 1)) = "…" Then
    strValue = "Â"
    
End If
If strValue <> "" Then
    mValue3 = Replace(mValue3, Right(mValue3, 1), strValue)
End If
Dim mEngLett As String
mEngLett = "e"
    Dim s As String
    mValue2 = mValue
    Select Case mTableName
    Case "jopstatus"
        If UCase(mValue) = "ACTIVE" Then
            mValue2 = "⁄·Ï ﬁÊ… «·⁄„·"
            
        End If
    Case "dean"
      If UCase(mValue) = "ISLAM" Then
            mValue2 = "„”·„"
       ElseIf UCase(mValue) = "CHRISTIAN" Then
            mValue2 = "„”ÌÕÏ"
        End If
    Case "Nationality"
        If UCase(mValue) = "JORDAN" Then
            mValue2 = "«—œ‰"
        ElseIf UCase(mValue) = "INDIA" Then
            mValue2 = "Â‰œ"
        ElseIf Trim(UCase(mValue)) = "" Then
            mValue2 = "”⁄ÊœÌ"
        ElseIf UCase(mValue) = "EGYPT" Then
            mValue2 = "„’—"
        ElseIf UCase(mValue) = "PAKISTAN" Then
            mValue2 = "»«ﬂ” «‰"
        ElseIf UCase(mValue) = "BANGLADESH" Then
            mValue2 = "»‰Ã·«œÌ‘"
        ElseIf UCase(mValue) = "SUDAN" Then
            mValue2 = "”Êœ«‰"
        ElseIf UCase(mValue) = "ETHIOPIA" Then
            mValue2 = "«ÀÌÊ»Ì«"
            
        ElseIf UCase(mValue) = "CAMEROON" Then
            mValue2 = "ﬂ«„Ì—Ê‰"
        ElseIf UCase(mValue) = "PALESTINE" Then
            mValue2 = "›·”ÿÌ‰"
        ElseIf UCase(mValue) = "SYRIA" Then
            mValue2 = "”Ê—Ì«"
        ElseIf UCase(mValue) = "JORDANIAN" Then
            mValue2 = "«—œ‰"
        ElseIf UCase(mValue) = "AMERICA" Then
            mValue2 = "«„—Ìﬂ«"
        ElseIf UCase(mValue) = "EGYPTIAN" Then
            mValue2 = "„’—"
        ElseIf UCase(mValue) = "KENYA" Then
            mValue2 = "ﬂÌ‰Ì«"
        ElseIf UCase(mValue) = "LEBANON" Then
            mValue2 = "·»‰«‰"
        ElseIf UCase(mValue) = "SIRLANKIAN" Then
            mValue2 = "”Ì—·«‰ﬂ"
        ElseIf UCase(mValue) = "YEMEN" Then
            mValue2 = "Ì„‰"
        ElseIf UCase(mValue) = "TUNIS" Then
            mValue2 = " Ê‰”"
        ElseIf UCase(mValue) = "MALAYSIA" Then
            mValue2 = "„«·Ì“Ì«"
         Else
            mValue2 = mValue
         
            
        End If
        If mValue = "" Then mValue2 = "”⁄ÊœÌ"
    Case Else
    End Select
    If mValue = "" Then
        Exit Sub
    End If
    mEngLett = "e"
    If UCase(mTableName) = "ACCOUNTS" Then
         mEngLett = "Eng"
    End If
    If UCase(mTableName) = "TBLCOUNTRIESGOVERNMENTS" Then
         mEngLett = ""
    End If

    
    s = "Select " & mFieldName & " ," & mFieldIDName & " ," & Trim(mFieldName) & mEngLett & "   "
    If UCase(mTableName) = "GROUPS" Or UCase(mTableName) = "GROUPSCUSTOMERS" Or UCase(mTableName) = "FIXEDASSETSGROUP" Then
        s = s & " ,ParentID,FullCode,GroupCode,Code,LastGroup "
    End If
    
    s = s & " from  " & mTableName
    s = s & " Where (" & mFieldName & " = '" & Trim(mValue2) & "' Or " & Trim(mFieldName) & mEngLett & "    = '" & Trim(mValue) & "')"
    s = s & " or (" & mFieldName & " = '" & Trim(mValue3) & "' Or " & Trim(mFieldName) & mEngLett & "   = '" & Trim(mValue3) & "')"
    If UCase(mTableName) = "GROUPS" Or UCase(mTableName) = "GROUPSCUSTOMERS" Or UCase(mTableName) = "FIXEDASSETSGROUP" Then
        s = s & " Or FullCode = '" & Trim(mValue3) & "' "
    End If
    Set rsDummy = New ADODB.Recordset
    rsDummy.Open s, Cn, adOpenKeyset, adLockOptimistic
    
    If rsDummy.EOF Then
        s = s & " Or ( " & mFieldName & " Like '%" & Trim(mValue2) & "%' Or " & Trim(mFieldName) & mEngLett & "    Like '%" & Trim(mValue) & "%')"
    
    End If
    If rsDummy.EOF And UCase(mTableName) = "ACCOUNTS" Then
        MsgBox "Â–« «·Õ”«» €Ì— „ÊÃÊœ ›Ï «·œ·Ì· " & mValue
        Exit Sub
    End If
    rsDummy.Close
    rsDummy.Open s, Cn, adOpenKeyset, adLockOptimistic
    If UCase(mTableName) = "GROUPS" And rsDummy.EOF Then
        rsDummy.Close
             s = "Select " & mFieldName & " ," & mFieldIDName & " ," & Trim(mFieldName) & "e   "
        If UCase(mTableName) = "GROUPS" Or UCase(mTableName) = "GROUPSCUSTOMERS" Or UCase(mTableName) = "FIXEDASSETSGROUP" Then
            s = s & " ,ParentID,FullCode,GroupCode,Code,LastGroup "
        End If
        Dim mValue4  As String
        mValue4 = Trim(mGrid.TextMatrix(mRow, mCol - 2))
        
        s = s & " from  " & mTableName
        s = s & " Where " & mFieldName & " Like '%" & Trim(mValue2) & "%' Or " & Trim(mFieldName) & "e Like '%" & Trim(mValue) & "%'"
        s = s & " Or Fullcode   Like '%" & Trim(mValue4) & "%' Or Code Like '%" & Trim(mValue4) & "%'"
        rsDummy.Open s, Cn, adOpenKeyset, adLockOptimistic
        If rsDummy.EOF Then
            mValue4 = mValue4
        End If
    End If
    
    If Not rsDummy.EOF Then
        If UCase(mTableName) = "ACCOUNTS" Then
            mGrid.TextMatrix(mRow, mCol) = Trim(rsDummy.Fields.Item(Trim(mFieldIDName)) & "")
        Else
            mGrid.TextMatrix(mRow, mCol) = val(rsDummy.Fields.Item(Trim(mFieldIDName)) & "")
            If UCase(mTableName) = "TBLEMPLOYEE" Then
                s = "Select Account_Code from TblBoxesData Where empid = " & val(rsDummy.Fields.Item(Trim(mFieldIDName)) & "")
                Set rsDummy = New ADODB.Recordset
                rsDummy.Open s, Cn, adOpenKeyset, adLockOptimistic
                If Not rsDummy.EOF Then
                    mGrid.TextMatrix(mRow, mGrid.ColIndex("Account_Code")) = rsDummy!Account_code & ""
                End If
            End If
        End If
        If mGrid.ColKey(mCol) = "ParentID" Then
            mGrid.TextMatrix(mRow, mGrid.ColIndex("Code")) = Trim(mGrid.TextMatrix(mRow, mGrid.ColIndex("FullCode")))
            Dim mmm As String
            mmm = SearchInGrid(mGrid, mValue, "GroupName")
            If mmm <> "" Then
                'mGrid.TextMatrix(mRow, mGrid.ColIndex("GroupCode")) = GetNewGroupCode(Val(mGrid.TextMatrix(CLng(mmm), mGrid.ColIndex("NewId"))))
            End If
            mGrid.TextMatrix(mRow, mGrid.ColIndex("LastGroup")) = 0
        End If

    Else
       
        rsDummy.AddNew
        rsDummy(Trim(mFieldName)) = mValue
        rsDummy(Trim(mFieldName) & mEngLett) = mValue
        If mGrid.ColKey(mCol) = "ParentID" Then
            'rsDummy("ParentID") = mValue
            Dim mm As String
            mm = SearchInGrid(mGrid, mValue, "GroupName")
            If mm <> "" Then
                rsDummy("ParentID") = val(mGrid.TextMatrix(CLng(mm), mCol))
                rsDummy("FullCode") = Trim(mGrid.TextMatrix(CLng(mm), mGrid.ColIndex("FullCode")))
                rsDummy("Code") = Trim(mGrid.TextMatrix(CLng(mm), mGrid.ColIndex("FullCode")))
            Else
                xx = Split(Trim(mGrid.TextMatrix(mRow, mGrid.ColIndex("FullCode"))), "-")
                rsDummy("ParentID") = 1
                rsDummy("FullCode") = xx(0)
                rsDummy("Code") = xx(0)
            End If
            rsDummy("GroupCode") = GetNewGroupCode(val(rsDummy("ParentID") & ""), mTableName)
            
            rsDummy("LastGroup") = 0
            If mm <> "" Then
                mGrid.TextMatrix(CLng(mm), mGrid.ColIndex("Code")) = Trim(mGrid.TextMatrix(CLng(mm), mGrid.ColIndex("FullCode")))
                mGrid.TextMatrix(CLng(mm), mGrid.ColIndex("GroupCode3")) = rsDummy("GroupCode") & ""
                mGrid.TextMatrix(CLng(mm), mGrid.ColIndex("LastGroup")) = 0
            End If
        End If
        s = "Select Max(" & mFieldIDName & ")  as MaxID  from  " & mTableName
        
        rsDummy2.Open s, Cn, adOpenKeyset, adLockOptimistic
        Dim mMaxId As Long
        If Not rsDummy2.EOF Then
            mMaxId = val(rsDummy2!MaxID & "") + 1
        Else
            mMaxId = 1
        End If
        If UCase(mTableName) <> "GROUPSCUSTOMERS" Then
            rsDummy(Trim(mFieldIDName)) = mMaxId
        End If
        rsDummy(Trim(mFieldName)) = mValue
        rsDummy.update
       ' mGrid.TextMatrix(mRow, mGrid.ColIndex("NewId")) = mMaxId
        mGrid.TextMatrix(mRow, mCol) = rsDummy(Trim(mFieldIDName) & "")
        CreateExpensType mMaxId, mValue, mRow
    End If

End Sub

Private Function SearchInGrid(ByVal mGrd As Object, ByVal mTxt As String, ByVal mFldName As String) As String
Dim i As Long
For i = 1 To mGrd.Rows - 1
    If Trim(mGrd.TextMatrix(i, mGrd.ColIndex(mFldName))) = mTxt Then
        SearchInGrid = i
        Exit Function
    End If
Next
SearchInGrid = ""
End Function
Function FileExists(filename) As Boolean
    On Error GoTo CheckError        ' Turn on error trapping so error handler                            ' responds if any error is detected.
    FileExists = (Dir(filename) <> "")
    Exit Function            ' Avoid executing error handler                             ' if no error occurs.

CheckError:        ' Branch here if error occurs.    ' Define constants to represent Visual Basic error code.
    FileExists = False
    Resume Next
End Function






Private Function GetNewGroupCode(LngParentGroupID As Long, Optional ByVal mTableName As String = "") As String
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
    StrSQL = "Select GroupCode From " & mTableName & "  Where GroupID=" & LngParentGroupID & ""
    Set rs = New ADODB.Recordset
    rs.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

    If Not (rs.EOF Or rs.BOF) Then
        StrParentCode = IIf(IsNull(rs("GroupCode").value), "", rs("GroupCode").value)
    End If

    rs.Close
    Set rs = New ADODB.Recordset
    StrSQL = "Select * From " & mTableName & "  Where ParentID=" & LngParentGroupID & " Order By GroupID"
    rs.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

    If rs.BOF Or rs.EOF Then
        StrNewGroupCode = StrParentCode & "1"
    Else
        rs.MoveLast
        StrLastGroupCode = IIf(IsNull(rs("GroupCode").value), "", rs("GroupCode").value)
        IntTemp = val(mId(StrLastGroupCode, Len(StrParentCode) + 1))
        StrNewGroupCode = StrParentCode & CStr(IntTemp + 1)
    End If

    rs.Close
    Set rs = Nothing
    GetNewGroupCode = StrNewGroupCode
    Exit Function
ErrTrap:
End Function





Private Sub CreateExpensType(ByVal mEmpId As Long, ByVal mName As String, ByVal mRow As Long)

    Dim rs As ADODB.Recordset
    Dim StrSQL As String
    Dim StrParentCode  As String
    Dim StrNewGroupCode As String
    Dim StrLastGroupCode As String
    Dim IntTemp As String
    Dim mTable  As String
    Dim sql As String
    Dim s As String
    Dim StrNewAccountCode As String
    Dim RsData As New ADODB.Recordset
    mTable = "TblBoxesData"
    Dim mMaxId As Long
    Dim mSer As Long
    Dim mParent_account  As String
    s = "SELECT parent_account FROM " & mTable & " AS te Where IsNull(parent_account,'') <> '' and IsNull(parent_account,'') In (Select Account_Code from Accounts)"
    Set rsDummy = New ADODB.Recordset
    rsDummy.Open s, Cn, adOpenKeyset, adLockOptimistic, adCmdText
    If Not rsDummy.EOF Then
        mParent_account = Trim(rsDummy!parent_account & "")
    End If
    rsDummy.Close
    
    
    
    sql = " select * from ACCOUNTS Where Account_Code= '" & Trim(mParent_account) & "'"
    Set rs = New ADODB.Recordset
    rs.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
   ' If rs.EOF Then
    
        
        
            StrNewAccountCode = AddNewAccount(Trim(rs!Account_code & ""), Trim$(mName), True, False, Trim$(mName), 1, False, False, False, , "", 0, 0, 0, 1, 0, IIf(rs("DepitOrCredit").value = 0, 0, 1), 0, 0, 0, 1, False)
            SaveBransh_UserAccount StrNewAccountCode
            'mSql = GetSqlQueryInsert(rs, ServerDb, "ACCOUNTS", "Account_ID", "", "", 0, 0, True)

        
       
    
    
    s = "SELECT Max(BoxID) MaxID  FROM " & mTable & " AS te "
    Set rsDummy = New ADODB.Recordset
    rsDummy.Open s, Cn, adOpenKeyset, adLockOptimistic, adCmdText
    If Not rsDummy.EOF Then
        mMaxId = val(rsDummy!MaxID & "")
    End If
    rsDummy.Close
    Set RsData = New ADODB.Recordset
    
    s = "Select * from TblBoxesData Where 1 = -1"
    RsData.Open s, Cn, adOpenKeyset, adLockOptimistic
    
      

    mSer = mMaxId

    mSer = mSer + 1
    RsData.AddNew
    'rsData!Fullcode = GetCode(mSer)
    'rsData!Code = GetCode(mSer)
    RsData!BoxID = mSer
    RsData!EmpID = mEmpId
    If Len(Trim(mName)) > 50 Then
        RsData!BoxName = Right(Trim(mName), 50)
    Else
        RsData!BoxName = Trim(mName)
    End If

    RsData!BoxNamee = Right(Trim(mName), 50)
    
  
    RsData!Type = 1
    
    RsData!BranchID = val(branch_id)
    RsData!Account_code = Trim(StrNewAccountCode)
    RsData!parent_account = Trim(mParent_account)
    
    'rsData!BranchID = val(rsDummy!BranchID & "")
    RsData.update
    grdExcel.TextMatrix(mRow, grdExcel.ColIndex("Account_Code")) = StrNewAccountCode


    

End Sub



Private Sub DBCboClientName_Click(Area As Integer)
Dim Fullcode As String


 GetCustomersDetail val(DBCboClientName.BoundText), , Fullcode, 1
  TxtSearchCode.Text = Fullcode
    
End Sub


Private Sub CheckAllCustomer_Click()
    If Me.CheckAllCustomer.value = vbChecked Then
        DBCboClientName.Enabled = False
        DBCboClientName.BoundText = 0
        CmdSelectCus.Enabled = True
        ChekCustomer.value = vbUnchecked
    End If
End Sub
Private Sub CheckAllEMp_Click()
    If Me.CheckAllEMp.value = vbChecked Then
        DcbEmployee.Enabled = False
        DcbEmployee.BoundText = 0
        CmdSelectEmp.Enabled = True
        CheckEmp.value = vbUnchecked
    End If
End Sub
Private Sub CheckEmp_Click()
    If Me.CheckEmp.value = vbChecked Then
        DcbEmployee.Enabled = True
        CmdSelectEmp.Enabled = False
        CheckAllEMp.value = vbUnchecked
        CurrenrEmployeeIDs.Text = ""
    End If
End Sub
Private Sub ChekCustomer_Click()
    If Me.ChekCustomer.value = vbChecked Then
        If Option1(2).value = True Then
            Reload 1
        Else
            Reload 2
        End If
        DBCboClientName.Enabled = True
        CmdSelectCus.Enabled = False
        CheckAllCustomer.value = vbUnchecked
        
    Else
        Reload 2
    End If
End Sub

Sub SaveBransh_UserAccount(Optional StrNewAccountCode As String)
Dim i As Integer
Dim sql As String
Dim Rs3 As ADODB.Recordset

sql = "Select * from  TblAccountBranch where 1=-1"
Set Rs3 = New ADODB.Recordset
Rs3.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText

Rs3.AddNew
Rs3("BranchID").value = branch_id
Rs3("Account_Code").value = Trim(StrNewAccountCode)
Rs3.update




    sql = "Select * from  TblAccountUser where 1=-1"
    Set Rs3 = New ADODB.Recordset
    Rs3.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
    
    Rs3.AddNew
    Rs3("UserID").value = user_id
    Rs3("Account_Code").value = Trim(StrNewAccountCode)
    Rs3.update


End Sub
Sub Reload(Optional Typ As Integer = -1)
    Dim Dcombos As ClsDataCombos
    Set Dcombos = New ClsDataCombos
    Dcombos.ClearMyDataCombo DBCboClientName
   Dcombos.GetCustomersSuppliers Typ, Me.DBCboClientName, True
End Sub
