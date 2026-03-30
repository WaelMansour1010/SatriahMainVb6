VERSION 5.00
Object = "{C115893A-A3BF-43AF-B28D-69DB846077F3}#1.0#0"; "vsflex8u.ocx"
Object = "{0AFE7BE0-11B7-4A3E-978D-D4501E9A57FE}#1.0#0"; "c1sizer.ocx"
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{FE5DCFAD-BC1D-11D2-94CF-004005455FAA}#1.4#0"; "ImpulseButton.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "mscomctl.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#12.0#0"; "SuiteCtrls.ocx"
Begin VB.Form FrmVizitScreen 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   9780
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   18660
   Icon            =   "FrmVizitScreen.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   RightToLeft     =   -1  'True
   ScaleHeight     =   9780
   ScaleWidth      =   18660
   ShowInTaskbar   =   0   'False
   Begin C1SizerLibCtl.C1Tab TabMain 
      Height          =   9585
      Left            =   0
      TabIndex        =   0
      Top             =   120
      Width           =   18780
      _cx             =   33126
      _cy             =   16907
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
      Caption         =   $"FrmVizitScreen.frx":57E2
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
         Height          =   9210
         Index           =   1
         Left            =   -21435
         TabIndex        =   1
         TabStop         =   0   'False
         Top             =   45
         Width           =   18690
         _cx             =   32967
         _cy             =   16245
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
            Height          =   690
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
                     Picture         =   "FrmVizitScreen.frx":586C
                     Key             =   "CompanyName"
                  EndProperty
                  BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                     Picture         =   "FrmVizitScreen.frx":5C06
                     Key             =   "Ser"
                  EndProperty
                  BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                     Picture         =   "FrmVizitScreen.frx":5FA0
                     Key             =   "Vac_Name"
                  EndProperty
                  BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                     Picture         =   "FrmVizitScreen.frx":633A
                     Key             =   "ShareCount"
                  EndProperty
                  BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                     Picture         =   "FrmVizitScreen.frx":66D4
                     Key             =   "Dis_Count"
                  EndProperty
                  BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                     Picture         =   "FrmVizitScreen.frx":6A6E
                     Key             =   "Bouns"
                  EndProperty
                  BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                     Picture         =   "FrmVizitScreen.frx":6E08
                     Key             =   "SharesValue"
                  EndProperty
                  BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                     Picture         =   "FrmVizitScreen.frx":73A2
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
               ButtonImage     =   "FrmVizitScreen.frx":773C
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
               ButtonImage     =   "FrmVizitScreen.frx":7AD6
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
               ButtonImage     =   "FrmVizitScreen.frx":7E70
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
               ButtonImage     =   "FrmVizitScreen.frx":820A
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
            Height          =   3675
            Left            =   1680
            RightToLeft     =   -1  'True
            TabIndex        =   5
            Top             =   5130
            Width           =   16740
            Begin VB.ComboBox CmbType 
               BackColor       =   &H80000018&
               Height          =   315
               ItemData        =   "FrmVizitScreen.frx":85A4
               Left            =   2280
               List            =   "FrmVizitScreen.frx":85B4
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
                  Format          =   215810049
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
                  ButtonImage     =   "FrmVizitScreen.frx":85CD
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
                  ButtonImage     =   "FrmVizitScreen.frx":8967
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
                  ButtonImage     =   "FrmVizitScreen.frx":F1C9
                  ColorButton     =   14871017
                  DrawFocusRectangle=   0   'False
               End
            End
         End
         Begin VSFlex8UCtl.VSFlexGrid FgItems 
            Height          =   9090
            Index           =   0
            Left            =   25935
            TabIndex        =   2
            Top             =   795
            Width           =   18555
            _cx             =   32729
            _cy             =   16034
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
            FormatString    =   $"FrmVizitScreen.frx":F563
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
            Height          =   1320
            Left            =   0
            TabIndex        =   43
            TabStop         =   0   'False
            Top             =   8010
            Width           =   16725
            _cx             =   29501
            _cy             =   2328
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
               ButtonImage     =   "FrmVizitScreen.frx":F623
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
               ButtonImage     =   "FrmVizitScreen.frx":F9BD
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
               ButtonImage     =   "FrmVizitScreen.frx":FD57
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
               ButtonImage     =   "FrmVizitScreen.frx":100F1
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
               ButtonImage     =   "FrmVizitScreen.frx":1048B
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
               ButtonImage     =   "FrmVizitScreen.frx":10A25
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
               ButtonImage     =   "FrmVizitScreen.frx":10DBF
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
               ButtonImage     =   "FrmVizitScreen.frx":11159
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
               ButtonImage     =   "FrmVizitScreen.frx":114F3
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
               ButtonImage     =   "FrmVizitScreen.frx":17D55
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
            Height          =   3570
            Left            =   135
            TabIndex        =   58
            Top             =   780
            Width           =   18285
            _cx             =   32253
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
            Cols            =   7
            FixedRows       =   1
            FixedCols       =   1
            RowHeightMin    =   320
            RowHeightMax    =   0
            ColWidthMin     =   0
            ColWidthMax     =   0
            ExtendLastCol   =   -1  'True
            FormatString    =   $"FrmVizitScreen.frx":1E5B7
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
         Height          =   9210
         Index           =   0
         Left            =   -21135
         TabIndex        =   3
         TabStop         =   0   'False
         Top             =   45
         Width           =   18690
         _cx             =   32967
         _cy             =   16245
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
         Begin VB.TextBox TXTIban 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFC0&
            Height          =   315
            Left            =   0
            Locked          =   -1  'True
            RightToLeft     =   -1  'True
            TabIndex        =   573
            Top             =   0
            Visible         =   0   'False
            Width           =   1455
         End
         Begin VB.TextBox TXTOrDer_no2 
            Alignment       =   1  'Right Justify
            Height          =   330
            Left            =   7245
            TabIndex        =   182
            Top             =   945
            Width           =   1680
         End
         Begin VB.TextBox TxtNoteSerial1 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFC0&
            Height          =   330
            Left            =   14700
            Locked          =   -1  'True
            RightToLeft     =   -1  'True
            TabIndex        =   129
            Top             =   1035
            Width           =   2010
         End
         Begin VB.CommandButton cmdPrintNote 
            Caption         =   "ÿ»«⁄Â «·ﬁÌœ"
            Height          =   480
            Left            =   3900
            RightToLeft     =   -1  'True
            TabIndex        =   127
            Top             =   7350
            Width           =   2790
         End
         Begin VB.CommandButton cmdDelNote 
            Caption         =   "Õ–› «·ﬁÌœ "
            Height          =   480
            Left            =   12000
            RightToLeft     =   -1  'True
            TabIndex        =   126
            Top             =   7290
            Visible         =   0   'False
            Width           =   3345
         End
         Begin VB.CommandButton CmdCreateV2 
            Caption         =   "≈‰‘«¡ «·ﬁÌœ "
            Height          =   480
            Left            =   15480
            RightToLeft     =   -1  'True
            TabIndex        =   125
            Top             =   7290
            Visible         =   0   'False
            Width           =   2940
         End
         Begin VB.TextBox TxtNoteSerial 
            Alignment       =   1  'Right Justify
            Enabled         =   0   'False
            Height          =   480
            Left            =   6690
            Locked          =   -1  'True
            RightToLeft     =   -1  'True
            TabIndex        =   124
            Top             =   7260
            Width           =   3765
         End
         Begin VB.TextBox Text1 
            Alignment       =   1  'Right Justify
            Height          =   330
            Left            =   660
            RightToLeft     =   -1  'True
            TabIndex        =   123
            Top             =   7485
            Visible         =   0   'False
            Width           =   2520
         End
         Begin VB.TextBox TxtCashCustomerName 
            Alignment       =   1  'Right Justify
            Height          =   360
            Left            =   2520
            Locked          =   -1  'True
            TabIndex        =   121
            Top             =   1410
            Width           =   3345
         End
         Begin VB.TextBox TXTOrDer_no 
            Alignment       =   1  'Right Justify
            Height          =   330
            Index           =   1
            Left            =   6420
            TabIndex        =   120
            Top             =   570
            Visible         =   0   'False
            Width           =   1545
         End
         Begin VB.ComboBox DcbType 
            Height          =   315
            Left            =   12000
            RightToLeft     =   -1  'True
            Style           =   2  'Dropdown List
            TabIndex        =   108
            Top             =   9180
            Visible         =   0   'False
            Width           =   1530
         End
         Begin VB.ComboBox DCOPrType 
            Height          =   315
            Left            =   14505
            RightToLeft     =   -1  'True
            Style           =   2  'Dropdown List
            TabIndex        =   107
            Top             =   9210
            Visible         =   0   'False
            Width           =   2220
         End
         Begin VB.ComboBox DcbyearFactor 
            Enabled         =   0   'False
            Height          =   315
            Left            =   13545
            Locked          =   -1  'True
            RightToLeft     =   -1  'True
            TabIndex        =   106
            Top             =   1980
            Width           =   2655
         End
         Begin VB.TextBox TxtPlatNo 
            Alignment       =   1  'Right Justify
            Height          =   330
            Left            =   690
            Locked          =   -1  'True
            TabIndex        =   105
            Top             =   1980
            Width           =   2790
         End
         Begin VB.TextBox TxtManualNo2 
            Alignment       =   1  'Right Justify
            BackColor       =   &H0080FFFF&
            Height          =   300
            Index           =   2
            Left            =   9345
            Locked          =   -1  'True
            TabIndex        =   104
            Top             =   2010
            Width           =   2655
         End
         Begin VB.TextBox TxtManualNo2 
            Alignment       =   1  'Right Justify
            BackColor       =   &H0080FFFF&
            Height          =   300
            Index           =   1
            Left            =   5310
            Locked          =   -1  'True
            TabIndex        =   103
            Top             =   2010
            Width           =   1935
         End
         Begin VB.TextBox TXTOrDer_no 
            Alignment       =   1  'Right Justify
            Height          =   330
            Index           =   0
            Left            =   3630
            TabIndex        =   101
            Top             =   975
            Width           =   1680
         End
         Begin VB.ComboBox CBoBasedON 
            Height          =   315
            ItemData        =   "FrmVizitScreen.frx":1E6C1
            Left            =   8925
            List            =   "FrmVizitScreen.frx":1E6C3
            RightToLeft     =   -1  'True
            Style           =   2  'Dropdown List
            TabIndex        =   100
            Top             =   960
            Width           =   1800
         End
         Begin VB.TextBox TxtNoteID 
            Alignment       =   1  'Right Justify
            Height          =   330
            Left            =   4890
            RightToLeft     =   -1  'True
            TabIndex        =   82
            Top             =   7770
            Visible         =   0   'False
            Width           =   2640
         End
         Begin VB.TextBox TxtSerial1 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFFC0&
            Enabled         =   0   'False
            Height          =   375
            Index           =   1
            Left            =   0
            RightToLeft     =   -1  'True
            TabIndex        =   73
            Top             =   1350
            Visible         =   0   'False
            Width           =   2370
         End
         Begin VB.TextBox TxtRemarks 
            Alignment       =   2  'Center
            Height          =   570
            Left            =   135
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   72
            Top             =   2805
            Width           =   16455
         End
         Begin VB.Frame Fra_Header 
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   0  'None
            Height          =   600
            Index           =   0
            Left            =   0
            RightToLeft     =   -1  'True
            TabIndex        =   60
            Top             =   0
            Width           =   20790
            Begin VB.CommandButton cmdcreate 
               Caption         =   "”œ«œ"
               Height          =   435
               Index           =   5
               Left            =   5250
               RightToLeft     =   -1  'True
               TabIndex        =   577
               Top             =   -90
               Visible         =   0   'False
               Width           =   1545
            End
            Begin VB.TextBox txtPassword 
               Height          =   315
               IMEMode         =   3  'DISABLE
               Left            =   4110
               PasswordChar    =   "*"
               TabIndex        =   576
               Top             =   0
               Width           =   1110
            End
            Begin VB.ComboBox DefaultInvoicetype 
               Height          =   315
               ItemData        =   "FrmVizitScreen.frx":1E6C5
               Left            =   8580
               List            =   "FrmVizitScreen.frx":1E6C7
               RightToLeft     =   -1  'True
               Style           =   2  'Dropdown List
               TabIndex        =   567
               Top             =   120
               Width           =   1890
            End
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
                     Picture         =   "FrmVizitScreen.frx":1E6C9
                     Key             =   "CompanyName"
                  EndProperty
                  BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                     Picture         =   "FrmVizitScreen.frx":1EA63
                     Key             =   "Ser"
                  EndProperty
                  BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                     Picture         =   "FrmVizitScreen.frx":1EDFD
                     Key             =   "Vac_Name"
                  EndProperty
                  BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                     Picture         =   "FrmVizitScreen.frx":1F197
                     Key             =   "ShareCount"
                  EndProperty
                  BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                     Picture         =   "FrmVizitScreen.frx":1F531
                     Key             =   "Dis_Count"
                  EndProperty
                  BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                     Picture         =   "FrmVizitScreen.frx":1F8CB
                     Key             =   "Bouns"
                  EndProperty
                  BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                     Picture         =   "FrmVizitScreen.frx":1FC65
                     Key             =   "SharesValue"
                  EndProperty
                  BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                     Picture         =   "FrmVizitScreen.frx":201FF
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
               ButtonImage     =   "FrmVizitScreen.frx":20599
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
               ButtonImage     =   "FrmVizitScreen.frx":20933
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
               ButtonImage     =   "FrmVizitScreen.frx":20CCD
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
               ButtonImage     =   "FrmVizitScreen.frx":21067
               ColorButton     =   14871017
               DrawFocusRectangle=   0   'False
               DisabledImageStyle=   1
            End
            Begin MSDataListLib.DataCombo DCDocTypes 
               Height          =   315
               Left            =   10650
               TabIndex        =   568
               Top             =   120
               Width           =   1245
               _ExtentX        =   2196
               _ExtentY        =   556
               _Version        =   393216
               BackColor       =   16777215
               Text            =   ""
               RightToLeft     =   -1  'True
            End
            Begin VB.Image ImgFavorites 
               Height          =   390
               Left            =   7560
               Picture         =   "FrmVizitScreen.frx":21401
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
               Left            =   13530
               RightToLeft     =   -1  'True
               TabIndex        =   70
               Top             =   90
               Width           =   2640
            End
         End
         Begin VSFlex8UCtl.VSFlexGrid FgItems 
            Height          =   9090
            Index           =   1
            Left            =   26070
            TabIndex        =   4
            Top             =   795
            Width           =   18420
            _cx             =   32491
            _cy             =   16034
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
            FormatString    =   $"FrmVizitScreen.frx":25069
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
            Height          =   315
            Left            =   11865
            TabIndex        =   74
            Top             =   1035
            Width           =   1800
            _ExtentX        =   3175
            _ExtentY        =   556
            _Version        =   393216
            Format          =   191168513
            CurrentDate     =   38784
         End
         Begin MSDataListLib.DataCombo dcBranch 
            Bindings        =   "FrmVizitScreen.frx":25129
            Height          =   315
            Index           =   1
            Left            =   0
            TabIndex        =   75
            Top             =   885
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
            Top             =   1440
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
            Top             =   8115
            Width           =   3345
            _ExtentX        =   5900
            _ExtentY        =   556
            _Version        =   393216
            Enabled         =   0   'False
            Text            =   ""
            RightToLeft     =   -1  'True
         End
         Begin ImpulseButton.ISButton btn_New 
            Height          =   300
            Index           =   1
            Left            =   14235
            TabIndex        =   84
            Top             =   8835
            Width           =   975
            _ExtentX        =   1720
            _ExtentY        =   529
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
            ButtonImage     =   "FrmVizitScreen.frx":2513E
            ColorButton     =   14871017
            DrawFocusRectangle=   0   'False
         End
         Begin ImpulseButton.ISButton btn_Save 
            Height          =   330
            Index           =   1
            Left            =   12000
            TabIndex        =   85
            Top             =   8805
            Width           =   975
            _ExtentX        =   1720
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
            ButtonImage     =   "FrmVizitScreen.frx":254D8
            ColorButton     =   14871017
            DrawFocusRectangle=   0   'False
         End
         Begin ImpulseButton.ISButton btn_Modify 
            Height          =   240
            Index           =   1
            Left            =   12975
            TabIndex        =   86
            Top             =   8835
            Width           =   960
            _ExtentX        =   1693
            _ExtentY        =   423
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
            ButtonImage     =   "FrmVizitScreen.frx":25872
            ColorButton     =   14871017
            DrawFocusRectangle=   0   'False
         End
         Begin ImpulseButton.ISButton Btn_Undo 
            Height          =   240
            Index           =   1
            Left            =   11025
            TabIndex        =   87
            Top             =   8835
            Width           =   975
            _ExtentX        =   1720
            _ExtentY        =   423
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
            ButtonImage     =   "FrmVizitScreen.frx":25C0C
            ColorButton     =   14871017
            DrawFocusRectangle=   0   'False
         End
         Begin ImpulseButton.ISButton btn_Delete 
            Height          =   330
            Index           =   1
            Left            =   10230
            TabIndex        =   88
            Top             =   8820
            Width           =   855
            _ExtentX        =   1508
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
            ButtonImage     =   "FrmVizitScreen.frx":25FA6
            ColorButton     =   14871017
            DrawFocusRectangle=   0   'False
         End
         Begin ImpulseButton.ISButton Btn_Update 
            Height          =   360
            Index           =   1
            Left            =   11310
            TabIndex        =   89
            TabStop         =   0   'False
            ToolTipText     =   " ÕœÌÀ ﬁ«⁄œ… «·»Ì«‰« "
            Top             =   8460
            Visible         =   0   'False
            Width           =   1110
            _ExtentX        =   1958
            _ExtentY        =   635
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
            ButtonImage     =   "FrmVizitScreen.frx":26540
            ColorButton     =   14871017
            DrawFocusRectangle=   0   'False
         End
         Begin ImpulseButton.ISButton btn_Cancel 
            Height          =   330
            Index           =   1
            Left            =   5865
            TabIndex        =   90
            Top             =   8775
            Width           =   960
            _ExtentX        =   1693
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
            ButtonImage     =   "FrmVizitScreen.frx":268DA
            ColorButton     =   14871017
            DrawFocusRectangle=   0   'False
            DisabledImageStyle=   1
         End
         Begin ImpulseButton.ISButton Btn_Print 
            Height          =   375
            Index           =   1
            Left            =   8520
            TabIndex        =   91
            TabStop         =   0   'False
            ToolTipText     =   "ÿ»«⁄… «·»Ì«‰«  "
            Top             =   8745
            Width           =   1380
            _ExtentX        =   2434
            _ExtentY        =   661
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
            ButtonImage     =   "FrmVizitScreen.frx":26C74
            ColorButton     =   14871017
            DrawFocusRectangle=   0   'False
            DisabledImageStyle=   1
         End
         Begin ImpulseButton.ISButton btn_Query 
            Height          =   420
            Index           =   1
            Left            =   6975
            TabIndex        =   92
            TabStop         =   0   'False
            ToolTipText     =   "(Ctrl+F)  ··»ÕÀ ≈÷€ÿ Â–« «·„› «Õ √Ê ≈÷€ÿ "
            Top             =   8700
            Width           =   1395
            _ExtentX        =   2461
            _ExtentY        =   741
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
            ButtonImage     =   "FrmVizitScreen.frx":2D4D6
            ColorButton     =   14871017
            DrawFocusRectangle=   0   'False
         End
         Begin ImpulseButton.ISButton Cmd_DeleteRow 
            Height          =   315
            Index           =   1
            Left            =   2520
            TabIndex        =   93
            Top             =   7950
            Width           =   1800
            _ExtentX        =   3175
            _ExtentY        =   556
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
            ButtonImage     =   "FrmVizitScreen.frx":2D870
            ColorButton     =   14871017
            DrawFocusRectangle=   0   'False
         End
         Begin ImpulseButton.ISButton Cmd_DeleteAll 
            Height          =   315
            Index           =   1
            Left            =   270
            TabIndex        =   94
            Top             =   7935
            Width           =   2100
            _ExtentX        =   3704
            _ExtentY        =   556
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
            ButtonImage     =   "FrmVizitScreen.frx":2DE0A
            ColorButton     =   14871017
            DrawFocusRectangle=   0   'False
         End
         Begin MSDataListLib.DataCombo DCEquipments 
            Height          =   315
            Left            =   11595
            TabIndex        =   109
            Top             =   2325
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
            Bindings        =   "FrmVizitScreen.frx":2E3A4
            Height          =   315
            Left            =   13545
            TabIndex        =   110
            Top             =   1575
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
            Bindings        =   "FrmVizitScreen.frx":2E3B9
            Height          =   315
            Left            =   690
            TabIndex        =   131
            Top             =   2355
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
            Height          =   3945
            Left            =   135
            TabIndex        =   134
            Top             =   3360
            Width           =   18555
            _cx             =   32729
            _cy             =   6959
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
               Height          =   3570
               Index           =   2
               Left            =   45
               TabIndex        =   135
               TabStop         =   0   'False
               Top             =   45
               Width           =   18465
               _cx             =   32570
               _cy             =   6297
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
               Begin VB.TextBox txt_Currency_rate 
                  Alignment       =   1  'Right Justify
                  Height          =   315
                  Left            =   0
                  RightToLeft     =   -1  'True
                  TabIndex        =   569
                  Text            =   "1"
                  Top             =   15
                  Width           =   765
               End
               Begin VB.PictureBox Picture1 
                  Height          =   90
                  Left            =   0
                  ScaleHeight     =   30
                  ScaleWidth      =   1710
                  TabIndex        =   566
                  Top             =   -120
                  Width           =   1770
               End
               Begin VB.Frame Frame6 
                  Caption         =   "«·«Ã„«·Ï «·⁄«„"
                  Height          =   2265
                  Left            =   0
                  RightToLeft     =   -1  'True
                  TabIndex        =   171
                  Top             =   1230
                  Width           =   3915
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
                  Height          =   2010
                  Left            =   4350
                  RightToLeft     =   -1  'True
                  TabIndex        =   154
                  Top             =   1545
                  Width           =   6705
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
                  Height          =   300
                  Left            =   2625
                  Locked          =   -1  'True
                  RightToLeft     =   -1  'True
                  TabIndex        =   150
                  Top             =   3570
                  Width           =   3780
               End
               Begin VB.Frame Frame5 
                  Caption         =   "»Ì«‰«   ›Ê« Ì— «·„»Ì⁄« "
                  Height          =   1560
                  Left            =   4350
                  RightToLeft     =   -1  'True
                  TabIndex        =   139
                  Top             =   0
                  Width           =   6705
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
                  Height          =   3465
                  Index           =   2
                  Left            =   27195
                  TabIndex        =   136
                  Top             =   675
                  Width           =   18315
                  _cx             =   32306
                  _cy             =   6112
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
                  FormatString    =   $"FrmVizitScreen.frx":2E3CE
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
                  Height          =   3270
                  Left            =   11055
                  TabIndex        =   151
                  Top             =   90
                  Width           =   7110
                  _cx             =   12541
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
                  FormatString    =   $"FrmVizitScreen.frx":2E48E
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
               Begin MSDataListLib.DataCombo DcCurrency 
                  Height          =   315
                  Left            =   840
                  TabIndex        =   570
                  Top             =   0
                  Width           =   855
                  _ExtentX        =   1508
                  _ExtentY        =   556
                  _Version        =   393216
                  Text            =   ""
                  RightToLeft     =   -1  'True
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  BackStyle       =   0  'Transparent
                  Caption         =   "««·⁄„·…"
                  Height          =   300
                  Index           =   32
                  Left            =   1170
                  RightToLeft     =   -1  'True
                  TabIndex        =   571
                  Top             =   30
                  Width           =   1080
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "«·’«›Ï »⁄œ «·ﬁÌ„… «·„÷«›…"
                  Height          =   240
                  Index           =   47
                  Left            =   6405
                  TabIndex        =   152
                  Top             =   3600
                  Width           =   2175
               End
            End
            Begin C1SizerLibCtl.C1Elastic ELe 
               Height          =   3570
               Index           =   3
               Left            =   19200
               TabIndex        =   137
               TabStop         =   0   'False
               Top             =   45
               Width           =   18465
               _cx             =   32570
               _cy             =   6297
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
                  Height          =   3435
                  Index           =   3
                  Left            =   27045
                  TabIndex        =   138
                  Top             =   765
                  Width           =   18315
                  _cx             =   32306
                  _cy             =   6059
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
                  FormatString    =   $"FrmVizitScreen.frx":2E59D
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
                  Height          =   3465
                  Left            =   435
                  TabIndex        =   153
                  Top             =   -60
                  Width           =   17445
                  _cx             =   30771
                  _cy             =   6112
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
                  FormatString    =   $"FrmVizitScreen.frx":2E65D
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
         Begin MSComCtl2.DTPicker txtDateRec 
            Height          =   300
            Left            =   3660
            TabIndex        =   572
            Top             =   660
            Width           =   1560
            _ExtentX        =   2752
            _ExtentY        =   529
            _Version        =   393216
            Enabled         =   0   'False
            Format          =   215875585
            CurrentDate     =   38784
         End
         Begin MSDataListLib.DataCombo cmbPaymentType 
            Height          =   315
            Left            =   5250
            TabIndex        =   574
            Top             =   2370
            Width           =   4605
            _ExtentX        =   8123
            _ExtentY        =   556
            _Version        =   393216
            ListField       =   "6"
            Text            =   ""
            RightToLeft     =   -1  'True
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "‰Ê⁄ «·œ›⁄"
            Height          =   255
            Index           =   37
            Left            =   9855
            RightToLeft     =   -1  'True
            TabIndex        =   575
            Top             =   2400
            Width           =   1260
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "—ﬁ„ «„— «·«’·«Õ"
            Height          =   270
            Index           =   20
            Left            =   4890
            TabIndex        =   133
            Top             =   1005
            Width           =   1800
         End
         Begin VB.Label lblModel 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "«·ÿ—«“ "
            Height          =   270
            Left            =   3765
            TabIndex        =   132
            Top             =   2355
            Width           =   1125
         End
         Begin VB.Label Label5 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "—ﬁ„ «·”‰œ"
            Height          =   270
            Index           =   1
            Left            =   17445
            TabIndex        =   130
            Top             =   1005
            Width           =   975
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            BackStyle       =   0  'Transparent
            Caption         =   "—ﬁ„ «·ﬁÌœ"
            Height          =   435
            Index           =   14
            Left            =   10455
            RightToLeft     =   -1  'True
            TabIndex        =   128
            Top             =   7380
            Width           =   1260
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "«”„ «·⁄„Ì· «·‰ﬁœÏ"
            Height          =   330
            Index           =   33
            Left            =   6135
            TabIndex        =   122
            Top             =   1470
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
            Top             =   5445
            Width           =   705
         End
         Begin VB.Label LblYear 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "„ÊœÌ· «·„⁄œÂ/«·”Ì«—…"
            Height          =   270
            Left            =   16530
            TabIndex        =   118
            Top             =   1920
            Width           =   1920
         End
         Begin VB.Label LblPla 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "—ﬁ„ «··ÊÕ…"
            Height          =   270
            Left            =   3900
            TabIndex        =   117
            Top             =   2010
            Width           =   1260
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "«·‰Ê⁄"
            Height          =   300
            Index           =   123
            Left            =   13530
            RightToLeft     =   -1  'True
            TabIndex        =   116
            Top             =   9435
            Visible         =   0   'False
            Width           =   840
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "‰Ê⁄ «·⁄„·Ì…"
            Height          =   300
            Index           =   124
            Left            =   17160
            RightToLeft     =   -1  'True
            TabIndex        =   115
            Top             =   9165
            Visible         =   0   'False
            Width           =   840
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "«·„⁄œÂ/«·”Ì«—…"
            Height          =   255
            Index           =   125
            Left            =   16200
            RightToLeft     =   -1  'True
            TabIndex        =   114
            Top             =   2355
            Visible         =   0   'False
            Width           =   1260
         End
         Begin VB.Label lbltycar 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "‰Ê⁄ «·„⁄œÂ/«·”Ì«—…"
            Height          =   255
            Left            =   16905
            TabIndex        =   113
            Top             =   1545
            Width           =   1560
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            BackStyle       =   0  'Transparent
            Caption         =   "—ﬁ„ «·‘«”ÌÂ"
            Height          =   210
            Index           =   119
            Left            =   11715
            TabIndex        =   112
            Top             =   2010
            Width           =   1665
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            BackStyle       =   0  'Transparent
            Caption         =   "⁄œ«œ «·ﬂÌ·Ê „ —"
            Height          =   210
            Index           =   118
            Left            =   7245
            TabIndex        =   111
            Top             =   2010
            Width           =   1545
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "»‰«¡ ⁄·Ï"
            Height          =   270
            Index           =   56
            Left            =   10725
            TabIndex        =   102
            Top             =   990
            Width           =   990
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Õ—— »Ê«”ÿ…  "
            Height          =   330
            Index           =   8
            Left            =   17160
            TabIndex        =   99
            Top             =   8040
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
            Top             =   8460
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
            Top             =   8460
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
            Top             =   8445
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
            Top             =   8445
            Width           =   1245
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   " «—ÌŒ «·”‰œ"
            Height          =   285
            Index           =   2
            Left            =   13635
            TabIndex        =   81
            Top             =   1035
            Width           =   990
         End
         Begin VB.Label lbl 
            Alignment       =   2  'Center
            BackColor       =   &H00E2E9E9&
            Caption         =   "—ﬁ„ «·⁄ﬁœ"
            Height          =   270
            Index           =   4
            Left            =   18690
            RightToLeft     =   -1  'True
            TabIndex        =   80
            Top             =   1335
            Width           =   1110
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "«·›—⁄"
            Height          =   270
            Index           =   7
            Left            =   2370
            TabIndex        =   79
            Top             =   915
            Width           =   705
         End
         Begin VB.Label Label1 
            Alignment       =   2  'Center
            BackColor       =   &H00E2E9E9&
            BackStyle       =   0  'Transparent
            Caption         =   "„·«ÕŸ«    "
            Height          =   300
            Index           =   11
            Left            =   18555
            RightToLeft     =   -1  'True
            TabIndex        =   78
            Top             =   3210
            Width           =   1245
         End
         Begin VB.Label lbl 
            Alignment       =   2  'Center
            BackColor       =   &H00E2E9E9&
            Caption         =   "«·⁄„Ì·"
            Height          =   495
            Index           =   15
            Left            =   11715
            TabIndex        =   77
            Top             =   1470
            Width           =   1395
         End
         Begin VB.Label Label4 
            Alignment       =   1  'Right Justify
            Caption         =   "„·«ÕŸ« "
            Height          =   270
            Index           =   0
            Left            =   16725
            RightToLeft     =   -1  'True
            TabIndex        =   71
            Top             =   2985
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
         Height          =   9210
         Index           =   4
         Left            =   -20835
         TabIndex        =   183
         TabStop         =   0   'False
         Top             =   45
         Width           =   18690
         _cx             =   32967
         _cy             =   16245
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
            Height          =   1740
            Index           =   2
            Left            =   135
            RightToLeft     =   -1  'True
            TabIndex        =   192
            Top             =   4785
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
               ItemData        =   "FrmVizitScreen.frx":2E881
               Left            =   2280
               List            =   "FrmVizitScreen.frx":2E891
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
            Height          =   615
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
                     Picture         =   "FrmVizitScreen.frx":2E8AA
                     Key             =   "CompanyName"
                  EndProperty
                  BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                     Picture         =   "FrmVizitScreen.frx":2EC44
                     Key             =   "Ser"
                  EndProperty
                  BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                     Picture         =   "FrmVizitScreen.frx":2EFDE
                     Key             =   "Vac_Name"
                  EndProperty
                  BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                     Picture         =   "FrmVizitScreen.frx":2F378
                     Key             =   "ShareCount"
                  EndProperty
                  BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                     Picture         =   "FrmVizitScreen.frx":2F712
                     Key             =   "Dis_Count"
                  EndProperty
                  BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                     Picture         =   "FrmVizitScreen.frx":2FAAC
                     Key             =   "Bouns"
                  EndProperty
                  BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                     Picture         =   "FrmVizitScreen.frx":2FE46
                     Key             =   "SharesValue"
                  EndProperty
                  BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                     Picture         =   "FrmVizitScreen.frx":303E0
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
               ButtonImage     =   "FrmVizitScreen.frx":3077A
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
               ButtonImage     =   "FrmVizitScreen.frx":30B14
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
               ButtonImage     =   "FrmVizitScreen.frx":30EAE
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
               ButtonImage     =   "FrmVizitScreen.frx":31248
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
            Height          =   375
            Index           =   2
            Left            =   7245
            TabIndex        =   200
            Top             =   8040
            Width           =   1125
            _ExtentX        =   1984
            _ExtentY        =   661
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
            ButtonImage     =   "FrmVizitScreen.frx":315E2
            ColorButton     =   14871017
            DrawFocusRectangle=   0   'False
         End
         Begin ImpulseButton.ISButton btn_Save 
            Height          =   360
            Index           =   2
            Left            =   5310
            TabIndex        =   201
            Top             =   8010
            Width           =   825
            _ExtentX        =   1455
            _ExtentY        =   635
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
            ButtonImage     =   "FrmVizitScreen.frx":3197C
            ColorButton     =   14871017
            DrawFocusRectangle=   0   'False
         End
         Begin ImpulseButton.ISButton btn_Modify 
            Height          =   375
            Index           =   2
            Left            =   6270
            TabIndex        =   202
            Top             =   8040
            Width           =   975
            _ExtentX        =   1720
            _ExtentY        =   661
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
            ButtonImage     =   "FrmVizitScreen.frx":31D16
            ColorButton     =   14871017
            DrawFocusRectangle=   0   'False
         End
         Begin ImpulseButton.ISButton Btn_Undo 
            Height          =   360
            Index           =   2
            Left            =   4320
            TabIndex        =   203
            Top             =   8010
            Width           =   990
            _ExtentX        =   1746
            _ExtentY        =   635
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
            ButtonImage     =   "FrmVizitScreen.frx":320B0
            ColorButton     =   14871017
            DrawFocusRectangle=   0   'False
         End
         Begin ImpulseButton.ISButton btn_Delete 
            Height          =   375
            Index           =   2
            Left            =   3480
            TabIndex        =   204
            Top             =   8040
            Width           =   840
            _ExtentX        =   1482
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
            ButtonImage     =   "FrmVizitScreen.frx":3244A
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
            Top             =   6765
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
            ButtonImage     =   "FrmVizitScreen.frx":329E4
            ColorButton     =   14871017
            DrawFocusRectangle=   0   'False
         End
         Begin ImpulseButton.ISButton btn_Cancel 
            Height          =   360
            Index           =   2
            Left            =   135
            TabIndex        =   206
            Top             =   7980
            Width           =   975
            _ExtentX        =   1720
            _ExtentY        =   635
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
            ButtonImage     =   "FrmVizitScreen.frx":32D7E
            ColorButton     =   14871017
            DrawFocusRectangle=   0   'False
            DisabledImageStyle=   1
         End
         Begin ImpulseButton.ISButton Btn_Print 
            Height          =   465
            Index           =   2
            Left            =   2370
            TabIndex        =   207
            TabStop         =   0   'False
            ToolTipText     =   "ÿ»«⁄… «·»Ì«‰«  "
            Top             =   7980
            Width           =   975
            _ExtentX        =   1720
            _ExtentY        =   820
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
            ButtonImage     =   "FrmVizitScreen.frx":33118
            ColorButton     =   14871017
            DrawFocusRectangle=   0   'False
            DisabledImageStyle=   1
         End
         Begin ImpulseButton.ISButton btn_Query 
            Height          =   570
            Index           =   2
            Left            =   1530
            TabIndex        =   208
            TabStop         =   0   'False
            ToolTipText     =   "(Ctrl+F)  ··»ÕÀ ≈÷€ÿ Â–« «·„› «Õ √Ê ≈÷€ÿ "
            Top             =   7905
            Width           =   705
            _ExtentX        =   1244
            _ExtentY        =   1005
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
            ButtonImage     =   "FrmVizitScreen.frx":3997A
            ColorButton     =   14871017
            DrawFocusRectangle=   0   'False
         End
         Begin VSFlex8Ctl.VSFlexGrid Grid2 
            Height          =   3645
            Left            =   0
            TabIndex        =   209
            Top             =   885
            Width           =   8790
            _cx             =   15505
            _cy             =   6429
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
            FormatString    =   $"FrmVizitScreen.frx":39D14
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
            Height          =   240
            Index           =   7
            Left            =   6135
            RightToLeft     =   -1  'True
            TabIndex        =   217
            Top             =   7020
            Width           =   2100
         End
         Begin VB.Label Label2 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "⁄œœ «·”Ã·« :"
            Height          =   240
            Index           =   6
            Left            =   2235
            RightToLeft     =   -1  'True
            TabIndex        =   216
            Top             =   7020
            Width           =   2085
         End
         Begin VB.Label LabCurr_Rec 
            BackColor       =   &H00E2E9E9&
            Height          =   255
            Index           =   0
            Left            =   4320
            RightToLeft     =   -1  'True
            TabIndex        =   215
            Top             =   7035
            Width           =   1545
         End
         Begin VB.Label LabCount_Rec 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Height          =   255
            Index           =   0
            Left            =   690
            RightToLeft     =   -1  'True
            TabIndex        =   214
            Top             =   7035
            Width           =   1125
         End
         Begin VB.Label Label2 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "«·”Ã· «·Õ«·Ì:"
            Height          =   240
            Index           =   4
            Left            =   6690
            RightToLeft     =   -1  'True
            TabIndex        =   213
            Top             =   7515
            Width           =   1545
         End
         Begin VB.Label Label2 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "⁄œœ «·”Ã·« :"
            Height          =   240
            Index           =   5
            Left            =   2790
            RightToLeft     =   -1  'True
            TabIndex        =   212
            Top             =   7515
            Width           =   2100
         End
         Begin VB.Label LabCurr_Rec 
            BackColor       =   &H00E2E9E9&
            Height          =   255
            Index           =   2
            Left            =   4890
            RightToLeft     =   -1  'True
            TabIndex        =   211
            Top             =   7530
            Width           =   1530
         End
         Begin VB.Label LabCount_Rec 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Height          =   255
            Index           =   2
            Left            =   1245
            RightToLeft     =   -1  'True
            TabIndex        =   210
            Top             =   7530
            Width           =   1275
         End
      End
      Begin C1SizerLibCtl.C1Elastic ELe 
         Height          =   9210
         Index           =   5
         Left            =   -20535
         TabIndex        =   218
         TabStop         =   0   'False
         Top             =   45
         Width           =   18690
         _cx             =   32967
         _cy             =   16245
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
            Height          =   465
            Left            =   5715
            RightToLeft     =   -1  'True
            TabIndex        =   369
            Top             =   5850
            Width           =   1395
         End
         Begin VB.CommandButton Command1 
            Caption         =   "«‰‘«¡ «·«’‰«›"
            Height          =   480
            Left            =   1380
            RightToLeft     =   -1  'True
            TabIndex        =   252
            Top             =   5655
            Visible         =   0   'False
            Width           =   2805
         End
         Begin VB.ComboBox cmbFlag 
            Height          =   315
            Index           =   2
            ItemData        =   "FrmVizitScreen.frx":39DA3
            Left            =   1380
            List            =   "FrmVizitScreen.frx":39DA5
            RightToLeft     =   -1  'True
            Style           =   2  'Dropdown List
            TabIndex        =   251
            Top             =   5280
            Visible         =   0   'False
            Width           =   1830
         End
         Begin VB.ComboBox cmbFlag 
            Height          =   315
            Index           =   1
            ItemData        =   "FrmVizitScreen.frx":39DA7
            Left            =   3345
            List            =   "FrmVizitScreen.frx":39DA9
            RightToLeft     =   -1  'True
            Style           =   2  'Dropdown List
            TabIndex        =   250
            Top             =   5310
            Visible         =   0   'False
            Width           =   1815
         End
         Begin VB.ComboBox cmbFlag 
            Height          =   315
            Index           =   0
            ItemData        =   "FrmVizitScreen.frx":39DAB
            Left            =   3075
            List            =   "FrmVizitScreen.frx":39DAD
            RightToLeft     =   -1  'True
            TabIndex        =   249
            Text            =   "cmbFlag"
            Top             =   2985
            Width           =   1815
         End
         Begin VB.Frame Fra_Header 
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   0  'None
            Height          =   615
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
                     Picture         =   "FrmVizitScreen.frx":39DAF
                     Key             =   "CompanyName"
                  EndProperty
                  BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                     Picture         =   "FrmVizitScreen.frx":3A149
                     Key             =   "Ser"
                  EndProperty
                  BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                     Picture         =   "FrmVizitScreen.frx":3A4E3
                     Key             =   "Vac_Name"
                  EndProperty
                  BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                     Picture         =   "FrmVizitScreen.frx":3A87D
                     Key             =   "ShareCount"
                  EndProperty
                  BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                     Picture         =   "FrmVizitScreen.frx":3AC17
                     Key             =   "Dis_Count"
                  EndProperty
                  BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                     Picture         =   "FrmVizitScreen.frx":3AFB1
                     Key             =   "Bouns"
                  EndProperty
                  BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                     Picture         =   "FrmVizitScreen.frx":3B34B
                     Key             =   "SharesValue"
                  EndProperty
                  BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                     Picture         =   "FrmVizitScreen.frx":3B8E5
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
               ButtonImage     =   "FrmVizitScreen.frx":3BC7F
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
               ButtonImage     =   "FrmVizitScreen.frx":3C019
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
               ButtonImage     =   "FrmVizitScreen.frx":3C3B3
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
               ButtonImage     =   "FrmVizitScreen.frx":3C74D
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
            Height          =   4365
            Index           =   3
            Left            =   135
            RightToLeft     =   -1  'True
            TabIndex        =   219
            Top             =   750
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
            Height          =   360
            Index           =   3
            Left            =   7245
            TabIndex        =   234
            Top             =   8670
            Width           =   1125
            _ExtentX        =   1984
            _ExtentY        =   635
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
            ButtonImage     =   "FrmVizitScreen.frx":3CAE7
            ColorButton     =   14871017
            DrawFocusRectangle=   0   'False
         End
         Begin ImpulseButton.ISButton btn_Save 
            Height          =   360
            Index           =   3
            Left            =   5310
            TabIndex        =   235
            Top             =   8640
            Width           =   825
            _ExtentX        =   1455
            _ExtentY        =   635
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
            ButtonImage     =   "FrmVizitScreen.frx":3CE81
            ColorButton     =   14871017
            DrawFocusRectangle=   0   'False
         End
         Begin ImpulseButton.ISButton btn_Modify 
            Height          =   360
            Index           =   3
            Left            =   6270
            TabIndex        =   236
            Top             =   8670
            Width           =   975
            _ExtentX        =   1720
            _ExtentY        =   635
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
            ButtonImage     =   "FrmVizitScreen.frx":3D21B
            ColorButton     =   14871017
            DrawFocusRectangle=   0   'False
         End
         Begin ImpulseButton.ISButton Btn_Undo 
            Height          =   360
            Index           =   3
            Left            =   4320
            TabIndex        =   237
            Top             =   8640
            Width           =   990
            _ExtentX        =   1746
            _ExtentY        =   635
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
            ButtonImage     =   "FrmVizitScreen.frx":3D5B5
            ColorButton     =   14871017
            DrawFocusRectangle=   0   'False
         End
         Begin ImpulseButton.ISButton btn_Delete 
            Height          =   360
            Index           =   3
            Left            =   3480
            TabIndex        =   238
            Top             =   8670
            Width           =   840
            _ExtentX        =   1482
            _ExtentY        =   635
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
            ButtonImage     =   "FrmVizitScreen.frx":3D94F
            ColorButton     =   14871017
            DrawFocusRectangle=   0   'False
         End
         Begin ImpulseButton.ISButton Btn_Update 
            Height          =   240
            Index           =   3
            Left            =   2655
            TabIndex        =   239
            TabStop         =   0   'False
            ToolTipText     =   " ÕœÌÀ ﬁ«⁄œ… «·»Ì«‰« "
            Top             =   7515
            Visible         =   0   'False
            Width           =   1935
            _ExtentX        =   3413
            _ExtentY        =   423
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
            ButtonImage     =   "FrmVizitScreen.frx":3DEE9
            ColorButton     =   14871017
            DrawFocusRectangle=   0   'False
         End
         Begin ImpulseButton.ISButton btn_Cancel 
            Height          =   360
            Index           =   3
            Left            =   135
            TabIndex        =   240
            Top             =   8610
            Width           =   975
            _ExtentX        =   1720
            _ExtentY        =   635
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
            ButtonImage     =   "FrmVizitScreen.frx":3E283
            ColorButton     =   14871017
            DrawFocusRectangle=   0   'False
            DisabledImageStyle=   1
         End
         Begin ImpulseButton.ISButton Btn_Print 
            Height          =   465
            Index           =   3
            Left            =   2370
            TabIndex        =   241
            TabStop         =   0   'False
            ToolTipText     =   "ÿ»«⁄… «·»Ì«‰«  "
            Top             =   8610
            Width           =   975
            _ExtentX        =   1720
            _ExtentY        =   820
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
            ButtonImage     =   "FrmVizitScreen.frx":3E61D
            ColorButton     =   14871017
            DrawFocusRectangle=   0   'False
            DisabledImageStyle=   1
         End
         Begin ImpulseButton.ISButton btn_Query 
            Height          =   570
            Index           =   3
            Left            =   1530
            TabIndex        =   242
            TabStop         =   0   'False
            ToolTipText     =   "(Ctrl+F)  ··»ÕÀ ≈÷€ÿ Â–« «·„› «Õ √Ê ≈÷€ÿ "
            Top             =   8535
            Width           =   705
            _ExtentX        =   1244
            _ExtentY        =   1005
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
            ButtonImage     =   "FrmVizitScreen.frx":44E7F
            ColorButton     =   14871017
            DrawFocusRectangle=   0   'False
         End
         Begin C1SizerLibCtl.C1Tab C1Tab1 
            Height          =   7650
            Left            =   7110
            TabIndex        =   282
            Top             =   750
            Width           =   11445
            _cx             =   20188
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
               Height          =   7275
               Index           =   16
               Left            =   -12000
               TabIndex        =   283
               TabStop         =   0   'False
               Top             =   45
               Width           =   11355
               _cx             =   20029
               _cy             =   12832
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
                  Height          =   360
                  Left            =   8775
                  TabIndex        =   292
                  Top             =   2055
                  Visible         =   0   'False
                  Width           =   1545
               End
               Begin VB.TextBox TxtModel 
                  Alignment       =   2  'Center
                  Height          =   330
                  Left            =   2460
                  TabIndex        =   291
                  Top             =   1230
                  Width           =   645
               End
               Begin VB.TextBox TxtColorCode 
                  Alignment       =   2  'Center
                  Height          =   315
                  Left            =   1155
                  TabIndex        =   290
                  Top             =   1215
                  Width           =   780
               End
               Begin VB.TextBox TxtSize 
                  Alignment       =   2  'Center
                  Height          =   315
                  Left            =   0
                  TabIndex        =   289
                  Top             =   1215
                  Width           =   780
               End
               Begin VB.CommandButton cmdLoadFile 
                  Caption         =   " Õ„Ì· «·„·›..."
                  Height          =   285
                  Left            =   10590
                  TabIndex        =   288
                  Top             =   4635
                  Visible         =   0   'False
                  Width           =   765
               End
               Begin VB.TextBox txtFile 
                  Height          =   360
                  Left            =   9555
                  Locked          =   -1  'True
                  TabIndex        =   287
                  Top             =   2700
                  Visible         =   0   'False
                  Width           =   900
               End
               Begin VB.CommandButton cmdSelectFile 
                  Caption         =   " ÕœÌœ «·„·›..."
                  Height          =   240
                  Left            =   8265
                  RightToLeft     =   -1  'True
                  TabIndex        =   286
                  Top             =   1395
                  Visible         =   0   'False
                  Width           =   1155
               End
               Begin VB.CommandButton Command8 
                  Caption         =   " ÕœÌÀ «·«’‰«›"
                  Height          =   195
                  Left            =   8130
                  TabIndex        =   285
                  Top             =   1935
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
                  Top             =   255
                  Width           =   3105
               End
               Begin VSFlex8UCtl.VSFlexGrid FgItems 
                  Height          =   7170
                  Index           =   7
                  Left            =   13425
                  TabIndex        =   293
                  Top             =   570
                  Width           =   9675
                  _cx             =   17066
                  _cy             =   12647
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
                  FormatString    =   $"FrmVizitScreen.frx":45219
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
                  Top             =   720
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
                  Top             =   2760
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
                  Top             =   3195
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
                  Top             =   1575
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
                  Top             =   3630
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
                  Top             =   2370
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
                  Top             =   6030
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
                  Top             =   5085
                  Width           =   3105
                  _ExtentX        =   5477
                  _ExtentY        =   556
                  _Version        =   393216
                  Style           =   2
                  Text            =   ""
                  RightToLeft     =   -1  'True
               End
               Begin VSFlex8Ctl.VSFlexGrid tmpGrd 
                  Height          =   360
                  Left            =   8910
                  TabIndex        =   303
                  Top             =   3600
                  Visible         =   0   'False
                  Width           =   1035
                  _cx             =   1826
                  _cy             =   635
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
                  Top             =   6015
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
                  Index           =   19
                  Left            =   9945
                  TabIndex        =   306
                  Top             =   2190
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
                  Top             =   2025
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
                  Top             =   6585
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
                  Top             =   5565
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
                  Top             =   6120
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
                  Top             =   6780
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
                  Top             =   4665
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
                  Top             =   4185
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
                  Top             =   5175
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
                  Top             =   2790
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
                  Top             =   1410
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
                  Top             =   2265
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
                  Height          =   555
                  Index           =   36
                  Left            =   7620
                  TabIndex        =   353
                  Top             =   1395
                  Visible         =   0   'False
                  Width           =   645
               End
               Begin VB.Label lbl«”„«·ÊÕœ… 
                  Alignment       =   2  'Center
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "Cylinder"
                  Height          =   585
                  Index           =   35
                  Left            =   7620
                  TabIndex        =   352
                  Top             =   2340
                  Visible         =   0   'False
                  Width           =   645
               End
               Begin VB.Label lbl«”„«·ÊÕœ… 
                  Alignment       =   2  'Center
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "Division"
                  Height          =   840
                  Index           =   33
                  Left            =   7620
                  TabIndex        =   351
                  Top             =   2925
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
                  Top             =   4710
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
                  Top             =   4185
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
                  Top             =   5055
                  Visible         =   0   'False
                  Width           =   645
               End
               Begin VB.Label lbl 
                  Alignment       =   2  'Center
                  BackStyle       =   0  'Transparent
                  Caption         =   "Age"
                  Height          =   465
                  Index           =   93
                  Left            =   8385
                  TabIndex        =   331
                  Top             =   6480
                  Visible         =   0   'False
                  Width           =   1425
               End
               Begin VB.Label lbl 
                  Alignment       =   2  'Center
                  BackStyle       =   0  'Transparent
                  Caption         =   "Gender"
                  Height          =   435
                  Index           =   92
                  Left            =   8385
                  TabIndex        =   330
                  Top             =   6135
                  Visible         =   0   'False
                  Width           =   1425
               End
               Begin VB.Label lbl«”„«·ÊÕœ… 
                  Alignment       =   2  'Center
                  AutoSize        =   -1  'True
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "Material "
                  Height          =   210
                  Index           =   15
                  Left            =   3870
                  RightToLeft     =   -1  'True
                  TabIndex        =   329
                  Top             =   3705
                  Width           =   525
               End
               Begin VB.Label lbl«”„«·ÊÕœ… 
                  Alignment       =   2  'Center
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "Shapes"
                  Height          =   330
                  Index           =   14
                  Left            =   7875
                  TabIndex        =   328
                  Top             =   180
                  Visible         =   0   'False
                  Width           =   1425
               End
               Begin VB.Label lbl«”„«·ÊÕœ… 
                  Alignment       =   2  'Center
                  AutoSize        =   -1  'True
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "Brand Type"
                  Height          =   210
                  Index           =   13
                  Left            =   3615
                  RightToLeft     =   -1  'True
                  TabIndex        =   327
                  Top             =   1650
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
                  Top             =   3240
                  Width           =   525
               End
               Begin VB.Label lbl«”„«·ÊÕœ… 
                  Alignment       =   2  'Center
                  AutoSize        =   -1  'True
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "Type"
                  Height          =   210
                  Index           =   10
                  Left            =   4005
                  RightToLeft     =   -1  'True
                  TabIndex        =   325
                  Top             =   2820
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
                  Top             =   840
                  Width           =   390
               End
               Begin VB.Label lbl«”„«·ÊÕœ… 
                  Alignment       =   2  'Center
                  AutoSize        =   -1  'True
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "Category"
                  Height          =   210
                  Index           =   16
                  Left            =   3750
                  RightToLeft     =   -1  'True
                  TabIndex        =   323
                  Top             =   2415
                  Width           =   645
               End
               Begin VB.Label lbl«”„«·ÊÕœ… 
                  Alignment       =   2  'Center
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "Source"
                  Height          =   345
                  Index           =   17
                  Left            =   8010
                  TabIndex        =   322
                  Top             =   660
                  Visible         =   0   'False
                  Width           =   1665
               End
               Begin VB.Label lbl«”„«·ÊÕœ… 
                  Alignment       =   2  'Center
                  AutoSize        =   -1  'True
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "Colors"
                  Height          =   210
                  Index           =   18
                  Left            =   4005
                  RightToLeft     =   -1  'True
                  TabIndex        =   321
                  Top             =   5160
                  Width           =   390
               End
               Begin VB.Label lbl«”„«·ÊÕœ… 
                  Alignment       =   2  'Center
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "Brand Type"
                  Height          =   465
                  Index           =   19
                  Left            =   9300
                  TabIndex        =   320
                  Top             =   1530
                  Visible         =   0   'False
                  Width           =   900
               End
               Begin VB.Label lbl«”„«·ÊÕœ… 
                  Alignment       =   2  'Center
                  AutoSize        =   -1  'True
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "Model"
                  Height          =   210
                  Index           =   20
                  Left            =   4005
                  RightToLeft     =   -1  'True
                  TabIndex        =   319
                  Top             =   1230
                  Width           =   390
               End
               Begin VB.Label lbl«”„«·ÊÕœ… 
                  Alignment       =   2  'Center
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "Color Code"
                  Height          =   195
                  Index           =   21
                  Left            =   1935
                  TabIndex        =   318
                  Top             =   1245
                  Width           =   525
               End
               Begin VB.Label lbl«”„«·ÊÕœ… 
                  Alignment       =   2  'Center
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "Size"
                  Height          =   255
                  Index           =   22
                  Left            =   645
                  TabIndex        =   317
                  Top             =   1215
                  Width           =   900
               End
               Begin VB.Label lbl«”„«·ÊÕœ… 
                  Alignment       =   2  'Center
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "Group"
                  Height          =   270
                  Index           =   24
                  Left            =   8520
                  TabIndex        =   316
                  Top             =   2325
                  Visible         =   0   'False
                  Width           =   1800
               End
               Begin VB.Label lbl«”„«·ÊÕœ… 
                  Alignment       =   2  'Center
                  AutoSize        =   -1  'True
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "Origin"
                  Height          =   210
                  Index           =   34
                  Left            =   4005
                  RightToLeft     =   -1  'True
                  TabIndex        =   315
                  Top             =   2085
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
                  Top             =   6600
                  Width           =   525
               End
               Begin VB.Label lbl«”„«·ÊÕœ… 
                  Alignment       =   2  'Center
                  AutoSize        =   -1  'True
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "Usage"
                  Height          =   210
                  Index           =   23
                  Left            =   3870
                  RightToLeft     =   -1  'True
                  TabIndex        =   313
                  Top             =   5550
                  Width           =   525
               End
               Begin VB.Label lbl«”„«·ÊÕœ… 
                  Alignment       =   2  'Center
                  AutoSize        =   -1  'True
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "Packing"
                  Height          =   210
                  Index           =   25
                  Left            =   3870
                  RightToLeft     =   -1  'True
                  TabIndex        =   312
                  Top             =   6195
                  Width           =   525
               End
            End
            Begin C1SizerLibCtl.C1Elastic ELe 
               Height          =   7275
               Index           =   27
               Left            =   45
               TabIndex        =   332
               TabStop         =   0   'False
               Top             =   45
               Width           =   11355
               _cx             =   20029
               _cy             =   12832
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
                  Height          =   7185
                  Index           =   8
                  Left            =   15240
                  TabIndex        =   333
                  Top             =   615
                  Width           =   11070
                  _cx             =   19526
                  _cy             =   12674
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
                  FormatString    =   $"FrmVizitScreen.frx":452D9
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
                  Height          =   7305
                  Index           =   28
                  Left            =   0
                  TabIndex        =   334
                  TabStop         =   0   'False
                  Top             =   0
                  Width           =   12255
                  _cx             =   21616
                  _cy             =   12885
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
                     Top             =   8790
                     Width           =   1500
                     _ExtentX        =   2646
                     _ExtentY        =   556
                     _Version        =   393216
                     Style           =   2
                     Text            =   ""
                     RightToLeft     =   -1  'True
                  End
                  Begin VSFlex8Ctl.VSFlexGrid Grid3 
                     Height          =   2400
                     Left            =   0
                     TabIndex        =   354
                     Top             =   0
                     Width           =   10920
                     _cx             =   19262
                     _cy             =   4233
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
                     FormatString    =   $"FrmVizitScreen.frx":45399
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
                     Height          =   1995
                     Left            =   0
                     TabIndex        =   355
                     Top             =   2430
                     Width           =   11055
                     _cx             =   19500
                     _cy             =   3519
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
                     FormatString    =   $"FrmVizitScreen.frx":45428
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
                     Height          =   2700
                     Left            =   0
                     TabIndex        =   356
                     Top             =   4920
                     Width           =   10920
                     _cx             =   19262
                     _cy             =   4762
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
                     FormatString    =   $"FrmVizitScreen.frx":4549A
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
                     Height          =   540
                     Index           =   26
                     Left            =   1500
                     TabIndex        =   336
                     Top             =   8940
                     Width           =   735
                  End
               End
            End
            Begin C1SizerLibCtl.C1Elastic ELe 
               Height          =   7275
               Index           =   9
               Left            =   12090
               TabIndex        =   337
               TabStop         =   0   'False
               Top             =   45
               Width           =   11355
               _cx             =   20029
               _cy             =   12832
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
                  Top             =   7320
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
                  Top             =   9270
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
                  Height          =   540
                  Index           =   1
                  Left            =   1425
                  TabIndex        =   341
                  Top             =   9420
                  Width           =   765
               End
               Begin VB.Label lbl«”„«·ÊÕœ… 
                  Alignment       =   2  'Center
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "Breaking"
                  Height          =   420
                  Index           =   0
                  Left            =   1545
                  TabIndex        =   340
                  Top             =   7515
                  Width           =   645
               End
            End
         End
         Begin VB.Label LabCurr_Rec 
            BackColor       =   &H00E2E9E9&
            Height          =   255
            Index           =   3
            Left            =   3630
            RightToLeft     =   -1  'True
            TabIndex        =   246
            Top             =   8160
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
            Top             =   8145
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
            Top             =   8145
            Width           =   1515
         End
         Begin VB.Label LabCount_Rec 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Height          =   255
            Index           =   3
            Left            =   0
            RightToLeft     =   -1  'True
            TabIndex        =   243
            Top             =   8040
            Width           =   1110
         End
      End
      Begin C1SizerLibCtl.C1Elastic ELe 
         Height          =   9210
         Index           =   6
         Left            =   -20235
         TabIndex        =   265
         TabStop         =   0   'False
         Top             =   45
         Width           =   18690
         _cx             =   32967
         _cy             =   16245
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
            Height          =   390
            Left            =   4755
            RightToLeft     =   -1  'True
            TabIndex        =   368
            Top             =   885
            Width           =   1245
         End
         Begin VB.Frame Frame9 
            Height          =   645
            Left            =   6270
            RightToLeft     =   -1  'True
            TabIndex        =   365
            Top             =   750
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
            Height          =   690
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
            Height          =   9090
            Index           =   4
            Left            =   25935
            TabIndex        =   268
            Top             =   795
            Width           =   18555
            _cx             =   32729
            _cy             =   16034
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
            FormatString    =   $"FrmVizitScreen.frx":45658
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
            Height          =   7635
            Left            =   270
            TabIndex        =   269
            Top             =   1755
            Width           =   18420
            _cx             =   32491
            _cy             =   13467
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
            FormatString    =   $"FrmVizitScreen.frx":45718
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
            Top             =   1065
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
            Height          =   300
            Index           =   5
            Left            =   15900
            RightToLeft     =   -1  'True
            TabIndex        =   281
            Top             =   1095
            Width           =   1680
         End
      End
      Begin C1SizerLibCtl.C1Elastic ELe 
         Height          =   9210
         Index           =   7
         Left            =   -19935
         TabIndex        =   270
         TabStop         =   0   'False
         Top             =   45
         Width           =   18690
         _cx             =   32967
         _cy             =   16245
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
            Height          =   9090
            Index           =   5
            Left            =   25935
            TabIndex        =   273
            Top             =   795
            Width           =   18555
            _cx             =   32729
            _cy             =   16034
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
            FormatString    =   $"FrmVizitScreen.frx":45A66
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
            Height          =   3255
            Left            =   135
            TabIndex        =   274
            Top             =   780
            Width           =   18285
            _cx             =   32253
            _cy             =   5741
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
            FormatString    =   $"FrmVizitScreen.frx":45B26
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
            Height          =   4230
            Left            =   270
            TabIndex        =   275
            Top             =   4680
            Width           =   18420
            _cx             =   32491
            _cy             =   7461
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
            FormatString    =   $"FrmVizitScreen.frx":45C5B
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
         Height          =   9210
         Index           =   8
         Left            =   -19635
         TabIndex        =   276
         TabStop         =   0   'False
         Top             =   45
         Width           =   18690
         _cx             =   32967
         _cy             =   16245
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
            Height          =   720
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
            Height          =   9090
            Index           =   6
            Left            =   25935
            TabIndex        =   279
            Top             =   795
            Width           =   18555
            _cx             =   32729
            _cy             =   16034
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
            FormatString    =   $"FrmVizitScreen.frx":45E35
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
            Height          =   7785
            Left            =   135
            TabIndex        =   361
            Top             =   1440
            Width           =   18420
            _cx             =   32491
            _cy             =   13732
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
            FormatString    =   $"FrmVizitScreen.frx":45EF5
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
            Top             =   975
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
            Height          =   300
            Index           =   7
            Left            =   16590
            RightToLeft     =   -1  'True
            TabIndex        =   363
            Top             =   1005
            Width           =   1680
         End
      End
      Begin C1SizerLibCtl.C1Elastic ELe 
         Height          =   9210
         Index           =   10
         Left            =   -19335
         TabIndex        =   370
         TabStop         =   0   'False
         Top             =   45
         Width           =   18690
         _cx             =   32967
         _cy             =   16245
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
            Height          =   615
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
                     Picture         =   "FrmVizitScreen.frx":46381
                     Key             =   "CompanyName"
                  EndProperty
                  BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                     Picture         =   "FrmVizitScreen.frx":4671B
                     Key             =   "Ser"
                  EndProperty
                  BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                     Picture         =   "FrmVizitScreen.frx":46AB5
                     Key             =   "Vac_Name"
                  EndProperty
                  BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                     Picture         =   "FrmVizitScreen.frx":46E4F
                     Key             =   "ShareCount"
                  EndProperty
                  BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                     Picture         =   "FrmVizitScreen.frx":471E9
                     Key             =   "Dis_Count"
                  EndProperty
                  BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                     Picture         =   "FrmVizitScreen.frx":47583
                     Key             =   "Bouns"
                  EndProperty
                  BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                     Picture         =   "FrmVizitScreen.frx":4791D
                     Key             =   "SharesValue"
                  EndProperty
                  BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                     Picture         =   "FrmVizitScreen.frx":47EB7
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
               ButtonImage     =   "FrmVizitScreen.frx":48251
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
               ButtonImage     =   "FrmVizitScreen.frx":485EB
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
               ButtonImage     =   "FrmVizitScreen.frx":48985
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
               ButtonImage     =   "FrmVizitScreen.frx":48D1F
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
            Height          =   1410
            Index           =   1
            Left            =   270
            RightToLeft     =   -1  'True
            TabIndex        =   371
            Top             =   810
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
               ItemData        =   "FrmVizitScreen.frx":490B9
               Left            =   2280
               List            =   "FrmVizitScreen.frx":490C9
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
               Bindings        =   "FrmVizitScreen.frx":490E2
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
               Format          =   215744513
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
            Height          =   360
            Index           =   7
            Left            =   7245
            TabIndex        =   383
            Top             =   8115
            Width           =   1125
            _ExtentX        =   1984
            _ExtentY        =   635
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
            ButtonImage     =   "FrmVizitScreen.frx":490F7
            ColorButton     =   14871017
            DrawFocusRectangle=   0   'False
         End
         Begin ImpulseButton.ISButton btn_Save 
            Height          =   360
            Index           =   7
            Left            =   5310
            TabIndex        =   384
            Top             =   8085
            Width           =   825
            _ExtentX        =   1455
            _ExtentY        =   635
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
            ButtonImage     =   "FrmVizitScreen.frx":49491
            ColorButton     =   14871017
            DrawFocusRectangle=   0   'False
         End
         Begin ImpulseButton.ISButton btn_Modify 
            Height          =   360
            Index           =   7
            Left            =   6270
            TabIndex        =   385
            Top             =   8145
            Width           =   975
            _ExtentX        =   1720
            _ExtentY        =   635
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
            ButtonImage     =   "FrmVizitScreen.frx":4982B
            ColorButton     =   14871017
            DrawFocusRectangle=   0   'False
         End
         Begin ImpulseButton.ISButton Btn_Undo 
            Height          =   375
            Index           =   7
            Left            =   4320
            TabIndex        =   386
            Top             =   8040
            Width           =   990
            _ExtentX        =   1746
            _ExtentY        =   661
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
            ButtonImage     =   "FrmVizitScreen.frx":49BC5
            ColorButton     =   14871017
            DrawFocusRectangle=   0   'False
         End
         Begin ImpulseButton.ISButton btn_Delete 
            Height          =   360
            Index           =   7
            Left            =   3480
            TabIndex        =   387
            Top             =   8085
            Width           =   840
            _ExtentX        =   1482
            _ExtentY        =   635
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
            ButtonImage     =   "FrmVizitScreen.frx":49F5F
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
            Top             =   6765
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
            ButtonImage     =   "FrmVizitScreen.frx":4A4F9
            ColorButton     =   14871017
            DrawFocusRectangle=   0   'False
         End
         Begin ImpulseButton.ISButton btn_Cancel 
            Height          =   360
            Index           =   7
            Left            =   135
            TabIndex        =   389
            Top             =   8010
            Width           =   975
            _ExtentX        =   1720
            _ExtentY        =   635
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
            ButtonImage     =   "FrmVizitScreen.frx":4A893
            ColorButton     =   14871017
            DrawFocusRectangle=   0   'False
            DisabledImageStyle=   1
         End
         Begin ImpulseButton.ISButton Btn_Print 
            Height          =   465
            Index           =   7
            Left            =   2370
            TabIndex        =   390
            TabStop         =   0   'False
            ToolTipText     =   "ÿ»«⁄… «·»Ì«‰«  "
            Top             =   8010
            Width           =   975
            _ExtentX        =   1720
            _ExtentY        =   820
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
            ButtonImage     =   "FrmVizitScreen.frx":4AC2D
            ColorButton     =   14871017
            DrawFocusRectangle=   0   'False
            DisabledImageStyle=   1
         End
         Begin ImpulseButton.ISButton btn_Query 
            Height          =   570
            Index           =   7
            Left            =   1530
            TabIndex        =   391
            TabStop         =   0   'False
            ToolTipText     =   "(Ctrl+F)  ··»ÕÀ ≈÷€ÿ Â–« «·„› «Õ √Ê ≈÷€ÿ "
            Top             =   7935
            Width           =   705
            _ExtentX        =   1244
            _ExtentY        =   1005
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
            ButtonImage     =   "FrmVizitScreen.frx":5148F
            ColorButton     =   14871017
            DrawFocusRectangle=   0   'False
         End
         Begin VSFlex8Ctl.VSFlexGrid GrdExcel 
            Height          =   4170
            Left            =   420
            TabIndex        =   392
            Top             =   2490
            Width           =   18540
            _cx             =   32702
            _cy             =   7355
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
            FormatString    =   $"FrmVizitScreen.frx":51829
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
            Top             =   7035
            Width           =   3360
            _ExtentX        =   5927
            _ExtentY        =   556
            _Version        =   393216
            Enabled         =   0   'False
            Text            =   ""
            RightToLeft     =   -1  'True
         End
         Begin ImpulseButton.ISButton Cmd_DeleteRow 
            Height          =   315
            Index           =   7
            Left            =   2685
            TabIndex        =   419
            Top             =   6780
            Width           =   1800
            _ExtentX        =   3175
            _ExtentY        =   556
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
            ButtonImage     =   "FrmVizitScreen.frx":51A10
            ColorButton     =   14871017
            DrawFocusRectangle=   0   'False
         End
         Begin ImpulseButton.ISButton Cmd_DeleteAll 
            Height          =   315
            Index           =   7
            Left            =   450
            TabIndex        =   420
            Top             =   6765
            Width           =   2100
            _ExtentX        =   3704
            _ExtentY        =   556
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
            ButtonImage     =   "FrmVizitScreen.frx":51FAA
            ColorButton     =   14871017
            DrawFocusRectangle=   0   'False
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Õ—— »Ê«”ÿ…  "
            Height          =   330
            Index           =   30
            Left            =   17130
            TabIndex        =   416
            Top             =   6975
            Width           =   1245
         End
         Begin VB.Label LabCount_Rec 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Height          =   255
            Index           =   7
            Left            =   1245
            RightToLeft     =   -1  'True
            TabIndex        =   397
            Top             =   7530
            Width           =   1275
         End
         Begin VB.Label LabCurr_Rec 
            BackColor       =   &H00E2E9E9&
            Height          =   255
            Index           =   7
            Left            =   4890
            RightToLeft     =   -1  'True
            TabIndex        =   396
            Top             =   7530
            Width           =   1530
         End
         Begin VB.Label Label2 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "⁄œœ «·”Ã·« :"
            Height          =   240
            Index           =   13
            Left            =   2850
            RightToLeft     =   -1  'True
            TabIndex        =   395
            Top             =   7545
            Width           =   2085
         End
         Begin VB.Label Label2 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "«·”Ã· «·Õ«·Ì:"
            Height          =   240
            Index           =   12
            Left            =   6690
            RightToLeft     =   -1  'True
            TabIndex        =   394
            Top             =   7515
            Width           =   1545
         End
         Begin VB.Label LabCurr_Rec 
            BackColor       =   &H00E2E9E9&
            Height          =   255
            Index           =   4
            Left            =   4320
            RightToLeft     =   -1  'True
            TabIndex        =   393
            Top             =   7035
            Width           =   1545
         End
      End
      Begin C1SizerLibCtl.C1Elastic ELe 
         Height          =   9210
         Index           =   11
         Left            =   45
         TabIndex        =   422
         TabStop         =   0   'False
         Top             =   45
         Width           =   18690
         _cx             =   32967
         _cy             =   16245
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
         Begin VB.CommandButton Command7 
            Caption         =   "«⁄„«— «·«’‰«›"
            Height          =   330
            Left            =   150
            RightToLeft     =   -1  'True
            TabIndex        =   565
            Top             =   1620
            Width           =   4320
         End
         Begin VB.OptionButton Option4 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "„” √Ã—Ì‰"
            ForeColor       =   &H00C00000&
            Height          =   255
            Left            =   5160
            RightToLeft     =   -1  'True
            TabIndex        =   468
            Top             =   1230
            Width           =   1215
         End
         Begin VB.OptionButton Option3 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "„·«ﬂ"
            ForeColor       =   &H00800000&
            Height          =   255
            Left            =   6840
            RightToLeft     =   -1  'True
            TabIndex        =   467
            Top             =   1230
            Width           =   735
         End
         Begin VB.TextBox txtTotalStill 
            Alignment       =   1  'Right Justify
            Height          =   465
            Left            =   7470
            RightToLeft     =   -1  'True
            TabIndex        =   460
            Top             =   8130
            Visible         =   0   'False
            Width           =   2805
         End
         Begin VB.CommandButton CmdSelectCus 
            Caption         =   " ÕœÌœ>>"
            Height          =   330
            Left            =   5130
            RightToLeft     =   -1  'True
            TabIndex        =   459
            Top             =   4020
            Width           =   4320
         End
         Begin VB.CommandButton CmdSelectEmp 
            Caption         =   " ÕœÌœ>>"
            Height          =   330
            Left            =   5160
            RightToLeft     =   -1  'True
            TabIndex        =   458
            Top             =   4830
            Width           =   4320
         End
         Begin VB.OptionButton Option1 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "⁄„·«¡"
            ForeColor       =   &H00800000&
            Height          =   255
            Index           =   2
            Left            =   9180
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
            Left            =   7575
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
            Left            =   7200
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
            TabIndex        =   436
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
            Format          =   192282627
            CurrentDate     =   37140
         End
         Begin ImpulseButton.ISButton CmdPrint 
            Height          =   330
            Left            =   4080
            TabIndex        =   437
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
            ButtonImage     =   "FrmVizitScreen.frx":52544
            ColorButton     =   14871017
            DrawFocusRectangle=   0   'False
         End
         Begin MSDataListLib.DataCombo dcBranch 
            Height          =   315
            Index           =   8
            Left            =   5100
            TabIndex        =   438
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
            TabIndex        =   439
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
            TabIndex        =   440
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
            TabIndex        =   441
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
            ButtonImage     =   "FrmVizitScreen.frx":528DE
            ColorButton     =   14871017
            DrawFocusRectangle=   0   'False
         End
         Begin ImpulseButton.ISButton BtnPrint 
            Height          =   555
            Index           =   1
            Left            =   2160
            TabIndex        =   442
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
            ButtonImage     =   "FrmVizitScreen.frx":52C78
            ColorButton     =   14871017
            DrawFocusRectangle=   0   'False
         End
         Begin XtremeSuiteControls.CheckBox ChekCustomer 
            Height          =   375
            Left            =   13290
            TabIndex        =   443
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
            TabIndex        =   444
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
            TabIndex        =   445
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
         Begin MSComCtl2.DTPicker FromDate1 
            Height          =   345
            Left            =   2925
            TabIndex        =   446
            TabStop         =   0   'False
            Top             =   8220
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
            Format          =   192282627
            CurrentDate     =   37140
         End
         Begin MSComCtl2.DTPicker DTPickerAccFrom 
            Height          =   345
            Left            =   3120
            TabIndex        =   447
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
            Format          =   192282627
            CurrentDate     =   37140
         End
         Begin VSFlex8Ctl.VSFlexGrid grdAging 
            Height          =   3630
            Left            =   1860
            TabIndex        =   435
            Top             =   7440
            Visible         =   0   'False
            Width           =   15420
            _cx             =   27199
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
            Cols            =   26
            FixedRows       =   1
            FixedCols       =   1
            RowHeightMin    =   320
            RowHeightMax    =   0
            ColWidthMin     =   0
            ColWidthMax     =   0
            ExtendLastCol   =   -1  'True
            FormatString    =   $"FrmVizitScreen.frx":53012
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
            Left            =   90
            TabIndex        =   457
            Top             =   6780
            Visible         =   0   'False
            Width           =   17370
            _cx             =   30639
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
            Cols            =   26
            FixedRows       =   1
            FixedCols       =   1
            RowHeightMin    =   320
            RowHeightMax    =   0
            ColWidthMin     =   0
            ColWidthMax     =   0
            ExtendLastCol   =   -1  'True
            FormatString    =   $"FrmVizitScreen.frx":533FF
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
            TabIndex        =   462
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
            ButtonImage     =   "FrmVizitScreen.frx":537F5
            ColorButton     =   14871017
            DrawFocusRectangle=   0   'False
            DisabledImageStyle=   1
         End
         Begin MSDataListLib.DataCombo dcClass 
            Height          =   315
            Left            =   5100
            TabIndex        =   463
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
            TabIndex        =   465
            Tag             =   "⁄›Ê« Ì—ÃÏ «œŒ«· √”„ «·ÕÌ"
            Top             =   1620
            Width           =   7545
            _ExtentX        =   13309
            _ExtentY        =   556
            _Version        =   393216
            Text            =   ""
            RightToLeft     =   -1  'True
         End
         Begin ImpulseButton.ISButton BtnPrint 
            Height          =   555
            Index           =   2
            Left            =   5580
            TabIndex        =   578
            Top             =   6000
            Width           =   1560
            _ExtentX        =   2752
            _ExtentY        =   979
            ButtonPositionImage=   1
            Caption         =   "ÿ»«⁄…  Õ·Ì·Ì 2"
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
            ButtonImage     =   "FrmVizitScreen.frx":53B8F
            ColorButton     =   14871017
            DrawFocusRectangle=   0   'False
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "‰Ê⁄ «·⁄„Ì·"
            Height          =   285
            Index           =   21
            Left            =   12615
            RightToLeft     =   -1  'True
            TabIndex        =   466
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
            TabIndex        =   464
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
            TabIndex        =   461
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
            TabIndex        =   456
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
            TabIndex        =   455
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
            TabIndex        =   454
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
            TabIndex        =   453
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
            TabIndex        =   452
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
            TabIndex        =   451
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
            TabIndex        =   450
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
            TabIndex        =   449
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
            TabIndex        =   448
            Top             =   8220
            Visible         =   0   'False
            Width           =   390
         End
      End
      Begin C1SizerLibCtl.C1Elastic ELe 
         Height          =   9210
         Index           =   12
         Left            =   19425
         TabIndex        =   469
         TabStop         =   0   'False
         Top             =   45
         Width           =   18690
         _cx             =   32967
         _cy             =   16245
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
         Begin VB.TextBox Text16 
            Alignment       =   1  'Right Justify
            Height          =   360
            Left            =   11520
            Locked          =   -1  'True
            TabIndex        =   508
            Top             =   7920
            Width           =   5985
         End
         Begin VB.TextBox Text15 
            Alignment       =   1  'Right Justify
            Height          =   360
            Left            =   11520
            Locked          =   -1  'True
            TabIndex        =   506
            Top             =   7425
            Width           =   5985
         End
         Begin VB.TextBox Text14 
            Alignment       =   1  'Right Justify
            Height          =   360
            Left            =   2880
            Locked          =   -1  'True
            TabIndex        =   504
            Top             =   7920
            Width           =   5985
         End
         Begin VB.TextBox Text13 
            Alignment       =   1  'Right Justify
            Height          =   360
            Left            =   2880
            Locked          =   -1  'True
            TabIndex        =   502
            Top             =   7425
            Width           =   5985
         End
         Begin VB.TextBox Text12 
            Alignment       =   1  'Right Justify
            Height          =   360
            Left            =   2880
            Locked          =   -1  'True
            TabIndex        =   500
            Top             =   6915
            Width           =   5985
         End
         Begin VB.Frame Frame1 
            BackColor       =   &H00E2E9E9&
            BorderStyle     =   0  'None
            Height          =   2415
            Index           =   4
            Left            =   240
            RightToLeft     =   -1  'True
            TabIndex        =   470
            Top             =   810
            Width           =   18150
            Begin VB.TextBox Text9 
               Alignment       =   1  'Right Justify
               Height          =   345
               Left            =   1440
               Locked          =   -1  'True
               TabIndex        =   491
               Top             =   1320
               Width           =   4185
            End
            Begin VB.TextBox Text6 
               Alignment       =   1  'Right Justify
               Height          =   345
               Left            =   1440
               Locked          =   -1  'True
               TabIndex        =   489
               Top             =   840
               Width           =   4185
            End
            Begin VB.ComboBox Combo2 
               BackColor       =   &H80000018&
               Height          =   315
               ItemData        =   "FrmVizitScreen.frx":53F29
               Left            =   2280
               List            =   "FrmVizitScreen.frx":53F39
               RightToLeft     =   -1  'True
               Style           =   2  'Dropdown List
               TabIndex        =   471
               Top             =   3150
               Visible         =   0   'False
               Width           =   1005
            End
            Begin MSDataListLib.DataCombo dcBranch 
               Bindings        =   "FrmVizitScreen.frx":53F52
               Height          =   315
               Index           =   0
               Left            =   2880
               TabIndex        =   472
               Top             =   0
               Visible         =   0   'False
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
            Begin MSDataListLib.DataCombo DcCustmer 
               Height          =   315
               Index           =   9
               Left            =   1440
               TabIndex        =   488
               Top             =   480
               Width           =   4185
               _ExtentX        =   7382
               _ExtentY        =   556
               _Version        =   393216
               IntegralHeight  =   0   'False
               Text            =   ""
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
            Begin XtremeSuiteControls.CheckBox CheckBox1 
               Height          =   375
               Left            =   1560
               TabIndex        =   494
               Top             =   1920
               Width           =   1155
               _Version        =   786432
               _ExtentX        =   2037
               _ExtentY        =   661
               _StockProps     =   79
               Caption         =   "Sell Out "
               BackColor       =   14871017
               UseVisualStyle  =   -1  'True
               TextAlignment   =   1
               RightToLeft     =   -1  'True
            End
            Begin XtremeSuiteControls.CheckBox CheckBox2 
               Height          =   375
               Left            =   2760
               TabIndex        =   495
               Top             =   1920
               Width           =   1155
               _Version        =   786432
               _ExtentX        =   2037
               _ExtentY        =   661
               _StockProps     =   79
               Caption         =   "Sell in"
               BackColor       =   14871017
               UseVisualStyle  =   -1  'True
               TextAlignment   =   1
               RightToLeft     =   -1  'True
            End
            Begin VB.Label Label1 
               BackColor       =   &H00E2E9E9&
               BackStyle       =   0  'Transparent
               Caption         =   "Promotion end day :   "
               Height          =   285
               Index           =   31
               Left            =   8160
               TabIndex        =   499
               Top             =   1560
               Width           =   1695
            End
            Begin VB.Label Label1 
               BackColor       =   &H00E2E9E9&
               BackStyle       =   0  'Transparent
               Caption         =   "Promotion Start day :   "
               Height          =   285
               Index           =   30
               Left            =   8160
               TabIndex        =   498
               Top             =   1200
               Width           =   1695
            End
            Begin VB.Label Label1 
               BackColor       =   &H00E2E9E9&
               BackStyle       =   0  'Transparent
               Caption         =   "Submission Date :  "
               Height          =   285
               Index           =   29
               Left            =   8160
               TabIndex        =   497
               Top             =   840
               Width           =   1695
            End
            Begin VB.Label Label1 
               BackColor       =   &H00E2E9E9&
               BackStyle       =   0  'Transparent
               Caption         =   "Promotion Week :"
               Height          =   285
               Index           =   28
               Left            =   8160
               TabIndex        =   496
               Top             =   480
               Width           =   1695
            End
            Begin VB.Label Label1 
               BackColor       =   &H00E2E9E9&
               BackStyle       =   0  'Transparent
               Caption         =   "Contract Type :   "
               Height          =   405
               Index           =   27
               Left            =   120
               TabIndex        =   493
               Top             =   1920
               Width           =   1095
            End
            Begin VB.Label Label1 
               BackColor       =   &H00E2E9E9&
               BackStyle       =   0  'Transparent
               Caption         =   "Contract No#"
               Height          =   405
               Index           =   26
               Left            =   120
               TabIndex        =   492
               Top             =   1440
               Width           =   1095
            End
            Begin VB.Label Label1 
               BackColor       =   &H00E2E9E9&
               BackStyle       =   0  'Transparent
               Caption         =   "Vendor Code "
               Height          =   405
               Index           =   25
               Left            =   120
               TabIndex        =   490
               Top             =   960
               Width           =   1095
            End
            Begin VB.Label Label6 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "Serial"
               Height          =   195
               Index           =   5
               Left            =   13980
               RightToLeft     =   -1  'True
               TabIndex        =   476
               Top             =   30
               Visible         =   0   'False
               Width           =   990
            End
            Begin VB.Label Label1 
               BackColor       =   &H00E2E9E9&
               BackStyle       =   0  'Transparent
               Caption         =   "Customer"
               Height          =   405
               Index           =   24
               Left            =   105
               TabIndex        =   475
               Top             =   570
               Width           =   1095
            End
            Begin VB.Label lbl 
               Alignment       =   2  'Center
               BackColor       =   &H00E2E9E9&
               Caption         =   "Date"
               Height          =   270
               Index           =   35
               Left            =   8220
               TabIndex        =   474
               Top             =   30
               Width           =   1215
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "«·›—⁄"
               Height          =   255
               Index           =   34
               Left            =   5790
               TabIndex        =   473
               Top             =   0
               Visible         =   0   'False
               Width           =   600
            End
         End
         Begin ImpulseButton.ISButton Btn_Update 
            Height          =   240
            Index           =   0
            Left            =   6555
            TabIndex        =   477
            TabStop         =   0   'False
            ToolTipText     =   " ÕœÌÀ ﬁ«⁄œ… «·»Ì«‰« "
            Top             =   5880
            Visible         =   0   'False
            Width           =   1815
            _ExtentX        =   3201
            _ExtentY        =   423
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
            ButtonImage     =   "FrmVizitScreen.frx":53F67
            ColorButton     =   14871017
            DrawFocusRectangle=   0   'False
         End
         Begin VSFlex8Ctl.VSFlexGrid VSFlexGrid1 
            Height          =   2550
            Left            =   240
            TabIndex        =   478
            Top             =   3240
            Width           =   18540
            _cx             =   32702
            _cy             =   4498
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
            FormatString    =   $"FrmVizitScreen.frx":54301
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
            WallPaperAlignment=   9
            AccessibleName  =   ""
            AccessibleDescription=   ""
            AccessibleValue =   ""
            AccessibleRole  =   24
         End
         Begin MSDataListLib.DataCombo DCboUserName 
            Height          =   315
            Index           =   0
            Left            =   13620
            TabIndex        =   479
            Top             =   7035
            Visible         =   0   'False
            Width           =   3360
            _ExtentX        =   5927
            _ExtentY        =   556
            _Version        =   393216
            Enabled         =   0   'False
            Text            =   ""
            RightToLeft     =   -1  'True
         End
         Begin ImpulseButton.ISButton Cmd_DeleteRow 
            Height          =   315
            Index           =   0
            Left            =   2685
            TabIndex        =   480
            Top             =   5895
            Width           =   1800
            _ExtentX        =   3175
            _ExtentY        =   556
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
            ButtonImage     =   "FrmVizitScreen.frx":54507
            ColorButton     =   14871017
            DrawFocusRectangle=   0   'False
         End
         Begin ImpulseButton.ISButton Cmd_DeleteAll 
            Height          =   315
            Index           =   0
            Left            =   450
            TabIndex        =   481
            Top             =   5880
            Width           =   2100
            _ExtentX        =   3704
            _ExtentY        =   556
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
            ButtonImage     =   "FrmVizitScreen.frx":54AA1
            ColorButton     =   14871017
            DrawFocusRectangle=   0   'False
         End
         Begin VB.Label Label1 
            BackColor       =   &H00E2E9E9&
            BackStyle       =   0  'Transparent
            Caption         =   "4) Final Approval:"
            Height          =   180
            Index           =   36
            Left            =   9120
            TabIndex        =   509
            Top             =   7920
            Width           =   2655
         End
         Begin VB.Label Label1 
            BackColor       =   &H00E2E9E9&
            BackStyle       =   0  'Transparent
            Caption         =   "2) Direct Head / Sales Manager"
            Height          =   165
            Index           =   35
            Left            =   9120
            TabIndex        =   507
            Top             =   7425
            Width           =   2655
         End
         Begin VB.Label Label1 
            BackColor       =   &H00E2E9E9&
            BackStyle       =   0  'Transparent
            Caption         =   "3) Finance Department:"
            Height          =   180
            Index           =   34
            Left            =   240
            TabIndex        =   505
            Top             =   7920
            Width           =   2655
         End
         Begin VB.Label Label1 
            BackColor       =   &H00E2E9E9&
            BackStyle       =   0  'Transparent
            Caption         =   "1) Employee Signature:"
            Height          =   165
            Index           =   33
            Left            =   240
            TabIndex        =   503
            Top             =   7425
            Width           =   2655
         End
         Begin VB.Label Label1 
            BackColor       =   &H00E2E9E9&
            BackStyle       =   0  'Transparent
            Caption         =   "Justification: (if Any)"
            Height          =   420
            Index           =   32
            Left            =   240
            TabIndex        =   501
            Top             =   6915
            Width           =   2655
         End
         Begin VB.Label LabCurr_Rec 
            BackColor       =   &H00E2E9E9&
            Height          =   255
            Index           =   6
            Left            =   4320
            RightToLeft     =   -1  'True
            TabIndex        =   487
            Top             =   7035
            Width           =   1545
         End
         Begin VB.Label Label2 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "«·”Ã· «·Õ«·Ì:"
            Height          =   240
            Index           =   9
            Left            =   6690
            RightToLeft     =   -1  'True
            TabIndex        =   486
            Top             =   6255
            Width           =   1545
         End
         Begin VB.Label Label2 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "⁄œœ «·”Ã·« :"
            Height          =   240
            Index           =   8
            Left            =   2850
            RightToLeft     =   -1  'True
            TabIndex        =   485
            Top             =   6285
            Width           =   2085
         End
         Begin VB.Label LabCurr_Rec 
            BackColor       =   &H00E2E9E9&
            Height          =   255
            Index           =   5
            Left            =   4890
            RightToLeft     =   -1  'True
            TabIndex        =   484
            Top             =   6270
            Width           =   1530
         End
         Begin VB.Label LabCount_Rec 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Height          =   255
            Index           =   4
            Left            =   1245
            RightToLeft     =   -1  'True
            TabIndex        =   483
            Top             =   6270
            Width           =   1275
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Õ—— »Ê«”ÿ…  "
            Height          =   330
            Index           =   36
            Left            =   17130
            TabIndex        =   482
            Top             =   6975
            Visible         =   0   'False
            Width           =   1245
         End
      End
      Begin C1SizerLibCtl.C1Elastic ELe 
         Height          =   9210
         Index           =   13
         Left            =   19725
         TabIndex        =   510
         TabStop         =   0   'False
         Top             =   45
         Width           =   18690
         _cx             =   32967
         _cy             =   16245
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
            Index           =   10
            Left            =   630
            MaxLength       =   50
            RightToLeft     =   -1  'True
            TabIndex        =   559
            Top             =   660
            Width           =   1065
         End
         Begin VB.TextBox txtNetSalesAfter 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   4
            Left            =   2430
            RightToLeft     =   -1  'True
            TabIndex        =   557
            Top             =   7500
            Width           =   1545
         End
         Begin VB.CommandButton cmdInsert 
            Caption         =   "⁄—÷"
            Height          =   390
            Left            =   13470
            RightToLeft     =   -1  'True
            TabIndex        =   555
            Top             =   600
            Width           =   1245
         End
         Begin VB.Frame Fra_Header 
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   0  'None
            Height          =   615
            Index           =   4
            Left            =   30
            RightToLeft     =   -1  'True
            TabIndex        =   544
            Top             =   -60
            Width           =   18690
            Begin VB.TextBox TxtModFlg2 
               Alignment       =   1  'Right Justify
               Appearance      =   0  'Flat
               BackColor       =   &H0000FF00&
               Enabled         =   0   'False
               Height          =   285
               Index           =   10
               Left            =   2580
               RightToLeft     =   -1  'True
               TabIndex        =   545
               Text            =   "modflag"
               Top             =   90
               Visible         =   0   'False
               Width           =   465
            End
            Begin MSComctlLib.ImageList GrdImageList2 
               Index           =   1
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
                     Picture         =   "FrmVizitScreen.frx":5503B
                     Key             =   "CompanyName"
                  EndProperty
                  BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                     Picture         =   "FrmVizitScreen.frx":553D5
                     Key             =   "Ser"
                  EndProperty
                  BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                     Picture         =   "FrmVizitScreen.frx":5576F
                     Key             =   "Vac_Name"
                  EndProperty
                  BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                     Picture         =   "FrmVizitScreen.frx":55B09
                     Key             =   "ShareCount"
                  EndProperty
                  BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                     Picture         =   "FrmVizitScreen.frx":55EA3
                     Key             =   "Dis_Count"
                  EndProperty
                  BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                     Picture         =   "FrmVizitScreen.frx":5623D
                     Key             =   "Bouns"
                  EndProperty
                  BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                     Picture         =   "FrmVizitScreen.frx":565D7
                     Key             =   "SharesValue"
                  EndProperty
                  BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                     Picture         =   "FrmVizitScreen.frx":56B71
                     Key             =   "BuyValue"
                  EndProperty
               EndProperty
            End
            Begin ImpulseButton.ISButton btn_Last 
               Height          =   315
               Index           =   10
               Left            =   90
               TabIndex        =   546
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
               ButtonImage     =   "FrmVizitScreen.frx":56F0B
               ColorButton     =   14871017
               AcclimateGrayTones=   -1  'True
               DrawFocusRectangle=   0   'False
               DisabledImageExtraction=   0
               DisabledImageStyle=   1
            End
            Begin ImpulseButton.ISButton btn_Next 
               Height          =   315
               Index           =   10
               Left            =   555
               TabIndex        =   547
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
               ButtonImage     =   "FrmVizitScreen.frx":572A5
               ColorButton     =   14871017
               DrawFocusRectangle=   0   'False
               DisabledImageStyle=   1
            End
            Begin ImpulseButton.ISButton btn_Previous 
               Height          =   315
               Index           =   10
               Left            =   1155
               TabIndex        =   548
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
               ButtonImage     =   "FrmVizitScreen.frx":5763F
               ColorButton     =   14871017
               DrawFocusRectangle=   0   'False
               DisabledImageStyle=   1
            End
            Begin ImpulseButton.ISButton btn_First 
               Height          =   315
               Index           =   10
               Left            =   1620
               TabIndex        =   549
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
               ButtonImage     =   "FrmVizitScreen.frx":579D9
               ColorButton     =   14871017
               DrawFocusRectangle=   0   'False
               DisabledImageStyle=   1
            End
            Begin MSComDlg.CommonDialog CommonDialog1 
               Left            =   0
               Top             =   120
               _ExtentX        =   847
               _ExtentY        =   847
               _Version        =   393216
            End
         End
         Begin VB.TextBox txtNetSalesAfter 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   3
            Left            =   10890
            RightToLeft     =   -1  'True
            TabIndex        =   534
            Top             =   5370
            Width           =   1545
         End
         Begin VB.TextBox txtNetSalesAfter 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   2
            Left            =   9330
            RightToLeft     =   -1  'True
            TabIndex        =   533
            Top             =   5370
            Width           =   1545
         End
         Begin VB.TextBox txtNetSalesAfter 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   1
            Left            =   7770
            RightToLeft     =   -1  'True
            TabIndex        =   532
            Top             =   5370
            Width           =   1545
         End
         Begin VB.TextBox txtNetSalesAfter 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   0
            Left            =   7560
            RightToLeft     =   -1  'True
            TabIndex        =   530
            Top             =   3360
            Width           =   3285
         End
         Begin VB.Frame Frame11 
            Caption         =   "”œ«œ «·„»·€ «·„” Õﬁ"
            Height          =   1395
            Index           =   4
            Left            =   5940
            RightToLeft     =   -1  'True
            TabIndex        =   526
            Top             =   7410
            Width           =   5775
            Begin VB.TextBox XPTxtID 
               Alignment       =   1  'Right Justify
               Height          =   375
               Index           =   4
               Left            =   0
               RightToLeft     =   -1  'True
               TabIndex        =   564
               Top             =   0
               Width           =   645
            End
            Begin VB.CommandButton cmdcreate 
               Caption         =   "”œ«œ"
               Height          =   1095
               Index           =   4
               Left            =   30
               RightToLeft     =   -1  'True
               TabIndex        =   527
               Top             =   240
               Width           =   615
            End
            Begin VSFlex8Ctl.VSFlexGrid grdAcc 
               Height          =   1080
               Index           =   4
               Left            =   660
               TabIndex        =   528
               Top             =   240
               Width           =   5130
               _cx             =   9049
               _cy             =   1905
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
               BackColor       =   12615935
               ForeColor       =   -2147483640
               BackColorFixed  =   14871017
               ForeColorFixed  =   -2147483630
               BackColorSel    =   -2147483635
               ForeColorSel    =   -2147483634
               BackColorBkg    =   -2147483636
               BackColorAlternate=   12615935
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
               Rows            =   3
               Cols            =   5
               FixedRows       =   1
               FixedCols       =   1
               RowHeightMin    =   320
               RowHeightMax    =   0
               ColWidthMin     =   0
               ColWidthMax     =   0
               ExtendLastCol   =   -1  'True
               FormatString    =   $"FrmVizitScreen.frx":57D73
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
         Begin VB.Frame Frame11 
            Caption         =   "Œ’„ „⁄Ã·"
            Height          =   1395
            Index           =   3
            Left            =   11700
            RightToLeft     =   -1  'True
            TabIndex        =   523
            Top             =   7380
            Width           =   5865
            Begin VB.TextBox XPTxtID 
               Alignment       =   1  'Right Justify
               Height          =   375
               Index           =   3
               Left            =   0
               RightToLeft     =   -1  'True
               TabIndex        =   563
               Top             =   0
               Width           =   645
            End
            Begin VB.CommandButton cmdcreate 
               Caption         =   "«‰‘«¡ «‘⁄«— ¬·Ì"
               Height          =   1095
               Index           =   3
               Left            =   30
               RightToLeft     =   -1  'True
               TabIndex        =   524
               Top             =   240
               Width           =   615
            End
            Begin VSFlex8Ctl.VSFlexGrid grdAcc 
               Height          =   1080
               Index           =   3
               Left            =   660
               TabIndex        =   525
               Top             =   240
               Width           =   5130
               _cx             =   9049
               _cy             =   1905
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
               BackColor       =   16744576
               ForeColor       =   -2147483640
               BackColorFixed  =   14871017
               ForeColorFixed  =   -2147483630
               BackColorSel    =   -2147483635
               ForeColorSel    =   -2147483634
               BackColorBkg    =   -2147483636
               BackColorAlternate=   16744576
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
               Rows            =   3
               Cols            =   5
               FixedRows       =   1
               FixedCols       =   1
               RowHeightMin    =   320
               RowHeightMax    =   0
               ColWidthMin     =   0
               ColWidthMax     =   0
               ExtendLastCol   =   -1  'True
               FormatString    =   $"FrmVizitScreen.frx":57E37
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
         Begin VB.Frame Frame11 
            Caption         =   "Œ’„ «·»—Ê„Ê‘‰"
            Height          =   1695
            Index           =   2
            Left            =   60
            RightToLeft     =   -1  'True
            TabIndex        =   520
            Top             =   5730
            Width           =   5745
            Begin VB.TextBox XPTxtID 
               Alignment       =   1  'Right Justify
               Height          =   375
               Index           =   2
               Left            =   0
               RightToLeft     =   -1  'True
               TabIndex        =   562
               Top             =   0
               Width           =   645
            End
            Begin VB.CommandButton cmdcreate 
               Caption         =   "«‰‘«¡ «‘⁄«— ¬·Ì"
               Height          =   1425
               Index           =   2
               Left            =   60
               RightToLeft     =   -1  'True
               TabIndex        =   521
               Top             =   180
               Width           =   615
            End
            Begin VSFlex8Ctl.VSFlexGrid grdAcc 
               Height          =   1380
               Index           =   2
               Left            =   690
               TabIndex        =   522
               Top             =   210
               Width           =   5040
               _cx             =   8890
               _cy             =   2434
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
               BackColor       =   65535
               ForeColor       =   -2147483640
               BackColorFixed  =   14871017
               ForeColorFixed  =   -2147483630
               BackColorSel    =   -2147483635
               ForeColorSel    =   -2147483634
               BackColorBkg    =   -2147483636
               BackColorAlternate=   65535
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
               Rows            =   4
               Cols            =   5
               FixedRows       =   1
               FixedCols       =   1
               RowHeightMin    =   320
               RowHeightMax    =   0
               ColWidthMin     =   0
               ColWidthMax     =   0
               ExtendLastCol   =   -1  'True
               FormatString    =   $"FrmVizitScreen.frx":57EFA
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
         Begin VB.Frame Frame11 
            Caption         =   "Œ’„ «·„’«—Ì› «· ”ÊÌﬁÌ…"
            Height          =   1695
            Index           =   1
            Left            =   5940
            RightToLeft     =   -1  'True
            TabIndex        =   517
            Top             =   5700
            Width           =   5745
            Begin VB.TextBox XPTxtID 
               Alignment       =   1  'Right Justify
               Height          =   375
               Index           =   1
               Left            =   0
               RightToLeft     =   -1  'True
               TabIndex        =   561
               Top             =   0
               Width           =   645
            End
            Begin VB.CommandButton cmdcreate 
               Caption         =   "«‰‘«¡ «‘⁄«— ¬·Ì"
               Height          =   1425
               Index           =   1
               Left            =   60
               RightToLeft     =   -1  'True
               TabIndex        =   518
               Top             =   180
               Width           =   615
            End
            Begin VSFlex8Ctl.VSFlexGrid grdAcc 
               Height          =   1380
               Index           =   1
               Left            =   690
               TabIndex        =   519
               Top             =   210
               Width           =   5040
               _cx             =   8890
               _cy             =   2434
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
               BackColor       =   12632256
               ForeColor       =   -2147483640
               BackColorFixed  =   14871017
               ForeColorFixed  =   -2147483630
               BackColorSel    =   -2147483635
               ForeColorSel    =   -2147483634
               BackColorBkg    =   -2147483636
               BackColorAlternate=   12632256
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
               Rows            =   4
               Cols            =   5
               FixedRows       =   1
               FixedCols       =   1
               RowHeightMin    =   320
               RowHeightMax    =   0
               ColWidthMin     =   0
               ColWidthMax     =   0
               ExtendLastCol   =   -1  'True
               FormatString    =   $"FrmVizitScreen.frx":57FBD
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
         Begin VB.Frame Frame11 
            Caption         =   "«·Œ’„ «·À«»  Rebate"
            Height          =   1695
            Index           =   0
            Left            =   11700
            RightToLeft     =   -1  'True
            TabIndex        =   514
            Top             =   5670
            Width           =   5865
            Begin VB.TextBox XPTxtID 
               Alignment       =   1  'Right Justify
               Height          =   375
               Index           =   0
               Left            =   30
               RightToLeft     =   -1  'True
               TabIndex        =   560
               Top             =   1170
               Width           =   645
            End
            Begin VB.CommandButton cmdcreate 
               Caption         =   "«‰‘«¡ «‘⁄«— ¬·Ì"
               Height          =   795
               Index           =   0
               Left            =   30
               RightToLeft     =   -1  'True
               TabIndex        =   516
               Top             =   240
               Width           =   615
            End
            Begin VSFlex8Ctl.VSFlexGrid grdAcc 
               Height          =   1380
               Index           =   0
               Left            =   660
               TabIndex        =   515
               Top             =   240
               Width           =   5130
               _cx             =   9049
               _cy             =   2434
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
               BackColor       =   12615935
               ForeColor       =   -2147483640
               BackColorFixed  =   14871017
               ForeColorFixed  =   -2147483630
               BackColorSel    =   -2147483635
               ForeColorSel    =   -2147483634
               BackColorBkg    =   -2147483636
               BackColorAlternate=   12615935
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
               Rows            =   4
               Cols            =   5
               FixedRows       =   1
               FixedCols       =   1
               RowHeightMin    =   320
               RowHeightMax    =   0
               ColWidthMin     =   0
               ColWidthMax     =   0
               ExtendLastCol   =   -1  'True
               FormatString    =   $"FrmVizitScreen.frx":58080
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
         Begin VSFlex8Ctl.VSFlexGrid grd 
            Height          =   2220
            Index           =   0
            Left            =   60
            TabIndex        =   511
            Top             =   1065
            Width           =   17460
            _cx             =   30797
            _cy             =   3916
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
            FormatString    =   $"FrmVizitScreen.frx":58143
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
            WallPaperAlignment=   9
            AccessibleName  =   ""
            AccessibleDescription=   ""
            AccessibleValue =   ""
            AccessibleRole  =   24
         End
         Begin VSFlex8Ctl.VSFlexGrid grd 
            Height          =   1695
            Index           =   1
            Left            =   30
            TabIndex        =   513
            Top             =   3630
            Width           =   17460
            _cx             =   30797
            _cy             =   2990
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
            FormatString    =   $"FrmVizitScreen.frx":58381
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
            WallPaperAlignment=   9
            AccessibleName  =   ""
            AccessibleDescription=   ""
            AccessibleValue =   ""
            AccessibleRole  =   24
         End
         Begin ImpulseButton.ISButton btn_New 
            Height          =   360
            Index           =   10
            Left            =   7350
            TabIndex        =   535
            Top             =   8850
            Width           =   1125
            _ExtentX        =   1984
            _ExtentY        =   635
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
            ButtonImage     =   "FrmVizitScreen.frx":585C2
            ColorButton     =   14871017
            DrawFocusRectangle=   0   'False
         End
         Begin ImpulseButton.ISButton btn_Save 
            Height          =   375
            Index           =   10
            Left            =   5415
            TabIndex        =   536
            Top             =   8805
            Width           =   825
            _ExtentX        =   1455
            _ExtentY        =   661
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
            ButtonImage     =   "FrmVizitScreen.frx":5895C
            ColorButton     =   14871017
            DrawFocusRectangle=   0   'False
         End
         Begin ImpulseButton.ISButton btn_Modify 
            Height          =   360
            Index           =   10
            Left            =   6375
            TabIndex        =   537
            Top             =   8880
            Width           =   975
            _ExtentX        =   1720
            _ExtentY        =   635
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
            ButtonImage     =   "FrmVizitScreen.frx":58CF6
            ColorButton     =   14871017
            DrawFocusRectangle=   0   'False
         End
         Begin ImpulseButton.ISButton Btn_Undo 
            Height          =   360
            Index           =   10
            Left            =   4425
            TabIndex        =   538
            Top             =   8775
            Width           =   990
            _ExtentX        =   1746
            _ExtentY        =   635
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
            ButtonImage     =   "FrmVizitScreen.frx":59090
            ColorButton     =   14871017
            DrawFocusRectangle=   0   'False
         End
         Begin ImpulseButton.ISButton btn_Delete 
            Height          =   375
            Index           =   10
            Left            =   3585
            TabIndex        =   539
            Top             =   8805
            Width           =   840
            _ExtentX        =   1482
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
            ButtonImage     =   "FrmVizitScreen.frx":5942A
            ColorButton     =   14871017
            DrawFocusRectangle=   0   'False
         End
         Begin ImpulseButton.ISButton btn_Cancel 
            Height          =   360
            Index           =   10
            Left            =   240
            TabIndex        =   540
            Top             =   8745
            Width           =   975
            _ExtentX        =   1720
            _ExtentY        =   635
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
            ButtonImage     =   "FrmVizitScreen.frx":599C4
            ColorButton     =   14871017
            DrawFocusRectangle=   0   'False
            DisabledImageStyle=   1
         End
         Begin ImpulseButton.ISButton Btn_Print 
            Height          =   465
            Index           =   10
            Left            =   2475
            TabIndex        =   541
            TabStop         =   0   'False
            ToolTipText     =   "ÿ»«⁄… «·»Ì«‰«  "
            Top             =   8745
            Width           =   975
            _ExtentX        =   1720
            _ExtentY        =   820
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
            ButtonImage     =   "FrmVizitScreen.frx":59D5E
            ColorButton     =   14871017
            DrawFocusRectangle=   0   'False
            DisabledImageStyle=   1
         End
         Begin ImpulseButton.ISButton btn_Query 
            Height          =   570
            Index           =   10
            Left            =   1635
            TabIndex        =   542
            TabStop         =   0   'False
            ToolTipText     =   "(Ctrl+F)  ··»ÕÀ ≈÷€ÿ Â–« «·„› «Õ √Ê ≈÷€ÿ "
            Top             =   8670
            Width           =   705
            _ExtentX        =   1244
            _ExtentY        =   1005
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
            ButtonImage     =   "FrmVizitScreen.frx":605C0
            ColorButton     =   14871017
            DrawFocusRectangle=   0   'False
         End
         Begin MSComCtl2.DTPicker XPDtbTrans10 
            Height          =   315
            Left            =   3180
            TabIndex        =   543
            Top             =   690
            Width           =   1800
            _ExtentX        =   3175
            _ExtentY        =   556
            _Version        =   393216
            Format          =   215547905
            CurrentDate     =   38784
         End
         Begin MSDataListLib.DataCombo DcCustmer 
            Height          =   315
            Index           =   10
            Left            =   5850
            TabIndex        =   552
            Top             =   660
            Width           =   4185
            _ExtentX        =   7382
            _ExtentY        =   556
            _Version        =   393216
            Text            =   ""
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
            Height          =   345
            Left            =   11790
            TabIndex        =   553
            TabStop         =   0   'False
            Top             =   630
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
            Format          =   215547907
            CurrentDate     =   37140
         End
         Begin MSComCtl2.DTPicker ToDate1 
            Height          =   345
            Left            =   14430
            TabIndex        =   554
            TabStop         =   0   'False
            Top             =   630
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
            Format          =   215547907
            CurrentDate     =   37140
         End
         Begin MSComCtl2.DTPicker FromDate 
            Height          =   345
            Left            =   10050
            TabIndex        =   556
            TabStop         =   0   'False
            Top             =   630
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
            Format          =   192806915
            CurrentDate     =   37140
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "TOTAL Notes"
            Height          =   285
            Index           =   41
            Left            =   -1620
            RightToLeft     =   -1  'True
            TabIndex        =   558
            Top             =   7560
            Width           =   3750
         End
         Begin VB.Label Label1 
            BackColor       =   &H00E2E9E9&
            BackStyle       =   0  'Transparent
            Caption         =   "Transfer Date : "
            Height          =   420
            Index           =   40
            Left            =   2070
            TabIndex        =   551
            Top             =   750
            Width           =   1095
         End
         Begin VB.Label Label1 
            BackColor       =   &H00E2E9E9&
            BackStyle       =   0  'Transparent
            Caption         =   "No"
            Height          =   420
            Index           =   23
            Left            =   150
            TabIndex        =   550
            Top             =   750
            Width           =   1095
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "TOTAL PAYMENT MADE"
            Height          =   285
            Index           =   39
            Left            =   3870
            RightToLeft     =   -1  'True
            TabIndex        =   531
            Top             =   5370
            Width           =   3750
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "Net sales after return"
            Height          =   285
            Index           =   38
            Left            =   3660
            RightToLeft     =   -1  'True
            TabIndex        =   529
            Top             =   3360
            Width           =   3750
         End
         Begin VB.Label Label1 
            BackColor       =   &H00E2E9E9&
            BackStyle       =   0  'Transparent
            Caption         =   "Customer"
            Height          =   420
            Index           =   37
            Left            =   5040
            TabIndex        =   512
            Top             =   720
            Width           =   1095
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
Dim ii As Long
Dim cSearch  As clsDCboSearch
Dim MintDone As Integer
Public mIndex As Integer
Dim mIndex2 As Integer
Dim Dcombos As ClsDataCombos
Dim mGridClicked As Boolean
Dim mDiscEnter As Boolean
Dim rsDummy As New ADODB.Recordset
Dim s As String
Dim i As Long
Dim j As Long
Public mToStoreID As Long
Dim mItemIDS As String
Dim mDateTrans As Date
Dim zatcaStatus As Integer
Dim mIndexVat As Integer
Dim Export As Integer
Public mTypeInvoice As Integer


Private Sub BtnPrint_Click(Index As Integer)
    
    If Index = 2 Then
    print_report66_proc
    Else
   print_report66 Index
   End If
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
   
 
MySQL = " Select TblAging.*,a.Account_Serial,ta.aqarNo,ta.aqarname"
MySQL = MySQL & " from TblAging LEFT OUTER JOIN ACCOUNTS AS a ON a.Account_Code = TblAging.Account_Code LEFT OUTER JOIN TblCustemers AS tc ON tc.Account_Code = a.Account_Code"
MySQL = MySQL & " LEFT OUTER JOIN TblAqar AS ta ON tc.CusID = ta.ownerid"
MySQL = MySQL & " Order By TblAging.AGEID"
    
    If Ind = 0 Then
    
     
            MySQL = " Select TblAging.*,a.Account_Serial,ta.aqarNo,ta.aqarname"
        MySQL = MySQL & " from TblAging LEFT OUTER JOIN ACCOUNTS AS a ON a.Account_Code = TblAging.Account_Code LEFT OUTER JOIN TblCustemers AS tc ON tc.Account_Code = a.Account_Code"
        MySQL = MySQL & " LEFT OUTER JOIN TblAqar AS ta ON tc.CusID = ta.ownerid "
        MySQL = MySQL & " WHERE ISNULL(StillAmount,0) <> 0"
        MySQL = MySQL & " Order By TblAging.AGEID"
            

     '   MySQL = "SELECT * FROM TblAging WHERE ISNULL(StillAmount,0) <> 0"
        If SystemOptions.UserInterface = ArabicInterface Then
            StrFileName = App.path & "\REPORTS\REPORTS NEW\" & "Aging1.rpt"
        Else
            StrFileName = App.path & "\REPORTS\REPORTS NEW\" & "Aging1E.rpt"
        End If
    Else
        MySQL = " SELECT CONCAT("
         MySQL = MySQL & "     TblCustemers.CusName,"
          MySQL = MySQL & "    CHAR(13) + CHAR(10),"
         MySQL = MySQL & "     N'„‰œÊ» : ',"
          MySQL = MySQL & "    ISNULL(TblEmployee.Emp_Name, '')"
       MySQL = MySQL & " ) AS CusName22"
        MySQL = MySQL & " ,TblEmployee.Emp_Code"
        MySQL = MySQL & " ,TblEmployee.Emp_Name"
        MySQL = MySQL & " ,TblEmployee.Emp_Namee"
        MySQL = MySQL & " ,TblCustemers.EmpId ,TblAging.* FROM TblAging INNER JOIN Ageng_type ON Ageng_type.id = TblAging.AGEID"
        MySQL = MySQL & " LEFT OUTER JOIN TblCustemers"
        MySQL = MySQL & " ON TblAging.Account_Code = TblCustemers.Account_Code"
        MySQL = MySQL & " LEFT OUTER JOIN TblEmployee"
        MySQL = MySQL & " ON TblCustemers.EmpId = TblEmployee.Emp_ID"
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
        MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
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
    xReport.ApplicationName = App.Title
    xReport.ReportAuthor = App.Title
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
    Dim mCusType As String
    Dim StrFileName As String
    Dim Msg As String
       Dim mSql1 As String
    Dim mSql2 As String

    Dim X As Integer
    MySQL = "Select TblCustemers.CusName"
   MySQL = MySQL & " ,TblEmployee.Emp_Code"
   MySQL = MySQL & ",TblEmployee.Emp_Name"
   MySQL = MySQL & ",TblEmployee.Emp_Namee"
   MySQL = MySQL & ",TblCustemers.EmpId"
   MySQL = MySQL & ",TblAging.* from TblAging  "
    MySQL = MySQL & "LEFT OUTER JOIN TblCustemers"
    MySQL = MySQL & " ON TblAging.Account_Code = TblCustemers.Account_Code"
    MySQL = MySQL & " LEFT OUTER JOIN TblEmployee"
    MySQL = MySQL & " ON TblCustemers.EmpId = TblEmployee.Emp_ID"
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
         If StrCusID.text <> "" Then
            MySQL = MySQL & " and TblAging.CusID in (" & (StrCusID.text) & ")"
        End If
    End If
     MySQL = MySQL & " ORDER BY DueDate  "
    RsData.Open MySQL, Cn, adOpenKeyset, adLockReadOnly
    If Not RsData.EOF Then
          If SystemOptions.UserInterface = ArabicInterface Then
                        X = MsgBox("ÌÊÃœ ⁄„— œÌ‰  „ ⁄„·Â „”»ﬁ« » «—ÌŒ " & DTP_Date.value & "" & " Â·  Êœ ⁄—÷Â ‰⁄„/·«", vbInformation + vbYesNo)
                    Else
                        X = MsgBox("No Contract For This Employee Create Contarct y / n", vbInformation + vbYesNo)
                    End If
     
                    If X = vbYes Then
                        loadgrid MySQL, grdAging, True, False
                        PrintAging Ind
                        Exit Function
                    End If
    End If

    RsData.Close
 
    

    
    grdAging2.rows = 1
    grdAging.rows = 1
    
Dim mWhereCus As String


'-------------------------------
   Dim mCusTypeStr As String
  
    If Option1(2).value = True Or Option4.value Then
        mCusType = 1
        mCusTypeStr = "(1,56)"
    Else
        mCusType = 2
        mCusTypeStr = "(2,57)"
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
MySQL = MySQL & "                   XB.NoteSerial1,"
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
MySQL = MySQL & "                   Notes.NoteSerial1,"
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

If Trim(Dcbranch(8).text) <> "" Then
    MySQL = MySQL & "          AND dev.branch_id = " & val(Dcbranch(8).BoundText)
End If
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
MySQL = MySQL & "                   Notes1.NoteSerial1,"
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
If Trim(Dcbranch(8).text) <> "" Then
    MySQL = MySQL & "          AND dev.branch_id = " & val(Dcbranch(8).BoundText)
End If
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
        mCusTypeStr = "(1)"
        'MySQL = MySQL & " and Account_Code  In (Select Account_Code from TblCustemers Where  Type = " & mCusType & ")"
        MySQL = MySQL & " and Account_Code  In (Select Account_Code from TblCustemers Where  Type In " & mCusTypeStr & ")"
        
    ElseIf Option2.value = True Then
        mCusType = 2
        mCusTypeStr = "(2) "
        MySQL = MySQL & " and Account_Code  In (Select Account_Code from TblCustemers Where  Type In " & mCusTypeStr & ")"
        ElseIf Option3.value = True Then
        mCusType = 57
        mCusTypeStr = "(57) "
        MySQL = MySQL & " and Account_Code  In (Select Account_Code from TblCustemers Where  Type In " & mCusTypeStr & ")"
        ElseIf Option4.value = True Then
        mCusType = 56
        mCusTypeStr = "(56) "
        MySQL = MySQL & " and Account_Code  In (Select Account_Code from TblCustemers Where  Type In " & mCusTypeStr & ")"
    End If
    If ChekCustomer.value = vbChecked Then
        If val(DBCboClientName.BoundText) <> 0 Then
            MySQL = MySQL & " and Account_Code  In (Select Account_Code from TblCustemers Where  TblCustemers.CusId = " & val(DBCboClientName.BoundText) & " and Type In " & mCusTypeStr & ")"
            'MySQL = MySQL & " and Account_Code  In (Select Account_Code from TblCustemers Where  Type In " & mCusType & ")"
        Else
            
        End If

    Else
         If val(DBCboClientName.BoundText) <> 0 Then
            MySQL = MySQL & " and Account_Code  In (Select Account_Code from TblCustemers Where  TblCustemers.CusId = " & val(DBCboClientName.BoundText) & " )"
        
        End If
        
    End If
    
        If dcClass.text <> "" And val(dcClass.BoundText) <> 0 Then
            
             MySQL = MySQL & " and Account_Code  In (Select Account_Code from TblCustemers Where  TblCustemers.ClassCustomersId = " & val(dcClass.BoundText) & " and Type In " & mCusTypeStr & ")"
        End If
        
            If (DcCustomerType.text) <> "" And val(DcCustomerType.BoundText) <> 0 Then
            
             MySQL = MySQL & " and Account_Code  In (Select Account_Code from TblCustemers Where  TblCustemers.CustomerTypeID = " & val(DcCustomerType.BoundText) & " and Type In " & mCusTypeStr & ")"
        End If
        
    If CheckAllCustomer.value = vbChecked Then
         If StrCusID.text <> "" Then
           ' MySQL = MySQL & " and TblCustemers.CusID in (" & (StrCusID.Text) & ")"
            MySQL = MySQL & " and Account_Code  In ( Select  TblCustemers.Account_Code from TblCustemers Where  Type In " & mCusTypeStr & " and  TblCustemers.CusID in (" & (StrCusID.text) & ") )"
        End If
    Else
        If StrCusID.text <> "" Then
           ' MySQL = MySQL & " and TblCustemers.CusID in (" & (StrCusID.Text) & ")"
            MySQL = MySQL & " and Account_Code  In ( Select  TblCustemers.Account_Code from TblCustemers Where  Type In " & mCusTypeStr & " and TblCustemers.CusID in (" & (StrCusID.text) & ") )"
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
           
If mCusType = 1 Or mCusType = 56 Then
MySQL = MySQL & " Order By"
MySQL = MySQL & "       Account_Code,"
MySQL = MySQL & "       XB.NoteSerial,"
MySQL = MySQL & "                   XB.NoteSerial1,"
MySQL = MySQL & "       Xb.DueDate"
'
Else
MySQL = MySQL & " GROUP BY Account_Code   "
End If
   
    
    
    mSql1 = MySQL
    
  
  
   
   MySQL = ""
   
If mCusType = 2 Or mCusType = 57 Then
    MySQL = MySQL & " SELECT SUM(XB.TransNet) AS TransNet ,SUM(XB.TransNet) AS TransNet22,Account_Code,CusName"


Else
      MySQL = MySQL & " Select  SUM (TransNet) AS TransNet,Account_Code,CusName"

End If
   
   


MySQL = ""
MySQL = MySQL & " SELECT "

If mCusType = 1 Or mCusType = 56 Then
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
MySQL = MySQL & "                   XB.NoteSerial1,"
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
MySQL = MySQL & "                   Notes.NoteSerial1,"
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
If Trim(Dcbranch(8).text) <> "" Then
    MySQL = MySQL & "          AND dev.branch_id = " & val(Dcbranch(8).BoundText)
End If
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
MySQL = MySQL & "                   Notes1.NoteSerial1,"
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
If Trim(Dcbranch(8).text) <> "" Then
    MySQL = MySQL & "          AND dev.branch_id = " & val(Dcbranch(8).BoundText)
End If
MySQL = MySQL & "       ) XB"
MySQL = MySQL & "       Right OUTER JOIN dbo.Ageng_type"
MySQL = MySQL & "            ON  XB.AgeID = dbo.Ageng_type.id"
MySQL = MySQL & " Where 1 = 1"
'
If Not IsNull(DTP_Date.value) Then
    MySQL = MySQL & " and XB.DueDate <=" & SQLDate(DTP_Date.value, True) & ""
End If
'    If Option1(2).value = True Then
        MySQL = MySQL & " and Account_Code  In (Select Account_Code from TblCustemers Where  Type In " & mCusTypeStr & ")"
'    ElseIf Option2.value = True Then
'        MySQL = MySQL & " and Account_Code  In (Select Account_Code from TblCustemers Where  Type In " & mCusTypeStr & ")"
'    End If
    

     If ChekCustomer.value = vbChecked Then
        If val(DBCboClientName.BoundText) <> 0 Then
            MySQL = MySQL & " and Account_Code  In (Select Account_Code from TblCustemers Where   Type In " & mCusTypeStr & " and TblCustemers.CusId = " & val(DBCboClientName.BoundText) & " )"
        
        End If
    Else
         If val(DBCboClientName.BoundText) <> 0 Then
            MySQL = MySQL & " and Account_Code  In (Select Account_Code from TblCustemers Where   Type In " & mCusTypeStr & " and TblCustemers.CusId = " & val(DBCboClientName.BoundText) & " )"
        
        End If
    End If
    
            
       If CheckAllCustomer.value = vbChecked Then
         If StrCusID.text <> "" Then
           ' MySQL = MySQL & " and TblCustemers.CusID in (" & (StrCusID.Text) & ")"
            MySQL = MySQL & " and Account_Code  In ( Select  TblCustemers.Account_Code from TblCustemers Where  Type In " & mCusTypeStr & " and  TblCustemers.CusID in (" & (StrCusID.text) & ") )"
        End If
    Else
        If StrCusID.text <> "" Then
           ' MySQL = MySQL & " and TblCustemers.CusID in (" & (StrCusID.Text) & ")"
            MySQL = MySQL & " and Account_Code  In ( Select  TblCustemers.Account_Code from TblCustemers Where  Type In " & mCusTypeStr & " and TblCustemers.CusID in (" & (StrCusID.text) & ") )"
        End If
    End If
         
            If Trim(dcClass.text) <> "" And val(dcClass.BoundText) <> 0 Then
            
             MySQL = MySQL & " and Account_Code  In (Select Account_Code from TblCustemers Where   Type In " & mCusTypeStr & " and TblCustemers.ClassCustomersId = " & val(dcClass.BoundText) & " and Type In " & mCusTypeStr & ")"
        End If
  
            
            If Trim(DcCustomerType.text) <> "" And val(DcCustomerType.BoundText) <> 0 Then
            
             MySQL = MySQL & " and Account_Code  In (Select Account_Code from TblCustemers Where  TblCustemers.CustomerTypeID = " & val(DcCustomerType.BoundText) & " and Type In " & mCusTypeStr & ")"
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


If mCusType = 2 Or mCusType = 57 Then
MySQL = MySQL & " Order By"
MySQL = MySQL & "       Account_Code,"
MySQL = MySQL & "       XB.NoteSerial,"
MySQL = MySQL & "                   XB.NoteSerial1,"
MySQL = MySQL & "       Xb.DueDate"
'
Else
MySQL = MySQL & " GROUP BY Account_Code  ,CusName"
End If
   


'MySQL = MySQL & " ,AgeID"

   
   
   mSql2 = MySQL

    If Option1(2).value = True Or Option4.value Then
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
   For i = 1 To grdAging.rows - 1
   
     
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
    
    
    
    saveGrid s, grdAging, "StillAmount", "", "Credit_Or_Debit", 0



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
            rs!ageid = rsDummyT2!ID
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



''========================================
'' «·ÃœÌœ: ‘€· √⁄„«— «·œÌÊ‰ „‰ «·‹ Stored Procedure
''========================================
'Public Function print_report66_proc()
'    Dim Cmd As New ADODB.Command
'    Dim rs As ADODB.Recordset
'
'    Dim mIsCustomer As Boolean
'    Dim mBranchID As Variant
'    Dim mCusId As Variant
'    Dim mCusList As Variant
'    Dim mClassId As Variant
'    Dim mCustTypeId As Variant
'    Dim mEmpId As Variant
'
'    ' 1) ‰Õœœ ÂÊ ⁄„Ì· Ê·« „Ê—œ „‰ «Œ Ì«—« ﬂ «·Õ«·Ì…
'    If Option1(2).value = True Or Option4.value = True Then   ' ⁄„·«¡
'        mIsCustomer = True
'    Else                                                       ' „Ê—œÌ‰
'        mIsCustomer = False
'    End If
'
'    ' 2) «·›—⁄
'    If Trim(dcBranch(8).text) <> "" Then
'        mBranchID = CLng(dcBranch(8).BoundText)
'    Else
'        mBranchID = Null
'    End If
'
'    ' 3) ⁄„Ì· Ê«Õœ
'    If ChekCustomer.value = vbChecked Then
'        If val(DBCboClientName.BoundText) <> 0 Then
'            mCusId = val(DBCboClientName.BoundText)
'        Else
'            mCusId = Null
'        End If
'    Else
'        mCusId = Null
'    End If
'
'    ' 4) ·Ì” … ⁄„·«¡
'    If CheckAllCustomer.value = vbChecked Then
'        If Trim(StrCusID.text) <> "" Then
'            mCusList = StrCusID.text       ' „À«·: "12,15,20"
'        Else
'            mCusList = Null
'        End If
'    Else
'        mCusList = Null
'    End If
'
'    ' 5) ﬂ·«” ⁄„Ì·
'    If Trim(dcClass.text) <> "" And val(dcClass.BoundText) <> 0 Then
'        mClassId = val(dcClass.BoundText)
'    Else
'        mClassId = Null
'    End If
'
'    ' 6) ‰Ê⁄ ⁄„Ì·
'    If Trim(DcCustomerType.text) <> "" And val(DcCustomerType.BoundText) <> 0 Then
'        mCustTypeId = val(DcCustomerType.BoundText)
'    Else
'        mCustTypeId = Null
'    End If
'
'    ' 7) „ÊŸ›
'    If CheckEmp.value = vbChecked Then
'        If val(DcbEmployee.BoundText) <> 0 Then
'            mEmpId = val(DcbEmployee.BoundText)
'        Else
'            mEmpId = Null
'        End If
'    Else
'        mEmpId = Null
'    End If
'
'    ' ‰Õ÷¯— «·√„—
'    Cmd.ActiveConnection = Cn
'    Cmd.CommandType = adCmdStoredProc
'    Cmd.CommandText = "Rpt_Aging_VBStyle"
'Dim sCusList As String
'sCusList = Trim$(StrCusID.text)   ' ?? ???CusList ?????
'
'
'
'    ' «·»—«„ —« 
'    Cmd.Parameters.Append Cmd.CreateParameter("@AsOfDate", adDate, adParamInput, , CDate(DTP_Date.value))
'    Cmd.Parameters.Append Cmd.CreateParameter("@IsCustomer", adBoolean, adParamInput, , mIsCustomer)
'    Cmd.Parameters.Append Cmd.CreateParameter("@BranchId", adInteger, adParamInput, , mBranchID)
'    Cmd.Parameters.Append Cmd.CreateParameter("@CusId", adInteger, adParamInput, , mCusId)
'
'
'sCusList = Trim$(StrCusID.text)   ' √Ê «·„CusList » «⁄ﬂ
'
'If sCusList = "" Then
'    ' „›Ì‘ ·Ì” … ? «»⁄  Null
'    Cmd.Parameters.Append Cmd.CreateParameter("@CusIdList", adVarWChar, adParamInput, 1, Null)
'ElseIf Len(sCusList) < 3900 Then
'    ' ·Ì” … „⁄ﬁÊ·… ? ‰»⁄ Â«
'    Cmd.Parameters.Append Cmd.CreateParameter("@CusIdList", adVarWChar, adParamInput, Len(sCusList), sCusList)
'Else
'    ' ·Ì” … „ÂÊ·… “Ï «··Ï ›Ï «·’Ê—… ? ADO »Ì ⁄»
'    ' Â‰« Ì« ≈„« ‰—Ã¯⁄ ·‹ Null Ê‰ÃÌ» «·ﬂ·° √Ê ‰⁄„· Õ· »œÌ·
'    Cmd.Parameters.Append Cmd.CreateParameter("@CusIdList", adVarWChar, adParamInput, 1, Null)
'End If
'
'
''    Cmd.Parameters.Append Cmd.CreateParameter("@CusIdList", adLongVarWChar, adParamInput, 4000, mCusList)
'    Cmd.Parameters.Append Cmd.CreateParameter("@ClassId", adInteger, adParamInput, , mClassId)
'    Cmd.Parameters.Append Cmd.CreateParameter("@CustomerTypeId", adInteger, adParamInput, , mCustTypeId)
'    Cmd.Parameters.Append Cmd.CreateParameter("@EmpId", adInteger, adParamInput, , mEmpId)
'
'    ' ‰‰›–
'    Set rs = Cmd.Execute
'
'    ' ‰›—¯€ «·Ã—ÌœÌ‰
'    grdAging.rows = 1
'    grdAging2.rows = 1
'
'
'
'  If mIsCustomer Then
'        ' ⁄„·«¡: √Ê· Ã—Ìœ = «·„œÌ‰° «· «‰Ì = «·œ«∆‰
'        If Not rs Is Nothing Then
'            loadgridRS rs, grdAging, True, False, True
'        End If
'        Set rs = rs.NextRecordset
'        If Not rs Is Nothing Then
'            loadgridRS rs, grdAging2, True, False, True
'        End If
'    Else
'        ' „Ê—œÌ‰: ‰⁄ﬂ” ⁄‘«‰ «··Ê» «·ﬁœÌ„ Ì‘ €· ’Õ
'        If Not rs Is Nothing Then
'            loadgridRS rs, grdAging2, True, False, True   ' «·œ«∆‰
'        End If
'        Set rs = rs.NextRecordset
'        If Not rs Is Nothing Then
'            loadgridRS rs, grdAging, True, False, True    ' «·„œÌ‰
'        End If
'    End If
'
'    ' «·ÃœÊ· «·√Ê· («·„œÌ‰)
''   If Not rs Is Nothing Then
''        loadgridRS rs, grdAging, True, False, True   ' True «·«ŒÌ—… ·Ê ⁄«Ì“  ⁄Ìœ «·√⁄„œ…
''    End If
''
''    ' «·ÃœÊ· «· «‰Ì («·œ«∆‰)
''    Set rs = rs.NextRecordset
''    If Not rs Is Nothing Then
''        loadgridRS rs, grdAging2, True, False, True
''    End If
'
'    ' »⁄œ „« «·Ã—ÌœÌ‰ « „·«Ê« ‰⁄„· ‰›” «··›… «·ﬁœÌ„… » «⁄  «·„ÿ«»ﬁ«  Ê«·Õ›Ÿ Ê«·ÿ»«⁄…
'    Call Aging_PostProcessAndPrint
'
'End Function
'========================================
' œÂ «·Ã“¡ «·ﬁœÌ„ » «⁄ «· Ê“Ì⁄ Ê«·Õ›Ÿ Ê«·ÿ»«⁄…
' Œ·Ì‰«Â ›Ì ›‰ﬂ‘‰ ·ÊÕœÂ


'========================================
' «·ÃœÌœ: ‘€· √⁄„«— «·œÌÊ‰ „‰ «·‹ Stored Procedure
'========================================
Public Function print_report66_proc()
    Dim Cmd As ADODB.Command
    Dim rs  As ADODB.Recordset

    Dim mIsCustomer   As Boolean
    Dim mBranchID     As Variant
    Dim mCusId        As Variant
    Dim mClassId      As Variant
    Dim mCustTypeId   As Variant
    Dim mEmpId        As Variant
    Dim sCusList      As String

    '-------------------------------------
    ' 1) ‰Õœœ ÂÊ ⁄„Ì· Ê·« „Ê—œ „‰ «Œ Ì«—« ﬂ «·Õ«·Ì…
    '-------------------------------------
    If Option1(2).value = True Or Option4.value = True Then   ' ⁄„·«¡
        mIsCustomer = True
    Else                                                       ' „Ê—œÌ‰
        mIsCustomer = False
    End If

    '-------------------------------------
    ' 2) «·›—⁄
    '-------------------------------------
    If Trim$(Dcbranch(8).text) <> "" Then
        mBranchID = CLng(Dcbranch(8).BoundText)
    Else
        mBranchID = Null
    End If

    '-------------------------------------
    ' 3) ⁄„Ì· Ê«Õœ
    '-------------------------------------
    If ChekCustomer.value = vbChecked Then
        If val(DBCboClientName.BoundText) <> 0 Then
            mCusId = val(DBCboClientName.BoundText)
        Else
            mCusId = Null
        End If
    Else
        mCusId = Null
    End If

    '-------------------------------------
    ' 4) ·Ì” … ⁄„·«¡ / „Ê—œÌ‰
    '-------------------------------------
    sCusList = Trim$(StrCusID.text)   ' ‰›” «··Ì ﬂ‰  ⁄«„· »ÌÂ IN (...)

    '-------------------------------------
    ' 5) ﬂ·«” ⁄„Ì·
    '-------------------------------------
    If Trim$(dcClass.text) <> "" And val(dcClass.BoundText) <> 0 Then
        mClassId = val(dcClass.BoundText)
    Else
        mClassId = Null
    End If

    '-------------------------------------
    ' 6) ‰Ê⁄ ⁄„Ì·
    '-------------------------------------
    If Trim$(DcCustomerType.text) <> "" And val(DcCustomerType.BoundText) <> 0 Then
        mCustTypeId = val(DcCustomerType.BoundText)
    Else
        mCustTypeId = Null
    End If

    '-------------------------------------
    ' 7) „ÊŸ›
    '-------------------------------------
    If CheckEmp.value = vbChecked Then
        If val(DcbEmployee.BoundText) <> 0 Then
            mEmpId = val(DcbEmployee.BoundText)
        Else
            mEmpId = Null
        End If
    Else
        mEmpId = Null
    End If

    '-------------------------------------
    ' ‰Õ÷¯— «·√„—
    '-------------------------------------
    Set Cmd = New ADODB.Command
    Cmd.ActiveConnection = Cn
    Cmd.CommandType = adCmdStoredProc
    Cmd.CommandText = "Rpt_Aging_VBStyle"

    ' «·»—«„ —«  »«· — Ì»
    Cmd.Parameters.Append Cmd.CreateParameter("@AsOfDate", adDate, adParamInput, , CDate(DTP_Date.value))
    Cmd.Parameters.Append Cmd.CreateParameter("@IsCustomer", adBoolean, adParamInput, , mIsCustomer)
    Cmd.Parameters.Append Cmd.CreateParameter("@BranchId", adInteger, adParamInput, , mBranchID)
    Cmd.Parameters.Append Cmd.CreateParameter("@CusId", adInteger, adParamInput, , mCusId)

    ' ·Ì” … «·⁄„·«¡/«·„Ê—œÌ‰
    If sCusList = "" Then
        ' „›Ì‘ ·Ì” … ? «»⁄  Null
        Cmd.Parameters.Append Cmd.CreateParameter("@CusIdList", adVarWChar, adParamInput, 1, Null)
    ElseIf Len(sCusList) < 3900 Then
        ' ·Ì” … „⁄ﬁÊ·… ? «»⁄ Â«
        Cmd.Parameters.Append Cmd.CreateParameter("@CusIdList", adVarWChar, adParamInput, Len(sCusList), sCusList)
    Else
        ' ·Ì” … „ÂÊ·… ? «»⁄  Null ÊŒ·Ì «·” Ê—œ Ì—Ã¯⁄ «·ﬂ·
        Cmd.Parameters.Append Cmd.CreateParameter("@CusIdList", adVarWChar, adParamInput, 1, Null)
    End If

    Cmd.Parameters.Append Cmd.CreateParameter("@ClassId", adInteger, adParamInput, , mClassId)
    Cmd.Parameters.Append Cmd.CreateParameter("@CustomerTypeId", adInteger, adParamInput, , mCustTypeId)
    Cmd.Parameters.Append Cmd.CreateParameter("@EmpId", adInteger, adParamInput, , mEmpId)

    '-------------------------------------
    ' ‰‰›¯–
    '-------------------------------------
    Set rs = Cmd.Execute

    ' œ«Ì„«: √Ê· ÃœÊ· = «·„” Õﬁ«  («·›Ê« Ì—)
    If Not rs Is Nothing Then
        loadgridRS rs, grdAging, True, False, True
    End If

    ' œ«Ì„«:  «‰Ì ÃœÊ· = «·”œ«œ
    Set rs = rs.NextRecordset
    If Not rs Is Nothing Then
        loadgridRS rs, grdAging2, True, False, True
    End If

    '-------------------------------------
    ' ‰›” «· Ê“Ì⁄ Ê«·Õ›Ÿ Ê«·ÿ»«⁄… «·ﬁœÌ„…
    '-------------------------------------
    Aging_PostProcessAndPrint_Light
End Function

Public Sub Aging_PostProcessAndPrint_Light()
    Dim i As Long, j As Long
    Dim mValue As Double, mPayedValue As Double, need As Double, takeAmt As Double
    Dim acct As String

    Dim idx_Acc As Long, idx_TransNet As Long, idx_Payed As Long, idx_Still As Long, idx_TNGrid2 As Long
    Dim idx2_Acc As Long, idx2_TransNet As Long, idx2_Payed As Long, idx2_Still As Long

    Dim dictTotal As Object ' ≈Ã„«·Ì «·”œ«œ ·ﬂ· Õ”«» („‰ «·Ã—Ìœ «· «‰Ì)
    Dim dictRemain As Object ' «·„ »ﬁÌ „‰ «·”œ«œ ·ﬂ· Õ”«» »⁄œ «· Ê“Ì⁄
    Dim usedLeft As Object   ' „ƒﬁ  · Ê“Ì⁄ «·‹ used ⁄·Ï ’›Ê› «·”œ«œ

    '-------------------------------------------
    ' 0)  Õ÷Ì— ≈‰œﬂ”«  «·√⁄„œ… »√„«‰
    '-------------------------------------------
    idx_Acc = grdAging.ColIndex("Account_Code")
    idx_TransNet = grdAging.ColIndex("TransNet")
    idx_Payed = grdAging.ColIndex("PayedValue")
    idx_Still = grdAging.ColIndex("StillAmount")
    idx_TNGrid2 = grdAging.ColIndex("TransNetGrid2") ' „„ﬂ‰ ÌﬂÊ‰ -1

    idx2_Acc = grdAging2.ColIndex("Account_Code")
    idx2_TransNet = grdAging2.ColIndex("TransNet")
    idx2_Payed = grdAging2.ColIndex("PayedValue")    ' „„ﬂ‰ -1
    idx2_Still = grdAging2.ColIndex("StillAmount")   ' „„ﬂ‰ -1

    If idx_Acc = -1 Or idx_TransNet = -1 Or idx_Payed = -1 Or idx_Still = -1 Or _
       idx2_Acc = -1 Or idx2_TransNet = -1 Then
        MsgBox "√⁄„œ… „ÿ·Ê»… €Ì— „ÊÃÊœ… ›Ì √Õœ «·Ã—ÌœÌ‰.", vbExclamation
        Exit Sub
    End If

    '-------------------------------------------
    ' 1) Ã„¯⁄ ﬂ· «·”œ«œ „‰ «·Ã—Ìœ «· «‰Ì Õ”» «·Õ”«»
    '    (»‰ﬁ—√ TransNet∫ Ê·Ê PayedValue „ÊÃÊœ Ê„‘ ’›— »‰«ŒœÂ ﬂ„—Ã⁄)
    '-------------------------------------------
    Set dictTotal = CreateObject("Scripting.Dictionary")
    For j = 1 To grdAging2.rows - 1
        acct = Trim$(grdAging2.TextMatrix(j, idx2_Acc))
        If acct <> "" Then
            Dim paidRaw As Double
            paidRaw = val(grdAging2.TextMatrix(j, idx2_TransNet))
            If idx2_Payed <> -1 Then
                ' ·Ê ›Ì ﬁÌ„… „”œœ… „ÊÃÊœ… ›⁄·« ›Ì «·Ã—Ìœ° «” Œœ„Â«
                Dim pv As Double
                pv = val(grdAging2.TextMatrix(j, idx2_Payed))
                If pv > 0 Then paidRaw = pv
            End If

            If paidRaw <> 0 Then
                If Not dictTotal.Exists(acct) Then dictTotal.Add acct, 0#
                dictTotal(acct) = dictTotal(acct) + paidRaw
            End If
        End If
    Next j

    ' «·„ »ﬁÌ = «·≈Ã„«·Ì ›Ì «·»œ«Ì…
    Set dictRemain = CreateObject("Scripting.Dictionary")
    Dim k As Variant
    For Each k In dictTotal.keys
        dictRemain.Add k, dictTotal(k)
    Next k

    '-------------------------------------------
    ' 2) Ê“¯⁄ «·”œ«œ «·„Ã„⁄ ⁄·Ï ›Ê« Ì— «·Ã—Ìœ «·√Ê·
    '-------------------------------------------
    txtTotalStill = ""
    For i = 1 To grdAging.rows - 1
        acct = Trim$(grdAging.TextMatrix(i, idx_Acc))
        mValue = val(grdAging.TextMatrix(i, idx_TransNet))
        mPayedValue = val(grdAging.TextMatrix(i, idx_Payed))

        need = mValue - mPayedValue
        If need > 0 And acct <> "" Then
            If dictRemain.Exists(acct) Then
                takeAmt = dictRemain(acct)
                If takeAmt > 0 Then
                    If takeAmt >= need Then
                        ' Ì”œ¯ »«·ﬂ«„·
                        grdAging.TextMatrix(i, idx_Payed) = CStr(mPayedValue + need)
                        dictRemain(acct) = takeAmt - need
                        ' √Ì »›«∆÷ ··”œ«œ ·Ì” „ÿ·Ê»« ≈ŸÂ«—Â ⁄·Ï „” ÊÏ «·›« Ê—… «·¬‰
                        If idx_TNGrid2 <> -1 Then grdAging.TextMatrix(i, idx_TNGrid2) = ""
                    Else
                        ' Ì”œ¯ Ã“∆Ì«
                        grdAging.TextMatrix(i, idx_Payed) = CStr(mPayedValue + takeAmt)
                        dictRemain(acct) = 0#
                        If idx_TNGrid2 <> -1 Then grdAging.TextMatrix(i, idx_TNGrid2) = ""
                    End If
                End If
            End If
        End If

        ' «Õ”» «·»«ﬁÌ ⁄·Ï «·›« Ê—… Ê √Œ›ˆ «·’› ·Ê ’›—
        Dim still As Double
        still = val(grdAging.TextMatrix(i, idx_TransNet)) - val(grdAging.TextMatrix(i, idx_Payed))
        If still = 0 Then
            grdAging.TextMatrix(i, idx_Still) = ""
            On Error Resume Next
            grdAging.RowHidden(i) = True
            On Error GoTo 0
        Else
            grdAging.TextMatrix(i, idx_Still) = CStr(still)
        End If

        txtTotalStill = val(txtTotalStill) + val(grdAging.TextMatrix(i, idx_Still))
    Next i

    '-------------------------------------------
    ' 3) —Ã¯⁄  Ê“Ì⁄ «·”œ«œ ÃÊ¯« «·Ã—Ìœ «· «‰Ì ‰›”Â
    '    (‰„·√ PayedValue Ê StillAmount ’›« ’›« ·ﬂ· Õ”«»)
    '-------------------------------------------
    Set usedLeft = CreateObject("Scripting.Dictionary")
    For Each k In dictTotal.keys
        Dim usedK As Double
        usedK = dictTotal(k) - dictRemain(k) ' «··Ì « ’—› ›⁄·«
        usedLeft.Add k, usedK
    Next k

    For j = 1 To grdAging2.rows - 1
        acct = Trim$(grdAging2.TextMatrix(j, idx2_Acc))
        If acct <> "" And usedLeft.Exists(acct) Then
            Dim rowTrans As Double, willPay As Double
            rowTrans = val(grdAging2.TextMatrix(j, idx2_TransNet))

            willPay = usedLeft(acct)
            If willPay <= 0 Then
                ' „›Ì‘ »«ﬁÚ ·Â–« «·Õ”«»
                If idx2_Payed <> -1 Then grdAging2.TextMatrix(j, idx2_Payed) = "0"
                If idx2_Still <> -1 Then grdAging2.TextMatrix(j, idx2_Still) = CStr(rowTrans)
            Else
                Dim payHere As Double
                If willPay >= rowTrans Then
                    payHere = rowTrans
                Else
                    payHere = willPay
                End If

                If idx2_Payed <> -1 Then grdAging2.TextMatrix(j, idx2_Payed) = CStr(payHere)
                If idx2_Still <> -1 Then grdAging2.TextMatrix(j, idx2_Still) = CStr(rowTrans - payHere)

                usedLeft(acct) = willPay - payHere
            End If
        Else
            ' ·« ÌÊÃœ «” Œœ«„ ·Â–« «·Õ”«»
            If idx2_Payed <> -1 Then grdAging2.TextMatrix(j, idx2_Payed) = "0"
            If idx2_Still <> -1 Then grdAging2.TextMatrix(j, idx2_Still) = _
                CStr(val(grdAging2.TextMatrix(j, idx2_TransNet)))
        End If
    Next j

    '-------------------------------------------
    ' 4) „”Õ/Õ›Ÿ À„ «·ÿ»«⁄…
    '-------------------------------------------
    On Error Resume Next
    Cn.Execute "DELETE FROM TblAging"
    On Error GoTo 0

    ' «Õ›Ÿ «·Ã—Ìœ «·√Ê· («·›Ê« Ì—) ›Ì TblAging
    ' ‰›” ‰œ«¡ﬂ «·ﬁœÌ„:  Ã«Â· StillAmount ›Ì saveGrid ⁄·‘«‰ ÂÊ »Ì Õ”»
    saveGrid "SELECT * FROM TblAging", grdAging, "StillAmount", "", "Credit_Or_Debit", 0

    ' «ÿ»⁄
    PrintAging 0
End Sub

'========================================
Public Sub Aging_PostProcessAndPrint()
    Dim i As Long, j As Long
    Dim mValue As Double, mValue2 As Double
    Dim mAccount_Code As String, mAccount_Code2 As String
    Dim mPayedValue As Double
    Dim mJ As Long

    txtTotalStill = ""
    mJ = 1

    For i = 1 To grdAging.rows - 1
        mValue = val(grdAging.TextMatrix(i, grdAging.ColIndex("TransNet")))
        mAccount_Code = Trim(grdAging.TextMatrix(i, grdAging.ColIndex("Account_Code")))

        If val(grdAging.TextMatrix(i, grdAging.ColIndex("PayedValue"))) <> mValue Then
            mJ = grdAging2.FindRow(mAccount_Code, grdAging2.FixedRows, grdAging2.ColIndex("Account_Code"), False, True)
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
                        grdAging.TextMatrix(i, grdAging.ColIndex("PayedValue")) = _
                            val(grdAging.TextMatrix(i, grdAging.ColIndex("PayedValue"))) + mValue2
                        grdAging2.TextMatrix(j, grdAging2.ColIndex("TransNet")) = 0
                    ElseIf mValue - mPayedValue < mValue2 Then
                        grdAging.TextMatrix(i, grdAging.ColIndex("PayedValue")) = mPayedValue + mValue - mPayedValue
                        grdAging2.TextMatrix(j, grdAging2.ColIndex("TransNet")) = mValue2 - (mValue - mPayedValue)
                        Dim c As Long
                        c = grdAging.ColIndex("TransNetGrid2")
                        If c <> -1 Then
                            grdAging.TextMatrix(i, c) = mValue2 - (mValue - mPayedValue)
                        End If

                        
                    End If
                End If

'                grdAging2.TextMatrix(j, grdAging2.ColIndex("StillAmount")) = _
'                    val(grdAging2.TextMatrix(j, grdAging2.ColIndex("TransNet"))) - _
'                    val(grdAging2.TextMatrix(j, grdAging2.ColIndex("PayedValue")))
            End If
        End If

        grdAging.TextMatrix(i, grdAging.ColIndex("StillAmount")) = _
            val(grdAging.TextMatrix(i, grdAging.ColIndex("TransNet"))) - _
            val(grdAging.TextMatrix(i, grdAging.ColIndex("PayedValue")))

        If val(grdAging.TextMatrix(i, grdAging.ColIndex("StillAmount"))) = 0 Then
            grdAging.TextMatrix(i, grdAging.ColIndex("StillAmount")) = ""
            grdAging.RowHidden(i) = True
        End If

        txtTotalStill = val(txtTotalStill) + val(grdAging.TextMatrix(i, grdAging.ColIndex("StillAmount")))
    Next

    ' ‰Õ›Ÿ “Ì “„«‰
    Cn.Execute "Delete TblAging"
    saveGrid "Select * from TblAging", grdAging, "StillAmount", "", "Credit_Or_Debit", 0

    ' ‰ÿ»⁄
    PrintAging 0
End Sub

Private Sub chkIsVat_Click()
GrdExcel.ColHidden(GrdExcel.ColIndex("VatValue")) = False
GrdExcel.ColHidden(GrdExcel.ColIndex("AmountNet")) = False
Dim Percentg As Double
Dim Notevalue As Double, mVat As Double, mValue As Double
If chkIsVat.value = vbUnchecked Then
    GrdExcel.TextMatrix(0, GrdExcel.ColIndex("AmountNet")) = "«·„»·€ «·«Ã„«·Ì"
Else
    GrdExcel.TextMatrix(0, GrdExcel.ColIndex("AmountNet")) = "«·„»·€ »œÊ‰ «·÷—Ì»…"
End If

PercentgValueAddedAccount_Transec XPDtbTrans7.value, 21, 1, , Percentg

Dim i As Long
For i = 1 To GrdExcel.rows - 1
    Notevalue = Abs(val(GrdExcel.TextMatrix(i, GrdExcel.ColIndex("Amount"))))
    If Notevalue <> 0 Then
        If chkIsVat.value = vbChecked Then
            If Percentg = 5 Then
                mValue = Notevalue / 1.05
            ElseIf Percentg = 15 Then
                mValue = Notevalue / 1.15
            End If
            mVat = Notevalue - mValue
            GrdExcel.TextMatrix(i, GrdExcel.ColIndex("VatValue")) = Round(mVat, 3)
 
        
             GrdExcel.TextMatrix(i, GrdExcel.ColIndex("AmountNet")) = Round(Notevalue - mVat, 3)
        Else
            If Percentg = 5 Then
                mValue = Notevalue / 1.05
            ElseIf Percentg = 15 Then
                mValue = Notevalue * 1.15
            End If
            mVat = mValue - Notevalue
            GrdExcel.TextMatrix(i, GrdExcel.ColIndex("VatValue")) = Round(mVat, 3)
        
            GrdExcel.TextMatrix(i, GrdExcel.ColIndex("AmountNet")) = Round(mValue, 3)
            
            
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

Private Sub cmdcreate_Click(Index As Integer)

If Index = 5 Then


        Dim mPay As Long
        Dim rsMPay As New ADODB.Recordset
        mPay = val(cmbPaymentType.BoundText)
        mPay = 6
        Dim mSerPos As Long
        Dim mSerPosString As String
        Dim mIsHiddenVat As Boolean
            If mPay <> 0 Then
            
                s = " SELECT"
                s = s & "        IsHiddenVat, TT = (CASE"
                s = s & "              WHEN bd.BankId > 9 THEN CAST(bd.BankId AS NVARCHAR)"
                s = s & "                     Else '0' + CAST(bd.BankId AS NVARCHAR)"
                s = s & "                 END)"
                s = s & "             From TblPaymentType"
                s = s & "             INNER JOIN BanksData bd"
                s = s & "                 ON bd.BankId = TblPaymentType.BankId"
                s = s & "             Where IsNull(IsNewCode, 0) = 1"
                s = s & " and PaymentID = " & mPay
                Set rsMPay = New ADODB.Recordset
                
                rsMPay.Open s, Cn, adOpenStatic, adLockReadOnly
                If Not rsMPay.EOF Then
                    mSerPos = val(rsMPay!tt & "")
                    mSerPosString = Trim(rsMPay!tt & "")
                    mIsHiddenVat = IIf(IsNull(rsMPay!IsHiddenVat & ""), False, rsMPay!IsHiddenVat & "")
                    
                End If
                rsMPay.Close
            End If
   
            Dim rsDummy As New ADODB.Recordset
            s = "Select * from TblHandWages where  PaymentId = " & mPay
            rsDummy.Open s, Cn, adOpenKeyset, adLockOptimistic
            Do While Not rsDummy.EOF
                rsDummy!NoteSerial1 = Voucher_coding(val(rsDummy!BranchID & ""), rsDummy!RecordDate, 81, 1100, , , , , , , "TblHandWages", , , mSerPosString)
                rsDummy.update
                rsDummy.MoveNext
            Loop
            
   
   
    
    

Exit Sub
End If


Dim Frm As New FrmDiscounts

If XPTxtID(Index) <> "" Then
    Frm.show
    Frm.Retrive val(XPTxtID(Index))
   
    Exit Sub
End If

If Index <> 4 Then

Frm.show
Frm.Cmd_Click (0)
Frm.CboDiscountType.ListIndex = 4
Frm.DBCboClientName = DcCustmer(10).BoundText
Frm.DcboDebitSide.BoundText = Trim(grdAcc(Index).TextMatrix(1, grdAcc(Index).ColIndex("Account_Code")))
If Index = 3 Then
    Frm.txtTotal = grdAcc(Index).TextMatrix(2, grdAcc(Index).ColIndex("Value"))
Else
    Frm.txtTotal = grdAcc(Index).TextMatrix(3, grdAcc(Index).ColIndex("Value"))
End If

Frm.mIsNoMsg = True
Frm.Cmd_Click (2)
XPTxtID(Index) = Frm.XPTxtID

s = "Update TblTamimi Set  XPTxtID" & Index + 1 & " = " & val(Frm.XPTxtID) & " Where Id = " & val(TxtSerial1(mIndex))
Cn.Execute s
cmdcreate(Index).Caption = "› Õ «·«‘⁄«—"

Unload Frm
Else
    FrmCashing.show
    FrmCashing.Cmd_Click 0
    FrmCashing.Option2 = True
    FrmCashing.DBCboClientName.BoundText = DcCustmer(10).BoundText
    FrmCashing.XPTxtVal = grdAcc(Index).TextMatrix(1, grdAcc(Index).ColIndex("Value"))
    
    
    'FrmCashing.TxtVAt2 = TxtVAt22
    'FrmCashing.txtTotal = txtTotalWithVat2
    'FrmCashing.DcboBox.BoundText = DcboBox.BoundText

End If
End Sub

Private Sub CmdCreateV7_Click()

'If (TxtNoteSerial7.Text) = "" Then
cmdDelNote7_Click
    If createVoucher7 Then
       'FindRec val(TXTLCNO.Text)
       
            s = "Update TblCaptinTrans Set NoteID = " & val(txtNoteID7) & ",NoteSerial = '" & Trim(TxtNoteSerial) & "' Where Id = " & val(TxtSerial1(mIndex))
            
                    
            Cn.Execute s
            
            FindRec val(TxtSerial1(mIndex).text)
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

Private Sub cmdCust_Click()
Dim mfrm As New FrmItemsClass
mfrm.mIndex = 11
mfrm.show
End Sub

Private Sub cmdDelNote7_Click()

Dim X As Integer
Dim Msg As String
Dim StrSQL As String
    
        X = vbYes

      If X = vbYes Then
        StrSQL = "Delete From DOUBLE_ENTREY_VOUCHERS Where Notes_ID=" & val(Me.txtNoteID7.text)
        Cn.Execute StrSQL, , adExecuteNoRecords
        StrSQL = "Delete From Notes Where NoteID=" & val(Me.txtNoteID7.text)
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

Private Sub cmdInsert_Click()
Dim s As String
Dim rsDummy As New ADODB.Recordset

    s = " Select * from ( SELECT    LblDiscountsTotal, dbo.Transactions.Transaction_Type, dbo.Transactions.Transaction_ID, dbo.Transactions.Transaction_Serial, dbo.Transactions.Transaction_Date,"
    s = s & "                      dbo.Transactions.CusID, dbo.TblCustemers.CusName, dbo.TblCustemers.CusNamee, dbo.TblCustemers.Cus_Phone, dbo.TblCustemers.Cus_mobile,"
    s = s & "                      dbo.TblCustemers.Fullcode, dbo.Transactions.BranchId, dbo.TblBranchesData.branch_name, dbo.TblBranchesData.branch_namee,"
    s = s & "                      dbo.Transactions.CashCustomerName, dbo.Transactions.CashCustomerPhone, dbo.Transactions.CashCustomerMobile, dbo.Transactions.CashCustomerAddress,"
    s = s & "                      dbo.Transactions.CashCustomerComment, dbo.Transactions.PaymentType, dbo.Transactions.NoteSerial1,       Transactions.Emp_id,"
    s = s & "                      TempName= ( SELECT     Emp_Name FROM         dbo.TblEmployee WHERE     (TblEmployee.Emp_ID =   transactions.Emp_id   ) )       ,"
    s = s & "                      (SELECT     SUM(Transaction_NetValue+isNull(0,0) ) AS SumValue                           FROM         dbo.Transactions AS A"
    s = s & "                      WHERE     (A.Transaction_Type = 9) AND (A.ReturnSerial = dbo.Transactions.NoteSerial1)) AS RetValue,"
    s = s & "                      (SELECT     SUM(dbo.TblNotesBillBuyPayment2.TransPayedValue) AS SumValue                           FROM         dbo.TblNotesBillBuyPayment2"
    s = s & "                      WHERE     (dbo.TblNotesBillBuyPayment2.NoteID = dbo.Transactions.Transaction_ID)) AS PayedVal, dbo.Transactions.Transaction_NetValue"
    s = s & "                      FROM         dbo.Transactions LEFT OUTER JOIN                    dbo.TblBranchesData ON dbo.Transactions.BranchId = dbo.TblBranchesData.branch_id"
    s = s & "                      LEFT OUTER JOIN                    dbo.TblCustemers ON dbo.Transactions.CusID = dbo.TblCustemers.CusID"
                        
    s = s & "                      Where (dbo.transactions.PaymentType = 1) And (dbo.transactions.Transaction_Type = 21 Or dbo.transactions.Transaction_Type = 9)"
                        
    
    
    
    If Not IsNull(FromDate.value) Then
       s = s & " and dbo.Transactions.Transaction_Date >=" & SQLDate(FromDate, True) & " "
    End If
    
    If Not IsNull(ToDate.value) Then
       s = s & " and dbo.Transactions.Transaction_Date <=" & SQLDate(ToDate, True) & " "
    End If
    s = s & " and Transactions.CusId = " & val(DcCustmer(10).BoundText)
    s = s & "   ) T"
    s = s & "                      Where (Round(IsNull(Transaction_NetValue, 0), 2) - Round(IsNull(retvalue, 0), 2) - Round(IsNull(PayedVal, 0), 2)) <> 0"
    s = s & "                      and     ( round( IsNull(Transaction_NetValue,0),2) -round(IsNull(retvalue,0),2) -  round(IsNull(PayedVal,0),2)) >1       ORDER by NoteSerial1"
    
    Set rsDummy = New ADODB.Recordset
    rsDummy.Open s, Cn, adOpenKeyset, adLockReadOnly
    grd(0).rows = 1
    grd(0).rows = grd(0).rows + 1
    Dim i As Long
    Dim mTotalNet As Double
    Dim mTotalDiscountNet As Double
    Dim mTransaction_NetValue As Double
    i = grd(0).rows - 1
    Do While Not rsDummy.EOF
        If val(rsDummy!Transaction_Type & "") = 9 Then
            grd(0).TextMatrix(i, grd(0).ColIndex("GROSSAMOUNT")) = val(rsDummy!Transaction_NetValue & "") * -1
            grd(0).TextMatrix(i, grd(0).ColIndex("DISCOUNT")) = val(rsDummy!LblDiscountsTotal & "") * -1
            grd(0).TextMatrix(i, grd(0).ColIndex("NETAMOUNT")) = (val(rsDummy!Transaction_NetValue & "") - val(rsDummy!LblDiscountsTotal & "")) * -1
        Else
            grd(0).TextMatrix(i, grd(0).ColIndex("GROSSAMOUNT")) = val(rsDummy!Transaction_NetValue & "")
            grd(0).TextMatrix(i, grd(0).ColIndex("DISCOUNT")) = val(rsDummy!LblDiscountsTotal & "")
            grd(0).TextMatrix(i, grd(0).ColIndex("NETAMOUNT")) = val(rsDummy!Transaction_NetValue & "") - val(rsDummy!LblDiscountsTotal & "")
        
        End If
        grd(0).TextMatrix(i, grd(0).ColIndex("INVOICENUMBER")) = Trim(rsDummy!NoteSerial1 & "")
        grd(0).TextMatrix(i, grd(0).ColIndex("RECEIVINGDATE")) = Trim(rsDummy!Transaction_Date & "")
        
        mTotalNet = mTotalNet + val(grd(0).TextMatrix(i, grd(0).ColIndex("GROSSAMOUNT")))
        mTotalDiscountNet = mTotalDiscountNet + val(grd(0).TextMatrix(i, grd(0).ColIndex("DISCOUNT")))
        
        mTransaction_NetValue = mTransaction_NetValue + val(grd(0).TextMatrix(i, grd(0).ColIndex("GROSSAMOUNT")))
        i = i + 1
        grd(0).rows = grd(0).rows + 1
        rsDummy.MoveNext
    Loop
    
    grd(1).rows = 1
    grd(1).rows = 10
     grd(1).TextMatrix(1, grd(1).ColIndex("TypeN")) = 1
    grd(1).TextMatrix(1, grd(1).ColIndex("GROSSAMOUNT")) = mTotalNet
    grd(1).cell(flexcpBackColor, 1, 1, 1, grd(1).Cols - 1) = vbRed
    
    
calcValues
End Sub

Private Sub cmdPrintNote7_Click()
ShowGL_cc Me.TxtNoteSerial7.text, , 23001
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
    ElseIf Me.Option3.value = True Then
        Indxx = 57
       
        FrmSelectVendor.Indxx = Indxx
        Load FrmSelectVendor
        FrmSelectVendor.Indxx = Indxx
        FrmSelectVendor.show
        FrmSelectVendor.Indxx = Indxx
       ElseIf Me.Option4.value = True Then
        Indxx = 56
       
        FrmSelectVendor.Indxx = Indxx
        Load FrmSelectVendor
        FrmSelectVendor.Indxx = Indxx
        FrmSelectVendor.show
        FrmSelectVendor.Indxx = Indxx
    Else
        Indxx = 4
         FrmSelectVendor.mEmpId = 0
        If DcbEmployee.text <> "" And val(DcbEmployee.BoundText) <> 0 Then
            
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
    FrmSelectEmployee.lblflag.Caption = 2
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
    For i = 1 To GrdItems.rows - 1
    
        tRs.AddNew
        ii = ii + 1
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
        tRs!fullcode = mNewCode
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
grdSphCYL.rows = 1
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

grdSphCYL.rows = 1


j = 1
i = grdSphCYL.rows
Do While Not rsDummy.EOF
    grdSphCYL.rows = i + 1
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

GrdItems.rows = 1
Dim n As Long
n = 1
For i = 1 To grdSphCYL.rows - 1
    For j = 0 To grdSphCYL.Cols - 1
        If j >= 3 Then
            GrdItems.rows = GrdItems.rows + 1
            GrdItems.TextMatrix(n, 0) = n
            GrdItems.TextMatrix(n, GrdItems.ColIndex("GroupID")) = cmbGroupId.BoundText
            GrdItems.TextMatrix(n, GrdItems.ColIndex("UnitID")) = cmbUnitID.BoundText
            GrdItems.TextMatrix(n, GrdItems.ColIndex("GroupName")) = cmbGroupId.text
            GrdItems.TextMatrix(n, GrdItems.ColIndex("UnitName")) = cmbUnitID.text
            
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
    
    For i = 1 To GrdItems.rows - 1
    
        
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
tmpGrd.rows = 1

Dim i As Long

    GrdExcel.ColHidden(GrdExcel.ColIndex("VatValue")) = True
    GrdExcel.ColHidden(GrdExcel.ColIndex("AmountNet")) = True
    
    GrdExcel.rows = 1
    FromExcel GrdExcel, tmpGrd, Me, , , txtFile.text, "TblEmployee"
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
txtFile.text = CD1.FileName
End Sub


Private Sub Command7_Click()
FrmAnalysItems.mIndex = 4
FrmAnalysItems.show
End Sub

Private Sub grd_AfterEdit(Index As Integer, ByVal Row As Long, ByVal Col As Long)
With grd(1)
Select Case .ColKey(Col)
   
    Case "TypeN"
        Select Case val(.TextMatrix(Row, .ColIndex("TypeN")))
        Case 1
            .cell(flexcpBackColor, Row, 1, Row, .Cols - 1) = vbRed
        Case 2
            .cell(flexcpBackColor, Row, 1, Row, .Cols - 1) = vbBlue
        Case 3
            .cell(flexcpBackColor, Row, 1, Row, .Cols - 1) = vbYellow
        
        End Select
    Case "GROSSAMOUNT"
         .TextMatrix(Row, .ColIndex("NETAMOUNT")) = .TextMatrix(Row, .ColIndex("GROSSAMOUNT"))
         calcValues
End Select
End With
End Sub
Private Sub calcValues()
Dim i As Long
Dim mType As Long

Dim mValue1 As Double
Dim mValue2 As Double
Dim mValue3 As Double
Dim mValue4 As Double
   
Dim mTotalNet2 As Double
   Dim mTotalNet As Double
    Dim mTotalDiscountNet As Double
    Dim mTransaction_NetValue As Double
    Dim mTransaction_NetValue2 As Double
    
    With grd(0)
    For i = 1 To .rows - 1
        
        mTotalNet = mTotalNet + val(grd(0).TextMatrix(i, grd(0).ColIndex("GROSSAMOUNT")))
        mTotalDiscountNet = mTotalDiscountNet + val(grd(0).TextMatrix(i, grd(0).ColIndex("DISCOUNT")))
        
        mTransaction_NetValue = mTransaction_NetValue + val(grd(0).TextMatrix(i, grd(0).ColIndex("NETAMOUNT")))
        
        mValue2 = mValue2 + val(.TextMatrix(i, .ColIndex("GROSSAMOUNT")))

        
    Next
End With

mValue2 = 0
mValue1 = 0
mValue3 = 0
With grd(1)
    For i = 1 To .rows - 1
        mType = val(.TextMatrix(i, .ColIndex("TypeN")))
        If mType = 1 Then
        ElseIf mType = 2 Then
            mValue2 = mValue2 + val(.TextMatrix(i, .ColIndex("GROSSAMOUNT")))
        ElseIf mType = 3 Then
            mValue3 = mValue3 + val(.TextMatrix(i, .ColIndex("GROSSAMOUNT")))
        End If
        mTotalNet2 = mTotalNet2 + val(.TextMatrix(i, .ColIndex("GROSSAMOUNT")))
        mTransaction_NetValue2 = mTransaction_NetValue2 + val(.TextMatrix(i, .ColIndex("NETAMOUNT")))
    Next
    
End With
mTotalNet2 = mTotalNet2 + mTotalNet
mTransaction_NetValue2 = mTransaction_NetValue + mTransaction_NetValue2
grd(1).ColComboList(grd(1).ColIndex("TypeN")) = "#1;Œ’„ Rebate |#2;Œ’„ «· ”ÊÌﬁ|#3;Œ’„ «·»—Ê„Ê‘‰|"


    txtNetSalesAfter(0) = mTotalNet
    grdAcc(0).TextMatrix(3, grdAcc(0).ColIndex("Value")) = Round(mTotalNet * val(grdAcc(0).TextMatrix(1, grdAcc(0).ColIndex("Percent"))) / 100, 2)
    grdAcc(0).TextMatrix(1, grdAcc(0).ColIndex("Value")) = Round((mTotalNet * val(grdAcc(0).TextMatrix(1, grdAcc(0).ColIndex("Percent"))) / 100) / 1.15, 2)
    grdAcc(0).TextMatrix(2, grdAcc(0).ColIndex("Value")) = Round(val(grdAcc(0).TextMatrix(3, grdAcc(0).ColIndex("Value"))) - val(grdAcc(0).TextMatrix(1, grdAcc(0).ColIndex("Value"))), 2)

    grdAcc(1).TextMatrix(3, grdAcc(1).ColIndex("Value")) = Round(mValue2, 2)
    grdAcc(1).TextMatrix(2, grdAcc(1).ColIndex("Value")) = Round((mValue2 * val(grdAcc(1).TextMatrix(2, grdAcc(1).ColIndex("Percent"))) / 100), 2)
    grdAcc(1).TextMatrix(1, grdAcc(1).ColIndex("Value")) = Round(mValue2, 2) - grdAcc(1).TextMatrix(2, grdAcc(1).ColIndex("Value"))
    
    grdAcc(2).TextMatrix(3, grdAcc(2).ColIndex("Value")) = Round(mValue3, 2)
    grdAcc(2).TextMatrix(2, grdAcc(2).ColIndex("Value")) = Round((mValue3 * val(grdAcc(2).TextMatrix(2, grdAcc(2).ColIndex("Percent"))) / 100), 2)
    grdAcc(2).TextMatrix(1, grdAcc(2).ColIndex("Value")) = Round(mValue3, 2) - grdAcc(2).TextMatrix(2, grdAcc(2).ColIndex("Value"))
    

    

    grdAcc(3).TextMatrix(1, grdAcc(3).ColIndex("Value")) = Round(mTotalDiscountNet * val(grdAcc(3).TextMatrix(1, grdAcc(3).ColIndex("Percent"))) / 100, 2)
    grdAcc(3).TextMatrix(2, grdAcc(3).ColIndex("Value")) = Round(mTotalDiscountNet * val(grdAcc(3).TextMatrix(1, grdAcc(3).ColIndex("Percent"))) / 100, 2)
    
    grdAcc(4).TextMatrix(1, grdAcc(4).ColIndex("Value")) = mTransaction_NetValue2
    grdAcc(4).TextMatrix(2, grdAcc(4).ColIndex("Value")) = mTransaction_NetValue2

txtNetSalesAfter(3) = mTransaction_NetValue2
txtNetSalesAfter(2) = mTotalDiscountNet
txtNetSalesAfter(1) = mTotalNet2
    

txtNetSalesAfter(4) = Round(val(grdAcc(0).TextMatrix(3, grdAcc(0).ColIndex("Value"))) + val(grdAcc(1).TextMatrix(3, grdAcc(1).ColIndex("Value"))) + val(grdAcc(2).TextMatrix(3, grdAcc(2).ColIndex("Value"))) + val(grdAcc(3).TextMatrix(2, grdAcc(3).ColIndex("Value"))), 2)
Dim mNETAMOUNT
With grd(0)
For i = 1 To grd(0).rows - 1
    
     mNETAMOUNT = val(.TextMatrix(i, .ColIndex("NETAMOUNT")))
     If mTransaction_NetValue2 * val(txtNetSalesAfter(4)) <> 0 Then
        grd(0).TextMatrix(i, grd(0).ColIndex("DISTRIBUTION")) = Round(mNETAMOUNT / mTransaction_NetValue2 * val(txtNetSalesAfter(4)), 2)
    End If
Next
End With
With grd(1)
For i = 1 To .rows - 1
    
     mNETAMOUNT = val(.TextMatrix(i, .ColIndex("NETAMOUNT")))
     If mTransaction_NetValue2 * val(txtNetSalesAfter(4)) <> 0 Then
     .TextMatrix(i, .ColIndex("DISTRIBUTION")) = mNETAMOUNT / mTransaction_NetValue2 * val(txtNetSalesAfter(4))
     End If
Next
End With


End Sub
Private Sub grd_CellButtonClick(Index As Integer, ByVal Row As Long, ByVal Col As Long)
With grd(1)
Select Case .ColKey(Col)
   
    Case "TypeN"
        Select Case val(.TextMatrix(Row, .ColIndex("TypeN")))
        Case 1
            .cell(flexcpBackColor, Row, 1, Row, .Cols - 1) = vbRed
        Case 2
            .cell(flexcpBackColor, Row, 1, Row, .Cols - 1) = vbBlue
        Case 3
            .cell(flexcpBackColor, Row, 1, Row, .Cols - 1) = vbYellow
    End Select
End Select
End With
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
    FindRec val(Me.Grid2.TextMatrix(Me.Grid2.Row, Me.Grid2.ColIndex("id")))
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
If val(TxtNoteSerial.text) = 0 Then
If createVoucher2 Then
       'FindRec val(TXTLCNO.Text)
       
            s = "Update TblHandWages Set NoteID = " & val(TxtNoteID) & ",NoteSerial = '" & Trim(TxtNoteSerial) & "' Where Id = " & val(TxtSerial1(mIndex))
            
                    
            Cn.Execute s
            
            FindRec val(TxtSerial1(mIndex).text)
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
des = "    Õ”«» «·" & TxtNoteSerial.text


Dim tablename As String
Dim Filedname As String
Dim NoteSerial1 As String
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
    CreateNotes NoteID, NoteDate, BranchID, notytype, Notevalue, NoteSerial, (NoteSerial1), tablename, Filedname, val(TxtSerial1(mIndex)), des                    ', recordDateH.value
                                              TxtNoteID.text = NoteID
                                                     TxtNoteSerial.text = NoteSerial

    If Not CREATE_VOUCHER_GE2(val(TxtNoteID.text), BranchID, val(DCboUserName(mIndex).BoundText), NoteDate) Then createVoucher2 = False Else createVoucher2 = True
    RsSavRec.Resync adAffectCurrent

    updateNotesValueAndNobytext val(TxtNoteSerial.text), Format(txtNet.text, "###.00")
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
des = "    Õ”«» «·" & TxtNoteSerial7.text


Dim tablename As String
Dim Filedname As String
Dim NoteSerial1 As Long
Dim BranchID As Integer
Dim sql As String
Dim mRate  As Double
tablename = "TblCaptinTrans"

Filedname = "ID"
NoteSerial1 = val(TxtNoteSerial17)

BranchID = val(Dcbranch(mIndex).BoundText)
mRate = 1

'

Dim i As Long
Notevalue = 0
For i = 1 To GrdExcel.rows - 1
    Notevalue = Notevalue + Abs(val(GrdExcel.TextMatrix(i, GrdExcel.ColIndex("Amount"))))
Next

notytype = 23001


'mAccNO = val(DboParentAccount.BoundText)
NoteDate = (XPDtbTrans7.value)
 
If Notevalue > 0 Then
    CreateNotes NoteID, NoteDate, BranchID, notytype, Notevalue, NoteSerial, (NoteSerial1), tablename, Filedname, val(TxtSerial1(mIndex)), des                   ', recordDateH.value
                                              txtNoteID7.text = NoteID
                                                     TxtNoteSerial7.text = NoteSerial

    If Not CREATE_VOUCHER_GECaptin(val(txtNoteID7.text), BranchID, val(DCboUserName(mIndex).BoundText), NoteDate) Then createVoucher7 = False Else createVoucher7 = True
    RsSavRec.Resync adAffectCurrent

    updateNotesValueAndNobytext val(TxtNoteSerial7.text), Format(Notevalue, "###.00")
'
'
'    StrSQL = "update  " & tablename & "   set NoteID=" & NoteID & ",NoteSerial='" & NoteSerial & "'"

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
    Msg = "    Õ”«» " & TxtSerial1(mIndex).text
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
    

   
    Notevalue = val(txtNet.text)
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


        Dim mPay As Long
        Dim rsMPay As New ADODB.Recordset
        mPay = val(cmbPaymentType.BoundText)
        Dim mSerPos As Long
        Dim mSerPosString As String
        Dim mIsHiddenVat As Boolean
            If mPay <> 0 Then
            
                s = " SELECT"
                s = s & "        IsHiddenVat, TT = (CASE"
                s = s & "              WHEN bd.BankId > 9 THEN CAST(bd.BankId AS NVARCHAR)"
                s = s & "                     Else '0' + CAST(bd.BankId AS NVARCHAR)"
                s = s & "                 END)"
                s = s & "             From TblPaymentType"
                s = s & "             INNER JOIN BanksData bd"
                s = s & "                 ON bd.BankId = TblPaymentType.BankId"
                s = s & "             Where IsNull(IsNewCode, 0) = 1"
                s = s & " and PaymentID = " & mPay
                Set rsMPay = New ADODB.Recordset
                
                rsMPay.Open s, Cn, adOpenStatic, adLockReadOnly
                If Not rsMPay.EOF Then
                    mSerPos = val(rsMPay!tt & "")
                    mSerPosString = Trim(rsMPay!tt & "")
                    mIsHiddenVat = IIf(IsNull(rsMPay!IsHiddenVat & ""), False, rsMPay!IsHiddenVat & "")
                    
                End If
                rsMPay.Close
            End If
   
   
   
    
    ' «·«ÿ—«›
    
     ' «·Ï Õ”«» «Ì—«œ«  «·Õ«ÊÌ« 
         
    Notevalue = val(txtTotal.text)
    If Notevalue > 0 Then
                If mIsHiddenVat Then
                    StrAccountCodeCridet = get_account_code_branch(2, my_branch)
                    Msg = Msg & " Õ”«» «Ã„«·Ï «·„»Ì⁄«  "
                Else
                
                    StrAccountCodeCridet = get_account_code_branch(77, my_branch)
                    Msg = Msg & " Õ”«» «Ì—«œ«  «·’Ì«‰… "
                End If
        
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

        
        
 
        
        If ModAccounts.AddNewDev(LngDevID, line_no, StrAccountCodeCridet, Notevalue, 1, Msg, val(notes_id), , , , XPDtbTrans.value, val(DCboUserName(mIndex).BoundText), , , , , , CLng(mRate), , , setfoxy_Line, , , , , , , , , _
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
        StrSQL = "Delete From DOUBLE_ENTREY_VOUCHERS Where Notes_ID=" & val(Me.TxtNoteID.text)
        Cn.Execute StrSQL, , adExecuteNoRecords
        StrSQL = "Delete From Notes Where NoteID=" & val(Me.TxtNoteID.text)
        Cn.Execute StrSQL, , adExecuteNoRecords
        Cn.Execute " Update ContainerContracts set NoteID=null ,NoteSerial=null where ID=" & val(TxtSerial1(mIndex).text)
       
        
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
    Dim X As Integer
   Dim AccountVATCreit As String
    Dim rs As New ADODB.Recordset
    Dim notes_serial As String
    Dim notes_id As String
    Msg = "    Õ”«» " & TxtSerial1(mIndex).text
    notes_id = general_noteid
    my_branch = val(Dcbranch(mIndex))
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
    For i = 1 To GrdExcel.rows - 1
            mCustName = Trim(GrdExcel.TextMatrix(i, GrdExcel.ColIndex("CompanyName")))
            Notevalue = val(GrdExcel.TextMatrix(i, GrdExcel.ColIndex("Amount")))
            If IsDate(GrdExcel.TextMatrix(i, GrdExcel.ColIndex("DateEntry"))) Then
                DateEntry = CDate(GrdExcel.TextMatrix(i, GrdExcel.ColIndex("DateEntry")))
            Else
                DateEntry = XPDtbTrans7.value
            End If
            If Notevalue < 0 Then
               StrAccountCodeDebt = Trim(GrdExcel.TextMatrix(i, GrdExcel.ColIndex("Account_Code")))
            Else
                'StrAccountCodeDebt = ModAccounts.GetMyAccountCode("TblCustemers", "CusID", val(Me.DcCustmer(mIndex).BoundText))
                StrAccountCodeDebt = get_bank_Account(val(Me.DcboBankName.BoundText), "Account_Code")
            End If
            
            mDisc = Trim(GrdExcel.TextMatrix(i, GrdExcel.ColIndex("CompanyName"))) & " " & Trim(GrdExcel.TextMatrix(i, GrdExcel.ColIndex("OperationName"))) & " " & Trim(GrdExcel.TextMatrix(i, GrdExcel.ColIndex("DateEntry")))
        
            Notevalue = Abs(Notevalue)
            PercentgValueAddedAccount_Transec XPDtbTrans7.value, 21, 1, AccountVATCreit, Percentg

          
            If Notevalue <> 0 Then
                mVat = val(GrdExcel.TextMatrix(i, GrdExcel.ColIndex("VatValue")))
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
                    Notevalue = Abs(val(GrdExcel.TextMatrix(i, GrdExcel.ColIndex("Amount"))))
                Else
                   Notevalue = val(GrdExcel.TextMatrix(i, GrdExcel.ColIndex("AmountNet")))
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

ShowGL_cc Me.TxtNoteSerial.text, , 1100

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
    
    If Me.TxtModFlg2(mIndex).text = "N" Or Me.TxtModFlg2(mIndex).text = "E" Then
        
        If TXTOrDer_no(0).text <> "" Then
            TXTOrDer_no(0).text = ""
            TXTOrDer_no(1).text = ""
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
                     StrSQL = StrSQL & " Where (t.Transaction_Type = 21) And (t.order_no = '" & val(TXTOrDer_no(0).text) & "')"
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
        sql = sql & "  Where (dbo.tblordermaintenancetypes.OrderID = " & val(TXTOrDer_no(1).text) & ") And (dbo.tblordermaintenancetypes.TypeTrans = 2)"
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
        StrSQL = StrSQL & " Where (t.Transaction_Type = 21) And (t.order_no = '" & val(TXTOrDer_no(0).text) & "')"
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
sql = sql & "  where ID =" & val(TXTOrDer_no(1).text) & ""
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
    If Me.TxtModFlg2(mIndex).text <> "R" Then
    
            FG.rows = 1
            FG.rows = 2

    
        
    End If
ElseIf mIndex = 7 Then
    If Me.TxtModFlg2(mIndex).text <> "R" Then
    
            GrdExcel.rows = 1
            GrdExcel.rows = 2

    
        
    End If

End If

End Sub

Private Sub Cmd_DeleteRow_Click(Index As Integer)
If Me.TxtModFlg.text <> "R" Then

    

    RemoveGridRow




End If
End Sub
Private Sub RemoveGridRow()
    If mIndex = 1 Then
        With Me.FG
    'MsgBox .Row
            If .Row <= 0 Then
                    .rows = 2
            Exit Sub
            Else
            .RemoveItem .Row
            End If
        End With
    ElseIf mIndex = 7 Then
        With Me.GrdExcel
    'MsgBox .Row
            If .Row <= 0 Then
                    .rows = 2
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
    If TxtSerial1(mIndex).text <> "" Then
        '    If CheckDelCountry(Val(Me.TxtVac_ID.text)) = False Then
        '        Msg = "·«Ì„ﬂ‰ Õ–› Â–« «·”Ã·...!!!"
        '        MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
        '        Exit Sub
        '    End If
        If SystemOptions.UserInterface = ArabicInterface Then
        MSGType = MsgBox("Â·  —€» ›Ì Õ–› Â–« «·”Ã·", vbQuestion + vbMsgBoxRtlReading + vbYesNo + vbMsgBoxRight, App.Title)
        Else
        MSGType = MsgBox("Confirm Delete", vbQuestion + vbMsgBoxRtlReading + vbYesNo + vbMsgBoxRight, App.Title)
        End If

        If MSGType = vbYes Then
            RsSavRec.Find "id=" & val(TxtSerial1(mIndex).text), , adSearchForward, 1
           ' CuurentLogdata ("D")
            RsSavRec.delete
           
            If mIndex = 1 Then
                s = " Delete From TblHandWages2 Where MasterID = " & val(TxtSerial1(mIndex).text)
                Cn.Execute s
            End If
            
            
            If SystemOptions.UserInterface = ArabicInterface Then
            MsgBox " „  ⁄„·Ì… «·Õ–› »‰Ã«Õ", vbInformation + vbMsgBoxRtlReading + vbOKOnly + vbMsgBoxRight, App.Title
            Else
            MsgBox "Delete Successfully", vbInformation + vbMsgBoxRtlReading + vbOKOnly + vbMsgBoxRight, App.Title
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
            MsgBox StrMSG, vbOKOnly + vbMsgBoxRight, App.Title
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

        If zatcaStatus = 1 Then
                    Msg = "·« Ì„ﬂ‰  ⁄œÌ· «Ê Õ–› «Ì „” ‰œ Ì„ﬂ‰ﬂ ⁄„· „” ‰œ ⁄ﬂ”Ì ›ﬁÿ"
                        Msg = Msg & CHR(13) & ""
                        MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
                    Exit Sub
            End If
  
 

      Set rsDummy = New ADODB.Recordset
      s = "select * from TblCardAuthorizationReform where WorkOrder = " & val(TXTOrDer_no(0).text) & " "
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
    If TxtSerial1(mIndex).text <> "" Then
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
    If TxtSerial1(mIndex).text <> "" Then
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
            MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading + vbOKOnly, App.Title
        
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
    TxtModFlg2(mIndex).text = "N"
    If mIndex = 1 Then
        My_SQL = "ContainerContracts"
        DCboUserName(mIndex).BoundText = user_id
        Dcbranch(1).BoundText = branch_id
            
        FG.rows = 1
        FG.rows = 2
       DefaultInvoicetype.ListIndex = SystemOptions.DefaultInvoicetype
            zatcaStatus = 0
        '    FlgAproved = 0
            txtDateRec.value = Date
   ElseIf mIndex = 2 Then
        My_SQL = "TblOffice"
   ElseIf mIndex = 3 Then
        My_SQL = "TblLensesTypes"
        
  ElseIf mIndex = 7 Then
        My_SQL = "TblCaptinTrans"
        DCboUserName(mIndex).BoundText = user_id
        Dcbranch(mIndex).BoundText = branch_id
        GrdExcel.rows = 1
        chkIsDiscountOnly.value = vbChecked
        chkIsAddOnly.value = vbChecked
    
   ElseIf mIndex = 10 Then
        My_SQL = "TblTamimi"
        
        DcCustmer(mIndex).BoundText = 0
        DcCustmer_Click 10, 1
        grd(0).rows = 1
        grd(1).rows = 1
        grd(1).rows = 10
                Dim rr As Long
       For i = 0 To grdAcc.count - 1
            For j = 1 To grdAcc(i).Cols - 1
                For rr = 1 To grdAcc(i).rows - 1
                    grdAcc(i).TextMatrix(rr, j) = ""
                 Next rr
            Next
        Next
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
MySQL = MySQL & "'" & CBoBasedON.text & "' as CBoBasedON,  "
MySQL = MySQL & "'" & DcbCarType.text & "' as DcbCarType,  "
MySQL = MySQL & "'" & DcbyearFactor.text & "' as DcbyearFactor,  "
MySQL = MySQL & "'" & DCEquipments.text & "' as DCEquipments,  "
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
        
MySQL = MySQL & "  And (TT.ID =" & val(TxtSerial1(mIndex).text) & ")"
   
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
        MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
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
    xReport.ApplicationName = App.Title
    xReport.ReportAuthor = App.Title
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

      If mZakamsg <> "" Then
            
        MsgBox mZakamsg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, " ÂÌ∆… «·“ﬂ«… Ê«·÷—Ì»… Ê«·Ã„«—ﬂ «·„—Õ·… «·À«‰Ì… - „—Õ·… «·—»ÿ Ê«· ﬂ«„·"
    End If
    
    For Each CtrlTxt In Me.Controls

        If TypeOf CtrlTxt Is TextBox Or TypeOf CtrlTxt Is ComboBox Then
            If CtrlTxt.text = "" And CtrlTxt.Tag <> "" And CtrlTxt.Enabled = True Then
                MsgBox CtrlTxt.Tag, vbOKOnly + vbMsgBoxRight, App.Title
                CtrlTxt.SetFocus
                Exit Sub
            End If
        End If

    Next

If mIndex < 2 Then
    If SystemOptions.UserInterface = ArabicInterface Then
        If TxtCashCustomerName.text = "" Then
            MsgBox "Ì—ÃÏ «œŒ«· «·⁄„Ì·"
            DcCustmer(mIndex).SetFocus
            Exit Sub
        End If
    Else
        If DcCustmer(mIndex).text = "" Then
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
            

            StrSQL = "SELECT NoteSerial1,OrDer_no2   FROM TblHandWages where  IsNull(OrDer_no,0)  = '" & val(TXTOrDer_no(0).text) & "' and Id <> " & val(TxtSerial1(mIndex).text)
            rs2.Open StrSQL, Cn, adOpenForwardOnly, adLockReadOnly
            If Not rs2.EOF Then
                MsgBox "Â–« «·«„— ·« Ì„ﬂ‰ «œ—«ÃÂ ›ﬁœ «œ—Ã „‰ ﬁ»· ›Ï «·›« Ê—… —ﬁ„" & rs2!NoteSerial1 & ""
                TXTOrDer_no2 = ""
                TXTOrDer_no(0) = ""
                TXTOrDer_no(1).text = ""
                
               ' Cmd(2).Enabled = True
              
                Exit Sub
            End If
        End If
   
             
            If checkCustomerdata(val(Me.DcCustmer(1).BoundText), val(txtTotalNet), val(DefaultInvoicetype.ListIndex), Dccurrency.text, Export) = False Then Exit Sub
        

ElseIf mIndex = 7 Then
    
        If DcboBankName.text = "" Then
            MsgBox "Ì—ÃÏ «œŒ«· «·»‰ﬂ"
            DcboBankName.SetFocus
            Exit Sub
        End If
 
        If Dcbranch(mIndex).text = "" Then
            MsgBox "Ì—ÃÏ «œŒ«· «·›—⁄"
            Dcbranch(mIndex).SetFocus
            Exit Sub
        End If
End If


   
    '------------------------------ check if Empcode exist ----------------------

   

    ' -------------------------------------- txtmodflg type -------------------
    Select Case TxtModFlg2(mIndex).text

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
            
                 
        If TxtNoteSerial17.text = "" Then
                If Voucher_coding(val(Dcbranch(mIndex).BoundText), XPDtbTrans7.value, 7, 23001, , , , , , , "TblCaptinTrans") = "error" Then
                    If SystemOptions.UserInterface = ArabicInterface Then
                        MsgBox " ·« Ì„ﬂ‰ «÷«›… ”‰œ  ÃœÌœ ·«‰ﬂ  ⁄œÌ  «·Õœ «·–Ì ﬁ„  » ÕœÌœ… „‰ «·”‰œ«   ": Exit Sub
                    Else
                        MsgBox " Cant't Create Expenses Voucher to this Process no You exceed the maximum number ": Exit Sub
                    End If

                Else
         
                    If Voucher_coding(val(Dcbranch(mIndex).BoundText), XPDtbTrans7.value, 7, 23001, , , , , , , "TblCaptinTrans") = "" Then
                        If SystemOptions.UserInterface = ArabicInterface Then
                            
                            TxtNoteSerial17.locked = False
                            MsgBox " ·«»œ „‰ ﬂ «»… —ﬁ„ «·”‰œ ÌœÊÌ« ﬂ„« Õœœ   ": Exit Sub
                        Else
                            MsgBox "  Enter Voucher No Manually or Define Coding ": Exit Sub
                        End If

                    Else
                        TxtNoteSerial17.text = Voucher_coding(val(Dcbranch(mIndex).BoundText), XPDtbTrans7.value, 7, 23001, , , , , , , "TblCaptinTrans")
                    End If
                End If
            End If
                
              '  AddNewRec
               FiLLRec7
                  ElseIf mIndex = 10 Then
                  FiLLRec10
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
          ElseIf mIndex = 10 Then
                FiLLRec10
          
            End If
    End Select

    Exit Sub
ErrTrap:
If SystemOptions.UserInterface = ArabicInterface Then
    MsgBox "Â‰«ﬂ Œÿ√ „« ›Ì ≈œŒ«· «·»Ì«‰« ", vbOKOnly + vbMsgBoxRight, App.Title
 Else
  MsgBox "Sorry...error douring insert data", vbOKOnly + vbMsgBoxRight, App.Title
End If
 
End Sub

Private Sub Btn_Undo_Click(Index As Integer)
    Undo
End Sub
Private Sub Undo()
    On Error GoTo ErrTrap

    Select Case TxtModFlg2(mIndex).text

        Case "N"
            clear_all Me
            TxtModFlg2(mIndex).text = "R"
           
            btn_First_Click (mIndex)
        Case "E"
            RsSavRec.Find "ID='" & val(TxtSerial1(mIndex).text) & "'", , adSearchForward, adBookmarkFirst

            If RsSavRec.EOF Or RsSavRec.BOF Then
                TxtModFlg2(mIndex).text = "R"
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
            ElseIf mIndex = 10 Then
                FiLLTXT10
        
                
            End If
            TxtModFlg2(mIndex).text = "R"
    End Select

    Exit Sub
ErrTrap:
End Sub
  Private Sub GetCardAuthorizationData()
   DcbyearFactor.text = ""
            TxtPlatNo = ""
            DcbCarType.BoundText = ""
            
            TxtManualNo2(2).text = """"
             TxtManualNo2(1).text = ""
  If val(TXTOrDer_no(1)) <> 0 Then
  
        Dim rs2 As New ADODB.Recordset
        Dim orderStatus As Integer
        Dim StrSQL As String
        MintDone = 0
    
        Set rs2 = New ADODB.Recordset
        StrSQL = "select * from TblCardAuthorizationReform where WorkOrder = " & val(TXTOrDer_no(0).text) & " "
        rs2.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText
        If rs2.RecordCount > 0 Then
            orderStatus = IIf(IsNull(rs2("OrderStatus").value), 0, rs2("OrderStatus").value)
            TxtCashCustomerName.text = IIf(IsNull(rs2("ClientName").value), "", rs2("ClientName").value)
            'DCOPrType =
                  
                  
                  
            
            DcbyearFactor.text = val(rs2!YearFact & "")
            TxtPlatNo = Trim(rs2!PlateNo & "")
            DcbCarType.BoundText = val(rs2!CarTypeID & "")
            DcbCarModel.BoundText = IIf(IsNull(rs2("CarModelID").value), "", rs2("CarModelID").value)
            TxtManualNo2(2).text = Trim(rs2!Shaseh & "")
             TxtManualNo2(1).text = Trim(rs2!CarMeter & "")
               DcCustmer(mIndex).BoundText = val(rs2!CusID & "")
                If val(rs2!CusID & "") = 0 Then
                    StrSQL = "SELECT tc.CusID FROM TblCustemers AS tc WHERE tc.CusName LIKE N'%" & Trim(TxtCashCustomerName.text) & "%'"
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
                     StrSQL = StrSQL & " Where (t.Transaction_Type = 21) And (t.NoteSerial1 = '" & val(TXTOrDer_no(0).text) & "')"
                          
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
    TxtSerial1(mIndex).text = IIf(IsNull(RsSavRec.Fields("id").value), "", RsSavRec.Fields("id").value)
    XPDtbTrans.value = IIf(IsNull(RsSavRec("RecordDate").value), Date, RsSavRec("RecordDate").value)
    Me.TxtNoteSerial1.text = IIf(IsNull(RsSavRec("NoteSerial1").value), "", RsSavRec("NoteSerial1").value)
   txtDiscValue = IIf(IsNull(RsSavRec("DiscValue").value), "", RsSavRec("DiscValue").value)
    txtDiscPercent = IIf(IsNull(RsSavRec("DiscPercent").value), "", RsSavRec("DiscPercent").value)
    Me.DcCustmer(1).BoundText = IIf(IsNull(RsSavRec("CusId").value), "", RsSavRec("CusId").value)
    Me.cmbPaymentType.BoundText = IIf(IsNull(RsSavRec("PaymentId").value), "", RsSavRec("PaymentId").value)
    
    txtTotal2 = IIf(IsNull(RsSavRec("Total2").value), "", RsSavRec("Total2").value)
    txtVat2 = IIf(IsNull(RsSavRec("Vat2").value), "", RsSavRec("Vat2").value)
    txtVatYou = IIf(IsNull(RsSavRec("VatYou").value), "", RsSavRec("VatYou").value)
    txtNet = IIf(IsNull(RsSavRec("Net").value), "", RsSavRec("Net").value)
    
     txtGeneralTotal = val(RsSavRec!GeneralTotal & "")
     txtTotalDisc = val(RsSavRec!TotalDisc & "")
     txtTotalBVat = val(RsSavRec!TotalBVat & "")
     txtTotalVat = val(RsSavRec!TotalVat & "")
     txtTotalNet = val(RsSavRec!TotalNet & "")
    
    
   CBoBasedON.ListIndex = IIf(IsNull(RsSavRec("CBoBasedON").value), -1, RsSavRec("CBoBasedON").value)
    TXTOrDer_no(0) = IIf(IsNull(RsSavRec("OrDer_no").value), "", RsSavRec("OrDer_no").value)
    TXTOrDer_no(1) = IIf(IsNull(RsSavRec("OrDer_no2").value), "", RsSavRec("OrDer_no2").value)
    
    TXTOrDer_no2 = IIf(IsNull(RsSavRec("RowsEstimatedID").value), "", RsSavRec("RowsEstimatedID").value)
    

   
    Dcbranch(mIndex).BoundText = IIf(IsNull(RsSavRec("BranchID").value), "", RsSavRec("BranchID").value)
    txtRemarks = IIf(IsNull(RsSavRec("Remarks").value), "", RsSavRec("Remarks").value)
    Me.DCboUserName(1).BoundText = IIf(IsNull(RsSavRec("UserID").value), "", RsSavRec("UserID").value)

    
    DCDocTypes.BoundText = IIf(IsNull(RsSavRec("Doctype").value), "", RsSavRec("Doctype").value)
    Me.Dccurrency.BoundText = IIf(IsNull(RsSavRec("Currency_id").value), "", RsSavRec("Currency_id").value)
    txt_Currency_rate.text = IIf(IsNull(RsSavRec("Currency_rate").value), 1, (RsSavRec("Currency_rate").value))
    txtDateRec.value = IIf(IsNull(RsSavRec("DateRec").value), Date, (RsSavRec("DateRec").value))
    zatcaStatus = IIf(IsNull(RsSavRec("zatcaStatus").value), 0, RsSavRec("zatcaStatus").value)
    TXTIban.text = IIf(IsNull(RsSavRec("CIBAN").value), "", (RsSavRec("CIBAN").value))
    
    DefaultInvoicetype.ListIndex = IIf(IsNull(RsSavRec("Invoicetype").value), 0, RsSavRec("Invoicetype").value)
     
    
 
     
    
    
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
    
    loadgrid s, FG, True, True
CalcTotal2



       

ErrTrap:

End Sub



Public Sub FiLLTXT10(Optional Lngid As Long = 0)

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
    TxtSerial1(mIndex).text = IIf(IsNull(RsSavRec.Fields("id").value), "", RsSavRec.Fields("id").value)
    XPDtbTrans10.value = IIf(IsNull(RsSavRec("RecordDate").value), Date, RsSavRec("RecordDate").value)
    
    FromDate.value = IIf(IsNull(RsSavRec("FromDate").value), Date, RsSavRec("FromDate").value)
    ToDate.value = IIf(IsNull(RsSavRec("toDate").value), Date, RsSavRec("toDate").value)
    
     
   Me.DcCustmer(mIndex).BoundText = IIf(IsNull(RsSavRec("CusId").value), "", RsSavRec("CusId").value)
    
    
    For i = 0 To txtNetSalesAfter.count - 1
          txtNetSalesAfter(i) = RsSavRec("NetSalesAfter" & i + 1).value & ""
          
    Next
    
    
    For i = 0 To XPTxtID.count - 1
          XPTxtID(i) = RsSavRec("XPTxtID" & i + 1).value & ""
          If Trim(XPTxtID(i)) <> "" Then
            cmdcreate(i).Caption = "› Õ «·«‘⁄«—"
          End If
    Next
     '*********************
         

    
    
      
   s = "Select * from TblTamimi2 Where MasterID = " & val(TxtSerial1(mIndex)) & " and TypeN2 = 1"
   loadgrid s, grd(0), True, False
    
    
   
   s = "Select * from TblTamimi2 Where MasterID = " & val(TxtSerial1(mIndex)) & " and TypeN2 = 2"
   loadgrid s, grd(1), True, False
    
   
   For i = 0 To grdAcc.count - 1
        s = "Select * from TblTamimi3 Where MasterID = " & val(TxtSerial1(mIndex)) & " and TypeN2 = " & i
        loadgrid s, grdAcc(i), True, False
   
   Next
                
    
  '  s = " Delete From  Where MasterID = " & val(TxtSerial1(mIndex).Text)
  
 
    
        
    
    
 '   LabCurr_Rec(mIndex).Caption = RsSavRec.AbsolutePosition
 '   LabCount_Rec(mIndex).Caption = RsSavRec.RecordCount

    
    
    





    loadgrid s, FG, True, True
calcValues
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
    TxtSerial1(mIndex).text = IIf(IsNull(RsSavRec.Fields("id").value), "", RsSavRec.Fields("id").value)
    XPDtbTrans7.value = IIf(IsNull(RsSavRec("RecordDate").value), Date, RsSavRec("RecordDate").value)
    Me.TxtNoteSerial17.text = IIf(IsNull(RsSavRec("NoteSerial1").value), "", RsSavRec("NoteSerial1").value)
    

   
    Dcbranch(mIndex).BoundText = IIf(IsNull(RsSavRec("BranchID").value), "", RsSavRec("BranchID").value)
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
    
    loadgrid s, GrdExcel, True, True
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
    TxtSerial1(mIndex).text = IIf(IsNull(RsSavRec.Fields("id").value), "", RsSavRec.Fields("id").value)
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
    TxtName(mIndex).text = IIf(IsNull(RsSavRec.Fields("name").value), "", RsSavRec.Fields("name").value)
    TxtNameE(mIndex).text = IIf(IsNull(RsSavRec.Fields("nameE").value), "", RsSavRec.Fields("nameE").value)
    
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

        For i = 1 To .rows - 1

            If Trim(TxtSerial1(mIndex).text) = .TextMatrix(i, .ColIndex("id")) Then
                TxtSerial1(mIndex).text = .TextMatrix(i, .ColIndex("Ser"))
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
    .IsSubtotal(.rows - 1) = True
    Dim SngTotal As Single
    If .rows > 1 Then
        txtTotal2 = .Aggregate(flexSTSum, .FixedRows, .ColIndex("Price"), .rows - 1, .ColIndex("Price"))
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



Public Sub CreateIssueVoucher(ByVal Row As Long, ByVal mGrid As VSFlexGrid)

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
    Dim mUserId As Long
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
        mUserId = user_id
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
               
getItemCostData Date, RSTransDetails("Item_ID").value, val(mStoreId), val(Transaction_ID), OldQty, OldCost, NewQty, NewCost, , LngUnitID
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
            getItemCostData Date, RSTransDetails("Item_ID").value, val(mStoreId), val(Transaction_ID), OldQty, OldCost, NewQty, NewCost, , LngUnitID
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



Public Sub CreatePurchOrder(ByVal Row As Long, ByVal mGrid As VSFlexGrid)

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
    Dim mUserId As Long
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
        mUserId = user_id
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
               
getItemCostData Date, RSTransDetails("Item_ID").value, val(mStoreId), val(Transaction_ID), OldQty, OldCost, NewQty, NewCost, , LngUnitID
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
        TxtNoteSerial1.text = ""
        TxtNoteSerial.text = ""
   End If
   End If
End Sub

Private Sub Dcbranch_Click(Index As Integer, Area As Integer)
    If Index <> 8 Then
    If Me.TxtModFlg2(mIndex) <> "R" Then
    TxtNoteSerial1.text = ""
   TxtNoteSerial.text = ""
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
    
If Me.TxtModFlg2(mIndex).text <> "R" Then
If val(DcCustmer(Index).BoundText) <> 0 Then

    If mIndex = 10 Then
       ' Dim j As Long
        Dim rr As Long
       For i = 0 To grdAcc.count - 1
            For j = 1 To grdAcc(i).Cols - 1
                For rr = 1 To grdAcc(i).rows - 1
                    grdAcc(i).TextMatrix(rr, j) = ""
                 Next rr
            Next
        Next
        Dim Percetage As Double
        Dim AccountVATCreit As String
        Dim AccountVATCreitName As String
        PercentgValueAddedAccount_Transec XPDtbTrans10.value, 21, 1, AccountVATCreit, Percetage
        Dim rsAcc As New ADODB.Recordset
        Set rsAcc = New ADODB.Recordset
        s = "Select Account_Name from accounts Where Account_code = N'" & Trim(AccountVATCreit) & "'"
        rsAcc.Open s, Cn, adOpenStatic, adLockReadOnly
        If Not rsAcc.EOF Then
           AccountVATCreitName = Trim(rsAcc!account_name & "")
        End If
        
        s = " SELECT        TblCustomerContract.Percent1 , TblCustomerContract.Percent2 , TblCustomerContract.Percent3 ,"
        s = s & "           TblCustomerContract.Percent4 , TblCustomerContract.AccCode1 , TblCustomerContract.AccCode2 ,"
        s = s & "                         TblCustomerContract.AccCode3 , TblCustomerContract.AccCode4 , TblCustemers.CusName, TblCustomerContract.IsLastMonth,"
        s = s & "                         ACCOUNTS_2.Account_Name Account_Name2, ACCOUNTS_1.Account_Name AS Account_Name1, ACCOUNTS_3.Account_Name AS Account_Name3 , ACCOUNTS_4.Account_Name AS Account_Name4,"
        s = s & "                         ACCOUNTS.Account_Name as CusAccName,ACCOUNTS.Account_Code as CusAcc"
        s = s & " FROM            TblCustemers INNER JOIN"
        s = s & "                          ACCOUNTS AS ACCOUNTS_4 INNER JOIN"
        s = s & "                          ACCOUNTS AS ACCOUNTS_3 INNER JOIN"
        s = s & "                          TblCustomerContract INNER JOIN"
        s = s & "                          ACCOUNTS AS ACCOUNTS_1 ON TblCustomerContract.AccCode1 = ACCOUNTS_1.Account_Code INNER JOIN"
        s = s & "                          ACCOUNTS AS ACCOUNTS_2 ON TblCustomerContract.AccCode2 = ACCOUNTS_2.Account_Code ON ACCOUNTS_3.Account_Code = TblCustomerContract.AccCode3 ON"
        s = s & "                          ACCOUNTS_4.Account_Code = TblCustomerContract.AccCode4 ON TblCustemers.CusID = TblCustomerContract.CustomerId INNER JOIN"
        s = s & "                          ACCOUNTS ON TblCustemers.Account_Code = ACCOUNTS.Account_Code"
       
        s = s & "                         Where TblCustemers.CusId = " & val(DcCustmer(mIndex).BoundText)
        Dim rsDummy As New ADODB.Recordset
        rsDummy.Open s, Cn, adOpenKeyset, adLockReadOnly
        Do While Not rsDummy.EOF
            grdAcc(0).TextMatrix(1, grdAcc(0).ColIndex("AccountName")) = Trim(rsDummy!account_name1 & "")
            grdAcc(0).TextMatrix(1, grdAcc(0).ColIndex("Account_Code")) = Trim(rsDummy!AccCode1 & "")
            grdAcc(0).TextMatrix(1, grdAcc(0).ColIndex("Percent")) = Trim(rsDummy!percent1 & "")
        
            grdAcc(0).TextMatrix(2, grdAcc(0).ColIndex("AccountName")) = AccountVATCreitName
            grdAcc(0).TextMatrix(2, grdAcc(0).ColIndex("Account_Code")) = AccountVATCreit
            grdAcc(0).TextMatrix(2, grdAcc(0).ColIndex("Percent")) = Percetage
        
            grdAcc(0).TextMatrix(3, grdAcc(0).ColIndex("AccountName")) = Trim(rsDummy!CusAccName & "")
            grdAcc(0).TextMatrix(3, grdAcc(0).ColIndex("Account_Code")) = Trim(rsDummy!CusAcc & "")
            grdAcc(0).TextMatrix(3, grdAcc(0).ColIndex("Percent")) = 0
        
        
            grdAcc(1).TextMatrix(1, grdAcc(1).ColIndex("AccountName")) = Trim(rsDummy!account_name2 & "")
            grdAcc(1).TextMatrix(1, grdAcc(1).ColIndex("Account_Code")) = Trim(rsDummy!AccCode2 & "")
            grdAcc(1).TextMatrix(1, grdAcc(1).ColIndex("Percent")) = Trim(rsDummy!percent2 & "")
        
            grdAcc(1).TextMatrix(2, grdAcc(1).ColIndex("AccountName")) = AccountVATCreitName
            grdAcc(1).TextMatrix(2, grdAcc(1).ColIndex("Account_Code")) = AccountVATCreit
            grdAcc(1).TextMatrix(2, grdAcc(1).ColIndex("Percent")) = Percetage
        
            grdAcc(1).TextMatrix(3, grdAcc(1).ColIndex("AccountName")) = Trim(rsDummy!CusAccName & "")
            grdAcc(1).TextMatrix(3, grdAcc(1).ColIndex("Account_Code")) = Trim(rsDummy!CusAcc & "")
            grdAcc(1).TextMatrix(3, grdAcc(1).ColIndex("Percent")) = 0
 
        
        
            grdAcc(2).TextMatrix(1, grdAcc(2).ColIndex("AccountName")) = Trim(rsDummy!account_name3 & "")
            grdAcc(2).TextMatrix(1, grdAcc(2).ColIndex("Account_Code")) = Trim(rsDummy!AccCode3 & "")
            grdAcc(2).TextMatrix(1, grdAcc(2).ColIndex("Percent")) = Trim(rsDummy!percent3 & "")
        
            grdAcc(2).TextMatrix(2, grdAcc(2).ColIndex("AccountName")) = AccountVATCreitName
            grdAcc(2).TextMatrix(2, grdAcc(2).ColIndex("Account_Code")) = AccountVATCreit
            grdAcc(2).TextMatrix(2, grdAcc(2).ColIndex("Percent")) = Percetage
        
            grdAcc(2).TextMatrix(3, grdAcc(2).ColIndex("AccountName")) = Trim(rsDummy!CusAccName & "")
            grdAcc(2).TextMatrix(3, grdAcc(2).ColIndex("Account_Code")) = Trim(rsDummy!CusAcc & "")
            grdAcc(2).TextMatrix(3, grdAcc(2).ColIndex("Percent")) = 0

        
        
            grdAcc(3).TextMatrix(1, grdAcc(3).ColIndex("AccountName")) = Trim(rsDummy!account_name4 & "")
            grdAcc(3).TextMatrix(1, grdAcc(3).ColIndex("Account_Code")) = Trim(rsDummy!AccCode4 & "")
            grdAcc(3).TextMatrix(1, grdAcc(3).ColIndex("Percent")) = Trim(rsDummy!percent4 & "")
        
            grdAcc(3).TextMatrix(2, grdAcc(3).ColIndex("AccountName")) = Trim(rsDummy!CusAccName & "")
            grdAcc(3).TextMatrix(2, grdAcc(3).ColIndex("Account_Code")) = Trim(rsDummy!CusAcc & "")
            grdAcc(3).TextMatrix(2, grdAcc(3).ColIndex("Percent")) = 0
 
             
            grdAcc(4).TextMatrix(2, grdAcc(4).ColIndex("AccountName")) = Trim(rsDummy!CusAccName & "")
            grdAcc(4).TextMatrix(2, grdAcc(4).ColIndex("Account_Code")) = Trim(rsDummy!CusAcc & "")
            

            
        
            rsDummy.MoveNext
        Loop
        
        
        s = " SELECT        BanksData.BankName,BanksData.Account_Code, ACCOUNTS.Account_Name"
        s = s & " FROM            BanksData INNER JOIN"
        s = s & "          ACCOUNTS ON BanksData.Account_Code = ACCOUNTS.Account_Code INNER JOIN"
        s = s & " TblUsers ON BanksData.BankID = TblUsers.BankID"
        s = s & " Where TblUsers.UserId = " & user_id
        Dim rsOut As New ADODB.Recordset
        rsOut.Open s, Cn, adOpenStatic, adLockReadOnly
        If Not rsOut.EOF Then
        
        
            grdAcc(4).TextMatrix(1, grdAcc(4).ColIndex("AccountName")) = Trim(rsOut!account_name & "")
            grdAcc(4).TextMatrix(1, grdAcc(4).ColIndex("Account_Code")) = Trim(rsOut!Account_code & "")
       
            
        End If

        

    End If

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

Private Sub Option3_Click()
  If Me.Option3.value = True Then
        Reload 57
    End If
End Sub

Private Sub Option4_Click()
  If Me.Option4.value = True Then
        Reload 56
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
            If CBoBasedON.ListIndex = 1 Then TXTOrDer_no(0) = TXTOrDer_no2.text
            If CBoBasedON.ListIndex = -1 Then Exit Sub
            'Else
                If Index <> 1 Then
                TXTOrDer_no(1).text = TXTOrDer_no2.text
                
                End If
            'End If
            
            Dim StrSQL As String
            Dim orderStatus As Integer
     
            MintDone = 0
            Set rs2 = New ADODB.Recordset
            If CBoBasedON.ListIndex = 1 Then
                StrSQL = "select * from TblCardAuthorizationReform where WorkOrder = " & val(TXTOrDer_no(1).text) & " "
            Else
                StrSQL = "select * from TblRowsEstimated where ID =" & val(TXTOrDer_no2.text) & " "
                
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
                TxtCashCustomerName.text = IIf(IsNull(rs2("ClientName").value), "", rs2("ClientName").value)
                'DCOPrType =
                
                DcCustmer(mIndex).BoundText = val(rs2!CusID & "")
                If val(rs2!CusID & "") = 0 Then
                    StrSQL = "SELECT tc.CusID FROM TblCustemers AS tc WHERE tc.CusName LIKE N'%" & Trim(TxtCashCustomerName.text) & "%'"
                    Dim rsDummy As New ADODB.Recordset
                    rsDummy.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText
                    If Not rsDummy.EOF Then
                        DcCustmer(mIndex).BoundText = val(rsDummy!CusID & "")
                    End If
                End If
                
                 TXTOrDer_no(1) = val(rs2!ID & "")
                DcbyearFactor.text = val(rs2!YearFact & "")
                TxtPlatNo = Trim(rs2!PlateNo & "")
                DcbCarType.BoundText = val(rs2!CarTypeID & "")
                
                TxtManualNo2(2).text = Trim(rs2!Shaseh & "")
                 TxtManualNo2(1).text = Trim(rs2!CarMeter & "")
                
                DcbCarModel.BoundText = IIf(IsNull(rs2("CarModelID").value), "", rs2("CarModelID").value)
                 
                If orderStatus = 2 Or orderStatus = 4 Or orderStatus = 5 Then
                    MintDone = 1
                End If
                If Me.TxtModFlg2(mIndex) = "N" Or Me.TxtModFlg2(mIndex) = "E" Then
                    
                    Dim RsData3 As New ADODB.Recordset
                    
                    
                    s = "Select TblCardAuthorizationReformItems.qty, tblitems.itemid,TblCardAuthorizationReformItems.Price ,TblCardAuthorizationReformItems.TotalWithVat ,tblItems.ItemCode,tblItems.ItemName from TblCardAuthorizationReformItems Left Outer Join tblItems On tblItems.ItemID =TblCardAuthorizationReformItems.ItemID Left Outer join TblCardAuthorizationReform On TblCardAuthorizationReform.Id = TblCardAuthorizationReformItems.id"
                    
                    s = s & "  Where (dbo.TblCardAuthorizationReform.WorkOrder = " & val(TXTOrDer_no(0).text) & ") "
                           
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
                TxtCashCustomerName.text = ""
                MintDone = -1
                TXTOrDer_no(0) = ""
                TXTOrDer_no(1) = ""
                TXTOrDer_no2 = ""
                DcbCarType.text = ""
                DcbyearFactor.text = ""
                DCEquipments.text = ""
                TxtManualNo2(2) = ""
                TxtManualNo2(1) = ""
                TxtPlatNo = ""
                DcbCarModel.text = ""
                DcCustmer(1).text = ""
            End If
            
            CalcTotal2
'End If
End Sub

Private Sub TXTOrDer_no2_Validate(Cancel As Boolean)
TXTOrDer_no_Validate 0, False
End Sub

Private Sub txtPassword_Change()
If Trim(txtPassword) = "UpdateSerial" Then
    cmdcreate(5).Caption = "UpdateSerial"
    cmdcreate(5).Visible = True
'        txtFromDateReSave.Visible = True
'    txtToDateReSave.Visible = True
'    chkIsBranch(0).Visible = True
End If
End Sub

Private Sub TxtSearchCode_KeyPress(KeyAscii As Integer)
 Dim CUSTID As Integer

    If KeyAscii = vbKeyReturn Then
        GetCustomersDetail CUSTID, , TxtSearchCode.text, 1
        DBCboClientName.BoundText = CUSTID
    End If
End Sub

Private Sub VSFlexGrid1_AfterEdit(ByVal Row As Long, ByVal Col As Long)
Dim StrAccountCode As String
Dim LngRow As Long
With VSFlexGrid1

   Select Case .ColKey(Col)
        Case "UnitPrice", "ShowQty"
            FG.TextMatrix(Row, FG.ColIndex("Total")) = val(FG.TextMatrix(Row, FG.ColIndex("UnitPrice"))) * val(FG.TextMatrix(Row, FG.ColIndex("ShowQty")))
        Case "UnitName"
                StrAccountCode = .ComboData
            
                LngRow = .FindRow(StrAccountCode, .FixedRows, .ColIndex("UnitName"), False, True)
                .TextMatrix(Row, .ColIndex("UnitId")) = StrAccountCode
        
        Case "UnitId"
                StrAccountCode = .ComboData
            
                LngRow = .FindRow(StrAccountCode, .FixedRows, .ColIndex("UnitId"), False, True)
                .TextMatrix(Row, .ColIndex("UnitName")) = StrAccountCode
        Case "ItemName"
                StrAccountCode = .ComboData
            
                LngRow = .FindRow(StrAccountCode, .FixedRows, .ColIndex("ItemId"), False, True)
                .TextMatrix(Row, .ColIndex("ItemId")) = StrAccountCode
                
        Case "GroupName"
                StrAccountCode = .ComboData
            
                LngRow = .FindRow(StrAccountCode, .FixedRows, .ColIndex("GroupID"), False, True)
                .TextMatrix(Row, .ColIndex("GroupID")) = StrAccountCode
        Case "Discount"
            
            
        End Select
        
    End With

End Sub

Private Sub VSFlexGrid1_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
With VSFlexGrid1

If .ColKey(Col) <> "ItemID" Then .ComboList = ""
End With
End Sub

Private Sub VSFlexGrid1_StartEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
   Dim rs As New ADODB.Recordset
    Dim StrSQL  As String
    Dim StrAccountType As String
    Dim StrComboList As String
    Dim Msg As String
    Dim StrAccountCode As String
    Dim StrAccountCode1 As String

    Dim ClsAcc As New ClsAccounts
    Dim LngRow As Long

    With VSFlexGrid1

        Select Case .ColKey(Col)
 
            Case "ItemName"
             .TextMatrix(Row, .ColIndex("ItemName")) = ""
                StrSQL = "select ItemID,ItemName,ItemNamee from TblItems "
                rs.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

                If SystemOptions.UserInterface = ArabicInterface Then
                    StrComboList = FG.BuildComboList(rs, "ItemName", "ItemID")
                Else
                    StrComboList = FG.BuildComboList(rs, "ItemNamee", "ItemID")
                End If
       
                If StrComboList <> "" Then
                    StrComboList = "|" & StrComboList
                End If
                 .ComboList = StrComboList
         End Select
        End With
End Sub

Private Sub XPDtbTrans_Change()
    If Me.TxtModFlg2(mIndex) <> "R" Then
        TxtNoteSerial1.text = ""
        TxtNoteSerial.text = ""
   End If
       
    CalcTotal2
End Sub


Private Sub btn_First_Click(Index As Integer)
  On Error GoTo ErrTrap

    Dim Msg As String

  
    If Me.TxtModFlg2(mIndex).text = "N" Then
        FindRec val(TxtSerial1(mIndex).text)
        TxtModFlg2(mIndex).text = "R"
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
   ElseIf mIndex = 10 Then
        FiLLTXT10
        
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
            MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading + vbOKOnly, App.Title
            RsSavRec.Requery
            Resume BegnieWork
    End Select
End Sub
Private Sub btn_Next_Click(Index As Integer)
On Error GoTo ErrTrap
    Dim Msg As String

    If Me.TxtModFlg2(mIndex).text = "N" Then
        FindRec val(TxtSerial1(mIndex).text)
        TxtModFlg2(mIndex).text = "R"
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
      ElseIf mIndex = 10 Then
        FiLLTXT10
     
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
            MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading + vbOKOnly, App.Title
            RsSavRec.Requery
            Resume BegnieWork
    End Select
End Sub

Private Sub btn_Previous_Click(Index As Integer)
  On Error GoTo ErrTrap
    Dim Msg As String

    If TxtModFlg2(mIndex).text = "N" Then
        FindRec val(TxtSerial1(mIndex).text)
        TxtModFlg2(mIndex).text = "R"
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
        ElseIf mIndex = 10 Then
        FiLLTXT10
   
        
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
            MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading + vbOKOnly, App.Title
            RsSavRec.Requery
            Resume BegnieWork
    End Select
End Sub
Private Sub btn_Last_Click(Index As Integer)
  On Error GoTo ErrTrap

    Dim Msg As String

    If Me.TxtModFlg.text = "N" Then
        FindRec val(TxtSerial1(mIndex).text)
        Me.TxtModFlg2(mIndex).text = "R"
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
        
          ElseIf mIndex = 10 Then
        FiLLTXT10
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
            MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading + vbOKOnly, App.Title
            RsSavRec.Requery
            Resume BegnieWork
    End Select
End Sub

Private Sub TxtModFlg2_Change(Index As Integer)
 On Error GoTo ErrTrap

    Select Case Me.TxtModFlg2(mIndex).text

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
'BtnPrint22.Caption = "Print"



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
    If TxtVac_ID.text <> "" Then
    If SystemOptions.UserInterface = ArabicInterface Then
        MSGType = MsgBox("Â·  —€» ›Ì Õ–› Â–« «·”Ã·", vbQuestion + vbMsgBoxRtlReading + vbYesNo + vbMsgBoxRight, App.Title)
        Else
        MSGType = MsgBox("ÂConfirm Delete", vbQuestion + vbMsgBoxRtlReading + vbYesNo + vbMsgBoxRight, App.Title)
     End If

        If MSGType = vbYes Then
        Cn.Execute "Update TblUserScreen set FlgWork=null where id=" & val(Me.DcbScreen.BoundText) & ""
            RsSavRec.Find "id=" & val(TxtVac_ID.text), , adSearchForward, 1
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
        If SystemOptions.UserInterface = ArabicInterface Then
            StrMSG = "⁄›Ê« ·« ÌÃÊ“ Õ–› «·”Ã· ·«— »«ÿÂ »»Ì«‰«  √Œ—Ì"
            Else
            StrMSG = "Sorry... Can not Delete.  is related to with other data"
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
        If SystemOptions.UserInterface = ArabicInterface Then
            Msg = "⁄›Ê« ·ﬁœ  „ Õ–› Â–« «·”Ã· «· «·Ï" & CHR(13)
            Msg = Msg & "„‰ ﬁ»· „” Œœ„ √Œ— ⁄·Ï «·‘»ﬂ… " & CHR(13)
            Msg = Msg & "”Ê› Ì „  ÕœÌÀ «·»Ì«‰« "
            Else
            Msg = "Sorry I've been to delete this record" & CHR(13)
            Msg = Msg & "From Another user on network " & CHR(13)
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
        If SystemOptions.UserInterface = ArabicInterface Then
            Msg = "⁄›Ê« ·ﬁœ  „ Õ–› Â–« «·”Ã· «· «·Ï" & CHR(13)
            Msg = Msg & "„‰ ﬁ»· „” Œœ„ √Œ— ⁄·Ï «·‘»ﬂ… " & CHR(13)
            Msg = Msg & "”Ê› Ì „  ÕœÌÀ «·»Ì«‰« "
              Else
            Msg = "Sorry I've been to delete this record" & CHR(13)
            Msg = Msg & "From Another user on network " & CHR(13)
            Msg = Msg & "Data will be updated"
            End If
            MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading + vbOKOnly, App.Title
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

    If TxtVac_ID.text <> "" Then
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

'    If DoPremis(Do_New, Me.name, True) = False Then
'        Exit Sub
'    End If

    On Error GoTo ErrTrap
    Set rs = New ADODB.Recordset
   ' Frm2.Enabled = True
   
    '-----------------------------------
    Me.TxtVac_ID.text = ""
 
    Frame3.Enabled = False
    '-----------------------------------
    TxtModFlg.text = "N"
clear_all Me
FillGridWithData
 My_SQL = "select ID,Name From TblUserScreen WHERE     (FlgWork IS NULL)"
    fill_combo Me.DcbScreen, My_SQL
    My_SQL = "TblVisitScreen"
    rs.Open My_SQL, Cn, adOpenKeyset, adLockOptimistic, adCmdTable

    If rs.RecordCount > 0 Then
        TxtSerial.text = rs.RecordCount + 1
    Else
        TxtSerial.text = 1
    End If

    rs.Close
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
        If SystemOptions.UserInterface = ArabicInterface Then
            Msg = "⁄›Ê« ·ﬁœ  „ Õ–› Â–« «·”Ã· «· «·Ï" & CHR(13)
            Msg = Msg & "„‰ ﬁ»· „” Œœ„ √Œ— ⁄·Ï «·‘»ﬂ… " & CHR(13)
            Msg = Msg & "”Ê› Ì „  ÕœÌÀ «·»Ì«‰« "
              Else
            Msg = "Sorry I've been to delete this record" & CHR(13)
            Msg = Msg & "From Another user on network " & CHR(13)
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
        If SystemOptions.UserInterface = ArabicInterface Then
            Msg = "⁄›Ê« ·ﬁœ  „ Õ–› Â–« «·”Ã· «· «·Ï" & CHR(13)
            Msg = Msg & "„‰ ﬁ»· „” Œœ„ √Œ— ⁄·Ï «·‘»ﬂ… " & CHR(13)
            Msg = Msg & "”Ê› Ì „  ÕœÌÀ «·»Ì«‰« "
              Else
            Msg = "Sorry I've been to delete this record" & CHR(13)
            Msg = Msg & "From Another user on network " & CHR(13)
            Msg = Msg & "Data will be updated"
            End If
            MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading + vbOKOnly, App.Title
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
            If CtrlTxt.text = "" And CtrlTxt.Tag <> "" And CtrlTxt.Enabled = True Then
                MsgBox CtrlTxt.Tag, vbOKOnly + vbMsgBoxRight, App.Title
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
If TxtEmpRemark.text = "" Then
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
    Select Case Me.TxtModFlg.text

            '------------------------------ new record ----------------------------
        Case "N"
      
            '------------------------- save record -----------------------------
            AddNewRec
            BtnLast_Click

        Case "E"

        If TxtNoteSerial17.text = "" Then
                If Voucher_coding(val(Dcbranch(mIndex).BoundText), XPDtbTrans7.value, 7, 23001, , , , , , , "TblCaptinTrans") = "error" Then
                    If SystemOptions.UserInterface = ArabicInterface Then
                        MsgBox " ·« Ì„ﬂ‰ «÷«›… ”‰œ  ÃœÌœ ·«‰ﬂ  ⁄œÌ  «·Õœ «·–Ì ﬁ„  » ÕœÌœ… „‰ «·”‰œ«   ": Exit Sub
                    Else
                        MsgBox " Cant't Create Expenses Voucher to this Process no You exceed the maximum number ": Exit Sub
                    End If

                Else
         
                    If Voucher_coding(val(Dcbranch(mIndex).BoundText), XPDtbTrans7.value, 7, 23001, , , , , , , "TblCaptinTrans") = "" Then
                        If SystemOptions.UserInterface = ArabicInterface Then
                            
                            TxtNoteSerial17.locked = False
                            MsgBox " ·«»œ „‰ ﬂ «»… —ﬁ„ «·”‰œ ÌœÊÌ« ﬂ„« Õœœ   ": Exit Sub
                        Else
                            MsgBox "  Enter Voucher No Manually or Define Coding ": Exit Sub
                        End If

                    Else
                        TxtNoteSerial17.text = Voucher_coding(val(Dcbranch(mIndex).BoundText), XPDtbTrans7.value, 7, 23001, , , , , , , "TblCaptinTrans")
                    End If
                End If
            End If
            '----------------------------- save edit -------------------------------
            FiLLRec
    End Select

    Exit Sub
ErrTrap:
If SystemOptions.UserInterface = ArabicInterface Then
    MsgBox "Â‰«ﬂ Œÿ√ „« ›Ì ≈œŒ«· «·»Ì«‰« ", vbOKOnly + vbMsgBoxRight, App.Title
   Else
   MsgBox "Sorry...error in douring enter data", vbOKOnly + vbMsgBoxRight, App.Title
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
    MsgBox Msg, vbOKOnly + vbMsgBoxRight + vbInformation, App.Title
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
If TxtEmpRemark.text = "" Then
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
If TxtEmpRemark.text = "" Then
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
    
    If SystemOptions.IsBlue Then
        Command7.Visible = True
    Else
        Command7.Visible = False
    End If
Frame3.Enabled = False
Frame2(0).Enabled = False
Frame4.Enabled = False
    RecordDate = Date
'DTPicker1.value = Date
'DTPicker2.value = Date
'DTPicker3.value = Date
XPDtbTrans10.value = Date
    TabMain.TabVisible(0) = False
     TabMain.TabVisible(1) = False
     TabMain.TabVisible(2) = False
     TabMain.TabVisible(3) = False
    TabMain.TabVisible(4) = False
     TabMain.TabVisible(5) = False
     TabMain.TabVisible(6) = False
     TabMain.TabVisible(7) = False
     TabMain.TabVisible(8) = False
     TabMain.TabVisible(10) = False
     TabMain.TabVisible(9) = False
    
          
    If mIndex = 0 Then
        TabMain.TabVisible(0) = True
        TabMain.CurrTab = 0
    ElseIf mIndex = 1 Then
        Me.Dcbranch(1).BoundText = branch_id
        TabMain.TabVisible(1) = True
        TabMain.CurrTab = 1
        Me.Caption = "√ÃÊ— «·Ìœ"
              With Me.DefaultInvoicetype
            .Clear
            
             


            .AddItem " ›« Ê—… ÷—Ì»Ì…  "
            .ItemData(0) = 0
     
            .AddItem " ›« Ê—… ÷—Ì»Ì… „»”ÿ… "
            .ItemData(1) = 2
         
        End With
 My_SQL = " select id,code from currency"
 
    fill_combo Me.Dccurrency, My_SQL
    
     ElseIf mIndex = 10 Then
        
        TabMain.TabVisible(1) = True
        TabMain.CurrTab = 10
        Me.Caption = "Tamimi Payment"
        Me.Width = Me.grd(1).Width + 400
        
        Set Dcombos = New ClsDataCombos
            
    grd(1).ColComboList(grd(1).ColIndex("TypeN")) = "#1;Œ’„ Rebate |#2;Œ’„ «· ”ÊÌﬁ|#3;Œ’„ «·»—Ê„Ê‘‰|"
            Dcombos.GetCustomersSuppliers 1, DcCustmer(10), , , 1
    ElseIf mIndex = 2 Then
        Me.Width = Grid2.Width + 400
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
            ToDate.value = Date
            FromDate.value = Date
            ToDate1.value = Date
            FromDate1.value = Date
            ToDate.value = Date

            
    If SystemOptions.UserInterface = ArabicInterface Then
        My_SQL = "  select  ID,Name  from ClassCustomers  "
    Else
        My_SQL = "  select  ID,Namee  from ClassCustomers  "
    End If

    fill_combo dcClass, My_SQL


       
            
            
            FromDate.value = ""
            ToDate.value = ""
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
            Dcombos.GetBranches Me.Dcbranch(mIndex)
    

            Dcombos.GetCustomerType Me.DcCustomerType



            
            
  
    
    
    
    
    


  
    
    
    
    
   
    
    Resize_Form Me
         
     
    ElseIf mIndex = 7 Then
        Me.Width = GrdExcel.Width + 400
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
    Dcombos.GetPaymentType cmbPaymentType
    Dcombos.GetTblCarsDataGroup Me.DcbCarType
    Dcombos.GetBranches Me.Dcbranch(1)
    Dcombos.GetBranches Me.Dcbranch(7)
    
    
   ' Dcombos.GetTblCarModels Me.DcbCarModel
    
    If SystemOptions.UserInterface = EnglishInterface Then
        My_SQL = "SELECT id,ISNULL(ModelE,Model) ModelName from TblCarModels"
    Else
        My_SQL = "SELECT id, Model from TblCarModels"
    End If
    fill_combo DcbCarModel, My_SQL
      'Dim ii As Integer
     
      For ii = 1900 To 2100
        Me.DcbyearFactor.AddItem (ii)
      Next ii
      
'Me.Dcbranch(1).BoundText = branch_id
    Resize_Form Me
    
    
    SetDtpickerDate Me.XPDtbTrans
    
    If mIndex = 1 Then
        My_SQL = "Select * from TblHandWages where Year(RecordDate) = year(getdate() ) "
       ' Set BKGrndPic = New ClsBackGroundPic
        Set RsSavRec = New ADODB.Recordset
        RsSavRec.CursorLocation = adUseClient
        RsSavRec.Open My_SQL, Cn, adOpenKeyset, adLockPessimistic
        TxtModFlg2(mIndex).text = "R"
        DCboUserName(mIndex).BoundText = user_id
       

        

        btn_First_Click (mIndex)
        
    ElseIf mIndex = 7 Then
        My_SQL = "TblCaptinTrans"
       ' Set BKGrndPic = New ClsBackGroundPic
        Set RsSavRec = New ADODB.Recordset
        RsSavRec.CursorLocation = adUseClient
        RsSavRec.Open My_SQL, Cn, adOpenKeyset, adLockPessimistic, adCmdTableDirect
        TxtModFlg2(mIndex).text = "R"
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
        Me.TxtModFlg.text = "R"
        Resize_Form Me
        
        'load tblUsers -----------------------------------------------
    
    
        FillGridWithData
    
        With Me.Grid
    '        .Cell(flexcpPicture, 0, .ColIndex("ContractNo")) = Me.GrdImageList.ListImages("Vac_Name").ExtractIcon
    '        .Cell(flexcpPicture, 0, .ColIndex("Ser")) = Me.GrdImageList.ListImages("Ser").ExtractIcon
    
            For i = 0 To .Cols - 1
                .cell(flexcpPictureAlignment, 0, i) = flexPicAlignRightCenter
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
            cmbStoreID2.BoundText = val(rsDummy!StoreID & "")
            
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
    
        If mIndex = 10 Then
        My_SQL = "TblTamimi"
       ' Set BKGrndPic = New ClsBackGroundPic
        Set RsSavRec = New ADODB.Recordset
        RsSavRec.CursorLocation = adUseClient
        RsSavRec.Open My_SQL, Cn, adOpenKeyset, adLockPessimistic, adCmdTableDirect
        TxtModFlg2(mIndex).text = "R"
       ' DCboUserName(mIndex).BoundText = user_id
       

       

        btn_First_Click (mIndex)
        
        
        End If
    ShowTip
TabMain.CurrTab = mIndex
    If SystemOptions.UserInterface = EnglishInterface Then
        SetInterface Me
        ChangeLang
    End If
Dcombos.GetCustomersSuppliers 1, DcCustmer(9), , , 1
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


Sub GetItemBalanceInStore(Optional LngRow As Long, Optional ColorID As Integer, Optional itemsize As Integer, Optional ClassId As Integer, Optional StrItemSerial As String, Optional LngItemID As Long, Optional TransactionDate As Variant, Optional ByVal mStoreId As Long = 1, Optional ByRef mGrid As VSFlexGrid = Nothing)
   
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


Private Sub CheckQtyFromStore(mGrid As VSFlexGrid)
Dim StrSQL  As String
Dim mDateTrans As Date
Dim Begin  As Boolean
Dim mItemId As Long
Dim mStoreId As Long
Dim mQuantity As Double
mDateTrans = Date
Dim mItemBalance As Double
Dim i As Long
For i = 1 To mGrid.rows - 1
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
        mGrid.TextMatrix(i, mGrid.ColIndex("StoreIDAvi")) = rsDummy!StoreID
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
    If mIndex = 0 Or mIndex = 8 Then
        mSelectModFlg = Me.TxtModFlg.text
    Else
    
        mSelectModFlg = Me.TxtModFlg2(mIndex).text
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

        IntResult = MsgBox(StrMSG, vbMsgBoxRight + vbYesNoCancel + vbMsgBoxRtlReading + vbQuestion, App.Title)

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
        
       ElseIf mIndex = 10 Then
        StrRecID = new_id("TblTamimi", "id", "")
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
    ElseIf mIndex = 10 Then
        FiLLRec10


    End If
    
ErrTrap:

End Sub

Public Sub FiLLRec()
    On Error GoTo ErrTrap
     RsSavRec.Fields("UserID").value = IIf(DcbUserID.BoundText <> 0, val(DcbUserID.BoundText), Null)
     RsSavRec.Fields("EmpUsrID").value = IIf(DcbEmpUsrID.BoundText <> 0, val(DcbEmpUsrID.BoundText), Null)
     RsSavRec.Fields("ScreenID").value = IIf(DcbScreen.BoundText <> 0, val(DcbScreen.BoundText), Null)
     RsSavRec.Fields("UserPass").value = TxtUserPass.text
     RsSavRec.Fields("EmpPass").value = TxtEmpPass.text
     RsSavRec.Fields("RecordDate").value = RecordDate.value
     RsSavRec.Fields("CusRemark").value = TxtCusRemark.text
     RsSavRec.Fields("EmpRemark").value = TxtEmpRemark.text
     RsSavRec.update
     Cn.Execute "Update TblUserScreen set FlgWork=1 where id=" & val(Me.DcbScreen.BoundText) & ""
    If SystemOptions.UserInterface = ArabicInterface Then
    MsgBox " „  ⁄„·Ì… «·Õ›Ÿ »‰Ã«Õ", vbInformation + vbOKOnly + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
    Else
    MsgBox "Save Successfully", vbInformation + vbOKOnly + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
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
    Dim s As String
        Dim mPay As Long
        Dim rsMPay As New ADODB.Recordset
        mPay = val(cmbPaymentType.BoundText)
        Dim mSerPos As Long
        Dim mSerPosString As String
        Dim mIsHiddenVat As Boolean
            If mPay <> 0 Then
            
                s = " SELECT"
                s = s & "        IsHiddenVat, TT = (CASE"
                s = s & "              WHEN bd.BankId > 9 THEN CAST(bd.BankId AS NVARCHAR)"
                s = s & "                     Else '0' + CAST(bd.BankId AS NVARCHAR)"
                s = s & "                 END)"
                s = s & "             From TblPaymentType"
                s = s & "             INNER JOIN BanksData bd"
                s = s & "                 ON bd.BankId = TblPaymentType.BankId"
                s = s & "             Where IsNull(IsNewCode, 0) = 1"
                s = s & " and PaymentID = " & mPay
                Set rsMPay = New ADODB.Recordset
                
                rsMPay.Open s, Cn, adOpenStatic, adLockReadOnly
                If Not rsMPay.EOF Then
                    mSerPos = val(rsMPay!tt & "")
                    mSerPosString = Trim(rsMPay!tt & "")
                    mIsHiddenVat = IIf(IsNull(rsMPay!IsHiddenVat & ""), False, rsMPay!IsHiddenVat & "")
                    
                End If
                rsMPay.Close
            End If
   
   
   
   
        If TxtNoteSerial1.text = "" Or TxtNoteSerial1.text = "00" Then
                If Voucher_coding(val(Dcbranch(mIndex).BoundText), XPDtbTrans.value, 81, 1100, , , , , , , "TblHandWages", , , mSerPosString) = "error" Then
                    If SystemOptions.UserInterface = ArabicInterface Then
                        MsgBox " ·« Ì„ﬂ‰ «÷«›… ”‰œ  ÃœÌœ ·«‰ﬂ  ⁄œÌ  «·Õœ «·–Ì ﬁ„  » ÕœÌœ… „‰ «·”‰œ«   ": Exit Sub
                    Else
                        MsgBox " Cant't Create Expenses Voucher to this Process no You exceed the maximum number ": Exit Sub
                    End If

                Else
         
                    If Voucher_coding(val(Dcbranch(mIndex).BoundText), XPDtbTrans.value, 81, 1100, , , , , , , "TblHandWages", , , mSerPosString) = "" Then
                        If SystemOptions.UserInterface = ArabicInterface Then
                            
                            TxtNoteSerial1.locked = False
                            MsgBox " ·«»œ „‰ ﬂ «»… —ﬁ„ «·”‰œ ÌœÊÌ« ﬂ„« Õœœ   ": Exit Sub
                        Else
                            MsgBox "  Enter Voucher No Manually or Define Coding ": Exit Sub
                        End If

                    Else
                        TxtNoteSerial1.text = Voucher_coding(val(Dcbranch(mIndex).BoundText), XPDtbTrans.value, 81, 1100, , , , , , , "TblHandWages", , , mSerPosString)
                    End If
                End If
            End If
    
    
    If TxtModFlg2(mIndex).text = "N" Then
           ' TxtSerial1(mIndex).Text = CStr(new_id("TblAdditionsAssest", "ID", "", True))

       
        RsSavRec.AddNew
        TxtSerial1(mIndex).text = new_id("TblHandWages", "id", "")
        RsSavRec.Fields("ID").value = val(TxtSerial1(mIndex).text)
    End If
    RsSavRec("NoteSerial1").value = Trim$(Me.TxtNoteSerial1.text)
    RsSavRec.Fields("BranchID").value = IIf(Dcbranch(mIndex).text <> "", Trim(Dcbranch(mIndex).BoundText), Null)
    RsSavRec("RecordDate").value = XPDtbTrans.value
    
    RsSavRec("CBoBasedON").value = CBoBasedON.ListIndex
   DCboUserName(mIndex).BoundText = IIf(DCboUserName(mIndex).text <> "", Trim(DCboUserName(mIndex).BoundText), user_id)
   'RsSavRec("RecType").value = cmbRecType.ListIndex
    'RsSavRec("ContractNo").value = txtContractNo.Text
    'RsSavRec("RecName").value = txtRecName.Text
    'RsSavRec("RecordTime").value = XPDtbTransTime.Value
    
    RsSavRec("SerPos") = val(mSerPos)
    RsSavRec("IsHiddenVat") = mIsHiddenVat
    
    RsSavRec.Fields("OrDer_no").value = val(TXTOrDer_no(0).text)
    RsSavRec.Fields("OrDer_no2").value = val(TXTOrDer_no(1).text)
    RsSavRec.Fields("RowsEstimatedID").value = val(TXTOrDer_no2.text)
 
    
    RsSavRec.Fields("CusId").value = IIf(DcCustmer(1).text <> "", Trim(DcCustmer(1).BoundText), Null)
    RsSavRec.Fields("PaymentId").value = IIf(cmbPaymentType.text <> "", Trim(cmbPaymentType.BoundText), Null)
    
    RsSavRec.Fields("DiscValue").value = val(txtDiscValue.text)
    RsSavRec.Fields("Total2").value = val(txtTotal2.text)
    RsSavRec.Fields("VatYou").value = val(txtVatYou.text)
    RsSavRec.Fields("DiscPercent").value = val(txtDiscPercent.text)
    
    RsSavRec.Fields("Total").value = val(txtTotal.text)
    RsSavRec.Fields("Vat2").value = val(txtVat2.text)
    RsSavRec.Fields("Net").value = val(txtNet.text)
    
    
      RsSavRec!GeneralTotal = val(txtGeneralTotal)
     RsSavRec!TotalDisc = val(txtTotalDisc)
     RsSavRec!TotalBVat = val(txtTotalBVat)
     RsSavRec!TotalVat = val(txtTotalVat)
     RsSavRec!TotalNet = val(txtTotalNet)
    
    
    
    RsSavRec("Remarks").value = txtRemarks.text
    
    
    '*********************
     
    
    
      
   

    RsSavRec.update
    
    savenewelectroncic
    cmdDelNote_Click
    
                
    If mIndex = 1 Then
        s = " Delete From TblHandWages2 Where MasterID = " & val(TxtSerial1(mIndex).text)
    
        
        
    End If
    Cn.Execute s
    
    s = "Select * from TblHandWages2 Where Id = -1"
    'saveGrid s, fg, "Name", "ID", "MasterID", val(TxtSerial1(mIndex).Text)
    saveGrid s, FG, "Name", "", "MasterID", val(TxtSerial1(mIndex).text)
    
    CmdCreateV2_Click
'
If SystemOptions.UserInterface = ArabicInterface Then
    MsgBox " „  ⁄„·Ì… «·Õ›Ÿ »‰Ã«Õ", vbInformation + vbOKOnly + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
 Else
   MsgBox "Save Successfully", vbInformation + vbOKOnly + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
End If
    'CuurentLogdata
    
    TxtModFlg2(mIndex) = "R"

    Exit Sub
ErrTrap:

    If RsSavRec.EditMode <> adEditNone Then
        RsSavRec.CancelUpdate
    End If

End Sub






Public Sub FiLLRec10()
    On Error GoTo ErrTrap
    
   
    
    
    If TxtModFlg2(mIndex).text = "N" Then
           ' TxtSerial1(mIndex).Text = CStr(new_id("TblAdditionsAssest", "ID", "", True))

       
        RsSavRec.AddNew
        TxtSerial1(mIndex).text = new_id("TblTamimi", "id", "")
        RsSavRec.Fields("ID").value = val(TxtSerial1(mIndex).text)
    End If
    
   
    RsSavRec("RecordDate").value = XPDtbTrans10.value
    RsSavRec("CusId").value = DcCustmer(mIndex).BoundText
    RsSavRec("FromDate").value = FromDate.value
    RsSavRec("ToDate").value = ToDate.value
    
    Dim i As Long
    For i = 0 To txtNetSalesAfter.count - 1
          RsSavRec("NetSalesAfter" & i + 1).value = txtNetSalesAfter(i)
    Next
     '*********************
     
    For i = 0 To XPTxtID.count - 1
          RsSavRec("XPTxtID" & i + 1).value = XPTxtID(i)
    Next
     
    
    
      
   

    RsSavRec.update
   
    Dim s As String
                
    
    s = " Delete From TblTamimi2 Where MasterID = " & val(TxtSerial1(mIndex).text)
    Cn.Execute s
    s = " Delete From TblTamimi3 Where MasterID = " & val(TxtSerial1(mIndex).text)
    
        
        
   
    Cn.Execute s
    
    s = "Select * from TblTamimi2 Where Id = -1"
    'saveGrid s, fg, "Name", "ID", "MasterID", val(TxtSerial1(mIndex).Text)
    saveGrid s, grd(0), "INVOICENUMBER", "", "MasterID", val(TxtSerial1(mIndex).text), "TypeN2", 1
    saveGrid s, grd(1), "GROSSAMOUNT", "", "MasterID", val(TxtSerial1(mIndex).text), "TypeN2", 2
    
    s = "Select * from TblTamimi3 Where Id = -1"
    'saveGrid s, fg, "Name", "ID", "MasterID", val(TxtSerial1(mIndex).Text)
    For i = 0 To grdAcc.count - 1
        saveGrid s, grdAcc(i), "Value", "", "MasterID", val(TxtSerial1(mIndex).text), "TypeN2", i
    Next
    
'
If SystemOptions.UserInterface = ArabicInterface Then
    MsgBox " „  ⁄„·Ì… «·Õ›Ÿ »‰Ã«Õ", vbInformation + vbOKOnly + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
 Else
   MsgBox "Save Successfully", vbInformation + vbOKOnly + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
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
    
   
    
    
    If TxtModFlg2(mIndex).text = "N" Then
           ' TxtSerial1(mIndex).Text = CStr(new_id("TblAdditionsAssest", "ID", "", True))

       
        RsSavRec.AddNew
        TxtSerial1(mIndex).text = new_id("TblCaptinTrans", "id", "")
   '     RsSavRec.Fields("ID").value = val(TxtSerial1(mIndex).Text)
    End If
    RsSavRec("NoteSerial1").value = Trim$(Me.TxtNoteSerial17.text)
    RsSavRec.Fields("BranchID").value = IIf(Dcbranch(mIndex).text <> "", Trim(Dcbranch(mIndex).BoundText), Null)
    RsSavRec.Fields("BankID").value = IIf(DcboBankName.text <> "", Trim(DcboBankName.BoundText), Null)
    RsSavRec("RecordDate").value = XPDtbTrans7.value
    
       If chkIsVat.value = vbChecked Then
            RsSavRec.Fields("IsVat").value = 1
        Else
            RsSavRec.Fields("IsVat").value = 0
        End If
               
   RsSavRec.Fields("UserID").value = IIf(DCboUserName(mIndex).text <> "", Trim(DCboUserName(mIndex).BoundText), user_id)
   'RsSavRec("RecType").value = cmbRecType.ListIndex
    'RsSavRec("ContractNo").value = txtContractNo.Text
    'RsSavRec("RecName").value = txtRecName.Text
    'RsSavRec("RecordTime").value = XPDtbTransTime.Value
    

    
   ' RsSavRec("Remarks").value = TxtRemarks.Text
    
    
    '*********************
     
    
    
      
   

    RsSavRec.update
    'cmdDelNote7
    Dim s As String
                
    
        s = " Delete  TblCaptinTrans2 Where MasterID = " & val(TxtSerial1(mIndex).text)
    
        
        
   
    Cn.Execute s
    
    s = "Select  MasterID,Emp_ID,CompanyName,OperationName,EmpName,typename,Account_Name,DateEntry,Amount"
    s = s & " from TblCaptinTrans2 Where Id = -1"
    'saveGrid s, fg, "Name", "ID", "MasterID", val(TxtSerial1(mIndex).Text)
    saveGrid s, GrdExcel, "CompanyName", "", "MasterID", val(TxtSerial1(mIndex).text)
    
    CmdCreateV7_Click
'
If SystemOptions.UserInterface = ArabicInterface Then
    MsgBox " „  ⁄„·Ì… «·Õ›Ÿ »‰Ã«Õ", vbInformation + vbOKOnly + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
 Else
   MsgBox "Save Successfully", vbInformation + vbOKOnly + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
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
    TxtVac_ID.text = IIf(IsNull(RsSavRec.Fields("id").value), "", RsSavRec.Fields("id").value)
    TxtSerial.text = IIf(IsNull(RsSavRec.Fields("id").value), "", RsSavRec.Fields("id").value)
    Me.DcbUserID.BoundText = IIf(IsNull(RsSavRec.Fields("UserID").value), "", RsSavRec.Fields("UserID").value)
   ' Me.DcbEmpUsrID.BoundText = IIf(IsNull(RsSavRec.Fields("EmpUsrID").value), "", RsSavRec.Fields("EmpUsrID").value)
  '  TxtEmpPass.Text = IIf(IsNull(RsSavRec.Fields("EmpPass").value), "", RsSavRec.Fields("EmpPass").value)
  '  TxtUserPass.Text = IIf(IsNull(RsSavRec.Fields("UserPass").value), "", RsSavRec.Fields("UserPass").value)
    RecordDate.value = IIf(IsNull(RsSavRec.Fields("RecordDate").value), Date, RsSavRec.Fields("RecordDate").value)
    TxtCusRemark.text = IIf(IsNull(RsSavRec.Fields("CusRemark").value), "", RsSavRec.Fields("CusRemark").value)
    TxtEmpRemark.text = IIf(IsNull(RsSavRec.Fields("EmpRemark").value), "", RsSavRec.Fields("EmpRemark").value)
    Me.DcbScreen.BoundText = IIf(IsNull(RsSavRec.Fields("ScreenID").value), "", RsSavRec.Fields("ScreenID").value)


    LabCurrRec.Caption = RsSavRec.AbsolutePosition
    LabCountRec.Caption = RsSavRec.RecordCount

    With Grid

        For i = 1 To .rows - 1

            If Trim(TxtVac_ID.text) = .TextMatrix(i, .ColIndex("id")) Then
                TxtSerial.text = .TextMatrix(i, .ColIndex("Ser"))
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
       StrSQL = "Select * From TblUserComp Where  Password='" & Trim(Me.TxtEmpPass.text) & "'"
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
       StrSQL = "Select * From TblUsers Where UserID=" & Me.DcbUserID.BoundText & " AND PassWord='" & Trim(Me.TxtUserPass.text) & "'"
Rs3.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText
If Rs3.RecordCount > 0 Then
CheckPassworUser = True
Else
CheckPassworUser = False
End If
End Function


Function print_report(Optional NoteSerial As String, Optional Ind As Integer = 0)

        SaveQRCode "TblHandWages", "ID", val(TxtSerial1(mIndex).text), TxtNoteSerial1.text, (XPDtbTrans.value), _
        (txtTotalNet.text), Picture1, 0, (txtTotalVat.text), (txtTotalNet.text)

    

    
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
     'MySQL = " SELECT    distinct  '" & DcbCarModel.Text & "' as CarModel,TblHandWages.Remarks,TblHandWages.NoteSerial1,TblHandWages.Total2,TblHandWages.OrDer_no,TblHandWages.Total,TblHandWages.VatYou,TblHandWages.DiscValue,TblHandWages.Net,TblHandWages.Net,"
     MySQL = " SELECT    '" & DcbCarModel.text & "' as CarModel,TblHandWages.Remarks,TblHandWages.QrCodeImage, TblHandWages.NoteSerial1,TblHandWages.Total2,TblHandWages.OrDer_no,TblHandWages.Total,TblHandWages.VatYou,TblHandWages.DiscValue,TblHandWages.Net,TblHandWages.Net,"
     MySQL = MySQL & "                     TblHandWages2.Name ,TblHandWages.Remarks,TblHandWages.RecordDate as bill_date,"
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
     MySQL = MySQL & "                     dbo.TblCardAuthorizationReform.TypeCustomer, dbo.TblCardAuthorizationReform.OverKM, dbo.TblCustemers.CusName, TblCustemers.VATNO, dbo.TblCustemers.CusNamee,TblCustemers.*,"
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
        MySQL = MySQL & "  Where (TblHandWages.Id  =  " & val(TxtSerial1(1).text) & ") "
     'and (dbo.TblCardAuthorizationReformDetails.type=0)"

     ' RsDetails1.Open MySQL, Cn, adOpenStatic, adLockReadOnly, adCmdText
    
    
    
     
                
                    
        StrSQL = "SELECT     Transaction_Details.ShowPrice,dbo.Transactions.Transaction_ID, dbo.Transactions.Transaction_Date, dbo.Transactions.Transaction_Type, dbo.Transactions.NoteSerial1, "
        StrSQL = StrSQL & "                      dbo.Transactions.Transaction_Serial, dbo.Transactions.Transaction_HijriDate, dbo.Transactions.TransactionComment, dbo.Transactions.OpOrderID,"
        StrSQL = StrSQL & "                      dbo.Transactions.OldOpOrderID,Transactions.QrCodeImage, dbo.Transaction_Details.UnitId,dbo.Transaction_Details.OperPrice, dbo.Transaction_Details.ID, dbo.Transaction_Details.ShowQty, dbo.Transaction_Details.Item_ID,"
        StrSQL = StrSQL & "                      dbo.TblItems.itemname , dbo.TblItems.ItemNamee, dbo.TblItems.fullcode , dbo.Transaction_Details.showPrice"
        StrSQL = StrSQL & " ,ShowPrice2 = (SELECT Top 1 TblItemsUnits.UnitSalesPrice"
        StrSQL = StrSQL & "                 From TblItemsUnits"
        StrSQL = StrSQL & "                 Where ItemID = Transaction_Details.Item_ID"
        StrSQL = StrSQL & "                        AND UnitID           = Transaction_Details.UnitId  )"
        StrSQL = StrSQL & " FROM         dbo.TblItems RIGHT OUTER JOIN"
        StrSQL = StrSQL & "                      dbo.Transaction_Details ON dbo.TblItems.ItemID = dbo.Transaction_Details.Item_ID RIGHT OUTER JOIN"
        StrSQL = StrSQL & "                      dbo.Transactions ON dbo.Transaction_Details.Transaction_ID = dbo.Transactions.Transaction_ID"
        StrSQL = StrSQL & " Where (dbo.Transactions.Transaction_Type = 21) And  (Transactions.order_no = '" & val(TXTOrDer_no(0).text) & "')"
            
            
            
            Set RsData2 = New ADODB.Recordset
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
            MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
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
        xReport.ApplicationName = App.Title
        xReport.ReportAuthor = App.Title
        Set CViewer = New ClsReportViewer
        CViewer.FireReport xReport, WindowTarget, "", , , , StrFileName, , MySQL
    
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
sql = sql & " where dbo.TblVisitScreen.ID=" & val(TxtVac_ID.text) & " "
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
    FrmReport.CRViewer.viewReport
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
    TxtMod = TxtModFlg.text
    TxtModFlg.text = ""
    TxtModFlg = TxtMod
End Sub

Public Function FindRec(ByVal RecId As Long)
    
    'RsSavRec.Filter = adFilterNone
    
    Dim My_SQL
    
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
        
        ElseIf mIndex = 10 Then
            FiLLTXT10
             
        End If
    End If
    If RecId <> 0 And mIndex = 1 And RsSavRec.EOF Then
        
        My_SQL = "Select * from TblHandWages where Year(RecordDate) = year(getdate())  Or  id =   " & RecId
        ' Set BKGrndPic = New ClsBackGroundPic
         Set RsSavRec = New ADODB.Recordset
         RsSavRec.CursorLocation = adUseClient
         RsSavRec.Open My_SQL, Cn, adOpenKeyset, adLockPessimistic
         FiLLTXT1
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
    If TxtModFlg.text = "N" Then
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
    
    ElseIf TxtModFlg.text = "R" Then
     '   Frm2.Enabled = False
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
    
    ElseIf TxtModFlg.text = "E" Then
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
        .rows = 2
        .Clear flexClearScrollable

        If rs.RecordCount > 0 Then
            .rows = rs.RecordCount + 1
            rs.MoveFirst

            For i = 1 To .rows - 1
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

    RsSavRec.Fields("name").value = IIf(TxtName(mIndex).text <> "", Trim(TxtName(mIndex).text), Null)
    RsSavRec.Fields("namee").value = IIf(TxtNameE(mIndex).text <> "", Trim(TxtNameE(mIndex).text), Null)
    

    RsSavRec.update
    MsgBox " „  ⁄„·Ì… «·Õ›Ÿ »‰Ã«Õ", vbInformation + vbOKOnly + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
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

  
    If TxtModFlg2(mIndex).text = "N" Then
           ' TxtSerial1(mIndex).Text = CStr(new_id("TblAdditionsAssest", "ID", "", True))

       
        RsSavRec.AddNew
        TxtSerial1(mIndex).text = new_id("TblLensesTypes", "id", "")
        RsSavRec.Fields("ID").value = val(TxtSerial1(mIndex).text)
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

    RsSavRec.Fields("name").value = IIf(TxtName(mIndex).text <> "", Trim(TxtName(mIndex).text), Null)
    RsSavRec.Fields("namee").value = IIf(TxtNameE(mIndex).text <> "", Trim(TxtNameE(mIndex).text), Null)
    

    RsSavRec.Fields("GroupId").value = IIf(cmbGroupId.text <> "", Trim(cmbGroupId.BoundText), Null)
    RsSavRec.Fields("UnitID").value = IIf(cmbUnitID.text <> "", Trim(cmbUnitID.BoundText), Null)
    
    RsSavRec.Fields("FromSPH").value = IIf(DCBoMain(2).text <> "", Trim(DCBoMain(2).BoundText), Null)
    RsSavRec.Fields("TOSPH").value = IIf(DCBoMain(5).text <> "", Trim(DCBoMain(5).BoundText), Null)
    RsSavRec.Fields("FROMCYL").value = IIf(DCBoMain(3).text <> "", Trim(DCBoMain(3).BoundText), Null)
    RsSavRec.Fields("TOCYL").value = IIf(DCBoMain(6).text <> "", Trim(DCBoMain(6).BoundText), Null)
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
    MsgBox " „  ⁄„·Ì… «·Õ›Ÿ »‰Ã«Õ", vbInformation + vbOKOnly + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
    
    Dim mmIDD2 As Long
    mmIDD2 = val(TxtSerial1(mIndex).text)
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

    With Me.Grid2
        .rows = 2
        .Clear flexClearScrollable

        If rs.RecordCount > 0 Then
            .rows = rs.RecordCount + 1
            rs.MoveFirst

            For i = 1 To .rows - 1
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
        .rows = 2
        .Clear flexClearScrollable

        If rs.RecordCount > 0 Then
            .rows = rs.RecordCount + 1
            rs.MoveFirst

            For i = 1 To .rows - 1
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
    TxtSerial1(mIndex).text = IIf(IsNull(RsSavRec.Fields("id").value), "", RsSavRec.Fields("id").value)
    TxtName(mIndex).text = IIf(IsNull(RsSavRec.Fields("name").value), "", RsSavRec.Fields("name").value)
    TxtNameE(mIndex).text = IIf(IsNull(RsSavRec.Fields("nameE").value), "", RsSavRec.Fields("nameE").value)
    
    LabCurr_Rec(mIndex).Caption = RsSavRec.AbsolutePosition
    LabCount_Rec(mIndex).Caption = RsSavRec.RecordCount

    With Grid2

        For i = 1 To .rows - 1

            If Trim(TxtSerial1(mIndex).text) = .TextMatrix(i, .ColIndex("id")) Then
                TxtSerial1(mIndex).text = .TextMatrix(i, .ColIndex("Ser"))
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
 
If Me.TxtModFlg2(mIndex).text = "N" Or Me.TxtModFlg2(mIndex).text = "E" Then
      If Index = 8 Then
    cmbGroupId.BoundText = cmbEyeDet(8).BoundText
      End If

        cmbEyeDet(7).text = ""
        DoEvents
        
 
        
         If 1 = 1 Then
    If cboMasterType.ListIndex = 1 Then 'frames
    Dim mNameAutoGen As String
    Dim mNameAutoGenEnG As String
    
  mNameAutoGen = cmbEyeDet(0).text      'brand
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
            mNameAutoGen = cmbEyeDet(0).text      'brand
   
             mNameAutoGen = mNameAutoGen & "," & TxtModel
             mNameAutoGen = mNameAutoGen & "," & cmbEyeDet(13).text 'index
             mNameAutoGen = mNameAutoGen & "," & cmbEyeDet(6).text 'imaterial
             mNameAutoGen = mNameAutoGen & "," & cmbEyeDet(2).text 'Design
             mNameAutoGen = mNameAutoGen & "," & cmbEyeDet(1).text 'Type
             
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
    

        cmbEyeDet(7).text = ""
        DoEvents
        
 
        
         If 1 = 1 Then
    If cboMasterType.ListIndex = 1 Then 'frames
    Dim mNameAutoGen As String
    Dim mNameAutoGenEnG As String
    
  mNameAutoGen = cmbEyeDet(0).text      'brand
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
            mNameAutoGen = cmbEyeDet(0).text      'brand
   
             mNameAutoGen = mNameAutoGen & "," & TxtModel
             mNameAutoGen = mNameAutoGen & "," & cmbEyeDet(13).text 'index
             mNameAutoGen = mNameAutoGen & "," & cmbEyeDet(6).text 'imaterial
             mNameAutoGen = mNameAutoGen & "," & cmbEyeDet(2).text 'Design
             mNameAutoGen = mNameAutoGen & "," & cmbEyeDet(1).text 'Type
             
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
        

FnGenrateName = mNameAutoGenEnG & "," & (cmbEyeDet(22).text) & "," & cmbEyeDet(23).text
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
If Me.TxtModFlg2(mIndex).text = "N" Or Me.TxtModFlg2(mIndex).text = "E" Then
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
    Dim indexs As String
    Dim MainContainerName As String

    On Error Resume Next
    indexs = G.Index

    MainContainerName = GetMainForm(G.Container)
    GlobalGridName = MainContainerName & "\" & G.Name & indexs & MainFormName
    GlobalGridName = "Import"
    GetGridFileName = App.path & GlobalGridName & ".xls"

End Function
Public Function GetMainForm(ByVal obj) As String
    Dim n As String
    On Error Resume Next
    n = obj.Container.Name

    If n = "" Then
        GetMainForm = obj.Name
    Else
        GetMainForm = GetMainForm(obj.Container)
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
    mGrid.rows = 1
    
    For i = 1 To mtmpGrd.rows - 1
        If i <= mtmpGrd.rows - 1 Then
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
    
    mGrid.rows = mtmpGrd.rows

    '********************************
    If Not ProgressBar Is Nothing Then
        ProgressBar.Min = 1
        ProgressBar.Max = IIf(mGrid.rows > 2, mGrid.rows - 1, 2)    ' mGrid.Rows - 1
        ProgressBar.Visible = True
        '********************************
    End If
        Set cProgress = New ClsProgress
       cProgress.ProgressType = Waiting
    

    



    
    Dim Hide As Integer
    For i = 1 To mtmpGrd.rows - 1
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
                mGrid.rows = i + 1:  Exit Sub
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
If (right(mValue, 1)) = "Â" Then
    strValue = "…"
ElseIf (right(mValue, 1)) = "…" Then
    strValue = "Â"
    
End If
If strValue <> "" Then
    mValue3 = Replace(mValue3, right(mValue3, 1), strValue)
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
For i = 1 To mGrd.rows - 1
    If Trim(mGrd.TextMatrix(i, mGrd.ColIndex(mFldName))) = mTxt Then
        SearchInGrid = i
        Exit Function
    End If
Next
SearchInGrid = ""
End Function
Function FileExists(FileName) As Boolean
    On Error GoTo CheckError        ' Turn on error trapping so error handler                            ' responds if any error is detected.
    FileExists = (Dir(FileName) <> "")
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
        RsData!BoxName = right(Trim(mName), 50)
    Else
        RsData!BoxName = Trim(mName)
    End If

    RsData!BoxNamee = right(Trim(mName), 50)
    
  
    RsData!type = 1
    
    RsData!BranchID = val(branch_id)
    RsData!Account_code = Trim(StrNewAccountCode)
    RsData!parent_account = Trim(mParent_account)
    
    'rsData!BranchID = val(rsDummy!BranchID & "")
    RsData.update
    GrdExcel.TextMatrix(mRow, GrdExcel.ColIndex("Account_Code")) = StrNewAccountCode


    

End Sub



Private Sub DBCboClientName_Click(Area As Integer)
Dim fullcode As String


 GetCustomersDetail val(DBCboClientName.BoundText), , fullcode, 1
  TxtSearchCode.text = fullcode
    
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
        CurrenrEmployeeIDs.text = ""
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



Private Sub DcCurrency_Click(Area As Integer)
'    DcCurrency_Change
End Sub






Function savenewelectroncic()
   'vat data
    Dim InvoiceTypeCodeID As Integer
    RsSavRec("CIBAN").value = TXTIban.text
    'vat data
      RsSavRec("RecTime").value = Time
            
      RsSavRec("Currency_id").value = IIf(Dccurrency.BoundText = "", 1, val(Dccurrency.BoundText))
    RsSavRec("Currency_rate").value = IIf(Not IsNumeric(txt_Currency_rate.text), 1, txt_Currency_rate.text)
    RsSavRec("DateRec").value = txtDateRec.value
    RsSavRec("Doctype").value = IIf(Me.DCDocTypes.BoundText = "", Null, val(DCDocTypes.BoundText))
    RsSavRec("CIBAN").value = TXTIban.text
    RsSavRec("Invoicetype").value = Me.DefaultInvoicetype.ListIndex

   
  
    
  If val(DCDocTypes.BoundText) <> 0 Then
  'wAEL
    getDocAccounts val(DCDocTypes.BoundText), , , , , , , , , , , , InvoiceTypeCodeID
  Else
 InvoiceTypeCodeID = 388
  End If
  RsSavRec("InvoiceTypeCodeID").value = InvoiceTypeCodeID
 
 
 
 If val(Me.DefaultInvoicetype.ListIndex) = 0 Then
   
   
    If Export = 1 Then
    RsSavRec("InvoiceTypeCodename").value = "0100100"
    Else
      RsSavRec("InvoiceTypeCodename").value = "0100000"
   End If
   
   
   
   
   Else
    RsSavRec("InvoiceTypeCodename").value = "0200000"
   End If
 
  
  RsSavRec("InvoiceTypeCodeID").value = InvoiceTypeCodeID
 
 
 
 If val(Me.DefaultInvoicetype.ListIndex) = 0 Then
   
   
    If Export = 1 Then
    RsSavRec("InvoiceTypeCodename").value = "0100100"
    Else
      RsSavRec("InvoiceTypeCodename").value = "0100000"
   End If
   
   
   
   
   Else
    RsSavRec("InvoiceTypeCodename").value = "0200000"
   End If

   RsSavRec("DocumentCurrencyCode").value = IIf(Dccurrency.text = "", "SAR", Dccurrency.text)
   RsSavRec("TaxCurrencyCode").value = IIf(Dccurrency.text = "", "SAR", Dccurrency.text)
  RsSavRec("ActualDeliveryDate").value = txtDateRec.value
 RsSavRec("LatestDeliveryDate").value = txtDateRec.value
Dim PaymentMeansCode As String
         
            '10 In cash
            '30 Credit
            '42 Payment to bank account
            '48 Bank card
            '1 Instrument not defined(Free text)
            Dim paymentnote
'        If CboPayMentType.ListIndex = 0 Then ' ‰ﬁœ«
'                  PaymentMeansCode = "10"
'                      paymentnote = "Payment by Cash"
'        ElseIf CboPayMentType.ListIndex = 1 Then ' ¬Ã·
'                 PaymentMeansCode = "30"
'                 paymentnote = "Payment by Credit"
'         ElseIf CboPayMentType.ListIndex = 2 Or CboPayMentType.ListIndex = 3 Then  '  ÕÊÌ· »‰ﬂÌ
'                    If SystemOptions.AllowSalesMultyPayed = True Then
'                     PaymentMeansCode = "48" 'ﬂ«— 
'                      paymentnote = "Payment by Bank Card"
'                    Else
'                    PaymentMeansCode = "42" '»‰ﬂ Õ”«»
'                    paymentnote = "Payment by Bank Account"
'                    End If
'
'         End If
         PaymentMeansCode = "30"
                 paymentnote = "Payment by Credit"
         RsSavRec("PaymentMeansCode").value = PaymentMeansCode
         
         
              
       ' If CboPayMentType.ListIndex = 0 Then ' ‰ﬁœ«
                  PaymentMeansCode = "10"
                      paymentnote = "Payment by Cash"
      '  ElseIf CboPayMentType.ListIndex = 1 Then ' ¬Ã·
       '          PaymentMeansCode = "30"
       '          paymentnote = "Payment by Credit"
       '  ElseIf CboPayMentType.ListIndex = 2 Or CboPayMentType.ListIndex = 3 Then  '  ÕÊÌ· »‰ﬂÌ
         '           If SystemOptions.AllowSalesMultyPayed = True Then
         '            PaymentMeansCode = "48" 'ﬂ«— 
          '            paymentnote = "Payment by Bank Card"
         '           Else
          '          PaymentMeansCode = "42" '»‰ﬂ Õ”«»
          '          paymentnote = "Payment by Bank Account"
          '          End If
         
        ' End If
         
         RsSavRec("PaymentMeansCode").value = PaymentMeansCode
      
RsSavRec("paymentnote").value = paymentnote
RsSavRec.update
End Function



