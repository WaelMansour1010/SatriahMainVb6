VERSION 5.00
Object = "{C115893A-A3BF-43AF-B28D-69DB846077F3}#1.0#0"; "vsflex8u.ocx"
Object = "{0AFE7BE0-11B7-4A3E-978D-D4501E9A57FE}#1.0#0"; "c1sizer.ocx"
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{FE5DCFAD-BC1D-11D2-94CF-004005455FAA}#1.4#0"; "ImpulseButton.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#12.0#0"; "SUITEC~1.OCX"
Object = "{D155F1AE-D9A4-458C-8CEE-498CB717DB7B}#1.0#0"; "DBPix20.ocx"
Begin VB.Form FrmCars 
   BackColor       =   &H00E2E9E9&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "»Ì«‰«  «·„⁄œ« /«·„—ﬂ»« "
   ClientHeight    =   9945
   ClientLeft      =   -15
   ClientTop       =   375
   ClientWidth     =   11640
   Icon            =   "FrmCars.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   RightToLeft     =   -1  'True
   ScaleHeight     =   9945
   ScaleWidth      =   11640
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
   Begin C1SizerLibCtl.C1Elastic frm_Main 
      Height          =   9945
      Left            =   0
      TabIndex        =   165
      TabStop         =   0   'False
      Top             =   0
      Width           =   11640
      _cx             =   20532
      _cy             =   17542
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
      Begin C1SizerLibCtl.C1Tab C1Tab1 
         Height          =   8475
         Left            =   0
         TabIndex        =   166
         Top             =   735
         Width           =   11625
         _cx             =   20505
         _cy             =   14949
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
         Caption         =   "»Ì‰«  «”«”Ì…|«·„—›ﬁ« |»Ì«‰«  «·ÕÊ«œÀ|«·„’—Ê›« |«·„·Õﬁ« "
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
         Begin C1SizerLibCtl.C1Elastic C1Elastic9 
            Height          =   8055
            Left            =   12270
            TabIndex        =   167
            TabStop         =   0   'False
            Top             =   45
            Width           =   11535
            _cx             =   20346
            _cy             =   14208
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
            Begin C1SizerLibCtl.C1Elastic C1Elastic11 
               Height          =   2355
               Left            =   150
               TabIndex        =   168
               TabStop         =   0   'False
               Top             =   -60
               Width           =   11235
               _cx             =   19817
               _cy             =   4154
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
               Caption         =   "„” ‰œ«  «·„—ﬂ»…"
               Align           =   0
               AutoSizeChildren=   7
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
               Begin VB.TextBox TxtOwnerName2 
                  Alignment       =   2  'Center
                  Height          =   360
                  Left            =   90
                  RightToLeft     =   -1  'True
                  TabIndex        =   307
                  Top             =   810
                  Width           =   8775
               End
               Begin VB.TextBox TxtOwnerName 
                  Alignment       =   2  'Center
                  Height          =   345
                  Left            =   90
                  RightToLeft     =   -1  'True
                  TabIndex        =   304
                  Top             =   360
                  Width           =   8775
               End
               Begin VB.TextBox FormOrignal 
                  Alignment       =   1  'Right Justify
                  Height          =   345
                  Left            =   600
                  RightToLeft     =   -1  'True
                  TabIndex        =   303
                  Top             =   1185
                  Width           =   3390
               End
               Begin MSDataListLib.DataCombo DcboCountryID2 
                  Height          =   315
                  Left            =   5490
                  TabIndex        =   305
                  Tag             =   "⁄›Ê« Ì—ÃÏ «œŒ«· √”„ «·œÊ·…"
                  Top             =   1185
                  Width           =   3315
                  _ExtentX        =   5847
                  _ExtentY        =   556
                  _Version        =   393216
                  Text            =   ""
                  RightToLeft     =   -1  'True
               End
               Begin MSDataListLib.DataCombo DcboGovernmentID 
                  Height          =   315
                  Left            =   5490
                  TabIndex        =   306
                  Tag             =   "⁄›Ê« Ì—ÃÏ «œŒ«· √”„ «·„œÌ‰…"
                  Top             =   1545
                  Width           =   3315
                  _ExtentX        =   5847
                  _ExtentY        =   556
                  _Version        =   393216
                  Text            =   ""
                  RightToLeft     =   -1  'True
               End
               Begin MSComCtl2.DTPicker DpExpireDate 
                  Height          =   345
                  Left            =   7170
                  TabIndex        =   322
                  Top             =   1860
                  Width           =   1665
                  _ExtentX        =   2937
                  _ExtentY        =   609
                  _Version        =   393216
                  Format          =   104792065
                  CurrentDate     =   38784
               End
               Begin MSComCtl2.DTPicker DpSensitiveWeightDate 
                  Height          =   345
                  Left            =   1950
                  TabIndex        =   324
                  Top             =   1860
                  Width           =   1665
                  _ExtentX        =   2937
                  _ExtentY        =   609
                  _Version        =   393216
                  Format          =   104792065
                  CurrentDate     =   38784
               End
               Begin Dynamic_Byte.NourHijriCal DpSensitiveWeightDateH 
                  Height          =   315
                  Left            =   510
                  TabIndex        =   326
                  Top             =   1860
                  Width           =   1365
                  _ExtentX        =   2408
                  _ExtentY        =   556
               End
               Begin Dynamic_Byte.NourHijriCal DpExpireDateH 
                  Height          =   285
                  Left            =   5790
                  TabIndex        =   327
                  Top             =   1860
                  Width           =   1365
                  _ExtentX        =   2408
                  _ExtentY        =   503
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   " «—ÌŒ  —ﬂÌ» Õ”«” «·Ê“‰"
                  Height          =   315
                  Index           =   56
                  Left            =   3690
                  RightToLeft     =   -1  'True
                  TabIndex        =   325
                  Top             =   1905
                  Width           =   1725
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   " «—ÌŒ «‰ Â«¡ ’·«ÕÌ… «·ﬂ«— "
                  Height          =   375
                  Index           =   55
                  Left            =   8850
                  RightToLeft     =   -1  'True
                  TabIndex        =   323
                  Top             =   1905
                  Width           =   2235
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "«· »œÌ· «·Ï"
                  Height          =   360
                  Index           =   54
                  Left            =   9795
                  RightToLeft     =   -1  'True
                  TabIndex        =   308
                  Top             =   810
                  Width           =   1290
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "„«·ﬂ «·„⁄œÂ/«·”Ì«—…"
                  Height          =   495
                  Index           =   38
                  Left            =   9795
                  RightToLeft     =   -1  'True
                  TabIndex        =   259
                  Top             =   390
                  Width           =   1290
               End
               Begin VB.Label Label1 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  BackStyle       =   0  'Transparent
                  Caption         =   "«”„ «·„œÌ‰Â"
                  Height          =   285
                  Index           =   4
                  Left            =   10020
                  RightToLeft     =   -1  'True
                  TabIndex        =   255
                  Top             =   1545
                  Width           =   1065
               End
               Begin VB.Label Label1 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  BackStyle       =   0  'Transparent
                  Caption         =   "«”„ «·œÊ·Â"
                  Height          =   435
                  Index           =   1
                  Left            =   9195
                  RightToLeft     =   -1  'True
                  TabIndex        =   254
                  Top             =   1185
                  Width           =   1890
               End
               Begin VB.Label Label5 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "«·«” „«—… «·«’·Ì…"
                  Height          =   375
                  Left            =   4140
                  RightToLeft     =   -1  'True
                  TabIndex        =   169
                  Top             =   1215
                  Width           =   1365
               End
            End
            Begin C1SizerLibCtl.C1Elastic C1Elastic10 
               Height          =   5565
               Left            =   150
               TabIndex        =   170
               TabStop         =   0   'False
               Top             =   2400
               Width           =   11235
               _cx             =   19817
               _cy             =   9816
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
               Caption         =   "„—›ﬁ«  «·„—ﬂ»…"
               Align           =   0
               AutoSizeChildren=   7
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
               Begin VB.CheckBox SideFrame 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "«ÿ«— Ã«‰»Ï »Ã‰ÿ"
                  Height          =   432
                  Left            =   2250
                  RightToLeft     =   -1  'True
                  TabIndex        =   317
                  Top             =   2550
                  Width           =   1815
               End
               Begin VB.CheckBox SideBarriers 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "ÕÊ«Ã“ Ã«‰»Ì…"
                  Height          =   432
                  Left            =   2250
                  RightToLeft     =   -1  'True
                  TabIndex        =   316
                  Top             =   2220
                  Width           =   1815
               End
               Begin VB.CheckBox Sail 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "‘—«⁄"
                  Height          =   432
                  Left            =   4215
                  RightToLeft     =   -1  'True
                  TabIndex        =   315
                  Top             =   2550
                  Width           =   1815
               End
               Begin VB.CheckBox Khabor 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "ﬁ›· Œ«»Ê—"
                  Height          =   432
                  Left            =   4215
                  RightToLeft     =   -1  'True
                  TabIndex        =   314
                  Top             =   2220
                  Width           =   1815
               End
               Begin VB.CheckBox Hock 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "ÂÊﬂ"
                  Height          =   432
                  Left            =   7155
                  RightToLeft     =   -1  'True
                  TabIndex        =   313
                  Top             =   2550
                  Width           =   1815
               End
               Begin VB.CheckBox Kafla 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "ﬁ›·…"
                  Height          =   432
                  Left            =   7155
                  RightToLeft     =   -1  'True
                  TabIndex        =   312
                  Top             =   2220
                  Width           =   1815
               End
               Begin VB.CheckBox Chains 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "”·«”·"
                  Height          =   432
                  Left            =   9270
                  RightToLeft     =   -1  'True
                  TabIndex        =   311
                  Top             =   2550
                  Width           =   1815
               End
               Begin VB.CheckBox Sabt 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "”» "
                  Height          =   432
                  Left            =   9270
                  RightToLeft     =   -1  'True
                  TabIndex        =   310
                  Top             =   2220
                  Width           =   1815
               End
               Begin VB.CheckBox keys 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "„›« ÌÕ"
                  Height          =   432
                  Left            =   7170
                  RightToLeft     =   -1  'True
                  TabIndex        =   78
                  Top             =   4830
                  Width           =   1800
               End
               Begin VB.CheckBox DriLicense 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "—Œ’… "
                  Height          =   432
                  Left            =   10185
                  RightToLeft     =   -1  'True
                  TabIndex        =   67
                  Top             =   1815
                  Width           =   900
               End
               Begin VB.TextBox TxtDriLicenseNo 
                  Alignment       =   1  'Right Justify
                  Height          =   315
                  Left            =   6705
                  RightToLeft     =   -1  'True
                  TabIndex        =   68
                  Top             =   1935
                  Width           =   2190
               End
               Begin VB.TextBox TxtAuthorType 
                  Alignment       =   2  'Center
                  Height          =   330
                  Left            =   6705
                  RightToLeft     =   -1  'True
                  TabIndex        =   60
                  Top             =   840
                  Width           =   2190
               End
               Begin VB.CheckBox BagAmbulance 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "ÕﬁÌ»… «·«”⁄«›"
                  Height          =   432
                  Left            =   2265
                  RightToLeft     =   -1  'True
                  TabIndex        =   84
                  Top             =   3390
                  Width           =   1800
               End
               Begin VB.CheckBox SubsBattery 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "«‘ —«ﬂ «·»ÿ«—Ì…"
                  Height          =   432
                  Left            =   7545
                  RightToLeft     =   -1  'True
                  TabIndex        =   75
                  Top             =   3390
                  Width           =   1425
               End
               Begin VB.CheckBox FireExtingui 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "ÿ›«Ì… «·Õ—Ìﬁ"
                  Height          =   432
                  Left            =   9270
                  RightToLeft     =   -1  'True
                  TabIndex        =   70
                  Top             =   3390
                  Width           =   1815
               End
               Begin VB.CheckBox TrackingDevice 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "ÃÂ«“ «·  »⁄"
                  Height          =   432
                  Left            =   2265
                  RightToLeft     =   -1  'True
                  TabIndex        =   83
                  Top             =   2910
                  Width           =   1800
               End
               Begin VB.CheckBox Triangle 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "«·„À·À «·÷Ê∆Ì"
                  Height          =   432
                  Left            =   4230
                  RightToLeft     =   -1  'True
                  TabIndex        =   79
                  Top             =   2910
                  Width           =   1800
               End
               Begin VB.CheckBox Receipt 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "«ﬁ—«— «·«” ·«„"
                  Height          =   432
                  Left            =   7545
                  RightToLeft     =   -1  'True
                  TabIndex        =   74
                  Top             =   2910
                  Width           =   1425
               End
               Begin VB.CheckBox KeyReserve 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "«·„› «Õ «·«Õ Ì«ÿÌ"
                  Height          =   432
                  Left            =   9270
                  RightToLeft     =   -1  'True
                  TabIndex        =   69
                  Top             =   2910
                  Width           =   1815
               End
               Begin VB.TextBox authorizeExamination 
                  Alignment       =   1  'Right Justify
                  Height          =   315
                  Left            =   6705
                  RightToLeft     =   -1  'True
                  TabIndex        =   66
                  Top             =   1575
                  Width           =   2190
               End
               Begin VB.TextBox authorizeLicense 
                  Alignment       =   1  'Right Justify
                  Height          =   330
                  Left            =   3390
                  RightToLeft     =   -1  'True
                  TabIndex        =   61
                  Top             =   840
                  Width           =   1965
               End
               Begin VB.CheckBox Exam 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "«·›Õ’ "
                  Height          =   432
                  Left            =   10185
                  RightToLeft     =   -1  'True
                  TabIndex        =   65
                  Top             =   1455
                  Width           =   900
               End
               Begin VB.CheckBox Authorization 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "«· ›ÊÌ÷"
                  Height          =   450
                  Left            =   10110
                  RightToLeft     =   -1  'True
                  TabIndex        =   59
                  Top             =   720
                  Width           =   975
               End
               Begin VB.TextBox TxtLicenseNO 
                  Alignment       =   2  'Center
                  Height          =   315
                  Left            =   6705
                  RightToLeft     =   -1  'True
                  TabIndex        =   64
                  Top             =   1215
                  Width           =   2190
               End
               Begin VB.CheckBox Licenses 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "«·«” „«—…"
                  Height          =   432
                  Left            =   9885
                  RightToLeft     =   -1  'True
                  TabIndex        =   63
                  Top             =   1095
                  Width           =   1200
               End
               Begin VB.TextBox TxtInsuranceNo 
                  Alignment       =   1  'Right Justify
                  Height          =   315
                  Left            =   3390
                  RightToLeft     =   -1  'True
                  TabIndex        =   57
                  Top             =   480
                  Width           =   1965
               End
               Begin VB.CheckBox Insurance 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   " √„Ì‰"
                  Height          =   432
                  Left            =   10110
                  RightToLeft     =   -1  'True
                  TabIndex        =   55
                  Top             =   360
                  Width           =   975
               End
               Begin VB.CheckBox cleaner 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "«·„”«Õ«  Ê√“—⁄ Â« "
                  Height          =   432
                  Left            =   9270
                  RightToLeft     =   -1  'True
                  TabIndex        =   71
                  Top             =   3870
                  Width           =   1815
               End
               Begin VB.CheckBox sideMirror 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "«·„—«Ì« «·Ã«‰»Ì…"
                  Height          =   432
                  Left            =   7170
                  RightToLeft     =   -1  'True
                  TabIndex        =   76
                  Top             =   3870
                  Width           =   1800
               End
               Begin VB.CheckBox driverMirror 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "„—¬… «·”«∆ﬁ"
                  Height          =   432
                  Left            =   4140
                  RightToLeft     =   -1  'True
                  TabIndex        =   81
                  Top             =   3870
                  Width           =   1890
               End
               Begin VB.CheckBox InnerLights 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "«·‰Ê— «·œ«Œ·Ï"
                  Height          =   432
                  Left            =   2265
                  RightToLeft     =   -1  'True
                  TabIndex        =   85
                  Top             =   3870
                  Width           =   1800
               End
               Begin VB.CheckBox Pedals 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "›—‘ «·œÊ”« "
                  Height          =   450
                  Left            =   225
                  RightToLeft     =   -1  'True
                  TabIndex        =   90
                  Top             =   4350
                  Width           =   1365
               End
               Begin VB.CheckBox SunScreens 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "Ê«ﬁÌ«  «·‘„” "
                  Height          =   435
                  Left            =   9270
                  RightToLeft     =   -1  'True
                  TabIndex        =   72
                  Top             =   4350
                  Width           =   1815
               End
               Begin VB.CheckBox Recorder 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "«·—«œÌÊ Ê«·„”Ã·"
                  Height          =   435
                  Left            =   7170
                  RightToLeft     =   -1  'True
                  TabIndex        =   77
                  Top             =   4350
                  Width           =   1800
               End
               Begin VB.CheckBox Anntena 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "«·ÂÊ«∆Ï"
                  Height          =   435
                  Left            =   4140
                  RightToLeft     =   -1  'True
                  TabIndex        =   82
                  Top             =   4350
                  Width           =   1890
               End
               Begin VB.CheckBox Battery 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "«·»ÿ«—Ì« "
                  Height          =   432
                  Left            =   4230
                  RightToLeft     =   -1  'True
                  TabIndex        =   80
                  Top             =   3390
                  Width           =   1800
               End
               Begin VB.CheckBox SpareTyre 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "≈ÿ«— «Õ Ì«ÿÏ"
                  Height          =   432
                  Left            =   225
                  RightToLeft     =   -1  'True
                  TabIndex        =   89
                  Top             =   3870
                  Width           =   1365
               End
               Begin VB.CheckBox Crane 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "—«›⁄…"
                  Height          =   432
                  Left            =   225
                  RightToLeft     =   -1  'True
                  TabIndex        =   87
                  Top             =   2910
                  Width           =   1365
               End
               Begin VB.CheckBox CoverKey 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "„› «Õ ⁄Ã·"
                  Height          =   432
                  Left            =   225
                  RightToLeft     =   -1  'True
                  TabIndex        =   88
                  Top             =   3390
                  Width           =   1365
               End
               Begin VB.CheckBox Guarantee 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "‘Â«œ… ÷„«‰"
                  Height          =   432
                  Left            =   9195
                  RightToLeft     =   -1  'True
                  TabIndex        =   73
                  Top             =   4830
                  Width           =   1890
               End
               Begin VB.CheckBox Stickers 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "«” Ìﬂ— „”«⁄œ… ⁄·Ï «·ÿ—Ìﬁ"
                  Height          =   450
                  Left            =   1665
                  RightToLeft     =   -1  'True
                  TabIndex        =   86
                  Top             =   4350
                  Width           =   2400
               End
               Begin MSDataListLib.DataCombo DCInsuranceCompanyId 
                  Height          =   315
                  Left            =   6705
                  TabIndex        =   56
                  Top             =   480
                  Width           =   2190
                  _ExtentX        =   3863
                  _ExtentY        =   556
                  _Version        =   393216
                  BackColor       =   16777215
                  Text            =   ""
                  RightToLeft     =   -1  'True
               End
               Begin Dynamic_Byte.NourHijriCal DpInsuranceExpireDateH 
                  Height          =   315
                  Left            =   150
                  TabIndex        =   58
                  Top             =   480
                  Width           =   1815
                  _ExtentX        =   3201
                  _ExtentY        =   556
               End
               Begin Dynamic_Byte.NourHijriCal DpLicenseExpireDateH 
                  Height          =   285
                  Left            =   3240
                  TabIndex        =   241
                  Top             =   1215
                  Width           =   1515
                  _ExtentX        =   2672
                  _ExtentY        =   503
               End
               Begin Dynamic_Byte.NourHijriCal AuthorDate 
                  Height          =   330
                  Left            =   150
                  TabIndex        =   62
                  Top             =   840
                  Width           =   1815
                  _ExtentX        =   3201
                  _ExtentY        =   582
               End
               Begin Dynamic_Byte.NourHijriCal DpTestExpireDateH 
                  Height          =   285
                  Left            =   3240
                  TabIndex        =   246
                  Top             =   1575
                  Width           =   1515
                  _ExtentX        =   2672
                  _ExtentY        =   503
               End
               Begin Dynamic_Byte.NourHijriCal DriLicenseDate 
                  Height          =   285
                  Left            =   3240
                  TabIndex        =   256
                  Top             =   1935
                  Width           =   1515
                  _ExtentX        =   2672
                  _ExtentY        =   503
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "«‰ Â«¡ —Œ’… «·”Ê«ﬁ…"
                  Height          =   285
                  Index           =   40
                  Left            =   4905
                  RightToLeft     =   -1  'True
                  TabIndex        =   258
                  Top             =   1935
                  Width           =   1575
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "—ﬁ„ —Œ’… «·”Ê«ﬁ…"
                  Height          =   300
                  Index           =   39
                  Left            =   8895
                  RightToLeft     =   -1  'True
                  TabIndex        =   257
                  Top             =   1935
                  Width           =   1290
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "—ﬁ„ «·›Õ’ «·œÊ—Ì"
                  Height          =   300
                  Index           =   36
                  Left            =   8895
                  RightToLeft     =   -1  'True
                  TabIndex        =   248
                  Top             =   1575
                  Width           =   1290
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   " «—ÌŒ ‰Â«Ì… «·›Õ’"
                  Height          =   285
                  Index           =   120
                  Left            =   4905
                  RightToLeft     =   -1  'True
                  TabIndex        =   247
                  Top             =   1575
                  Width           =   1575
               End
               Begin VB.Label Label10 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "—ﬁ„ «· ›ÊÌ÷"
                  Height          =   390
                  Left            =   5130
                  RightToLeft     =   -1  'True
                  TabIndex        =   245
                  Top             =   840
                  Width           =   1350
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "‰Ê⁄ «· ›ÊÌ÷"
                  Height          =   405
                  Index           =   35
                  Left            =   8520
                  RightToLeft     =   -1  'True
                  TabIndex        =   244
                  Top             =   840
                  Width           =   1515
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "«‰ Â«¡ «· ›ÊÌ÷"
                  Height          =   285
                  Index           =   34
                  Left            =   1890
                  RightToLeft     =   -1  'True
                  TabIndex        =   243
                  Top             =   870
                  Width           =   1350
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   " «—ÌŒ ‰Â«Ì… «·«” „«—…"
                  Height          =   285
                  Index           =   128
                  Left            =   4905
                  RightToLeft     =   -1  'True
                  TabIndex        =   242
                  Top             =   1215
                  Width           =   1575
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "—ﬁ„ «·«” „«—…"
                  Height          =   300
                  Index           =   106
                  Left            =   8745
                  RightToLeft     =   -1  'True
                  TabIndex        =   240
                  Top             =   1215
                  Width           =   1290
               End
               Begin VB.Label Label9 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "—ﬁ„ «· √„Ì‰"
                  Height          =   375
                  Left            =   5130
                  RightToLeft     =   -1  'True
                  TabIndex        =   239
                  Top             =   480
                  Width           =   1350
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "‘—ﬂ… «· √„Ì‰"
                  Height          =   390
                  Index           =   3
                  Left            =   8520
                  RightToLeft     =   -1  'True
                  TabIndex        =   238
                  Top             =   480
                  Width           =   1515
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "‰Â«Ì… «· √„Ì‰"
                  Height          =   270
                  Index           =   127
                  Left            =   2040
                  RightToLeft     =   -1  'True
                  TabIndex        =   237
                  Top             =   510
                  Width           =   1200
               End
            End
         End
         Begin C1SizerLibCtl.C1Elastic Ele 
            Height          =   8055
            Index           =   2
            Left            =   45
            TabIndex        =   171
            TabStop         =   0   'False
            Top             =   45
            Width           =   11535
            _cx             =   20346
            _cy             =   14208
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
            Begin C1SizerLibCtl.C1Elastic C1Elastic1 
               Height          =   7995
               Left            =   0
               TabIndex        =   172
               TabStop         =   0   'False
               Top             =   0
               Width           =   11535
               _cx             =   20346
               _cy             =   14102
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
               Begin VB.Frame Frame12 
                  BackColor       =   &H00E2E9E9&
                  Height          =   615
                  Index           =   2
                  Left            =   0
                  RightToLeft     =   -1  'True
                  TabIndex        =   318
                  Top             =   2880
                  Visible         =   0   'False
                  Width           =   4320
                  Begin VB.TextBox TxtAccount 
                     Alignment       =   1  'Right Justify
                     Height          =   315
                     Left            =   4560
                     RightToLeft     =   -1  'True
                     TabIndex        =   319
                     Top             =   240
                     Width           =   705
                  End
                  Begin MSDataListLib.DataCombo DcbAccount 
                     Height          =   315
                     Left            =   120
                     TabIndex        =   320
                     Top             =   240
                     Width           =   4335
                     _ExtentX        =   7646
                     _ExtentY        =   556
                     _Version        =   393216
                     Text            =   ""
                     RightToLeft     =   -1  'True
                  End
                  Begin VB.Label lbl 
                     Alignment       =   2  'Center
                     BackColor       =   &H00E2E9E9&
                     Caption         =   "Õ”«» «·«Ì—«œ"
                     Height          =   285
                     Index           =   91
                     Left            =   5160
                     TabIndex        =   321
                     Top             =   240
                     Width           =   1305
                  End
               End
               Begin VB.CheckBox chkIsUsed 
                  Alignment       =   1  'Right Justify
                  Caption         =   "„” ⁄„·"
                  Height          =   225
                  Left            =   1410
                  RightToLeft     =   -1  'True
                  TabIndex        =   309
                  Top             =   180
                  Width           =   1185
               End
               Begin VB.CommandButton DeleteImage 
                  Caption         =   "Õ–› ’Ê—…"
                  Height          =   495
                  Left            =   360
                  RightToLeft     =   -1  'True
                  TabIndex        =   302
                  Top             =   3510
                  Width           =   1740
               End
               Begin VB.CommandButton btnAddImage 
                  Caption         =   "√œ—«Ã ’Ê—…"
                  Height          =   495
                  Left            =   2490
                  RightToLeft     =   -1  'True
                  TabIndex        =   47
                  Top             =   3480
                  Width           =   1620
               End
               Begin VB.TextBox TxtNotes 
                  Alignment       =   1  'Right Justify
                  Height          =   705
                  Left            =   75
                  MultiLine       =   -1  'True
                  RightToLeft     =   -1  'True
                  ScrollBars      =   2  'Vertical
                  TabIndex        =   54
                  Top             =   6405
                  Width           =   3390
               End
               Begin C1SizerLibCtl.C1Elastic C1Elastic8 
                  Height          =   1110
                  Left            =   75
                  TabIndex        =   173
                  TabStop         =   0   'False
                  Top             =   5280
                  Width           =   4215
                  _cx             =   7435
                  _cy             =   1958
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
                  Caption         =   "„⁄œ· «” Â·«ﬂ «·ÊﬁÊœ"
                  Align           =   0
                  AutoSizeChildren=   7
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
                  Begin VB.TextBox LetterPrice 
                     Alignment       =   2  'Center
                     Height          =   375
                     Left            =   2775
                     MaxLength       =   50
                     RightToLeft     =   -1  'True
                     TabIndex        =   52
                     Top             =   810
                     Width           =   1740
                  End
                  Begin VB.TextBox Total 
                     Alignment       =   2  'Center
                     Height          =   375
                     Left            =   135
                     Locked          =   -1  'True
                     MaxLength       =   50
                     RightToLeft     =   -1  'True
                     TabIndex        =   53
                     Top             =   810
                     Width           =   1710
                  End
                  Begin VB.TextBox LetterCount 
                     Alignment       =   2  'Center
                     Height          =   375
                     Left            =   5565
                     MaxLength       =   50
                     RightToLeft     =   -1  'True
                     TabIndex        =   51
                     Top             =   810
                     Width           =   1575
                  End
                  Begin VB.Label lbl 
                     Alignment       =   2  'Center
                     BackColor       =   &H00E2E9E9&
                     Caption         =   "*"
                     BeginProperty Font 
                        Name            =   "MS Sans Serif"
                        Size            =   13.5
                        Charset         =   178
                        Weight          =   400
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     Height          =   375
                     Index           =   33
                     Left            =   4635
                     RightToLeft     =   -1  'True
                     TabIndex        =   178
                     Top             =   810
                     Width           =   930
                  End
                  Begin VB.Label lbl 
                     Alignment       =   2  'Center
                     BackColor       =   &H00E2E9E9&
                     Caption         =   "="
                     BeginProperty Font 
                        Name            =   "MS Sans Serif"
                        Size            =   13.5
                        Charset         =   178
                        Weight          =   400
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     Height          =   375
                     Index           =   32
                     Left            =   1845
                     RightToLeft     =   -1  'True
                     TabIndex        =   177
                     Top             =   810
                     Width           =   930
                  End
                  Begin VB.Label lbl 
                     Alignment       =   1  'Right Justify
                     BackColor       =   &H00E2E9E9&
                     Caption         =   "«·«Ã„«·Ï"
                     Height          =   330
                     Index           =   31
                     Left            =   930
                     RightToLeft     =   -1  'True
                     TabIndex        =   176
                     Top             =   420
                     Width           =   915
                  End
                  Begin VB.Label lbl 
                     Alignment       =   1  'Right Justify
                     BackColor       =   &H00E2E9E9&
                     Caption         =   "”⁄— «·· —"
                     Height          =   330
                     Index           =   30
                     Left            =   3300
                     RightToLeft     =   -1  'True
                     TabIndex        =   175
                     Top             =   420
                     Width           =   1335
                  End
                  Begin VB.Label lbl 
                     Alignment       =   1  'Right Justify
                     BackColor       =   &H00E2E9E9&
                     Caption         =   "⁄œœ «·· —« "
                     Height          =   330
                     Index           =   28
                     Left            =   5700
                     RightToLeft     =   -1  'True
                     TabIndex        =   174
                     Top             =   420
                     Width           =   1440
                  End
               End
               Begin C1SizerLibCtl.C1Elastic C1Elastic4 
                  Height          =   705
                  Left            =   75
                  TabIndex        =   179
                  TabStop         =   0   'False
                  Top             =   7230
                  Width           =   4140
                  _cx             =   7303
                  _cy             =   1244
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
                  Begin VB.Label lbl 
                     Alignment       =   1  'Right Justify
                     BackColor       =   &H00E2E9E9&
                     Caption         =   " «·”Ã· «·Õ«·Ì:"
                     Height          =   525
                     Index           =   125
                     Left            =   15555
                     RightToLeft     =   -1  'True
                     TabIndex        =   183
                     Top             =   180
                     Width           =   5715
                  End
                  Begin VB.Label lbl 
                     Alignment       =   1  'Right Justify
                     BackColor       =   &H00E2E9E9&
                     Caption         =   " ⁄œœ «·”Ã·« :"
                     Height          =   525
                     Index           =   126
                     Left            =   9045
                     RightToLeft     =   -1  'True
                     TabIndex        =   182
                     Top             =   180
                     Width           =   3180
                  End
                  Begin VB.Label XPTxtCount 
                     Alignment       =   1  'Right Justify
                     BackColor       =   &H00E2E9E9&
                     Height          =   540
                     Left            =   105
                     RightToLeft     =   -1  'True
                     TabIndex        =   181
                     Top             =   120
                     Width           =   1080
                  End
                  Begin VB.Label XPTxtCurrent 
                     Alignment       =   1  'Right Justify
                     BackColor       =   &H00E2E9E9&
                     Height          =   420
                     Left            =   2625
                     RightToLeft     =   -1  'True
                     TabIndex        =   180
                     Top             =   240
                     Width           =   1290
                  End
               End
               Begin C1SizerLibCtl.C1Elastic C1Elastic3 
                  Height          =   4530
                  Left            =   4455
                  TabIndex        =   184
                  TabStop         =   0   'False
                  Top             =   3420
                  Width           =   7080
                  _cx             =   12488
                  _cy             =   7990
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
                  Begin VB.TextBox TxtTrackingNo 
                     Alignment       =   2  'Center
                     Height          =   360
                     Left            =   3690
                     MaxLength       =   50
                     RightToLeft     =   -1  'True
                     TabIndex        =   26
                     Top             =   600
                     Width           =   1890
                  End
                  Begin VB.TextBox txtSetCount 
                     Alignment       =   2  'Center
                     Height          =   360
                     Left            =   3690
                     MaxLength       =   50
                     RightToLeft     =   -1  'True
                     TabIndex        =   32
                     Top             =   1815
                     Width           =   1890
                  End
                  Begin VB.TextBox Chesis 
                     Alignment       =   2  'Center
                     Height          =   345
                     Left            =   3690
                     MaxLength       =   50
                     RightToLeft     =   -1  'True
                     TabIndex        =   43
                     Top             =   3705
                     Width           =   1890
                  End
                  Begin VB.TextBox TxtGearno1 
                     Alignment       =   2  'Center
                     Height          =   345
                     Left            =   150
                     MaxLength       =   50
                     RightToLeft     =   -1  'True
                     TabIndex        =   42
                     Top             =   3330
                     Width           =   1665
                  End
                  Begin VB.TextBox TxtMachineno1 
                     Alignment       =   2  'Center
                     Height          =   345
                     Left            =   3690
                     MaxLength       =   50
                     RightToLeft     =   -1  'True
                     TabIndex        =   40
                     Top             =   3330
                     Width           =   1890
                  End
                  Begin VB.TextBox TxtGearno 
                     Alignment       =   2  'Center
                     Height          =   330
                     Left            =   150
                     MaxLength       =   50
                     RightToLeft     =   -1  'True
                     TabIndex        =   38
                     Top             =   2985
                     Width           =   1665
                  End
                  Begin VB.TextBox TxtMachineno 
                     Alignment       =   2  'Center
                     Height          =   330
                     Left            =   3690
                     MaxLength       =   50
                     RightToLeft     =   -1  'True
                     TabIndex        =   37
                     Top             =   2985
                     Width           =   1890
                  End
                  Begin VB.TextBox txtOperatorN 
                     Alignment       =   2  'Center
                     Height          =   375
                     Left            =   150
                     MaxLength       =   50
                     RightToLeft     =   -1  'True
                     TabIndex        =   36
                     Top             =   2550
                     Width           =   1665
                  End
                  Begin VB.TextBox txtRep 
                     Alignment       =   2  'Center
                     Height          =   315
                     Left            =   150
                     MaxLength       =   50
                     RightToLeft     =   -1  'True
                     TabIndex        =   34
                     Top             =   2190
                     Width           =   1665
                  End
                  Begin VB.TextBox txtMax 
                     Alignment       =   2  'Center
                     Height          =   405
                     Left            =   3690
                     Locked          =   -1  'True
                     MaxLength       =   50
                     RightToLeft     =   -1  'True
                     TabIndex        =   35
                     Top             =   2550
                     Width           =   1890
                  End
                  Begin VB.TextBox txtBoardNO 
                     Alignment       =   2  'Center
                     Enabled         =   0   'False
                     Height          =   375
                     Left            =   4065
                     MaxLength       =   50
                     RightToLeft     =   -1  'True
                     TabIndex        =   17
                     Top             =   120
                     Width           =   1515
                  End
                  Begin VB.TextBox txtModel 
                     Alignment       =   2  'Center
                     Height          =   360
                     Left            =   3690
                     MaxLength       =   50
                     RightToLeft     =   -1  'True
                     TabIndex        =   28
                     Top             =   1005
                     Width           =   1890
                  End
                  Begin VB.TextBox txtLastKMCounter 
                     Alignment       =   2  'Center
                     Height          =   330
                     Left            =   150
                     MaxLength       =   50
                     RightToLeft     =   -1  'True
                     TabIndex        =   31
                     Top             =   1410
                     Width           =   1665
                  End
                  Begin VB.TextBox VehicleLong 
                     Alignment       =   2  'Center
                     Height          =   300
                     Left            =   3690
                     MaxLength       =   50
                     RightToLeft     =   -1  'True
                     TabIndex        =   30
                     Top             =   1410
                     Width           =   1890
                  End
                  Begin VB.TextBox TxtEquQty 
                     Alignment       =   2  'Center
                     Height          =   315
                     Left            =   900
                     MaxLength       =   50
                     RightToLeft     =   -1  'True
                     TabIndex        =   29
                     Top             =   1050
                     Width           =   915
                  End
                  Begin VB.TextBox txtCapacity 
                     Alignment       =   2  'Center
                     Height          =   315
                     Left            =   3690
                     Locked          =   -1  'True
                     MaxLength       =   50
                     RightToLeft     =   -1  'True
                     TabIndex        =   33
                     Top             =   2205
                     Width           =   1890
                  End
                  Begin MSComCtl2.DTPicker DpPurchaseDate 
                     Height          =   345
                     Left            =   150
                     TabIndex        =   27
                     Top             =   630
                     Width           =   1665
                     _ExtentX        =   2937
                     _ExtentY        =   609
                     _Version        =   393216
                     Format          =   104792065
                     CurrentDate     =   38784
                  End
                  Begin C1SizerLibCtl.C1Elastic C1Elastic7 
                     Height          =   435
                     Left            =   150
                     TabIndex        =   185
                     TabStop         =   0   'False
                     Top             =   120
                     Width           =   3690
                     _cx             =   6509
                     _cy             =   767
                     BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                        Name            =   "Arial"
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
                     Begin VB.TextBox txtLetter1 
                        Alignment       =   2  'Center
                        Height          =   390
                        Left            =   3240
                        MaxLength       =   1
                        RightToLeft     =   -1  'True
                        TabIndex        =   18
                        Top             =   0
                        Width           =   450
                     End
                     Begin VB.TextBox txtLetter2 
                        Alignment       =   2  'Center
                        Height          =   390
                        Left            =   2865
                        MaxLength       =   1
                        RightToLeft     =   -1  'True
                        TabIndex        =   19
                        Top             =   0
                        Width           =   375
                     End
                     Begin VB.TextBox txtLetter3 
                        Alignment       =   2  'Center
                        Height          =   390
                        Left            =   2415
                        MaxLength       =   1
                        RightToLeft     =   -1  'True
                        TabIndex        =   20
                        Top             =   0
                        Width           =   525
                     End
                     Begin VB.TextBox txtNum1 
                        Alignment       =   2  'Center
                        Height          =   390
                        Left            =   1350
                        MaxLength       =   1
                        RightToLeft     =   -1  'True
                        TabIndex        =   22
                        Top             =   0
                        Width           =   615
                     End
                     Begin VB.TextBox txtNum2 
                        Alignment       =   2  'Center
                        Height          =   390
                        Left            =   825
                        MaxLength       =   1
                        RightToLeft     =   -1  'True
                        TabIndex        =   23
                        Top             =   0
                        Width           =   525
                     End
                     Begin VB.TextBox txtNum3 
                        Alignment       =   2  'Center
                        Height          =   390
                        Left            =   450
                        MaxLength       =   1
                        RightToLeft     =   -1  'True
                        TabIndex        =   24
                        Top             =   0
                        Width           =   450
                     End
                     Begin VB.TextBox txtLetter4 
                        Alignment       =   2  'Center
                        Height          =   390
                        Left            =   1890
                        MaxLength       =   1
                        RightToLeft     =   -1  'True
                        TabIndex        =   21
                        Top             =   0
                        Width           =   675
                     End
                     Begin VB.TextBox txtNum4 
                        Alignment       =   2  'Center
                        Height          =   390
                        Left            =   0
                        MaxLength       =   1
                        RightToLeft     =   -1  'True
                        TabIndex        =   25
                        Top             =   0
                        Width           =   525
                     End
                  End
                  Begin MSDataListLib.DataCombo VColor 
                     Height          =   315
                     Left            =   150
                     TabIndex        =   44
                     Tag             =   "Õœœ «”„ «·„⁄œ…"
                     Top             =   3705
                     Width           =   1665
                     _ExtentX        =   2937
                     _ExtentY        =   556
                     _Version        =   393216
                     Text            =   ""
                     RightToLeft     =   -1  'True
                  End
                  Begin MSDataListLib.DataCombo VModel 
                     Height          =   315
                     Left            =   3690
                     TabIndex        =   45
                     Tag             =   "Õœœ «”„ «·„⁄œ…"
                     Top             =   4080
                     Width           =   1890
                     _ExtentX        =   3334
                     _ExtentY        =   556
                     _Version        =   393216
                     Text            =   ""
                     RightToLeft     =   -1  'True
                  End
                  Begin MSDataListLib.DataCombo LocationID 
                     Height          =   315
                     Left            =   150
                     TabIndex        =   46
                     Tag             =   "Õœœ «”„ «·„⁄œ…"
                     Top             =   4065
                     Width           =   1665
                     _ExtentX        =   2937
                     _ExtentY        =   556
                     _Version        =   393216
                     Text            =   ""
                     RightToLeft     =   -1  'True
                  End
                  Begin MSDataListLib.DataCombo VType 
                     Height          =   315
                     Left            =   8430
                     TabIndex        =   186
                     Tag             =   "Õœœ «”„ «·„⁄œ…"
                     Top             =   4935
                     Visible         =   0   'False
                     Width           =   1965
                     _ExtentX        =   3466
                     _ExtentY        =   556
                     _Version        =   393216
                     Text            =   ""
                     RightToLeft     =   -1  'True
                  End
                  Begin VB.Label lbl 
                     Alignment       =   1  'Right Justify
                     BackColor       =   &H00E2E9E9&
                     Caption         =   "—ﬁ„ «·  »⁄"
                     Height          =   360
                     Index           =   37
                     Left            =   5505
                     RightToLeft     =   -1  'True
                     TabIndex        =   253
                     Top             =   600
                     Width           =   1350
                  End
                  Begin VB.Label lbl 
                     Alignment       =   1  'Right Justify
                     BackColor       =   &H00E2E9E9&
                     Caption         =   "„Ê«ﬁ⁄ «·⁄„· "
                     Height          =   315
                     Index           =   27
                     Left            =   2115
                     RightToLeft     =   -1  'True
                     TabIndex        =   209
                     Top             =   4065
                     Width           =   1350
                  End
                  Begin VB.Label lbl 
                     Alignment       =   1  'Right Justify
                     BackColor       =   &H00E2E9E9&
                     Caption         =   "‰Ê⁄ «·„—ﬂ»… "
                     Height          =   330
                     Index           =   26
                     Left            =   9120
                     RightToLeft     =   -1  'True
                     TabIndex        =   208
                     Top             =   4935
                     Visible         =   0   'False
                     Width           =   1575
                  End
                  Begin VB.Label lbl 
                     Alignment       =   1  'Right Justify
                     BackColor       =   &H00E2E9E9&
                     Caption         =   "««·ÿ—«“"
                     Height          =   210
                     Index           =   25
                     Left            =   5505
                     RightToLeft     =   -1  'True
                     TabIndex        =   207
                     Top             =   4080
                     Width           =   1425
                  End
                  Begin VB.Label lbl 
                     Alignment       =   1  'Right Justify
                     BackColor       =   &H00E2E9E9&
                     Caption         =   "·Ê‰ «·„—ﬂ»…"
                     Height          =   225
                     Index           =   24
                     Left            =   1965
                     RightToLeft     =   -1  'True
                     TabIndex        =   206
                     Top             =   3705
                     Width           =   1500
                  End
                  Begin VB.Label lbl 
                     Alignment       =   1  'Right Justify
                     BackColor       =   &H00E2E9E9&
                     Caption         =   "—ﬁ„ «·ÂÌﬂ·"
                     Height          =   300
                     Index           =   23
                     Left            =   5505
                     RightToLeft     =   -1  'True
                     TabIndex        =   205
                     Top             =   3705
                     Width           =   1350
                  End
                  Begin VB.Label lbl 
                     Alignment       =   1  'Right Justify
                     BackColor       =   &H00E2E9E9&
                     Caption         =   "—ﬁ„ «⁄œ«œ «·ÃÌ—"
                     Height          =   330
                     Index           =   22
                     Left            =   1890
                     RightToLeft     =   -1  'True
                     TabIndex        =   204
                     Top             =   3330
                     Width           =   1725
                  End
                  Begin VB.Label lbl 
                     Alignment       =   1  'Right Justify
                     BackColor       =   &H00E2E9E9&
                     Caption         =   "—ﬁ„ «⁄œ«œ «·„Õ—ﬂ"
                     Height          =   330
                     Index           =   21
                     Left            =   5040
                     RightToLeft     =   -1  'True
                     TabIndex        =   203
                     Top             =   3330
                     Width           =   1890
                  End
                  Begin VB.Label lbl 
                     Alignment       =   1  'Right Justify
                     BackColor       =   &H00E2E9E9&
                     Caption         =   "—ﬁ„  ”·”· «·ÃÌ—"
                     Height          =   315
                     Index           =   20
                     Left            =   1890
                     RightToLeft     =   -1  'True
                     TabIndex        =   202
                     Top             =   2985
                     Width           =   1725
                  End
                  Begin VB.Label lbl 
                     Alignment       =   1  'Right Justify
                     BackColor       =   &H00E2E9E9&
                     Caption         =   "—ﬁ„  ”·”· «·„Õ—ﬂ"
                     Height          =   315
                     Index           =   19
                     Left            =   5190
                     RightToLeft     =   -1  'True
                     TabIndex        =   201
                     Top             =   2985
                     Width           =   1890
                  End
                  Begin VB.Label lbl 
                     Alignment       =   1  'Right Justify
                     BackColor       =   &H00E2E9E9&
                     Caption         =   "«·—ﬁ„ «· ‘€Ì·Ï"
                     Height          =   375
                     Index           =   17
                     Left            =   1965
                     RightToLeft     =   -1  'True
                     TabIndex        =   200
                     Top             =   2550
                     Width           =   1650
                  End
                  Begin VB.Label lbl 
                     Alignment       =   1  'Right Justify
                     BackColor       =   &H00E2E9E9&
                     Caption         =   "«·Õœ «·«ﬁ’Ï ··—œÊœ"
                     Height          =   315
                     Index           =   15
                     Left            =   1965
                     RightToLeft     =   -1  'True
                     TabIndex        =   199
                     Top             =   2190
                     Width           =   1650
                  End
                  Begin VB.Label lbl 
                     Alignment       =   1  'Right Justify
                     BackColor       =   &H00E2E9E9&
                     Caption         =   "«ﬁ’Ï ”⁄…"
                     Height          =   375
                     Index           =   16
                     Left            =   5355
                     RightToLeft     =   -1  'True
                     TabIndex        =   198
                     Top             =   2550
                     Width           =   1500
                  End
                  Begin VB.Label txtRate 
                     Alignment       =   2  'Center
                     BackColor       =   &H00FFFFFF&
                     Caption         =   "1.3"
                     Height          =   345
                     Left            =   150
                     RightToLeft     =   -1  'True
                     TabIndex        =   197
                     Top             =   1845
                     Width           =   1665
                  End
                  Begin VB.Label lbl 
                     Alignment       =   1  'Right Justify
                     BackColor       =   &H00E2E9E9&
                     Caption         =   "‰”»… «·«—ﬂ«»"
                     Height          =   345
                     Index           =   13
                     Left            =   1965
                     RightToLeft     =   -1  'True
                     TabIndex        =   196
                     Top             =   1845
                     Width           =   1650
                  End
                  Begin VB.Label lbl 
                     Alignment       =   1  'Right Justify
                     BackColor       =   &H00E2E9E9&
                     Caption         =   "⁄œœ «·„ﬁ«⁄œ"
                     Height          =   360
                     Index           =   12
                     Left            =   5355
                     RightToLeft     =   -1  'True
                     TabIndex        =   195
                     Top             =   1815
                     Width           =   1425
                  End
                  Begin VB.Label lbl 
                     Alignment       =   1  'Right Justify
                     BackColor       =   &H00E2E9E9&
                     Caption         =   "«·„ÊœÌ·"
                     Height          =   360
                     Index           =   107
                     Left            =   5355
                     RightToLeft     =   -1  'True
                     TabIndex        =   194
                     Top             =   1005
                     Width           =   1425
                  End
                  Begin VB.Label lbl 
                     Alignment       =   1  'Right Justify
                     BackColor       =   &H00E2E9E9&
                     Caption         =   "—ﬁ„ «··ÊÕ…"
                     Height          =   315
                     Index           =   105
                     Left            =   5355
                     RightToLeft     =   -1  'True
                     TabIndex        =   193
                     Top             =   240
                     Width           =   1425
                  End
                  Begin VB.Label lbl 
                     Alignment       =   1  'Right Justify
                     BackColor       =   &H00E2E9E9&
                     Caption         =   " «—ÌŒ  «·‘—«¡"
                     Height          =   375
                     Index           =   1
                     Left            =   1890
                     RightToLeft     =   -1  'True
                     TabIndex        =   192
                     Top             =   645
                     Width           =   1725
                  End
                  Begin VB.Label lbl 
                     Alignment       =   1  'Right Justify
                     BackColor       =   &H00E2E9E9&
                     Caption         =   "«Œ— ﬁ—«¡… ··⁄œ«œ"
                     Height          =   330
                     Index           =   2
                     Left            =   1965
                     RightToLeft     =   -1  'True
                     TabIndex        =   191
                     Top             =   1395
                     Width           =   1650
                  End
                  Begin VB.Label lbl 
                     Alignment       =   1  'Right Justify
                     BackColor       =   &H00E2E9E9&
                     Caption         =   "ÿÊ· «·„⁄œ…"
                     Height          =   300
                     Index           =   6
                     Left            =   5355
                     RightToLeft     =   -1  'True
                     TabIndex        =   190
                     Top             =   1395
                     Width           =   1425
                  End
                  Begin VB.Label lbl 
                     Alignment       =   1  'Right Justify
                     BackColor       =   &H00E2E9E9&
                     Caption         =   "«·Õ„Ê·…"
                     Height          =   315
                     Index           =   7
                     Left            =   2265
                     RightToLeft     =   -1  'True
                     TabIndex        =   189
                     Top             =   1050
                     Width           =   1350
                  End
                  Begin VB.Label Label4 
                     Alignment       =   1  'Right Justify
                     BackStyle       =   0  'Transparent
                     Caption         =   "ÿ‰"
                     Height          =   360
                     Left            =   300
                     RightToLeft     =   -1  'True
                     TabIndex        =   188
                     Top             =   1050
                     Width           =   450
                  End
                  Begin VB.Label lbl 
                     Alignment       =   1  'Right Justify
                     BackColor       =   &H00E2E9E9&
                     Caption         =   "”⁄… «·„⁄œ…"
                     Height          =   315
                     Index           =   8
                     Left            =   5355
                     RightToLeft     =   -1  'True
                     TabIndex        =   187
                     Top             =   2190
                     Width           =   1500
                  End
               End
               Begin C1SizerLibCtl.C1Elastic C1Elastic2 
                  Height          =   3420
                  Left            =   4455
                  TabIndex        =   210
                  TabStop         =   0   'False
                  Top             =   60
                  Width           =   7080
                  _cx             =   12488
                  _cy             =   6033
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
                  Begin VB.ComboBox DcbStuts 
                     Height          =   315
                     Left            =   1275
                     RightToLeft     =   -1  'True
                     TabIndex        =   5
                     Top             =   480
                     Width           =   1440
                  End
                  Begin VB.TextBox TxtDepartment 
                     Alignment       =   2  'Center
                     Height          =   315
                     Left            =   300
                     RightToLeft     =   -1  'True
                     TabIndex        =   16
                     Top             =   2925
                     Width           =   5730
                  End
                  Begin VB.TextBox TxtJob 
                     Alignment       =   2  'Center
                     Height          =   330
                     Left            =   300
                     RightToLeft     =   -1  'True
                     TabIndex        =   15
                     Top             =   2550
                     Width           =   2340
                  End
                  Begin VB.TextBox TxtNatinality 
                     Alignment       =   2  'Center
                     Height          =   330
                     Left            =   3540
                     RightToLeft     =   -1  'True
                     TabIndex        =   14
                     Top             =   2550
                     Width           =   2490
                  End
                  Begin VB.TextBox TxtEqupName 
                     Alignment       =   2  'Center
                     Height          =   285
                     Left            =   75
                     RightToLeft     =   -1  'True
                     TabIndex        =   6
                     Top             =   840
                     Width           =   5280
                  End
                  Begin VB.TextBox txtid 
                     Alignment       =   1  'Right Justify
                     Height          =   330
                     Left            =   4440
                     MaxLength       =   50
                     RightToLeft     =   -1  'True
                     TabIndex        =   3
                     Top             =   465
                     Width           =   915
                  End
                  Begin XtremeSuiteControls.RadioButton Rd 
                     Height          =   270
                     Index           =   0
                     Left            =   5190
                     TabIndex        =   0
                     Top             =   120
                     Width           =   1815
                     _Version        =   786432
                     _ExtentX        =   3201
                     _ExtentY        =   476
                     _StockProps     =   79
                     Caption         =   "—»ÿ „‰ „·› «·«’Ê·"
                     UseVisualStyle  =   -1  'True
                     TextAlignment   =   1
                     RightToLeft     =   -1  'True
                  End
                  Begin MSDataListLib.DataCombo DCPreFix 
                     Height          =   315
                     Left            =   3615
                     TabIndex        =   4
                     Top             =   465
                     Width           =   825
                     _ExtentX        =   1455
                     _ExtentY        =   556
                     _Version        =   393216
                     BackColor       =   16777215
                     Text            =   ""
                     RightToLeft     =   -1  'True
                  End
                  Begin MSDataListLib.DataCombo DCGroup 
                     Height          =   315
                     Left            =   75
                     TabIndex        =   8
                     Top             =   1170
                     Width           =   5280
                     _ExtentX        =   9313
                     _ExtentY        =   556
                     _Version        =   393216
                     BackColor       =   16777215
                     Text            =   ""
                     RightToLeft     =   -1  'True
                  End
                  Begin MSDataListLib.DataCombo DcFixedAssets 
                     Height          =   315
                     Left            =   75
                     TabIndex        =   7
                     Tag             =   "Õœœ «”„ «·„⁄œ…"
                     Top             =   855
                     Width           =   5280
                     _ExtentX        =   9313
                     _ExtentY        =   556
                     _Version        =   393216
                     Text            =   ""
                     RightToLeft     =   -1  'True
                  End
                  Begin XtremeSuiteControls.RadioButton Rd 
                     Height          =   270
                     Index           =   1
                     Left            =   3615
                     TabIndex        =   1
                     Top             =   120
                     Width           =   1350
                     _Version        =   786432
                     _ExtentX        =   2381
                     _ExtentY        =   476
                     _StockProps     =   79
                     Caption         =   "ÌœÊÌ"
                     UseVisualStyle  =   -1  'True
                     TextAlignment   =   1
                     RightToLeft     =   -1  'True
                  End
                  Begin C1SizerLibCtl.C1Elastic C1Elastic12 
                     Height          =   870
                     Left            =   0
                     TabIndex        =   234
                     TabStop         =   0   'False
                     Top             =   1575
                     Width           =   7005
                     _cx             =   12356
                     _cy             =   1535
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
                     Begin VB.TextBox TxtLeaderName 
                        Alignment       =   1  'Right Justify
                        Height          =   300
                        Left            =   150
                        RightToLeft     =   -1  'True
                        TabIndex        =   13
                        Top             =   480
                        Width           =   4740
                     End
                     Begin VB.TextBox Text6 
                        Alignment       =   1  'Right Justify
                        Height          =   330
                        Left            =   3990
                        RightToLeft     =   -1  'True
                        TabIndex        =   10
                        Top             =   120
                        Width           =   900
                     End
                     Begin XtremeSuiteControls.RadioButton ChDrievType 
                        Height          =   255
                        Index           =   0
                        Left            =   4890
                        TabIndex        =   9
                        Top             =   120
                        Width           =   1215
                        _Version        =   786432
                        _ExtentX        =   2143
                        _ExtentY        =   450
                        _StockProps     =   79
                        Caption         =   "„ÊŸ›"
                        UseVisualStyle  =   -1  'True
                        TextAlignment   =   1
                        RightToLeft     =   -1  'True
                     End
                     Begin XtremeSuiteControls.RadioButton ChDrievType 
                        Height          =   255
                        Index           =   1
                        Left            =   4515
                        TabIndex        =   12
                        Top             =   495
                        Width           =   1590
                        _Version        =   786432
                        _ExtentX        =   2805
                        _ExtentY        =   450
                        _StockProps     =   79
                        Caption         =   "≈÷«›… „ÊŸ›"
                        UseVisualStyle  =   -1  'True
                        TextAlignment   =   1
                        RightToLeft     =   -1  'True
                     End
                     Begin MSDataListLib.DataCombo DcEmployee 
                        Height          =   315
                        Left            =   150
                        TabIndex        =   11
                        Top             =   120
                        Width           =   3840
                        _ExtentX        =   6773
                        _ExtentY        =   556
                        _Version        =   393216
                        BackColor       =   16777215
                        Text            =   ""
                        RightToLeft     =   -1  'True
                     End
                     Begin VB.Label lbl 
                        Alignment       =   1  'Right Justify
                        BackColor       =   &H00E2E9E9&
                        Caption         =   "ﬁ«∆œ «·„⁄œ…"
                        Height          =   270
                        Index           =   29
                        Left            =   5430
                        RightToLeft     =   -1  'True
                        TabIndex        =   236
                        Top             =   0
                        Width           =   1500
                     End
                     Begin VB.Label Label8 
                        Alignment       =   1  'Right Justify
                        BackColor       =   &H00E2E9E9&
                        Height          =   510
                        Left            =   150
                        RightToLeft     =   -1  'True
                        TabIndex        =   235
                        Top             =   240
                        Width           =   975
                     End
                  End
                  Begin MSDataListLib.DataCombo DcbDept 
                     Height          =   315
                     Left            =   3240
                     TabIndex        =   265
                     Top             =   120
                     Visible         =   0   'False
                     Width           =   2865
                     _ExtentX        =   5054
                     _ExtentY        =   556
                     _Version        =   393216
                     BackColor       =   16777215
                     Text            =   ""
                     RightToLeft     =   -1  'True
                  End
                  Begin MSDataListLib.DataCombo DcbJob 
                     Height          =   315
                     Left            =   75
                     TabIndex        =   2
                     Top             =   120
                     Visible         =   0   'False
                     Width           =   2865
                     _ExtentX        =   5054
                     _ExtentY        =   556
                     _Version        =   393216
                     BackColor       =   16777215
                     Text            =   ""
                     RightToLeft     =   -1  'True
                  End
                  Begin MSDataListLib.DataCombo dcBranch 
                     Height          =   315
                     Left            =   75
                     TabIndex        =   273
                     Top             =   105
                     Width           =   2865
                     _ExtentX        =   5054
                     _ExtentY        =   556
                     _Version        =   393216
                     BackColor       =   16777215
                     Text            =   ""
                     RightToLeft     =   -1  'True
                  End
                  Begin XtremeSuiteControls.PushButton PushButton1 
                     Height          =   330
                     Left            =   75
                     TabIndex        =   275
                     Top             =   480
                     Width           =   1125
                     _Version        =   786432
                     _ExtentX        =   1984
                     _ExtentY        =   582
                     _StockProps     =   79
                     Caption         =   " ÕœÌÀ «·Õ«·…"
                     BackColor       =   12640511
                     UseVisualStyle  =   -1  'True
                  End
                  Begin VB.Label lbl 
                     Alignment       =   1  'Right Justify
                     BackColor       =   &H00E2E9E9&
                     Caption         =   "Õ«·…«·„⁄œ…"
                     Height          =   210
                     Index           =   47
                     Left            =   2790
                     RightToLeft     =   -1  'True
                     TabIndex        =   272
                     Top             =   480
                     Width           =   675
                  End
                  Begin VB.Label lbl 
                     Alignment       =   1  'Right Justify
                     BackColor       =   &H00E2E9E9&
                     Caption         =   "«·«œ«—…"
                     Height          =   345
                     Index           =   43
                     Left            =   5580
                     RightToLeft     =   -1  'True
                     TabIndex        =   262
                     Top             =   2925
                     Width           =   1425
                  End
                  Begin VB.Label lbl 
                     Alignment       =   1  'Right Justify
                     BackColor       =   &H00E2E9E9&
                     Caption         =   "«·ÊŸÌ›…"
                     Height          =   360
                     Index           =   42
                     Left            =   2040
                     RightToLeft     =   -1  'True
                     TabIndex        =   261
                     Top             =   2550
                     Width           =   1425
                  End
                  Begin VB.Label lbl 
                     Alignment       =   1  'Right Justify
                     BackColor       =   &H00E2E9E9&
                     Caption         =   "«·Ã‰”Ì…"
                     Height          =   360
                     Index           =   41
                     Left            =   5580
                     RightToLeft     =   -1  'True
                     TabIndex        =   260
                     Top             =   2550
                     Width           =   1425
                  End
                  Begin VB.Label lbl 
                     Alignment       =   1  'Right Justify
                     BackColor       =   &H00E2E9E9&
                     Caption         =   "«·„⁄œ…"
                     Height          =   315
                     Index           =   18
                     Left            =   3015
                     RightToLeft     =   -1  'True
                     TabIndex        =   215
                     Top             =   -240
                     Width           =   600
                     WordWrap        =   -1  'True
                  End
                  Begin VB.Label lbl 
                     Alignment       =   1  'Right Justify
                     BackColor       =   &H00E2E9E9&
                     Caption         =   "ﬂÊœ «·„⁄œ…"
                     Height          =   210
                     Index           =   101
                     Left            =   5355
                     RightToLeft     =   -1  'True
                     TabIndex        =   214
                     Top             =   465
                     Width           =   1425
                  End
                  Begin VB.Label lbl 
                     Alignment       =   1  'Right Justify
                     BackColor       =   &H00E2E9E9&
                     Caption         =   "«”„ «·„⁄œ…"
                     Height          =   255
                     Index           =   102
                     Left            =   5355
                     RightToLeft     =   -1  'True
                     TabIndex        =   213
                     Top             =   855
                     Width           =   1500
                  End
                  Begin VB.Label lbl 
                     Alignment       =   1  'Right Justify
                     BackColor       =   &H00E2E9E9&
                     Caption         =   "‰Ê⁄ «·„⁄œ…"
                     Height          =   270
                     Index           =   103
                     Left            =   5355
                     RightToLeft     =   -1  'True
                     TabIndex        =   212
                     Top             =   1170
                     Width           =   1500
                  End
                  Begin VB.Label lbl 
                     Alignment       =   1  'Right Justify
                     BackColor       =   &H00E2E9E9&
                     Caption         =   "«·›—⁄"
                     Height          =   315
                     Index           =   117
                     Left            =   2940
                     RightToLeft     =   -1  'True
                     TabIndex        =   211
                     Top             =   105
                     Width           =   525
                  End
               End
               Begin DBPIXLib.DBPix20 DBPix201 
                  Height          =   2385
                  Left            =   75
                  TabIndex        =   216
                  Top             =   660
                  Width           =   3990
                  _Version        =   131072
                  _ExtentX        =   7038
                  _ExtentY        =   4207
                  _StockProps     =   1
                  BackColor       =   16777152
                  _Image          =   "FrmCars.frx":000C
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
               Begin C1SizerLibCtl.C1Elastic C1Elastic6 
                  Height          =   1275
                  Left            =   75
                  TabIndex        =   249
                  TabStop         =   0   'False
                  Top             =   3975
                  Width           =   4215
                  _cx             =   7435
                  _cy             =   2249
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
                  Begin VB.TextBox txtContractNo 
                     Alignment       =   2  'Center
                     Height          =   285
                     Left            =   225
                     RightToLeft     =   -1  'True
                     TabIndex        =   48
                     Top             =   120
                     Width           =   2340
                  End
                  Begin MSComCtl2.DTPicker dtpEndContractDate 
                     Height          =   285
                     Left            =   1200
                     TabIndex        =   49
                     Top             =   480
                     Width           =   1365
                     _ExtentX        =   2408
                     _ExtentY        =   503
                     _Version        =   393216
                     Format          =   104792065
                     CurrentDate     =   38784
                  End
                  Begin Dynamic_Byte.NourHijriCal dtpEndContractDateH 
                     Height          =   285
                     Left            =   1200
                     TabIndex        =   50
                     Top             =   870
                     Width           =   1365
                     _ExtentX        =   2408
                     _ExtentY        =   503
                  End
                  Begin VB.Label lbl 
                     Alignment       =   1  'Right Justify
                     BackColor       =   &H00E2E9E9&
                     Caption         =   " «—ÌŒ ‰Â«Ì… «· Œ’Ì’"
                     Height          =   405
                     Index           =   10
                     Left            =   2640
                     RightToLeft     =   -1  'True
                     TabIndex        =   252
                     Top             =   480
                     Width           =   1425
                  End
                  Begin VB.Label lbl 
                     Alignment       =   1  'Right Justify
                     BackColor       =   &H00E2E9E9&
                     Caption         =   " «—ÌŒ ‰Â«Ì… «· Œ’Ì’ Â‹‹ "
                     Height          =   330
                     Index           =   14
                     Left            =   2640
                     RightToLeft     =   -1  'True
                     TabIndex        =   251
                     Top             =   870
                     Width           =   1425
                  End
                  Begin VB.Label lbl 
                     Alignment       =   1  'Right Justify
                     BackColor       =   &H00E2E9E9&
                     Caption         =   "—ﬁ„ ⁄ﬁœ «· Œ’Ì’"
                     Height          =   390
                     Index           =   11
                     Left            =   2715
                     RightToLeft     =   -1  'True
                     TabIndex        =   250
                     Top             =   120
                     Width           =   1350
                  End
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "Ê’› „Œ ’—"
                  Height          =   420
                  Index           =   124
                  Left            =   3465
                  RightToLeft     =   -1  'True
                  TabIndex        =   217
                  Top             =   6450
                  Width           =   600
               End
               Begin VB.Image Image2 
                  Height          =   2295
                  Left            =   75
                  Picture         =   "FrmCars.frx":0024
                  Stretch         =   -1  'True
                  Top             =   630
                  Visible         =   0   'False
                  Width           =   2415
               End
               Begin VB.Image Image1 
                  Height          =   2505
                  Left            =   75
                  Picture         =   "FrmCars.frx":3BD0
                  Stretch         =   -1  'True
                  Top             =   645
                  Visible         =   0   'False
                  Width           =   2415
               End
            End
         End
         Begin C1SizerLibCtl.C1Elastic C1Elastic13 
            Height          =   8055
            Left            =   12570
            TabIndex        =   266
            TabStop         =   0   'False
            Top             =   45
            Width           =   11535
            _cx             =   20346
            _cy             =   14208
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
            CaptionPos      =   6
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
            Begin C1SizerLibCtl.C1Elastic C1Elastic17 
               Height          =   8085
               Left            =   0
               TabIndex        =   267
               TabStop         =   0   'False
               Top             =   0
               Width           =   11535
               _cx             =   20346
               _cy             =   14261
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
               Begin VSFlex8Ctl.VSFlexGrid Fg 
                  Height          =   6525
                  Left            =   135
                  TabIndex        =   268
                  Top             =   525
                  Width           =   11280
                  _cx             =   19897
                  _cy             =   11509
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
                  BackColor       =   16777215
                  ForeColor       =   -2147483640
                  BackColorFixed  =   14871017
                  ForeColorFixed  =   -2147483630
                  BackColorSel    =   -2147483635
                  ForeColorSel    =   -2147483634
                  BackColorBkg    =   -2147483636
                  BackColorAlternate=   16776960
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
                  Cols            =   8
                  FixedRows       =   1
                  FixedCols       =   1
                  RowHeightMin    =   320
                  RowHeightMax    =   0
                  ColWidthMin     =   0
                  ColWidthMax     =   0
                  ExtendLastCol   =   0   'False
                  FormatString    =   $"FrmCars.frx":60EE
                  ScrollTrack     =   0   'False
                  ScrollBars      =   3
                  ScrollTips      =   0   'False
                  MergeCells      =   0
                  MergeCompare    =   0
                  AutoResize      =   0   'False
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
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "0"
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   9.75
                     Charset         =   178
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H00C00000&
                  Height          =   240
                  Index           =   46
                  Left            =   2895
                  RightToLeft     =   -1  'True
                  TabIndex        =   271
                  Top             =   7395
                  Width           =   6630
               End
               Begin VB.Label lbl 
                  Alignment       =   2  'Center
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "«·«Ã„«·Ì"
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   9.75
                     Charset         =   178
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H00C00000&
                  Height          =   360
                  Index           =   45
                  Left            =   8730
                  RightToLeft     =   -1  'True
                  TabIndex        =   270
                  Top             =   7395
                  Width           =   2730
               End
               Begin VB.Label lbl 
                  Alignment       =   2  'Center
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "„’—Ê›«  «·ÕÊ«œÀ"
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   12
                     Charset         =   178
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H00C00000&
                  Height          =   360
                  Index           =   44
                  Left            =   4455
                  RightToLeft     =   -1  'True
                  TabIndex        =   269
                  Top             =   90
                  Width           =   2625
               End
            End
         End
         Begin C1SizerLibCtl.C1Elastic C1Elastic14 
            Height          =   8055
            Left            =   12870
            TabIndex        =   276
            TabStop         =   0   'False
            Top             =   45
            Width           =   11535
            _cx             =   20346
            _cy             =   14208
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
            Begin C1SizerLibCtl.C1Elastic C1Elastic15 
               Height          =   6030
               Left            =   -120
               TabIndex        =   278
               TabStop         =   0   'False
               Top             =   0
               Width           =   11535
               _cx             =   20346
               _cy             =   10636
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
               Begin VB.TextBox txtCode1 
                  Alignment       =   1  'Right Justify
                  Height          =   315
                  Left            =   0
                  MaxLength       =   50
                  RightToLeft     =   -1  'True
                  TabIndex        =   292
                  Top             =   -240
                  Visible         =   0   'False
                  Width           =   1080
               End
               Begin VB.TextBox TxtVlue 
                  Alignment       =   1  'Right Justify
                  Height          =   315
                  Left            =   3720
                  RightToLeft     =   -1  'True
                  TabIndex        =   287
                  Top             =   5055
                  Width           =   6345
               End
               Begin VB.TextBox TxtRemarks 
                  Alignment       =   1  'Right Justify
                  Height          =   435
                  Left            =   3720
                  MultiLine       =   -1  'True
                  RightToLeft     =   -1  'True
                  ScrollBars      =   2  'Vertical
                  TabIndex        =   286
                  Top             =   5505
                  Width           =   6345
               End
               Begin VSFlex8Ctl.VSFlexGrid VSFlexGrid1 
                  Height          =   3975
                  Left            =   150
                  TabIndex        =   279
                  Top             =   540
                  Width           =   12210
                  _cx             =   21537
                  _cy             =   7011
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
                  BackColor       =   16777215
                  ForeColor       =   -2147483640
                  BackColorFixed  =   14871017
                  ForeColorFixed  =   -2147483630
                  BackColorSel    =   -2147483635
                  ForeColorSel    =   -2147483634
                  BackColorBkg    =   -2147483636
                  BackColorAlternate=   16776960
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
                  ExtendLastCol   =   0   'False
                  FormatString    =   $"FrmCars.frx":622A
                  ScrollTrack     =   0   'False
                  ScrollBars      =   3
                  ScrollTips      =   0   'False
                  MergeCells      =   0
                  MergeCompare    =   0
                  AutoResize      =   0   'False
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
               Begin ImpulseButton.ISButton CmdAdd 
                  Height          =   270
                  Left            =   735
                  TabIndex        =   283
                  Top             =   4635
                  Width           =   1005
                  _ExtentX        =   1773
                  _ExtentY        =   476
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
                  ButtonImage     =   "FrmCars.frx":62F6
                  ColorButton     =   14871017
                  DrawFocusRectangle=   0   'False
               End
               Begin ImpulseButton.ISButton CmdDelete 
                  Height          =   300
                  Left            =   975
                  TabIndex        =   284
                  Top             =   5205
                  Width           =   765
                  _ExtentX        =   1349
                  _ExtentY        =   529
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
                  ButtonImage     =   "FrmCars.frx":6690
                  ColorButton     =   14871017
                  DrawFocusRectangle=   0   'False
               End
               Begin ImpulseButton.ISButton btnModify 
                  Height          =   240
                  Left            =   735
                  TabIndex        =   285
                  Top             =   4950
                  Width           =   1005
                  _ExtentX        =   1773
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
                  ButtonImage     =   "FrmCars.frx":CEF2
                  ColorButton     =   14871017
                  DrawFocusRectangle=   0   'False
               End
               Begin MSDataListLib.DataCombo DcbExpenSiv 
                  Height          =   315
                  Left            =   3720
                  TabIndex        =   288
                  Top             =   4620
                  Width           =   6345
                  _ExtentX        =   11192
                  _ExtentY        =   556
                  _Version        =   393216
                  Text            =   ""
                  RightToLeft     =   -1  'True
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "„·«ÕŸ« "
                  Height          =   330
                  Index           =   52
                  Left            =   9975
                  RightToLeft     =   -1  'True
                  TabIndex        =   291
                  Top             =   5580
                  Visible         =   0   'False
                  Width           =   1170
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "«·ﬁÌ„…"
                  Height          =   330
                  Index           =   51
                  Left            =   9855
                  RightToLeft     =   -1  'True
                  TabIndex        =   290
                  Top             =   5025
                  Width           =   1170
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "«·„’—Ê›"
                  Height          =   330
                  Index           =   49
                  Left            =   9975
                  RightToLeft     =   -1  'True
                  TabIndex        =   289
                  Top             =   4575
                  Width           =   1170
               End
               Begin VB.Label lbl 
                  Alignment       =   2  'Center
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "«·„’—Ê›« "
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   12
                     Charset         =   178
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H00C00000&
                  Height          =   480
                  Index           =   48
                  Left            =   6225
                  RightToLeft     =   -1  'True
                  TabIndex        =   282
                  Top             =   165
                  Width           =   4560
               End
               Begin VB.Label TotalValue 
                  Alignment       =   2  'Center
                  BackColor       =   &H00E2E9E9&
                  BackStyle       =   0  'Transparent
                  Caption         =   "0"
                  Height          =   255
                  Left            =   165
                  RightToLeft     =   -1  'True
                  TabIndex        =   281
                  Top             =   5550
                  Width           =   2430
               End
               Begin VB.Label Label1 
                  Alignment       =   2  'Center
                  BackColor       =   &H00E2E9E9&
                  BackStyle       =   0  'Transparent
                  Caption         =   "«·«Ã„«·Ì «·⁄«„"
                  Height          =   255
                  Index           =   48
                  Left            =   2595
                  RightToLeft     =   -1  'True
                  TabIndex        =   280
                  Top             =   5550
                  Width           =   975
               End
            End
            Begin ImpulseButton.ISButton CmdPrint 
               Height          =   420
               Left            =   3690
               TabIndex        =   293
               Top             =   6660
               Width           =   2490
               _ExtentX        =   4392
               _ExtentY        =   741
               ButtonPositionImage=   1
               Caption         =   "ÿ»«⁄… «·„’—Ê›« "
               BackColor       =   14871017
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
            Begin VB.Label lbl 
               Alignment       =   2  'Center
               BackColor       =   &H00E2E9E9&
               Caption         =   "„’—Ê›«  «·ÕÊ«œÀ"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   12
                  Charset         =   178
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00C00000&
               Height          =   405
               Index           =   50
               Left            =   3915
               RightToLeft     =   -1  'True
               TabIndex        =   277
               Top             =   90
               Width           =   3240
            End
         End
         Begin C1SizerLibCtl.C1Elastic C1Elastic16 
            Height          =   8055
            Left            =   13170
            TabIndex        =   295
            TabStop         =   0   'False
            Top             =   45
            Width           =   11535
            _cx             =   20346
            _cy             =   14208
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
            Begin VB.CommandButton printPartsRep 
               Caption         =   "ÿ»«⁄… »Ì«‰ »«·„·Õﬁ« "
               Height          =   510
               Left            =   825
               RightToLeft     =   -1  'True
               TabIndex        =   301
               Top             =   7245
               Width           =   1740
            End
            Begin VSFlex8UCtl.VSFlexGrid VSFlexGrid13 
               Height          =   6105
               Left            =   225
               TabIndex        =   296
               Top             =   975
               Width           =   11160
               _cx             =   19685
               _cy             =   10769
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
               BackColor       =   16777215
               ForeColor       =   -2147483640
               BackColorFixed  =   14871017
               ForeColorFixed  =   -2147483630
               BackColorSel    =   -2147483635
               ForeColorSel    =   -2147483634
               BackColorBkg    =   -2147483633
               BackColorAlternate=   16777088
               GridColor       =   -2147483633
               GridColorFixed  =   -2147483633
               TreeColor       =   -2147483632
               FloodColor      =   192
               SheetBorder     =   -2147483633
               FocusRect       =   1
               HighLight       =   1
               AllowSelection  =   -1  'True
               AllowBigSelection=   -1  'True
               AllowUserResizing=   1
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
               FormatString    =   $"FrmCars.frx":D28C
               ScrollTrack     =   -1  'True
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
            Begin MSDataListLib.DataCombo PartDC 
               Height          =   315
               Left            =   4680
               TabIndex        =   297
               Top             =   360
               Width           =   5580
               _ExtentX        =   9843
               _ExtentY        =   556
               _Version        =   393216
               Text            =   ""
               RightToLeft     =   -1  'True
            End
            Begin ImpulseButton.ISButton ISButton2 
               Height          =   405
               Left            =   3465
               TabIndex        =   298
               Top             =   240
               Width           =   825
               _ExtentX        =   1455
               _ExtentY        =   714
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
               ButtonImage     =   "FrmCars.frx":D33A
               ColorButton     =   14871017
               DrawFocusRectangle=   0   'False
            End
            Begin ImpulseButton.ISButton ISButton1 
               Height          =   375
               Left            =   2415
               TabIndex        =   299
               Top             =   240
               Width           =   750
               _ExtentX        =   1323
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
               ButtonImage     =   "FrmCars.frx":D6D4
               ColorButton     =   14871017
               DrawFocusRectangle=   0   'False
            End
            Begin ImpulseButton.ISButton ISButton3 
               Height          =   435
               Left            =   1350
               TabIndex        =   300
               Top             =   240
               Width           =   765
               _ExtentX        =   1349
               _ExtentY        =   767
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
               ButtonImage     =   "FrmCars.frx":DA6E
               ColorButton     =   14871017
               DrawFocusRectangle=   0   'False
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "«”„ «·„·Õﬁ"
               Height          =   345
               Index           =   53
               Left            =   10335
               RightToLeft     =   -1  'True
               TabIndex        =   294
               Top             =   360
               Width           =   1050
            End
         End
      End
      Begin C1SizerLibCtl.C1Elastic Ele 
         Height          =   765
         Index           =   0
         Left            =   0
         TabIndex        =   218
         TabStop         =   0   'False
         Top             =   0
         Width           =   11685
         _cx             =   20611
         _cy             =   1349
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
         Caption         =   "»Ì«‰«  «·„⁄œ« /«·„—ﬂ»« "
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
         Begin ImpulseButton.ISButton XPBtnMove 
            Height          =   375
            Index           =   0
            Left            =   1155
            TabIndex        =   219
            Top             =   120
            Width           =   495
            _ExtentX        =   873
            _ExtentY        =   661
            ButtonStyle     =   1
            ButtonPositionImage=   4
            Caption         =   ""
            BackColor       =   -2147483634
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   178
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ButtonImage     =   "FrmCars.frx":142D0
            ColorButton     =   -2147483634
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
            Height          =   375
            Index           =   2
            Left            =   90
            TabIndex        =   220
            Top             =   120
            Width           =   495
            _ExtentX        =   873
            _ExtentY        =   661
            ButtonStyle     =   1
            ButtonPositionImage=   4
            Caption         =   ""
            BackColor       =   -2147483634
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   178
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ButtonImage     =   "FrmCars.frx":1466A
            ColorButton     =   -2147483634
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
            Height          =   375
            Index           =   1
            Left            =   1680
            TabIndex        =   221
            Top             =   120
            Width           =   495
            _ExtentX        =   873
            _ExtentY        =   661
            ButtonStyle     =   1
            ButtonPositionImage=   4
            Caption         =   ""
            BackColor       =   -2147483634
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   178
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ButtonImage     =   "FrmCars.frx":14A04
            ColorButton     =   -2147483634
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
            Height          =   375
            Index           =   3
            Left            =   615
            TabIndex        =   222
            Top             =   120
            Width           =   495
            _ExtentX        =   873
            _ExtentY        =   661
            ButtonStyle     =   1
            ButtonPositionImage=   4
            Caption         =   ""
            BackColor       =   -2147483634
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   178
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ButtonImage     =   "FrmCars.frx":14D9E
            ColorButton     =   -2147483634
            ColorHighlight  =   4194304
            ColorHoverText  =   16777215
            ColorShadow     =   -2147483631
            ColorOutline    =   -2147483631
            DrawFocusRectangle=   0   'False
            DisabledImageStyle=   1
            ColorToggledHoverText=   16777215
            ColorTextShadow =   16777215
         End
         Begin VB.Label Label6 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
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
            Height          =   495
            Left            =   3360
            RightToLeft     =   -1  'True
            TabIndex        =   274
            Top             =   120
            Width           =   3615
         End
      End
      Begin C1SizerLibCtl.C1Elastic C1Elastic5 
         Height          =   705
         Left            =   0
         TabIndex        =   223
         TabStop         =   0   'False
         Top             =   9240
         Width           =   11640
         _cx             =   20532
         _cy             =   1244
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
         Begin ImpulseButton.ISButton Cmd 
            Height          =   660
            Index           =   0
            Left            =   10635
            TabIndex        =   224
            Top             =   60
            Width           =   990
            _ExtentX        =   1746
            _ExtentY        =   1164
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
            ColorButton     =   14871017
            ColorHighlight  =   16777215
            ColorHoverText  =   16711680
            ColorShadow     =   -2147483637
            ColorOutline    =   0
            DrawFocusRectangle=   0   'False
            DisabledImageExtraction=   0
            ColorToggledHoverText=   16711680
            ColorTextShadow =   -2147483637
         End
         Begin ImpulseButton.ISButton Cmd 
            Height          =   660
            Index           =   1
            Left            =   9180
            TabIndex        =   225
            Top             =   60
            Width           =   1170
            _ExtentX        =   2064
            _ExtentY        =   1164
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
            ColorButton     =   14871017
            ColorHighlight  =   16777215
            ColorHoverText  =   16711680
            ColorShadow     =   -2147483637
            ColorOutline    =   0
            DrawFocusRectangle=   0   'False
            ColorToggledHoverText=   16711680
            ColorTextShadow =   -2147483637
         End
         Begin ImpulseButton.ISButton Cmd 
            Height          =   780
            Index           =   2
            Left            =   8085
            TabIndex        =   226
            Top             =   -60
            Width           =   945
            _ExtentX        =   1667
            _ExtentY        =   1376
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
            ColorButton     =   14871017
            ColorHighlight  =   16777215
            ColorHoverText  =   16711680
            ColorShadow     =   -2147483637
            ColorOutline    =   0
            DrawFocusRectangle=   0   'False
            ColorToggledHoverText=   16711680
            ColorTextShadow =   -2147483637
         End
         Begin ImpulseButton.ISButton Cmd 
            Height          =   660
            Index           =   3
            Left            =   6750
            TabIndex        =   227
            Top             =   60
            Width           =   1125
            _ExtentX        =   1984
            _ExtentY        =   1164
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
            ColorButton     =   14871017
            ColorHighlight  =   16777215
            ColorHoverText  =   16711680
            ColorShadow     =   -2147483637
            ColorOutline    =   0
            DrawFocusRectangle=   0   'False
            ColorToggledHoverText=   16711680
            ColorTextShadow =   -2147483637
         End
         Begin ImpulseButton.ISButton Cmd 
            Height          =   660
            Index           =   4
            Left            =   5760
            TabIndex        =   228
            Top             =   60
            Width           =   975
            _ExtentX        =   1720
            _ExtentY        =   1164
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
            ColorButton     =   14871017
            ColorHighlight  =   16777215
            ColorHoverText  =   16711680
            ColorShadow     =   -2147483637
            ColorOutline    =   0
            DrawFocusRectangle=   0   'False
            ColorToggledHoverText=   16711680
            ColorTextShadow =   -2147483637
         End
         Begin ImpulseButton.ISButton Cmd 
            Height          =   660
            Index           =   6
            Left            =   2805
            TabIndex        =   229
            Top             =   60
            Width           =   975
            _ExtentX        =   1720
            _ExtentY        =   1164
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
            ColorButton     =   14871017
            ColorHighlight  =   16777215
            ColorHoverText  =   16711680
            ColorShadow     =   -2147483637
            ColorOutline    =   0
            DrawFocusRectangle=   0   'False
            ColorToggledHoverText=   16711680
            ColorTextShadow =   -2147483637
         End
         Begin ImpulseButton.ISButton Cmd 
            Height          =   660
            Index           =   5
            Left            =   -105
            TabIndex        =   230
            Top             =   165
            Width           =   705
            _ExtentX        =   1244
            _ExtentY        =   1164
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
         Begin ImpulseButton.ISButton Cmd 
            Height          =   660
            Index           =   10
            Left            =   1485
            TabIndex        =   231
            Top             =   135
            Width           =   660
            _ExtentX        =   1164
            _ExtentY        =   1164
            ButtonPositionImage=   1
            Caption         =   "ŒÿÂ «·’Ì«‰Â"
            BackColor       =   14871017
            ForeColor       =   16711680
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
            ColorToggledText=   16711680
            ColorToggledHoverText=   16711680
            ColorTextShadow =   -2147483637
         End
         Begin ImpulseButton.ISButton Cmd 
            Height          =   660
            Index           =   12
            Left            =   1965
            TabIndex        =   232
            Top             =   60
            Width           =   840
            _ExtentX        =   1482
            _ExtentY        =   1164
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
            ColorButton     =   14871017
            ColorHighlight  =   16777215
            ColorHoverText  =   16711680
            ColorShadow     =   -2147483637
            ColorOutline    =   0
            DrawFocusRectangle=   0   'False
            ColorToggledHoverText=   16711680
            ColorTextShadow =   -2147483637
         End
         Begin ImpulseButton.ISButton Cmd 
            Height          =   780
            Index           =   17
            Left            =   4995
            TabIndex        =   233
            Top             =   0
            Width           =   690
            _ExtentX        =   1217
            _ExtentY        =   1376
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
            ColorButton     =   14871017
            ColorHighlight  =   16777215
            ColorHoverText  =   16711680
            ColorShadow     =   -2147483637
            ColorOutline    =   0
            DrawFocusRectangle=   0   'False
            ColorToggledHoverText=   16711680
            ColorTextShadow =   -2147483637
         End
         Begin ImpulseButton.ISButton Cmd 
            Height          =   660
            Index           =   18
            Left            =   3750
            TabIndex        =   263
            Top             =   60
            Width           =   1200
            _ExtentX        =   2117
            _ExtentY        =   1164
            ButtonPositionImage=   1
            Caption         =   " ﬁ—Ì— ‘«„·"
            BackColor       =   14871017
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
         Begin ImpulseButton.ISButton Cmd 
            Height          =   660
            Index           =   19
            Left            =   525
            TabIndex        =   264
            Top             =   135
            Width           =   945
            _ExtentX        =   1667
            _ExtentY        =   1164
            ButtonPositionImage=   1
            Caption         =   "„’—Ê›«  «·„⁄œÂ/«·”Ì«—…"
            BackColor       =   14871017
            ForeColor       =   16711680
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
            ColorToggledText=   16711680
            ColorToggledHoverText=   16711680
            ColorTextShadow =   -2147483637
         End
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00E2E9E9&
      Caption         =   "»Ì«‰«  „Ã„Ê⁄Â «·«’·"
      Enabled         =   0   'False
      Height          =   2895
      Left            =   18360
      RightToLeft     =   -1  'True
      TabIndex        =   146
      Top             =   3840
      Width           =   6375
      Begin VB.TextBox TXtPercentage1 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   2760
         MaxLength       =   50
         RightToLeft     =   -1  'True
         TabIndex        =   156
         Top             =   360
         Width           =   1245
      End
      Begin VB.TextBox txtPercentage2 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   2760
         MaxLength       =   50
         RightToLeft     =   -1  'True
         TabIndex        =   155
         Top             =   720
         Width           =   1245
      End
      Begin VB.TextBox TXT40 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   120
         MaxLength       =   50
         RightToLeft     =   -1  'True
         TabIndex        =   154
         Top             =   2520
         Width           =   3885
      End
      Begin VB.TextBox TXT31 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   120
         MaxLength       =   50
         RightToLeft     =   -1  'True
         TabIndex        =   153
         Top             =   2160
         Width           =   3885
      End
      Begin VB.TextBox TXT25 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   120
         MaxLength       =   50
         RightToLeft     =   -1  'True
         TabIndex        =   152
         Top             =   1800
         Width           =   3885
      End
      Begin VB.TextBox TXT26 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   120
         MaxLength       =   50
         RightToLeft     =   -1  'True
         TabIndex        =   151
         Top             =   1440
         Width           =   3885
      End
      Begin VB.TextBox TXT24 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   120
         MaxLength       =   50
         RightToLeft     =   -1  'True
         TabIndex        =   150
         Top             =   1080
         Width           =   3885
      End
      Begin VB.OptionButton Opt 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "·Â «Â·«ﬂ"
         Enabled         =   0   'False
         Height          =   225
         Index           =   0
         Left            =   2370
         RightToLeft     =   -1  'True
         TabIndex        =   149
         Top             =   120
         Width           =   1815
      End
      Begin VB.OptionButton Opt 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "·Ì” ·Â «Â·«ﬂ"
         Enabled         =   0   'False
         Height          =   255
         Index           =   1
         Left            =   -120
         RightToLeft     =   -1  'True
         TabIndex        =   148
         Top             =   120
         Value           =   -1  'True
         Width           =   2295
      End
      Begin VB.TextBox TxtAge 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   120
         MaxLength       =   50
         RightToLeft     =   -1  'True
         TabIndex        =   147
         Top             =   360
         Width           =   525
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "‰”»… «·«Â·«ﬂ ⁄‰œ «·«Ìﬁ«›"
         Height          =   255
         Index           =   110
         Left            =   4080
         RightToLeft     =   -1  'True
         TabIndex        =   164
         Top             =   720
         Width           =   1995
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "‰”»… «·«Â·«ﬂ"
         Height          =   255
         Index           =   109
         Left            =   4080
         RightToLeft     =   -1  'True
         TabIndex        =   163
         Top             =   360
         Width           =   1995
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   " Õ”«»   Œ”«∆— »Ì⁄"
         Height          =   255
         Index           =   115
         Left            =   4080
         RightToLeft     =   -1  'True
         TabIndex        =   162
         Top             =   2520
         Width           =   2115
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   " Õ”«»   «—»«Õ »Ì⁄"
         Height          =   255
         Index           =   114
         Left            =   4080
         RightToLeft     =   -1  'True
         TabIndex        =   161
         Top             =   2160
         Width           =   2115
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   " Õ”«»    „’—Ê›«  «·«Â·«ﬂ"
         Height          =   255
         Index           =   113
         Left            =   4080
         RightToLeft     =   -1  'True
         TabIndex        =   160
         Top             =   1800
         Width           =   2115
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   " Õ”«» „Ã„⁄ «·«Â·«ﬂ"
         Height          =   255
         Index           =   112
         Left            =   4080
         RightToLeft     =   -1  'True
         TabIndex        =   159
         Top             =   1440
         Width           =   2115
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   " Õ”«» «·«’·  »«·„Ì“«‰Ì…"
         Height          =   255
         Index           =   111
         Left            =   4200
         RightToLeft     =   -1  'True
         TabIndex        =   158
         Top             =   1080
         Width           =   1995
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "«·⁄„— «·«› —«÷Ì ··«’· »«·‘Â—"
         Height          =   255
         Index           =   9
         Left            =   600
         RightToLeft     =   -1  'True
         TabIndex        =   157
         Top             =   360
         Width           =   2115
      End
   End
   Begin VB.TextBox TxtName 
      Alignment       =   1  'Right Justify
      Height          =   315
      Left            =   17760
      MaxLength       =   50
      RightToLeft     =   -1  'True
      TabIndex        =   145
      Top             =   1800
      Visible         =   0   'False
      Width           =   4605
   End
   Begin VB.Frame Frame4 
      Height          =   5295
      Left            =   16920
      RightToLeft     =   -1  'True
      TabIndex        =   140
      Top             =   0
      Width           =   3255
      Begin ImpulseButton.ISButton Cmd 
         Height          =   735
         Index           =   13
         Left            =   0
         TabIndex        =   141
         Top             =   120
         Width           =   2415
         _ExtentX        =   4260
         _ExtentY        =   1296
         ButtonPositionImage=   1
         Caption         =   "„Œ“‰ «·Ê—‘…"
         BackColor       =   255
         ForeColor       =   16777215
         FontSize        =   13.5
         FontBold        =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ButtonImage     =   "FrmCars.frx":15138
         ColorButton     =   255
         ColorHighlight  =   16777215
         ColorHoverText  =   16711680
         ColorShadow     =   -2147483637
         ColorOutline    =   16777215
         DrawFocusRectangle=   0   'False
         DisabledImageExtraction=   0
         ColorToggledText=   -2147483637
         ColorToggledHoverText=   16711680
         ColorTextShadow =   -2147483637
      End
      Begin ImpulseButton.ISButton Cmd 
         Height          =   1095
         Index           =   14
         Left            =   0
         TabIndex        =   142
         Top             =   1920
         Width           =   3255
         _ExtentX        =   5741
         _ExtentY        =   1931
         ButtonPositionImage=   1
         Caption         =   "«· ÕÊÌ· »Ì‰ «·«ﬁ”«„"
         BackColor       =   255
         ForeColor       =   16777215
         FontSize        =   13.5
         FontBold        =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ButtonImage     =   "FrmCars.frx":1667A
         ColorButton     =   255
         ColorHighlight  =   16777215
         ColorHoverText  =   16711680
         ColorShadow     =   -2147483637
         ColorOutline    =   16777215
         DrawFocusRectangle=   0   'False
         DisabledImageExtraction=   0
         ColorToggledText=   -2147483637
         ColorToggledHoverText=   16711680
         ColorTextShadow =   -2147483637
      End
      Begin ImpulseButton.ISButton Cmd 
         Height          =   855
         Index           =   15
         Left            =   0
         TabIndex        =   143
         Top             =   840
         Width           =   2535
         _ExtentX        =   4471
         _ExtentY        =   1508
         ButtonPositionImage=   1
         Caption         =   "’—› ﬁÿ⁄ €Ì«— "
         BackColor       =   255
         ForeColor       =   16777215
         FontSize        =   13.5
         FontBold        =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ButtonImage     =   "FrmCars.frx":17B61
         ColorButton     =   255
         ColorHighlight  =   16777215
         ColorHoverText  =   16711680
         ColorShadow     =   -2147483637
         ColorOutline    =   16777215
         DrawFocusRectangle=   0   'False
         DisabledImageExtraction=   0
         ColorToggledText=   -2147483637
         ColorToggledHoverText=   16711680
         ColorTextShadow =   -2147483637
      End
      Begin ImpulseButton.ISButton Cmd 
         Height          =   975
         Index           =   16
         Left            =   0
         TabIndex        =   144
         Top             =   3000
         Width           =   3255
         _ExtentX        =   5741
         _ExtentY        =   1720
         ButtonPositionImage=   1
         Caption         =   " ”·Ì„ «·„⁄œ« /«·”Ì«—« "
         BackColor       =   255
         ForeColor       =   16777215
         FontSize        =   13.5
         FontBold        =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ButtonImage     =   "FrmCars.frx":18CDE
         ColorButton     =   255
         ColorHighlight  =   16777215
         ColorHoverText  =   16711680
         ColorShadow     =   -2147483637
         ColorOutline    =   16777215
         DrawFocusRectangle=   0   'False
         DisabledImageExtraction=   0
         ColorToggledText=   -2147483637
         ColorToggledHoverText=   16711680
         ColorTextShadow =   -2147483637
      End
   End
   Begin VB.TextBox txtopening_balance_voucher_id 
      Alignment       =   1  'Right Justify
      Height          =   495
      Left            =   18960
      RightToLeft     =   -1  'True
      TabIndex        =   136
      Top             =   6960
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.TextBox txtNoteSerial1 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   19200
      RightToLeft     =   -1  'True
      TabIndex        =   135
      Top             =   6720
      Visible         =   0   'False
      Width           =   2295
   End
   Begin VB.TextBox txtNoteID1 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   19200
      RightToLeft     =   -1  'True
      TabIndex        =   134
      Top             =   6600
      Visible         =   0   'False
      Width           =   2295
   End
   Begin VB.TextBox Text2 
      Alignment       =   1  'Right Justify
      Enabled         =   0   'False
      Height          =   315
      Left            =   18840
      MaxLength       =   50
      RightToLeft     =   -1  'True
      TabIndex        =   133
      Top             =   5880
      Width           =   2325
   End
   Begin VB.TextBox TxtSalePrice 
      Alignment       =   1  'Right Justify
      Enabled         =   0   'False
      Height          =   315
      Left            =   18360
      MaxLength       =   50
      RightToLeft     =   -1  'True
      TabIndex        =   132
      Top             =   5400
      Width           =   2325
   End
   Begin VB.TextBox txtNoteID 
      Height          =   285
      Left            =   18720
      RightToLeft     =   -1  'True
      TabIndex        =   127
      Top             =   7200
      Visible         =   0   'False
      Width           =   2295
   End
   Begin VB.TextBox txtNoteSerial 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   18840
      RightToLeft     =   -1  'True
      TabIndex        =   126
      Top             =   6120
      Visible         =   0   'False
      Width           =   2295
   End
   Begin VB.Frame Frame3 
      Height          =   435
      Left            =   15240
      RightToLeft     =   -1  'True
      TabIndex        =   123
      Top             =   75
      Width           =   2175
      Begin VB.OptionButton Option1 
         Alignment       =   1  'Right Justify
         Caption         =   "ÃœÌœ"
         Height          =   195
         Left            =   960
         RightToLeft     =   -1  'True
         TabIndex        =   125
         Top             =   120
         Width           =   915
      End
      Begin VB.OptionButton Option2 
         Alignment       =   1  'Right Justify
         Caption         =   "«›  «ÕÌ"
         Height          =   195
         Left            =   120
         RightToLeft     =   -1  'True
         TabIndex        =   124
         Top             =   120
         Width           =   975
      End
   End
   Begin VB.TextBox txtPurchaseBillId 
      Alignment       =   1  'Right Justify
      Height          =   315
      Left            =   15360
      MaxLength       =   50
      RightToLeft     =   -1  'True
      TabIndex        =   120
      Top             =   360
      Width           =   1245
   End
   Begin VB.TextBox TxtKhordaPrice 
      Alignment       =   2  'Center
      Enabled         =   0   'False
      Height          =   315
      Left            =   15240
      MaxLength       =   50
      RightToLeft     =   -1  'True
      TabIndex        =   118
      Top             =   720
      Width           =   2325
   End
   Begin VB.TextBox TxtCurrentValue 
      Alignment       =   2  'Center
      Enabled         =   0   'False
      Height          =   315
      Left            =   15360
      MaxLength       =   50
      RightToLeft     =   -1  'True
      TabIndex        =   117
      Top             =   600
      Width           =   1605
   End
   Begin VB.Frame Frame2 
      Caption         =   "»Ì«‰«  «·«Â·«ﬂ"
      Height          =   2415
      Left            =   480
      RightToLeft     =   -1  'True
      TabIndex        =   116
      Top             =   11400
      Width           =   8535
   End
   Begin VB.TextBox txtinstallDo 
      Alignment       =   1  'Right Justify
      Enabled         =   0   'False
      Height          =   315
      Left            =   18720
      MaxLength       =   50
      RightToLeft     =   -1  'True
      TabIndex        =   114
      Top             =   3360
      Width           =   1605
   End
   Begin VB.TextBox txtinstallmentresult 
      Alignment       =   1  'Right Justify
      Enabled         =   0   'False
      Height          =   315
      Left            =   15960
      MaxLength       =   50
      RightToLeft     =   -1  'True
      TabIndex        =   112
      Top             =   720
      Width           =   2325
   End
   Begin VB.ComboBox cStatus 
      Height          =   288
      ItemData        =   "FrmCars.frx":19B39
      Left            =   15120
      List            =   "FrmCars.frx":19B49
      RightToLeft     =   -1  'True
      TabIndex        =   108
      Top             =   360
      Width           =   2415
   End
   Begin VB.ComboBox CBoDepreciation_Type_id 
      Enabled         =   0   'False
      Height          =   288
      ItemData        =   "FrmCars.frx":19B89
      Left            =   14400
      List            =   "FrmCars.frx":19B93
      TabIndex        =   107
      Top             =   360
      Width           =   1575
   End
   Begin VB.TextBox TxtnoOfInst 
      Alignment       =   1  'Right Justify
      Enabled         =   0   'False
      Height          =   315
      Left            =   17520
      MaxLength       =   50
      RightToLeft     =   -1  'True
      TabIndex        =   104
      Top             =   3360
      Width           =   1605
   End
   Begin VB.TextBox txtinstallValue 
      Alignment       =   1  'Right Justify
      Enabled         =   0   'False
      Height          =   315
      Left            =   15840
      MaxLength       =   50
      RightToLeft     =   -1  'True
      TabIndex        =   102
      Top             =   600
      Width           =   2325
   End
   Begin VB.TextBox TxtAccDepreciation 
      Alignment       =   1  'Right Justify
      Enabled         =   0   'False
      Height          =   315
      Left            =   15840
      MaxLength       =   50
      RightToLeft     =   -1  'True
      TabIndex        =   100
      Top             =   480
      Width           =   1605
   End
   Begin VB.TextBox XPTxtID 
      Height          =   285
      Left            =   6960
      TabIndex        =   93
      Text            =   "Text1"
      Top             =   11880
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.TextBox TxtPurchasePrice 
      Alignment       =   2  'Center
      Height          =   315
      Left            =   15240
      MaxLength       =   50
      RightToLeft     =   -1  'True
      TabIndex        =   41
      Top             =   720
      Width           =   1605
   End
   Begin VB.TextBox TxtModFlg 
      Alignment       =   1  'Right Justify
      Height          =   345
      Left            =   3960
      RightToLeft     =   -1  'True
      TabIndex        =   39
      Top             =   11880
      Visible         =   0   'False
      Width           =   855
   End
   Begin MSComCtl2.DTPicker XPDtbTrans 
      Height          =   348
      Left            =   9960
      TabIndex        =   94
      Top             =   12120
      Visible         =   0   'False
      Width           =   1332
      _ExtentX        =   2355
      _ExtentY        =   609
      _Version        =   393216
      Format          =   104792065
      CurrentDate     =   38784
   End
   Begin MSDataListLib.DataCombo DCboUserName 
      Height          =   312
      Left            =   120
      TabIndex        =   95
      Top             =   12240
      Width           =   1740
      _ExtentX        =   3069
      _ExtentY        =   556
      _Version        =   393216
      Enabled         =   0   'False
      Text            =   ""
      RightToLeft     =   -1  'True
   End
   Begin MSComCtl2.DTPicker DpLicenseExpireDate 
      Height          =   348
      Left            =   18360
      TabIndex        =   98
      Top             =   5040
      Width           =   1452
      _ExtentX        =   2566
      _ExtentY        =   609
      _Version        =   393216
      Format          =   104792065
      CurrentDate     =   38784
   End
   Begin MSComCtl2.DTPicker DPReceiveDate 
      Height          =   348
      Left            =   14880
      TabIndex        =   106
      Top             =   600
      Width           =   2172
      _ExtentX        =   3836
      _ExtentY        =   609
      _Version        =   393216
      Format          =   104792065
      CurrentDate     =   38784
   End
   Begin ImpulseButton.ISButton Cmd 
      Height          =   372
      Index           =   7
      Left            =   21720
      TabIndex        =   109
      Top             =   5640
      Width           =   1512
      _ExtentX        =   2672
      _ExtentY        =   661
      ButtonPositionImage=   1
      Caption         =   "«Ìﬁ«› «·«Â·«ﬂ"
      BackColor       =   14871017
      Enabled         =   0   'False
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
      DisabledImageExtraction=   0
      ColorToggledHoverText=   16711680
      ColorTextShadow =   -2147483637
   End
   Begin ImpulseButton.ISButton Cmd 
      Height          =   372
      Index           =   8
      Left            =   20160
      TabIndex        =   110
      Top             =   5640
      Width           =   1512
      _ExtentX        =   2672
      _ExtentY        =   661
      ButtonPositionImage=   1
      Caption         =   "≈⁄«œ…  ‘€Ì· «·«Â·«ﬂ"
      BackColor       =   14871017
      Enabled         =   0   'False
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
      DisabledImageExtraction=   0
      ColorToggledHoverText=   16711680
      ColorTextShadow =   -2147483637
   End
   Begin ImpulseButton.ISButton Cmd 
      Height          =   372
      Index           =   9
      Left            =   18600
      TabIndex        =   111
      Top             =   5640
      Width           =   1512
      _ExtentX        =   2672
      _ExtentY        =   661
      ButtonPositionImage=   1
      Caption         =   "«· Œ·’ „‰ «·«’·"
      BackColor       =   14871017
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
      DisabledImageExtraction=   0
      ColorToggledHoverText=   16711680
      ColorTextShadow =   -2147483637
   End
   Begin MSComCtl2.DTPicker DpTestExpireDate 
      Height          =   348
      Left            =   18120
      TabIndex        =   122
      Top             =   1560
      Width           =   2172
      _ExtentX        =   3836
      _ExtentY        =   609
      _Version        =   393216
      Format          =   104792065
      CurrentDate     =   38784
   End
   Begin ImpulseButton.ISButton Cmd 
      Height          =   372
      Index           =   11
      Left            =   18120
      TabIndex        =   128
      Top             =   5232
      Visible         =   0   'False
      Width           =   912
      _ExtentX        =   1614
      _ExtentY        =   661
      ButtonPositionImage=   1
      Caption         =   "ÿ»«⁄Â «·ﬁÌœ"
      BackColor       =   14871017
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
   Begin MSComCtl2.DTPicker DpInsuranceExpireDate 
      Height          =   348
      Left            =   15000
      TabIndex        =   137
      Top             =   600
      Width           =   2172
      _ExtentX        =   3836
      _ExtentY        =   609
      _Version        =   393216
      Format          =   104792065
      CurrentDate     =   38784
   End
   Begin VSFlex8Ctl.VSFlexGrid Grid 
      Height          =   1860
      Left            =   16920
      TabIndex        =   138
      Top             =   3960
      Width           =   10140
      _cx             =   17886
      _cy             =   3281
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
      SelectionMode   =   1
      GridLines       =   1
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   1
      Cols            =   21
      FixedRows       =   1
      FixedCols       =   2
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   -1  'True
      FormatString    =   $"FrmCars.frx":19BB6
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
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   "ŒÿÂ «·’Ì«‰Â"
      Height          =   312
      Index           =   4
      Left            =   17160
      RightToLeft     =   -1  'True
      TabIndex        =   139
      Top             =   4200
      Width           =   1248
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "«·„ﬂ”» «Ê «·Œ”«—…"
      Height          =   372
      Left            =   19440
      RightToLeft     =   -1  'True
      TabIndex        =   131
      Top             =   5760
      Width           =   852
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "”⁄— «·»Ì⁄"
      Height          =   252
      Index           =   0
      Left            =   14040
      RightToLeft     =   -1  'True
      TabIndex        =   130
      Top             =   5160
      Width           =   1092
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   "—ﬁ„ «·ﬁÌœ"
      Height          =   192
      Index           =   0
      Left            =   18480
      RightToLeft     =   -1  'True
      TabIndex        =   129
      Top             =   6120
      Visible         =   0   'False
      Width           =   1128
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   "—ﬁ„ ›« Ê—… «·‘—«¡"
      Height          =   255
      Index           =   116
      Left            =   2520
      RightToLeft     =   -1  'True
      TabIndex        =   121
      Top             =   360
      Width           =   1275
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   "ﬁÌ„… «·«’· ﬂŒ—œ…"
      Height          =   372
      Index           =   121
      Left            =   15240
      RightToLeft     =   -1  'True
      TabIndex        =   119
      Top             =   720
      Width           =   1272
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   "„‰›–"
      Height          =   252
      Index           =   130
      Left            =   18120
      RightToLeft     =   -1  'True
      TabIndex        =   115
      Top             =   3360
      Width           =   672
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   "„ »ﬁÏ"
      Height          =   252
      Index           =   123
      Left            =   15120
      RightToLeft     =   -1  'True
      TabIndex        =   113
      Top             =   600
      Width           =   1272
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   " «—ÌŒ «·«” ·«„"
      Height          =   372
      Index           =   119
      Left            =   18480
      RightToLeft     =   -1  'True
      TabIndex        =   105
      Top             =   1920
      Width           =   1152
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   "«Œ— ﬁ—«¡… ··⁄œ«œ"
      Height          =   252
      Index           =   108
      Left            =   17400
      RightToLeft     =   -1  'True
      TabIndex        =   103
      Top             =   3360
      Width           =   1392
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   "ﬁÌ„… ﬁ”ÿ «·«Â·«ﬂ"
      Height          =   252
      Index           =   122
      Left            =   16200
      RightToLeft     =   -1  'True
      TabIndex        =   101
      Top             =   600
      Width           =   1272
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   "„Ã„⁄ «·«Â·«ﬂ"
      Height          =   252
      Index           =   129
      Left            =   15120
      RightToLeft     =   -1  'True
      TabIndex        =   99
      Top             =   720
      Width           =   912
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   "Õ«·… «·«’·"
      Height          =   255
      Index           =   118
      Left            =   2520
      RightToLeft     =   -1  'True
      TabIndex        =   97
      Top             =   -120
      Width           =   1275
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   "Õ—— »Ê«”ÿ… : "
      Height          =   312
      Index           =   5
      Left            =   2160
      RightToLeft     =   -1  'True
      TabIndex        =   96
      Top             =   12240
      Width           =   912
   End
   Begin VB.Label Label2 
      Caption         =   "—ﬁ„ «·ﬁÌœ"
      Height          =   372
      Left            =   8280
      TabIndex        =   92
      Top             =   12000
      Visible         =   0   'False
      Width           =   972
   End
   Begin VB.Label LngDevID 
      Height          =   375
      Left            =   6960
      TabIndex        =   91
      Top             =   8040
      Visible         =   0   'False
      Width           =   975
   End
End
Attribute VB_Name = "FrmCars"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rs As ADODB.Recordset
Dim RSAss As New ADODB.Recordset
Dim FirstPeriodDateInthisYear  As Date
Dim TTP As clstooltip
Dim rs_CarExpenses As ADODB.Recordset
Dim Dcombos As New ClsDataCombos
Dim rs_CarParts As ADODB.Recordset
    
Function CheckCarAssest() As Boolean
Dim Rs5 As ADODB.Recordset
Set Rs5 = New ADODB.Recordset
Dim sql As String
sql = "Select id from  TblCarsData where id =" & val(XPTxtID.Text) & "and FlagFixedasset=1  "
Rs5.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
If Rs5.RecordCount > 0 Then
CheckCarAssest = True
Else
CheckCarAssest = False
End If
End Function
Sub SaveDriveAssest(Optional FexdID As Double = 0, Optional EmpID As Double = 0)
Dim sql As String
Dim StrSQL As String
Dim Msg As String
Dim ID As Double
Dim RsDetails As ADODB.Recordset
Dim Rs5 As ADODB.Recordset
Set Rs5 = New ADODB.Recordset
sql = "select * from TblEmpAsest where 1=-1 "
Rs5.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
  ID = CStr(new_id("TblEmpAsest", "EmpAsID", "", True))
  Rs5.AddNew
     
        Rs5("EmpAsID").value = ID
      Rs5("BranchID").value = IIf(Me.Dcbranch.BoundText = "", Null, Me.Dcbranch.BoundText)
        Rs5("RecordDate").value = Date
      Rs5("OperatorN").value = TxtOperatorN.Text
      Rs5("BoardNO").value = TxtBoardNO.Text
      
        Rs5("DriveDate").value = Date
        Rs5("PostedDate").value = Date
    Rs5("EmpAsestID").value = EmpID
        Rs5("AsID").value = FexdID
        Rs5("ISCar").value = 1
        Rs5("FlgCar").value = 1
        Rs5("CrsID").value = val(XPTxtID.Text)
       Rs5.update
sql = "Select * from TblEmpAsestDetails where 1=-1"
Set RsDetails = New ADODB.Recordset
RsDetails.Open sql, Cn, adOpenKeyset, adLockOptimistic, adCmdText
If SystemOptions.UserInterface = ArabicInterface Then
Msg = "„‰ „·› «·„⁄œ« /«·”Ì«—« "
Else
Msg = "From Cars File"
End If
RsDetails.AddNew
RsDetails("IDAseset").value = ID
RsDetails("EmpID").value = EmpID
RsDetails("Remarks").value = Msg
RsDetails("CrsID").value = val(XPTxtID.Text)
RsDetails("AsID").value = FexdID
RsDetails("Qunt").value = 1
RsDetails.update
End Sub
Sub SaveEmployee()
Dim sql As String
Dim StrSQL As String
Dim Rs5 As ADODB.Recordset
Set Rs5 = New ADODB.Recordset
Dim Emp_id As Double

If Me.TxtModFlg.Text = "N" Or val(DcEmployee.BoundText) = 0 Then
sql = "Select * from TblEmployee where 1=-1"
Rs5.Open sql, Cn, adOpenKeyset, adLockOptimistic, adCmdText
Emp_id = CStr(new_id("TblEmployee", "Emp_ID", "", True))
Rs5.AddNew
Else
Emp_id = val(DcEmployee.BoundText)
sql = "Select * from TblEmployee where Emp_ID=" & Emp_id & ""
Rs5.Open sql, Cn, adOpenKeyset, adLockOptimistic, adCmdText
End If
Rs5("Emp_ID").value = Emp_id
Rs5("CrsID").value = val(XPTxtID.Text)
Rs5("BranchId").value = val(Dcbranch.BoundText)
Rs5("FlagDriver").value = 1
Rs5("Emp_Name").value = TxtLeaderName.Text
Rs5("Emp_Name1").value = TxtLeaderName.Text
Rs5("Emp_Namee").value = TxtLeaderName.Text
Rs5("Emp_Namee1").value = TxtLeaderName.Text

Rs5.update
sql = "Update TblCarsData set Emp_id=" & Emp_id & "  where id =" & val(XPTxtID.Text) & ""
Cn.Execute sql
sql = "Update TblEmpAsestDetails set EmpID=" & Emp_id & "  where CrsID =" & val(XPTxtID.Text) & ""
Cn.Execute sql
sql = "Update TblEmpAsest set EmpAsestID=" & Emp_id & "  where CrsID =" & val(XPTxtID.Text) & ""
Cn.Execute sql
End Sub
Function chekEmoloyee(Optional emp_Name As String) As Boolean
Dim sql As String
Dim Rs7 As ADODB.Recordset
Set Rs7 = New ADODB.Recordset
sql = " select Emp_Name from TblEmployee where Emp_Name=N'" & emp_Name & "'  "
Rs7.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
If Rs7.RecordCount > 0 Then
chekEmoloyee = True
Else
chekEmoloyee = False
End If
End Function

Private Sub btnAddImage_Click()
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
If val(XPTxtID.Text) <> 0 Then
DBPix201.ImageSaveFile (App.path & "\" & SystemOptions.ImagesPath & "\cars\" & XPTxtID.Text & ".JPG")
End If
    'DBPix201.ImageSaveFile (App.path & "\" & SystemOptions.ImagesPath & "\" & xptxtid.text & ".JPG")

End Sub
Function GetEmpValue(Optional Acd As Double) As Double
Dim Rs3 As ADODB.Recordset
Set Rs3 = New ADODB.Recordset
Dim sql As String
sql = " SELECT     Valuee"
sql = sql & " From dbo.TblAccidentReportDet"
sql = sql & " Where (typ = 2) And (AccID = " & Acd & ")"
Rs3.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
If Rs3.RecordCount > 0 Then
GetEmpValue = IIf(IsNull(Rs3("Valuee").value), 0, Rs3("Valuee").value)
Else
GetEmpValue = 0
End If
End Function
Sub FillAcced()
 fg.Clear flexClearScrollable, flexClearEverything
 fg.Rows = fg.FixedRows + 1
 Dim i As Integer
Dim sql As String
Dim Rs4 As ADODB.Recordset
Set Rs4 = New ADODB.Recordset
sql = " SELECT     dbo.TblAccidentReport.BranchID, dbo.TblBranchesData.branch_name, dbo.TblBranchesData.branch_namee, dbo.TblAccidentReport.ID,"
sql = sql & "                       dbo.TblAccidentReport.AccTime , dbo.TblAccidentReport.AccDate, dbo.TblAccidentReport.PlateNo, dbo.TblAccidentReport.CompValue"
sql = sql & " FROM         dbo.TblAccidentReport LEFT OUTER JOIN"
sql = sql & "                       dbo.TblBranchesData ON dbo.TblAccidentReport.BranchID = dbo.TblBranchesData.branch_id"
sql = sql & " WHERE     (dbo.TblAccidentReport.PlateNo = N'" & TxtBoardNO.Text & "')"
Rs4.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
If Rs4.RecordCount > 0 Then
With fg
.Rows = Rs4.RecordCount + 1
Rs4.MoveFirst
For i = 1 To .Rows - 1
.TextMatrix(i, .ColIndex("Serial")) = i
.TextMatrix(i, .ColIndex("ID")) = IIf(IsNull(Rs4("ID").value), 0, Rs4("ID").value)
.TextMatrix(i, .ColIndex("AccTime")) = IIf(IsNull(Rs4("AccTime").value), "", Rs4("AccTime").value)
.TextMatrix(i, .ColIndex("AccDate")) = IIf(IsNull(Rs4("AccDate").value), "", Rs4("AccDate").value)
.TextMatrix(i, .ColIndex("CompValue")) = IIf(IsNull(Rs4("CompValue").value), 0, Rs4("CompValue").value)
.TextMatrix(i, .ColIndex("EmpValue")) = GetEmpValue(val(.TextMatrix(i, .ColIndex("ID"))))
If SystemOptions.UserInterface = ArabicInterface Then
.TextMatrix(i, .ColIndex("branch_name")) = IIf(IsNull(Rs4("branch_name").value), "", Rs4("branch_name").value)
Else
.TextMatrix(i, .ColIndex("branch_name")) = IIf(IsNull(Rs4("branch_namee").value), "", Rs4("branch_namee").value)
End If
Rs4.MoveNext
Next i
End With
End If
Relin
End Sub

Function CheckOrderMainte() As Boolean
Dim sql As String
Dim Rs3 As ADODB.Recordset
Set Rs3 = New ADODB.Recordset
sql = "SELECT  EquepID"
sql = sql & " FROM  dbo.TblOrderMaint"
sql = sql & " where EquepID=" & val(XPTxtID.Text) & ""
Rs3.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
If Rs3.RecordCount > 0 Then
CheckOrderMainte = True
Else
CheckOrderMainte = False
End If
End Function

Private Sub btnModify_Click()
Update_CarsExpens
End Sub
Private Sub Update_CarsExpens()
Dim BeginTrans As Boolean
  Dim StrSQL As String
  Dim Msg As String
  On Error GoTo errortrap
Dim str As String, sr As String
 If val(TxtId.Text) = 0 Then
 If SystemOptions.UserInterface = ArabicInterface Then
 MsgBox "Ì—ÃÏ Õ›Ÿ »Ì«‰«  «·„⁄œ… «Ê·«"
 Else
 MsgBox "Please Save Data"
 End If
 Exit Sub
 End If
 If val(DcbExpenSiv.BoundText) = 0 Or DcbExpenSiv.Text = "" Then
 If SystemOptions.UserInterface = ArabicInterface Then
        MsgBox "Ì—ÃÏ «Œ Ì«— «·„’—Ê›"
 Else
        MsgBox "Please Select Expenses"
 End If
 DcbExpenSiv.SetFocus
 Exit Sub
 End If
 
 If val(TxtVlue.Text) = 0 Then
 If SystemOptions.UserInterface = ArabicInterface Then
        MsgBox ("«œŒ· «·ﬁÌ„…")
 Else
        MsgBox ("Enter Value ")
 End If
 TxtVlue.SetFocus
 Exit Sub
 End If
 
    str = VSFlexGrid1.TextMatrix(VSFlexGrid1.Row, VSFlexGrid1.ColIndex("id"))
    sr = VSFlexGrid1.TextMatrix(VSFlexGrid1.Row, VSFlexGrid1.ColIndex("serial"))
    If str <> "" Then
    Cn.BeginTrans
    BeginTrans = True
    StrSQL = "Update TblCarExpenses Set ExpID=" & val(Me.DcbExpenSiv.BoundText) & ", Vlue=" & val(TxtVlue.Text) & ", Remarks='" & TxtRemarks.Text & "' "
    StrSQL = StrSQL & " Where ID = " & val(str) & " And CarID = " & val(XPTxtID.Text) & ""
    Cn.Execute StrSQL, , adExecuteNoRecords
           
    Cn.CommitTrans
    BeginTrans = False
    If SystemOptions.UserInterface = ArabicInterface Then
    MsgBox (" „ Õ›Ÿ  ⁄œÌ· «·»Ì«‰«  »‰Ã«Õ")
    Else
    MsgBox "Save Successfully"
    End If
    Retrive_CarsExp
    Clear_CarsExpens
    
Exit Sub
errortrap:

    If BeginTrans = True Then
        BeginTrans = False
        Cn.RollbackTrans
    End If

    If Err.Number = -2147217900 Then
    If SystemOptions.UserInterface = ArabicInterface Then
        Msg = "·« Ì„ﬂ‰ Õ›Ÿ Â–Â «·»Ì«‰«  " & CHR(13)
        Msg = Msg + "·ﬁœ  „ «œŒ«· ﬁÌ„ €Ì— ’«·Õ… " & CHR(13)
        Msg = Msg + " √ﬂœ „‰ œﬁ… «·»Ì«‰«  Ê√⁄œ «·„Õ«Ê·…"
    Else
    Msg = "Can not save make sure of the validity of the data"
    End If
        MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
        Exit Sub
    End If
If SystemOptions.UserInterface = ArabicInterface Then
    Msg = "⁄›Ê«...ÕœÀ Œÿ√ „« √À‰«¡ Õ›Ÿ Â–Â «·»Ì«‰«  " & CHR(13)
   Else
   Msg = "Sory..error douring save data"
   End If
    MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
   End If
End Sub

Sub RelinExp()
Dim i As Integer
Dim SumVl As Double
SumVl = 0
With VSFlexGrid1
For i = 1 To .Rows - 1
If val(.TextMatrix(i, .ColIndex("Vlue"))) > 0 Then
SumVl = SumVl + val(.TextMatrix(i, .ColIndex("Vlue")))
End If
Next i
End With
TotalValue.Caption = SumVl
End Sub

Sub Reload()
    Dim Dcombos As New ClsDataCombos
    Dcombos.GetFixedAssets Me.DcFixedAssets, True
    fillComboParts
End Sub

Private Sub ChDrievType_Click(Index As Integer)
If Me.TxtModFlg.Text <> "R" Then
If ChDrievType(0).value = True Then
Text6.Enabled = True
DcEmployee.Enabled = True
TxtLeaderName.Enabled = False
TxtLeaderName.Text = ""
ElseIf ChDrievType(1).value = True Then
Text6.Enabled = False
DcEmployee.Enabled = False
TxtLeaderName.Enabled = True
DcEmployee.BoundText = 0
Text6.Text = ""
End If
End If
End Sub

Private Sub cmdAdd_Click()
Add_CarExpenses
End Sub
Private Sub Add_CarExpenses()
  Dim BeginTrans As Boolean
   Dim StrSQL As String
  Dim Msg As String
  On Error GoTo errortrap
  
 If val(TxtId.Text) = 0 Then
 If SystemOptions.UserInterface = ArabicInterface Then
 MsgBox "Ì—ÃÏ Õ›Ÿ »Ì«‰«  «·„⁄œ… «Ê·«"
 Else
 MsgBox "Please Save Data"
 End If
 Exit Sub
 End If
 If val(DcbExpenSiv.BoundText) = 0 Or DcbExpenSiv.Text = "" Then
 If SystemOptions.UserInterface = ArabicInterface Then
        MsgBox "Ì—ÃÏ «Œ Ì«— «·„’—Ê›"
 Else
        MsgBox "Please Select Expenses"
 End If
 DcbExpenSiv.SetFocus
 Exit Sub
 End If
 
 If val(TxtVlue.Text) = 0 Then
 If SystemOptions.UserInterface = ArabicInterface Then
        MsgBox ("«œŒ· «·ﬁÌ„…")
 Else
        MsgBox ("Enter Value ")
 End If
 TxtVlue.SetFocus
 Exit Sub
 End If
  
    Cn.BeginTrans
    BeginTrans = True
    
    Set rs_CarExpenses = New ADODB.Recordset
    StrSQL = "SELECT  *  From TblCarExpenses  "
    rs_CarExpenses.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText
    rs_CarExpenses.AddNew
    txtCode1.Text = CStr(new_id("TblCarExpenses", "ID", "", True))
    rs_CarExpenses("ID") = IIf(txtCode1.Text = "", Null, val(txtCode1.Text))
    rs_CarExpenses("CarID") = IIf(XPTxtID.Text = "", Null, val(XPTxtID.Text))
    rs_CarExpenses("Vlue") = IIf(TxtVlue.Text = "", Null, val(TxtVlue.Text))
    rs_CarExpenses("Remarks") = IIf(TxtRemarks.Text = "", Null, TxtRemarks.Text)
    rs_CarExpenses("ExpID") = IIf(Me.DcbExpenSiv.BoundText = "", Null, val(DcbExpenSiv.BoundText))
    rs_CarExpenses.update
    Cn.CommitTrans
    BeginTrans = False
  If SystemOptions.UserInterface = ArabicInterface Then
    MsgBox (" „ Õ›Ÿ «·»Ì«‰«  »‰Ã«Õ")
    Else
    MsgBox "Save Successfully"
    End If
    Retrive_CarsExp
    Clear_CarsExpens
Exit Sub
errortrap:

    If BeginTrans = True Then
        BeginTrans = False
        Cn.RollbackTrans
    End If

    If Err.Number = -2147217900 Then
    If SystemOptions.UserInterface = ArabicInterface Then
        Msg = "·« Ì„ﬂ‰ Õ›Ÿ Â–Â «·»Ì«‰«  " & CHR(13)
        Msg = Msg + "·ﬁœ  „ «œŒ«· ﬁÌ„ €Ì— ’«·Õ… " & CHR(13)
        Msg = Msg + " √ﬂœ „‰ œﬁ… «·»Ì«‰«  Ê√⁄œ «·„Õ«Ê·…"
      Else
      Msg = "Can not save make sure of the validity of the data"
      End If
        MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
        Exit Sub
    End If
If SystemOptions.UserInterface = ArabicInterface Then
    Msg = "⁄›Ê«...ÕœÀ Œÿ√ „« √À‰«¡ Õ›Ÿ Â–Â «·»Ì«‰«  " & CHR(13)
  Else
  Msg = "Sory..error douring save data "
  End If
    MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
End Sub
Private Sub Clear_CarsExpens()
TxtRemarks.Text = ""
TxtVlue.Text = ""
Me.DcbExpenSiv.BoundText = ""
End Sub
Private Sub Retrive_CarsExp()
Dim i As Integer
     Set rs_CarExpenses = New ADODB.Recordset
    Dim StrSQL As String
    
 StrSQL = " SELECT  dbo.TblCarExpenses.ID,   dbo.TblCarExpenses.CarID, dbo.TblCarExpenses.Vlue, dbo.TblCarExpenses.Remarks, dbo.TblCarExpenses.ExpID, dbo.TblDataTypeExchange.name,"
 StrSQL = StrSQL & "                      dbo.TblDataTypeExchange.NameE"
 StrSQL = StrSQL & "  FROM         dbo.TblCarExpenses LEFT OUTER JOIN"
 StrSQL = StrSQL & "                     dbo.TblDataTypeExchange ON dbo.TblCarExpenses.ExpID = dbo.TblDataTypeExchange.Id"
 StrSQL = StrSQL & " Where (dbo.TblCarExpenses.CarID = " & val(Me.XPTxtID.Text) & ")"
rs_CarExpenses.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText
    VSFlexGrid1.Rows = 1
    If rs_CarExpenses.RecordCount > 0 Then
        rs_CarExpenses.MoveFirst
        With VSFlexGrid1
        .Rows = rs_CarExpenses.RecordCount + 1
         For i = 1 To .Rows - 1
         .TextMatrix(i, .ColIndex("Serial")) = i
         .TextMatrix(i, .ColIndex("id")) = IIf(IsNull(rs_CarExpenses("id").value), "", rs_CarExpenses("id").value)
         .TextMatrix(i, .ColIndex("ExpID")) = IIf(IsNull(rs_CarExpenses("ExpID").value), 0, rs_CarExpenses("ExpID").value)
         .TextMatrix(i, .ColIndex("Remarks")) = IIf(IsNull(rs_CarExpenses("Remarks").value), "", rs_CarExpenses("Remarks").value)
         .TextMatrix(i, .ColIndex("Vlue")) = IIf(IsNull(rs_CarExpenses("Vlue").value), 0, rs_CarExpenses("Vlue").value)
         If SystemOptions.UserInterface = ArabicInterface Then
         .TextMatrix(i, .ColIndex("Name")) = IIf(IsNull(rs_CarExpenses("name").value), "", rs_CarExpenses("name").value)
         Else
         .TextMatrix(i, .ColIndex("Name")) = IIf(IsNull(rs_CarExpenses("namee").value), "", rs_CarExpenses("namee").value)
         End If
          rs_CarExpenses.MoveNext
         Next
         End With
    End If
    RelinExp
End Sub

Private Sub CmdDelete_Click()
Del_CarExpen
End Sub
Private Sub Del_CarExpen()
    Dim Msg As String
    Dim StrSQL As String
    Dim RsTemp As New ADODB.Recordset
 
    On Error GoTo ErrTrap
        Dim str As String, sr As String
        str = VSFlexGrid1.TextMatrix(VSFlexGrid1.Row, VSFlexGrid1.ColIndex("id"))
        sr = VSFlexGrid1.TextMatrix(VSFlexGrid1.Row, VSFlexGrid1.ColIndex("serial"))
        
        If str <> "" Then
 If SystemOptions.UserInterface = ArabicInterface Then
        Msg = "”Ì „ Õ–› »Ì«‰«  ”ÿ— —ﬁ„ " & CHR(13)
        Msg = Msg + (sr) & CHR(13)
        Msg = Msg + " Â·  —€» ›Ì Õ–› Â–Â «·»Ì«‰« ø"
Else
Msg = "Confirm Delete"
End If
        If MsgBox(Msg, vbYesNo + vbQuestion + vbMsgBoxRight + vbMsgBoxRtlReading, App.title) = vbYes Then
            
            If Not rs_CarExpenses.RecordCount < 1 Then
                StrSQL = "delete From TblCarExpenses  where  ID =" & val(str) & "  and carID=" & val(Me.XPTxtID.Text) & ""
                   Cn.Execute StrSQL, , adExecuteNoRecords
                   StrSQL = "SELECT  *  From TblCarExpenses"
                   Set rs_CarExpenses = New ADODB.Recordset
                   rs_CarExpenses.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText
                    Clear_CarsExpens
                If rs_CarExpenses.RecordCount < 1 Then
               
                Else
                   Retrive_CarsExp
                End If
            End If
        End If

    Else
    If SystemOptions.UserInterface = ArabicInterface Then
        Msg = "Â–Â «·⁄„·Ì… €Ì— „ «Õ… ÕÌÀ √‰Â ·«ÌÊÃœ √Ì ”Ã·« "
        Else
        MsgBox "This process is not available. There are no records"
        End If
        MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
        Exit Sub
    End If
 Retrive_CarsExp
    
    Exit Sub
ErrTrap:
 If SystemOptions.UserInterface = ArabicInterface Then
    Msg = "·«Ì„ﬂ‰ Õ–› Â–« «·”Ã· · ﬂ«„· «·»Ì«‰«  " & CHR(13) & " ÊÃœ »Ì«‰«  „— »ÿ… »Â–« «·„Œ«·›… "
    Else
    Msg = "Can not delete "
  End If
    Msg = Msg & CHR(13) & Err.description
    MsgBox Msg, vbMsgBoxRight + vbMsgBoxRtlReading + vbExclamation, App.title
    rs_CarExpenses.CancelUpdate
    'End If
End Sub

Private Sub CmdPrint_Click()
print_reportExp
End Sub

Private Sub Dcbranch_Change()
Dcbranch_Click (0)
End Sub

Private Sub Dcbranch_Click(Area As Integer)
If Me.TxtModFlg.Text <> "R" Then
ladData
End If
End Sub



Private Sub DCEmployee_KeyUp(KeyCode As Integer, Shift As Integer)
   Dim str As String
    If KeyCode = vbKeyF3 Then
        FrmEmployeeSearch.lbltype = 18
        Set FrmEmployeeSearch.RetrunFrm = Me

        FrmEmployeeSearch.show
  
    End If
    
     If KeyCode = vbKeyF5 Then
   If SystemOptions.UserInterface = ArabicInterface Then
      str = " SELECT     dbo.TblEmployee.Emp_ID, dbo.TblEmployee.Emp_Name, dbo.TblEmployee.FlagDriver, dbo.TblEmpJobsTypes.JobTypeName, dbo.TblEmpJobsTypes.JobTypeNamee, "
    str = str & "                   dbo.TblEmployee.Emp_Namee"
   Else
   str = " SELECT     dbo.TblEmployee.Emp_ID, dbo.TblEmployee.Emp_Namee, dbo.TblEmployee.FlagDriver, dbo.TblEmpJobsTypes.JobTypeName, dbo.TblEmpJobsTypes.JobTypeNamee, "
    str = str & "                   dbo.TblEmployee.Emp_Name"
   End If
    str = str & "    FROM         dbo.TblEmployee LEFT OUTER JOIN"
    str = str & "                    dbo.TblEmpJobsTypes ON dbo.TblEmployee.JobTypeID = dbo.TblEmpJobsTypes.JobTypeID"
       If SystemOptions.ShowDriverOnly = True Then

    str = str & "     where  ( JobTypeName like '%”«∆ﬁ%'  or JobTypeNamee like '%driver%' )or (FlagDriver=1) "
End If
    fill_combo DcEmployee, str
    

End If
End Sub
Function SumTotalExpen(Optional EqpID As Double) As Double
Dim sql As String
Dim Rs6 As ADODB.Recordset
Set Rs6 = New ADODB.Recordset
sql = " SELECT     SUM(Total) AS SumTotal"
sql = sql & " From dbo.TblOrderMaint"
sql = sql & " Where (EquepID = " & EqpID & ")"
Rs6.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
If Rs6.RecordCount > 0 Then
SumTotalExpen = IIf(IsNull(Rs6("SumTotal").value), 0, Rs6("SumTotal").value)
Else
SumTotalExpen = 0
End If
End Function
Private Sub DcFixedAssets_Change()
TxtName.Text = DcFixedAssets.Text
End Sub

Private Sub DcFixedAssets_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF3 Then
         FixedAssetsSearch.RetrunType = 3
        FixedAssetsSearch.show vbModal
  
    End If
End Sub

Private Sub DCGroup_KeyUp(KeyCode As Integer, Shift As Integer)
 
 
 If KeyCode = vbKeyF5 Then
  Dcombos.GetTblCarsDataGroup DCGroup
End If


End Sub

Private Sub DCInsuranceCompanyId_KeyUp(KeyCode As Integer, Shift As Integer)
Dim My_SQL As String
 If KeyCode = vbKeyF5 Then
    My_SQL = "  select id,name from insurance_companies   "
   Else
   My_SQL = "  select id,Namee from insurance_companies   "
   End If
    fill_combo DCInsuranceCompanyId, My_SQL
 
    
End Sub

Private Sub DeleteImage_Click()
DBPix201.ImageLoadFile (App.path & "\" & SystemOptions.ImagesPath & "\cars\DefaultCar.JPG")
'DBPix201.ImageBPP
DoEvents
DBPix201.ImageSaveFile (App.path & "\" & SystemOptions.ImagesPath & "\cars\" & XPTxtID.Text & ".JPG")
End Sub

Private Sub DpExpireDate_Change()
If Me.TxtModFlg.Text <> "R" Then
        DpExpireDateH.value = ToHijriDate(DpExpireDate.value)
End If

End Sub

Private Sub DpExpireDateH_LostFocus()
 If Me.TxtModFlg.Text <> "R" Then
              VBA.Calendar = vbCalGreg
            DpExpireDate.value = ToGregorianDate(DpExpireDateH.value)
        End If
End Sub

Private Sub DpSensitiveWeightDate_Change()
If Me.TxtModFlg.Text <> "R" Then
        DpSensitiveWeightDateH.value = ToHijriDate(DpSensitiveWeightDate.value)
End If
End Sub

Private Sub DpSensitiveWeightDateH_LostFocus()
 If Me.TxtModFlg.Text <> "R" Then
              VBA.Calendar = vbCalGreg
            DpSensitiveWeightDate.value = ToGregorianDate(DpSensitiveWeightDateH.value)
        End If
End Sub

Private Sub dtpEndContractDate_Change()
If Me.TxtModFlg.Text <> "R" Then
        dtpEndContractDateH.value = ToHijriDate(dtpEndContractDate.value)
End If
End Sub

Private Sub dtpEndContractDateH_LostFocus()
  If Me.TxtModFlg.Text <> "R" Then
              VBA.Calendar = vbCalGreg
            dtpEndContractDate.value = ToGregorianDate(dtpEndContractDateH.value)
        End If
End Sub

Private Sub FG_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)

With fg
If .ColKey(Col) <> "Shw" Then
Cancel = True
Else
.ComboList = ""
End If
End With
End Sub
Sub Relin()
Dim i As Integer
Dim count As Integer
Dim SumVal As Double
With fg
SumVal = 0
count = 0
For i = 1 To .Rows - 1
If val(.TextMatrix(i, .ColIndex("ID"))) <> 0 Then
count = count + 1
.TextMatrix(i, .ColIndex("Serial")) = count
SumVal = SumVal + val(.TextMatrix(i, .ColIndex("CompValue")))
End If
Next i
lbl(46).Caption = SumVal
End With
End Sub

Private Sub FG_CellButtonClick(ByVal Row As Long, ByVal Col As Long)
With fg
Select Case .ColKey(Col)
Case "Shw"
Unload FrmAccidentReport
Load FrmAccidentReport
FrmAccidentReport.FindRec val(.TextMatrix(Row, .ColIndex("ID")))
FrmAccidentReport.show
End Select
End With
End Sub

Private Sub Fg_StartEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
With fg
Select Case .ColKey(Col)
Case "Shw"
.ColComboList(.ColIndex("Shw")) = "..."
End Select
End With
End Sub

Private Sub Grid_AfterEdit(ByVal Row As Long, _
                           ByVal Col As Long)
    On Error Resume Next
    Dim StrAccountCode As String
    Dim Msg As String
    Dim rs As New ADODB.Recordset
    Dim StrSQL As String
    Dim ClsAcc As New ClsAccounts
    Dim LngRow As Long
    Dim code  As String

    With Grid

        Select Case .ColKey(Col)
 
            Case "MaintenanceType"
                code = .ComboData
           
                '   LngRow = .FindRow(Code, .FixedRows, .ColIndex("UnitID"), False, True)
                .TextMatrix(Row, .ColIndex("MaintenanceID")) = code
                .TextMatrix(Row, .ColIndex("MaintenanceType")) = .ComboItem
                Dim km As String
                Dim alarmBfore As Double
                getMaintenancetypeInformations val(code), , km, , alarmBfore
 
                .TextMatrix(Row, .ColIndex("km")) = km
                .TextMatrix(Row, .ColIndex("AlarmBefore")) = alarmBfore
                .TextMatrix(Row, .ColIndex("AlarmIn")) = val(txtLastKMCounter.Text) + val(km) - alarmBfore
 
        End Select
   
        If Row = .Rows - 1 Then
    
            .Rows = .Rows + 1
        End If

        ReLineGrid
    End With
 
    'If Me.TxtModFlg <> "E" Then Exit Sub
 
    'If Col = Grid.ColIndex("ItemCode") Or Col = Grid.ColIndex("ItemName") Then
    'RegisterItemData Me.name, Me.TxtModFlg, Grid.TextMatrix(Row, Grid.ColIndex("ItemCode")), Grid.TextMatrix(Row, Grid.ColIndex("ItemName")), , , , , , , , , , , Me.TxtTblCustomerContractD
    'ElseIf Col = Grid.ColIndex("UnitName") Then
    'RegisterItemData Me.name, Me.TxtModFlg, Grid.TextMatrix(Row, Grid.ColIndex("ItemCode")), Grid.TextMatrix(Row, Grid.ColIndex("ItemName")), Grid.TextMatrix(Row, Grid.ColIndex("UnitName")), , , , , , , , , , Me.TxtTblCustomerContractD
    ' ElseIf Col = Grid.ColIndex("Price") Then
    'RegisterItemData Me.name, Me.TxtModFlg, Grid.TextMatrix(Row, Grid.ColIndex("ItemCode")), Grid.TextMatrix(Row, Grid.ColIndex("ItemName")), , , (Grid.TextMatrix(Row, Grid.ColIndex("Price"))), , , , , , , , Me.TxtTblCustomerContractD
    ' ElseIf Col = Grid.ColIndex("Discount") Then
    'RegisterItemData Me.name, Me.TxtModFlg, Grid.TextMatrix(Row, Grid.ColIndex("ItemCode")), Grid.TextMatrix(Row, Grid.ColIndex("ItemName")), , , , , , , , , Grid.TextMatrix(Row, Grid.ColIndex("Discount")), , Me.TxtTblCustomerContractD

    'End If
    '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\LogFile Saving//////////////////////////////////////////////////////////////////////////

End Sub

Private Sub ReLineGrid()
    Dim IntCounter As Integer
    IntCounter = 0
    Dim i As Integer

    With Me.Grid

        For i = .FixedRows To .Rows - 1
    
            If .TextMatrix(i, .ColIndex("MaintenanceID")) <> "" Then
                IntCounter = IntCounter + 1
                .TextMatrix(i, .ColIndex("Ser")) = IntCounter
  
            End If

        Next i
   
    End With

End Sub

Private Sub Grid_BeforeEdit(ByVal Row As Long, _
                            ByVal Col As Long, _
                            Cancel As Boolean)

    With Grid

        If .ColKey(Col) <> "MaintenanceType" Then
            .ComboList = ""
                            
        End If
                   
        If .ColKey(Col) <> "MaintenanceType" Then
            Cancel = True
                            
        End If
                   
    End With

End Sub

Private Sub Grid_StartEdit(ByVal Row As Long, _
                           ByVal Col As Long, _
                           Cancel As Boolean)
    Dim rs As New ADODB.Recordset
    Dim StrSQL  As String
    Dim StrAccountType As String
    Dim StrComboList As String
    Dim Msg As String
    Dim LngItemID As Integer
    Dim MyStrList As String

    With Me.Grid

        Select Case .ColKey(Col)

            Case "MaintenanceType"
 
                StrSQL = "SELECT id,Name,km "
                StrSQL = StrSQL + " FROM   MaintenanceTypes "
                StrSQL = StrSQL + " Order BY  Name "
                Set rs = New ADODB.Recordset
                rs.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText

                If Not (rs.BOF Or rs.EOF) Then
                    MyStrList = .BuildComboList(rs, "Name", "id")
                    '                    Grid.ColComboList = MyStrList
                    Grid.ColComboList(.ColIndex("Name")) = "|" & MyStrList
                Else
                    Cancel = True
                End If
            
        End Select

    End With

End Sub

Private Sub Cmd_Click(Index As Integer)
    Dim msgstr  As String
'    On Error GoTo ErrTrap
    getFirstPeriodDateInthisYear FirstPeriodDateInthisYear
    XPDtbTrans.value = FirstPeriodDateInthisYear

    Select Case Index

        Case 0
            TxtModFlg.Text = "N"
            clear_all Me
            ladData
            txtRate.Caption = "1.3"
            'Rd(0).value = True
            
             If mdifrmmain.MNUFixedAssets.Visible = True Then
            Rd(0).value = True
            Else
            Rd(1).value = True
            Rd(0).Visible = False
            End If
            
            
            Rd_Click (0)
            Me.DCboUserName.BoundText = user_id
            Me.Dcbranch.BoundText = Current_branch
            DpLicenseExpireDateH.value = ToHijriDate(Date)
            DcbStuts.ListIndex = 0
            DpInsuranceExpireDateH.value = ToHijriDate(Date)
            DpTestExpireDateH.value = ToHijriDate(Date)
            Grid.Clear flexClearScrollable, flexClearEverything
            Grid.Rows = 1
             Dcbranch.BoundText = Current_branch


            dtpEndContractDateH.Visible = False
            TxtContractNo.Visible = False
            dtpEndContractDate.Visible = False
            lbl(11).Visible = False
            lbl(10).Visible = False
            lbl(14).Visible = False
            txtRep.Text = "1"

            Dim Str_Path As String
            Str_Path = App.path & "\" & SystemOptions.ImagesPath & "\cars\DefaultCar.JPG"
            If Dir(Str_Path) <> "" Then
                    DBPix201.ImageLoadFile (App.path & "\" & SystemOptions.ImagesPath & "\cars\DefaultCar.JPG")
            Else
                    DBPix201.ImageClear
            End If
    Label6.Caption = ""
        Case 1
        
        If Not SystemOptions.CanEditCars Then
            If CheckOrderMainte() = True Then
                If SystemOptions.UserInterface = ArabicInterface Then
                    MsgBox "·«Ì„ﬂ‰ «· ⁄œÌ· Â–Â «·„—ﬂ»… „— »ÿ… »«„— ‘€· ’Ì«‰… "
                Else
                    MsgBox "Can not be edited. This vehicle linked to maintenance "
                End If
                Exit Sub
            End If
        End If
        
        If val(DcbStuts.ListIndex) = 2 Then
        If SystemOptions.UserInterface = ArabicInterface Then
        MsgBox "·«Ì„ﬂ‰ «· ⁄œÌ· Â–Â «·„⁄œÂ/«·”Ì«—… „»«⁄…"
        Else
        MsgBox "Can not be edited. This Equipment Sold"
        End If
        Exit Sub
        End If
        If CheckCarAssest() = True Then
        If SystemOptions.UserInterface = ArabicInterface Then
        MsgBox "·«Ì„ﬂ‰ «· ⁄œÌ· Â–Â «·Õ—ﬂ… „— »ÿ… » ”·Ì„ «·⁄Âœ"
        Else
        MsgBox "Can not be edited. Linked to deliver the custody of the staff"
        End If
        Exit Sub
        End If
            TxtModFlg.Text = "E"
'          ladData
            ' Me.dcBranch.BoundText = my_branch
            CuurentLogdata
            Grid.Rows = Grid.Rows + 1
            Grid.Enabled = True
        
        
dtpEndContractDateH.Visible = True
TxtContractNo.Visible = True
dtpEndContractDate.Visible = True
lbl(11).Visible = True
lbl(10).Visible = True
lbl(14).Visible = True


        Case 2
            Dim currentcode As String
            If TxtId.Text = "" Then
                currentcode = get_coding(branch_id, "TblCarsData", 7, Me.DCPreFix.Text)
            If SystemOptions.UserInterface = ArabicInterface Then
                If currentcode = "miniError" Then
                    If SystemOptions.UserInterface = ArabicInterface Then
                        MsgBox "⁄œœ «·Œ«‰«  «· Ì ﬁ„  » ÕœÌœ…  ·Â–« ««ﬂÊœ ’€Ì—… Ãœ« Ì—ÃÌ  €ÌÌ—Â« ›Ì ‘«‘…  ﬂÊÌœ «·ÕﬁÊ· «Ê «·« ’«· »„”∆Ê· «·‰Ÿ«„"
                    Else
                        MsgBox "The number of fields in entered code is too small ... please change the code or contact your systen admin"
                    End If
                    Exit Sub
                ElseIf currentcode = "Manual" Then
                    If SystemOptions.UserInterface = ArabicInterface Then
                        MsgBox "«œŒ· «·ﬂÊœ ÌœÊÌ« ﬂ„« Õœœ  ›Ì  ﬂÊÌœ «·ÕﬁÊ·"
                    Else
                        MsgBox "Please enter the code manually"
                   End If
                Else
                    TxtId = currentcode
                End If
              Else
                       If currentcode = "miniError" Then
                    MsgBox "Number of small fields coding please change"
                    Exit Sub
                ElseIf currentcode = "Manual" Then
                    MsgBox "Enter the code manually"
                Else
                    TxtId = currentcode
                End If
              End If
            End If
        If Rd(1).value = True Then
        If Me.TxtEqupName.Text = "" Then
        If SystemOptions.UserInterface = ArabicInterface Then
        MsgBox "Ì—ÃÏ «œŒ«· «”„ «·„⁄œ…"
        Else
        MsgBox "Please Eneter Equipment"
        End If
        TxtEqupName.SetFocus
        Exit Sub
        End If
        End If
       If ChDrievType(1).value = True Then
        If Me.TxtLeaderName.Text = "" Then
        If SystemOptions.UserInterface = ArabicInterface Then
        MsgBox "Ì—ÃÏ «œŒ«· «”„ «·„ÊŸ›"
        Else
        MsgBox "Please Eneter Employee"
        End If
        TxtLeaderName.SetFocus
        Exit Sub
        End If
        End If

            SaveData
               
               
        Case 3
            Call Undo
        Case 4
        If CheckOrderMainte() = True Then
        If SystemOptions.UserInterface = ArabicInterface Then
        MsgBox "·«Ì„ﬂ‰ «·Õ–›. Â–Â «·„—ﬂ»… „— »ÿ… »«„— ‘€· ’Ì«‰… "
        Else
        MsgBox "Can not be delete. This vehicle linked to maintenance "
        End If
        Exit Sub
        End If
        If val(DcbStuts.ListIndex) = 2 Then
        If SystemOptions.UserInterface = ArabicInterface Then
        MsgBox "·«Ì„ﬂ‰ «·Õ–›. Â–Â «·„⁄œÂ/«·”Ì«—… „»«⁄…"
        Else
        MsgBox "Can not be delete. This Equipment Sold"
        End If
        Exit Sub
        End If
         If CheckCarAssest() = True Then
        If SystemOptions.UserInterface = ArabicInterface Then
        MsgBox "·«Ì„ﬂ‰ «·Õ–› Â–Â «·Õ—ﬂ… „— »ÿ… » ”·Ì„ «·⁄Âœ"
        Else
        MsgBox "Can not be delete. Linked to deliver the custody of the staff"
        End If
        Exit Sub
        End If
            Del_AssetType
        Case 5
              If DoPremis(Do_Attach, Me.Name, True) = False Then
                Exit Sub
            End If
            VIEW_ATTACH
        Case 6
            Unload Me
        Case 7 ' «Ìﬁ«› «·«Â·«ﬂ
            cStatus.ListIndex = 1
            Cmd(7).Enabled = False
        Case 8 ' «⁄«œ…  ‘€Ì· «·«Â·«ﬂ
            cStatus.ListIndex = 0
            Cmd(8).Enabled = False
        Case 9 ' «· Œ·’ „‰ «·«’·
            cStatus.ListIndex = 3
            Cmd(9).Enabled = False
        Case 10
        Dim Msg As String
        If CheckPlan(val(XPTxtID.Text)) = True Then
            publicCarId = val(Me.XPTxtID.Text)
            FrmCarsPlan.show
          Else
          If SystemOptions.UserInterface = ArabicInterface Then
          Msg = "·«ÌÊÃœ ŒÿÂ „”»ﬁ… Â·  —Ìœ ⁄„· ŒÿÂ ÃœÌœ…"
          Else
          Msg = "There is no plan. Do you want to make a new plan? "
          End If
      If MsgBox(Msg, vbYesNo + vbQuestion + vbMsgBoxRight + vbMsgBoxRtlReading, App.title) = vbYes Then
      FrmCarsPlan.Cmd_Click (0)
      FrmCarsPlan.DCCar.BoundText = val(Me.XPTxtID.Text)
      FrmCarsPlan.TXTCurrentKM.Text = Me.txtLastKMCounter.Text
      FrmCarsPlan.show
      Else
      Exit Sub
      End If
          End If
        Case 11
            ShowGL_cc Me.TxtNoteSerial.Text, , 200
        Case 12
               Unload FrmCasrShearches
                   Load FrmCasrShearches
                    FrmCasrShearches.show vbModal
             '   TblCarsDataSearch.RetrunType = 0
            '    TblCarsDataSearch.Show vbModal
       Case 17
            print_report
        Case 18
            print_report2
         Case 19
         print_reportExpense
    End Select

    Exit Sub
ErrTrap:
End Sub
Function CheckPlan(Optional CarID As Double) As Boolean
Dim StrSQL As String
Dim rs2 As ADODB.Recordset
Set rs2 = New ADODB.Recordset
StrSQL = "select * From TblCarMaintenancePlan where CarId=" & CarID
rs2.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText
If rs2.RecordCount > 0 Then
CheckPlan = True
Else
CheckPlan = False
End If
End Function
Function VIEW_ATTACH()
    'On Error Resume Next''
 
    'If TxtEmp_Code.text = "" Then MsgBox "·«»œ „‰ «Õ Ì«— „ÊŸ› «Ê·«": Exit Sub

    imaged.show
    imaged.Label9.Caption = "„—›ﬁ«  «·„⁄œÂ/«·”Ì«—… —ﬁ„"
    imaged.Caption = "„—›ﬁ«  «·„⁄œÂ/«·”Ì«—…  "
    imaged.txtopeation_type = "„—›ﬁ«  «·„⁄œÂ/«·”Ì«—…"
    imaged.SUBJECT_NO = XPTxtID 'TxtEmp_Code.text
    imaged.Label6.Caption = "ﬂÊœ «·„⁄œÂ/«·”Ì«—…"
    imaged.Adodc1.CommandType = adCmdText
    imaged.Adodc1.RecordSource = "SELECT * FROM subjects_images WHERE operation_type = '„—›ﬁ«  «·„⁄œÂ/«·”Ì«—…' and subject_no='" & XPTxtID & "'"
    imaged.Adodc1.Refresh

    If imaged.Adodc1.Recordset.RecordCount > 0 Then

        imaged.DBPix201.Visible = True
    Else
        imaged.DBPix201.Visible = False
    End If

End Function
 
Function CuurentLogdata(Optional Currentmode As String)
    LogTextA = "    ‘«‘… " & ScreenNameArabic & CHR(13) & "ﬂÊœ «·«’· " & DCPreFix & TxtId.Text & CHR(13) & " «”„  «·«’·   " & TxtName & CHR(13) & "   «·„Ã„Ê⁄Â   " & DCGroup & CHR(13) & "      «·›—⁄   " & Dcbranch & CHR(13) & " Õ«·… «·«’· " & cStatus

    If Option1.value = True Then
        LogTextA = LogTextA & CHR(13) & "      ÃœÌœ     "
    ElseIf Option2.value = True Then
        LogTextA = LogTextA & CHR(13) & "   «›  «ÕÌ  "
                
    End If
                    
     LogTextA = LogTextA & CHR(13) & "   ÿ—Ìﬁ… «·«Â·«ﬂ   " & CBoDepreciation_Type_id & CHR(13) & "    «—ÌŒ »œ«Ì…  «·«Â·«ﬂ   " & "" & CHR(13) & "    «—ÌŒ «Œ—  «·«Â·«ﬂ   " & "" & CHR(13) & "        ﬁÌ„… ‘—«¡ «·«’·    " & TxtPurchasePrice & CHR(13) & "         «—ÌŒ ‘—«¡ «·«’·   " & DpPurchaseDate & CHR(13) & "        ﬁÌ„… «·«’· ﬂŒ—œ…   " & TxtKhordaPrice & CHR(13) & " «·ﬁÌ„… «·œ› —Ì…  " & TxtCurrentValue & CHR(13) & "  „Ã„⁄ «·«Â·«ﬂ   " & TxtAccDepreciation & CHR(13) & "     ﬁÌ„… ﬁ”ÿ  «·«Â·«ﬂ   " & txtinstallValue & CHR(13) & "          «ﬁ”«ÿ «·«Â·«ﬂ  «·„‰›–…  " & txtinstallDo & CHR(13) & "        ⁄œœ «ﬁ”«ÿ «·«Â·«ﬂ  «·„ »ﬁÌ… " & txtinstallmentresult & CHR(13) & " «·⁄„— «·«› —«÷Ì ··«’· »«·‘Â—" & TxtAge & CHR(13) & "  Ê’› «·«’· " & TxtNotes
       LogTexte = "    Screen  " & ScreenNameEnglish & CHR(13) & "F.A. Code   " & DCPreFix & TxtId.Text & CHR(13) & " F.A. Name " & TxtName & CHR(13) & "   Group   " & DCGroup & CHR(13) & "      Branch   " & Dcbranch & CHR(13) & "    Status " & cStatus

    If Option1.value = True Then
        LogTexte = LogTextA & CHR(13) & "      New     "
    ElseIf Option2.value = True Then
        LogTexte = LogTextA & CHR(13) & "   Opening  "
                
    End If
                    
     LogTexte = LogTexte & CHR(13) & " Depreciation Type  " & CBoDepreciation_Type_id & CHR(13) & "  Start Depreciation Date    " & "" & CHR(13) & "  LastDepreciationDate   " & "" & CHR(13) & "   PurchasePrice    " & TxtPurchasePrice & CHR(13) & " PurchaseDate" & DpPurchaseDate & CHR(13) & "       Khorda Price  " & TxtKhordaPrice & CHR(13) & " CurrentValue" & TxtCurrentValue & CHR(13) & "  Acc. Depreciation   " & TxtAccDepreciation & CHR(13) & "    intinstallment Value   " & txtinstallValue & CHR(13) & "  installment Done " & txtinstallDo & CHR(13) & "      Remain installment " & txtinstallmentresult & CHR(13) & " Age Range By Month " & TxtAge & CHR(13) & " Remarks " & TxtNotes
       If Currentmode <> "D" Then
        AddToLogFile CInt(user_id), 0, Date, Time, LogTextA, LogTexte, Me.Name, Me.TxtModFlg, , , val(TxtNoteSerial)
    Else
        AddToLogFile CInt(user_id), 0, Date, Time, LogTextA, LogTexte, Me.Name, "D", , , val(TxtNoteSerial)
    End If
    
End Function
Sub LodR()
Dim str As String
  If SystemOptions.UserInterface = ArabicInterface Then
      str = " SELECT     dbo.TblEmployee.Emp_ID, dbo.TblEmployee.Emp_Name, dbo.TblEmployee.FlagDriver, dbo.TblEmpJobsTypes.JobTypeName, dbo.TblEmpJobsTypes.JobTypeNamee, "
    str = str & "                   dbo.TblEmployee.Emp_Namee"
   Else
   str = " SELECT     dbo.TblEmployee.Emp_ID, dbo.TblEmployee.Emp_Namee, dbo.TblEmployee.FlagDriver, dbo.TblEmpJobsTypes.JobTypeName, dbo.TblEmpJobsTypes.JobTypeNamee, "
    str = str & "                   dbo.TblEmployee.Emp_Name"
   End If
    str = str & "    FROM         dbo.TblEmployee LEFT OUTER JOIN"
    str = str & "                    dbo.TblEmpJobsTypes ON dbo.TblEmployee.JobTypeID = dbo.TblEmpJobsTypes.JobTypeID"
    
   If SystemOptions.ShowDriverOnly = True Then
   str = str & "     where  ( JobTypeName like '%”«∆ﬁ%'  or JobTypeNamee like '%driver%' )or (FlagDriver=1) "
   End If
    fill_combo DcEmployee, str

End Sub
Private Sub Form_KeyDown(KeyCode As Integer, _
                         Shift As Integer)
    On Error GoTo ErrTrap

    If KeyCode = vbKeyReturn Then
        If Me.TxtModFlg.Text = "R" Then
            Cmd_Click (0)
        Else
            SendKeys "{TAB}"
        End If
    End If

    If Me.TxtModFlg.Text = "R" Then
        If KeyCode = vbKeyDown Or KeyCode = vbKeyEnd Then
            XPBtnMove_Click (2)
        ElseIf KeyCode = vbKeyUp Or KeyCode = vbKeyHome Then
            XPBtnMove_Click (1)
        ElseIf KeyCode = vbKeyRight Or KeyCode = vbKeyPageDown Then
            XPBtnMove_Click (3)
        ElseIf KeyCode = vbKeyLeft Or KeyCode = vbKeyPageUp Then
            XPBtnMove_Click (0)
        End If
    End If

    If KeyCode = vbKeyF12 Then
        If Cmd(0).Enabled = False Then Exit Sub
        Cmd_Click (0)
    End If

    If KeyCode = vbKeyF11 Then
        If Cmd(1).Enabled = False Then Exit Sub
        Cmd_Click (1)
    End If

    If KeyCode = vbKeyF10 Then
        If Cmd(2).Enabled = False Then Exit Sub
        Cmd_Click (2)
    End If

    If KeyCode = vbKeyF9 Then
        If Cmd(3).Enabled = False Then Exit Sub
        Cmd_Click (3)
    End If

    If KeyCode = vbKeyF8 Then
        If Cmd(4).Enabled = False Then Exit Sub
        Cmd_Click (4)
    End If

    If Shift = 2 Then
        If KeyCode = vbKeyX Then
            If Cmd(6).Enabled = False Then Exit Sub
            Cmd_Click (6)
        End If
    End If

    Exit Sub
ErrTrap:
End Sub
Sub ladData()
 Dim Dcombos As New ClsDataCombos
     Dcombos.ClearMyDataCombo DcEmployee
Dim str As String
If SystemOptions.UserInterface = ArabicInterface Then
    str = " SELECT     dbo.TblEmployee.Emp_ID, dbo.TblEmployee.Emp_Name, dbo.TblEmployee.FlagDriver, dbo.TblEmpJobsTypes.JobTypeName, dbo.TblEmpJobsTypes.JobTypeNamee, "
    str = str & "                   dbo.TblEmployee.Emp_Namee ,dbo.TblEmployee.BranchId"
   Else
   str = " SELECT     dbo.TblEmployee.Emp_ID, dbo.TblEmployee.Emp_Namee, dbo.TblEmployee.FlagDriver, dbo.TblEmpJobsTypes.JobTypeName, dbo.TblEmpJobsTypes.JobTypeNamee, "
    str = str & "                   dbo.TblEmployee.Emp_Name ,dbo.TblEmployee.BranchId "
   End If
    str = str & "    FROM         dbo.TblEmployee LEFT OUTER JOIN"
    str = str & "                    dbo.TblEmpJobsTypes ON dbo.TblEmployee.JobTypeID = dbo.TblEmpJobsTypes.JobTypeID"
       If SystemOptions.ShowDriverOnly = True Then

    str = str & "     where  (( JobTypeName like '%”«∆ﬁ%'  or JobTypeNamee like '%driver%' )or (FlagDriver=1)) and (dbo.TblEmployee.BranchId=" & val(Me.Dcbranch.BoundText) & ")"
    End If
   
    fill_combo DcEmployee, str
End Sub
Private Sub Form_Load()
    'On Error GoTo ErrTrap
    Dim Dcombos As New ClsDataCombos
    'Dcombos.ClearMyDataCombo
    'Dcombos.GetBoxes Me.DcboBox
    Dcombos.GetUsers Me.DCboUserName
    Dcombos.GetTblCarsDataGroup DCGroup
    Dcombos.GetTypeExchange Me.DcbExpenSiv
    Dcombos.GetFixedAssets Me.DcFixedAssets, True
    Dcombos.GetCountriesNames Me.DcboCountryID2
    Dcombos.GetPrefix Me.DCPreFix, 7, 0
Dcombos.GetAccountingCodes Me.DcbAccount, True, False
If SystemOptions.CarsRevenuePerOwner = True Then
Frame12(2).Visible = True

End If

    LodR
    Dim My_SQL As String

    ScreenNameArabic = " »Ì«‰«  «·„⁄œ« /«·”Ì«—«   "
    ScreenNameEnglish = "Cars Data"
    RegisterLogInOut Me.Name, ScreenNameArabic, ScreenNameEnglish, "1"
    
    If SystemOptions.UserInterface = ArabicInterface Then
        My_SQL = "  select id,name from insurance_companies   "
    Else
        My_SQL = "  select id,Namee from insurance_companies   "
    End If
    
    fill_combo DCInsuranceCompanyId, My_SQL

    If SystemOptions.UserInterface = ArabicInterface Then
        With DcbStuts
            .Clear
            .AddItem "‰‘ÿ"
            .AddItem " Õ  «·’Ì«‰…"
            .AddItem "„»«⁄"
        End With
    Else
        With DcbStuts
            .Clear
            .AddItem "Active"
            .AddItem "Under Maintenance "
            .AddItem "Sold"
        End With
    End If
    
    Dcombos.GetBranches Dcbranch
    Dcombos.GetEmpDepartments Me.DcbDept
    Dcombos.GetEmpJobsTypes Me.DcbJob
    
    '//////////////////
    Dim str As String
    ScreenNameEnglish = "Cars Data"
    RegisterLogInOut Me.Name, ScreenNameArabic, ScreenNameEnglish, "1"
    If SystemOptions.UserInterface = ArabicInterface Then
        str = " select id , name from tblcolor "
    Else
        str = " select id , namee from tblcolor "
    End If
    fill_combo VColor, str

    Dcombos.GetTblCarsDataGroup VType
    If SystemOptions.UserInterface = ArabicInterface Then
        str = " select id , Model  from TblCarModels"
    Else
          str = " select id , ModelE  from TblCarModels"
    End If
    fill_combo VModel, str
    
    Dcombos.GetEmpLocations LocationID
    
    '///////////////////
    'My_SQL = "  select Emp_ID,Emp_name  from TblEmployee  where not(driverid is null) order by Emp_name   "
    'My_SQL = "  select Emp_ID,Emp_name  from TblEmployee  order by Emp_name "
    'fill_combo DCEmployee, My_SQL
    'Dcombos.GetEmployees Me.DcEmployee, , True
    
    'If SystemOptions.UserInterface = ArabicInterface Then
    
    Set Cmd(0).ButtonImage = mdifrmmain.ImgLstTree.ListImages("New").Picture
    Set Cmd(1).ButtonImage = mdifrmmain.ImgLstTree.ListImages("Edit").Picture
    Set Cmd(2).ButtonImage = mdifrmmain.ImgLstTree.ListImages("save").Picture
    Set Cmd(3).ButtonImage = mdifrmmain.ImgLstTree.ListImages("Undo").Picture
    Set Cmd(4).ButtonImage = mdifrmmain.ImgLstTree.ListImages("Del").Picture
    Set Cmd(6).ButtonImage = mdifrmmain.ImgLstTree.ListImages("Exit").Picture
    
    'Dcombos.GetAccountingCodes Me.DcboCreditSide

    Resize_Form Me

    If SystemOptions.UserInterface = EnglishInterface Then
        SetInterface Me
        ChangeLang
    End If

    Dim GrdBack As ClsBackGroundPic
    Set GrdBack = New ClsBackGroundPic

    With Me.Grid
        Set .WallPaper = GrdBack.Picture
    End With

    With Me.Grid
        .Rows = 1
        .ExplorerBar = flexExSortShowAndMove
        .RowHeightMin = 300
        .ExtendLastCol = True
    End With

    AddTip
    Dim StrSQL As String
    Set rs = New ADODB.Recordset
    'rs.Open "TblCarsData", Cn, adOpenStatic, adLockOptimistic, adCmdTable
    StrSQL = "select * from TblCarsData"
    If SystemOptions.usertype <> UserAdminAll Then
        StrSQL = StrSQL & " where branch_no in (select BranchID from TblUsersBranches where userid =" & user_id & ")"
    End If
    rs.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText

    If SystemOptions.usertype <> UserAdminAll Then
        Dcbranch.BoundText = Current_branch
        Dcbranch.Enabled = False
    End If
    
    Me.TxtModFlg.Text = "R"

    Retrive
    Retrive_CarsExp
    If OPEN_NEW_SCREEN = True Then
        Cmd_Click (0)
    End If

    C1Tab1.CurrTab = 0

    Exit Sub
ErrTrap:
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, _
                             UnloadMode As Integer)
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
       
                SaveData

            Case vbCancel
                Cancel = True
        End Select

    End If

    Exit Sub
ErrTrap:
End Sub

Private Sub Form_Unload(Cancel As Integer)
    On Error GoTo ErrTrap
    RegisterLogInOut Me.Name, ScreenNameArabic, ScreenNameEnglish

    If rs.State = adStateOpen Then
        If Not (rs.EOF Or rs.BOF) Then
            If rs.EditMode <> adEditNone Then
                rs.CancelUpdate
            End If
        End If

        rs.Close
    End If

    Set rs = Nothing
    Set TTP = Nothing
    Exit Sub
ErrTrap:
End Sub
 








Private Sub LetterCount_Change()
Total.Text = val(LetterCount.Text) * val(LetterPrice.Text)
End Sub

Private Sub LetterCount_KeyPress(KeyAscii As Integer)
KeyAscii = KeyAscii_Num(KeyAscii, Me.LetterCount.Text, 0)
End Sub

Private Sub LetterPrice_Change()
        Total.Text = val(LetterCount.Text) * val(LetterPrice.Text)
End Sub

Private Sub LetterPrice_KeyPress(KeyAscii As Integer)
        KeyAscii = KeyAscii_Num(KeyAscii, Me.LetterPrice.Text, 0)
End Sub

Private Sub LocationID_KeyUp(KeyCode As Integer, Shift As Integer)
 
 
 If KeyCode = vbKeyF5 Then
     Dcombos.GetEmpLocations LocationID
 End If
 
End Sub

Private Sub printPartsRep_Click()
    print_report_Parts
End Sub

Private Sub PushButton1_Click()
If Me.TxtModFlg.Text = "R" Then
If val(DcbStuts.ListIndex) <> -1 Then
Cn.Execute "Update TblCarsData set StutsID=" & val(DcbStuts.ListIndex) & "  where  id =" & val(Me.XPTxtID.Text) & ""
rs.Resync
If SystemOptions.UserInterface = ArabicInterface Then
MsgBox " „ «· ÕœÌÀ"
Else
MsgBox "Update Successfully"
End If
End If
End If
End Sub

Private Sub Rd_Click(Index As Integer)

If Rd(1).value = True Then
DcFixedAssets.Visible = False
If Me.TxtModFlg.Text <> "R" Then
DcFixedAssets.BoundText = ""
End If
TxtEqupName.Visible = True
ElseIf Rd(0).value = True Then
DcFixedAssets.Visible = True
TxtEqupName.Visible = False
If Me.TxtModFlg.Text <> "R" Then
TxtEqupName.Text = ""
End If
End If

End Sub

Private Sub TxtAccount_Change()
DcbAccount.BoundText = getAccountSerial_Code("Account_Code", "Account_Serial", TxtAccount.Text)
End Sub

Private Sub txtBoardNO_Change()
    Label6.Caption = Me.TxtBoardNO.Text
End Sub

Private Sub txtCapacity_Change()
txtMax.Text = val(TxtCapacity.Text) * val(txtRep.Text)
End Sub

Private Sub TxtEquQty_KeyPress(KeyAscii As Integer)
  KeyAscii = KeyAscii_Num(KeyAscii, TxtEquQty.Text, 0)
End Sub

Private Sub TxtKhordaPrice_KeyPress(KeyAscii As Integer)
    KeyAscii = KeyAscii_Num(KeyAscii, TxtKhordaPrice.Text, 0)
End Sub

Private Sub txtLastKMCounter_KeyPress(KeyAscii As Integer)
    KeyAscii = KeyAscii_Num(KeyAscii, txtLastKMCounter.Text, 0)

End Sub



Private Sub txtLetter1_KeyPress(KeyAscii As Integer)
txtLetter1.Text = ""
If Len(txtLetter1.Text) > 0 Then
KeyAscii = 0
End If
Select Case CHR(KeyAscii)
    Case 0 To 9
        KeyAscii = 0
    Case 8
        Exit Sub
    Case Else
        txtLetter2.SetFocus
End Select

End Sub

Private Sub txtLetter1_KeyUp(KeyCode As Integer, Shift As Integer)
Cal_Board
End Sub

Private Sub txtLetter2_KeyPress(KeyAscii As Integer)
txtLetter2.Text = ""
If Len(txtLetter2.Text) > 0 Then
KeyAscii = 0
End If
Select Case CHR(KeyAscii)
    Case 0 To 9
        KeyAscii = 0
    Case Else
        txtLetter3.SetFocus
End Select
Cal_Board
End Sub

Private Sub txtLetter2_KeyUp(KeyCode As Integer, Shift As Integer)
Cal_Board
End Sub

Private Sub txtLetter3_KeyPress(KeyAscii As Integer)
txtLetter3.Text = ""
If Len(txtLetter3.Text) > 0 Then
KeyAscii = 0
End If
Select Case CHR(KeyAscii)
    Case 0 To 9
        KeyAscii = 0
    Case Else
        txtLetter4.SetFocus
End Select
Cal_Board
End Sub

Private Sub txtLetter3_KeyUp(KeyCode As Integer, Shift As Integer)
Cal_Board
End Sub

Private Sub txtLetter4_KeyPress(KeyAscii As Integer)
txtLetter4.Text = ""
If Len(txtLetter4.Text) > 0 Then
KeyAscii = 0
End If
Select Case CHR(KeyAscii)
    Case 0 To 9
        KeyAscii = 0
    Case Else
        txtNum1.SetFocus
End Select
Cal_Board
End Sub
Private Sub DcboCountryID2_Change()
  Dim Dcombos As ClsDataCombos
    Set Dcombos = New ClsDataCombos
     Dcombos.getCountriesGovernments Me.DcboGovernmentID, val(Me.DcboCountryID2.BoundText)
End Sub

Private Sub DcboCountryID2_Click(Area As Integer)
DcboCountryID2_Change
End Sub

Private Sub Cal_Board()
    TxtBoardNO.Text = txtLetter1.Text & " " & txtLetter2.Text & " " & txtLetter3.Text & " " & txtLetter4.Text & " " & txtNum1.Text & " " & txtNum2.Text & " " & txtNum3.Text & " " & txtNum4.Text
End Sub
Sub ChanEnabel(Optional Ind As Integer = 0)
If Ind = 0 Then
Dcbranch.Enabled = True
Rd(0).Enabled = True
Rd(1).Enabled = True
DcbJob.Enabled = True
TxtId.Enabled = True
DCPreFix.Enabled = True
TxtEqupName.Enabled = True
DcFixedAssets.Enabled = True
DCGroup.Enabled = True
C1Elastic12.Enabled = True
TxtNatinality.Enabled = True
txtJob.Enabled = True
TxtDepartment.Enabled = True
Else
Dcbranch.Enabled = False
Rd(0).Enabled = False
Rd(1).Enabled = False
DcbJob.Enabled = False
TxtId.Enabled = False
DCPreFix.Enabled = False
TxtEqupName.Enabled = False
DcFixedAssets.Enabled = False
DCGroup.Enabled = False
C1Elastic12.Enabled = False
TxtNatinality.Enabled = False
txtJob.Enabled = False
TxtDepartment.Enabled = False
End If
End Sub
Private Sub txtLetter4_KeyUp(KeyCode As Integer, Shift As Integer)
Cal_Board
End Sub

Private Sub TxtModFlg_Change()
    On Error GoTo ErrTrap
ChanEnabel
PushButton1.Enabled = False
C1Elastic16.Enabled = False
    Select Case Me.TxtModFlg.Text

        Case "R"
        C1Elastic16.Enabled = True
        Rd(0).Enabled = False
        Rd(1).Enabled = False
        LodR
        ChanEnabel 1
        PushButton1.Enabled = True
            '  txtLastKMCounter.locked = True
            '   Me.Caption = "«·«’Ê· «·À«» …"
            Me.Cmd(2).Enabled = False
            Me.Cmd(3).Enabled = False
            cStatus.Enabled = False
            Me.Cmd(0).Enabled = True
            Me.Cmd(1).Enabled = True
            Me.Cmd(4).Enabled = True
            Me.Cmd(12).Enabled = True
        
            Me.XPBtnMove(0).Enabled = True
            Me.XPBtnMove(1).Enabled = True
            Me.XPBtnMove(2).Enabled = True
            Me.XPBtnMove(3).Enabled = True
            Frame3.Enabled = False
        
            If rs.RecordCount < 1 Then
                     
              '       Me.XPBtnMove(0).Enabled = False
               '      Me.XPBtnMove(1).Enabled = False
                '      Me.XPBtnMove(2).Enabled = False
                '      Me.XPBtnMove(3).Enabled = False
                Me.Cmd(1).Enabled = False
                Me.Cmd(4).Enabled = False
            End If
            
           ' C1Elastic2.Enabled = False
            C1Elastic3.Enabled = False
            C1Elastic6.Enabled = False
        Case "N"
              Rd(0).Enabled = True
             Rd(1).Enabled = True
            '   Me.Caption = "√‰Ê«⁄ «·„’—Ê›« ( ÃœÌœ )"
            txtLastKMCounter.locked = False
            Me.Cmd(2).Enabled = True
            Me.Cmd(3).Enabled = True
            cStatus.Enabled = True
            Me.Cmd(0).Enabled = False
            Me.Cmd(1).Enabled = False
            Me.Cmd(4).Enabled = False
        Me.Cmd(12).Enabled = False
        
            Me.XPBtnMove(0).Enabled = False
            Me.XPBtnMove(1).Enabled = False
            Me.XPBtnMove(2).Enabled = False
            Me.XPBtnMove(3).Enabled = False
         
            C1Elastic2.Enabled = True
            C1Elastic3.Enabled = True
            C1Elastic6.Enabled = True
        Case "E"
        Rd(0).Enabled = False
        Rd(1).Enabled = False
            ' txtLastKMCounter.locked = True
            Frame3.Enabled = False
            '   Me.Caption = "√‰Ê«⁄ «·„’—Ê›« (  ⁄œÌ· )"
            Me.Cmd(2).Enabled = True
            Me.Cmd(3).Enabled = True
            cStatus.Enabled = False
            Me.Cmd(0).Enabled = False
            Me.Cmd(1).Enabled = False
            Me.Cmd(4).Enabled = False
        Me.Cmd(12).Enabled = False
        
            Me.XPBtnMove(0).Enabled = False
            Me.XPBtnMove(1).Enabled = False
            Me.XPBtnMove(2).Enabled = False
            Me.XPBtnMove(3).Enabled = False
            
            C1Elastic2.Enabled = True
            C1Elastic3.Enabled = True
            C1Elastic6.Enabled = True
    End Select

    Exit Sub
ErrTrap:
End Sub

Private Sub DcbAccount_Click(Area As Integer)
TxtAccount.Text = getAccountSerial_Code("Account_Serial", "Account_Code", DcbAccount.BoundText)
 End Sub






Public Sub Retrive(Optional Lngid As Long = 0)
    'On Error GoTo ErrTrap
LodR
Reload
    If rs.RecordCount < 1 Then
        XPTxtCurrent.Caption = 0
        XPTxtCount.Caption = 0
        Exit Sub
    End If

    If Not (rs.EOF Or rs.BOF) Then
        If Lngid <> 0 Then
            rs.find "id=" & Lngid, , adSearchForward, adBookmarkFirst

            If rs.EOF Or rs.BOF Then
                Exit Sub
            End If
        End If

    End If
 
    Me.XPTxtID.Text = IIf(val(rs("id").value) = 0, 0, val(rs("id").value))
    DcbAccount.BoundText = IIf(IsNull(rs("AccountPaym").value), "", rs("AccountPaym").value)



    Dcbranch.BoundText = IIf(val(rs("Branch_NO").value) = 0, 0, (rs("Branch_NO").value))
    Me.TxtId.Text = IIf(IsNull(rs("code").value), "", rs("code").value)
    Me.DcbStuts.ListIndex = IIf(IsNull(rs("StutsID").value), -1, rs("StutsID").value)
    DCPreFix.Text = IIf(IsNull(rs("prifix").value), "", rs("prifix").value)
    DCGroup.BoundText = IIf(IsNull(rs("CarsTypeId").value), "", rs("CarsTypeId").value)
    DcFixedAssets.BoundText = IIf(IsNull(rs("fixedAssetid").value), "", (rs("fixedAssetid").value))
    Me.VehicleLong.Text = IIf(IsNull(rs("VehicleLong").value), "", rs("VehicleLong").value)
    Me.TxtLicenseNO.Text = IIf(IsNull(rs("LicenseNO").value), "", rs("LicenseNO").value)
    Me.TxtName.Text = IIf(IsNull(rs("Name").value), "", rs("Name").value)
    'BoardNO
    Me.TxtBoardNO.Text = IIf(IsNull(rs("BoardNO").value), "", rs("BoardNO").value)
    Label6.Caption = Me.TxtBoardNO.Text
    Me.TxtModel.Text = IIf(IsNull(rs("Model").value), "", rs("Model").value)

    DpPurchaseDate.value = IIf(IsNull(rs("PurchaseDate").value), Date, rs("PurchaseDate").value) ' rs("PurchaseDate").value
    
    DpExpireDate.value = IIf(IsNull(rs("ExpireDate").value), Date, rs("ExpireDate").value) ' rs("ExpireDate").value
    DpSensitiveWeightDate.value = IIf(IsNull(rs("SensitiveWeightDate").value), Date, rs("SensitiveWeightDate").value) ' rs("SensitiveWeightDate").value
 
    txtLastKMCounter.Text = IIf(IsNull(rs("LastKMCounter").value), "", rs("LastKMCounter").value) 'val(rs("LastKMCounter").value)
    'InsuranceCompanyId
    DCInsuranceCompanyId.BoundText = IIf(IsNull(rs("InsuranceCompanyId").value), 0, (rs("InsuranceCompanyId").value))
    DcEmployee.BoundText = IIf(IsNull(rs("Emp_id").value), "", rs("Emp_id").value)

    Me.TxtNotes.Text = IIf(IsNull(rs("Notes").value), "", rs("Notes").value)
        Me.TxtEquQty.Text = IIf(IsNull(rs("EquQty").value), "", rs("EquQty").value)
        
    '

    'DpLicenseExpireDate.value = rs("LicenseExpireDate").value
    'DpInsuranceExpireDate.value = rs("InsuranceExpireDate").value
    'DpTestExpireDate.value = rs("TestExpireDate").value

    DpLicenseExpireDateH.value = IIf(IsNull(rs("LicenseExpireDateH").value), ToHijriDate(Date), rs("LicenseExpireDateH").value)
    DpInsuranceExpireDateH.value = IIf(IsNull(rs("InsuranceExpireDateH").value), ToHijriDate(Date), rs("InsuranceExpireDateH").value)
    DpTestExpireDateH.value = IIf(IsNull(rs("TestExpireDateH").value), ToHijriDate(Date), rs("TestExpireDateH").value)

    txtSetCount.Text = IIf(IsNull(rs("SetCount").value), "", rs("SetCount").value)
    'txtRate.Caption = IIf(IsNull(rs("Rate").value), "", rs("Rate").value)
    TxtCapacity.Text = IIf(IsNull(rs("capacity").value), "", rs("capacity").value)
    TxtContractNo.Text = IIf(IsNull(rs("ContractID").value), "", rs("ContractID").value)
    dtpEndContractDate.value = IIf(IsNull(rs("EndContractDate").value), Date, rs("EndContractDate").value)
    dtpEndContractDateH.value = IIf(IsNull(rs("EndContractDateh").value), "", rs("EndContractDateh").value)
    
        DpExpireDateH.value = IIf(IsNull(rs("ExpireDateH").value), "", rs("ExpireDateH").value)
            DpSensitiveWeightDateH.value = IIf(IsNull(rs("SensitiveWeightDateH").value), "", rs("SensitiveWeightDateH").value)

    txtRep.Text = IIf(IsNull(rs("rep").value), "", rs("rep").value)
    Me.txtMax.Text = IIf(IsNull(rs("MaxCap").value), "", rs("MaxCap").value)
    
    XPTxtCurrent.Caption = rs.AbsolutePosition
    XPTxtCount.Caption = rs.RecordCount
    TxtOperatorN.Text = IIf(IsNull(rs("OperatorN").value), "", rs("OperatorN").value)
    
   TxtGearno.Text = IIf(IsNull(rs("Gearno").value), "", rs("Gearno").value)
   TxtGearno1.Text = IIf(IsNull(rs("Gearno1").value), "", rs("Gearno1").value)
   TxtMachineno.Text = IIf(IsNull(rs("Machineno").value), "", rs("Machineno").value)
   TxtMachineno1.Text = IIf(IsNull(rs("Machineno1").value), "", rs("Machineno1").value)
       
    VModel.BoundText = IIf(IsNull(rs("VModel").value), "", rs("VModel").value)
    VType.BoundText = IIf(IsNull(rs("VType").value), "", rs("VType").value)
    VColor.BoundText = IIf(IsNull(rs("VColor").value), "", rs("VColor").value)
    Chesis.Text = IIf(IsNull(rs("Chesis").value), "", rs("Chesis").value)
    LocationID.BoundText = IIf(IsNull(rs("LocationID").value), "", rs("LocationID").value)
    TxtEqupName.Text = IIf(IsNull(rs("EqupName").value), "", rs("EqupName").value)
    
    If Not (IsNull(rs("IsUsed").value)) Then
        If rs("IsUsed").value = True Then
            chkIsUsed.value = vbChecked
        Else
            chkIsUsed.value = vbUnchecked
        End If
        Else
        chkIsUsed.value = vbUnchecked
    End If
    
    If Not (IsNull(rs("TypeCar").value)) Then
    If rs("TypeCar").value = 1 Then
    Rd(1).value = True
    Else
    Rd(0).value = True
    End If
    Else
    Rd(0).value = True
    End If
    
    
           If mdifrmmain.MNUFixedAssets.Visible = True Then
            Rd(0).value = True
            Else
            Rd(1).value = True
            Rd(0).Visible = False
            End If
            
            
     TxtLeaderName.Text = IIf(IsNull(rs("LeaderName").value), "", rs("LeaderName").value)
    If Not (IsNull(rs("EmpType").value)) Then
    If rs("EmpType").value = 1 Then
    ChDrievType(1).value = True
    Else
    ChDrievType(0).value = True
    End If
    Else
    ChDrievType(0).value = True
    End If
    
    dtpEndContractDateH.Visible = True
    TxtContractNo.Visible = True
    dtpEndContractDate.Visible = True
    lbl(11).Visible = True
    lbl(10).Visible = True
    lbl(14).Visible = True

    Dim Str_Path As String
    Str_Path = App.path & "\" & SystemOptions.ImagesPath & "\cars\" & XPTxtID.Text & ".JPG"
    
     If Dir(Str_Path) <> "" Then
                DBPix201.ImageLoadFile (App.path & "\" & SystemOptions.ImagesPath & "\cars\" & XPTxtID.Text & ".JPG")
     Else
     DBPix201.ImageClear
                'Str_Path = App.path & "\" & SystemOptions.ImagesPath & "\DefaultCar.JPG"
                'If Dir(Str_Path) <> "" Then
                '        DBPix201.ImageLoadFile (App.path & "\" & SystemOptions.ImagesPath & "\cars\DefaultCar.JPG")
                'Else
                '        DBPix201.ImageClear
                'End If
     End If
    
     Total.Text = IIf(IsNull(rs("Total").value), "", rs("Total").value)
     LetterCount.Text = IIf(IsNull(rs("LetterCount").value), "", rs("LetterCount").value)
     LetterPrice.Text = IIf(IsNull(rs("LetterPrice").value), "", rs("LetterPrice").value)
        
    
    
    FormOrignal.Text = IIf(IsNull(rs("FormOrignal").value), "", rs("FormOrignal").value)
    authorizeLicense.Text = IIf(IsNull(rs("authorizeLicense").value), "", rs("authorizeLicense").value)
    authorizeExamination.Text = IIf(IsNull(rs("authorizeExamination").value), "", rs("authorizeExamination").value)
    cleaner.value = IIf(IsNull(rs("cleaner").value), False, rs("cleaner").value)
    sideMirror.value = IIf(IsNull(rs("sideMirror").value), False, rs("sideMirror").value)
    driverMirror.value = IIf(IsNull(rs("driverMirror").value), False, rs("driverMirror").value)
    InnerLights.value = IIf(IsNull(rs("InnerLights").value), False, rs("InnerLights").value)
    Pedals.value = IIf(IsNull(rs("Pedals").value), False, rs("Pedals").value)
    SunScreens.value = IIf(IsNull(rs("SunScreens").value), False, rs("SunScreens").value)
    Recorder.value = IIf(IsNull(rs("Recorder").value), False, rs("Recorder").value)
    Anntena.value = IIf(IsNull(rs("Anntena").value), False, rs("Anntena").value)
    Battery.value = IIf(IsNull(rs("Battery").value), False, rs("Battery").value)
    SpareTyre.value = IIf(IsNull(rs("SpareTyre").value), False, rs("SpareTyre").value)
    Crane.value = IIf(IsNull(rs("Crane").value), False, rs("Crane").value)
    CoverKey.value = IIf(IsNull(rs("CoverKey").value), False, rs("CoverKey").value)
    Guarantee.value = IIf(IsNull(rs("Guarantee").value), False, rs("Guarantee").value)
    Stickers.value = IIf(IsNull(rs("Stickers").value), False, rs("Stickers").value)
   ''/////////31 07 2016
   DcboCountryID2.BoundText = IIf(IsNull(rs("CounryID").value), 0, rs("CounryID").value)
   DcboGovernmentID.BoundText = IIf(IsNull(rs("CityID").value), 0, rs("CityID").value)
   TxtOwnerName.Text = IIf(IsNull(rs("OwnerName").value), "", rs("OwnerName").value)
   TxtOwnerName2.Text = IIf(IsNull(rs("OwnerName2").value), "", rs("OwnerName2").value)
   TxtTrackingNo.Text = IIf(IsNull(rs("TrackingNo").value), "", rs("TrackingNo").value)
   TxtAuthorType.Text = IIf(IsNull(rs("AuthorType").value), "", rs("AuthorType").value)
   
   TxtNatinality.Text = IIf(IsNull(rs("Natinality").value), "", rs("Natinality").value)
   txtJob.Text = IIf(IsNull(rs("Job").value), "", rs("Job").value)
   TxtDepartment.Text = IIf(IsNull(rs("Department").value), "", rs("Department").value)
   TxtDriLicenseNo.Text = IIf(IsNull(rs("DriLicenseNo").value), "", rs("DriLicenseNo").value)
   DriLicenseDate.value = IIf(IsNull(rs("DriLicenseDate").value), ToHijriDate(Date), rs("DriLicenseDate").value)
   TxtInsuranceNo.Text = IIf(IsNull(rs("InsuranceNO").value), "", rs("InsuranceNO").value)
If Not IsNull(rs("Keys").value) Then
    If (rs("Keys").value) = True Then
    keys.value = vbChecked
    Else
    keys.value = vbUnchecked
    End If
   Else
   keys.value = vbUnchecked
End If
   
   
   If Not IsNull(rs("Insurance").value) Then
    If (rs("Insurance").value) = True Then
    Insurance.value = vbChecked
    Else
    Insurance.value = vbUnchecked
    End If
   Else
   Insurance.value = vbUnchecked
   End If
 If Not IsNull(rs("Authorization2").value) Then
    If (rs("Authorization2").value) = True Then
       Authorization.value = vbChecked
    Else
        Authorization.value = vbUnchecked
    End If
 Else
      Insurance.value = vbUnchecked
 End If
 
If Not IsNull(rs("Licenses").value) Then
    If (rs("Licenses").value) = True Then
       Licenses.value = vbChecked
    Else
        Licenses.value = vbUnchecked
    End If
 Else
      Licenses.value = vbUnchecked
 End If
 ''////
 If Not IsNull(rs("Exam").value) Then
    If (rs("Exam").value) = True Then
       Exam.value = vbChecked
    Else
        Exam.value = vbUnchecked
    End If
 Else
      Exam.value = vbUnchecked
 End If
 ''/////
 If Not IsNull(rs("KeyReserve").value) Then
    If (rs("KeyReserve").value) = True Then
       KeyReserve.value = vbChecked
    Else
        KeyReserve.value = vbUnchecked
    End If
 Else
      KeyReserve.value = vbUnchecked
 End If
 ''
  If Not IsNull(rs("Receipt").value) Then
    If (rs("Receipt").value) = True Then
       Receipt.value = vbChecked
    Else
        Receipt.value = vbUnchecked
    End If
 Else
      Receipt.value = vbUnchecked
 End If
''////
  If Not IsNull(rs("Triangle").value) Then
    If (rs("Triangle").value) = True Then
       Triangle.value = vbChecked
    Else
        Triangle.value = vbUnchecked
    End If
 Else
      Triangle.value = vbUnchecked
 End If
''///////
  If Not IsNull(rs("TrackingDevice").value) Then
    If (rs("TrackingDevice").value) = True Then
       TrackingDevice.value = vbChecked
    Else
        TrackingDevice.value = vbUnchecked
    End If
 Else
      TrackingDevice.value = vbUnchecked
 End If
''/////


  If Not IsNull(rs("IsUsed").value) Then
    If (rs("IsUsed").value) = True Then
       chkIsUsed.value = vbChecked
    Else
        chkIsUsed.value = vbUnchecked
    End If
 Else
      chkIsUsed.value = vbUnchecked
 End If

  If Not IsNull(rs("Sabt").value) Then
    If (rs("Sabt").value) = True Then
       Sabt.value = vbChecked
    Else
        Sabt.value = vbUnchecked
    End If
 Else
      Sabt.value = vbUnchecked
 End If



  If Not IsNull(rs("Chains").value) Then
    If (rs("Chains").value) = True Then
       Chains.value = vbChecked
    Else
        Chains.value = vbUnchecked
    End If
 Else
      Chains.value = vbUnchecked
 End If



  If Not IsNull(rs("Kafla").value) Then
    If (rs("Kafla").value) = True Then
       Kafla.value = vbChecked
    Else
        Kafla.value = vbUnchecked
    End If
 Else
      Kafla.value = vbUnchecked
 End If



  If Not IsNull(rs("Hock").value) Then
    If (rs("Hock").value) = True Then
       Hock.value = vbChecked
    Else
        Hock.value = vbUnchecked
    End If
 Else
      Hock.value = vbUnchecked
 End If


  If Not IsNull(rs("Khabor").value) Then
    If (rs("Khabor").value) = True Then
       Khabor.value = vbChecked
    Else
        Khabor.value = vbUnchecked
    End If
 Else
      Khabor.value = vbUnchecked
 End If


  If Not IsNull(rs("Sail").value) Then
    If (rs("Sail").value) = True Then
       Sail.value = vbChecked
    Else
        Sail.value = vbUnchecked
    End If
 Else
      Sail.value = vbUnchecked
 End If



  If Not IsNull(rs("SideBarriers").value) Then
    If (rs("SideBarriers").value) = True Then
       SideBarriers.value = vbChecked
    Else
        SideBarriers.value = vbUnchecked
    End If
 Else
      SideBarriers.value = vbUnchecked
 End If


  If Not IsNull(rs("SideFrame").value) Then
    If (rs("SideFrame").value) = True Then
       SideFrame.value = vbChecked
    Else
        SideFrame.value = vbUnchecked
    End If
 Else
      SideFrame.value = vbUnchecked
 End If




  If Not IsNull(rs("FireExtingui").value) Then
    If (rs("FireExtingui").value) = True Then
       FireExtingui.value = vbChecked
    Else
        FireExtingui.value = vbUnchecked
    End If
 Else
      FireExtingui.value = vbUnchecked
 End If
 ''///////
    If Not IsNull(rs("SubsBattery").value) Then
    If (rs("SubsBattery").value) = True Then
       SubsBattery.value = vbChecked
    Else
        SubsBattery.value = vbUnchecked
    End If
 Else
      SubsBattery.value = vbUnchecked
 End If
 
   If Not IsNull(rs("BagAmbulance").value) Then
    If (rs("BagAmbulance").value) = True Then
       BagAmbulance.value = vbChecked
    Else
        BagAmbulance.value = vbUnchecked
    End If
 Else
      BagAmbulance.value = vbUnchecked
 End If
   If Not IsNull(rs("DriLicense").value) Then
    If (rs("DriLicense").value) = True Then
       DriLicense.value = vbChecked
    Else
        DriLicense.value = vbUnchecked
    End If
 Else
      DriLicense.value = vbUnchecked
 End If
 FillAcced
 Retrive_CarsExp
 '#################### Khaled ###############################
 fillComboParts
 Retrive_CarParts
 
    Exit Sub
ErrTrap:
End Sub
 

Private Sub txtNum1_KeyPress(KeyAscii As Integer)
txtNum1.Text = ""
If Len(txtNum1.Text) > 0 Then
KeyAscii = 0
End If
If Not (CHR(KeyAscii) >= 0 And CHR(KeyAscii) <= 9) Then
KeyAscii = 0
Else
        txtNum2.SetFocus
End If
Cal_Board
End Sub

Private Sub txtNum1_KeyUp(KeyCode As Integer, Shift As Integer)
Cal_Board
End Sub

Private Sub txtNum2_KeyPress(KeyAscii As Integer)
txtNum2.Text = ""
If Len(txtNum2.Text) > 0 Then
KeyAscii = 0
End If
If Not (CHR(KeyAscii) >= 0 And CHR(KeyAscii) <= 9) Then
KeyAscii = 0
Else
        txtNum3.SetFocus
End If
Cal_Board
End Sub

Private Sub txtNum2_KeyUp(KeyCode As Integer, Shift As Integer)
Cal_Board
End Sub

Private Sub txtNum3_KeyPress(KeyAscii As Integer)
txtNum3.Text = ""
If Len(txtNum3.Text) > 0 Then
KeyAscii = 0
End If
If Not (CHR(KeyAscii) >= 0 And CHR(KeyAscii) <= 9) Then
KeyAscii = 0
Else
        txtNum4.SetFocus
End If
Cal_Board
End Sub

Private Sub txtNum3_KeyUp(KeyCode As Integer, Shift As Integer)
Cal_Board
End Sub

Private Sub txtNum4_KeyPress(KeyAscii As Integer)
txtNum4.Text = ""
If Len(txtNum4.Text) > 0 Then
KeyAscii = 0
End If
If Not (CHR(KeyAscii) >= 0 And CHR(KeyAscii) <= 9) Then
KeyAscii = 0
End If

Cal_Board
End Sub


Private Sub txtNum4_KeyUp(KeyCode As Integer, Shift As Integer)
Cal_Board
End Sub

Private Sub txtRep_Change()
txtMax.Text = val(TxtCapacity.Text) * val(txtRep.Text)
End Sub

Private Sub txtSetCount_Change()
If IsNumeric(txtSetCount.Text) Then
TxtCapacity.Text = val(txtSetCount.Text) * val(txtRate.Caption)
Else
txtSetCount.Text = ""
End If
End Sub

Private Sub TxtVlue_KeyPress(KeyAscii As Integer)
KeyAscii = KeyAscii_Num(KeyAscii, Me.TxtVlue.Text, 0)
End Sub

Private Sub VColor_KeyUp(KeyCode As Integer, Shift As Integer)
Dim str As String
 If KeyCode = vbKeyF5 Then
 If SystemOptions.UserInterface = ArabicInterface Then
    str = " select id , name from tblcolor "
    Else
    str = " select id , namee from tblcolor "
    End If
    fill_combo VColor, str
End If


End Sub

Private Sub VehicleLong_KeyPress(KeyAscii As Integer)
    KeyAscii = KeyAscii_Num(KeyAscii, VehicleLong.Text, 0)

End Sub

Private Sub DcEmployee_Change()
Dim Nationality As String
Dim JobTypeID As Double
Dim DepartmentID As Double
Dim DriverLicense As String
Dim DriverLicenseendH As String
    If val(DcEmployee.BoundText) = 0 Then Exit Sub
      Dim EmpCode  As String
      
      GetEmployeeIDFromCode , , DcEmployee.BoundText, EmpCode
      Text6.Text = EmpCode
       If Me.TxtModFlg = "R" Then Exit Sub
        get_employee_information val(Me.DcEmployee.BoundText), , DepartmentID, , JobTypeID, , , , , Nationality, , , , , , , , , , , , , , , , , , , , , , , , , , , , , , DriverLicenseendH, DriverLicense
        TxtNatinality.Text = Nationality
        Me.DcbDept.BoundText = DepartmentID
        Me.DcbJob.BoundText = JobTypeID
        txtJob.Text = Me.DcbJob.Text
        TxtDepartment.Text = Me.DcbDept.Text
        TxtDriLicenseNo.Text = DriverLicense
        DriLicenseDate.value = DriverLicenseendH
        
End Sub

Private Sub DcEmployee_Click(Area As Integer)
DcEmployee_Change
End Sub

Private Sub Text6_KeyPress(KeyAscii As Integer)
  Dim EmpID As Integer
    If KeyAscii = vbKeyReturn Then
        GetEmployeeIDFromCode Text6.Text, EmpID
        DcEmployee.BoundText = EmpID
    End If
End Sub


Private Sub VModel_KeyUp(KeyCode As Integer, Shift As Integer)
Dim str As String
 If KeyCode = vbKeyF5 Then
    If SystemOptions.UserInterface = ArabicInterface Then
            str = " select id , Model  from TblCarModels"
    Else
          str = " select id , ModelE  from TblCarModels"
    End If
    fill_combo VModel, str
    
    End If
    
End Sub

Private Sub VSFlexGrid1_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
Cancel = True
End Sub

Private Sub VSFlexGrid1_Click()
    With Me.VSFlexGrid1
        If .Row > 0 Then
            txtCode1.Text = .TextMatrix(.Row, .ColIndex("id"))
            Me.DcbExpenSiv.BoundText = val(.TextMatrix(.Row, .ColIndex("ExpID")))
            Me.TxtVlue.Text = val(.TextMatrix(.Row, .ColIndex("Vlue")))
            Me.TxtRemarks.Text = .TextMatrix(.Row, .ColIndex("Remarks"))
        End If
    End With
End Sub

Private Sub XPBtnMove_Click(Index As Integer)
    'On Error GoTo ErrTrap
 
    If Me.TxtModFlg.Text = "N" Then
        clear_all Me
        Me.TxtModFlg.Text = "R"
        XPBtnMove_Click (1)
    End If

    Select Case Index

        Case 0

            If Not (rs.EOF Or rs.BOF) Then
                rs.MovePrevious

                If rs.BOF Then rs.MoveFirst
            End If

        Case 1

            If Not (rs.EOF Or rs.BOF) Then
                rs.MoveFirst
            End If

        Case 2

            If Not (rs.EOF Or rs.BOF) Then
                rs.MoveLast
            End If

        Case 3

            If Not (rs.EOF Or rs.BOF) Then
                rs.MoveNext

                If rs.EOF Then rs.MoveLast
            End If

    End Select

    Retrive
    Exit Sub
ErrTrap:
End Sub
Sub SaveFixed()
Dim sql As String
Dim StrSQL As String
'StrSQL = "Delete From FixedAssets Where CarsDataID =" & val(Me.XPTxtID.text) & "  And FlgCarNotFixed = 1"
'           Cn.Execute StrSQL, , adExecuteNoRecords
Dim Rs5 As ADODB.Recordset
Set Rs5 = New ADODB.Recordset
Dim FexdID As Double

If Me.TxtModFlg.Text = "N" Then
sql = "Select * from FixedAssets where 1=-1"
Rs5.Open sql, Cn, adOpenKeyset, adLockOptimistic, adCmdText
FexdID = CStr(new_id("FixedAssets", "id", "", True))
Rs5.AddNew
Else
FexdID = val(DcFixedAssets.BoundText)
If FexdID = 0 Then
sql = "Select * from FixedAssets where 1=-1"
Rs5.Open sql, Cn, adOpenKeyset, adLockOptimistic, adCmdText
FexdID = CStr(new_id("FixedAssets", "id", "", True))
Rs5.AddNew

Else
sql = "Select * from FixedAssets where id=" & FexdID & ""
Rs5.Open sql, Cn, adOpenKeyset, adLockOptimistic, adCmdText
End If
End If
If FexdID = 0 Then Exit Sub
Rs5("id").value = FexdID
Rs5("CarsDataID").value = val(XPTxtID.Text)
Rs5("FlgCarNotFixed").value = 1
Rs5("ISEQUP").value = 1
Rs5("HaveDepreciation").value = 1
Rs5("PurchasePrice").value = 1
Rs5("branch_no").value = val(Dcbranch.BoundText)
Rs5("NameE").value = TxtEqupName.Text
Rs5("Name").value = TxtEqupName.Text
Rs5("code").value = TxtId.Text

 Rs5("isUsed").value = IIf(chkIsUsed.value = vbChecked, 1, 0)

Rs5.update
sql = "Update TblCarsData set fixedAssetid=" & FexdID & "  where id =" & val(XPTxtID.Text) & ""
Cn.Execute sql
SaveAssest FexdID

End Sub
Private Sub DcbAccount_Change()
DcbAccount_Click (0)
End Sub

Private Sub DcbAccount_KeyUp(KeyCode As Integer, Shift As Integer)
 If KeyCode = vbKeyF3 Then

        Unload Account_search
        Account_search.show
        Account_search.case_id = 90519
        

        
            
    End If
    
    
End Sub




Private Sub SaveData()
    Dim sql As String
    Dim TblCarKMFOLLOWid As Double
    Dim Msg As String
    Dim StrSQL As String
    Dim RsTemp As New ADODB.Recordset
    Dim RsTempM As New ADODB.Recordset
    Dim RsDev As New ADODB.Recordset
    Dim RsNot As New ADODB.Recordset

    Dim BeginTrans As Boolean
'    On Error GoTo ErrTrap

    If Me.TxtModFlg.Text <> "R" Then
 
 If SystemOptions.CarsRevenuePerOwner = True Then
 
If DcbAccount.BoundText = "" Or DcbAccount.Text = "" Then
        If SystemOptions.UserInterface = ArabicInterface Then
                    MsgBox "Ì—ÃÏ «Œ Ì«— «·Õ”«»"
        Else
                 MsgBox "Please Select Account"
         End If
                 DcbAccount.SetFocus
        Exit Sub
End If

End If



 
 
 
 If Rd(0).value = True Then
        If DcFixedAssets.Text = "" And val(DcFixedAssets.BoundText) = 0 Then
            If SystemOptions.UserInterface = ArabicInterface Then
                MsgBox "Õœœ «”„ «·„⁄œÂ/«·”Ì«—…  «Ê·«", vbCritical
            Else
                MsgBox "Select Name Firstly    ", vbCritical
            End If
    
            DcFixedAssets.SetFocus
            Exit Sub
        End If
End If
    
    
    
           If TxtBoardNO.Text = "" Then
            If SystemOptions.UserInterface = ArabicInterface Then
                MsgBox "Õœœ  —ﬁ„ «··ÊÕ… «Ê·«    ", vbCritical
            Else
                MsgBox "Select   Board No firstly    ", vbCritical
            End If
    
'            TxtBoardNO.SetFocus
            Exit Sub
        End If
      If Me.TxtModFlg.Text = "N" Then
     If ChDrievType(1).value = True Then
   If chekEmoloyee(TxtLeaderName.Text) = True Then
   If SystemOptions.UserInterface = ArabicInterface Then
   MsgBox "«”„ «·„ÊŸ› „ÊÃÊœ „”»ﬁ« Ì—ÃÏ «œŒ«· «”„ «Œ—"
   Else
   MsgBox "This is Name of Employee Already Exists"
   End If
   Exit Sub
   End If
    End If
End If
        If val(Me.Dcbranch.BoundText) = 0 Then
            If SystemOptions.UserInterface = ArabicInterface Then
                MsgBox "Õœœ «·›—⁄ «Ê·«", vbCritical
            Else
                MsgBox "Select Branch Firstly    ", vbCritical
            End If

            Dcbranch.SetFocus
            SendKeys "{F4}"
            Exit Sub
        End If
 
        If val(Me.DCGroup.BoundText) = 0 Then
            If SystemOptions.UserInterface = ArabicInterface Then
                MsgBox "Õœœ ‰Ê⁄ «·„⁄œÂ/«·”Ì«—…  «Ê·«", vbCritical
            Else
                MsgBox "Select Group Firstly    ", vbCritical
            End If

            DCGroup.SetFocus
             SendKeys "{F4}"
            Exit Sub
        End If
    
        If val(Me.DcEmployee.BoundText) = 0 Then
'            If SystemOptions.UserInterface = ArabicInterface Then
'                MsgBox "Õœœ   «·”Ì«— … »⁄ÂœÂ  «Ê·«", vbCritical
'            Else
'                MsgBox "Select Holder Name   ", vbCritical
'            End If
'
'            DcEmployee.SetFocus
'            SendKeys "{F4}"
'            Exit Sub
        End If
 
        Select Case Me.TxtModFlg.Text

            Case "N"

                StrSQL = "select * From  TblCarsData where Name='" & Trim(TxtName.Text) & "'"
                RsTemp.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText

                If RsTemp.RecordCount > 0 Then
                If SystemOptions.UserInterface = ArabicInterface Then
                    Msg = "Â‰«ﬂ „⁄œ…  „”Ã· „”»ﬁ« »Â–« «·«”„" & CHR(13)
                    Msg = Msg + "»—Ã«¡ «· √ﬂœ „‰ «·«”„ «·’ÕÌÕ " & CHR(13)
                    Msg = Msg + "√Ê  €ÌÌ— √Ê  „ÌÌ“ «·„⁄œÂ/«·”Ì«—… «·„Õœœ"
                Else
                     Msg = "There Equepment already registered by that name" & CHR(13)
                              End If
                MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
                   TxtName.SetFocus
                    Exit Sub
                End If

            Case "E"
        
        End Select

      '  If Me.TxtModFlg.text = "N" Then
   
      '  End If

        Cn.BeginTrans
        BeginTrans = True

        Select Case Me.TxtModFlg.Text

            Case "N"
            
                XPTxtID.Text = CStr(new_id("TblCarsData", "id", "", True))
            
                rs.AddNew
            
            Case "E"
                Cn.Execute "delete TblCarMaintenancePlan where CarId=" & val(Me.XPTxtID.Text)
                Cn.Execute "delete TblEmpAsest where CrsID=" & val(Me.XPTxtID.Text)
        Cn.Execute "delete TblEmpAsestDetails where CrsID=" & val(Me.XPTxtID.Text)
        
                 rs("ContractID").value = IIf(Not IsNumeric(TxtContractNo.Text), Null, val(TxtContractNo.Text))
                 rs("EndContractDate").value = dtpEndContractDate.value
                 rs("EndContractDateH").value = dtpEndContractDateH.value
                 
                 
               
                        
        End Select

        rs("id").value = val(Me.XPTxtID.Text)
rs("AccountPaym").value = IIf(Trim(DcbAccount.BoundText) = "", Null, DcbAccount.BoundText)
        
        rs("Branch_NO").value = IIf(val(Dcbranch.BoundText) = 0, Null, Dcbranch.BoundText)
        rs("fixedAssetid").value = IIf(val(DcFixedAssets.BoundText) = 0, Null, val(DcFixedAssets.BoundText))
        
        rs("code").value = TxtId.Text
        rs("Fullcode").value = IIf(DCPreFix.BoundText = "", Null, DCPreFix.Text) & IIf(Trim(TxtId.Text) = "", Null, TxtId.Text)
        rs("prifix").value = IIf(DCPreFix.Text = "", Null, DCPreFix.Text)
        rs("CarsTypeId").value = IIf(val(DCGroup.BoundText) = 0, Null, DCGroup.BoundText)
        '
        rs("LicenseNO").value = IIf(Trim(TxtLicenseNO.Text) = "", Null, TxtLicenseNO.Text)
 
        rs("Name").value = IIf(Trim(TxtName.Text) = "", Null, TxtName.Text)
        rs("BoardNO").value = IIf(Trim(TxtBoardNO.Text) = "", Null, TxtBoardNO.Text)
        rs("Model").value = IIf(Trim(TxtModel.Text) = "", Null, TxtModel.Text)
        rs("VehicleLong").value = IIf(Trim(VehicleLong.Text) = "", Null, VehicleLong.Text)
        rs("StutsID").value = IIf(val(Me.DcbStuts.ListIndex) = -1, Null, val(Me.DcbStuts.ListIndex))
        'VehicleLong
        
        
        If Rd(1).value = True Then
           rs("TypeCar").value = 1
        Else
           rs("TypeCar").value = 0
        End If
        If ChDrievType(1).value = True Then
           rs("EmpType").value = 1
        Else
           rs("EmpType").value = Null
        End If
        rs("LeaderName").value = TxtLeaderName.Text
        rs("EqupName").value = TxtEqupName.Text

        rs("PurchaseDate").value = DpPurchaseDate.value
        rs("ExpireDate").value = DpExpireDate.value
        rs("SensitiveWeightDate").value = DpSensitiveWeightDate.value
        
 
        
        rs("LastKMCounter").value = IIf(val(txtLastKMCounter.Text) = 0, 0, val(txtLastKMCounter.Text))
        rs("InsuranceCompanyId").value = IIf(val(DCInsuranceCompanyId.BoundText) = 0, Null, DCInsuranceCompanyId.BoundText)
        
        rs("Emp_id").value = IIf(val(DcEmployee.BoundText) = 0, Null, DcEmployee.BoundText)
       
        'ToGregorianDate '
        rs("LicenseExpireDateH").value = DpLicenseExpireDateH.value
        rs("LicenseExpireDate").value = ToGregorianDate(DpLicenseExpireDateH.value)
       
        rs("InsuranceExpireDateH").value = DpInsuranceExpireDateH.value
        rs("InsuranceExpireDate").value = ToGregorianDate(DpInsuranceExpireDateH.value)
         
        rs("TestExpireDateh").value = DpTestExpireDateH.value
        rs("TestExpireDate").value = ToGregorianDate(DpTestExpireDateH.value)
                   
        rs("Notes").value = IIf(Trim(TxtNotes.Text) = "", Null, TxtNotes.Text)
        rs("EquQty").value = IIf(val(TxtEquQty.Text) = 0, Null, TxtEquQty.Text)
        
        rs("capacity").value = IIf(Not IsNumeric(TxtCapacity.Text), Null, val(TxtCapacity.Text))
        rs("SetCount").value = IIf(txtSetCount.Text = "", Null, val(txtSetCount.Text))
        rs("Rate").value = IIf(txtRate.Caption = "", Null, val(txtRate.Caption))
        rs("rep").value = IIf(txtRep.Text = "", 1, val(txtRep.Text))
        rs("MaxCap").value = IIf(txtMax.Text = "", 1, val(txtMax.Text))
        rs("OperatorN").value = IIf(TxtOperatorN.Text = "", "", TxtOperatorN.Text)
        
        rs("Gearno").value = IIf(TxtGearno.Text = "", "", TxtGearno.Text)
        rs("Gearno1").value = IIf(TxtGearno1.Text = "", "", TxtGearno1.Text)
        
        rs("Machineno").value = IIf(TxtMachineno.Text = "", "", TxtMachineno.Text)
        rs("Machineno1").value = IIf(TxtMachineno1.Text = "", "", TxtMachineno1.Text)
            
            
        rs("Chesis").value = IIf(Chesis.Text = "", "", Chesis.Text)
        rs("VColor").value = IIf(VColor.BoundText = "", Null, VColor.BoundText)
        rs("VModel").value = IIf(VModel.BoundText = "", Null, VModel.BoundText)
        rs("VType").value = IIf(VType.BoundText = "", Null, VType.BoundText)
             
        rs("LocationID").value = IIf(LocationID.BoundText = "", Null, LocationID.BoundText)
        
        rs("Total").value = val(Total.Text)
        rs("LetterCount").value = val(LetterCount.Text)
        rs("LetterPrice").value = val(LetterPrice.Text)
       
       
        rs("FormOrignal").value = IIf(FormOrignal.Text = "", "", FormOrignal.Text)
        rs("authorizeLicense").value = IIf(authorizeLicense.Text = "", "", authorizeLicense.Text)
        rs("authorizeExamination").value = IIf(authorizeExamination.Text = "", "", authorizeExamination.Text)
             
        rs("cleaner").value = cleaner.value
        rs("sideMirror").value = sideMirror.value
        rs("driverMirror").value = driverMirror.value
        rs("InnerLights").value = InnerLights.value
        rs("Pedals").value = Pedals.value
        rs("SunScreens").value = SunScreens.value
        rs("Recorder").value = Recorder.value
        rs("Anntena").value = Anntena.value
        rs("Battery").value = Battery.value
        rs("SpareTyre").value = SpareTyre.value
        rs("Crane").value = Crane.value
        rs("CoverKey").value = CoverKey.value
        rs("Guarantee").value = Guarantee.value
        rs("Stickers").value = Stickers.value
        ''''/31 07 2016
        rs("CounryID").value = val(DcboCountryID2.BoundText)
        rs("CityID").value = val(DcboGovernmentID.BoundText)
        rs("OwnerName").value = TxtOwnerName.Text
        rs("OwnerName2").value = TxtOwnerName2.Text
        rs("TrackingNo").value = TxtTrackingNo.Text
        rs("AuthorType").value = TxtAuthorType.Text
        rs("AuthorDate").value = AuthorDate.value
        
        rs("Natinality").value = TxtNatinality.Text
        rs("Job").value = txtJob.Text
        rs("Department").value = TxtDepartment.Text
        rs("DriLicenseNo").value = TxtDriLicenseNo.Text
        rs("DriLicenseDate").value = DriLicenseDate.value
        rs("InsuranceNO").value = TxtInsuranceNo.Text
         If DriLicense.value = vbChecked Then
            rs("DriLicense").value = 1
        Else
            rs("DriLicense").value = 0
        End If
        
        If Insurance.value = vbChecked Then
            rs("Insurance").value = 1
        Else
            rs("Insurance").value = 0
        End If
        If Authorization.value = vbChecked Then
            rs("Authorization2").value = 1
        Else
            rs("Authorization2").value = 0
            
        End If
        If Licenses.value = vbChecked Then
            rs("Licenses").value = 1
        Else
            rs("Licenses").value = 0
        End If
        If Exam.value = vbChecked Then
            rs("Exam").value = 1
        Else
            rs("Exam").value = 0
        End If
       
       If KeyReserve.value = vbChecked Then
            rs("KeyReserve").value = 1
       Else
            rs("KeyReserve").value = 0
       End If
        
        If Receipt.value = vbChecked Then
            rs("Receipt").value = 1
        Else
            rs("Receipt").value = 0
       End If
      If Triangle.value = vbChecked Then
         rs("Triangle").value = 1
      Else
        rs("Triangle").value = 0
       End If
    If TrackingDevice.value = vbChecked Then
        rs("TrackingDevice").value = 1
    Else
        rs("TrackingDevice").value = 0
     End If
    If FireExtingui.value = vbChecked Then
        rs("FireExtingui").value = 1
    Else
        rs("FireExtingui").value = 0
     End If
   rs("ExpireDateH").value = DpExpireDateH.value
   rs("SensitiveWeightDateH").value = DpSensitiveWeightDateH.value


     
       
If chkIsUsed.value = vbChecked Then
    rs("IsUsed").value = 1
Else
    rs("IsUsed").value = 0
End If

If Sabt.value = vbChecked Then
    rs("Sabt").value = 1
Else
    rs("Sabt").value = 0
End If

If Chains.value = vbChecked Then
    rs("Chains").value = 1
Else
    rs("Chains").value = 1
End If

If Kafla.value = vbChecked Then
    rs("Kafla").value = 1
Else
    rs("Kafla").value = 0
End If

If Hock.value = vbChecked Then
    rs("Hock").value = 1
Else
    rs("Hock").value = 0
End If

If Khabor.value = vbChecked Then
    rs("Khabor").value = 1
Else
    rs("Khabor").value = 0
End If

If Sail.value = vbChecked Then
    rs("Sail").value = 1
Else
    rs("Sail").value = 0
End If

If SideBarriers.value = vbChecked Then
    rs("SideBarriers").value = 1
Else
    rs("SideBarriers").value = 0
End If

If SideFrame.value = vbChecked Then
    rs("SideFrame").value = 1
Else
    rs("SideFrame").value = 0
End If

   
     
   If SubsBattery.value = vbChecked Then
        rs("SubsBattery").value = 1
    Else
        rs("SubsBattery").value = 0
     End If
   If BagAmbulance.value = vbChecked Then
        rs("BagAmbulance").value = 1
    Else
        rs("BagAmbulance").value = 0
     End If
  If keys.value = vbChecked Then
        rs("Keys").value = 1
    Else
        rs("Keys").value = 0
   End If
     
        rs.update
    End If
 If Rd(1).value = True Then
 SaveFixed
 End If
  If ChDrievType(1).value = True Then
SaveEmployee
End If
    '**************************************************************************
rs.Resync

    Cn.CommitTrans
    BeginTrans = False
    XPTxtCurrent.Caption = rs.AbsolutePosition
    XPTxtCount.Caption = rs.RecordCount
    CuurentLogdata
        DBPix201.ImageSaveFile (App.path & "\" & SystemOptions.ImagesPath & "\cars\" & XPTxtID.Text & ".JPG")
        
    Select Case Me.TxtModFlg.Text

        Case "N"
        If SystemOptions.UserInterface = ArabicInterface Then
            Msg = "  „ Õ›Ÿ »Ì«‰«  Â–« «·‰Ê⁄" & CHR(13)
            Msg = Msg + "Â·  —€» ›Ì ≈÷«›… »Ì«‰«  √Œ—Ì"
        Else
            Msg = " This is record already saved " & CHR(13)
            Msg = Msg + "You want enter another record"
          End If
            If MsgBox(Msg, vbYesNo + vbQuestion + vbMsgBoxRight + vbMsgBoxRtlReading + vbDefaultButton2, App.title) = vbYes Then
                Cmd_Click (0)
                Exit Sub
            End If

            'Õ›Ÿ «Œ— ﬁ—«¡… ··⁄œ«œ
            '    If txtLastKMCounter.text = 0 Then
            '    TblCarKMFOLLOWid = val(CStr(new_id("TblCarKMFOLLOW", "TblCarKMFOLLOWid", "", True)))
            '    sql = "insert into TblCarKMFOLLOW (TblCarKMFOLLOWid,CarId,beforeKM,CurrentKM,RecordDate) "
            '    sql = sql & " Values(" & TblCarKMFOLLOWid & "," & val(XPTxtID.text) & ",0," & val(txtLastKMCounter.text) & ",'" & SQLDate(Date) & "')"
            '    End If
            '    Cn.Execute sql
           
        Case "E"
            '     If txtLastKMCounter.text = 0 Then
            '                TblCarKMFOLLOWid = val(CStr(new_id("TblCarKMFOLLOW", "TblCarKMFOLLOWid", "", True)))
            '    sql = "insert into TblCarKMFOLLOW (TblCarKMFOLLOWid,CarId,beforeKM,CurrentKM,RecordDate) "
            '    sql = sql & " Values(" & TblCarKMFOLLOWid & "," & val(XPTxtID.text) & ",0," & val(txtLastKMCounter.text) & ",'" & SQLDate(Date) & "')"
            '    Cn.Execute sql
            '    End If
            If SystemOptions.UserInterface = ArabicInterface Then
            MsgBox " „ Õ›Ÿ Â–Â «· ⁄œÌ·« ", vbInformation + vbOKOnly + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
            Else
            MsgBox "Saved successfully", vbInformation + vbOKOnly + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
          End If
    End Select
    
    TxtModFlg.Text = "R"
 Retrive
    Exit Sub
ErrTrap:

    If BeginTrans = True Then
        BeginTrans = False
        Cn.RollbackTrans
    End If

    If rs.EditMode <> adEditNone Then
        rs.CancelUpdate
    End If

    If Err.Number = -2147217900 Then
    If SystemOptions.UserInterface = ArabicInterface Then
        Msg = "·« Ì„ﬂ‰ Õ›Ÿ Â–Â «·»Ì«‰«  " & CHR(13)
        Msg = Msg + "·ﬁœ  „ «œŒ«· ﬁÌ„ €Ì— ’«·Õ… " & CHR(13)
        Msg = Msg + " √ﬂœ „‰ œﬁ… «·»Ì«‰«  Ê√⁄œ «·„Õ«Ê·…"
    Else
        Msg = "Can not Save  " & CHR(13)
        Msg = Msg + "It was invalid input data " & CHR(13)
        Msg = Msg + "Make sure you retry data"
    End If
        MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
        Exit Sub
    End If
If SystemOptions.UserInterface = ArabicInterface Then
    Msg = "⁄›Ê«...ÕœÀ Œÿ√ „« √À‰«¡ Õ›Ÿ Â–Â «·»Ì«‰«  " & CHR(13)
  Else
   Msg = "Sorry... error douring save " & CHR(13)
  End If
    MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
End Sub

Private Sub Undo()
    On Error GoTo ErrTrap

    Select Case TxtModFlg.Text

        Case "N"
            clear_all Me
            Me.TxtModFlg.Text = "R"
            XPBtnMove_Click (1)

        Case "E"
            rs.find "ID=" & val(XPTxtID.Text) & "", , adSearchForward, adBookmarkFirst

            If rs.EOF Or rs.BOF Then
                Me.TxtModFlg.Text = "R"
                Exit Sub
            End If
Me.TxtModFlg.Text = "R"
            Retrive
            
    End Select

    Exit Sub
ErrTrap:
End Sub

Private Sub Del_AssetType()
    Dim msgstr  As String
Dim StrSQL As String
    Dim sql As String

    Dim Msg As String
    On Error GoTo ErrTrap

    If XPTxtID.Text <> "" Then
    If SystemOptions.UserInterface = ArabicInterface Then
        Msg = "”Ì „ Õ–› »Ì«‰«  «·⁄„·Ì… —ﬁ„ " & CHR(13)
        Msg = Msg + (Me.XPTxtID.Text) & CHR(13)
        Msg = Msg + " Â·  —€» ›Ì Õ–› Â–Â «·»Ì«‰« ø"
Else
Msg = "Confirm Delete"
End If
        If MsgBox(Msg, vbYesNo + vbQuestion + vbMsgBoxRight + vbMsgBoxRtlReading, App.title) = vbYes Then
            If Not rs.RecordCount < 1 Then
                '       sql = "Delete   from notes where NoteID=" & Val(txtNoteID.text)
                '        Cn.Execute sql
        
                ' sql = "delete From DOUBLE_ENTREY_VOUCHERS1 where opening_balance_voucher_id=" & Val(txtopening_balance_voucher_id.text)
                '    Cn.Execute sql, , adExecuteNoRecords
                '   sql = "delete  FixedAssetInstallmentsDetails where FixedAssetID=" & Val(Me.XPTxtID.text)
                '   Cn.Execute sql, , adExecuteNoRecords
                
                CuurentLogdata ("D")
                StrSQL = "Delete From TblAssestes Where CarsDataID =" & val(Me.XPTxtID.Text) & "  And FlgCarNotFixed = 1"
              Cn.Execute StrSQL, , adExecuteNoRecords
               Cn.Execute "delete TblEmpAsest where CrsID=" & val(Me.XPTxtID.Text)
        Cn.Execute "delete TblEmpAsestDetails where CrsID=" & val(Me.XPTxtID.Text)
                StrSQL = "Delete From FixedAssets Where CarsDataID =" & val(Me.XPTxtID.Text) & "  And FlgCarNotFixed = 1"
              Cn.Execute StrSQL, , adExecuteNoRecords
                rs.delete
                rs.MoveFirst

                If rs.RecordCount < 1 Then
                    clear_all Me
                    TxtModFlg_Change
             
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
        Msg = "This process is not available where there was no records"
        End If
        MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
        TxtModFlg_Change
        Exit Sub
    End If

    TxtModFlg_Change
    Exit Sub
ErrTrap:
If SystemOptions.UserInterface = ArabicInterface Then
    Msg = "⁄›Ê«...ÕœÀ Œÿ√ „« √À‰«¡ Õ–› Â–Â «·»Ì«‰«  " & CHR(13)
 Else
 Msg = "Sorry...error douring delete " & CHR(13)
 End If
    MsgBox Msg, vbMsgBoxRight + vbMsgBoxRtlReading + vbExclamation, App.title
    rs.CancelUpdate

End Sub
 
Private Sub AddTip()
    Dim Wrap As String
    On Error GoTo ErrTrap
    Set TTP = New clstooltip
    Wrap = CHR(13) + CHR(10)

    If SystemOptions.UserInterface = ArabicInterface Then

        With TTP
            .Create Me.hwnd, "√‰Ê«⁄ «·„’—Ê›« ", 1, 15204351, -2147483630
            .MaxWidth = 4000
            .VisibleTime = 9000
            .DelayTime = 600
            .AddControl Cmd(0), "ÃœÌœ ..." & Wrap & "·«÷«›… »Ì«‰«  ‰Ê⁄ ÃœÌœ" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
        End With

        With TTP
            .Create Me.hwnd, "√‰Ê«⁄ «·„’—Ê›« ", 1, 15204351, -2147483630
            .MaxWidth = 4000
            .VisibleTime = 9000
            .DelayTime = 600
            .AddControl Cmd(1), " ⁄œÌ· ..." & Wrap & "· ⁄œÌ· »Ì«‰«  Â–« «·‰Ê⁄" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
        End With

        With TTP
            .Create Me.hwnd, "√‰Ê«⁄ «·„’—Ê›« ", 1, 15204351, -2147483630
            .MaxWidth = 4000
            .VisibleTime = 9000
            .DelayTime = 600
            .AddControl Cmd(2), "Õ›Ÿ ..." & Wrap & "·Õ›Ÿ »Ì«‰«  «·‰Ê⁄ «·ÃœÌœ" & Wrap & "·Õ›Ÿ «· ⁄œÌ·« " & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
        End With

        With TTP
            .Create Me.hwnd, "√‰Ê«⁄ «·„’—Ê›« ", 1, 15204351, -2147483630
            .MaxWidth = 4000
            .VisibleTime = 9000
            .DelayTime = 600
            .AddControl Cmd(3), " —«Ã⁄ ..." & Wrap & "·· —«Ã⁄ ⁄‰ ⁄„·Ì… «·«÷«›…" & Wrap & "··· —«Ã⁄ ⁄‰ ⁄„·Ì… «· ⁄œÌ·" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
        End With

        With TTP
            .Create Me.hwnd, "√‰Ê«⁄ «·„’—Ê›« ", 1, 15204351, -2147483630
            .MaxWidth = 4000
            .VisibleTime = 9000
            .DelayTime = 600
            .AddControl Cmd(4), "Õ–› ..." & Wrap & "·Õ–› »Ì«‰«  Â–« «·‰Ê⁄" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
        End With

        With TTP
            .Create Me.hwnd, "√‰Ê«⁄ «·„’—Ê›« ", 1, 15204351, -2147483630
            .MaxWidth = 4000
            .VisibleTime = 9000
            .DelayTime = 600
            .AddControl Cmd(6), "Œ—ÊÃ ..." & Wrap & "·«€·«ﬁ Â–Â «·‰«›–…" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
        End With

        With TTP
            .Create Me.hwnd, "√‰Ê«⁄ «·„’—Ê›« ", 1, 15204351, -2147483630
            .MaxWidth = 4000
            .VisibleTime = 9000
            .DelayTime = 600
            .AddControl XPBtnMove(1), "«·√Ê· ..." & Wrap & "··«‰ ﬁ«· «·Ï √Ê· ”Ã·" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
        End With

        With TTP
            .Create Me.hwnd, "√‰Ê«⁄ «·„’—Ê›« ", 1, 15204351, -2147483630
            .MaxWidth = 4000
            .VisibleTime = 9000
            .DelayTime = 600
            .AddControl XPBtnMove(0), "«·”«»ﬁ ..." & Wrap & "··«‰ ﬁ«· «·Ï «·”Ã· «·”«»ﬁ" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
        End With

        With TTP
            .Create Me.hwnd, "√‰Ê«⁄ «·„’—Ê›« ", 1, 15204351, -2147483630
            .MaxWidth = 4000
            .VisibleTime = 9000
            .DelayTime = 600
            .AddControl XPBtnMove(3), "«· «·Ì ..." & Wrap & "··«‰ ﬁ«· «·Ï «·”Ã· «· «·Ì" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
        End With

        With TTP
            .Create Me.hwnd, "√‰Ê«⁄ «·„’—Ê›« ", 1, 15204351, -2147483630
            .MaxWidth = 4000
            .VisibleTime = 9000
            .DelayTime = 600
            .AddControl XPBtnMove(2), "«·√ŒÌ— ..." & Wrap & "··«‰ ﬁ«· «·Ï ¬Œ— ”Ã·" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
        End With

    ElseIf SystemOptions.UserInterface = EnglishInterface Then

    End If

    Exit Sub
ErrTrap:
End Sub

Private Sub ChangeLang()
Cmd(19).Caption = "Car Expenses"
printPartsRep.Caption = "Print"
    Dim XPic As IPictureDisp
    Set XPic = Me.XPBtnMove(1).ButtonImage
    Set Me.XPBtnMove(1).ButtonImage = Me.XPBtnMove(2).ButtonImage
    Set Me.XPBtnMove(2).ButtonImage = XPic
C1Tab1.TabCaption(0) = "Data"
C1Tab1.TabCaption(1) = "Data"
lbl(54).Caption = "Switch To"
    Set XPic = Me.XPBtnMove(0).ButtonImage
    Set Me.XPBtnMove(0).ButtonImage = Me.XPBtnMove(3).ButtonImage
    Set Me.XPBtnMove(3).ButtonImage = XPic
    lbl(30).Caption = "Price"
    lbl(124).Caption = "Description"
    lbl(31).Caption = "Total"
    lbl(20).Caption = "Ser.Kir"
    lbl(22).Caption = "No.Kir"
     lbl(19).Caption = "Ser.Machine"
    lbl(21).Caption = "No.Machine"
    lbl(24).Caption = "Color"
    lbl(27).Caption = "Location"
 Rd(0).RightToLeft = False
 lbl(28).Caption = "No.liters"
 C1Elastic8.Caption = "Fuel consumption rate"
 btnAddImage.Caption = "Add Image"
 Rd(1).RightToLeft = False
 Rd(0).Caption = "From FixedAssest"
  Rd(1).Caption = "Manual"
  Cmd(17).Caption = "Print"
    Frame1.Caption = "Group Data"
lbl(6).Caption = "Equipment Long"
    Me.Caption = "Equipment Data"
   Me.ELe(0).Caption = Me.Caption
    Me.lbl(101).Caption = "Code"
    Me.lbl(102).Caption = "Name"
lbl(18).Caption = "Equipment"
    Me.lbl(103).Caption = "Type"
    Me.lbl(117).Caption = "Branch"
   ' Me.lbl(104).Caption = "Employee"
    Me.lbl(3).Caption = "Insur. Co."
    Me.lbl(105).Caption = "Board No."
    Me.lbl(1).Caption = "Purchase Date"
    Me.lbl(120).Caption = "Check Up Date"
    Me.lbl(106).Caption = "License No."
    lbl(25).Caption = "Type"
    lbl(23).Caption = "Structure No"
    Me.lbl(102).Caption = "Name"
    Me.lbl(128).Caption = "License Expire"
    Me.lbl(127).Caption = "Insurance Expire"

    Me.lbl(107).Caption = "Model"
    Me.lbl(2).Caption = "Last Km Count"
    Me.lbl(124).Caption = "Remarks"
    Cmd(10).Caption = "Maintenance Plan"
    Cmd(12).Caption = "Search"

    Cmd(8).Caption = "Depreciation Restart"
    Cmd(9).Caption = "Asset Disposal"
    Cmd(5).Caption = "Asset Image"
    Cmd(10).Caption = " Show bill"

    Me.lbl(125).Caption = "Current Record:"
    Me.lbl(126).Caption = "Records NO:"
    Me.Cmd(0).Caption = "New"
    Me.Cmd(1).Caption = "Edit"
    Me.Cmd(2).Caption = "Save"
    Me.Cmd(3).Caption = "Undo"
    Me.Cmd(4).Caption = "Delete"
    Me.Cmd(6).Caption = "Exit"
    lbl(17).Caption = "Operational No."
    lbl(16).Caption = "Max.Speed"
    Cmd(7).Caption = "Stop Dep"
    lbl(14).Caption = "End Date"
    lbl(15).Caption = "Max.No.Trip"
lbl(12).Caption = "Set Count"
lbl(13).Caption = "Rate"
lbl(7).Caption = "Load"
lbl(11).Caption = "Begin Contact No"
lbl(10).Caption = "End Contract Date"
lbl(8).Caption = "Capacity"
Label4.Caption = "Ton"

    With CBoDepreciation_Type_id
        .Clear
        .AddItem "fixed "
        .AddItem "Decreasing"
    End With

'############################# khaled ############################
'****************** tab 1 **********************
lbl(47).Caption = "Equipment status"
lbl(29).Caption = "Equipment User"
ChDrievType(0).Caption = "Employee"
ChDrievType(1).Caption = "Add Employee"
lbl(41).Caption = "Nationality "
lbl(42).Caption = "Job Title"
lbl(43).Caption = "Department"
lbl(37).Caption = "Tracking  No."
PushButton1.Caption = "Refresh status"
Cmd(18).Caption = "Extensive Report"
C1Tab1.TabCaption(2) = "Accidents data"
C1Tab1.TabCaption(3) = "Expenses"
C1Tab1.TabCaption(4) = "Parts"
'**************** tab 2 ************************
C1Elastic11.Caption = "Equipment Documents"
lbl(38).Caption = "Equipment owner"
Label1(1).Caption = "Country Name"
Label5.Caption = "original form"
Label1(4).Caption = "City Name"
C1Elastic10.Caption = "Equipment Attachments"
Insurance.Caption = "Insurance"
lbl(3).Caption = "Insurance Company"
Label9.Caption = "Insurance No."
lbl(127).Caption = "Insurance Expiration"
Authorization.Caption = "Authorization"
lbl(35).Caption = "Authorization type"
Label10.Caption = "Authorization No."
lbl(34).Caption = "Authorization Expiration date"
Licenses.Caption = "Equipment Licenses"
lbl(106).Caption = "Licenses No."
lbl(128).Caption = "Licenses Expiration"
Exam.Caption = "examination"
lbl(36).Caption = "Examination No."
lbl(120).Caption = "Examination Expiration date"
DriLicense.Caption = "Driver Licenses"
lbl(39).Caption = "Driver Licenses No."
lbl(40).Caption = "Driver Licenses Expiration"
KeyReserve.Caption = "Spare key"
Receipt.Caption = "Receipt"
Triangle.Caption = "Emergency triangle"
TrackingDevice.Caption = "Tracking Device"
Crane.Caption = "Crane"
FireExtingui.Caption = "Fire Extinguisher"

'IsUsed.Caption = "Is Used"
Sabt.Caption = "Sabt"
Chains.Caption = "Chains"
Kafla.Caption = "Kafla"
Hock.Caption = "Hock"
Khabor.Caption = "Khabor"
Sail.Caption = "Sail"
SideBarriers.Caption = "Side Barriers"
SideFrame.Caption = "Side Frame"


SubsBattery.Caption = "Emergency battary connector"
Battery.Caption = "Battery"
BagAmbulance.Caption = "First aid kit"
'CoverKey.Caption = ""
cleaner.Caption = "Windscreen wipers"
sideMirror.Caption = "Wing mirror"
driverMirror.Caption = "Rear view mirror"
InnerLights.Caption = "Inner Lights"
SpareTyre.Caption = "Spare Tyre"
SunScreens.Caption = "Sun Screens"
Recorder.Caption = "Multimedia device"
Anntena.Caption = "Anntena"
Stickers.Caption = "Emergency Stickers"
Pedals.Caption = "Pedals"
Guarantee.Caption = "Guarantee certificate"
keys.Caption = "Keys"
'************************tab 3**************************
lbl(44).Caption = "Accidents expenses"
With fg
    .TextMatrix(0, .ColIndex("Serial")) = "No."
    .TextMatrix(0, .ColIndex("ID")) = "Accident No."
    .TextMatrix(0, .ColIndex("AccDate")) = "Accident Date"
    .TextMatrix(0, .ColIndex("AccTime")) = "Accident Time"
    .TextMatrix(0, .ColIndex("branch_name")) = "Branch"
    .TextMatrix(0, .ColIndex("CompValue")) = "Company bearing ratio"
    .TextMatrix(0, .ColIndex("EmpValue")) = "Employee bearing ratio"
    .TextMatrix(0, .ColIndex("shw")) = "Show"
    lbl(45).Caption = "Total"
End With
'******************** tab 4 *****************************
lbl(48).Caption = "Expenses"
With VSFlexGrid1
    .TextMatrix(0, .ColIndex("serial")) = "No."
    .TextMatrix(0, .ColIndex("Name")) = "Expense"
    .TextMatrix(0, .ColIndex("Vlue")) = "Value"
    .TextMatrix(0, .ColIndex("Remarks")) = "Notes"
End With
lbl(49).Caption = "Expense"
lbl(51).Caption = "Value"
lbl(52).Caption = "Notes"
Label1(48).Caption = "Total"
cmdAdd.Caption = "Add"
btnModify.Caption = "Edit"
CmdDelete.Caption = "Delete"
cmdPrint.Caption = "Print Expenses"
'****************** tab 5 ******************************
lbl(53).Caption = "Part Name"
ISButton2.Caption = "Add"
ISButton1.Caption = "Edit"
ISButton3.Caption = "Delete"
With VSFlexGrid13
    .TextMatrix(0, .ColIndex("Serial")) = "No."
    .TextMatrix(0, .ColIndex("Code")) = "Code"
    .TextMatrix(0, .ColIndex("Name")) = "Name"
End With
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
MySQL = "SELECT     dbo.TblCarsData.id, dbo.TblCarsData.code, dbo.TblCarsData.Fullcode, dbo.TblCarsData.prifix, dbo.TblCarsData.TypeCar, dbo.TblCarsData.Branch_NO, "
MySQL = MySQL & "                      dbo.TblBranchesData.branch_name, dbo.TblBranchesData.branch_namee, dbo.TblCarsData.CarsTypeId, dbo.TBLCarTypes.name, dbo.TBLCarTypes.namee,"
MySQL = MySQL & "                      dbo.TblCarsData.LicenseNO, dbo.TblCarsData.BoardNO, dbo.TblCarsData.EqupName, dbo.FixedAssets.Name AS CarName, dbo.FixedAssets.namee AS CarNameE,"
MySQL = MySQL & "                      dbo.TblCarsData.Name AS HName, dbo.TblCarsData.Emp_id, dbo.TblEmployee.Emp_Name, dbo.TblEmployee.Emp_Name1, dbo.TblEmployee.Emp_Name2,"
MySQL = MySQL & "                      dbo.TblEmployee.Emp_Name3, dbo.TblEmployee.Emp_Name4, dbo.TblEmployee.Fullcode AS EmpFullcode, dbo.TblEmployee.Emp_Namee4,"
MySQL = MySQL & "                      dbo.TblEmployee.Emp_Namee3, dbo.TblEmployee.Emp_Namee2, dbo.TblEmployee.Emp_Namee1, dbo.TblEmployee.Emp_Namee, dbo.TblCarsData.Model,"
MySQL = MySQL & "                      dbo.TblCarsData.PurchaseDate, dbo.TblCarsData.LastKMCounter, dbo.TblCarsData.LicenseExpireDate, dbo.TblCarsData.LicenseExpireDateH,"
MySQL = MySQL & "                      dbo.TblCarsData.InsuranceExpireDateH, dbo.TblCarsData.TestExpireDateH, dbo.TblCarsData.fixedAssetid, dbo.TblCarsData.VehicleLong,"
MySQL = MySQL & "                      dbo.TblCarsData.InsuranceExpireDate, dbo.TblCarsData.TestExpireDate, dbo.TblCarsData.Notes, dbo.TblCarsData.SetCount, dbo.TblCarsData.Rate,"
MySQL = MySQL & "                      dbo.TblCarsData.Capacity, dbo.TblCarsData.MaxCap, dbo.TblCarsData.Machineno, dbo.TblCarsData.Machineno1, dbo.TblCarsData.Chesis, dbo.TblCarsData.VModel,"
MySQL = MySQL & "                      dbo.TblCarModels.Model AS ModelName, dbo.TblCarModels.ModelE AS ModelNameE, dbo.TblCarsData.EndContractDateH, dbo.TblCarsData.EndAllocationDate,"
MySQL = MySQL & "                      dbo.TblCarsData.LocationID, dbo.EmpGroupDep.GroupName, dbo.TblCarsData.EquQty, dbo.TblCarsData.Rep, dbo.TblCarsData.EndContractDate,"
MySQL = MySQL & "                      dbo.TblCarsData.OperatorN, dbo.TblCarsData.Gearno, dbo.TblCarsData.Gearno1, dbo.TblCarsData.VColor, dbo.TblColor.name AS Colorname,"
MySQL = MySQL & "                      dbo.TblColor.namee AS ColornameE, dbo.TblCarsData.InsuranceCompanyId, dbo.insurance_companies.name AS CompInsurName,"
MySQL = MySQL & "                      dbo.insurance_companies.Namee AS CompInsurNameE, dbo.TblCarsData.Total, dbo.TblCarsData.LetterCount, dbo.TblCarsData.LetterPrice,"
MySQL = MySQL & "                      dbo.TblCarsData.FormOrignal, dbo.TblCarsData.authorizeLicense, dbo.TblCarsData.authorizeExamination, dbo.TblCarsData.cleaner, dbo.TblCarsData.sideMirror,"
MySQL = MySQL & "                      dbo.TblCarsData.driverMirror, dbo.TblCarsData.InnerLights, dbo.TblCarsData.Pedals, dbo.TblCarsData.SunScreens, dbo.TblCarsData.Recorder,"
MySQL = MySQL & "                      dbo.TblCarsData.Anntena, dbo.TblCarsData.Battery, dbo.TblCarsData.SpareTyre, dbo.TblCarsData.Crane, dbo.TblCarsData.CoverKey, dbo.TblCarsData.Guarantee,"
MySQL = MySQL & "                      dbo.TblCarsData.Stickers , dbo.TblCarsData.LeaderName, dbo.TblCarsData.EmpType"
MySQL = MySQL & " FROM         dbo.TblCarsData LEFT OUTER JOIN"
MySQL = MySQL & "                      dbo.insurance_companies ON dbo.TblCarsData.InsuranceCompanyId = dbo.insurance_companies.id LEFT OUTER JOIN"
MySQL = MySQL & "                      dbo.TblColor ON dbo.TblCarsData.VColor = dbo.TblColor.Id LEFT OUTER JOIN"
MySQL = MySQL & "                      dbo.EmpGroupDep ON dbo.TblCarsData.LocationID = dbo.EmpGroupDep.GroupID LEFT OUTER JOIN"
MySQL = MySQL & "                      dbo.TblCarModels ON dbo.TblCarsData.VModel = dbo.TblCarModels.Id LEFT OUTER JOIN"
MySQL = MySQL & "                      dbo.TblEmployee ON dbo.TblCarsData.Emp_id = dbo.TblEmployee.Emp_ID LEFT OUTER JOIN"
MySQL = MySQL & "                      dbo.FixedAssets ON dbo.TblCarsData.fixedAssetid = dbo.FixedAssets.id LEFT OUTER JOIN"
MySQL = MySQL & "                      dbo.TBLCarTypes ON dbo.TblCarsData.CarsTypeId = dbo.TBLCarTypes.id LEFT OUTER JOIN"
MySQL = MySQL & "                      dbo.TblBranchesData ON dbo.TblCarsData.Branch_NO = dbo.TblBranchesData.branch_id"
MySQL = MySQL & "  where    TblCarsData.ID  =" & val(XPTxtID.Text)


 If SystemOptions.UserInterface = ArabicInterface Then
          StrFileName = App.path & "\REPORTS\REPORTS NEW\" & "rpt_Cars.rpt"
     Else
        StrFileName = App.path & "\REPORTS\REPORTS NEW\" & "rpt_CarsE.rpt"
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
        StrReportTitle = "" '& StrAccountName
    Else
        xReport.ParameterFields(1).AddCurrentValue cCompanyInfo.ArabCompanyName ' RPTCompany_Name_Eng
      '  xReport.ParameterFields(2).AddCurrentValue RPTComment_Eng
        xReport.ParameterFields(4).AddCurrentValue get_branch_name(val(my_branch))
        StrReportTitle = ""
    End If
   '
    xReport.ParameterFields(3).AddCurrentValue user_name
    xReport.ParameterFields(4).AddCurrentValue SumTotalExpen(val(DcFixedAssets.BoundText)) + val(TotalValue.Caption) + val(lbl(46).Caption)
    xReport.reporttitle = StrReportTitle
    xReport.EnableParameterPrompting = False
    xReport.ApplicationName = App.title
    xReport.ReportAuthor = App.title
              Dim xLogo As CRAXDRT.OLEObject
         If Dir(App.path & "\" & SystemOptions.ImagesPath & "\" & val(XPTxtID.Text) & ".JPG") <> "" Then
            Set xLogo = xReport.Areas(1).Sections(1).AddPictureObject(App.path & "\" & SystemOptions.ImagesPath & "\" & val(XPTxtID.Text) & ".JPG", 250, 2400)
            xLogo.Width = 7500
            xLogo.Height = 3500
            xLogo.backcolor = vbWhite
            xLogo.BorderColor = 255
            xLogo.CloseAtPageBreak = True
          End If
    ''///////
    Set CViewer = New ClsReportViewer
    CViewer.FireReport xReport, WindowTarget, "", , , , StrFileName

    RsData.Close
    Set RsData = Nothing
    Screen.MousePointer = vbDefault
End Function
Function print_reportExp(Optional NoteSerial As String)
    
     
    Dim MySQL As String
    Dim RsData As New ADODB.Recordset
    Dim xApp As New CRAXDRT.Application
    Dim xReport As CRAXDRT.Report
    Dim CViewer As ClsReportViewer
    Dim StrReportTitle As String
    Dim StrFileName As String
    Dim Msg As String
MySQL = " SELECT        dbo.TblCarExpenses.Remarks, dbo.TblCarExpenses.Vlue, dbo.TblCarExpenses.CarID, dbo.TblCarExpenses.ExpID, dbo.TblDataTypeExchange.name, dbo.TblDataTypeExchange.namee, dbo.TblCarsData.code, "
MySQL = MySQL & "                          dbo.TblCarsData.Fullcode, dbo.TblCarsData.Name AS CarName2, dbo.TblCarsData.BoardNO, dbo.TblCarsData.fixedAssetid, dbo.FixedAssets.Name AS CarName, dbo.FixedAssets.namee AS CarNameE,"
MySQL = MySQL & "                         dbo.FixedAssets.Fullcode AS CarFullCode"
MySQL = MySQL & " FROM            dbo.FixedAssets RIGHT OUTER JOIN"
MySQL = MySQL & "                         dbo.TblCarsData ON dbo.FixedAssets.id = dbo.TblCarsData.fixedAssetid RIGHT OUTER JOIN"
MySQL = MySQL & "                         dbo.TblCarExpenses ON dbo.TblCarsData.id = dbo.TblCarExpenses.CarID LEFT OUTER JOIN"
MySQL = MySQL & "                         dbo.TblDataTypeExchange ON dbo.TblCarExpenses.ExpID = dbo.TblDataTypeExchange.Id"
MySQL = MySQL & "  where    TblCarExpenses.CarID  =" & val(XPTxtID.Text)


 If SystemOptions.UserInterface = ArabicInterface Then
          StrFileName = App.path & "\REPORTS\REPORTS NEW\" & "rpt_CarsExpens.rpt"
     Else
        StrFileName = App.path & "\REPORTS\REPORTS NEW\" & "rpt_CarsExpens.rpt"
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
        StrReportTitle = "" '& StrAccountName
    Else
        xReport.ParameterFields(1).AddCurrentValue cCompanyInfo.ArabCompanyName ' RPTCompany_Name_Eng
      '  xReport.ParameterFields(2).AddCurrentValue RPTComment_Eng
        xReport.ParameterFields(4).AddCurrentValue get_branch_name(val(my_branch))
        StrReportTitle = ""
    End If
   '
    xReport.ParameterFields(3).AddCurrentValue user_name
    xReport.reporttitle = StrReportTitle
    xReport.EnableParameterPrompting = False
    xReport.ApplicationName = App.title
    xReport.ReportAuthor = App.title

    ''///////
    Set CViewer = New ClsReportViewer
    CViewer.FireReport xReport, WindowTarget, "", , , , StrFileName

    RsData.Close
    Set RsData = Nothing
    Screen.MousePointer = vbDefault
End Function
Function print_reportExpense(Optional NoteSerial As String)
    
     
    Dim StrSQL As String
    Dim RsData As New ADODB.Recordset
    Dim xApp As New CRAXDRT.Application
    Dim xReport As CRAXDRT.Report
    Dim CViewer As ClsReportViewer
    Dim StrReportTitle As String
    Dim StrFileName As String
    Dim Msg As String
  StrSQL = " SELECT     dbo.TblOrderMaint.ID, dbo.TblOrderMaint.RecordDate, dbo.TblOrderMaint.BranchID, TblBranchesData_2.branch_name, TblBranchesData_2.branch_namee, "
  StrSQL = StrSQL & "                     dbo.TblOrderMaint.UserID, dbo.FixedAssets.Name, dbo.FixedAssets.namee, TblEmployee_1.Emp_Name, TblEmployee_1.Emp_Name1, TblEmployee_1.Emp_Name2,"
  StrSQL = StrSQL & "                     TblEmployee_1.Emp_Name3, TblEmployee_1.Emp_Name4, TblEmployee_1.Fullcode, TblEmployee_1.Emp_Namee, TblEmployee_1.Emp_Namee1,"
  StrSQL = StrSQL & "                    TblEmployee_1.Emp_Namee2, TblEmployee_1.Emp_Namee3, dbo.TblOrderMaint.TypeMaint, dbo.TblOrderMaint.Jiha, dbo.TblOrderMaint.Remarks,"
  StrSQL = StrSQL & "                    dbo.TblOrderMaint.Cost, dbo.TblOrderMaint.Des, dbo.TblOrderMaint.startmaintenanceTime, dbo.TblOrderMaint.endmaintenanceTime,"
  StrSQL = StrSQL & "                    dbo.TblOrderMaint.RecmaintenanceTime, dbo.TblOrderMaint.RecmaintenanceDate, dbo.TblOrderMaint.reciverRemarks, dbo.TblOrderMaint.TechNote,"
  StrSQL = StrSQL & "                    dbo.TblOrderMaint.reciverid, TblEmployee_1.Emp_Name AS ReciEmp_Name, TblEmployee_1.Emp_Name1 AS ReciEmp_Name1,"
  StrSQL = StrSQL & "                    TblEmployee_1.Emp_Name2 AS ReciEmp_Name2, TblEmployee_1.Emp_Name3 AS ReciEmp_Name3, TblEmployee_1.Fullcode AS ReciFullcode,"
  StrSQL = StrSQL & "                    TblEmployee_1.Emp_Namee4 AS ReciEmp_Namee4, TblEmployee_1.Emp_Namee3 AS ReciEmp_Namee3, TblEmployee_1.Emp_Namee2 AS ReciEmp_Namee2,"
  StrSQL = StrSQL & "                    TblEmployee_1.Emp_Namee1 AS ReciEmp_Namee1, TblEmployee_1.Emp_Namee AS RecieEmp_Namee, dbo.TblOrderMaint.endmaintenanceDate,"
  StrSQL = StrSQL & "                    dbo.TblOrderMaint.ended, dbo.TblOrderMaint.ReqMainID, TblEmployee_1.Emp_Namee4, dbo.tblordermaintenancetypes.Qty,"
  StrSQL = StrSQL & "                    dbo.tblordermaintenancetypes.Remarks AS RemarksDet, dbo.tblordermaintenancetypes.ID AS IDDet, dbo.tblordermaintenancetypes.ORderID,"
  StrSQL = StrSQL & "                    dbo.tblordermaintenancetypes.maintenanceid, dbo.TblMaintenanceType.name AS nameMType, dbo.TblMaintenanceType.namee AS nameMTypeE,"
  StrSQL = StrSQL & "                    dbo.TblMaintenanceType.id AS MainID, dbo.Transaction_Details.Item_ID, dbo.TblItems.ItemCode, dbo.TblItems.ItemName, dbo.TblItems.ItemNamee,"
  StrSQL = StrSQL & "                    dbo.Transaction_Details.showPrice, dbo.Transaction_Details.ShowQty, dbo.Transaction_Details.ID AS TrnsID, dbo.TblOrderMaint.LeaderName,"
  StrSQL = StrSQL & "                    dbo.TblOrderMaint.LeaderType, dbo.TblOrderMaint.DrievType, dbo.TblOrderMaint.DrievName, dbo.TblOrderMaint.EquepmentName, dbo.TblOrderMaint.Total,"
  StrSQL = StrSQL & "                    dbo.TblOrderMaint.DcbBranchFrom, TblBranchesData_1.branch_name AS Frombranch_name, TblBranchesData_1.branch_namee AS Frombranch_nameE,"
  StrSQL = StrSQL & "                    dbo.TblOrderMaint.LeaderID, TblEmployee_3.Emp_Name AS LeaderEmp_Name, TblEmployee_3.Fullcode AS LeaderFullcode,"
  StrSQL = StrSQL & "                    TblEmployee_3.Emp_Namee AS LeaderEmp_NameE, dbo.TblOrderMaint.SuperVisor, TblEmployee_2.Emp_Name AS SuperEmp_Name,"
  StrSQL = StrSQL & "                    TblEmployee_2.Fullcode AS SuperFullcode, TblEmployee_2.Emp_Namee AS SuperEmp_NameE, dbo.TblOrderMaint.DrievID,"
  StrSQL = StrSQL & "                    TblEmployee_4.Emp_Name AS DevEmp_Name, TblEmployee_4.Fullcode AS DevFullcode, TblEmployee_4.Emp_Namee AS DevEmp_NameE,"
  StrSQL = StrSQL & "                    dbo.tblordermaintenancetypes.LocaMaint, dbo.tblordermaintenancetypes.Company, dbo.tblordermaintenancetypes.Price,"
  StrSQL = StrSQL & "                    dbo.tblordermaintenancetypes.Total AS TotalDet, dbo.tblordermaintenancetypes.BillNo, dbo.tblordermaintenancetypes.CusMobile,"
  StrSQL = StrSQL & "                    dbo.tblordermaintenancetypes.PartName, dbo.tblordermaintenancetypes.CusID, dbo.TblCustemers.CusName, dbo.TblCustemers.CusNamee,"
  StrSQL = StrSQL & "                    dbo.TblCustemers.Fullcode AS CusFullcode, dbo.tblordermaintenancetypes.EmpID, TblEmployee_5.Emp_Name AS FiterEmp_Name,"
  StrSQL = StrSQL & "                    TblEmployee_5.Fullcode AS FiterFullcode, TblEmployee_5.Emp_Namee AS FiterEmp_NameE, dbo.tblordermaintenancetypes.SuperID,"
  StrSQL = StrSQL & "                    TblEmployee_6.Emp_Name AS SuperEmp_NameDet, TblEmployee_6.Fullcode AS SuperFullcodeDet, TblEmployee_6.Emp_Namee AS SuperEmp_NameDetE,"
  StrSQL = StrSQL & "                    dbo.tblordermaintenancetypes.DeptID, dbo.TblEmpDepartments.DepartmentName, dbo.TblEmpDepartments.DepartmentNamee,"
  StrSQL = StrSQL & "                    dbo.tblordermaintenancetypes.Transaction_ID AS Transaction_IDH, dbo.tblordermaintenancetypes.Transaction_IDDet, dbo.Transactions.Transaction_Date,"
  StrSQL = StrSQL & "                    dbo.Transactions.Transaction_Type, dbo.Transactions.NoteSerial1, dbo.Transactions.Transaction_Serial, dbo.Transactions.Transaction_HijriDate,"
  StrSQL = StrSQL & "                    dbo.Transactions.TransactionComment, dbo.tblordermaintenancetypes.TypeTrans, dbo.Transactions.OpOrderID, dbo.Transactions.OldOpOrderID,"
  StrSQL = StrSQL & "                    dbo.Transaction_Details.OperPrice, dbo.TblOrderMaint.TotalSand, dbo.TblOrderMaint.TotalSpare, dbo.TblOrderMaint.TotalMaint, dbo.TblOrderMaint.OperatorN,"
  StrSQL = StrSQL & "                    dbo.TblOrderMaint.BoardNO, dbo.TblOrderMaint.StutsMaint, dbo.TblOrderMaint.EnterDate, dbo.TblOrderMaint.EnterTime, dbo.TblOrderMaint.startmaintenanceDate,"
  StrSQL = StrSQL & "                    dbo.FixedAssets.code, dbo.TblOrderMaint.EquepID, dbo.TblCarsData.Fullcode AS CarFullcode, dbo.TblCarsData.Model, dbo.TblCarsData.CarsTypeId,"
  StrSQL = StrSQL & "                    dbo.TBLCarTypes.name AS CrTypname, dbo.TBLCarTypes.namee AS CrTypnameE, dbo.TblCarsData.VModel, dbo.TblCarModels.Model AS ModelName,"
  StrSQL = StrSQL & "                    dbo.TblCarModels.ModelE AS ModelNameE"
  StrSQL = StrSQL & "    FROM         dbo.TblCarModels RIGHT OUTER JOIN"
  StrSQL = StrSQL & "                    dbo.TblCarsData ON dbo.TblCarModels.Id = dbo.TblCarsData.VModel LEFT OUTER JOIN"
  StrSQL = StrSQL & "                    dbo.TBLCarTypes ON dbo.TblCarsData.CarsTypeId = dbo.TBLCarTypes.id RIGHT OUTER JOIN"
  StrSQL = StrSQL & "                    dbo.FixedAssets ON dbo.TblCarsData.fixedAssetid = dbo.FixedAssets.id RIGHT OUTER JOIN"
  StrSQL = StrSQL & "                    dbo.TblOrderMaint LEFT OUTER JOIN"
  StrSQL = StrSQL & "                    dbo.TblEmployee TblEmployee_4 ON dbo.TblOrderMaint.DrievID = TblEmployee_4.Emp_ID LEFT OUTER JOIN"
  StrSQL = StrSQL & "                    dbo.TblMaintenanceType RIGHT OUTER JOIN"
  StrSQL = StrSQL & "                    dbo.TblEmployee TblEmployee_5 RIGHT OUTER JOIN"
  StrSQL = StrSQL & "                    dbo.TblEmpDepartments RIGHT OUTER JOIN"
  StrSQL = StrSQL & "                    dbo.TblItems RIGHT OUTER JOIN"
  StrSQL = StrSQL & "                    dbo.tblordermaintenancetypes LEFT OUTER JOIN"
  StrSQL = StrSQL & "                    dbo.Transaction_Details LEFT OUTER JOIN"
  StrSQL = StrSQL & "                    dbo.Transactions ON dbo.Transaction_Details.Transaction_ID = dbo.Transactions.Transaction_ID ON"
  StrSQL = StrSQL & "                    dbo.tblordermaintenancetypes.Transaction_IDDet = dbo.Transaction_Details.ID LEFT OUTER JOIN"
  StrSQL = StrSQL & "                    dbo.TblCustemers ON dbo.tblordermaintenancetypes.CusID = dbo.TblCustemers.CusID ON dbo.TblItems.ItemID = dbo.Transaction_Details.Item_ID ON"
  StrSQL = StrSQL & "                    dbo.TblEmpDepartments.DeparmentID = dbo.tblordermaintenancetypes.DeptID LEFT OUTER JOIN"
  StrSQL = StrSQL & "                    dbo.TblEmployee TblEmployee_6 ON dbo.tblordermaintenancetypes.SuperID = TblEmployee_6.Emp_ID ON"
  StrSQL = StrSQL & "                    TblEmployee_5.Emp_ID = dbo.tblordermaintenancetypes.EmpID ON dbo.TblMaintenanceType.id = dbo.tblordermaintenancetypes.maintenanceid ON"
  StrSQL = StrSQL & "                    dbo.TblOrderMaint.ID = dbo.tblordermaintenancetypes.ORderID LEFT OUTER JOIN"
  StrSQL = StrSQL & "                    dbo.TblEmployee TblEmployee_2 ON dbo.TblOrderMaint.SuperVisor = TblEmployee_2.Emp_ID LEFT OUTER JOIN"
  StrSQL = StrSQL & "                    dbo.TblEmployee TblEmployee_1 ON dbo.TblOrderMaint.reciverid = TblEmployee_1.Emp_ID LEFT OUTER JOIN"
  StrSQL = StrSQL & "                    dbo.TblEmployee TblEmployee_3 ON dbo.TblOrderMaint.LeaderID = TblEmployee_3.Emp_ID LEFT OUTER JOIN"
  StrSQL = StrSQL & "                    dbo.TblBranchesData TblBranchesData_1 ON dbo.TblOrderMaint.DcbBranchFrom = TblBranchesData_1.branch_id ON"
  StrSQL = StrSQL & "                    dbo.FixedAssets.id = dbo.TblOrderMaint.EquepID LEFT OUTER JOIN"
  StrSQL = StrSQL & "                    dbo.TblBranchesData TblBranchesData_2 ON dbo.TblOrderMaint.BranchID = TblBranchesData_2.branch_id"
  StrSQL = StrSQL & " Where (dbo.TblOrderMaint.EquepID = " & val(DcFixedAssets.BoundText) & ")"

    If SystemOptions.UserInterface = ArabicInterface Then
        Msg = "Â–Â «·ÿ»«⁄…  Õ·Ì·Ì " & CHR(13)
        Msg = Msg + " Â·  —Ìœ ÿ»«⁄… «Ã„«·Ì"
    Else
        Msg = "This Analytical print do you want to print the total "
    End If
        If MsgBox(Msg, vbYesNo + vbQuestion + vbMsgBoxRight + vbMsgBoxRtlReading, App.title) = vbNo Then
      
             If SystemOptions.UserInterface = ArabicInterface Then
        
            StrFileName = App.path & "\REPORTS\REPORTS NEW\" & "RepMaintinancCarExpenses2.rpt"
            Else
            StrFileName = App.path & "\REPORTS\REPORTS NEW\" & "RepMaintinancCarExpenses2.rpt"
       End If
      
Else
         If SystemOptions.UserInterface = ArabicInterface Then
        
            StrFileName = App.path & "\REPORTS\REPORTS NEW\" & "RepMaintinancCarExpensesTotal2.rpt"
            Else
            StrFileName = App.path & "\REPORTS\REPORTS NEW\" & "RepMaintinancCarExpensesTotal2.rpt"
       End If
     
End If


    If Dir(StrFileName) = "" Then
        'GetMsgs 139, vbExclamation
        Screen.MousePointer = vbDefault
        Exit Function
    End If

    Set RsData = New ADODB.Recordset
    RsData.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

    If RsData.BOF Or RsData.EOF Then
        'GetMsgs 138, vbExclamation
print_reportAccedent
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
      '  xReport.ParameterFields(2).AddCurrentValue RPTComment_Eng
        xReport.ParameterFields(4).AddCurrentValue get_branch_name(val(my_branch))
        StrReportTitle = ""
    End If
   '
    xReport.ParameterFields(3).AddCurrentValue user_name
   xReport.ParameterFields(12).AddCurrentValue val(lbl(46).Caption)
   xReport.ParameterFields(13).AddCurrentValue val(TotalValue.Caption)
    xReport.reporttitle = StrReportTitle
    xReport.EnableParameterPrompting = False
    xReport.ApplicationName = App.title
    xReport.ReportAuthor = App.title
    ''///////
    Set CViewer = New ClsReportViewer
    CViewer.FireReport xReport, WindowTarget, "", , , , StrFileName

    RsData.Close
    Set RsData = Nothing
    Screen.MousePointer = vbDefault
End Function
Function print_reportAccedent(Optional NoteSerial As String)
     
    Dim StrSQL As String
    Dim RsData As New ADODB.Recordset
    Dim xApp As New CRAXDRT.Application
    Dim xReport As CRAXDRT.Report
    Dim CViewer As ClsReportViewer
    Dim StrReportTitle As String
    Dim StrFileName As String
    Dim Msg As String
  StrSQL = " SELECT     dbo.TblAccidentReport.PlateNo, SUM(dbo.TblAccidentReport.CompValue) AS SmVal, dbo.TblCarsData.code, dbo.TblCarsData.Fullcode, dbo.TblCarsData.Name, "
  StrSQL = StrSQL & "                     dbo.TblCarsData.Model, dbo.TblCarsData.fixedAssetid, dbo.FixedAssets.Name AS NameFix"
  StrSQL = StrSQL & "   FROM         dbo.FixedAssets RIGHT OUTER JOIN"
  StrSQL = StrSQL & "                    dbo.TblCarsData ON dbo.FixedAssets.id = dbo.TblCarsData.fixedAssetid RIGHT OUTER JOIN"
  StrSQL = StrSQL & "                    dbo.TblAccidentReport ON dbo.TblCarsData.BoardNO = dbo.TblAccidentReport.PlateNo"
  StrSQL = StrSQL & "    GROUP BY dbo.TblAccidentReport.PlateNo, dbo.TblCarsData.code, dbo.TblCarsData.Fullcode, dbo.TblCarsData.Name, dbo.TblCarsData.Model, dbo.TblCarsData.fixedAssetid,"
  StrSQL = StrSQL & "                     dbo.FixedAssets.name"
  StrSQL = StrSQL & " HAVING      (dbo.TblAccidentReport.PlateNo = N'" & TxtBoardNO.Text & "')"
  
             If SystemOptions.UserInterface = ArabicInterface Then
        
            StrFileName = App.path & "\REPORTS\REPORTS NEW\" & "RepMaintinancCarAccedrntTotal2.rpt"
            Else
            StrFileName = App.path & "\REPORTS\REPORTS NEW\" & "RepMaintinancCarAccedrntTotal2.rpt"
       End If


    If Dir(StrFileName) = "" Then
        'GetMsgs 139, vbExclamation
        Screen.MousePointer = vbDefault
        Exit Function
    End If

    Set RsData = New ADODB.Recordset
    RsData.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

    If RsData.BOF Or RsData.EOF Then
        'GetMsgs 138, vbExclamation
        If SystemOptions.UserInterface = ArabicInterface Then
        Msg = "·« ÊÃœ »Ì«‰«  ··⁄—÷"
        Else
        Msg = "No Data"
        End If
        print_reportOtherExppens
        Exit Function
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
      '  xReport.ParameterFields(2).AddCurrentValue RPTComment_Eng
        xReport.ParameterFields(4).AddCurrentValue get_branch_name(val(my_branch))
        StrReportTitle = ""
    End If
   '
    xReport.ParameterFields(3).AddCurrentValue user_name
   xReport.ParameterFields(12).AddCurrentValue val(lbl(46).Caption)
   xReport.ParameterFields(13).AddCurrentValue val(TotalValue.Caption)
    xReport.reporttitle = StrReportTitle
    xReport.EnableParameterPrompting = False
    xReport.ApplicationName = App.title
    xReport.ReportAuthor = App.title

    ''///////
    Set CViewer = New ClsReportViewer
    CViewer.FireReport xReport, WindowTarget, "", , , , StrFileName

    RsData.Close
    Set RsData = Nothing
    Screen.MousePointer = vbDefault
End Function
Function print_reportOtherExppens(Optional NoteSerial As String)
     
     
    Dim MySQL As String
    Dim RsData As New ADODB.Recordset
    Dim xApp As New CRAXDRT.Application
    Dim xReport As CRAXDRT.Report
    Dim CViewer As ClsReportViewer
    Dim StrReportTitle As String
    Dim StrFileName As String
    Dim Msg As String
MySQL = "SELECT     dbo.TblCarsData.id, dbo.TblCarsData.code, dbo.TblCarsData.Fullcode, dbo.TblCarsData.prifix, dbo.TblCarsData.TypeCar, dbo.TblCarsData.Branch_NO, "
MySQL = MySQL & "                      dbo.TblBranchesData.branch_name, dbo.TblBranchesData.branch_namee, dbo.TblCarsData.CarsTypeId, dbo.TBLCarTypes.name, dbo.TBLCarTypes.namee,"
MySQL = MySQL & "                      dbo.TblCarsData.LicenseNO, dbo.TblCarsData.BoardNO, dbo.TblCarsData.EqupName, dbo.FixedAssets.Name AS CarName, dbo.FixedAssets.namee AS CarNameE,"
MySQL = MySQL & "                      dbo.TblCarsData.Name AS HName, dbo.TblCarsData.Emp_id, dbo.TblEmployee.Emp_Name, dbo.TblEmployee.Emp_Name1, dbo.TblEmployee.Emp_Name2,"
MySQL = MySQL & "                      dbo.TblEmployee.Emp_Name3, dbo.TblEmployee.Emp_Name4, dbo.TblEmployee.Fullcode AS EmpFullcode, dbo.TblEmployee.Emp_Namee4,"
MySQL = MySQL & "                      dbo.TblEmployee.Emp_Namee3, dbo.TblEmployee.Emp_Namee2, dbo.TblEmployee.Emp_Namee1, dbo.TblEmployee.Emp_Namee, dbo.TblCarsData.Model,"
MySQL = MySQL & "                      dbo.TblCarsData.PurchaseDate, dbo.TblCarsData.LastKMCounter, dbo.TblCarsData.LicenseExpireDate, dbo.TblCarsData.LicenseExpireDateH,"
MySQL = MySQL & "                      dbo.TblCarsData.InsuranceExpireDateH, dbo.TblCarsData.TestExpireDateH, dbo.TblCarsData.fixedAssetid, dbo.TblCarsData.VehicleLong,"
MySQL = MySQL & "                      dbo.TblCarsData.InsuranceExpireDate, dbo.TblCarsData.TestExpireDate, dbo.TblCarsData.Notes, dbo.TblCarsData.SetCount, dbo.TblCarsData.Rate,"
MySQL = MySQL & "                      dbo.TblCarsData.Capacity, dbo.TblCarsData.MaxCap, dbo.TblCarsData.Machineno, dbo.TblCarsData.Machineno1, dbo.TblCarsData.Chesis, dbo.TblCarsData.VModel,"
MySQL = MySQL & "                      dbo.TblCarModels.Model AS ModelName, dbo.TblCarModels.ModelE AS ModelNameE, dbo.TblCarsData.EndContractDateH, dbo.TblCarsData.EndAllocationDate,"
MySQL = MySQL & "                      dbo.TblCarsData.LocationID, dbo.EmpGroupDep.GroupName, dbo.TblCarsData.EquQty, dbo.TblCarsData.Rep, dbo.TblCarsData.EndContractDate,"
MySQL = MySQL & "                      dbo.TblCarsData.OperatorN, dbo.TblCarsData.Gearno, dbo.TblCarsData.Gearno1, dbo.TblCarsData.VColor, dbo.TblColor.name AS Colorname,"
MySQL = MySQL & "                      dbo.TblColor.namee AS ColornameE, dbo.TblCarsData.InsuranceCompanyId, dbo.insurance_companies.name AS CompInsurName,"
MySQL = MySQL & "                      dbo.insurance_companies.Namee AS CompInsurNameE, dbo.TblCarsData.Total, dbo.TblCarsData.LetterCount, dbo.TblCarsData.LetterPrice,"
MySQL = MySQL & "                      dbo.TblCarsData.FormOrignal, dbo.TblCarsData.authorizeLicense, dbo.TblCarsData.authorizeExamination, dbo.TblCarsData.cleaner, dbo.TblCarsData.sideMirror,"
MySQL = MySQL & "                      dbo.TblCarsData.driverMirror, dbo.TblCarsData.InnerLights, dbo.TblCarsData.Pedals, dbo.TblCarsData.SunScreens, dbo.TblCarsData.Recorder,"
MySQL = MySQL & "                      dbo.TblCarsData.Anntena, dbo.TblCarsData.Battery, dbo.TblCarsData.SpareTyre, dbo.TblCarsData.Crane, dbo.TblCarsData.CoverKey, dbo.TblCarsData.Guarantee,"
MySQL = MySQL & "                      dbo.TblCarsData.Stickers , dbo.TblCarsData.LeaderName, dbo.TblCarsData.EmpType"
MySQL = MySQL & " FROM         dbo.TblCarsData LEFT OUTER JOIN"
MySQL = MySQL & "                      dbo.insurance_companies ON dbo.TblCarsData.InsuranceCompanyId = dbo.insurance_companies.id LEFT OUTER JOIN"
MySQL = MySQL & "                      dbo.TblColor ON dbo.TblCarsData.VColor = dbo.TblColor.Id LEFT OUTER JOIN"
MySQL = MySQL & "                      dbo.EmpGroupDep ON dbo.TblCarsData.LocationID = dbo.EmpGroupDep.GroupID LEFT OUTER JOIN"
MySQL = MySQL & "                      dbo.TblCarModels ON dbo.TblCarsData.VModel = dbo.TblCarModels.Id LEFT OUTER JOIN"
MySQL = MySQL & "                      dbo.TblEmployee ON dbo.TblCarsData.Emp_id = dbo.TblEmployee.Emp_ID LEFT OUTER JOIN"
MySQL = MySQL & "                      dbo.FixedAssets ON dbo.TblCarsData.fixedAssetid = dbo.FixedAssets.id LEFT OUTER JOIN"
MySQL = MySQL & "                      dbo.TBLCarTypes ON dbo.TblCarsData.CarsTypeId = dbo.TBLCarTypes.id LEFT OUTER JOIN"
MySQL = MySQL & "                      dbo.TblBranchesData ON dbo.TblCarsData.Branch_NO = dbo.TblBranchesData.branch_id"
MySQL = MySQL & "  where    TblCarsData.ID  =" & val(XPTxtID.Text)


 If SystemOptions.UserInterface = ArabicInterface Then
          StrFileName = App.path & "\REPORTS\REPORTS NEW\" & "rpt_CarsOtherExpens.rpt"
     Else
        StrFileName = App.path & "\REPORTS\REPORTS NEW\" & "rpt_CarsOtherExpens.rpt"
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
        StrReportTitle = "" '& StrAccountName
    Else
        xReport.ParameterFields(1).AddCurrentValue cCompanyInfo.ArabCompanyName ' RPTCompany_Name_Eng
      '  xReport.ParameterFields(2).AddCurrentValue RPTComment_Eng
        xReport.ParameterFields(4).AddCurrentValue get_branch_name(val(my_branch))
        StrReportTitle = ""
    End If
   '
    xReport.ParameterFields(3).AddCurrentValue user_name
    xReport.ParameterFields(4).AddCurrentValue val(TotalValue.Caption)
    xReport.reporttitle = StrReportTitle
    xReport.EnableParameterPrompting = False
    xReport.ApplicationName = App.title
    xReport.ReportAuthor = App.title

    ''///////
    Set CViewer = New ClsReportViewer
    CViewer.FireReport xReport, WindowTarget, "", , , , StrFileName

    RsData.Close
    Set RsData = Nothing
    Screen.MousePointer = vbDefault
End Function

Function print_report2(Optional NoteSerial As String)
    
     
    Dim MySQL As String
    Dim RsData As New ADODB.Recordset
    Dim xApp As New CRAXDRT.Application
    Dim xReport As CRAXDRT.Report
    Dim CViewer As ClsReportViewer
    Dim StrReportTitle As String
    Dim StrFileName As String
    Dim Msg As String
MySQL = " SELECT     dbo.TblCarsData.id, dbo.TblCarsData.Branch_NO, dbo.TblBranchesData.branch_name, dbo.TblBranchesData.branch_namee, dbo.TblCarsData.code, "
MySQL = MySQL & "                      dbo.TblCarsData.Fullcode, dbo.TblCarsData.prifix, dbo.TblCarsData.CarsTypeId, dbo.TBLCarTypes.name AS CasrName, dbo.TBLCarTypes.namee AS CasrNameE,"
MySQL = MySQL & "                      dbo.TblCarsData.Name, dbo.TblCarsData.BoardNO, dbo.TblCarsData.Model, dbo.TblCarsData.PurchaseDate, dbo.TblCarsData.LastKMCounter,"
MySQL = MySQL & "                      dbo.TblCarsData.fixedAssetid, dbo.FixedAssets.Name AS FixedName, dbo.FixedAssets.namee AS FixedNameE, dbo.TblCarsData.InsuranceCompanyId,"
MySQL = MySQL & "                      dbo.insurance_companies.name AS Insuname, dbo.insurance_companies.Namee AS InsunameE, dbo.TblCarsData.LicenseExpireDate, dbo.TblCarsData.Emp_id,"
MySQL = MySQL & "                      dbo.TblEmployee.Emp_Name, dbo.TblEmployee.Emp_Name1, dbo.TblEmployee.Emp_Name2, dbo.TblEmployee.Emp_Name3, dbo.TblEmployee.Emp_Name4,"
MySQL = MySQL & "                      dbo.TblEmployee.Fullcode AS EmpFullcode, dbo.TblEmployee.Emp_Namee4, dbo.TblEmployee.Emp_Namee3, dbo.TblEmployee.Emp_Namee2,"
MySQL = MySQL & "                      dbo.TblEmployee.Emp_Namee1, dbo.TblEmployee.Emp_Namee, dbo.TblEmployee.Nationality, dbo.TblEmployee.NumEkama, dbo.TblEmployee.NationalityE,"
MySQL = MySQL & "                      dbo.TblEmployee.JobTypeID, dbo.TblEmpJobsTypes.JobTypeName, dbo.TblEmpJobsTypes.JobTypeNamee, dbo.TblEmployee.DepartmentID,"
MySQL = MySQL & "                      dbo.TblEmpDepartments.DepartmentName, dbo.TblEmpDepartments.DepartmentNamee, dbo.TblCarsData.InsuranceExpireDate, dbo.TblCarsData.TestExpireDate,"
MySQL = MySQL & "                      dbo.TblCarsData.Notes, dbo.TblCarsData.LicenseExpireDateH, dbo.TblCarsData.InsuranceExpireDateH, dbo.TblCarsData.TestExpireDateH,"
MySQL = MySQL & "                      dbo.TblCarsData.VehicleLong, dbo.TblCarsData.EquQty, dbo.TblCarsData.Capacity, dbo.TblCarsData.LicenseNO, dbo.TblCarsData.EndContractDate,"
MySQL = MySQL & "                      dbo.TblCarsData.ContractID, dbo.TblCarsData.SetCount, dbo.TblCarsData.Rate, dbo.TblCarsData.EndContractDateH, dbo.TblCarsData.EndAllocationDate,"
MySQL = MySQL & "                      dbo.TblCarsData.Rep, dbo.TblCarsData.MaxCap, dbo.TblCarsData.OperatorN, dbo.TblCarsData.EqupName, dbo.TblCarsData.DriLicense,"
MySQL = MySQL & "                      dbo.TblCarsData.DriLicenseDate, dbo.TblCarsData.DriLicenseNo, dbo.TblCarsData.Department, dbo.TblCarsData.Job, dbo.TblCarsData.Natinality,"
MySQL = MySQL & "                      dbo.TblCarsData.Authorization2, dbo.TblCarsData.BagAmbulance, dbo.TblCarsData.SubsBattery, dbo.TblCarsData.FireExtingui, dbo.TblCarsData.TrackingDevice,"
MySQL = MySQL & "                      dbo.TblCarsData.Triangle, dbo.TblCarsData.Receipt, dbo.TblCarsData.KeyReserve, dbo.TblCarsData.Exam, dbo.TblCarsData.Licenses, dbo.TblCarsData.AuthorDate,"
MySQL = MySQL & "                      dbo.TblCarsData.AuthorType, dbo.TblCarsData.Insurance, dbo.TblCarsData.TrackingNo, dbo.TblCarsData.OwnerName, dbo.TblCarsData.LeaderName,"
MySQL = MySQL & "                      dbo.TblCarsData.EmpType, dbo.TblCarsData.Stickers, dbo.TblCarsData.Guarantee, dbo.TblCarsData.CoverKey, dbo.TblCarsData.Crane, dbo.TblCarsData.SpareTyre,"
MySQL = MySQL & "                      dbo.TblCarsData.Battery, dbo.TblCarsData.Anntena, dbo.TblCarsData.Recorder, dbo.TblCarsData.SunScreens, dbo.TblCarsData.Pedals, dbo.TblCarsData.CounryID,"
MySQL = MySQL & "                      dbo.TblCountriesData.CountryName, dbo.TblCarsData.CityID, dbo.TblCountriesGovernments.GovernmentName, dbo.TblCarsData.InnerLights,"
MySQL = MySQL & "                      dbo.TblCarsData.driverMirror, dbo.TblCarsData.sideMirror, dbo.TblCarsData.cleaner, dbo.TblCarsData.authorizeExamination, dbo.TblCarsData.authorizeLicense,"
MySQL = MySQL & "                      dbo.TblCarsData.FormOrignal, dbo.TblCarsData.LetterPrice, dbo.TblCarsData.LetterCount, dbo.TblCarsData.Total, dbo.TblCarsData.Machineno,"
MySQL = MySQL & "                      dbo.TblCarsData.Machineno1, dbo.TblCarsData.Gearno1, dbo.TblCarsData.Gearno, dbo.TblCarsData.Chesis, dbo.TblCarsData.TypeCar, dbo.TblCarsData.VColor,"
MySQL = MySQL & "                      dbo.TblColor.name AS ColorName, dbo.TblColor.namee AS ColorNameE, dbo.TblCarsData.VModel, dbo.TblCarModels.Model AS NameModel,"
MySQL = MySQL & "                      dbo.TblCarModels.ModelE AS NameModelE, dbo.TblCarsData.LocationID, dbo.EmpGroupDep.GroupName, dbo.TblCarsData.VType, dbo.TblCarsData.Keys,"
MySQL = MySQL & "                      dbo.TblCarsData.InsuranceNO, dbo.TblCarsDataDet.PartID, FixedAssets_1.code AS Partcode, FixedAssets_1.Name AS PartName,"
MySQL = MySQL & "                      FixedAssets_1.namee AS PartNameE"
MySQL = MySQL & " FROM         dbo.TblCarsDataDet LEFT OUTER JOIN"
MySQL = MySQL & "                      dbo.FixedAssets FixedAssets_1 ON dbo.TblCarsDataDet.PartID = FixedAssets_1.id RIGHT OUTER JOIN"
MySQL = MySQL & "                      dbo.TblCarsData ON dbo.TblCarsDataDet.EqupID = dbo.TblCarsData.fixedAssetid LEFT OUTER JOIN"
MySQL = MySQL & "                      dbo.EmpGroupDep ON dbo.TblCarsData.LocationID = dbo.EmpGroupDep.GroupID LEFT OUTER JOIN"
MySQL = MySQL & "                      dbo.TblCarModels ON dbo.TblCarsData.VModel = dbo.TblCarModels.Id LEFT OUTER JOIN"
MySQL = MySQL & "                      dbo.TblColor ON dbo.TblCarsData.VColor = dbo.TblColor.Id LEFT OUTER JOIN"
MySQL = MySQL & "                      dbo.TblCountriesGovernments ON dbo.TblCarsData.CityID = dbo.TblCountriesGovernments.GovernmentID LEFT OUTER JOIN"
MySQL = MySQL & "                      dbo.TblCountriesData ON dbo.TblCarsData.CounryID = dbo.TblCountriesData.CountryID LEFT OUTER JOIN"
MySQL = MySQL & "                      dbo.TblEmpDepartments RIGHT OUTER JOIN"
MySQL = MySQL & "                      dbo.TblEmployee ON dbo.TblEmpDepartments.DeparmentID = dbo.TblEmployee.DepartmentID LEFT OUTER JOIN"
MySQL = MySQL & "                      dbo.TblEmpJobsTypes ON dbo.TblEmployee.JobTypeID = dbo.TblEmpJobsTypes.JobTypeID ON"
MySQL = MySQL & "                      dbo.TblCarsData.Emp_id = dbo.TblEmployee.Emp_ID LEFT OUTER JOIN"
MySQL = MySQL & "                      dbo.insurance_companies ON dbo.TblCarsData.InsuranceCompanyId = dbo.insurance_companies.id LEFT OUTER JOIN"
MySQL = MySQL & "                      dbo.FixedAssets ON dbo.TblCarsData.fixedAssetid = dbo.FixedAssets.id LEFT OUTER JOIN"
MySQL = MySQL & "                      dbo.TBLCarTypes ON dbo.TblCarsData.CarsTypeId = dbo.TBLCarTypes.id LEFT OUTER JOIN"
MySQL = MySQL & "                      dbo.TblBranchesData ON dbo.TblCarsData.Branch_NO = dbo.TblBranchesData.branch_id"
MySQL = MySQL & "  where    TblCarsData.ID  =" & val(XPTxtID.Text)


 If SystemOptions.UserInterface = ArabicInterface Then
          StrFileName = App.path & "\REPORTS\REPORTS NEW\" & "rpt_CarsAll.rpt"
     Else
        StrFileName = App.path & "\REPORTS\REPORTS NEW\" & "rpt_CarsAllE.rpt"
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
        StrReportTitle = "" '& StrAccountName
    Else
        xReport.ParameterFields(1).AddCurrentValue cCompanyInfo.ArabCompanyName ' RPTCompany_Name_Eng
      '  xReport.ParameterFields(2).AddCurrentValue RPTComment_Eng
     '   xReport.ParameterFields(4).AddCurrentValue get_branch_name(val(my_branch))
        StrReportTitle = ""
    End If
   '
    xReport.ParameterFields(3).AddCurrentValue user_name
    xReport.ParameterFields(5).AddCurrentValue WriteNo(SumTotalExpen(val(DcFixedAssets.BoundText)) + val(lbl(46).Caption) + val(TotalValue.Caption), 0)
    xReport.ParameterFields(6).AddCurrentValue SumTotalExpen(val(DcFixedAssets.BoundText)) + val(TotalValue.Caption)
    
    xReport.ParameterFields(7).AddCurrentValue val(lbl(46).Caption)
    xReport.reporttitle = StrReportTitle
    xReport.EnableParameterPrompting = False
    xReport.ApplicationName = App.title
    xReport.ReportAuthor = App.title
    ''///////
    Set CViewer = New ClsReportViewer
    CViewer.FireReport xReport, WindowTarget, "", , , , StrFileName

    RsData.Close
    Set RsData = Nothing
    Screen.MousePointer = vbDefault
End Function
Sub SaveAssest(Optional FexdID As Double = 0)
Dim sql As String
Dim StrSQL As String
Dim AsID As Double
Dim Msg As String
Dim Rs5 As ADODB.Recordset
Set Rs5 = New ADODB.Recordset
If Me.TxtModFlg.Text = "E" Then
sql = "Delete  from TblAssestes where AsFixedID=" & FexdID & ""
Cn.Execute sql
End If
sql = "Select * from TblAssestes where 1=-1"
Rs5.Open sql, Cn, adOpenKeyset, adLockOptimistic, adCmdText
Rs5.AddNew
Rs5("BranchID").value = val(Dcbranch.BoundText)
Rs5("CarsDataID").value = val(XPTxtID.Text)
Rs5("FlgCarNotFixed").value = 1
If SystemOptions.UserInterface = ArabicInterface Then
Msg = "„‰ „·› «·„⁄œ« /«·”Ì«—« "
Else
Msg = "From Cars File"
End If
Rs5("AsFixedID").value = FexdID
Rs5("AsDes").value = Msg
Rs5("AsName").value = TxtEqupName.Text
Rs5("AsCode").value = val(TxtId.Text)
Rs5.update
AsID = IIf(IsNull(Rs5("AsID").value), 0, Rs5("AsID").value)
SaveDriveAssest AsID, val(DcEmployee.BoundText)
End Sub
Sub fillComboParts()
Dim StrSQL As String
  '  str = " select id , EqupName  from TblCarsData where ID <> " & XPTxtID.Text & " "
            If SystemOptions.UserInterface = ArabicInterface Then
                   StrSQL = " SELECT     id, Name"
                   StrSQL = StrSQL & " from dbo.FixedAssets where ISEQUP = 1 and id <>" & val(DcFixedAssets.BoundText) & ""
                   StrSQL = StrSQL & " order by Name   "
                Else

                    StrSQL = " SELECT     id, Namee"
                    StrSQL = StrSQL & " from dbo.FixedAssets where ISEQUP = 1 and id <>" & val(DcFixedAssets.BoundText) & ""
                    StrSQL = StrSQL & " order by Namee "
                        
                End If
    fill_combo PartDC, StrSQL
End Sub
Private Sub Add_CarParts()
    Dim BeginTrans As Boolean
    Dim StrSQL As String
    Dim Msg As String
    Dim i As Integer
  'On Error GoTo errortrap
  
 If val(TxtId.Text) = 0 Then
    If SystemOptions.UserInterface = ArabicInterface Then
        MsgBox "Ì—ÃÏ Õ›Ÿ »Ì«‰«  «·„⁄œ… «Ê·«"
     Else
        MsgBox "Please Save Data"
    End If
 Exit Sub
 End If
 
    With VSFlexGrid13
        .Rows = rs_CarParts.RecordCount + 1
            For i = 1 To .Rows - 1
                If .TextMatrix(i, .ColIndex("PartID")) = PartDC.BoundText Then
                    If SystemOptions.UserInterface = ArabicInterface Then
                        MsgBox ("Â–« «·„·Õﬁ „ÊÃÊœ „”»ﬁ«")
                    Else
                         MsgBox "This Part is already added"
                    End If
                    Exit Sub
                End If
            Next
    End With
  
    Cn.BeginTrans
    BeginTrans = True
    
    Set rs_CarParts = New ADODB.Recordset
    StrSQL = "SELECT  *  From TblCarsDataDet  "
    rs_CarParts.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText
    rs_CarParts.AddNew
    rs_CarParts("EqupID") = IIf(DcFixedAssets.BoundText = "", Null, val(DcFixedAssets.BoundText))
    rs_CarParts("PartID") = IIf(PartDC.BoundText = "", Null, val(PartDC.BoundText))
    rs_CarParts.update
    
    Cn.CommitTrans
    BeginTrans = False
    If SystemOptions.UserInterface = ArabicInterface Then
        MsgBox (" „ Õ›Ÿ «·»Ì«‰«  »‰Ã«Õ")
    Else
        MsgBox "Save Successfully"
    End If
    
    Retrive_CarParts
    'Clear_CarsExpens
Exit Sub
errortrap:

    If BeginTrans = True Then
        BeginTrans = False
        Cn.RollbackTrans
    End If

    If Err.Number = -2147217900 Then
    If SystemOptions.UserInterface = ArabicInterface Then
        Msg = "·« Ì„ﬂ‰ Õ›Ÿ Â–Â «·»Ì«‰«  " & CHR(13)
        Msg = Msg + "·ﬁœ  „ «œŒ«· ﬁÌ„ €Ì— ’«·Õ… " & CHR(13)
        Msg = Msg + " √ﬂœ „‰ œﬁ… «·»Ì«‰«  Ê√⁄œ «·„Õ«Ê·…"
      Else
      Msg = "Can not save make sure of the validity of the data"
      End If
        MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
        Exit Sub
    End If
If SystemOptions.UserInterface = ArabicInterface Then
    Msg = "⁄›Ê«...ÕœÀ Œÿ√ „« √À‰«¡ Õ›Ÿ Â–Â «·»Ì«‰«  " & CHR(13)
  Else
  Msg = "Sory..error douring save data "
  End If
    MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
End Sub
Private Sub Retrive_CarParts()
    Dim i As Integer
        VSFlexGrid13.Clear flexClearScrollable, flexClearEverything
    VSFlexGrid13.Rows = 1
    Set rs_CarParts = New ADODB.Recordset
    Dim StrSQL As String
    
    StrSQL = " SELECT     dbo.TblCarsDataDet.ID AS PID, dbo.FixedAssets.Name, dbo.FixedAssets.code, dbo.FixedAssets.namee, dbo.TblCarsDataDet.EqupID, dbo.TblCarsDataDet.PartID"
    StrSQL = StrSQL & "             FROM         dbo.TblCarsDataDet LEFT OUTER JOIN"
    StrSQL = StrSQL & "                  dbo.FixedAssets ON dbo.TblCarsDataDet.PartID = dbo.FixedAssets.id"
    StrSQL = StrSQL & "      Where (dbo.TblCarsDataDet.EqupID = " & val(DcFixedAssets.BoundText) & ")"

    
    rs_CarParts.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText
    
    VSFlexGrid13.Rows = 1
    If rs_CarParts.RecordCount > 0 Then
        rs_CarParts.MoveFirst
        With VSFlexGrid13
            .Rows = rs_CarParts.RecordCount + 1
            For i = 1 To .Rows - 1
                .TextMatrix(i, .ColIndex("Serial")) = i
                .TextMatrix(i, .ColIndex("ID")) = IIf(IsNull(rs_CarParts("PID").value), 0, rs_CarParts("PID").value)
                .TextMatrix(i, .ColIndex("PartID")) = IIf(IsNull(rs_CarParts("PartID").value), 0, rs_CarParts("PartID").value)
                .TextMatrix(i, .ColIndex("Code")) = IIf(IsNull(rs_CarParts("code").value), "", rs_CarParts("code").value)
                If SystemOptions.UserInterface = ArabicInterface Then
                .TextMatrix(i, .ColIndex("Name")) = IIf(IsNull(rs_CarParts("Name").value), "", rs_CarParts("Name").value)
                Else
                .TextMatrix(i, .ColIndex("Name")) = IIf(IsNull(rs_CarParts("namee").value), "", rs_CarParts("namee").value)
                End If
                rs_CarParts.MoveNext
            Next
         End With
    End If
End Sub
Private Sub Update_CarParts()
    Dim BeginTrans As Boolean
    Dim StrSQL As String
    Dim Msg As String
    Dim str As String, sr As String
    
    On Error GoTo errortrap

    If val(TxtId.Text) = 0 Then
        If SystemOptions.UserInterface = ArabicInterface Then
            MsgBox "Ì—ÃÏ Õ›Ÿ »Ì«‰«  «·„⁄œ… «Ê·«"
        Else
            MsgBox "Please Save Data"
        End If
        Exit Sub
    End If
    
    str = VSFlexGrid13.TextMatrix(VSFlexGrid13.Row, VSFlexGrid13.ColIndex("id"))
    If str <> "" Then
        Cn.BeginTrans
        BeginTrans = True
        
        StrSQL = "Update TblCarsDataDet Set PartID =" & val(Me.PartDC.BoundText) & " Where TblCarsDataDet.ID = " & val(str) & " "
        
        Cn.Execute StrSQL, , adExecuteNoRecords
           
        Cn.CommitTrans
        BeginTrans = False
        
        If SystemOptions.UserInterface = ArabicInterface Then
            MsgBox (" „ Õ›Ÿ  ⁄œÌ· «·»Ì«‰«  »‰Ã«Õ")
        Else
            MsgBox "Save Successfully"
        End If
    Retrive_CarParts
Exit Sub
errortrap:
    If BeginTrans = True Then
        BeginTrans = False
        Cn.RollbackTrans
    End If

    If Err.Number = -2147217900 Then
        If SystemOptions.UserInterface = ArabicInterface Then
            Msg = "·« Ì„ﬂ‰ Õ›Ÿ Â–Â «·»Ì«‰«  " & CHR(13)
            Msg = Msg + "·ﬁœ  „ «œŒ«· ﬁÌ„ €Ì— ’«·Õ… " & CHR(13)
            Msg = Msg + " √ﬂœ „‰ œﬁ… «·»Ì«‰«  Ê√⁄œ «·„Õ«Ê·…"
        Else
            Msg = " Can not save make sure of the validity of the data"
        End If
        MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
        Exit Sub
    End If
    
    If SystemOptions.UserInterface = ArabicInterface Then
        Msg = "⁄›Ê«...ÕœÀ Œÿ√ „« √À‰«¡ Õ›Ÿ Â–Â «·»Ì«‰«  " & CHR(13)
    Else
        Msg = "Sory..error douring save data"
    End If
    MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
End If
End Sub
Private Sub Del_CarParts()
    Dim Msg As String
    Dim StrSQL As String
    Dim RsTemp As New ADODB.Recordset
    Dim str As String, sr As String
    
    On Error GoTo ErrTrap

    str = VSFlexGrid13.TextMatrix(VSFlexGrid13.Row, VSFlexGrid13.ColIndex("id"))
        
     If str <> "" And str <> "id" Then
        If SystemOptions.UserInterface = ArabicInterface Then
            Msg = "”Ì „ Õ–› »Ì«‰«  ”ÿ— —ﬁ„ " & CHR(13)
            Msg = Msg + (VSFlexGrid13.TextMatrix(VSFlexGrid13.Row, VSFlexGrid13.ColIndex("Serial"))) & CHR(13)
            Msg = Msg + " Â·  —€» ›Ì Õ–› Â–Â «·»Ì«‰« ø"
        Else
            Msg = "Confirm Delete"
        End If
        If MsgBox(Msg, vbYesNo + vbQuestion + vbMsgBoxRight + vbMsgBoxRtlReading, App.title) = vbYes Then
            If Not rs_CarParts.RecordCount < 1 Then
            
                StrSQL = "delete From TblCarsDataDet  where  ID =" & val(str) & " "
                
                Cn.Execute StrSQL, , adExecuteNoRecords
                
                Set rs_CarParts = New ADODB.Recordset
                StrSQL = "SELECT  *  From TblCarsDataDet"
                rs_CarParts.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText
                         VSFlexGrid13.Clear flexClearScrollable, flexClearEverything
    VSFlexGrid13.Rows = 1
                
                If rs_CarParts.RecordCount < 1 Then
                Else
                    Retrive_CarParts
                End If
            End If
            Else
   
        End If
    Else
        If SystemOptions.UserInterface = ArabicInterface Then
            Msg = "Â–Â «·⁄„·Ì… €Ì— „ «Õ… ÕÌÀ √‰Â ·«ÌÊÃœ √Ì ”Ã·« "
        Else
            MsgBox "This process is not available. There are no records"
        End If
        MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
        Exit Sub
    End If
    Retrive_CarsExp
    Exit Sub
ErrTrap:
    If SystemOptions.UserInterface = ArabicInterface Then
        Msg = "·«Ì„ﬂ‰ Õ–› Â–« «·”Ã· · ﬂ«„· «·»Ì«‰«  " & CHR(13) & " ÊÃœ »Ì«‰«  „— »ÿ… »Â–« «·„Œ«·›… "
    Else
        Msg = "Can not delete "
    End If
        Msg = Msg & CHR(13) & Err.description
        MsgBox Msg, vbMsgBoxRight + vbMsgBoxRtlReading + vbExclamation, App.title
        rs_CarParts.CancelUpdate
End Sub

Private Sub VSFlexGrid13_Click()
    With Me.VSFlexGrid13
        If .Row > 0 Then
            PartDC.BoundText = .TextMatrix(.Row, .ColIndex("PartID"))
        End If
    End With
End Sub
Private Sub ISButton2_Click()
    Add_CarParts
End Sub
Private Sub ISButton1_Click()
    Update_CarParts
End Sub
Private Sub ISButton3_Click()
    Del_CarParts
End Sub
Function print_report_Parts(Optional NoteSerial As String)
    Dim MySQL As String
    Dim RsData As New ADODB.Recordset
    Dim xApp As New CRAXDRT.Application
    Dim xReport As CRAXDRT.Report
    Dim CViewer As ClsReportViewer
    Dim StrReportTitle As String
    Dim StrFileName As String
    Dim Msg As String
    Dim ActivID As String
    Dim Arche As String
    Dim Room As String
    Dim Box As String
    Dim Shelf As String
    Dim dep As String
    Dim docType As String
    Dim i As Integer
    
MySQL = " SELECT     dbo.TblCarsDataDet.ID AS PID, dbo.FixedAssets.Name, dbo.FixedAssets.code, dbo.FixedAssets.namee, dbo.TblCarsDataDet.EqupID, dbo.TblCarsDataDet.PartID, "
MySQL = MySQL & "                      FixedAssets_1.code AS codeH, FixedAssets_1.Name AS NameH, FixedAssets_1.namee AS nameeH"
MySQL = MySQL & " FROM         dbo.TblCarsDataDet LEFT OUTER JOIN"
MySQL = MySQL & "                      dbo.FixedAssets FixedAssets_1 ON dbo.TblCarsDataDet.EqupID = FixedAssets_1.id LEFT OUTER JOIN"
MySQL = MySQL & "                      dbo.FixedAssets ON dbo.TblCarsDataDet.PartID = dbo.FixedAssets.id"
MySQL = MySQL & " Where (dbo.TblCarsDataDet.EqupID = " & val(DcFixedAssets.BoundText) & ")"

        If SystemOptions.UserInterface = ArabicInterface Then
            StrFileName = App.path & "\REPORTS\REPORTS NEW\" & "RepParts.rpt"
        Else
            StrFileName = App.path & "\REPORTS\REPORTS NEW\" & "RepPartsE.rpt"
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


